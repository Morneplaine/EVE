"""
EVE Online Overview Color Alert

Monitors a window (e.g. the Overview / ship list) and plays an alarm when
teal, yellow, or red color bands appear in a narrow band on the right side.
Purple is explicitly ignored.

Alarm sound: WAV file stored in the program folder (alarm.wav). Use
"Select sound file..." in the GUI to pick a WAV; it is copied into the
program folder so the path never changes. 16-bit PCM WAV is recommended.

Run every 1 second. Can be run standalone (GUI or CLI) or imported.
"""

import ctypes
from ctypes import wintypes
import os
import re
import shutil
import time
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog
from pathlib import Path

# Alarm sound: WAV file in program folder (best: 16-bit PCM; use "Select sound file..." in GUI to set)
APP_DIR = Path(__file__).resolve().parent
ALARM_WAV_NAME = "alarm.wav"
ALARM_SOUND_PATH = APP_DIR / ALARM_WAV_NAME
# Subfolder for screenshots and debug captures (keeps program folder tidy)
CAPTURES_DIR = APP_DIR / "overview_alert_captures"


def _next_capture_number(prefix: str, suffix: str) -> int:
    """Return next number N such that prefix_NNNsuffix does not exist in CAPTURES_DIR. E.g. prefix='overview_alert_screenshot', suffix='.png'."""
    CAPTURES_DIR.mkdir(parents=True, exist_ok=True)
    pattern = re.compile(re.escape(prefix) + r"_(\d+)" + re.escape(suffix))
    max_n = 0
    for p in CAPTURES_DIR.iterdir():
        if not p.is_file():
            continue
        m = pattern.match(p.name)
        if m:
            max_n = max(max_n, int(m.group(1)))
    return max_n + 1

# Optional: Pillow for screen capture
try:
    from PIL import ImageGrab
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

# Windows API for window and monitor enumeration
user32 = ctypes.windll.user32
MONITORINFOF_PRIMARY = 1

# ---------------------------------------------------------------------------
# Color detection (RGB). Tuned for EVE Overview highlights.
# Purple ~ #5B1C5B is excluded.
# Can use built-in thresholds or optional ranges: {"teal": (r_min,r_max,g_min,g_max,b_min,b_max), ...}.
# ---------------------------------------------------------------------------

# Default RGB ranges (r_min, r_max, g_min, g_max, b_min, b_max); purple = exclude band.
DEFAULT_COLOR_RANGES = {
    "teal": (0, 120, 140, 255, 140, 255),
    "yellow": (160, 255, 140, 255, 0, 165),
    "red": (180, 255, 0, 100, 0, 100),
    "purple": (50, 130, 0, 55, 50, 130),
}


def _in_range(r: int, g: int, b: int, r_min: int, r_max: int, g_min: int, g_max: int, b_min: int, b_max: int) -> bool:
    return r_min <= r <= r_max and g_min <= g <= g_max and b_min <= b <= b_max


def _is_purple(r: int, g: int, b: int, color_ranges: dict | None = None) -> bool:
    """Exclude purple (e.g. EVE overview selection ~ #5B1C5B)."""
    if color_ranges and "purple" in color_ranges:
        t = color_ranges["purple"]
        return _in_range(r, g, b, t[0], t[1], t[2], t[3], t[4], t[5])
    if g > 55:
        return False
    return 50 <= r <= 130 and 50 <= b <= 130


def is_teal(r: int, g: int, b: int, color_ranges: dict | None = None) -> bool:
    """Teal/cyan: low R, high G and B (or use ranges)."""
    if _is_purple(r, g, b, color_ranges):
        return False
    if color_ranges and "teal" in color_ranges:
        t = color_ranges["teal"]
        return _in_range(r, g, b, t[0], t[1], t[2], t[3], t[4], t[5])
    return r < 120 and g > 140 and b > 140 and abs(g - b) < 80


def is_yellow(r: int, g: int, b: int, color_ranges: dict | None = None) -> bool:
    """Yellow/orange band: high R and G, low B (or use ranges)."""
    if _is_purple(r, g, b, color_ranges):
        return False
    if color_ranges and "yellow" in color_ranges:
        t = color_ranges["yellow"]
        return _in_range(r, g, b, t[0], t[1], t[2], t[3], t[4], t[5])
    return r > 160 and g > 140 and b < 165


def is_red(r: int, g: int, b: int, color_ranges: dict | None = None) -> bool:
    """Red: high R, low G and B (or use ranges)."""
    if _is_purple(r, g, b, color_ranges):
        return False
    if color_ranges and "red" in color_ranges:
        t = color_ranges["red"]
        return _in_range(r, g, b, t[0], t[1], t[2], t[3], t[4], t[5])
    return r > 180 and g < 100 and b < 100


def pixel_matches_alert_color(r: int, g: int, b: int, color_ranges: dict | None = None) -> str | None:
    """Return 'teal'|'yellow'|'red' if pixel matches an alert color, else None. color_ranges: optional dict of name -> (r_min,r_max,g_min,g_max,b_min,b_max)."""
    if is_teal(r, g, b, color_ranges):
        return "teal"
    if is_yellow(r, g, b, color_ranges):
        return "yellow"
    if is_red(r, g, b, color_ranges):
        return "red"
    return None


# ---------------------------------------------------------------------------
# Window finding (Windows)
# ---------------------------------------------------------------------------

# Used by EnumWindows callback (ctypes can't pass Python list via LPARAM)
_enum_windows_list = []


def _enum_callback(hwnd, _lparam):
    if not user32.IsWindowVisible(hwnd):
        return True
    length = user32.GetWindowTextLengthW(hwnd) + 1
    buf = ctypes.create_unicode_buffer(length)
    user32.GetWindowTextW(hwnd, buf, length)
    title = buf.value or ""
    _enum_windows_list.append((hwnd, title))
    return True


WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)


def get_window_rect(hwnd):
    """Return (left, top, width, height) in pixels."""
    rect = wintypes.RECT()
    if not user32.GetWindowRect(hwnd, ctypes.byref(rect)):
        return None
    return (
        rect.left,
        rect.top,
        rect.right - rect.left,
        rect.bottom - rect.top,
    )


def list_all_visible_windows():
    """Return list of (hwnd, title) for all visible windows that have a non-empty title and valid size. Deduplicated by hwnd."""
    global _enum_windows_list
    _enum_windows_list = []
    callback = WNDENUMPROC(_enum_callback)
    user32.EnumWindows(callback, 0)
    results = _enum_windows_list
    _enum_windows_list = []
    seen_hwnd = set()
    out = []
    for hwnd, title in results:
        if hwnd in seen_hwnd:
            continue
        seen_hwnd.add(hwnd)
        title = (title or "").strip()
        if not title:
            continue
        rect = get_window_rect(hwnd)
        if not rect or rect[2] < 20 or rect[3] < 20:
            continue
        out.append((hwnd, title))
    return out


def find_windows_by_title(substring: str):
    """Return list of (hwnd, title) for visible windows whose title contains substring (case-insensitive)."""
    all_w = list_all_visible_windows()
    sub = substring.lower()
    return [(h, t) for h, t in all_w if sub in (t or "").lower()]


def get_window_rect_by_title(substring: str):
    """
    Find first visible window whose title contains substring.
    Return (left, top, width, height) or None if not found.
    """
    candidates = find_windows_by_title(substring)
    for hwnd, _ in candidates:
        rect = get_window_rect(hwnd)
        if rect and rect[2] > 0 and rect[3] > 0:
            return rect
    return None


# ---------------------------------------------------------------------------
# Monitor enumeration (for full-screen capture, multi-monitor)
# ---------------------------------------------------------------------------

_monitor_list = []


def _monitor_enum_callback(h_monitor, h_dc, lprc_rect, lparam):
    r = lprc_rect.contents
    left, top = r.left, r.top
    width = r.right - r.left
    height = r.bottom - r.top
    # GetMonitorInfo to know if primary and device name
    class MONITORINFOEXW(ctypes.Structure):
        _fields_ = [
            ("cbSize", wintypes.DWORD),
            ("rcMonitor", wintypes.RECT),
            ("rcWork", wintypes.RECT),
            ("dwFlags", wintypes.DWORD),
            ("szDevice", wintypes.WCHAR * 32),
        ]
    info = MONITORINFOEXW()
    info.cbSize = ctypes.sizeof(MONITORINFOEXW)
    if user32.GetMonitorInfoW(h_monitor, ctypes.byref(info)):
        is_primary = bool(info.dwFlags & MONITORINFOF_PRIMARY)
        name = info.szDevice or ""
    else:
        is_primary = False
        name = ""
    _monitor_list.append({
        "left": left, "top": top, "width": width, "height": height,
        "is_primary": is_primary, "device": name,
    })
    return True


def get_monitors():
    """Return list of monitors: [{"left", "top", "width", "height", "is_primary", "device"}, ...]."""
    global _monitor_list
    _monitor_list = []
    MONITORENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HMONITOR, wintypes.HDC, ctypes.POINTER(wintypes.RECT), wintypes.LPARAM)
    callback = MONITORENUMPROC(_monitor_enum_callback)
    user32.EnumDisplayMonitors(None, None, callback, 0)
    return list(_monitor_list)


# ---------------------------------------------------------------------------
# Screen capture and band sampling
# ---------------------------------------------------------------------------

def capture_region(left: int, top: int, width: int, height: int):
    """Capture screen region. Returns PIL Image or None."""
    if not HAS_PIL:
        return None
    bbox = (left, top, left + width, top + height)
    return ImageGrab.grab(bbox)


def sample_band(image, x_start_ratio: float = 0.85, x_end_ratio: float = 0.98, y_start_ratio: float = 0.0, y_end_ratio: float = 0.5, step: int = 2):
    """
    Sample pixels in a band of the image: X 85%–98% of width, Y top to middle (0–50%).
    step: sample every N pixels in both directions to keep it fast.
    Returns list of (r, g, b) tuples.
    """
    w, h = image.size
    x0 = int(w * x_start_ratio)
    x1 = int(w * x_end_ratio)
    if x0 >= x1:
        x1 = min(x0 + 1, w)
    y0 = int(h * y_start_ratio)
    y1 = int(h * y_end_ratio)
    if y0 >= y1:
        y1 = min(y0 + 1, h)
    pixels = []
    for x in range(x0, x1, max(1, step)):
        for y in range(y0, y1, step):
            try:
                px = image.getpixel((x, y))
            except IndexError:
                continue
            if isinstance(px, int):
                # Grayscale
                pixels.append((px, px, px))
            elif len(px) >= 3:
                pixels.append((px[0], px[1], px[2]))
    return pixels


def check_band_for_alert_colors(pixels, require_count: int = 8, color_ranges: dict | None = None):
    """
    Check sampled pixels for teal, yellow, or red (excluding purple).
    If at least require_count pixels match the same alert color, return that color name.
    color_ranges: optional dict of name -> (r_min,r_max,g_min,g_max,b_min,b_max).
    """
    counts = {"teal": 0, "yellow": 0, "red": 0}
    for r, g, b in pixels:
        color = pixel_matches_alert_color(r, g, b, color_ranges)
        if color:
            counts[color] += 1
    for color, count in counts.items():
        if count >= require_count:
            return color
    return None


def save_debug_capture(
    hwnd=None,
    monitor_rect=None,
    x_start: float = 0.85,
    x_end: float = 0.98,
    y_start: float = 0.0,
    y_end: float = 0.5,
    sample_step: int = 2,
    require_pixels: int = 8,
    color_ranges: dict | None = None,
    save_dir: str | None = None,
    capture_index: int | None = None,
):
    """
    Capture the region (monitor or window) and the sampled band, save as images, and write a debug report.
    Pass monitor_rect=(left, top, width, height) for full screen, or hwnd for a window.
    Returns (success: bool, message: str).
    """
    if not HAS_PIL:
        return False, "Pillow (PIL) is required for capture."
    if monitor_rect is not None:
        left, top, width, height = monitor_rect
    elif hwnd is not None:
        rect = get_window_rect(hwnd)
        if not rect:
            return False, "Window no longer valid."
        left, top, width, height = rect
    else:
        return False, "Provide hwnd or monitor_rect."
    full = capture_region(left, top, width, height)
    if full is None:
        return False, "Failed to capture window."
    if save_dir is None:
        save_dir = str(CAPTURES_DIR)
    os.makedirs(save_dir, exist_ok=True)
    if capture_index is None:
        capture_index = _next_capture_number("overview_alert_full", ".png")
    num = f"{capture_index:03d}"
    full_path = os.path.join(save_dir, f"overview_alert_full_{num}.png")
    band_path = os.path.join(save_dir, f"overview_alert_band_{num}.png")
    report_path = os.path.join(save_dir, f"overview_alert_debug_{num}.txt")
    try:
        full.save(full_path)
    except Exception as e:
        return False, f"Could not save full capture: {e}"
    w, h = full.size
    x0 = int(w * x_start)
    x1 = int(w * x_end)
    if x1 <= x0:
        x1 = x0 + 1
    y0 = int(h * y_start)
    y1 = int(h * y_end)
    if y1 <= y0:
        y1 = y0 + 1
    band = full.crop((x0, y0, min(x1, w), min(y1, h)))
    try:
        band.save(band_path)
    except Exception as e:
        return False, f"Could not save band image: {e}"
    pixels = sample_band(full, x_start_ratio=x_start, x_end_ratio=x_end, y_start_ratio=y_start, y_end_ratio=y_end, step=sample_step)
    counts = {"teal": 0, "yellow": 0, "red": 0}
    for r, g, b in pixels:
        color = pixel_matches_alert_color(r, g, b, color_ranges)
        if color:
            counts[color] += 1
    detected = check_band_for_alert_colors(pixels, require_count=require_pixels, color_ranges=color_ranges)
    # Pixels that are "yellow-ish" (high R, high G) for tuning
    yellow_candidates = [(r, g, b) for r, g, b in pixels if r > 150 and g > 150 and b < 200]
    lines = [
        "Overview Alert debug report",
        "=" * 50,
        f"Full capture: {full_path}",
        f"Band only:    {band_path}",
        "",
        f"Window: {width}x{height}, band X: {x0}-{x1} ({(x_start*100):.0f}%-{(x_end*100):.0f}%), Y: {y0}-{y1} ({(y_start*100):.0f}%-{(y_end*100):.0f}%)",
        f"Pixels sampled (step={sample_step}): {len(pixels)}",
        "",
        "Detection counts (current thresholds):",
        f"  teal:   {counts['teal']} (need >= {require_pixels} to trigger)",
        f"  yellow: {counts['yellow']} (need >= {require_pixels} to trigger)",
        f"  red:    {counts['red']} (need >= {require_pixels} to trigger)",
        f"Detected color: {detected or 'none'}",
        "",
        "RGB ranges: adjust in Settings tab if using custom ranges.",
        "",
        "Sample of up to 30 raw RGB values from the band:",
    ]
    for i, (r, g, b) in enumerate(pixels[:30]):
        lines.append(f"  {i+1}. R={r:3d} G={g:3d} B={b:3d}")
    if yellow_candidates:
        lines.append("")
        lines.append("Pixels that are yellow-ish (R>150, G>150, B<200) - for tuning is_yellow():")
        for i, (r, g, b) in enumerate(yellow_candidates[:20]):
            lines.append(f"  R={r:3d} G={g:3d} B={b:3d}")
    try:
        with open(report_path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
    except Exception as e:
        return False, f"Could not write report: {e}"
    return True, f"Saved: {full_path}\nBand: {band_path}\nReport: {report_path}"


# ---------------------------------------------------------------------------
# Alarm (WAV in program folder, played via winsound — no extra deps, no focus steal)
# ---------------------------------------------------------------------------

def play_alarm():
    """Play the alarm WAV from the program folder, or Windows beep if missing."""
    path = Path(ALARM_SOUND_PATH)
    if path.is_file():
        try:
            import winsound
            winsound.PlaySound(str(path), winsound.SND_FILENAME | winsound.SND_ASYNC)
            return
        except Exception:
            pass
    try:
        import winsound
        winsound.MessageBeep(winsound.MB_ICONEXCLAMATION)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Main loop
# ---------------------------------------------------------------------------

def run_once(
    window_title_substring: str = "EVE",
    hwnd=None,
    monitor_rect=None,
    x_start: float = 0.85,
    x_end: float = 0.98,
    y_start: float = 0.0,
    y_end: float = 0.5,
    sample_step: int = 2,
    require_pixels: int = 8,
    color_ranges: dict | None = None,
):
    """
    Perform one check: capture band and detect teal/yellow/red.
    If monitor_rect (left, top, width, height) is given, capture that screen region (full monitor).
    Else if hwnd is given, use that window; else find by window_title_substring.
    Returns (detected_color or None, error_message or None).
    """
    if not HAS_PIL:
        return None, "Pillow (PIL) is required. Install with: pip install Pillow"

    if monitor_rect is not None:
        left, top, width, height = monitor_rect
        if width < 10 or height < 10:
            return None, "Invalid monitor rect"
    elif hwnd is not None:
        rect = get_window_rect(hwnd)
        if not rect:
            return None, "Window no longer valid"
        left, top, width, height = rect
    else:
        rect = get_window_rect_by_title(window_title_substring)
        if not rect:
            return None, f"No window found containing '{window_title_substring}'"
        left, top, width, height = rect
    img = capture_region(left, top, width, height)
    if img is None:
        return None, "Failed to capture window"

    pixels = sample_band(img, x_start_ratio=x_start, x_end_ratio=x_end, y_start_ratio=y_start, y_end_ratio=y_end, step=sample_step)
    if not pixels:
        return None, "No pixels in band"

    color = check_band_for_alert_colors(pixels, require_count=require_pixels, color_ranges=color_ranges)
    return color, None


def run_loop(
    window_title_substring: str = "EVE",
    hwnd=None,
    monitor_rect=None,
    interval_seconds: float = 1.0,
    x_start: float = 0.85,
    x_end: float = 0.98,
    y_start: float = 0.0,
    y_end: float = 0.5,
    sample_step: int = 2,
    require_pixels: int = 8,
    color_ranges: dict | None = None,
    on_alert=None,
    stop_event: threading.Event | None = None,
    blink_only: bool = False,
    sound_once: bool = False,
    should_play_sound=None,
):
    """
    Run the check every interval_seconds. On detection, play alarm (if should_play_sound()) and call on_alert(color).
    If stop_event is set, stop the loop when it is set.
    If blink_only is True, only alert when a color appears after it was absent.
    If sound_once is True, only beep once per detection (no repeat while color remains).
    should_play_sound: callable returning bool; if given, play_alarm only when it returns True (e.g. for mute).
    """
    if stop_event is None:
        stop_event = threading.Event()
    if should_play_sound is None:
        should_play_sound = lambda: True
    last_color = None
    already_beeped_for = None  # when sound_once: don't beep again until color changes
    while not stop_event.is_set():
        color, err = run_once(
            window_title_substring=window_title_substring,
            hwnd=hwnd,
            monitor_rect=monitor_rect,
            x_start=x_start,
            x_end=x_end,
            y_start=y_start,
            y_end=y_end,
            sample_step=sample_step,
            require_pixels=require_pixels,
            color_ranges=color_ranges,
        )
        if err and not color:
            pass
        if color:
            do_beep = should_play_sound()
            if sound_once:
                if color != already_beeped_for and do_beep:
                    play_alarm()
                    already_beeped_for = color
                if on_alert:
                    on_alert(color)
            else:
                if (not blink_only or color != last_color) and do_beep:
                    play_alarm()
                if on_alert:
                    on_alert(color)
            last_color = color
        else:
            last_color = None
            already_beeped_for = None
        stop_event.wait(interval_seconds)


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

def run_gui():
    """GUI with Monitor tab (source, start/stop, captures, sound) and Settings tab (band, threshold, RGB ranges)."""
    root = tk.Tk()
    root.title("Overview Alert")
    root.minsize(420, 420)
    root.geometry("580x520")

    # State
    stop_event = threading.Event()
    monitor_thread = None
    mute_until = [0.0]  # use list so inner fn can rebind
    sound_once_var = tk.BooleanVar(value=True)
    source_var = tk.StringVar(value="screen")  # "screen" or "window"
    monitors_list = []  # list of monitor dicts after refresh
    windows_data = []  # list of (hwnd, title) after refresh

    # Settings vars (used by Monitor tab when starting/saving)
    x_start_var = tk.IntVar(value=85)
    x_end_var = tk.IntVar(value=98)
    y_start_var = tk.IntVar(value=0)
    y_end_var = tk.IntVar(value=50)
    require_pixels_var = tk.IntVar(value=8)
    # RGB range vars: (r_min, r_max, g_min, g_max, b_min, b_max) per color
    def _default_ranges():
        d = DEFAULT_COLOR_RANGES
        return {
            "teal": (tk.IntVar(value=d["teal"][0]), tk.IntVar(value=d["teal"][1]), tk.IntVar(value=d["teal"][2]), tk.IntVar(value=d["teal"][3]), tk.IntVar(value=d["teal"][4]), tk.IntVar(value=d["teal"][5])),
            "yellow": (tk.IntVar(value=d["yellow"][0]), tk.IntVar(value=d["yellow"][1]), tk.IntVar(value=d["yellow"][2]), tk.IntVar(value=d["yellow"][3]), tk.IntVar(value=d["yellow"][4]), tk.IntVar(value=d["yellow"][5])),
            "red": (tk.IntVar(value=d["red"][0]), tk.IntVar(value=d["red"][1]), tk.IntVar(value=d["red"][2]), tk.IntVar(value=d["red"][3]), tk.IntVar(value=d["red"][4]), tk.IntVar(value=d["red"][5])),
            "purple": (tk.IntVar(value=d["purple"][0]), tk.IntVar(value=d["purple"][1]), tk.IntVar(value=d["purple"][2]), tk.IntVar(value=d["purple"][3]), tk.IntVar(value=d["purple"][4]), tk.IntVar(value=d["purple"][5])),
        }
    rgb_vars = _default_ranges()

    notebook = ttk.Notebook(root)
    notebook.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

    # ---- Monitor tab ----
    monitor_tab = ttk.Frame(notebook, padding=6)
    notebook.add(monitor_tab, text="Monitor")

    status_var = tk.StringVar(value="Select a screen or window, then Start.")
    ttk.Label(monitor_tab, textvariable=status_var).pack(anchor=tk.W, padx=4, pady=2)

    # Source: Screen (monitor) or Window
    source_frame = ttk.LabelFrame(monitor_tab, text="Capture from", padding=8)
    source_frame.pack(fill=tk.X, padx=4, pady=6)
    ttk.Radiobutton(source_frame, text="Screen (full monitor)", variable=source_var, value="screen", command=lambda: _toggle_source()).pack(anchor=tk.W)
    ttk.Radiobutton(source_frame, text="Window (select from list)", variable=source_var, value="window", command=lambda: _toggle_source()).pack(anchor=tk.W)

    # Monitor selector (when source = screen)
    mon_frame = ttk.Frame(source_frame)
    mon_frame.pack(fill=tk.X, pady=4)
    ttk.Label(mon_frame, text="Monitor:").pack(side=tk.LEFT, padx=(0, 4))
    monitor_combo = ttk.Combobox(mon_frame, state="readonly", width=42)
    monitor_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)

    def _fill_monitors():
        nonlocal monitors_list
        monitors_list = get_monitors()
        opts = []
        for i, m in enumerate(monitors_list):
            label = f"Monitor {i + 1}"
            if m.get("is_primary"):
                label += " (Primary)"
            label += f" - {m['width']}x{m['height']}"
            opts.append(label)
        monitor_combo["values"] = opts
        if opts:
            monitor_combo.current(0)

    def _toggle_source():
        if source_var.get() == "screen":
            list_frame.pack_forget()
            mon_frame.pack(fill=tk.X, pady=4)
            _fill_monitors()
        else:
            mon_frame.pack_forget()
            list_frame.pack(fill=tk.BOTH, expand=True, padx=4, pady=6)
            refresh_list()

    def _get_band_ratios():
        """Return (x_start, x_end, y_start, y_end) as fractions in 0–1, clamped and validated."""
        xs = max(0, min(100, x_start_var.get())) / 100.0
        xe = max(0, min(100, x_end_var.get())) / 100.0
        ys = max(0, min(100, y_start_var.get())) / 100.0
        ye = max(0, min(100, y_end_var.get())) / 100.0
        if xs >= xe:
            xe = min(1.0, xs + 0.01)
        if ys >= ye:
            ye = min(1.0, ys + 0.01)
        return xs, xe, ys, ye

    def _get_require_pixels():
        return max(1, min(100, require_pixels_var.get()))

    def _get_color_ranges():
        out = {}
        for name, vars_tuple in rgb_vars.items():
            v = vars_tuple
            out[name] = (v[0].get(), v[1].get(), v[2].get(), v[3].get(), v[4].get(), v[5].get())
        return out

    # Capture buttons (always visible; work for both screen and window mode)
    capture_btn_frame = ttk.LabelFrame(monitor_tab, text="Save captures (to overview_alert_captures/)", padding=8)
    capture_btn_frame.pack(fill=tk.X, padx=4, pady=6)
    capture_btn_row = ttk.Frame(capture_btn_frame)
    capture_btn_row.pack(fill=tk.X)
    ttk.Label(capture_btn_frame, text="Screenshot = full capture; debug = full + band + report", font=("", 8)).pack(anchor=tk.W)

    # Window list (when source = window)
    list_frame = ttk.LabelFrame(monitor_tab, text="Select window to monitor", padding=8)
    listbox = tk.Listbox(list_frame, height=6, selectmode=tk.SINGLE, font=("Segoe UI", 9))
    scroll = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=listbox.yview)
    listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scroll.pack(side=tk.RIGHT, fill=tk.Y)
    listbox.configure(yscrollcommand=scroll.set)

    def refresh_list():
        nonlocal windows_data
        listbox.delete(0, tk.END)
        windows_data = list_all_visible_windows()
        for hwnd, title in windows_data:
            display = title[:80] + ("..." if len(title) > 80 else "")
            listbox.insert(tk.END, display)
        status_var.set(f"Found {len(windows_data)} window(s). Select one and click Start.")

    btn_refresh = ttk.Frame(list_frame)
    btn_refresh.pack(pady=4)
    ttk.Button(btn_refresh, text="Refresh list", command=refresh_list).pack(side=tk.LEFT)

    def do_save_debug_capture():
        if source_var.get() == "screen":
            if not monitors_list:
                status_var.set("No monitors found.")
                return
            idx = monitor_combo.current()
            if idx < 0 or idx >= len(monitors_list):
                status_var.set("Select a monitor first.")
                return
            m = monitors_list[idx]
            rect = (m["left"], m["top"], m["width"], m["height"])
            xs, xe, ys, ye = _get_band_ratios()
            ok, msg = save_debug_capture(monitor_rect=rect, x_start=xs, x_end=xe, y_start=ys, y_end=ye, sample_step=2, require_pixels=_get_require_pixels(), color_ranges=_get_color_ranges())
        else:
            sel = listbox.curselection()
            if not sel or not windows_data:
                status_var.set("Select a window first, then click Save debug capture.")
                return
            idx = int(sel[0])
            if idx >= len(windows_data):
                return
            hwnd, title = windows_data[idx]
            xs, xe, ys, ye = _get_band_ratios()
            ok, msg = save_debug_capture(hwnd=hwnd, x_start=xs, x_end=xe, y_start=ys, y_end=ye, sample_step=2, require_pixels=_get_require_pixels(), color_ranges=_get_color_ranges())
        if ok:
            status_var.set("Debug capture saved in overview_alert_captures/")
            root.clipboard_clear()
            root.clipboard_append(msg)
        else:
            status_var.set(msg)

    def do_take_screenshot():
        """Save current source (monitor or window) as a single image for calibration."""
        if not HAS_PIL:
            status_var.set("Pillow (PIL) is required for screenshots.")
            return
        if source_var.get() == "screen":
            if not monitors_list:
                status_var.set("No monitors found.")
                return
            idx = monitor_combo.current()
            if idx < 0 or idx >= len(monitors_list):
                status_var.set("Select a monitor first.")
                return
            m = monitors_list[idx]
            left, top, width, height = m["left"], m["top"], m["width"], m["height"]
        else:
            sel = listbox.curselection()
            if not sel or not windows_data:
                status_var.set("Select a window first.")
                return
            idx = int(sel[0])
            if idx >= len(windows_data):
                return
            hwnd, _ = windows_data[idx]
            rect = get_window_rect(hwnd)
            if not rect:
                status_var.set("Window no longer valid.")
                return
            left, top, width, height = rect
        img = capture_region(left, top, width, height)
        if img is None:
            status_var.set("Failed to capture.")
            return
        CAPTURES_DIR.mkdir(parents=True, exist_ok=True)
        n = _next_capture_number("overview_alert_screenshot", ".png")
        path = CAPTURES_DIR / f"overview_alert_screenshot_{n:03d}.png"
        try:
            img.save(str(path))
            status_var.set(f"Screenshot saved: {path.parent.name}/{path.name}")
        except Exception as e:
            status_var.set(f"Could not save: {e}")

    ttk.Button(capture_btn_row, text="Take screenshot", command=do_take_screenshot).pack(side=tk.LEFT, padx=(0, 8))
    ttk.Button(capture_btn_row, text="Save debug capture", command=do_save_debug_capture).pack(side=tk.LEFT)

    # Controls
    ctrl = ttk.Frame(monitor_tab)
    ctrl.pack(fill=tk.X, padx=4, pady=6)
    start_btn = ttk.Button(ctrl, text="Start", command=lambda: None)
    stop_btn = ttk.Button(ctrl, text="Stop", command=lambda: None, state=tk.DISABLED)

    def should_play_sound():
        return time.time() > mute_until[0]

    def do_start():
        nonlocal monitor_thread
        if source_var.get() == "screen":
            if not monitors_list:
                status_var.set("No monitors found. Try switching source or restart.")
                return
            idx = monitor_combo.current()
            if idx < 0 or idx >= len(monitors_list):
                status_var.set("Select a monitor first.")
                return
            m = monitors_list[idx]
            monitor_rect = (m["left"], m["top"], m["width"], m["height"])
            stop_event.clear()
            xs, xe, ys, ye = _get_band_ratios()
            status_var.set(f"Monitoring screen {idx + 1} ({m['width']}x{m['height']}).")
            def run():
                run_loop(
                    monitor_rect=monitor_rect,
                    interval_seconds=1.0,
                    x_start=xs,
                    x_end=xe,
                    y_start=ys,
                    y_end=ye,
                    sample_step=2,
                    require_pixels=_get_require_pixels(),
                    color_ranges=_get_color_ranges(),
                    on_alert=lambda c: root.after(0, lambda: status_var.set(f"Alert: {c}")),
                    stop_event=stop_event,
                    blink_only=False,
                    sound_once=sound_once_var.get(),
                    should_play_sound=should_play_sound,
                )
            monitor_thread = threading.Thread(target=run, daemon=True)
            monitor_thread.start()
            start_btn.configure(state=tk.DISABLED)
            stop_btn.configure(state=tk.NORMAL)
            return
        # Window mode
        sel = listbox.curselection()
        if not sel or not windows_data:
            status_var.set("Select a window from the list first (click Refresh if empty).")
            return
        idx = int(sel[0])
        if idx >= len(windows_data):
            return
        hwnd, title = windows_data[idx]
        stop_event.clear()
        xs, xe, ys, ye = _get_band_ratios()
        status_var.set(f"Monitoring: {title[:50]}...")

        def run():
            run_loop(
                hwnd=hwnd,
                interval_seconds=1.0,
                x_start=xs,
                x_end=xe,
                y_start=ys,
                y_end=ye,
                sample_step=2,
                require_pixels=_get_require_pixels(),
                color_ranges=_get_color_ranges(),
                on_alert=lambda c: root.after(0, lambda: status_var.set(f"Alert: {c}")),
                stop_event=stop_event,
                blink_only=False,
                sound_once=sound_once_var.get(),
                should_play_sound=should_play_sound,
            )

        monitor_thread = threading.Thread(target=run, daemon=True)
        monitor_thread.start()
        start_btn.configure(state=tk.DISABLED)
        stop_btn.configure(state=tk.NORMAL)

    def do_stop():
        stop_event.set()
        start_btn.configure(state=tk.NORMAL)
        stop_btn.configure(state=tk.DISABLED)
        status_var.set("Stopped. Select a window and Start again or close.")

    start_btn.configure(command=do_start)
    stop_btn.configure(command=do_stop)
    start_btn.pack(side=tk.LEFT, padx=2)
    stop_btn.pack(side=tk.LEFT, padx=2)

    # Sound options
    sound_frame = ttk.Frame(monitor_tab)
    sound_frame.pack(fill=tk.X, padx=4, pady=4)
    ttk.Checkbutton(
        sound_frame,
        text="Sound once per detection (no repeat while color is present)",
        variable=sound_once_var,
    ).pack(anchor=tk.W)
    btn_row = ttk.Frame(sound_frame)
    btn_row.pack(anchor=tk.W, pady=4)
    ttk.Button(btn_row, text="Play sound", command=play_alarm).pack(side=tk.LEFT, padx=(0, 8))

    def do_select_sound():
        path = filedialog.askopenfilename(
            title="Select alarm sound (WAV recommended)",
            filetypes=[("WAV files", "*.wav"), ("All files", "*.*")],
            initialdir=str(APP_DIR),
        )
        if not path:
            return
        dest = APP_DIR / ALARM_WAV_NAME
        try:
            shutil.copy2(path, dest)
            status_var.set(f"Sound saved as {ALARM_WAV_NAME} in program folder.")
        except Exception as e:
            status_var.set(f"Could not save sound: {e}")

    ttk.Button(btn_row, text="Select sound file...", command=do_select_sound).pack(side=tk.LEFT, padx=(0, 8))

    def do_mute():
        mute_until[0] = time.time() + 30
        status_var.set("Muted for 30 seconds.")

    ttk.Button(btn_row, text="Stop sound / Mute for 30 sec", command=do_mute).pack(side=tk.LEFT)

    # ---- Settings tab ----
    settings_tab = ttk.Frame(notebook, padding=8)
    notebook.add(settings_tab, text="Settings")

    band_frame = ttk.LabelFrame(settings_tab, text="Alert band (percent of width / height)", padding=8)
    band_frame.pack(fill=tk.X, pady=6)
    band_row = ttk.Frame(band_frame)
    band_row.pack(fill=tk.X)
    ttk.Label(band_row, text="X start %:").pack(side=tk.LEFT, padx=(0, 2))
    ttk.Spinbox(band_row, from_=0, to=100, width=5, textvariable=x_start_var).pack(side=tk.LEFT, padx=(0, 12))
    ttk.Label(band_row, text="X end %:").pack(side=tk.LEFT, padx=(0, 2))
    ttk.Spinbox(band_row, from_=0, to=100, width=5, textvariable=x_end_var).pack(side=tk.LEFT, padx=(0, 12))
    ttk.Label(band_row, text="Y start %:").pack(side=tk.LEFT, padx=(0, 2))
    ttk.Spinbox(band_row, from_=0, to=100, width=5, textvariable=y_start_var).pack(side=tk.LEFT, padx=(0, 12))
    ttk.Label(band_row, text="Y end %:").pack(side=tk.LEFT, padx=(0, 2))
    ttk.Spinbox(band_row, from_=0, to=100, width=5, textvariable=y_end_var).pack(side=tk.LEFT)

    thresh_frame = ttk.LabelFrame(settings_tab, text="Detection", padding=8)
    thresh_frame.pack(fill=tk.X, pady=6)
    ttk.Label(thresh_frame, text="Min pixels to trigger (same color):").pack(side=tk.LEFT, padx=(0, 8))
    ttk.Spinbox(thresh_frame, from_=1, to=100, width=5, textvariable=require_pixels_var).pack(side=tk.LEFT)
    ttk.Label(thresh_frame, text="(default 8)").pack(side=tk.LEFT, padx=(8, 0))

    rgb_frame = ttk.LabelFrame(settings_tab, text="RGB ranges (R min–max, G min–max, B min–max; 0–255). Purple = excluded.", padding=8)
    rgb_frame.pack(fill=tk.BOTH, expand=True, pady=6)
    for color_name in ("teal", "yellow", "red", "purple"):
        row = ttk.Frame(rgb_frame)
        row.pack(fill=tk.X, pady=2)
        ttk.Label(row, text=color_name.capitalize() + ":", width=8, anchor=tk.W).pack(side=tk.LEFT, padx=(0, 4))
        v = rgb_vars[color_name]
        for i, label in enumerate(("R min", "R max", "G min", "G max", "B min", "B max")):
            ttk.Label(row, text=label, font=("", 8)).pack(side=tk.LEFT, padx=(8, 2))
            ttk.Spinbox(row, from_=0, to=255, width=4, textvariable=v[i]).pack(side=tk.LEFT, padx=(0, 4))

    def on_closing():
        stop_event.set()
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_closing)
    _toggle_source()  # show screen selector and fill monitors (default)
    root.mainloop()


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    import argparse
    p = argparse.ArgumentParser(description="EVE Overview color alert (teal/yellow/red, no purple)")
    p.add_argument(
        "--gui",
        action="store_true",
        help="Launch GUI to pick window and start/stop (default if no other args)",
    )
    p.add_argument(
        "--window", "-w",
        default="EVE",
        help="Window title substring to monitor (default: EVE)",
    )
    p.add_argument(
        "--interval", "-i",
        type=float,
        default=1.0,
        help="Check interval in seconds (default: 1.0)",
    )
    p.add_argument(
        "--x-start",
        type=float,
        default=0.85,
        help="Band left edge as fraction of width (default: 0.85)",
    )
    p.add_argument(
        "--x-end",
        type=float,
        default=0.98,
        help="Band right edge as fraction of width (default: 0.98)",
    )
    p.add_argument("--y-start", type=float, default=0.0, help="Band top as fraction of height (default: 0)")
    p.add_argument("--y-end", type=float, default=0.5, help="Band bottom as fraction of height (default: 0.5 = middle)")
    p.add_argument(
        "--step",
        type=int,
        default=4,
        help="Pixel sample step (default: 4)",
    )
    p.add_argument(
        "--require",
        type=int,
        default=8,
        help="Min number of matching pixels to trigger (default: 8)",
    )
    p.add_argument(
        "--once",
        action="store_true",
        help="Run once and print result, then exit",
    )
    p.add_argument(
        "--blink-only",
        action="store_true",
        help="Only alert when color appears after being absent (fewer repeated beeps)",
    )
    args = p.parse_args()

    # No args or --gui: launch GUI. Otherwise use CLI.
    if len(sys.argv) == 1 or args.gui:
        try:
            run_gui()
            return 0
        except Exception as e:
            print(e, file=sys.stderr)
            return 1

    if args.once:
        color, err = run_once(
            window_title_substring=args.window,
            x_start=args.x_start,
            x_end=args.x_end,
            y_start=args.y_start,
            y_end=args.y_end,
            sample_step=args.step,
            require_pixels=args.require,
        )
        if err:
            print(err, file=sys.stderr)
        print("Detected:", color if color else "none")
        return 0 if color else 1

    print(f"Monitoring window containing '{args.window}' every {args.interval}s (band X {args.x_start*100:.0f}%–{args.x_end*100:.0f}%, Y {args.y_start*100:.0f}%–{args.y_end*100:.0f}%). Ctrl+C to stop.")
    try:
        run_loop(
            window_title_substring=args.window,
            interval_seconds=args.interval,
            x_start=args.x_start,
            x_end=args.x_end,
            y_start=args.y_start,
            y_end=args.y_end,
            sample_step=args.step,
            require_pixels=args.require,
            on_alert=lambda c: print(f"Alert: {c}") or None,
            blink_only=args.blink_only,
        )
    except KeyboardInterrupt:
        print("\nStopped.")
    return 0


if __name__ == "__main__":
    sys.exit(main() or 0)
