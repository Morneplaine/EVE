"""
EVE Online Overview Color Alert

Monitors a window (e.g. the Overview / ship list) and plays an alarm when
teal, yellow, or red color bands appear in a narrow band on the right side
(90%–95% of window width). Purple is explicitly ignored.

Run every 1 second. Can be run standalone (GUI or CLI) or imported.
"""

import ctypes
from ctypes import wintypes
import os
import time
import sys
import threading
import tkinter as tk
from tkinter import ttk

# Alarm sound: MP3 file played on detection (and via "Play sound" in GUI)
ALARM_SOUND_PATH = r"c:\Users\nicol\OneDrive\Pictures\alarm.mp3"

# Optional: pygame for in-process MP3 playback (no external player, no focus steal)
try:
    import pygame
    PYGAME_AVAILABLE = True
except ImportError:
    PYGAME_AVAILABLE = False

# Optional: Pillow for screen capture
try:
    from PIL import ImageGrab
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

# Windows API for window enumeration
user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

# ---------------------------------------------------------------------------
# Color detection (RGB). Tuned for EVE Overview highlights.
# Purple ~ #5B1C5B is excluded.
# ---------------------------------------------------------------------------

def _is_purple(r: int, g: int, b: int) -> bool:
    """Exclude purple (e.g. EVE overview selection ~ #5B1C5B)."""
    if g > 55:
        return False
    # Purple: R and B similar and relatively high, G low
    return 50 <= r <= 130 and 50 <= b <= 130


def is_teal(r: int, g: int, b: int) -> bool:
    """Teal/cyan: low R, high G and B, G and B similar."""
    if _is_purple(r, g, b):
        return False
    return r < 120 and g > 140 and b > 140 and abs(g - b) < 80


def is_yellow(r: int, g: int, b: int) -> bool:
    """Yellow/orange band: high R and G, low B (relaxed to catch game UI yellows)."""
    if _is_purple(r, g, b):
        return False
    return r > 160 and g > 140 and b < 165


def is_red(r: int, g: int, b: int) -> bool:
    """Red: high R, low G and B."""
    if _is_purple(r, g, b):
        return False
    return r > 180 and g < 100 and b < 100


def pixel_matches_alert_color(r: int, g: int, b: int) -> str | None:
    """Return 'teal'|'yellow'|'red' if pixel matches an alert color, else None."""
    if is_teal(r, g, b):
        return "teal"
    if is_yellow(r, g, b):
        return "yellow"
    if is_red(r, g, b):
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
# Screen capture and band sampling
# ---------------------------------------------------------------------------

def capture_region(left: int, top: int, width: int, height: int):
    """Capture screen region. Returns PIL Image or None."""
    if not HAS_PIL:
        return None
    bbox = (left, top, left + width, top + height)
    return ImageGrab.grab(bbox)


def sample_band(image, x_start_ratio: float = 0.75, x_end_ratio: float = 0.98, step: int = 2):
    """
    Sample pixels in a vertical band of the image (default 75%–98% of width, catches top-right bar).
    step: sample every N pixels in both directions to keep it fast.
    Returns list of (r, g, b) tuples.
    """
    w, h = image.size
    x0 = int(w * x_start_ratio)
    x1 = int(w * x_end_ratio)
    if x0 >= x1:
        x1 = min(x0 + 1, w)
    pixels = []
    for x in range(x0, x1, max(1, step)):
        for y in range(0, h, step):
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


def check_band_for_alert_colors(pixels, require_count: int = 3):
    """
    Check sampled pixels for teal, yellow, or red (excluding purple).
    If at least require_count pixels match the same alert color, return that color name.
    Otherwise return None.
    """
    counts = {"teal": 0, "yellow": 0, "red": 0}
    for r, g, b in pixels:
        color = pixel_matches_alert_color(r, g, b)
        if color:
            counts[color] += 1
    for color, count in counts.items():
        if count >= require_count:
            return color
    return None


def save_debug_capture(
    hwnd,
    x_start: float = 0.75,
    x_end: float = 0.98,
    sample_step: int = 2,
    require_pixels: int = 3,
    save_dir: str | None = None,
):
    """
    Capture the window and the sampled band, save as images, and write a debug report
    (pixel counts, sample RGBs) so you can see why a color might not be detected.
    Returns (success: bool, message: str).
    """
    if not HAS_PIL:
        return False, "Pillow (PIL) is required for capture."
    rect = get_window_rect(hwnd)
    if not rect:
        return False, "Window no longer valid."
    left, top, width, height = rect
    full = capture_region(left, top, width, height)
    if full is None:
        return False, "Failed to capture window."
    if save_dir is None:
        save_dir = os.path.dirname(os.path.abspath(__file__))
    os.makedirs(save_dir, exist_ok=True)
    full_path = os.path.join(save_dir, "overview_alert_full.png")
    band_path = os.path.join(save_dir, "overview_alert_band.png")
    report_path = os.path.join(save_dir, "overview_alert_debug.txt")
    try:
        full.save(full_path)
    except Exception as e:
        return False, f"Could not save full capture: {e}"
    w, h = full.size
    x0 = int(w * x_start)
    x1 = int(w * x_end)
    if x1 <= x0:
        x1 = x0 + 1
    band = full.crop((x0, 0, min(x1, w), h))
    try:
        band.save(band_path)
    except Exception as e:
        return False, f"Could not save band image: {e}"
    pixels = sample_band(full, x_start_ratio=x_start, x_end_ratio=x_end, step=sample_step)
    counts = {"teal": 0, "yellow": 0, "red": 0}
    for r, g, b in pixels:
        color = pixel_matches_alert_color(r, g, b)
        if color:
            counts[color] += 1
    detected = check_band_for_alert_colors(pixels, require_count=require_pixels)
    # Pixels that are "yellow-ish" (high R, high G) for tuning
    yellow_candidates = [(r, g, b) for r, g, b in pixels if r > 150 and g > 150 and b < 200]
    lines = [
        "Overview Alert debug report",
        "=" * 50,
        f"Full capture: {full_path}",
        f"Band only:    {band_path}",
        "",
        f"Window: {width}x{height}, band X: {x0}-{x1} ({(x_start*100):.0f}%-{(x_end*100):.0f}% of width)",
        f"Pixels sampled (step={sample_step}): {len(pixels)}",
        "",
        "Detection counts (current thresholds):",
        f"  teal:   {counts['teal']} (need >= {require_pixels} to trigger)",
        f"  yellow: {counts['yellow']} (need >= {require_pixels} to trigger)",
        f"  red:    {counts['red']} (need >= {require_pixels} to trigger)",
        f"Detected color: {detected or 'none'}",
        "",
        "Yellow threshold in code: R>160, G>140, B<165. If your yellow is dimmer, lower R/G or raise B.",
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
# Alarm (in-process via pygame: no external player, no delay, no focus steal)
# ---------------------------------------------------------------------------

_alarm_loaded = False


def init_alarm_sound():
    """Load the alarm MP3 once so play_alarm() is instant. Call at GUI startup."""
    global _alarm_loaded
    if _alarm_loaded or not PYGAME_AVAILABLE:
        return
    path = (os.path.expanduser(ALARM_SOUND_PATH) or "").strip()
    if not path or not os.path.isfile(path):
        return
    try:
        pygame.mixer.init(frequency=22050, size=-16, channels=2, buffer=512)
        pygame.mixer.music.load(path)
        _alarm_loaded = True
    except Exception:
        _alarm_loaded = False


def play_alarm():
    """Play the alarm sound in-process (no external player, no focus steal)."""
    global _alarm_loaded
    if not _alarm_loaded and PYGAME_AVAILABLE:
        init_alarm_sound()
    if PYGAME_AVAILABLE and _alarm_loaded:
        try:
            pygame.mixer.music.rewind()
            pygame.mixer.music.play()
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
    x_start: float = 0.75,
    x_end: float = 0.98,
    sample_step: int = 2,
    require_pixels: int = 3,
):
    """
    Perform one check: capture band 90–95% of the window, detect teal/yellow/red.
    If hwnd is given, use that window; else find by window_title_substring.
    Returns (detected_color or None, error_message or None).
    """
    if not HAS_PIL:
        return None, "Pillow (PIL) is required. Install with: pip install Pillow"

    if hwnd is not None:
        rect = get_window_rect(hwnd)
        if not rect:
            return None, "Window no longer valid"
    else:
        rect = get_window_rect_by_title(window_title_substring)
        if not rect:
            return None, f"No window found containing '{window_title_substring}'"

    left, top, width, height = rect
    img = capture_region(left, top, width, height)
    if img is None:
        return None, "Failed to capture window"

    pixels = sample_band(img, x_start_ratio=x_start, x_end_ratio=x_end, step=sample_step)
    if not pixels:
        return None, "No pixels in band"

    color = check_band_for_alert_colors(pixels, require_count=require_pixels)
    return color, None


def run_loop(
    window_title_substring: str = "EVE",
    hwnd=None,
    interval_seconds: float = 1.0,
    x_start: float = 0.75,
    x_end: float = 0.98,
    sample_step: int = 2,
    require_pixels: int = 3,
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
            x_start=x_start,
            x_end=x_end,
            sample_step=sample_step,
            require_pixels=require_pixels,
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
    """Small GUI: choose window from list, Start/Stop, sound once, mute button."""
    root = tk.Tk()
    root.title("Overview Alert")
    root.minsize(400, 320)
    root.geometry("480x360")

    # Preload alarm sound so playback is instant (no external player, no focus steal)
    init_alarm_sound()

    # State
    stop_event = threading.Event()
    monitor_thread = None
    mute_until = [0.0]  # use list so inner fn can rebind: mute_until[0] = time.time() + 30
    sound_once_var = tk.BooleanVar(value=True)

    # Window list
    list_frame = ttk.LabelFrame(root, text="Select window to monitor", padding=8)
    list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=6)
    listbox = tk.Listbox(list_frame, height=8, selectmode=tk.SINGLE, font=("Segoe UI", 9))
    scroll = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=listbox.yview)
    listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scroll.pack(side=tk.RIGHT, fill=tk.Y)
    listbox.configure(yscrollcommand=scroll.set)
    windows_data = []  # list of (hwnd, title) after refresh

    def refresh_list():
        nonlocal windows_data
        listbox.delete(0, tk.END)
        windows_data = list_all_visible_windows()
        for hwnd, title in windows_data:
            display = title[:80] + ("..." if len(title) > 80 else "")
            listbox.insert(tk.END, display)
        status_var.set(f"Found {len(windows_data)} window(s). Select one and click Start.")

    status_var = tk.StringVar(value="Click Refresh to list windows.")
    ttk.Label(root, textvariable=status_var).pack(anchor=tk.W, padx=10, pady=2)
    btn_refresh = ttk.Frame(list_frame)
    btn_refresh.pack(pady=4)
    ttk.Button(btn_refresh, text="Refresh list", command=refresh_list).pack(side=tk.LEFT, padx=(0, 8))

    def do_save_debug_capture():
        sel = listbox.curselection()
        if not sel or not windows_data:
            status_var.set("Select a window first, then click Save debug capture.")
            return
        idx = int(sel[0])
        if idx >= len(windows_data):
            return
        hwnd, title = windows_data[idx]
        ok, msg = save_debug_capture(hwnd, x_start=0.75, x_end=0.98, sample_step=2, require_pixels=3)
        if ok:
            status_var.set("Debug capture saved. Check overview_alert_full.png, overview_alert_band.png, overview_alert_debug.txt")
            root.clipboard_clear()
            root.clipboard_append(msg)
        else:
            status_var.set(msg)

    ttk.Button(btn_refresh, text="Save debug capture", command=do_save_debug_capture).pack(side=tk.LEFT)
    ttk.Label(btn_refresh, text="(saves full window + band + report in script folder)", font=("", 8)).pack(side=tk.LEFT, padx=8)

    # Controls
    ctrl = ttk.Frame(root)
    ctrl.pack(fill=tk.X, padx=10, pady=6)
    start_btn = ttk.Button(ctrl, text="Start", command=lambda: None)
    stop_btn = ttk.Button(ctrl, text="Stop", command=lambda: None, state=tk.DISABLED)

    def should_play_sound():
        return time.time() > mute_until[0]

    def do_start():
        sel = listbox.curselection()
        if not sel or not windows_data:
            status_var.set("Select a window from the list first (click Refresh if empty).")
            return
        idx = int(sel[0])
        if idx >= len(windows_data):
            return
        hwnd, title = windows_data[idx]
        stop_event.clear()
        status_var.set(f"Monitoring: {title[:50]}...")

        def run():
            run_loop(
                hwnd=hwnd,
                interval_seconds=1.0,
                x_start=0.75,
                x_end=0.98,
                sample_step=2,
                require_pixels=3,
                on_alert=lambda c: root.after(0, lambda: status_var.set(f"Alert: {c}")),
                stop_event=stop_event,
                blink_only=False,
                sound_once=sound_once_var.get(),
                should_play_sound=should_play_sound,
            )

        nonlocal monitor_thread
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
    sound_frame = ttk.Frame(root)
    sound_frame.pack(fill=tk.X, padx=10, pady=4)
    ttk.Checkbutton(
        sound_frame,
        text="Sound once per detection (no repeat while color is present)",
        variable=sound_once_var,
    ).pack(anchor=tk.W)
    btn_row = ttk.Frame(sound_frame)
    btn_row.pack(anchor=tk.W, pady=4)
    ttk.Button(btn_row, text="Play sound", command=play_alarm).pack(side=tk.LEFT, padx=(0, 8))

    def do_mute():
        mute_until[0] = time.time() + 30
        status_var.set("Muted for 30 seconds.")

    ttk.Button(btn_row, text="Stop sound / Mute for 30 sec", command=do_mute).pack(side=tk.LEFT)

    def on_closing():
        stop_event.set()
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_closing)
    refresh_list()
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
        default=0.75,
        help="Band left edge as fraction of width (default: 0.75)",
    )
    p.add_argument(
        "--x-end",
        type=float,
        default=0.98,
        help="Band right edge as fraction of width (default: 0.98)",
    )
    p.add_argument(
        "--step",
        type=int,
        default=4,
        help="Pixel sample step (default: 4)",
    )
    p.add_argument(
        "--require",
        type=int,
        default=3,
        help="Min number of matching pixels to trigger (default: 3)",
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
            sample_step=args.step,
            require_pixels=args.require,
        )
        if err:
            print(err, file=sys.stderr)
        print("Detected:", color if color else "none")
        return 0 if color else 1

    print(f"Monitoring window containing '{args.window}' every {args.interval}s (band {args.x_start*100:.0f}%–{args.x_end*100:.0f}%). Ctrl+C to stop.")
    try:
        run_loop(
            window_title_substring=args.window,
            interval_seconds=args.interval,
            x_start=args.x_start,
            x_end=args.x_end,
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
