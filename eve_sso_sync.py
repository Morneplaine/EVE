"""
EVE SSO (Single Sign-On) and ESI sync module.

Uses OAuth 2.0 with PKCE (desktop flow) to authenticate and fetch:
- Wallet transactions
- Wallet journal
- Character industry jobs (manufacturing, reprocessing, etc.)

Required: Create an SSO application at https://developers.eveonline.com/
with callback URL http://localhost:8765/callback/ and the scopes listed in
DEFAULT_SCOPES below.

Store client_id and client_secret in config or environment (e.g. EVE_SSO_CLIENT_ID,
EVE_SSO_CLIENT_SECRET) or in eve_sso_credentials.json (gitignored).
"""

import base64
import hashlib
import json
import logging
import secrets
import sqlite3
import threading
import time
import urllib.parse
from http.server import HTTPServer, BaseHTTPRequestHandler
from pathlib import Path

import requests

logger = logging.getLogger(__name__)

# EVE SSO v2
SSO_AUTHORIZE_URL = "https://login.eveonline.com/v2/oauth/authorize"
SSO_TOKEN_URL = "https://login.eveonline.com/v2/oauth/token"
SSO_VERIFY_URL = "https://login.eveonline.com/oauth/verify"
ESI_BASE = "https://esi.evetech.net/latest"

# Scopes enabled on the EVE developer-portal application.
# Must exactly match (or be a subset of) what the application is configured for,
# otherwise EVE SSO will reject the authorization request.
DEFAULT_SCOPES = " ".join([
    "publicData",
    "esi-skills.read_skills.v1",
    "esi-skills.read_skillqueue.v1",
    "esi-wallet.read_character_wallet.v1",
    "esi-wallet.read_corporation_wallet.v1",
    "esi-corporations.read_corporation_membership.v1",
    "esi-assets.read_assets.v1",
    "esi-ui.open_window.v1",
    "esi-ui.write_waypoint.v1",
    "esi-markets.structure_markets.v1",
    "esi-characters.read_agents_research.v1",
    "esi-industry.read_character_jobs.v1",
    "esi-markets.read_character_orders.v1",
    "esi-characters.read_blueprints.v1",
    "esi-corporations.track_members.v1",
    "esi-wallet.read_corporation_wallets.v1",
    "esi-assets.read_corporation_assets.v1",
    "esi-corporations.read_blueprints.v1",
    "esi-industry.read_corporation_jobs.v1",
    "esi-markets.read_corporation_orders.v1",
    "esi-corporations.read_container_logs.v1",
    "esi-industry.read_character_mining.v1",
    "esi-industry.read_corporation_mining.v1",
    "esi-corporations.read_projects.v1",
    "esi-corporations.read_freelance_jobs.v1",
    "esi-activities.read_character.v1",
    "esi-access.read_lists.v1",
])
CALLBACK_HOST = "localhost"
CALLBACK_PORT = 8765
CALLBACK_PATH = "/callback/"
REDIRECT_URI = f"http://{CALLBACK_HOST}:{CALLBACK_PORT}{CALLBACK_PATH}"


def _base64url_encode(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).rstrip(b"=").decode("ascii")


def make_pkce_pair():
    """Return (code_verifier, code_challenge) for PKCE."""
    verifier = secrets.token_urlsafe(32)
    if isinstance(verifier, bytes):
        verifier = verifier.decode("ascii")
    digest = hashlib.sha256(verifier.encode("utf-8")).digest()
    challenge = _base64url_encode(digest)
    return verifier, challenge


def get_authorize_url(client_id: str, state: str | None = None, scopes: str = DEFAULT_SCOPES) -> tuple[str, str]:
    """
    Build the SSO authorize URL and return (url, code_verifier).
    Open the URL in a browser; after login the callback server will receive the code.
    """
    verifier, challenge = make_pkce_pair()
    state = state or secrets.token_urlsafe(16)
    params = {
        "response_type": "code",
        "client_id": client_id,
        "redirect_uri": REDIRECT_URI,
        "scope": scopes,
        "code_challenge": challenge,
        "code_challenge_method": "S256",
        "state": state,
    }
    url = SSO_AUTHORIZE_URL + "?" + urllib.parse.urlencode(params)
    return url, verifier


def exchange_code_for_tokens(
    code: str,
    client_id: str,
    code_verifier: str,
    client_secret: str | None = None,
) -> dict:
    """
    Exchange the authorization code for access_token and refresh_token.
    Returns dict with access_token, expires_in, refresh_token, and optionally character info.
    """
    data = {
        "grant_type": "authorization_code",
        "client_id": client_id,
        "code": code,
        "code_verifier": code_verifier,
        "redirect_uri": REDIRECT_URI,
    }
    if client_secret:
        data["client_secret"] = client_secret
    resp = requests.post(
        SSO_TOKEN_URL,
        data=data,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        timeout=30,
    )
    resp.raise_for_status()
    out = resp.json()
    return out


def refresh_tokens(
    refresh_token: str,
    client_id: str,
    client_secret: str,
) -> dict:
    """Get a new access token using the refresh token."""
    data = {
        "grant_type": "refresh_token",
        "client_id": client_id,
        "client_secret": client_secret,
        "refresh_token": refresh_token,
    }
    resp = requests.post(
        SSO_TOKEN_URL,
        data=data,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()


def decode_jwt_payload(token: str) -> dict:
    """Decode JWT payload (middle part) without verification (SSO returns valid tokens)."""
    try:
        parts = token.split(".")
        if len(parts) != 3:
            return {}
        payload = parts[1]
        padding = 4 - len(payload) % 4
        if padding != 4:
            payload += "=" * padding
        decoded = base64.urlsafe_b64decode(payload)
        return json.loads(decoded)
    except Exception:
        return {}


def run_callback_server(timeout_seconds: int = 300) -> tuple[str | None, str | None]:
    """
    Start a local HTTP server to receive the SSO callback; returns (code, state) or (None, None) on timeout.
    """
    code_result = []
    state_result = []

    class CallbackHandler(BaseHTTPRequestHandler):
        def do_GET(self):
            parsed = urllib.parse.urlparse(self.path)
            if parsed.path.rstrip("/") != CALLBACK_PATH.rstrip("/"):
                self.send_response(404)
                self.end_headers()
                return
            qs = urllib.parse.parse_qs(parsed.query)
            code_result.append(qs.get("code", [None])[0])
            state_result.append(qs.get("state", [None])[0])
            self.send_response(200)
            self.send_header("Content-type", "text/html")
            self.end_headers()
            self.wfile.write(b"<html><body><p>Login successful. You can close this window.</p></body></html>")

        def log_message(self, format, *args):
            logger.debug(format, *args)

    server = HTTPServer((CALLBACK_HOST, CALLBACK_PORT), CallbackHandler)
    server.timeout = 1
    deadline = time.time() + timeout_seconds
    while time.time() < deadline and not code_result:
        server.handle_request()
    return (code_result[0] if code_result else None, state_result[0] if state_result else None)


def _ensure_sso_columns(conn: sqlite3.Connection):
    """Add columns to sso_character if missing (idempotent migration)."""
    cols = {row[1] for row in conn.execute("PRAGMA table_info(sso_character)").fetchall()}
    if "corporation_id" not in cols:
        conn.execute("ALTER TABLE sso_character ADD COLUMN corporation_id INTEGER")
    if "corporation_name" not in cols:
        conn.execute("ALTER TABLE sso_character ADD COLUMN corporation_name TEXT")
    if "last_synced_at" not in cols:
        conn.execute("ALTER TABLE sso_character ADD COLUMN last_synced_at TEXT")
    conn.commit()


def ensure_sso_tables(conn: sqlite3.Connection):
    """Create SSO/ESI tables if missing."""
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS sso_character (
            character_id INTEGER PRIMARY KEY,
            character_name TEXT,
            refresh_token TEXT,
            access_token TEXT,
            access_token_expires_at REAL,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS esi_wallet_transactions (
            character_id INTEGER NOT NULL,
            transaction_id BIGINT NOT NULL,
            date_utc TEXT NOT NULL,
            type_id INTEGER,
            quantity INTEGER,
            unit_price REAL,
            client_id INTEGER,
            location_id INTEGER,
            is_buy INTEGER,
            is_personal INTEGER,
            journal_ref_id BIGINT,
            synced_at TEXT DEFAULT CURRENT_TIMESTAMP,
            PRIMARY KEY (character_id, transaction_id)
        );
        CREATE TABLE IF NOT EXISTS esi_wallet_journal (
            character_id INTEGER NOT NULL,
            ref_id BIGINT NOT NULL,
            date_utc TEXT NOT NULL,
            ref_type TEXT,
            amount REAL,
            balance REAL,
            context_id_type TEXT,
            context_id BIGINT,
            description TEXT,
            first_party_id INTEGER,
            second_party_id INTEGER,
            reason TEXT,
            synced_at TEXT DEFAULT CURRENT_TIMESTAMP,
            PRIMARY KEY (character_id, ref_id)
        );
        CREATE TABLE IF NOT EXISTS esi_industry_jobs (
            character_id INTEGER NOT NULL,
            job_id INTEGER NOT NULL,
            activity_id INTEGER,
            blueprint_id BIGINT,
            blueprint_type_id INTEGER,
            blueprint_location_id BIGINT,
            output_location_id BIGINT,
            runs INTEGER,
            cost REAL,
            licensed_runs INTEGER,
            probability REAL,
            product_type_id INTEGER,
            status TEXT,
            duration INTEGER,
            start_date_utc TEXT,
            end_date_utc TEXT,
            completed_date_utc TEXT,
            facility_id BIGINT,
            installer_id INTEGER,
            synced_at TEXT DEFAULT CURRENT_TIMESTAMP,
            PRIMARY KEY (character_id, job_id)
        );
        CREATE TABLE IF NOT EXISTS esi_corporation_industry_jobs (
            corporation_id INTEGER NOT NULL,
            job_id INTEGER NOT NULL,
            activity_id INTEGER,
            blueprint_id BIGINT,
            blueprint_type_id INTEGER,
            blueprint_location_id BIGINT,
            output_location_id BIGINT,
            runs INTEGER,
            cost REAL,
            licensed_runs INTEGER,
            probability REAL,
            product_type_id INTEGER,
            status TEXT,
            duration INTEGER,
            start_date_utc TEXT,
            end_date_utc TEXT,
            completed_date_utc TEXT,
            facility_id BIGINT,
            installer_id INTEGER,
            location_id BIGINT,
            successful_runs INTEGER,
            synced_at TEXT DEFAULT CURRENT_TIMESTAMP,
            PRIMARY KEY (corporation_id, job_id)
        );
        CREATE INDEX IF NOT EXISTS idx_corp_jobs_status
            ON esi_corporation_industry_jobs(corporation_id, status);
        CREATE TABLE IF NOT EXISTS esi_corporation_wallet_transactions (
            corporation_id INTEGER NOT NULL,
            division INTEGER NOT NULL,
            transaction_id BIGINT NOT NULL,
            date_utc TEXT NOT NULL,
            type_id INTEGER,
            quantity INTEGER,
            unit_price REAL,
            client_id INTEGER,
            location_id INTEGER,
            is_buy INTEGER,
            is_personal INTEGER,
            journal_ref_id BIGINT,
            synced_at TEXT DEFAULT CURRENT_TIMESTAMP,
            PRIMARY KEY (corporation_id, division, transaction_id)
        );
        CREATE INDEX IF NOT EXISTS idx_corp_wallet_tx_corp_tid
            ON esi_corporation_wallet_transactions(corporation_id, transaction_id);
    """)
    _ensure_sso_columns(conn)
    conn.commit()


def esi_get_public(path: str, params: dict | None = None) -> list | dict:
    """GET an unauthenticated ESI endpoint (e.g. /characters/{id}/, /corporations/{id}/)."""
    url = ESI_BASE.rstrip("/") + "/" + path.lstrip("/")
    r = requests.get(url, params=params or {}, timeout=30)
    r.raise_for_status()
    return r.json()


def fetch_character_public(character_id: int) -> dict:
    """Public character info: name, corporation_id, etc."""
    return esi_get_public(f"characters/{character_id}/")


def fetch_corporation_public(corporation_id: int) -> dict:
    """Public corporation info: name, ticker, etc."""
    return esi_get_public(f"corporations/{corporation_id}/")


def list_sso_characters(conn: sqlite3.Connection) -> list[dict]:
    """Return all SSO-linked characters with their metadata."""
    ensure_sso_tables(conn)
    rows = conn.execute(
        """
        SELECT character_id, character_name, corporation_id, corporation_name,
               access_token_expires_at, updated_at, last_synced_at
        FROM sso_character
        ORDER BY character_name COLLATE NOCASE
        """
    ).fetchall()
    out = []
    for r in rows:
        out.append({
            "character_id": r[0],
            "character_name": r[1] or "",
            "corporation_id": r[2],
            "corporation_name": r[3] or "",
            "access_token_expires_at": r[4],
            "updated_at": r[5],
            "last_synced_at": r[6],
        })
    return out


def delete_sso_character(conn: sqlite3.Connection, character_id: int) -> None:
    """Remove an SSO-linked character and all of its synced rows."""
    ensure_sso_tables(conn)
    conn.execute("DELETE FROM sso_character WHERE character_id = ?", (character_id,))
    conn.execute("DELETE FROM esi_wallet_transactions WHERE character_id = ?", (character_id,))
    conn.execute("DELETE FROM esi_wallet_journal WHERE character_id = ?", (character_id,))
    conn.execute("DELETE FROM esi_industry_jobs WHERE character_id = ?", (character_id,))
    conn.commit()


def mark_synced(conn: sqlite3.Connection, character_id: int) -> None:
    """Update last_synced_at on a character row."""
    ensure_sso_tables(conn)
    conn.execute(
        "UPDATE sso_character SET last_synced_at = CURRENT_TIMESTAMP WHERE character_id = ?",
        (character_id,),
    )
    conn.commit()


def get_valid_access_token(conn: sqlite3.Connection, character_id: int, client_id: str, client_secret: str) -> str | None:
    """Return a valid access token for the character, refreshing if needed."""
    row = conn.execute(
        "SELECT access_token, access_token_expires_at, refresh_token FROM sso_character WHERE character_id = ?",
        (character_id,),
    ).fetchone()
    if not row:
        return None
    access_token, expires_at, refresh_token = row
    now = time.time()
    if expires_at and now < expires_at - 60:
        return access_token
    if not refresh_token:
        return None
    try:
        data = refresh_tokens(refresh_token, client_id, client_secret)
    except Exception as e:
        logger.warning("Token refresh failed for character %s: %s", character_id, e)
        return None
    new_access = data.get("access_token")
    expires_in = data.get("expires_in", 1200)
    new_refresh = data.get("refresh_token") or refresh_token
    new_expires = now + expires_in
    conn.execute(
        "UPDATE sso_character SET access_token = ?, access_token_expires_at = ?, refresh_token = ?, updated_at = CURRENT_TIMESTAMP WHERE character_id = ?",
        (new_access, new_expires, new_refresh, character_id),
    )
    conn.commit()
    return new_access


def esi_get(path: str, access_token: str, params: dict | None = None) -> list | dict:
    """GET an ESI authenticated endpoint; path is e.g. /characters/123/wallet/transactions/."""
    url = ESI_BASE.rstrip("/") + "/" + path.lstrip("/")
    headers = {"Authorization": f"Bearer {access_token}"}
    r = requests.get(url, params=params or {}, headers=headers, timeout=30)
    r.raise_for_status()
    return r.json()


def fetch_wallet_transactions(character_id: int, access_token: str, from_id: int | None = None) -> list:
    """Fetch wallet transactions (last 365 days). Paginate with from_id if needed."""
    path = f"characters/{character_id}/wallet/transactions/"
    params = {}
    if from_id is not None:
        params["from_id"] = from_id
    return esi_get(path, access_token, params)


def fetch_corporation_wallet_transactions(
    corporation_id: int, division: int, access_token: str, from_id: int | None = None
) -> list:
    """Corp wallet division 1–7 market transactions (requires wallet role + corp scope)."""
    path = f"corporations/{corporation_id}/wallets/{division}/transactions/"
    params = {}
    if from_id is not None:
        params["from_id"] = from_id
    return esi_get(path, access_token, params)


def fetch_wallet_journal(character_id: int, access_token: str, page: int = 1) -> list:
    """Fetch one page of wallet journal (30 days)."""
    path = f"characters/{character_id}/wallet/journal/"
    return esi_get(path, access_token, {"page": page})


def fetch_character_skills(character_id: int, access_token: str) -> dict:
    """Trained skills. Requires scope esi-skills.read_skills.v1.

    Returns {skills: [{skill_id, trained_skill_level, active_skill_level,
    skillpoints_in_skill}], total_sp, unallocated_sp}.
    """
    return esi_get(f"characters/{character_id}/skills/", access_token)


def fetch_skill_queue(character_id: int, access_token: str) -> list:
    """Character skill queue. Requires scope esi-skills.read_skillqueue.v1.

    Each entry: skill_id, finished_level, queue_position, level_start_sp,
    level_end_sp, training_start_sp, start_date, finish_date.
    """
    return esi_get(f"characters/{character_id}/skillqueue/", access_token)


def fetch_type_attributes(type_id: int) -> dict:
    """Public type info incl. dogma_attributes (primary=180, secondary=181)."""
    return esi_get_public(f"universe/types/{type_id}/", {"datasource": "tranquility"})


def fetch_industry_jobs(character_id: int, access_token: str, include_completed: bool = True, page: int = 1) -> list:
    """Fetch character industry jobs (manufacturing, etc.). include_completed: last 90 days."""
    path = f"characters/{character_id}/industry/jobs/"
    params = {"include_completed": "true" if include_completed else "false", "page": page}
    return esi_get(path, access_token, params)


def fetch_corporation_industry_jobs(
    corporation_id: int, access_token: str, include_completed: bool = True, page: int = 1
) -> list:
    """Fetch corporation industry jobs (requires Factory_Manager role + esi-industry.read_corporation_jobs.v1)."""
    path = f"corporations/{corporation_id}/industry/jobs/"
    params = {"include_completed": "true" if include_completed else "false", "page": page}
    return esi_get(path, access_token, params)


def sync_character(
    conn: sqlite3.Connection,
    character_id: int,
    access_token: str,
    refresh_token: str,
    expires_in: int = 1200,
    character_name: str | None = None,
    corporation_id: int | None = None,
    corporation_name: str | None = None,
):
    """Upsert character and tokens into sso_character (with optional corp info)."""
    ensure_sso_tables(conn)
    expires_at = time.time() + expires_in
    conn.execute(
        """
        INSERT INTO sso_character (character_id, character_name, refresh_token, access_token,
                                   access_token_expires_at, corporation_id, corporation_name, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
        ON CONFLICT(character_id) DO UPDATE SET
            character_name = COALESCE(excluded.character_name, character_name),
            refresh_token = excluded.refresh_token,
            access_token = excluded.access_token,
            access_token_expires_at = excluded.access_token_expires_at,
            corporation_id = COALESCE(excluded.corporation_id, corporation_id),
            corporation_name = COALESCE(excluded.corporation_name, corporation_name),
            updated_at = CURRENT_TIMESTAMP
        """,
        (
            character_id,
            character_name or "",
            refresh_token,
            access_token,
            expires_at,
            corporation_id,
            corporation_name,
        ),
    )
    conn.commit()


def sync_wallet_transactions(conn: sqlite3.Connection, character_id: int, access_token: str) -> int:
    """Fetch and store wallet transactions; returns count of new/updated rows."""
    ensure_sso_tables(conn)
    all_tx = []
    from_id = None
    while True:
        page = fetch_wallet_transactions(character_id, access_token, from_id=from_id)
        if not page:
            break
        all_tx.extend(page)
        from_id = min(t["transaction_id"] for t in page)
        if len(page) < 1000:
            break
    count = 0
    for t in all_tx:
        conn.execute(
            """
            INSERT OR REPLACE INTO esi_wallet_transactions
            (character_id, transaction_id, date_utc, type_id, quantity, unit_price, client_id, location_id, is_buy, is_personal, journal_ref_id, synced_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
            """,
            (
                character_id,
                t["transaction_id"],
                t["date"],
                t.get("type_id"),
                t.get("quantity"),
                t.get("unit_price"),
                t.get("client_id"),
                t.get("location_id"),
                1 if t.get("is_buy") else 0,
                1 if t.get("is_personal") else 0,
                t.get("journal_ref_id"),
            ),
        )
        count += 1
    conn.commit()
    return count


def sync_wallet_journal(conn: sqlite3.Connection, character_id: int, access_token: str, pages: int = 10) -> int:
    """Fetch and store wallet journal (multiple pages); returns count of new/updated rows."""
    ensure_sso_tables(conn)
    count = 0
    for page_num in range(1, pages + 1):
        try:
            page = fetch_wallet_journal(character_id, access_token, page=page_num)
        except requests.HTTPError as e:
            if e.response.status_code == 404 or page_num > 1:
                break
            raise
        if not page:
            break
        for j in page:
            conn.execute(
                """
                INSERT OR REPLACE INTO esi_wallet_journal
                (character_id, ref_id, date_utc, ref_type, amount, balance, context_id_type, context_id, description, first_party_id, second_party_id, reason, synced_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
                """,
                (
                    character_id,
                    j["id"],
                    j["date"],
                    j.get("ref_type"),
                    j.get("amount"),
                    j.get("balance"),
                    j.get("context_id_type"),
                    j.get("context_id"),
                    j.get("description", ""),
                    j.get("first_party_id"),
                    j.get("second_party_id"),
                    j.get("reason", ""),
                ),
            )
            count += 1
    conn.commit()
    return count


def sync_corporation_industry_jobs(
    conn: sqlite3.Connection, corporation_id: int, access_token: str, include_completed: bool = True
) -> int:
    """Fetch and store corporation industry jobs; returns count of rows inserted/updated."""
    ensure_sso_tables(conn)
    count = 0
    page = 1
    while True:
        try:
            jobs = fetch_corporation_industry_jobs(
                corporation_id, access_token, include_completed=include_completed, page=page
            )
        except requests.HTTPError as e:
            if page == 1:
                raise
            break
        if not jobs:
            break
        for j in jobs:
            conn.execute(
                """
                INSERT OR REPLACE INTO esi_corporation_industry_jobs
                (corporation_id, job_id, activity_id, blueprint_id, blueprint_type_id, blueprint_location_id,
                 output_location_id, runs, cost, licensed_runs, probability, product_type_id, status, duration,
                 start_date_utc, end_date_utc, completed_date_utc, facility_id, installer_id, location_id,
                 successful_runs, synced_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
                """,
                (
                    corporation_id, j["job_id"], j.get("activity_id"), j.get("blueprint_id"),
                    j.get("blueprint_type_id"), j.get("blueprint_location_id"), j.get("output_location_id"),
                    j.get("runs"), j.get("cost"), j.get("licensed_runs"), j.get("probability"),
                    j.get("product_type_id"), j.get("status"), j.get("duration"),
                    j.get("start_date"), j.get("end_date"), j.get("completed_date"),
                    j.get("facility_id"), j.get("installer_id"), j.get("location_id"),
                    j.get("successful_runs"),
                ),
            )
            count += 1
        if len(jobs) < 50:
            break
        page += 1
    conn.commit()
    return count


def sync_corporation_wallet_transactions(
    conn: sqlite3.Connection, corporation_id: int, access_token: str
) -> tuple[int, str | None]:
    """Fetch and store corp wallet market transactions for divisions 1–7. Returns (row_count, optional_note)."""
    ensure_sso_tables(conn)
    total = 0
    skipped: list[int] = []
    for division in range(1, 8):
        from_id = None
        while True:
            try:
                page = fetch_corporation_wallet_transactions(
                    corporation_id, division, access_token, from_id=from_id
                )
            except requests.HTTPError as e:
                status = e.response.status_code if e.response is not None else None
                if from_id is None and status in (401, 403, 404):
                    if status in (401, 403):
                        skipped.append(division)
                    break
                if from_id is not None and status in (400, 404):
                    break
                raise
            if not page:
                break
            for t in page:
                conn.execute(
                    """
                    INSERT OR REPLACE INTO esi_corporation_wallet_transactions
                    (corporation_id, division, transaction_id, date_utc, type_id, quantity, unit_price,
                     client_id, location_id, is_buy, is_personal, journal_ref_id, synced_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
                    """,
                    (
                        corporation_id,
                        division,
                        t["transaction_id"],
                        t["date"],
                        t.get("type_id"),
                        t.get("quantity"),
                        t.get("unit_price"),
                        t.get("client_id"),
                        t.get("location_id"),
                        1 if t.get("is_buy") else 0,
                        1 if t.get("is_personal") else 0,
                        t.get("journal_ref_id"),
                    ),
                )
                total += 1
            from_id = min(x["transaction_id"] for x in page)
            if len(page) < 1000:
                break
    conn.commit()
    note = None
    if skipped:
        note = (
            f"Some corp wallet divisions unreadable (HTTP 403/401): {skipped}. "
            "In-game Director or Accountant (wallet access) plus esi-wallet.read_corporation_wallets.v1 are required."
        )
    return total, note


def sync_industry_jobs(conn: sqlite3.Connection, character_id: int, access_token: str, include_completed: bool = True) -> int:
    """Fetch and store industry jobs; returns count of new/updated rows."""
    ensure_sso_tables(conn)
    count = 0
    page = 1
    while True:
        try:
            jobs = fetch_industry_jobs(character_id, access_token, include_completed=include_completed, page=page)
        except requests.HTTPError as e:
            if page == 1:
                raise
            break
        if not jobs:
            break
        for j in jobs:
            conn.execute(
                """
                INSERT OR REPLACE INTO esi_industry_jobs
                (character_id, job_id, activity_id, blueprint_id, blueprint_type_id, blueprint_location_id, output_location_id,
                 runs, cost, licensed_runs, probability, product_type_id, status, duration, start_date_utc, end_date_utc,
                 completed_date_utc, facility_id, installer_id, synced_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
                """,
                (
                    character_id,
                    j["job_id"],
                    j.get("activity_id"),
                    j.get("blueprint_id"),
                    j.get("blueprint_type_id"),
                    j.get("blueprint_location_id"),
                    j.get("output_location_id"),
                    j.get("runs"),
                    j.get("cost"),
                    j.get("licensed_runs"),
                    j.get("probability"),
                    j.get("product_type_id"),
                    j.get("status"),
                    j.get("duration"),
                    j.get("start_date"),
                    j.get("end_date"),
                    j.get("completed_date"),
                    j.get("facility_id"),
                    j.get("installer_id"),
                ),
            )
            count += 1
        if len(jobs) < 50:
            break
        page += 1
    conn.commit()
    return count


def run_full_sync(
    conn: sqlite3.Connection,
    character_id: int,
    client_id: str,
    client_secret: str,
) -> dict:
    """
    Ensure valid token, then sync wallet transactions, journal, industry jobs,
    and (when in a corporation) corporation industry jobs and corporation wallet
    market transactions (divisions 1–7).
    Returns dict with counts and any error message.
    """
    ensure_sso_tables(conn)
    access = get_valid_access_token(conn, character_id, client_id, client_secret)
    if not access:
        return {
            "error": "No valid token. Please log in with EVE SSO first.",
            "tx": 0,
            "journal": 0,
            "jobs": 0,
            "corp_jobs": 0,
            "corp_wallet_tx": 0,
        }
    out = {"tx": 0, "journal": 0, "jobs": 0, "corp_jobs": 0, "corp_wallet_tx": 0}
    try:
        corp_id: int | None = None
        try:
            pub = fetch_character_public(character_id)
            cid_pub = pub.get("corporation_id")
            corp_id = int(cid_pub) if cid_pub is not None else None
            corp_nm = ""
            if corp_id is not None:
                try:
                    corp_nm = fetch_corporation_public(corp_id).get("name") or ""
                except Exception:
                    logger.warning("Could not fetch corporation public info for %s", corp_id)
            ch_nm = pub.get("name") or ""
            conn.execute(
                """
                UPDATE sso_character SET
                    character_name = CASE WHEN ? != '' THEN ? ELSE character_name END,
                    corporation_id = ?,
                    corporation_name = CASE WHEN ? != '' THEN ? ELSE corporation_name END,
                    updated_at = CURRENT_TIMESTAMP
                WHERE character_id = ?
                """,
                (ch_nm, ch_nm, corp_id, corp_nm, corp_nm, int(character_id)),
            )
            conn.commit()
        except Exception as e:
            logger.warning("Public character/corp refresh failed for %s: %s", character_id, e)
            row = conn.execute(
                "SELECT corporation_id FROM sso_character WHERE character_id = ?", (int(character_id),)
            ).fetchone()
            corp_id = int(row[0]) if row and row[0] else None

        out["tx"] = sync_wallet_transactions(conn, character_id, access)
        out["journal"] = sync_wallet_journal(conn, character_id, access)
        out["jobs"] = sync_industry_jobs(conn, character_id, access)
        if corp_id is None:
            row = conn.execute(
                "SELECT corporation_id FROM sso_character WHERE character_id = ?", (int(character_id),)
            ).fetchone()
            corp_id = int(row[0]) if row and row[0] else None
        if corp_id:
            try:
                out["corp_jobs"] = sync_corporation_industry_jobs(conn, corp_id, access)
                out["corporation_id"] = corp_id
            except requests.HTTPError as e:
                status = e.response.status_code if e.response is not None else None
                if status in (403, 401):
                    out["corp_jobs_note"] = (
                        f"Corp jobs not accessible (HTTP {status}). "
                        "Character needs Factory_Manager role + esi-industry.read_corporation_jobs.v1 scope."
                    )
                else:
                    raise
            try:
                n_cw, cw_note = sync_corporation_wallet_transactions(conn, corp_id, access)
                out["corp_wallet_tx"] = n_cw
                if cw_note:
                    out["corp_wallet_note"] = cw_note
            except requests.HTTPError as e:
                status = e.response.status_code if e.response is not None else None
                if status in (403, 401):
                    out["corp_wallet_note"] = (
                        f"Corp wallet transactions not accessible (HTTP {status}). "
                        "Need wallet role (e.g. Director/Accountant) + esi-wallet.read_corporation_wallets.v1."
                    )
                else:
                    raise
        mark_synced(conn, character_id)
    except Exception as e:
        out["error"] = str(e)
        logger.exception("Sync failed")
    return out


def login_flow(client_id: str, client_secret: str, db_path: str = "eve_manufacturing.db") -> dict:
    """
    Run the desktop SSO flow: open browser, start callback server, exchange code for tokens,
    decode JWT to get character_id, store in DB, return character info.
    Returns dict with character_id, character_name, and optional error.
    """
    import webbrowser
    url, code_verifier = get_authorize_url(client_id)
    webbrowser.open(url)
    code, state = run_callback_server(timeout_seconds=300)
    if not code:
        return {"error": "No authorization code received (timeout or user cancelled)."}
    try:
        data = exchange_code_for_tokens(code, client_id, code_verifier, client_secret)
    except requests.HTTPError as e:
        return {"error": f"Token exchange failed: {e.response.status_code} {e.response.text}"}
    access = data.get("access_token")
    refresh = data.get("refresh_token")
    expires_in = data.get("expires_in", 1200)
    if not access or not refresh:
        return {"error": "Missing access or refresh token in response."}
    payload = decode_jwt_payload(access)
    sub = payload.get("sub", "")
    # sub is like "CHARACTER:EVE:12345"
    if ":" in sub:
        character_id = int(sub.split(":")[-1])
    else:
        character_id = int(sub)
    character_name = payload.get("name") or ""
    corporation_id = None
    corporation_name = None
    try:
        char_info = fetch_character_public(character_id)
        corporation_id = char_info.get("corporation_id")
        if not character_name:
            character_name = char_info.get("name") or ""
        if corporation_id:
            try:
                corp_info = fetch_corporation_public(corporation_id)
                corporation_name = corp_info.get("name") or ""
            except Exception:
                logger.warning("Could not fetch corporation %s public info", corporation_id)
    except Exception:
        logger.warning("Could not fetch character %s public info", character_id)
    Path(db_path).parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path)
    try:
        ensure_sso_tables(conn)
        sync_character(
            conn,
            character_id,
            access,
            refresh,
            expires_in,
            character_name,
            corporation_id=corporation_id,
            corporation_name=corporation_name,
        )
    finally:
        conn.close()
    return {
        "character_id": character_id,
        "character_name": character_name,
        "corporation_id": corporation_id,
        "corporation_name": corporation_name,
    }
