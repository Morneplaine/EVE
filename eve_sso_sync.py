"""
EVE SSO (Single Sign-On) and ESI sync module.

Uses OAuth 2.0 with PKCE (desktop flow) to authenticate and fetch:
- Wallet transactions
- Wallet journal
- Character industry jobs (manufacturing, reprocessing, etc.)

Required: Create an SSO application at https://developers.eveonline.com/
with callback URL http://localhost:8765/callback/ and scopes:
  esi-wallet.read_character_wallet.v1
  esi-industry.read_character_jobs.v1

Store client_id and client_secret in config or environment (e.g. EVE_SSO_CLIENT_ID, EVE_SSO_CLIENT_SECRET).
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

# Scopes for wallet and industry (for profitability tracking)
DEFAULT_SCOPES = "esi-wallet.read_character_wallet.v1 esi-industry.read_character_jobs.v1"
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
    """)
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


def fetch_wallet_journal(character_id: int, access_token: str, page: int = 1) -> list:
    """Fetch one page of wallet journal (30 days)."""
    path = f"characters/{character_id}/wallet/journal/"
    return esi_get(path, access_token, {"page": page})


def fetch_industry_jobs(character_id: int, access_token: str, include_completed: bool = True, page: int = 1) -> list:
    """Fetch character industry jobs (manufacturing, etc.). include_completed: last 90 days."""
    path = f"characters/{character_id}/industry/jobs/"
    params = {"include_completed": "true" if include_completed else "false", "page": page}
    return esi_get(path, access_token, params)


def sync_character(
    conn: sqlite3.Connection,
    character_id: int,
    access_token: str,
    refresh_token: str,
    expires_in: int = 1200,
    character_name: str | None = None,
):
    """Upsert character and tokens into sso_character."""
    expires_at = time.time() + expires_in
    conn.execute(
        """
        INSERT INTO sso_character (character_id, character_name, refresh_token, access_token, access_token_expires_at, updated_at)
        VALUES (?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
        ON CONFLICT(character_id) DO UPDATE SET
            character_name = COALESCE(excluded.character_name, character_name),
            refresh_token = excluded.refresh_token,
            access_token = excluded.access_token,
            access_token_expires_at = excluded.access_token_expires_at,
            updated_at = CURRENT_TIMESTAMP
        """,
        (character_id, character_name or "", refresh_token, access_token, expires_at),
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
    Ensure valid token, then sync wallet transactions, journal, and industry jobs.
    Returns dict with counts and any error message.
    """
    ensure_sso_tables(conn)
    access = get_valid_access_token(conn, character_id, client_id, client_secret)
    if not access:
        return {"error": "No valid token. Please log in with EVE SSO first.", "tx": 0, "journal": 0, "jobs": 0}
    out = {"tx": 0, "journal": 0, "jobs": 0}
    try:
        out["tx"] = sync_wallet_transactions(conn, character_id, access)
        out["journal"] = sync_wallet_journal(conn, character_id, access)
        out["jobs"] = sync_industry_jobs(conn, character_id, access)
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
    Path(db_path).parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path)
    try:
        ensure_sso_tables(conn)
        sync_character(conn, character_id, access, refresh, expires_in, character_name)
    finally:
        conn.close()
    return {"character_id": character_id, "character_name": character_name}
