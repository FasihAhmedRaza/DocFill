"""
history.py — persistent extraction history with a pluggable backend.

Two interchangeable stores, auto-selected at runtime:

  • Google Sheets (via a Google Apps Script Web App)  — used when
    st.secrets["GSHEET_WEBAPP_URL"] is set. Permanent in the cloud AND local.
  • SQLite (data/history.db)                          — local fallback when
    no Web App URL is configured. Survives restarts on one machine.

Public API (backend-agnostic):
    init(), active_backend(), add_records(), fetch_all(), delete_ids(),
    clear_all(), stats(), test_connection()

The Google side speaks plain JSON over HTTPS to the Apps Script Web App — no
service account, no OAuth, no extra Python dependency (stdlib urllib).
See apps_script/Code.gs for the script to paste into Google.
"""

import json
import os
import sqlite3
import urllib.request
from contextlib import closing
from datetime import datetime

# ── shared schema ──────────────────────────────────────────────────────────
RECORD_COLS = [
    "source_file", "english_name", "arabic_name", "iban", "account_number",
    "swift", "currency", "bank_name_ar", "bank_name_en", "nationality",
    "emirates_id", "date_of_birth", "id_expiry", "emirate", "iban_valid",
]

DB_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
DB_PATH = os.path.join(DB_DIR, "history.db")


# ── config / backend selection ─────────────────────────────────────────────
def _secret(key: str, default: str = "") -> str:
    """Read config from st.secrets first, then environment, then default."""
    val = ""
    try:
        import streamlit as st
        val = st.secrets.get(key, "") or ""
    except Exception:
        val = ""
    if not val:
        val = os.environ.get(key, "")
    return val or default


def _webapp_url() -> str:
    return _secret("GSHEET_WEBAPP_URL")


def _token() -> str:
    return _secret("GSHEET_TOKEN")


def active_backend() -> str:
    """'gsheet' when a Web App URL is configured, else 'sqlite'."""
    return "gsheet" if _webapp_url() else "sqlite"


def _now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _row_record(rec: dict) -> dict:
    """Normalize one record to the persisted shape (iban_valid as 0/1)."""
    out = {}
    for c in RECORD_COLS:
        v = rec.get(c, "")
        if c == "iban_valid":
            v = 1 if v else 0
        out[c] = "" if v is None else v
    return out


def _match(row: dict, search: str) -> bool:
    s = search.strip().lower()
    if not s:
        return True
    hay = " ".join(str(row.get(k, "")) for k in
                   ("english_name", "arabic_name", "iban", "emirates_id", "source_file"))
    return s in hay.lower()


# ════════════════════════════════════════════════════════════════════════════
#  GOOGLE SHEETS BACKEND (Apps Script Web App)
# ════════════════════════════════════════════════════════════════════════════
def _gs_call(action: str, **payload) -> dict:
    url = _webapp_url()
    body = json.dumps({"action": action, "token": _token(), **payload}).encode("utf-8")
    req = urllib.request.Request(
        url, data=body, headers={"Content-Type": "application/json"}, method="POST"
    )
    with urllib.request.urlopen(req, timeout=25) as resp:   # follows Apps Script redirect
        data = json.loads(resp.read().decode("utf-8"))
    if not data.get("ok", False):
        raise RuntimeError(data.get("error", "Google Sheets request failed"))
    return data


def _gs_add(records: list, batch_id: str) -> int:
    ts = _now()
    rows = [{"ts": ts, "batch_id": batch_id, **_row_record(r)} for r in records]
    return int(_gs_call("add", records=rows).get("added", 0))


def _gs_fetch(search: str) -> list:
    rows = _gs_call("list").get("records", [])
    rows.sort(key=lambda r: str(r.get("ts", "")), reverse=True)   # newest first
    return [r for r in rows if _match(r, search)]


def _gs_delete(ids: list) -> int:
    return int(_gs_call("delete", ids=ids).get("deleted", 0))


def _gs_clear() -> int:
    return int(_gs_call("clear").get("removed", 0))


# ════════════════════════════════════════════════════════════════════════════
#  SQLITE BACKEND (local)
# ════════════════════════════════════════════════════════════════════════════
def _sq_connect() -> sqlite3.Connection:
    os.makedirs(DB_DIR, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def _sq_init() -> None:
    cols = ",\n            ".join(f"{c} TEXT" for c in RECORD_COLS if c != "iban_valid")
    with closing(_sq_connect()) as conn:
        conn.execute(
            f"""CREATE TABLE IF NOT EXISTS records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ts TEXT NOT NULL,
            batch_id TEXT NOT NULL,
            {cols},
            iban_valid INTEGER DEFAULT 0
        )"""
        )
        conn.commit()


def _sq_add(records: list, batch_id: str) -> int:
    ts = _now()
    placeholders = ", ".join(["?"] * (2 + len(RECORD_COLS)))
    colnames = "ts, batch_id, " + ", ".join(RECORD_COLS)
    rows = []
    for rec in records:
        r = _row_record(rec)
        rows.append(tuple([ts, batch_id] + [r[c] for c in RECORD_COLS]))
    with closing(_sq_connect()) as conn:
        conn.executemany(f"INSERT INTO records ({colnames}) VALUES ({placeholders})", rows)
        conn.commit()
    return len(rows)


def _sq_fetch(search: str) -> list:
    query, params = "SELECT * FROM records", ()
    if search.strip():
        like = f"%{search.strip()}%"
        query += (" WHERE english_name LIKE ? OR arabic_name LIKE ? OR iban LIKE ? "
                  "OR emirates_id LIKE ? OR source_file LIKE ?")
        params = (like, like, like, like, like)
    query += " ORDER BY id DESC"
    with closing(_sq_connect()) as conn:
        return [dict(r) for r in conn.execute(query, params).fetchall()]


def _sq_delete(ids: list) -> int:
    if not ids:
        return 0
    marks = ", ".join("?" * len(ids))
    with closing(_sq_connect()) as conn:
        cur = conn.execute(f"DELETE FROM records WHERE id IN ({marks})", tuple(ids))
        conn.commit()
        return cur.rowcount


def _sq_clear() -> int:
    with closing(_sq_connect()) as conn:
        cur = conn.execute("DELETE FROM records")
        conn.commit()
        return cur.rowcount


# ════════════════════════════════════════════════════════════════════════════
#  PUBLIC API (dispatches to the active backend)
# ════════════════════════════════════════════════════════════════════════════
def init() -> None:
    """Prepare the local store. (Google Sheets needs no client-side init.)"""
    _sq_init()


# Backwards-compatible alias.
init_db = init


def add_records(records: list, batch_id: str | None = None) -> int:
    if not records:
        return 0
    batch_id = batch_id or _now()
    if active_backend() == "gsheet":
        return _gs_add(records, batch_id)
    return _sq_add(records, batch_id)


def fetch_all(search: str = "") -> list:
    if active_backend() == "gsheet":
        return _gs_fetch(search)
    return _sq_fetch(search)


def delete_ids(ids: list) -> int:
    if active_backend() == "gsheet":
        return _gs_delete(ids)
    return _sq_delete(ids)


def clear_all() -> int:
    if active_backend() == "gsheet":
        return _gs_clear()
    return _sq_clear()


def stats() -> dict:
    rows = fetch_all()
    batches = {r.get("batch_id") for r in rows}
    valid = sum(1 for r in rows if str(r.get("iban_valid")) in ("1", "True", "true"))
    last = max((r.get("ts", "") for r in rows), default="") or "—"
    return {"total": len(rows), "batches": len(batches), "valid_ibans": valid, "last": last}


def test_connection() -> tuple:
    """(ok, message) — verify the active backend is reachable."""
    try:
        if active_backend() == "gsheet":
            _gs_call("list")
            return True, "Connected to Google Sheets"
        _sq_init()
        return True, "Local SQLite ready"
    except Exception as e:
        return False, str(e)
