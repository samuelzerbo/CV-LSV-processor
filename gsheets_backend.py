"""
gsheets_backend.py
=====================
Logs login events (name, email, timestamp) to a Google Sheet, using a
service account. This is a DIFFERENT credential type than the OAuth client
used for st.login() -- a service account is a robot identity for the app
itself to read/write the Sheet, separate from "let a user log in".

Intentionally minimal: no file storage, no per-user processing history.
Only tracks who has signed in and when.

SETUP (see README.md for full step-by-step):
  1. Create a service account in Google Cloud Console, enable the Sheets
     API, download its JSON key.
  2. Create a Google Sheet with one tab named "Logins", header row:
     Name, Email, Timestamp
  3. Share the Sheet with the service account's email (Editor access).
  4. Add the service account JSON fields and the Sheet ID to Streamlit
     secrets under [gcp_service_account] and [gsheets].

log_login() is best-effort: if the integration isn't configured or fails
for any reason, it logs the error and returns quietly -- this must never
block someone from actually using the app.
"""

import datetime
import streamlit as st

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
LOGINS_SHEET = "Logins"


def backend_configured() -> bool:
    """Cheap check for whether login-tracking secrets exist at all."""
    try:
        return "gcp_service_account" in st.secrets and "gsheets" in st.secrets
    except Exception:
        return False


def _get_gspread_client():
    import gspread
    from google.oauth2.service_account import Credentials
    info = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return gspread.authorize(creds)


def log_login(name: str, email: str):
    """Append one row to the Logins tab. Silently no-ops on any failure,
    but stashes the error in session state so the app can optionally
    surface it for debugging (see app.py)."""
    try:
        gc = _get_gspread_client()
        sh = gc.open_by_key(st.secrets["gsheets"]["spreadsheet_id"])
        ws = sh.worksheet(LOGINS_SHEET)
        ws.append_row([name, email, datetime.datetime.utcnow().isoformat()])
        st.session_state["_gsheets_last_error"] = None
    except Exception as e:
        print(f"[gsheets_backend] log_login failed: {e}")
        st.session_state["_gsheets_last_error"] = str(e)
