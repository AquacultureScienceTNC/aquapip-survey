"""
AquaPIP — response storage.

save_response(submission) appends one row to a Google Sheet, but ONLY when the
practitioner has ticked the consent box AND you have configured credentials
(see README + .streamlit/secrets.toml.template).

This module is intentionally defensive: if anything is missing or fails, it
returns (False, reason) and NEVER raises, so a storage problem can never break
the survey for a user. The privacy-safe default is simply "not stored".
"""

import json
import datetime


def is_configured():
    """True only if Streamlit secrets contain a service account + sheet id."""
    try:
        import streamlit as st
    except Exception:
        return False
    try:
        return ("gcp_service_account" in st.secrets
                and bool(st.secrets.get("responses", {}).get("sheet_id")))
    except Exception:
        return False


def save_response(submission):
    """Append the submission to the configured Google Sheet.
    Returns (ok: bool, status: str)."""
    try:
        import streamlit as st
    except Exception:
        return False, "streamlit_unavailable"

    try:
        if "gcp_service_account" not in st.secrets:
            return False, "not_configured"
        sheet_id = st.secrets.get("responses", {}).get("sheet_id")
        if not sheet_id:
            return False, "no_sheet_id"

        import gspread
        from google.oauth2.service_account import Credentials

        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]
        creds = Credentials.from_service_account_info(
            dict(st.secrets["gcp_service_account"]), scopes=scopes)
        gc = gspread.authorize(creds)
        ws = gc.open_by_key(sheet_id).sheet1

        # Write a header row once, if the sheet is empty.
        try:
            if not ws.get_all_values():
                ws.append_row(
                    ["timestamp_utc", "farm_name", "interested_in_measuring",
                     "indicator_codes", "farm_profile_json", "consent"],
                    value_input_option="RAW")
        except Exception:
            pass

        row = [
            datetime.datetime.utcnow().isoformat(timespec="seconds") + "Z",
            submission.get("farm_name", ""),
            "; ".join(submission.get("farmer_goals", [])),
            "; ".join(submission.get("codes", [])),
            json.dumps(submission.get("profile", {}), ensure_ascii=False),
            "consented",
        ]
        ws.append_row(row, value_input_option="RAW")
        return True, "ok"
    except Exception as e:  # never let storage break the app
        return False, f"error: {type(e).__name__}: {e}"
