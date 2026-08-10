"""
Filtration Service Estimator — dry-weight contribution storage.

save_contribution(payload) appends the farmer's measured dry-weight data to a
Google Sheet, ONE ROW PER SIZE CLASS (each row is a shell-height -> dry-weight
datapoint, which is exactly what improves the length-weight conversions).

Mirrors AquaPIP's storage.py: lazy imports, fully defensive — if anything is
missing or fails it returns (False, reason) and NEVER raises, so a storage
problem can never break the calculator. Nothing is written unless the farmer
opts in (the app only calls this after they tick the consent box).

Configuration (in the filtration app's Streamlit secrets — the SAME service
account you use for AquaPIP is fine):

    [gcp_service_account]
    ... service-account JSON fields ...

    [contributions]
    sheet_id = "the target spreadsheet id"
    worksheet = "dry_weight_contributions"   # optional; created if missing
"""

import uuid
import datetime

HEADER = [
    "timestamp_utc", "submission_id", "species", "shell_height_mm",
    "dry_weight_g", "number_of_organisms", "temperature_c",
    "contributor_name", "contributor_email",
]


def _clean(x):
    return "" if x is None else x


def build_rows(payload):
    """Turn one submission into a list of rows (one per size class).
    Kept import-free and pure so it can be unit-tested without gspread."""
    ts = datetime.datetime.now(datetime.timezone.utc).isoformat(timespec="seconds").replace("+00:00", "Z")
    sid = uuid.uuid4().hex[:12]
    species = _clean(payload.get("species"))
    temp = _clean(payload.get("temperature_c"))
    name = _clean(payload.get("name"))
    email = _clean(payload.get("email"))
    rows = []
    for c in payload.get("classes", []):
        # skip empty rows
        if c.get("shell_mm") is None and c.get("dtw_g") is None and not c.get("count"):
            continue
        rows.append([
            ts, sid, species,
            _clean(c.get("shell_mm")), _clean(c.get("dtw_g")),
            _clean(c.get("count")), temp, name, email,
        ])
    return rows


def is_configured():
    """True only if secrets contain a service account + contributions sheet id."""
    try:
        import streamlit as st
    except Exception:
        return False
    try:
        return ("gcp_service_account" in st.secrets
                and bool(st.secrets.get("contributions", {}).get("sheet_id")))
    except Exception:
        return False


def save_contribution(payload):
    """Append the contribution rows to the configured sheet.
    Returns (ok: bool, status: str). Never raises."""
    try:
        import streamlit as st
    except Exception:
        return False, "streamlit_unavailable"

    try:
        if "gcp_service_account" not in st.secrets:
            return False, "not_configured"
        conf = st.secrets.get("contributions", {})
        sheet_id = conf.get("sheet_id")
        if not sheet_id:
            return False, "no_sheet_id"
        ws_name = conf.get("worksheet", "dry_weight_contributions")

        rows = build_rows(payload)
        if not rows:
            return False, "no_data"

        import gspread
        from google.oauth2.service_account import Credentials

        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]
        creds = Credentials.from_service_account_info(
            dict(st.secrets["gcp_service_account"]), scopes=scopes)
        gc = gspread.authorize(creds)
        sh = gc.open_by_key(sheet_id)

        # get or create the contributions worksheet
        try:
            ws = sh.worksheet(ws_name)
        except Exception:
            ws = sh.add_worksheet(title=ws_name, rows=1000, cols=len(HEADER))

        # header once
        try:
            if not ws.get_all_values():
                ws.append_row(HEADER, value_input_option="RAW")
        except Exception:
            pass

        ws.append_rows(rows, value_input_option="RAW")
        return True, "ok"
    except Exception as e:                     # never let storage break the app
        return False, f"error: {type(e).__name__}: {e}"


if __name__ == "__main__":
    demo = {
        "species": "Crassostrea gigas", "temperature_c": 20.0,
        "name": "Jane Farmer", "email": "jane@example.com",
        "classes": [
            {"shell_mm": 70, "dtw_g": 1.5, "count": 8000},
            {"shell_mm": 40, "dtw_g": 0.4, "count": 4000},
            {"shell_mm": None, "dtw_g": None, "count": 0},   # should be skipped
        ],
    }
    for r in build_rows(demo):
        print(r)
