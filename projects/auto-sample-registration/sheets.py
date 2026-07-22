import logging

import gspread
from google.oauth2.service_account import Credentials

from config import GOOGLE_SHEET_ID, SERVICE_ACCOUNT_JSON, WORKSHEET_NAME, HEADERS

log = logging.getLogger(__name__)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
]


def append_records(records: list[dict]) -> int:
    """Append parsed records to the Google Sheet. Returns rows written."""
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_JSON, scopes=SCOPES)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(GOOGLE_SHEET_ID)
    ws = sh.worksheet(WORKSHEET_NAME)

    rows = [[r[h] for h in HEADERS] for r in records]
    ws.append_rows(rows, value_input_option="RAW")
    log.info("Wrote %d rows to sheet", len(rows))
    return len(rows)
