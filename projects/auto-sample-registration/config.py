import os
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()

BASE_DIR = Path(__file__).resolve().parent

PDF_FOLDER = os.path.expanduser(os.environ.get("PDF_FOLDER", ""))
GOOGLE_SHEET_ID = os.environ.get("GOOGLE_SHEET_ID", "")
SERVICE_ACCOUNT_JSON = os.environ.get("SERVICE_ACCOUNT_JSON", "credentials.json")
PROCESSED_LOG = BASE_DIR / "processed_log.json"

WORKSHEET_NAME = "受付情報一覧"
HEADERS = ["試験名", "被験者番号", "性別", "採取日", "ポイント名", "検査項目"]

OCR_DPI = 300
OCR_LANG = "jpn"

# --- opendataloader-pdf settings ---
# Set to "false" to disable opendataloader-pdf and use Tesseract only
USE_OPENDATALOADER = os.environ.get("USE_OPENDATALOADER", "true").lower() in ("1", "true", "yes")
# Enable hybrid mode for AI-backed OCR (requires running: opendataloader-pdf-hybrid --port 5002)
OPENDATALOADER_HYBRID = os.environ.get("OPENDATALOADER_HYBRID", "false").lower() in ("1", "true", "yes")
# Force OCR even on digitally-created PDFs (recommended for scanned documents)
OPENDATALOADER_FORCE_OCR = os.environ.get("OPENDATALOADER_FORCE_OCR", "true").lower() in ("1", "true", "yes")

# --- Tesseract fallback settings ---
# Windows-specific paths (ignored on macOS/Linux when tools are on PATH)
POPPLER_PATH = os.environ.get("POPPLER_PATH", "")
TESSERACT_CMD = os.environ.get("TESSERACT_CMD", "")
