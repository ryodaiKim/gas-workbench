#!/usr/bin/env python3
"""
Auto Sample Registration Pipeline

Scans a PDF folder, OCRs new files via tesseract, parses clinical specimen
receipt tables, and appends the results to a Google Sheet.
Already-processed files are tracked in processed_log.json and skipped.
"""

import json
import logging
import os
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from config import PDF_FOLDER, GOOGLE_SHEET_ID, PROCESSED_LOG
from ocr import ocr_pdf
from parser import parse_text
from sheets import append_records

log = logging.getLogger("pipeline")


# ---------------------------------------------------------------------------
# Processed-file tracking
# ---------------------------------------------------------------------------

def load_processed(path: Path) -> dict:
    if path.exists():
        with open(path) as f:
            return json.load(f)
    return {}


def save_processed(path: Path, data: dict) -> None:
    # Atomic write via temp file + rename
    fd, tmp = tempfile.mkstemp(dir=path.parent, suffix=".tmp")
    try:
        with os.fdopen(fd, "w") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        os.replace(tmp, path)
    except BaseException:
        os.unlink(tmp)
        raise


# ---------------------------------------------------------------------------
# Pipeline
# ---------------------------------------------------------------------------

def run(dry_run: bool = False) -> None:
    if not PDF_FOLDER:
        log.error("PDF_FOLDER is not set. Check your .env file.")
        sys.exit(1)

    folder = Path(PDF_FOLDER)
    if not folder.is_dir():
        log.error("PDF_FOLDER does not exist: %s", folder)
        sys.exit(1)

    pdf_files = sorted(folder.glob("*.pdf"))
    if not pdf_files:
        log.info("No PDF files found in %s", folder)
        return

    processed = load_processed(PROCESSED_LOG)
    new_files = [f for f in pdf_files if f.name not in processed]

    if not new_files:
        log.info("All %d PDFs already processed. Nothing to do.", len(pdf_files))
        return

    log.info(
        "Found %d PDFs total, %d new to process", len(pdf_files), len(new_files)
    )

    total_records = 0
    errors = 0

    for pdf in new_files:
        log.info("Processing: %s", pdf.name)
        try:
            text = ocr_pdf(str(pdf))
            records = parse_text(text)
            log.info("  %d records parsed", len(records))

            if not dry_run:
                if GOOGLE_SHEET_ID:
                    append_records(records)
                else:
                    log.warning("  GOOGLE_SHEET_ID not set — skipping sheet write")

            processed[pdf.name] = datetime.now(timezone.utc).isoformat()
            save_processed(PROCESSED_LOG, processed)
            total_records += len(records)

        except Exception:
            log.exception("  FAIL: %s", pdf.name)
            errors += 1

    mode = "DRY RUN" if dry_run else "DONE"
    log.info(
        "%s: %d files processed, %d records written, %d errors",
        mode,
        len(new_files) - errors,
        total_records,
        errors,
    )


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def setup_logging() -> None:
    fmt = "%(asctime)s [%(levelname)s] %(message)s"
    logging.basicConfig(level=logging.INFO, format=fmt, datefmt="%Y-%m-%d %H:%M:%S")


if __name__ == "__main__":
    setup_logging()
    dry = "--dry-run" in sys.argv
    run(dry_run=dry)
