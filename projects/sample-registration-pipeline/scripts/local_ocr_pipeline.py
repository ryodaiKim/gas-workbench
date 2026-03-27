#!/usr/bin/env python3
"""
Local OCR pipeline for clinical specimen receipt PDFs.
Reads PDFs from a folder, extracts text via image-based OCR (tesseract),
parses table data, and outputs CSV matching the project sheet format.

Usage:
    python3 scripts/local_ocr_pipeline.py <pdf_folder> [--out <output.csv>] [--sheet-id <google_sheet_id>]

Requirements (install in venv):
    pip install pdf2image pytesseract Pillow gspread google-auth
"""

import argparse
import csv
import io
import os
import re
import sys
from pathlib import Path

from pdf2image import convert_from_path
import pytesseract


# ---------------------------------------------------------------------------
# OCR
# ---------------------------------------------------------------------------

def ocr_pdf(pdf_path: str, dpi: int = 300, lang: str = "jpn") -> str:
    """Convert PDF to images, OCR each page, return combined text."""
    images = convert_from_path(pdf_path, dpi=dpi)
    pages = []
    for img in images:
        text = pytesseract.image_to_string(img, lang=lang)
        pages.append(text)
    return "\n".join(pages)


# ---------------------------------------------------------------------------
# Parser
# ---------------------------------------------------------------------------

HEADERS = ["試験名", "被験者番号", "性別", "採取日", "ポイント名", "検査項目"]


def extract_trial_name(text: str) -> str:
    m = re.search(r"試験名\s*[:：]\s*(.+)", text)
    return m.group(1).strip() if m else "レジストリ研究"


def extract_reception_date(text: str) -> tuple[str, str, str]:
    """Extract year, month, day from 治験受付日, falling back to 発信日."""
    for pat in [
        r"(?:治験)?受付日\s*[:：]\s*(\d{4})\s*年?\s*(\d{1,2})\s*月?\s*(\d{1,2})",
        r"発信日\s*[:：]\s*(\d{4})\s*年?\s*(\d{1,2})\s*月?\s*(\d{1,2})",
    ]:
        m = re.search(pat, text)
        if m:
            return m.group(1), m.group(2), m.group(3)
    return str(__import__("datetime").datetime.now().year), "1", "1"



def find_last_before(items, position):
    """Find the last item whose position <= position."""
    best = None
    for item in items:
        if item[-1] <= position:
            best = item
    return best


def expand_item_group(raw: str) -> list[str]:
    """Expand 【血清分離・血漿分離・DNA】 into normalized item names."""
    parts = re.split(r"[・·、,]", raw)
    items = []
    for part in parts:
        t = part.strip()
        if not t:
            continue
        if re.search(r"(?i)dna|ＤＮＡ|DNA", t):
            items.append("ＤＮＡ抽出（Ｎ）")
        elif re.search(r"株化.*リンパ|リンパ.*株化|リンパ球", t):
            items.append("リンパ球株化１１")
        elif re.search(r"血清", t):
            items.append("血清分離（用手法）")
        elif re.search(r"血漿|血禁|血茜|血呆", t):
            # tesseract sometimes misreads 漿 as 禁/茜/呆
            items.append("血漿分離（用手法）")
        else:
            items.append(t)
    return list(dict.fromkeys(items))  # deduplicate preserving order


def collect_with_positions(text: str, pattern: str, group: int = 0):
    """Return list of (matched_text, position) tuples."""
    results = []
    for m in re.finditer(pattern, text):
        results.append((m.group(group), m.start()))
    return results


def collect_subjects(text: str):
    """Deduplicate consecutive identical subject IDs."""
    raw = collect_with_positions(text, r"CIDP-([A-Z]{3})-(\d{4})")
    results = []
    for full_match_unused, pos in raw:
        m = re.search(r"CIDP-[A-Z]{3}-\d{4}", text[pos:pos+20])
        sid = m.group(0) if m else full_match_unused
        if not results or sid != results[-1][0] or pos - results[-1][1] > 50:
            results.append((sid, pos))
    return results


def find_date_near(text: str, position: int, window: int = 500):
    """Find month/day date near the given position."""
    start = max(0, position - 200)
    end = min(len(text), position + window)
    snippet = text[start:end]
    # Prefer explicit 月日 format
    m = re.search(r"(\d{1,2})\s*月\s*(\d{1,2})\s*日", snippet)
    if m:
        month, day = int(m.group(1)), int(m.group(2))
        if 1 <= month <= 12 and 1 <= day <= 31:
            return (str(month), str(day))
    return None


def parse_text(text: str) -> list[dict]:
    """Parse OCR text into records matching the sheet format."""
    trial_name = extract_trial_name(text)
    doc_year, doc_month, doc_day = extract_reception_date(text)
    fallback_date = f"{doc_year}{int(doc_month):02d}{int(doc_day):02d}"
    subjects = collect_subjects(text)

    # Genders: standalone 男/女
    genders = collect_with_positions(text, r"(?:^|[\s\t])([男女])", group=1)

    # Points: 初回登録時 or 追跡時(N年目)
    points = collect_with_positions(text, r"(初回登録時|追跡時\s*[(（][^)）]*[)）])", group=1)

    # Item groups: NNN【content】
    item_groups = []
    for m in re.finditer(r"\d{3}【([^】\n]+)】?", text):
        raw = m.group(1).strip()
        if raw:
            item_groups.append((raw, m.start()))

    if not item_groups:
        raise ValueError(f"No item groups (【...】) found. Text preview: {text[:500]}")

    records = []
    for raw_items, ig_pos in item_groups:
        subject = find_last_before(subjects, ig_pos)
        gender = find_last_before(genders, ig_pos)
        point = find_last_before(points, ig_pos)
        subject_id = subject[0] if subject else ""
        gender_str = gender[0] if gender else ""
        point_str = point[0] if point else "初回登録時"

        for item in expand_item_group(raw_items):
            records.append({
                "試験名": trial_name,
                "被験者番号": subject_id,
                "性別": gender_str,
                "採取日": fallback_date,
                "ポイント名": point_str,
                "検査項目": item,
            })

    return records


# ---------------------------------------------------------------------------
# Sheet writing (optional)
# ---------------------------------------------------------------------------

def write_to_sheet(records: list[dict], sheet_id: str):
    """Append records to the 受付情報一覧 sheet via gspread."""
    import gspread
    from google.auth import default

    creds, _ = default(scopes=["https://www.googleapis.com/auth/spreadsheets"])
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(sheet_id)
    ws = sh.worksheet("受付情報一覧")

    rows = [[r[h] for h in HEADERS] for r in records]
    ws.append_rows(rows, value_input_option="RAW")
    print(f"  → Wrote {len(rows)} rows to sheet")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def process_folder(folder: str, out_csv: str | None, sheet_id: str | None):
    pdf_files = sorted(Path(folder).glob("*.pdf"))
    if not pdf_files:
        print(f"No PDF files found in {folder}")
        return

    all_records = []
    for pdf in pdf_files:
        print(f"Processing: {pdf.name}")
        try:
            text = ocr_pdf(str(pdf))
            records = parse_text(text)
            all_records.extend(records)
            print(f"  → {len(records)} records extracted")
        except Exception as e:
            print(f"  → FAIL: {e}", file=sys.stderr)

    if not all_records:
        print("No records extracted from any file.")
        return

    # Output CSV
    if out_csv:
        with open(out_csv, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=HEADERS)
            writer.writeheader()
            writer.writerows(all_records)
        print(f"\nWrote {len(all_records)} records to {out_csv}")
    else:
        # Print to stdout
        writer = csv.DictWriter(sys.stdout, fieldnames=HEADERS)
        writer.writeheader()
        writer.writerows(all_records)

    # Optionally write to sheet
    if sheet_id:
        write_to_sheet(all_records, sheet_id)

    print(f"\nTotal: {len(all_records)} records from {len(pdf_files)} files")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Local OCR pipeline for clinical specimen PDFs")
    parser.add_argument("folder", help="Path to folder containing PDF files")
    parser.add_argument("--out", help="Output CSV file path (default: stdout)")
    parser.add_argument("--sheet-id", help="Google Sheet ID to write results to")
    args = parser.parse_args()

    process_folder(args.folder, args.out, args.sheet_id)
