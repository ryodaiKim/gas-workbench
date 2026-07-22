import re

from config import HEADERS


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
    from datetime import datetime
    now = datetime.now()
    return str(now.year), str(now.month), str(now.day)


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
    return list(dict.fromkeys(items))


def collect_with_positions(text: str, pattern: str, group: int = 0):
    results = []
    for m in re.finditer(pattern, text):
        results.append((m.group(group), m.start()))
    return results


def collect_subjects(text: str):
    """Find all CIDP subject IDs, deduplicating consecutive repeats."""
    raw = collect_with_positions(text, r"CIDP-([A-Z]{3})-(\d{4})")
    results = []
    for _unused, pos in raw:
        m = re.search(r"CIDP-[A-Z]{3}-\d{4}", text[pos : pos + 20])
        sid = m.group(0) if m else _unused
        if not results or sid != results[-1][0] or pos - results[-1][1] > 50:
            results.append((sid, pos))
    return results


def parse_text(text: str) -> list[dict]:
    """Parse OCR text into records matching the sheet format."""
    trial_name = extract_trial_name(text)
    doc_year, doc_month, doc_day = extract_reception_date(text)
    fallback_date = f"{doc_year}{int(doc_month):02d}{int(doc_day):02d}"
    subjects = collect_subjects(text)

    genders = collect_with_positions(text, r"(?:^|[\s\t])([男女])", group=1)
    points = collect_with_positions(
        text, r"(初回登録時|追跡時\s*[(（][^)）]*[)）])", group=1
    )

    item_groups = []
    for m in re.finditer(r"\d{3}【([^】\n]+)】?", text):
        raw = m.group(1).strip()
        if raw:
            item_groups.append((raw, m.start()))

    if not item_groups:
        raise ValueError(
            f"No item groups (【...】) found. Text preview: {text[:500]}"
        )

    records = []
    for raw_items, ig_pos in item_groups:
        subject = find_last_before(subjects, ig_pos)
        gender = find_last_before(genders, ig_pos)
        point = find_last_before(points, ig_pos)

        for item in expand_item_group(raw_items):
            records.append({
                "試験名": trial_name,
                "被験者番号": subject[0] if subject else "",
                "性別": gender[0] if gender else "",
                "採取日": fallback_date,
                "ポイント名": point[0] if point else "初回登録時",
                "検査項目": item,
            })

    return records
