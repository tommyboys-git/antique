"""
Standalone eBay File Exchange CSV exporter.
Run: python ebay_export.py
Outputs: output/ebay_upload_ready.csv
"""

import csv
import json
import re
from pathlib import Path

from dotenv import load_dotenv

BASE_DIR = Path(__file__).parent
load_dotenv(BASE_DIR / ".env", override=True)

import db
from config import OUTPUT_DIR

OUTPUT_FILE = OUTPUT_DIR / "ebay_upload_ready.csv"

# ── eBay File Exchange column order ──────────────────────────────────────────
COLUMNS = [
    "Action",
    "Category",
    "Title",
    "Description",
    "BuyItNowPrice",
    "Quantity",
    "ConditionID",
    "Format",
    "Duration",
    "Location",
    "Country",
    "Currency",
    "ReturnsAcceptedOption",
    "ReturnsWithinOption",
    "RefundOption",
    "ShippingType",
    "ShippingService-1:Option",
    "ShippingService-1:FreeShipping",
    "DispatchTimeMax",
    "PicURL",
]

# Static values applied to every row
STATIC_FIELDS = {
    "Action":                        "Add",
    "Quantity":                      "1",
    "ConditionID":                   "3000",   # Used
    "Format":                        "FixedPrice",
    "Duration":                      "GTC",
    "Location":                      "Saint Charles, MO",
    "Country":                       "US",
    "Currency":                      "USD",
    "ReturnsAcceptedOption":         "ReturnsAccepted",
    "ReturnsWithinOption":           "Days_30",
    "RefundOption":                  "MoneyBack",
    "ShippingType":                  "Calculated",
    "ShippingService-1:Option":      "USPSPriority",
    "ShippingService-1:FreeShipping":"false",
    "DispatchTimeMax":               "3",
}


def _format_description_html(raw: str) -> str:
    """
    Convert plain-text pipeline description to eBay-safe HTML.
    - Split on blank lines → individual <p> blocks
    - Single newlines within a paragraph → <br>
    - Strip leading/trailing whitespace
    """
    if not raw:
        return ""
    # Normalise line endings
    text = raw.replace("\r\n", "\n").replace("\r", "\n")
    # Split into paragraphs on one or more blank lines
    paragraphs = re.split(r"\n{2,}", text.strip())
    html_parts = []
    for para in paragraphs:
        para = para.strip()
        if not para:
            continue
        # Replace single newlines inside a paragraph with <br>
        para = para.replace("\n", "<br>")
        html_parts.append(f"<p>{para}</p>")
    return "".join(html_parts)


def _truncate_title(title: str, max_len: int = 80) -> str:
    if not title:
        return ""
    title = title.strip()
    if len(title) <= max_len:
        return title
    # Truncate at last whole word before the limit
    truncated = title[:max_len].rsplit(" ", 1)[0]
    return truncated


def build_rows() -> tuple[list[dict], list[str]]:
    """
    Returns (rows, skipped_reasons).
    Rows are individual item listings only — lots excluded (no category).
    """
    db.init_db()
    all_items = db.get_all_items()

    rows = []
    skipped = []

    for item in all_items:
        item_id = item["item_id"]

        # Skip items assigned to a lot or flagged yard-sale
        if item.get("lot_id"):
            skipped.append(f"  SKIP  {item_id}  — in a lot ({item['lot_id']})")
            continue
        if item.get("yard_sale"):
            skipped.append(f"  SKIP  {item_id}  — FB Marketplace / yard sale")
            continue

        title    = _truncate_title(item.get("ebay_title") or "")
        category = (item.get("category_ebay") or "").strip()
        price    = item.get("suggested_price") or 0

        # Rule 1: drop rows missing Title or Category
        if not title:
            skipped.append(f"  SKIP  {item_id}  — empty Title")
            continue
        if not category:
            skipped.append(f"  SKIP  {item_id}  — empty Category")
            continue
        if price <= 0:
            skipped.append(f"  SKIP  {item_id}  — price is $0")
            continue

        # Description → HTML
        raw_desc = item.get("description") or ""
        html_desc = _format_description_html(raw_desc)

        # PicURL — resolve to actual file location (may have moved to processed_photos)
        from config import PROCESSED_PHOTOS_DIR
        photo_paths = json.loads(item.get("photo_paths") or "[]")
        pic_url = ""
        if photo_paths:
            p = Path(photo_paths[0])
            if p.exists():
                pic_url = str(p)
            else:
                alt = PROCESSED_PHOTOS_DIR / p.name
                pic_url = str(alt) if alt.exists() else str(p)

        row = {**STATIC_FIELDS}
        row["Title"]         = title
        row["Category"]      = category
        row["Description"]   = html_desc
        row["BuyItNowPrice"] = f"{price:.2f}"
        row["PicURL"]        = pic_url
        rows.append((item_id, row))

    return rows, skipped


def export_csv() -> Path:
    OUTPUT_DIR.mkdir(exist_ok=True)
    rows_with_ids, skipped = build_rows()

    with open(OUTPUT_FILE, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=COLUMNS, extrasaction="ignore")
        writer.writeheader()
        for _, row in rows_with_ids:
            writer.writerow(row)

    return rows_with_ids, skipped


def main():
    print("\n=== eBay File Exchange CSV Export ===\n")

    rows_with_ids, skipped = export_csv()

    if skipped:
        print("Skipped:")
        for s in skipped:
            print(s)
        print()

    print(f"Included listings ({len(rows_with_ids)}):")
    for item_id, row in rows_with_ids:
        title = row["Title"]
        over = " [TRUNCATED]" if len((db.get_item(item_id) or {}).get("ebay_title") or "") > 80 else ""
        print(f"  {item_id:45s}  ${row['BuyItNowPrice']:>7}  {title[:55]}{over}")

    print(f"\nOutput: {OUTPUT_FILE}")
    print(f"Total rows in CSV: {len(rows_with_ids)}")
    print(f"Total skipped:     {len(skipped)}")


if __name__ == "__main__":
    main()
