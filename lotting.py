"""Phase 3.5: Smart Lotting — group sub-$20 items into themed, marketable bundles."""

import io
import json
import math
import re
import uuid
from pathlib import Path

import anthropic
from dotenv import load_dotenv
from PIL import Image as PILImage

import db
from config import ANTHROPIC_API_KEY, LISTING_MODEL, OUTPUT_DIR

load_dotenv(Path(__file__).parent.parent / ".env.txt", override=True)
load_dotenv(Path(__file__).parent / ".env", override=True)

YARD_SALE_CEILING = 10.0   # Items below this that don't group go to yard sale
LOT_MIN = 3
LOT_MAX = 5
LOT_PRICE_MIN = 35.0
LOT_PRICE_MAX = 75.0
THUMB_SIZE = 600            # Composite thumbnail canvas width/height

LOTTING_PROMPT = """You are an expert vintage reseller who specialises in curating themed lots that
sell for more than the sum of their parts. Below is a list of items that individually price under $20.
Your job is to group them into attractive, marketable lots.

ITEMS:
{items_json}

RULES:
1. Lot size: 3–5 items per lot.
2. Lot price: target $35–$75 per lot (typically 65–80% of the sum of individual prices to make it
   feel like a deal, but never below $35).
3. Grouping priority:
   a. Same style_keywords + same probable_era → strongest match
   b. Same item_type across eras → good match
   c. Shared aesthetic theme across types (e.g. "cottagecore", "coastal", "boho") → valid match
4. lot_name must be specific and marketable — never "Misc Lot". Examples:
   "Vintage 1970s Ceramic Elephant Collection", "Mid-Century Asian Decor Bundle",
   "Boho Elephant Figurine Trio", "Japan Import Porcelain Animal Set"
5. lot_description: 2 short paragraphs. Para 1 = the story/appeal of the collection.
   Para 2 = what's included and condition summary.
6. suggested_platform: "eBay" for branded/collectible lots, "Etsy" for aesthetic/decorative lots,
   "Facebook Marketplace" for purely local/bulky lots.
7. Any item under ${yard_sale_ceiling} that genuinely doesn't fit any group → put in yard_sale list.
   Items $10–$20 that don't fit a group → force into the closest available lot rather than yard sale.

Return ONLY a JSON object with this exact structure — no preamble, no explanation:
{{
  "lots": [
    {{
      "lot_name": "string",
      "item_ids": ["item_id_1", "item_id_2", "item_id_3"],
      "lot_price": 45.00,
      "lot_description": "string",
      "suggested_platform": "Etsy"
    }}
  ],
  "yard_sale": ["item_id_a", "item_id_b"]
}}"""


def _build_item_summary(item: dict) -> dict:
    """Slim down an item dict to what the model needs for grouping."""
    return {
        "item_id": item["item_id"],
        "item_type": item.get("item_type", ""),
        "probable_maker": item.get("probable_maker", ""),
        "probable_era": item.get("probable_era", ""),
        "material": item.get("material", ""),
        "style_keywords": json.loads(item.get("style_keywords") or "[]"),
        "condition_notes": item.get("condition_notes", "")[:120],
        "suggested_price": item.get("suggested_price", 0),
        "ebay_title": item.get("ebay_title", ""),
    }


def _call_claude_for_lots(items: list[dict]) -> dict:
    summaries = [_build_item_summary(i) for i in items]
    prompt = LOTTING_PROMPT.format(
        items_json=json.dumps(summaries, indent=2),
        yard_sale_ceiling=YARD_SALE_CEILING,
    )
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    response = client.messages.create(
        model=LISTING_MODEL,
        max_tokens=4096,
        messages=[{"role": "user", "content": prompt}],
    )
    raw = response.content[0].text
    raw = re.sub(r"```(?:json)?\s*", "", raw).strip().rstrip("`").strip()
    return json.loads(raw)


def _make_composite_thumbnail(item_ids: list[str], lot_id: str) -> Path | None:
    """Arrange the first photo of each item in a grid, save as PNG."""
    images = []
    for iid in item_ids:
        item = db.get_item(iid)
        if not item:
            continue
        paths = json.loads(item.get("photo_paths") or "[]")
        if not paths:
            continue
        src = Path(paths[0])
        # Try processed_photos if not at original path
        if not src.exists():
            from config import PROCESSED_PHOTOS_DIR
            alt = PROCESSED_PHOTOS_DIR / src.name
            if alt.exists():
                src = alt
            else:
                continue
        try:
            img = PILImage.open(src).convert("RGB")
            images.append(img)
        except Exception:
            continue

    if not images:
        return None

    n = len(images)
    cols = math.ceil(math.sqrt(n))
    rows = math.ceil(n / cols)
    cell_w = THUMB_SIZE // cols
    cell_h = THUMB_SIZE // rows

    canvas = PILImage.new("RGB", (cols * cell_w, rows * cell_h), (30, 30, 30))
    for idx, img in enumerate(images):
        img.thumbnail((cell_w - 4, cell_h - 4), PILImage.LANCZOS)
        col = idx % cols
        row = idx // cols
        x = col * cell_w + (cell_w - img.width) // 2
        y = row * cell_h + (cell_h - img.height) // 2
        canvas.paste(img, (x, y))

    thumb_dir = OUTPUT_DIR / "lot_thumbs"
    thumb_dir.mkdir(parents=True, exist_ok=True)
    out_path = thumb_dir / f"{lot_id}_composite.png"
    canvas.save(out_path, "PNG")
    return out_path


def run_lotting(price_ceiling: float = 20.0) -> dict:
    """
    Main entry point. Groups eligible items into lots, flags yard-sale stragglers.
    Returns summary dict: {lots_created, items_lotted, yard_sale_count}.
    """
    items = db.get_lottable_items(price_ceiling)
    if not items:
        print("  No eligible items for lotting (all either >= $20, already lotted, or not yet listed).")
        return {"lots_created": 0, "items_lotted": 0, "yard_sale_count": 0}

    print(f"  {len(items)} items eligible for lotting (suggested_price < ${price_ceiling:.0f})")

    # Ask Claude to group them
    print("  Asking Claude to group items...", end=" ", flush=True)
    try:
        result = _call_claude_for_lots(items)
    except Exception as exc:
        print(f"FAILED: {exc}")
        return {"lots_created": 0, "items_lotted": 0, "yard_sale_count": 0}
    print("OK")

    lots_created = 0
    items_lotted = 0

    for lot_data in result.get("lots", []):
        lot_id = "lot_" + uuid.uuid4().hex[:8]
        item_ids = lot_data.get("item_ids", [])

        if len(item_ids) < LOT_MIN:
            print(f"  SKIP lot '{lot_data.get('lot_name')}' — only {len(item_ids)} items (min {LOT_MIN})")
            continue

        # Validate all item_ids exist
        valid_ids = [iid for iid in item_ids if db.get_item(iid)]
        if len(valid_ids) < LOT_MIN:
            print(f"  SKIP lot '{lot_data.get('lot_name')}' — too few valid item IDs")
            continue

        # Build composite thumbnail
        thumb_path = _make_composite_thumbnail(valid_ids, lot_id)

        db.upsert_lot(
            lot_id,
            lot_name=lot_data.get("lot_name", "Untitled Lot"),
            item_ids=json.dumps(valid_ids),
            lot_price=float(lot_data.get("lot_price") or 0),
            lot_description=lot_data.get("lot_description", ""),
            suggested_platform=lot_data.get("suggested_platform", "eBay"),
            thumbnail_path=str(thumb_path) if thumb_path else "",
            status="draft",
        )

        # Flag each item with this lot_id so it's excluded from individual export
        for iid in valid_ids:
            db.upsert_item(iid, lot_id=lot_id)

        price_sum = sum(
            (db.get_item(iid) or {}).get("suggested_price") or 0
            for iid in valid_ids
        )
        print(
            f"  LOT  {lot_id}  '{lot_data.get('lot_name')}'  "
            f"({len(valid_ids)} items, sum=${price_sum:.2f} -> lot=${lot_data.get('lot_price'):.2f})"
            f"  {lot_data.get('suggested_platform')}"
        )
        lots_created += 1
        items_lotted += len(valid_ids)

    # Flag yard-sale items
    yard_sale_count = 0
    for iid in result.get("yard_sale", []):
        item = db.get_item(iid)
        if item and not item.get("lot_id"):
            db.upsert_item(iid, yard_sale=1, suggested_platform="Facebook Marketplace")
            print(f"  YARD SALE  {iid}  (${item.get('suggested_price', 0):.2f})")
            yard_sale_count += 1

    print(f"\n  Summary: {lots_created} lots created, {items_lotted} items lotted, {yard_sale_count} yard-sale flags")
    return {"lots_created": lots_created, "items_lotted": items_lotted, "yard_sale_count": yard_sale_count}
