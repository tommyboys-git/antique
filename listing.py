"""Phase 3: Generate listing drafts using Claude Sonnet."""

import json
import re
import time
from pathlib import Path

import anthropic
from dotenv import load_dotenv

import db
from config import ANTHROPIC_API_KEY, LISTING_MODEL, MAX_RETRIES

load_dotenv(Path(__file__).parent.parent / ".env.txt", override=True)
load_dotenv(Path(__file__).parent / ".env", override=True)

LISTING_PROMPT = """You are an expert vintage/antique seller on eBay and Etsy with a track record of
high sell-through rates. Create optimized listing content for this item.

ITEM DATA:
{item_data}

COMP PRICING:
{comp_data}

Return a JSON object with exactly these fields:
{{
  "ebay_title": "max 80 chars, keyword-rich, no ALL CAPS spam",
  "etsy_title": "max 140 chars, descriptive and story-driven",
  "fb_marketplace_title": "max 60 chars, casual and local-friendly",
  "description": "2-3 paragraphs: para1=what it is and why it's special, para2=material/construction/marks, para3=condition and measurements if estimable",
  "suggested_price": float (based on median comp adjusted for condition; be conservative),
  "suggested_platform": "eBay|Etsy|Facebook Marketplace (pick best fit: eBay for collectibles/branded, Etsy for decorative/handmade, FB for bulky/furniture/local)",
  "tags": ["10-13 Etsy-style tags, lowercase, max 20 chars each"],
  "category_ebay": "eBay browse node path e.g. 'Pottery & Glass > Pottery & China > China & Dinnerware'"
}}

Return ONLY the JSON object."""


def generate_listing(item_id: str) -> dict:
    item = db.get_item(item_id)
    if not item:
        raise ValueError(f"Item {item_id} not found in DB")

    style_keywords = json.loads(item.get("style_keywords") or "[]")
    comp_listings = json.loads(item.get("comp_listings") or "[]")

    item_data = {
        "item_type": item.get("item_type"),
        "probable_maker": item.get("probable_maker"),
        "makers_marks_observed": item.get("makers_marks_observed"),
        "probable_era": item.get("probable_era"),
        "material": item.get("material"),
        "style_keywords": style_keywords,
        "condition_notes": item.get("condition_notes"),
    }

    comp_data = {
        "low": item.get("comp_low"),
        "median": item.get("comp_median"),
        "high": item.get("comp_high"),
        "sample_size": item.get("comp_sample_size"),
        "confidence": item.get("comp_confidence"),
        "sample_titles": [l["title"] for l in comp_listings[:5]],
    }

    prompt = LISTING_PROMPT.format(
        item_data=json.dumps(item_data, indent=2),
        comp_data=json.dumps(comp_data, indent=2),
    )

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    last_error = None

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = client.messages.create(
                model=LISTING_MODEL,
                max_tokens=1500,
                messages=[{"role": "user", "content": prompt}],
            )
            raw = response.content[0].text
            raw = re.sub(r"```(?:json)?\s*", "", raw).strip().rstrip("`").strip()
            result = json.loads(raw)

            result["suggested_price"] = float(result.get("suggested_price") or 0)
            if not isinstance(result.get("tags"), list):
                result["tags"] = []

            db.upsert_item(
                item_id,
                status="listing_done",
                ebay_title=result.get("ebay_title", "")[:80],
                etsy_title=result.get("etsy_title", "")[:140],
                fb_marketplace_title=result.get("fb_marketplace_title", "")[:60],
                description=result.get("description", ""),
                suggested_price=result["suggested_price"],
                suggested_platform=result.get("suggested_platform", ""),
                tags=json.dumps(result.get("tags", [])),
                category_ebay=result.get("category_ebay", ""),
            )
            db.log_run(item_id, "listing", "success",
                       f"platform={result.get('suggested_platform')} price=${result['suggested_price']}")
            return result

        except Exception as exc:
            last_error = exc
            db.log_run(item_id, "listing", f"error_attempt_{attempt}", str(exc))
            if attempt < MAX_RETRIES:
                time.sleep(2 ** attempt)

    db.append_error(item_id, f"listing failed: {last_error}")
    raise RuntimeError(f"Listing generation failed for {item_id}: {last_error}") from last_error
