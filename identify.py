"""Phase 1: Send item photos to Claude for structured identification."""

import base64
import io
import json
import re
import time
from pathlib import Path

import anthropic
from dotenv import load_dotenv
from PIL import Image as PILImage

import db
from config import ANTHROPIC_API_KEY, CONFIDENCE_THRESHOLD, IDENTIFY_MODEL, MAX_RETRIES

load_dotenv(Path(__file__).parent.parent / ".env.txt", override=True)
load_dotenv(Path(__file__).parent / ".env", override=True)

SYSTEM_PROMPT = """You are an expert antique and vintage item appraiser with decades of experience
identifying items for resale. You analyze photos carefully, noting maker's marks, materials,
construction methods, stylistic details, and condition. You always respond with valid JSON."""

IDENTIFY_PROMPT = """Examine all photos of this antique/vintage item and return a JSON object with
exactly these fields. Be specific and accurate — your output drives pricing and listing research.

{
  "item_type": "precise description e.g. 'ceramic elephant figurine' or 'sterling silver sugar tongs'",
  "probable_maker": "manufacturer/artist name, or 'unknown'",
  "makers_marks_observed": "verbatim text from any marks, stamps, stickers, or labels visible — or 'none visible'",
  "probable_era": "decade or range e.g. '1950s-1960s' or 'Victorian era (1880s-1900s)'",
  "material": "primary material(s) e.g. 'porcelain with hand-painted glaze'",
  "style_keywords": ["3-6 searchable terms for eBay/Etsy e.g. 'mid century modern', 'hand painted', 'Japan import'"],
  "condition_notes": "honest description of chips, cracks, wear, repairs, fading — or 'excellent, no flaws observed'",
  "confidence_score": 0-100 integer,
  "confidence_reasoning": "one sentence explaining your confidence level",
  "suggested_search_query": "the exact eBay search string to find sold comps e.g. 'Andrea Sadek elephant figurine porcelain'"
}

Return ONLY the JSON object. No preamble, no explanation."""


MAX_IMAGE_BYTES = 4 * 1024 * 1024  # 4MB safety margin under Claude's 5MB limit


def encode_image(path: Path) -> tuple[str, str]:
    """Return (base64_data, media_type). Auto-resizes images over MAX_IMAGE_BYTES."""
    raw = path.read_bytes()
    if len(raw) <= MAX_IMAGE_BYTES:
        ext = path.suffix.lower()
        media_type = {".png": "image/png", ".webp": "image/webp", ".gif": "image/gif"}.get(ext, "image/jpeg")
        return base64.standard_b64encode(raw).decode("utf-8"), media_type

    # Resize until it fits
    with PILImage.open(path) as img:
        if img.mode in ("RGBA", "P"):
            img = img.convert("RGB")
        for max_dim in (2000, 1500, 1200, 900):
            img.thumbnail((max_dim, max_dim), PILImage.LANCZOS)
            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=85)
            data = buf.getvalue()
            if len(data) <= MAX_IMAGE_BYTES:
                return base64.standard_b64encode(data).decode("utf-8"), "image/jpeg"

    return base64.standard_b64encode(data).decode("utf-8"), "image/jpeg"


def build_message_content(photo_paths: list[Path]) -> list[dict]:
    content = []
    for path in photo_paths:
        data, media_type = encode_image(path)
        content.append({
            "type": "image",
            "source": {"type": "base64", "media_type": media_type, "data": data},
        })
    content.append({"type": "text", "text": IDENTIFY_PROMPT})
    return content


def parse_identification(raw: str) -> dict:
    """Extract and validate JSON from model response."""
    # Strip markdown code fences if present
    raw = re.sub(r"```(?:json)?\s*", "", raw).strip().rstrip("`").strip()
    data = json.loads(raw)

    required = [
        "item_type", "probable_maker", "makers_marks_observed", "probable_era",
        "material", "style_keywords", "condition_notes", "confidence_score",
        "confidence_reasoning", "suggested_search_query",
    ]
    for field in required:
        if field not in data:
            raise ValueError(f"Missing field in identification response: {field}")

    data["confidence_score"] = int(data["confidence_score"])
    if not isinstance(data["style_keywords"], list):
        data["style_keywords"] = [data["style_keywords"]]

    return data


def identify_item(item_id: str, photo_paths: list[Path]) -> dict:
    """Run identification for one item. Returns parsed identification dict."""
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    db.upsert_item(item_id, status="identifying")
    db.log_run(item_id, "identify", "started", f"{len(photo_paths)} photos")

    last_error = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            content = build_message_content(photo_paths)
            response = client.messages.create(
                model=IDENTIFY_MODEL,
                max_tokens=1024,
                system=SYSTEM_PROMPT,
                messages=[{"role": "user", "content": content}],
            )
            raw = response.content[0].text
            result = parse_identification(raw)

            needs_review = 1 if result["confidence_score"] < CONFIDENCE_THRESHOLD else 0

            db.upsert_item(
                item_id,
                status="identified",
                item_type=result["item_type"],
                probable_maker=result["probable_maker"],
                makers_marks_observed=result["makers_marks_observed"],
                probable_era=result["probable_era"],
                material=result["material"],
                style_keywords=json.dumps(result["style_keywords"]),
                condition_notes=result["condition_notes"],
                confidence_score=result["confidence_score"],
                confidence_reasoning=result["confidence_reasoning"],
                suggested_search_query=result["suggested_search_query"],
                needs_review=needs_review,
            )
            db.log_run(item_id, "identify", "success",
                       f"confidence={result['confidence_score']} needs_review={bool(needs_review)}")
            return result

        except Exception as exc:
            last_error = exc
            db.log_run(item_id, "identify", f"error_attempt_{attempt}", str(exc))
            if attempt < MAX_RETRIES:
                time.sleep(2 ** attempt)

    db.upsert_item(item_id, status="failed")
    db.append_error(item_id, f"identify failed after {MAX_RETRIES} attempts: {last_error}")
    raise RuntimeError(f"Identification failed for {item_id}: {last_error}") from last_error
