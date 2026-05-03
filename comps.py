"""
Phase 2: Research eBay sold comps via Claude web search (Anthropic API).

eBay blocks all headless-browser scraping via Akamai (HTTP 403).
Claude's built-in web_search_20250305 tool routes through Anthropic's
infrastructure, which eBay does not block.
"""

import json
import re
import time
import random
from pathlib import Path

import anthropic
from dotenv import load_dotenv

BASE_DIR = Path(__file__).parent
load_dotenv(BASE_DIR / ".env", override=True)

import db
from config import COMP_DELAY_SECONDS, MAX_RETRIES, LISTING_MODEL

_client: anthropic.Anthropic | None = None

WEB_SEARCH_TOOL = {"type": "web_search_20250305", "name": "web_search"}


def _get_client() -> anthropic.Anthropic:
    global _client
    if _client is None:
        _client = anthropic.Anthropic()
    return _client


def close_browser():
    """No-op — kept for pipeline.py compatibility (was Playwright cleanup)."""
    pass


# ── Query builder ─────────────────────────────────────────────────────────────

def make_search_query(item: dict) -> str:
    """
    Build a tight 3-5 keyword eBay search query from item data.
    Priority: maker + item_type [+ model/mark if short].
    Falls back to first 4 words of suggested_search_query.
    """
    maker  = (item.get("probable_maker") or "").strip()
    itype  = (item.get("item_type") or "").strip()
    marks  = (item.get("makers_marks_observed") or "").strip()
    era    = (item.get("probable_era") or "").strip()
    query  = (item.get("suggested_search_query") or "").strip()

    if maker and itype:
        # Shorten item_type to first 2 meaningful words
        itype_short = " ".join(itype.split()[:2])
        parts = [maker, itype_short]

        # Add a model/catalogue number if it's compact (e.g. "7837", "No. 1234")
        num_match = re.search(r"\b(\d{3,6})\b", marks)
        if num_match:
            parts.append(num_match.group(1))
        elif era and len(era.split()) <= 2:
            parts.append(era)

        return " ".join(parts)

    # No maker: trim suggested_search_query to first 4 words
    if query:
        return " ".join(query.split()[:4])

    return " ".join(itype.split()[:4]) if itype else "antique figurine"


# ── Claude web search ─────────────────────────────────────────────────────────

def _search_comps_via_claude(query: str) -> list[dict]:
    """
    Use Claude's web_search tool to find eBay sold listings.
    Returns a list of dicts: {price, date, title, url, condition}.
    """
    client = _get_client()

    prompt = (
        f'Search eBay completed/sold listings for "{query}". '
        "I need SOLD prices — items that have already sold, not active listings. "
        "Find 8–15 recent sold prices. For each result list:\n"
        "PRICE: $XX.XX\n"
        "DATE: Mon YYYY\n"
        "TITLE: <listing title>\n"
        "---\n"
        "Only include actual completed sales from within the last 18 months."
    )

    messages = [{"role": "user", "content": prompt}]

    for _turn in range(8):          # safety cap on the tool loop
        resp = client.messages.create(
            model=LISTING_MODEL,
            max_tokens=2048,
            tools=[WEB_SEARCH_TOOL],
            messages=messages,
        )

        # Collect any text blocks this turn
        text_parts = [
            block.text
            for block in resp.content
            if hasattr(block, "text") and block.text
        ]

        if resp.stop_reason == "end_turn":
            return _parse_prices_from_text("\n".join(text_parts))

        if resp.stop_reason == "tool_use":
            # Append assistant turn, then provide tool_result so Claude
            # can incorporate Anthropic's server-side search results.
            messages.append({"role": "assistant", "content": resp.content})
            tool_results = [
                {
                    "type": "tool_result",
                    "tool_use_id": block.id,
                    "content": "",      # Anthropic injects real results server-side
                }
                for block in resp.content
                if getattr(block, "type", None) == "tool_use"
            ]
            if tool_results:
                messages.append({"role": "user", "content": tool_results})
            # If we already have text (from a partial response), hold it
            continue

        # Any other stop reason — use whatever text we have
        if text_parts:
            return _parse_prices_from_text("\n".join(text_parts))
        break

    return []


def _parse_prices_from_text(text: str) -> list[dict]:
    """
    Extract individual sold-listing records from Claude's text response.
    Handles the structured "PRICE: / DATE: / TITLE:" format as well as
    looser inline price mentions.
    """
    listings = []

    # ── Structured block parser ───────────────────────────────────────────────
    # Split on "---" separators and try to parse each block
    blocks = re.split(r"-{2,}", text)
    for block in blocks:
        price = _extract_price(block)
        if price is None:
            continue

        date_m = re.search(
            r"DATE:\s*(.+?)(?:\n|$)",
            block, re.IGNORECASE
        )
        title_m = re.search(
            r"TITLE:\s*(.+?)(?:\n|$)",
            block, re.IGNORECASE
        )
        cond_m = re.search(
            r"CONDITION:\s*(.+?)(?:\n|$)",
            block, re.IGNORECASE
        )
        url_m = re.search(
            r"(https?://[^\s]+)",
            block
        )

        listings.append({
            "price": price,
            "date":      (date_m.group(1).strip()  if date_m  else ""),
            "title":     (title_m.group(1).strip() if title_m else ""),
            "condition": (cond_m.group(1).strip()  if cond_m  else ""),
            "url":       (url_m.group(1).strip()   if url_m   else ""),
        })

    if listings:
        return listings

    # ── Fallback: grab all standalone dollar amounts from text ────────────────
    price_lines = re.findall(
        r"\$\s*(\d{1,4}(?:\.\d{2})?)",
        text
    )
    for p_str in price_lines:
        try:
            price = float(p_str)
            if price > 0:
                listings.append({"price": price, "date": "", "title": "", "condition": "", "url": ""})
        except ValueError:
            pass

    return listings


def _extract_price(text: str) -> float | None:
    """Extract the first dollar-amount from a text block."""
    # "PRICE: $12.50" or "$12.50" or "12.50"
    m = re.search(
        r"PRICE:\s*\$?\s*([\d,]+\.?\d*)"
        r"|\$\s*([\d,]+\.?\d*)",
        text, re.IGNORECASE
    )
    if not m:
        return None
    raw = (m.group(1) or m.group(2) or "").replace(",", "")
    try:
        val = float(raw)
        return val if val > 0 else None
    except ValueError:
        return None


# ── Stats helpers (unchanged interface) ───────────────────────────────────────

def _percentile(values: list[float], pct: float) -> float:
    if not values:
        return 0.0
    sv = sorted(values)
    idx = (len(sv) - 1) * pct / 100
    lo, hi = int(idx), min(int(idx) + 1, len(sv) - 1)
    return sv[lo] + (sv[hi] - sv[lo]) * (idx - lo)


def filter_outliers(listings: list[dict]) -> list[dict]:
    prices = [l["price"] for l in listings]
    if len(prices) < 3:
        return listings
    median = _percentile(prices, 50)
    if median == 0:
        return listings
    return [l for l in listings if median / 3 <= l["price"] <= median * 3]


# ── Public entry point ────────────────────────────────────────────────────────

def compute_comps(item_id: str, query: str) -> dict:
    """
    Run comp research for item_id using the given query string.
    query should already be shortened via make_search_query() before calling.
    Writes results to DB and returns stats dict.
    """
    db.upsert_item(item_id, status="comps_pending")

    try:
        raw = _search_comps_via_claude(query)
        time.sleep(COMP_DELAY_SECONDS + random.uniform(0, 1.5))
    except Exception as exc:
        db.append_error(item_id, f"comps web-search error: {exc}")
        db.upsert_item(item_id, status="comps_failed")
        return {"error": str(exc)}

    filtered = filter_outliers(raw)
    prices   = [l["price"] for l in filtered]

    if not prices:
        db.upsert_item(
            item_id, status="comps_done",
            comp_sample_size=0, comp_confidence="no data found",
            comp_listings=json.dumps([]),
        )
        return {"sample_size": 0, "confidence": "no data found"}

    stats = {
        "low":         round(_percentile(prices, 25), 2),
        "median":      round(_percentile(prices, 50), 2),
        "high":        round(_percentile(prices, 75), 2),
        "sample_size": len(prices),
        "confidence":  "low (< 5 comps)" if len(prices) < 5 else "ok",
    }

    db.upsert_item(
        item_id,
        status="comps_done",
        comp_low=stats["low"],
        comp_median=stats["median"],
        comp_high=stats["high"],
        comp_sample_size=stats["sample_size"],
        comp_confidence=stats["confidence"],
        comp_listings=json.dumps(filtered[:20]),
    )
    db.log_run(item_id, "comps", "success",
               f"n={stats['sample_size']} median=${stats['median']}")
    return stats
