"""Main orchestrator. Run: python pipeline.py [--phase identify|comps|listing|lotting|export|master_export|all]"""

import argparse
import json
import re
import shutil
import sys
from pathlib import Path

from dotenv import load_dotenv

BASE_DIR = Path(__file__).parent
load_dotenv(BASE_DIR.parent / ".env.txt", override=True)
load_dotenv(BASE_DIR / ".env", override=True)

import db
import identify as id_module
import comps as comps_module
import listing as listing_module
import lotting as lotting_module
import export as export_module
import master_export as master_export_module
from config import (
    BATCH_SIZE, CONFIDENCE_THRESHOLD, INPUT_PHOTOS_DIR,
    PROCESSED_PHOTOS_DIR, FAILED_PHOTOS_DIR, SUPPORTED_EXTENSIONS,
    SKIP_COMPS_UNDER_VALUE,
)


def group_photos(directory: Path) -> dict[str, list[Path]]:
    """
    Group photo files into items. Supports two naming conventions:
      1. item_001_front.jpg  → grouped by 'item_001' prefix
      2. Cookie Jar Elephant 1.png → grouped by base name without trailing number
    Returns {item_id: [path, path, ...]}
    """
    groups: dict[str, list[Path]] = {}
    for path in sorted(directory.iterdir()):
        if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            continue
        stem = path.stem

        # Convention 1: item_XXX_ prefix
        m1 = re.match(r"^(item_\d+)", stem, re.IGNORECASE)
        if m1:
            key = m1.group(1).lower()
        else:
            # Convention 2: strip trailing space+digits (e.g. "Elephant 1" → "Elephant")
            key = re.sub(r"\s+\d+$", "", stem).strip()

        item_id = re.sub(r"[^\w]+", "_", key).strip("_").lower()
        groups.setdefault(item_id, []).append(path)

    # Sort photos within each group so the "primary" (no trailing number) comes first
    for item_id, paths in groups.items():
        groups[item_id] = sorted(paths, key=lambda p: (
            bool(re.search(r"\s\d+$", p.stem)),  # numbered photos last
            p.stem,
        ))

    return groups


def run_identify(limit: int | None = None):
    groups = group_photos(INPUT_PHOTOS_DIR)
    if not groups:
        print(f"No photos found in {INPUT_PHOTOS_DIR}")
        return

    processed = 0
    for item_id, photo_paths in groups.items():
        if limit and processed >= limit:
            break

        existing = db.get_item(item_id)
        if existing and existing["status"] not in ("pending", "failed", "identifying"):
            print(f"  SKIP {item_id} (status={existing['status']})")
            continue

        db.upsert_item(item_id, photo_paths=json.dumps([str(p) for p in photo_paths]))
        print(f"  ID   {item_id} ({len(photo_paths)} photos)...", end=" ", flush=True)

        try:
            result = id_module.identify_item(item_id, photo_paths)
            flag = " [NEEDS REVIEW]" if result["confidence_score"] < CONFIDENCE_THRESHOLD else ""
            print(f"OK  confidence={result['confidence_score']}{flag}")
            summary = f"       -> {result['item_type']} | {result['probable_era']} | {result['probable_maker']}"
            print(summary.encode("ascii", "replace").decode("ascii"))
            for p in photo_paths:
                dest = PROCESSED_PHOTOS_DIR / p.name
                shutil.move(str(p), str(dest))
        except Exception as exc:
            print(f"FAILED: {exc}")
            for p in photo_paths:
                dest = FAILED_PHOTOS_DIR / p.name
                try:
                    shutil.move(str(p), str(dest))
                except Exception:
                    pass

        processed += 1

    print(f"\nIdentification complete. {processed} items processed.")


def run_comps():
    items = db.get_items_by_status("identified")
    if not items:
        print("No items ready for comps (status=identified).")
        return

    for item in items:
        item_id = item["item_id"]
        # Build a tight 3-5 keyword query from item metadata
        query = comps_module.make_search_query(item)
        median = item.get("comp_median") or 0

        if not query:
            print(f"  SKIP {item_id} — no search query")
            continue
        if median and median < SKIP_COMPS_UNDER_VALUE:
            print(f"  SKIP {item_id} — estimated value below threshold")
            continue

        print(f"  COMPS {item_id}: '{query}'...", end=" ", flush=True)
        try:
            stats = comps_module.compute_comps(item_id, query)
            print(f"OK  n={stats.get('sample_size',0)} median=${stats.get('median',0)}")
        except Exception as exc:
            print(f"FAILED: {exc}")


def run_listing():
    for status in ("comps_done", "comps_failed"):
        items = db.get_items_by_status(status)
        for item in items:
            item_id = item["item_id"]
            print(f"  LISTING {item_id}...", end=" ", flush=True)
            try:
                result = listing_module.generate_listing(item_id)
                print(f"OK  platform={result.get('suggested_platform')} price=${result.get('suggested_price')}")
            except Exception as exc:
                print(f"FAILED: {exc}")


def run_export():
    out = export_module.export_to_excel()
    print(f"Exported to: {out}")


def run_master_export():
    out = master_export_module.export()
    print(f"Master export saved to: {out}")


def main():
    parser = argparse.ArgumentParser(description="Antique Research Pipeline")
    parser.add_argument(
        "--phase",
        choices=["identify", "comps", "listing", "lotting", "export", "master_export", "all"],
        default="identify",
        help="Which phase to run (default: identify)",
    )
    parser.add_argument(
        "--batch",
        type=int,
        default=BATCH_SIZE,
        help="Max items to process in this run",
    )
    args = parser.parse_args()

    db.init_db()

    print(f"\n=== Antique Research Pipeline — phase: {args.phase} ===\n")

    if args.phase in ("identify", "all"):
        print("--- Phase 1: Identification ---")
        run_identify(limit=args.batch if args.phase == "identify" else None)

    if args.phase in ("comps", "all"):
        print("\n--- Phase 2: Comp Research ---")
        try:
            run_comps()
        finally:
            comps_module.close_browser()

    if args.phase in ("listing", "all"):
        print("\n--- Phase 3: Listing Drafts ---")
        run_listing()

    if args.phase in ("lotting", "all"):
        print("\n--- Phase 3.5: Smart Lotting ---")
        lotting_module.run_lotting()

    if args.phase in ("export", "all"):
        print("\n--- Phase 4: Export ---")
        run_export()

    if args.phase in ("master_export", "all"):
        print("\n--- Phase 5: Master Listings Export ---")
        run_master_export()


if __name__ == "__main__":
    main()
