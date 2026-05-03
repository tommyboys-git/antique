"""Streamlit review UI. Run: python -m streamlit run review_app.py"""

import json
from pathlib import Path

import streamlit as st

import db
from config import OUTPUT_DIR
import export as export_module

st.set_page_config(page_title="Antique Pipeline Review", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
  .stApp { background-color: #1a1a1a; color: #e0e0e0; }
  .block-container { padding-top: 1rem; }
  .badge-reviewed   { background:#1a6634; color:#fff; padding:2px 10px; border-radius:12px; font-size:.8rem; }
  .badge-needs-review { background:#7a3a00; color:#fff; padding:2px 10px; border-radius:12px; font-size:.8rem; }
  .badge-lot        { background:#1a3a5c; color:#fff; padding:2px 10px; border-radius:12px; font-size:.8rem; }
  .badge-yardsale   { background:#4a2020; color:#fff; padding:2px 10px; border-radius:12px; font-size:.8rem; }
  .badge-pending    { background:#444; color:#aaa; padding:2px 10px; border-radius:12px; font-size:.8rem; }
</style>
""", unsafe_allow_html=True)


def load_data():
    db.init_db()
    return db.get_all_items(), db.get_all_lots()


# ── Sidebar ────────────────────────────────────────────────────────────────────
st.sidebar.title("Antique Pipeline")
view = st.sidebar.radio("View", ["Individual Items", "Lots"], label_visibility="collapsed")

all_items, all_lots = load_data()

if st.sidebar.button("Export All to Excel"):
    out_path = export_module.export_to_excel()
    st.sidebar.success(f"Exported: {out_path.name}")

st.sidebar.markdown("---")

# ══════════════════════════════════════════════════════════════════════════════
# VIEW: INDIVIDUAL ITEMS
# ══════════════════════════════════════════════════════════════════════════════
if view == "Individual Items":
    filter_status = st.sidebar.selectbox(
        "Filter",
        ["all", "identified", "comps_done", "listing_done", "reviewed",
         "failed", "needs_review", "in_a_lot", "yard_sale"],
    )

    if filter_status == "needs_review":
        items = [i for i in all_items if i.get("needs_review") and not i.get("lot_id")]
    elif filter_status == "in_a_lot":
        items = [i for i in all_items if i.get("lot_id")]
    elif filter_status == "yard_sale":
        items = [i for i in all_items if i.get("yard_sale")]
    elif filter_status == "all":
        items = all_items
    else:
        items = [i for i in all_items if i.get("status") == filter_status]

    if not items:
        st.info("No items match this filter.")
        st.stop()

    item_labels = []
    for i in items:
        badge = ""
        if i.get("lot_id"):     badge = " [LOT]"
        elif i.get("yard_sale"): badge = " [YARD]"
        elif i.get("needs_review"): badge = " [REVIEW]"
        item_labels.append(f"{i['item_id']}{badge} [{i.get('status','')}]")

    selected_label = st.sidebar.radio("Select item", item_labels, label_visibility="collapsed")
    selected_idx   = item_labels.index(selected_label)
    item = items[selected_idx]
    st.sidebar.caption(f"{len(items)} items shown")

    # ── Main panel ──
    item_id      = item["item_id"]
    photo_paths  = json.loads(item.get("photo_paths") or "[]")
    style_keywords = json.loads(item.get("style_keywords") or "[]")
    tags         = json.loads(item.get("tags") or "[]")
    comp_listings = json.loads(item.get("comp_listings") or "[]")
    is_lotted    = bool(item.get("lot_id"))
    is_yard_sale = bool(item.get("yard_sale"))

    col_photos, col_data = st.columns([1, 2])

    with col_photos:
        st.subheader("Photos")
        for p in photo_paths:
            path = Path(p)
            if not path.exists():
                from config import PROCESSED_PHOTOS_DIR
                path = PROCESSED_PHOTOS_DIR / path.name
            if path.exists():
                st.image(str(path), use_container_width=True)
            else:
                st.caption(f"(missing: {Path(p).name})")

        if is_lotted:
            lot = db.get_lot(item["lot_id"])
            st.markdown(f'<span class="badge-lot">In lot: {lot["lot_name"] if lot else item["lot_id"]}</span>',
                        unsafe_allow_html=True)
        if is_yard_sale:
            st.markdown('<span class="badge-yardsale">Yard Sale — FB Marketplace</span>', unsafe_allow_html=True)
        if item.get("needs_review"):
            st.markdown('<span class="badge-needs-review">Needs Review</span>', unsafe_allow_html=True)
            st.caption(f"Confidence: {item.get('confidence_score')}% — {item.get('confidence_reasoning')}")

    with col_data:
        is_reviewed = item.get("status") == "reviewed"
        st.subheader(item_id)
        st.caption(f"Status: **{item.get('status')}**")

        with st.expander("Identification", expanded=True):
            item_type = st.text_input("Item Type", value=item.get("item_type") or "", disabled=is_reviewed)
            c1, c2 = st.columns(2)
            maker    = c1.text_input("Maker",    value=item.get("probable_maker") or "", disabled=is_reviewed)
            era      = c2.text_input("Era",      value=item.get("probable_era") or "", disabled=is_reviewed)
            material = c1.text_input("Material", value=item.get("material") or "", disabled=is_reviewed)
            marks    = c2.text_input("Maker's Marks", value=item.get("makers_marks_observed") or "", disabled=is_reviewed)
            condition   = st.text_area("Condition", value=item.get("condition_notes") or "", height=80, disabled=is_reviewed)
            keywords    = st.text_input("Style Keywords", value=", ".join(style_keywords), disabled=is_reviewed)
            search_query = st.text_input("Search Query", value=item.get("suggested_search_query") or "", disabled=is_reviewed)

        with st.expander("Pricing & Comps"):
            c1, c2, c3 = st.columns(3)
            c1.metric("Low (25th)",    f"${item.get('comp_low') or 0:.2f}")
            c2.metric("Median (50th)", f"${item.get('comp_median') or 0:.2f}")
            c3.metric("High (75th)",   f"${item.get('comp_high') or 0:.2f}")
            st.caption(f"n={item.get('comp_sample_size') or 0} — {item.get('comp_confidence') or 'not run'}")
            for l in comp_listings[:5]:
                st.markdown(f"- [{l.get('title','')} — ${l.get('price',0):.2f}]({l.get('url','')})")

        with st.expander("Listing Drafts"):
            ebay_title = st.text_input("eBay Title (80)",  value=item.get("ebay_title") or "", disabled=is_reviewed)
            etsy_title = st.text_input("Etsy Title (140)", value=item.get("etsy_title") or "", disabled=is_reviewed)
            fb_title   = st.text_input("FB Title",         value=item.get("fb_marketplace_title") or "", disabled=is_reviewed)
            description = st.text_area("Description", value=item.get("description") or "", height=150, disabled=is_reviewed)
            c1, c2 = st.columns(2)
            price    = c1.number_input("Price ($)", value=float(item.get("suggested_price") or 0), min_value=0.0, step=0.50, disabled=is_reviewed)
            platform_opts = ["eBay", "Etsy", "Facebook Marketplace"]
            cur_plat = item.get("suggested_platform", "eBay")
            platform = c1.selectbox("Platform", platform_opts,
                                    index=platform_opts.index(cur_plat) if cur_plat in platform_opts else 0,
                                    disabled=is_reviewed)
            category = c2.text_input("eBay Category", value=item.get("category_ebay") or "", disabled=is_reviewed)
            tag_str  = st.text_input("Tags", value=", ".join(tags), disabled=is_reviewed)

        if not is_reviewed:
            c1, c2 = st.columns(2)
            if c1.button("Save Edits", use_container_width=True):
                db.upsert_item(item_id,
                    item_type=item_type, probable_maker=maker, probable_era=era,
                    material=material, makers_marks_observed=marks, condition_notes=condition,
                    style_keywords=json.dumps([k.strip() for k in keywords.split(",") if k.strip()]),
                    suggested_search_query=search_query,
                    ebay_title=ebay_title[:80], etsy_title=etsy_title[:140],
                    fb_marketplace_title=fb_title[:60], description=description,
                    suggested_price=price, suggested_platform=platform,
                    category_ebay=category,
                    tags=json.dumps([t.strip() for t in tag_str.split(",") if t.strip()]),
                )
                st.success("Saved.")
            if c2.button("Approve & Lock", type="primary", use_container_width=True):
                db.upsert_item(item_id,
                    item_type=item_type, probable_maker=maker, probable_era=era,
                    material=material, makers_marks_observed=marks, condition_notes=condition,
                    style_keywords=json.dumps([k.strip() for k in keywords.split(",") if k.strip()]),
                    ebay_title=ebay_title[:80], etsy_title=etsy_title[:140],
                    fb_marketplace_title=fb_title[:60], description=description,
                    suggested_price=price, suggested_platform=platform, category_ebay=category,
                    tags=json.dumps([t.strip() for t in tag_str.split(",") if t.strip()]),
                    status="reviewed", needs_review=0,
                )
                st.rerun()
        else:
            st.success("Item approved and locked.")
            if st.button("Unlock for Editing"):
                db.upsert_item(item_id, status="listing_done")
                st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# VIEW: LOTS
# ══════════════════════════════════════════════════════════════════════════════
else:
    if not all_lots:
        st.title("Lots")
        st.info("No lots created yet. Run `python pipeline.py --phase lotting` first.")
        st.stop()

    lot_labels = [f"{l['lot_id']} — {l['lot_name']} (${l.get('lot_price',0):.2f})" for l in all_lots]
    selected_label = st.sidebar.radio("Select lot", lot_labels, label_visibility="collapsed")
    selected_idx   = lot_labels.index(selected_label)
    lot = all_lots[selected_idx]
    st.sidebar.caption(f"{len(all_lots)} lots")

    lot_id    = lot["lot_id"]
    item_ids  = json.loads(lot.get("item_ids") or "[]")
    lot_items = [db.get_item(iid) for iid in item_ids if db.get_item(iid)]
    is_approved = lot.get("status") == "reviewed"

    col_photos, col_data = st.columns([1, 2])

    with col_photos:
        st.subheader("Composite")
        thumb = Path(lot.get("thumbnail_path") or "")
        if thumb.exists():
            st.image(str(thumb), use_container_width=True)
        else:
            st.caption("(no composite thumbnail)")

        st.markdown("**Items in this lot:**")
        for li in lot_items:
            price = li.get("suggested_price") or 0
            st.caption(f"• {li['item_id']} — {li.get('item_type','')} (${price:.2f})")

        item_sum = sum((li.get("suggested_price") or 0) for li in lot_items)
        st.caption(f"Individual sum: ${item_sum:.2f}")

    with col_data:
        st.subheader(lot.get("lot_name", lot_id))
        st.caption(f"Status: **{lot.get('status')}** | Platform: **{lot.get('suggested_platform')}**")

        lot_name   = st.text_input("Lot Name",  value=lot.get("lot_name") or "", disabled=is_approved)
        lot_price  = st.number_input("Lot Price ($)", value=float(lot.get("lot_price") or 0),
                                     min_value=0.0, step=0.50, disabled=is_approved)
        platform_opts = ["eBay", "Etsy", "Facebook Marketplace"]
        cur_plat = lot.get("suggested_platform", "eBay")
        platform = st.selectbox("Platform", platform_opts,
                                index=platform_opts.index(cur_plat) if cur_plat in platform_opts else 0,
                                disabled=is_approved)
        lot_desc = st.text_area("Lot Description", value=lot.get("lot_description") or "",
                                height=200, disabled=is_approved)

        if not is_approved:
            c1, c2 = st.columns(2)
            if c1.button("Save Lot Edits", use_container_width=True):
                db.upsert_lot(lot_id, lot_name=lot_name, lot_price=lot_price,
                              suggested_platform=platform, lot_description=lot_desc)
                st.success("Saved.")
            if c2.button("Approve Lot", type="primary", use_container_width=True):
                db.upsert_lot(lot_id, lot_name=lot_name, lot_price=lot_price,
                              suggested_platform=platform, lot_description=lot_desc,
                              status="reviewed")
                st.rerun()
        else:
            st.success("Lot approved.")
            if st.button("Unlock Lot"):
                db.upsert_lot(lot_id, status="draft")
                st.rerun()

        # Quick breakdown
        with st.expander("Item details"):
            for li in lot_items:
                st.markdown(f"**{li['item_id']}** — {li.get('item_type')} | "
                            f"{li.get('probable_era')} | ${li.get('suggested_price',0):.2f}")
                st.caption(li.get("condition_notes", "")[:150])
