# antique
Claude Code made this for me to research items in my house and give me the details to list it on the recommended selling platform.
Turns photos of antique/vintage items into a research-ready spreadsheet with AI identification,
eBay sold-comp pricing, platform recommendations, and listing drafts.

---

## Setup (one time)

### 1. Install Python 3.11+
Download from https://python.org — check "Add to PATH" during install.

### 2. Open a terminal in this folder
Right-click the `antique-pipeline` folder → "Open in Terminal" (or use Command Prompt).

### 3. Install dependencies
```
pip install -r requirements.txt
```

### 4. Add your API key
Your `.env` file is already populated. Keep it private.

---

## Dropping in Photos

Place photos in the `input_photos/` folder. Two naming styles work:

**Style A — descriptive names (what you already have):**
```
Cookie Jar Elephant.png
Cookie Jar Elephant 1.png
Cookie Jar Elephant 2.png
```
Photos with the same base name (minus the trailing number) are treated as the same item.

**Style B — numbered prefix:**
```
item_001_front.jpg
item_001_back.jpg
item_001_mark.jpg
```

---

## Running the Pipeline

### Phase 1 — Identify items (start here)
```
python pipeline.py --phase identify
```
Processes everything in `input_photos/`. Adds `--batch 5` to do only 5 at a time.

### Phase 2 — Fetch eBay sold comps
```
python pipeline.py --phase comps
```
Runs after identify. Scrapes eBay for sold listings.

### Phase 3 — Generate listing drafts
```
python pipeline.py --phase listing
```

### Phase 4 — Export to Excel
```
python pipeline.py --phase export
```
Output file appears in `output/`.

### Run all phases at once
```
python pipeline.py --phase all
```

### Resume after interruption
Just re-run the same command. Items already completed are skipped automatically.

---

## Review UI

```
streamlit run review_app.py
```

Opens a browser with your items side-by-side: photos on the left, editable AI data on the right.
Click **Approve & Lock** when an item is ready.

---

## Using Your Existing Photos

Your photos are already in `C:\Users\Owner\Pictures\Storefront\Edited\`.
Copy or move 5 items worth of photos into `input_photos/` to test:

```
# Example: copy all Cookie Jar Elephant photos
copy "..\Cookie Jar Elephant*.png" input_photos\
```

---

## Output

- `output/antique_pipeline_TIMESTAMP.xlsx` — main spreadsheet
- Sheet 1: full data with thumbnails
- Sheet 2: eBay File Exchange format for bulk upload
- `db.sqlite` — intermediate state (don't delete between runs)

---

## Troubleshooting

| Problem | Fix |
|---|---|
| `ModuleNotFoundError` | Run `pip install -r requirements.txt` again |
| API key error | Check `.env` file has your key |
| No photos found | Make sure photos are in `input_photos/` |
| eBay scraping blocked | Normal — pipeline logs and continues |
| Low confidence items | Review them in the Streamlit UI |
