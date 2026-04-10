# Supalai City Resort Rama 8 — Resale Price Tracker

Automated tracker for resale condo listings at **Supalai City Resort Rama 8**
(ศุภาลัย ซิตี้ รีสอร์ท พระราม 8), Bangkok.

---

## Memory / Context for Future Chat Sessions

> **Read this first if you are starting a new chat.**

### Owner's property
| Field       | Value |
|-------------|-------|
| Project     | Supalai City Resort Rama 8 (ศุภาลัย ซิตี้ รีสอร์ท พระราม 8) |
| Size        | ~35 sqm |
| Bedrooms    | 1 |
| Goal        | Track resale market price to know the current market value of this unit |

### What this repo does
- `scraper.py` — headless-browser scraper (Playwright + Chromium) that hits 5 Thai property portals:
  DDproperty, Fazwaz, Hipflat, Baania, PropertyScout
- Deduplicates: same unit listed on multiple sites → merged into one row
- Deduplicates across runs: same unit at same price already recorded → `last_seen` updated, not re-added
- Saves cumulative history to `data/property_history.json` (committed to git = memory across sessions)
- Generates `Supalai_Rama8_Price_Track.xlsx` with 3 sheets:
  1. **All Listings** — every scraped listing ever (most recent first)
  2. **My Unit (~35sqm 1BR)** — filtered view matching owner's unit specs (±5 sqm, 1 BR)
  3. **Run Summary** — per-run statistics and price ranges

### Key files
| File | Purpose |
|------|---------|
| `scraper.py` | Main script — run this to collect data |
| `requirements.txt` | pip dependencies |
| `data/property_history.json` | **Cumulative database** — committed to git |
| `data/raw/` | Raw per-run snapshots — gitignored |
| `Supalai_Rama8_Price_Track.xlsx` | Output Excel — regenerated on every run |

### Data schema (`property_history.json`)
Each record has these fields:
- `first_seen` — date this listing was first collected (YYYY-MM-DD)
- `last_seen` — most recent date it was still active
- `listed_month_year` — when the seller listed it (YYYY-MM, extracted from site)
- `price_thb` — asking price in Thai Baht (integer)
- `price_per_sqm` — price ÷ size
- `size_sqm` — unit size in square metres
- `floor` — floor number
- `building` — building/tower name (if multi-building project)
- `bedrooms` / `bathrooms`
- `furnishing` — "Fully Furnished" / "Partially Furnished" / "Unfurnished" / "Unknown"
- `view` — view direction if available
- `title` — listing title (first 120 chars)
- `source` — site(s) where found (pipe-separated if merged)
- `url` — listing URL(s)
- `fingerprint` — dedup key: `s{size}_f{floor}_br{beds}_ba{baths}`

### Deduplication logic
1. **Within a run**: same fingerprint from multiple sites → merged (richest record kept, all URLs joined)
2. **Across runs**: same fingerprint + price within ±50,000 THB already in history for same month → skip; otherwise add

---

## Setup

```bash
# 1. Clone / pull this repo
git clone <repo-url>
cd Supalai-rama-8-price-track

# 2. Install Python dependencies
pip3 install -r requirements.txt

# 3. Install Playwright browser (first time only)
playwright install chromium

# 4. Run the scraper
python3 scraper.py
```

The Excel file `Supalai_Rama8_Price_Track.xlsx` will be created/updated in the project root.

---

## Run Frequency Recommendation

| Interval | Pros | Cons |
|----------|------|------|
| **Every 2 weeks** ✅ (recommended) | Captures new listings and price drops without redundant data; manageable file size | May miss very short listings |
| Monthly | Minimal data, easy to manage | May miss listings that sold or changed price mid-month |
| Weekly | Best trend resolution | More duplicate suppression needed; sites may throttle |

**Recommendation: run every 2 weeks (1st and 15th of each month).**
The Thai resale condo market moves slowly enough that bi-weekly is more than sufficient
to catch price trends. If you're actively considering selling, switch to weekly.

---

## Troubleshooting

| Problem | Fix |
|---------|-----|
| `0 cards found` on a site | The site likely updated its HTML — open the URL in a real browser, inspect the listing cards, and update the CSS selectors in the matching `scrape_*` function in `scraper.py` |
| `403 / blocked` | Try increasing `wait_for_timeout` or add `page.wait_for_load_state("networkidle")` |
| Excel file locked | Close the file in Excel before running |
| `playwright install` needed | Run `playwright install chromium` once after pip install |

---

## Adding More Sites

Each scraper follows the same pattern:
1. `page.goto(url, …)`
2. Find listing cards with `_extract_cards(page, [selectors])`
3. Extract text, parse with helpers (`parse_price`, `parse_size`, etc.)
4. Append dict to `records`

Copy any existing `scrape_*` function and adapt selectors.

---

## Status Log

| Date | Action | Notes |
|------|--------|-------|
| 2026-04-10 | Initial setup | Repo created; 5 scrapers written; Excel with 3 sheets |
