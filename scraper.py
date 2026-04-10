#!/usr/bin/env python3
"""
Supalai City Resort Rama 8 — Resale Price Tracker
===================================================
Scrapes DDproperty, Fazwaz, Hipflat, Baania, and PropertyScout
for resale condo listings, deduplicates across sites, and saves
a running history to Excel.

My unit: ~35 sqm, 1 bedroom — tracked in a separate sheet.

Usage:
    python3 scraper.py

Run frequency recommendation: every 2 weeks (see README.md)
"""

import json
import re
import time
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from playwright.sync_api import sync_playwright, Page

# ─────────────────────────────────────────
#  CONFIG
# ─────────────────────────────────────────
CONDO_EN = "Supalai City Resort Rama 8"
CONDO_TH = "ศุภาลัย ซิตี้ รีสอร์ท พระราม 8"
CONDO_TH_SHORT = "ศุภาลัย"   # fallback keyword

MY_SIZE_SQM = 35
MY_SIZE_TOL = 5     # ±5 sqm  →  30–40 sqm
MY_BEDROOMS = 1

DATA_DIR     = Path("data")
RAW_DIR      = DATA_DIR / "raw"
HISTORY_FILE = DATA_DIR / "property_history.json"
EXCEL_FILE   = Path("Supalai_Rama8_Price_Track.xlsx")

COLUMNS = [
    "first_seen", "last_seen", "listed_month_year",
    "price_thb", "price_per_sqm",
    "size_sqm", "floor", "building",
    "bedrooms", "bathrooms", "furnishing", "view",
    "title", "source", "url", "fingerprint",
]

# ─────────────────────────────────────────
#  PARSING HELPERS
# ─────────────────────────────────────────

def parse_price(text: str) -> int | None:
    """Extract integer THB from various formats: '2.5 ล้าน', '2,500,000', '฿2.5M'."""
    if not text:
        return None
    text = str(text).strip()
    # Millions written as X.X ล้าน or X.XM
    m = re.search(r'([\d,.]+)\s*(?:ล้าน|M\b|million)', text, re.I)
    if m:
        return int(float(m.group(1).replace(',', '')) * 1_000_000)
    # Plain number (may have commas)
    m = re.search(r'[\d,]+', text.replace('.', ''))  # remove decimal separators
    if m:
        val = int(m.group(0).replace(',', ''))
        # Sanity: condo prices should be between 500k and 50M THB
        if 500_000 <= val <= 50_000_000:
            return val
    # Re-try allowing decimal (e.g. 2500000.00)
    m = re.search(r'[\d,]+\.?\d*', text)
    if m:
        val = int(float(m.group(0).replace(',', '')))
        if 500_000 <= val <= 50_000_000:
            return val
    return None


def parse_size(text: str) -> float | None:
    """Extract sqm from '35 sqm', '35.5 ตร.ม.', etc."""
    if not text:
        return None
    m = re.search(r'([\d]+\.?[\d]*)\s*(?:sqm|sq\.?m|ตร\.?ม|ตารางเมตร)', str(text), re.I)
    return float(m.group(1)) if m else None


def parse_int(text: str) -> int | None:
    """Extract first integer from a string."""
    if not text:
        return None
    m = re.search(r'\d+', str(text))
    return int(m.group(0)) if m else None


def parse_furnishing(text: str) -> str:
    """Normalise furnishing status to one of three values."""
    if not text:
        return "Unknown"
    t = str(text).lower()
    # Check unfurnished BEFORE furnished to avoid substring collision
    if any(w in t for w in ['unfurnish', 'empty', 'ไม่มีเฟ', 'เปล่า']):
        return "Unfurnished"
    if any(w in t for w in ['partial', 'partly', 'semi', 'บางส่วน']):
        return "Partially Furnished"
    if any(w in t for w in ['fully', 'full furnished', 'เฟอร์นิเจอร์ครบ', 'furnished']):
        return "Fully Furnished"
    return "Unknown"


def parse_relative_date(text: str) -> str:
    """
    Convert 'X days ago', 'X months ago', or explicit dates to 'YYYY-MM'.
    Falls back to current month.
    """
    now = datetime.now()
    if not text:
        return now.strftime("%Y-%m")
    text = str(text).lower().strip()

    # Explicit ISO / Thai date patterns
    m = re.search(r'(20\d{2})[/-](\d{1,2})', text)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}"

    # Month name patterns: "Jan 2024", "January 2024"
    months = {
        'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
        'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12,
        'มกราคม': 1, 'กุมภาพันธ์': 2, 'มีนาคม': 3, 'เมษายน': 4,
        'พฤษภาคม': 5, 'มิถุนายน': 6, 'กรกฎาคม': 7, 'สิงหาคม': 8,
        'กันยายน': 9, 'ตุลาคม': 10, 'พฤศจิกายน': 11, 'ธันวาคม': 12,
    }
    for name, num in months.items():
        if name in text:
            yr_m = re.search(r'20\d{2}', text)
            yr = int(yr_m.group(0)) if yr_m else now.year
            return f"{yr}-{num:02d}"

    # Relative: X days ago
    m = re.search(r'(\d+)\s*day', text)
    if m:
        from datetime import timedelta
        d = now - timedelta(days=int(m.group(1)))
        return d.strftime("%Y-%m")

    # Relative: X weeks ago
    m = re.search(r'(\d+)\s*week', text)
    if m:
        from datetime import timedelta
        d = now - timedelta(weeks=int(m.group(1)))
        return d.strftime("%Y-%m")

    # Relative: X months ago
    m = re.search(r'(\d+)\s*month', text)
    if m:
        months_back = int(m.group(1))
        month = now.month - months_back
        year = now.year
        while month <= 0:
            month += 12
            year -= 1
        return f"{year}-{month:02d}"

    return now.strftime("%Y-%m")


def is_relevant(title: str) -> bool:
    """Check if a listing title mentions our condo."""
    if not title:
        return False
    t = title.lower()
    return (
        "supalai city resort rama 8" in t
        or "supalai city resort" in t
        or CONDO_TH in title
        or CONDO_TH_SHORT in title
        or "rama 8" in t
        or "พระราม 8" in title
        or "พระราม8" in title
    )


# ─────────────────────────────────────────
#  FINGERPRINT & DEDUPLICATION
# ─────────────────────────────────────────

def make_fingerprint(size_sqm, floor, bedrooms, bathrooms) -> str:
    """
    Stable identifier for a physical unit.
    Size rounded to nearest 0.5 sqm to absorb minor discrepancies.
    """
    sz = round(float(size_sqm or 0) * 2) / 2
    return f"s{sz}_f{floor}_br{bedrooms}_ba{bathrooms}"


def deduplicate_within_run(records: list[dict]) -> list[dict]:
    """
    Merge records with the same fingerprint from different sites in one run.
    Keeps the record with the most non-null fields; appends all source URLs.
    """
    groups: dict[str, list[dict]] = {}
    for r in records:
        fp = r.get("fingerprint", "")
        groups.setdefault(fp, []).append(r)

    result = []
    for fp, grp in groups.items():
        if len(grp) == 1:
            result.append(grp[0])
        else:
            # pick the richest record
            best = max(grp, key=lambda r: sum(1 for v in r.values() if v is not None))
            best["source"] = " | ".join(sorted({r["source"] for r in grp}))
            best["url"]    = " | ".join({r["url"] for r in grp if r.get("url")})
            result.append(best)
    return result


def is_new_or_changed(record: dict, history: list[dict]) -> bool:
    """
    Return True if this record should be added to history.
    Skips if same fingerprint + price within ±50,000 THB already exists
    (same active listing, still on the market at same price).
    Returns True (add) if price changed more than ±50k or it's a brand-new unit.
    """
    fp    = record.get("fingerprint", "")
    price = record.get("price_thb")

    for h in history:
        if h.get("fingerprint") == fp:
            h_price = h.get("price_thb")
            if h_price and price and abs(h_price - price) <= 50_000:
                return False   # same listing, same price → skip
    return True   # new fingerprint or price changed


# ─────────────────────────────────────────
#  BROWSER HELPERS
# ─────────────────────────────────────────

# JavaScript injected into the page to find listing cards.
# Works on React/Next.js sites where class names are dynamically hashed,
# by doing substring matching on class names and inspecting rendered text.
_JS_FIND_CARDS = """
() => {
    const isCard = (el) => {
        const text = el.innerText || el.textContent || '';
        if (text.length < 80 || text.length > 5000) return false;
        if (!el.querySelector('a[href]'))             return false;
        return /\\d[\\d,]{4,}/.test(text);  // contains a 5+ digit number (price)
    };

    const patterns = [
        '[class*="card"]',     '[class*="Card"]',
        '[class*="listing"]',  '[class*="Listing"]',
        '[class*="property"]', '[class*="Property"]',
        '[class*="result"]',   '[class*="Result"]',
        '[class*="item"]',     '[class*="Item"]',
        'article',             'li[class]',
    ];

    let best = [];
    for (const p of patterns) {
        try {
            const found = Array.from(document.querySelectorAll(p)).filter(isCard);
            if (found.length > best.length) best = found;
        } catch(e) {}
    }

    // Keep only the most-specific cards (remove ancestors of other cards)
    const final = best.filter(el =>
        !best.some(other => other !== el && el.contains(other))
    );

    return final.slice(0, 60).map(el => ({
        text: (el.innerText || el.textContent || '').substring(0, 1500).trim(),
        href: (el.querySelector('a[href]') || {}).href || '',
    }));
}
"""


def _load_page(page: Page, url: str, extra_wait_ms: int = 5000):
    """Navigate, wait for network to settle, then scroll to trigger lazy loads."""
    page.goto(url, wait_until="domcontentloaded", timeout=45_000)
    try:
        page.wait_for_load_state("networkidle", timeout=12_000)
    except Exception:
        pass   # networkidle timeout is fine — proceed anyway
    page.wait_for_timeout(extra_wait_ms)
    # Log page title + final URL so we can see if we were redirected / got a CAPTCHA
    try:
        title   = page.title()
        final   = page.url
        body_snippet = (page.evaluate("document.body.innerText") or "")[:300].replace("\n", " ")
        print(f"    title={title!r}  url={final}")
        print(f"    body[0:300]={body_snippet!r}")
    except Exception:
        pass
    # Scroll halfway then to bottom to trigger lazy-loaded cards
    try:
        page.evaluate("window.scrollTo(0, document.body.scrollHeight * 0.5)")
        page.wait_for_timeout(800)
        page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        page.wait_for_timeout(800)
    except Exception:
        pass


def _js_cards(page: Page) -> list[dict]:
    """Run the JS card-finder in the live browser DOM; returns list of {text, href}."""
    try:
        cards = page.evaluate(_JS_FIND_CARDS) or []
        print(f"    js_cards found: {len(cards)}")
        return cards
    except Exception as e:
        print(f"    js_cards error: {e}")
        return []


def _save_debug_html(page: Page, site_name: str):
    """Save page HTML for debugging when no cards are found."""
    try:
        debug_dir = DATA_DIR / "debug"
        debug_dir.mkdir(parents=True, exist_ok=True)
        path = debug_dir / f"{site_name}.html"   # fixed name → overwrites each run
        path.write_text(page.content(), encoding="utf-8")
        print(f"    Debug HTML saved → {path}")
    except Exception:
        pass


def _parse_card(card: dict, source: str, run_date: str) -> dict | None:
    """
    Parse a single card's text blob into a structured record.
    Returns None if no valid price can be found.
    """
    text = card.get("text", "")
    href = card.get("href", "")

    # ── Price ──────────────────────────────────────────────────
    price = None
    # Thai million: "2.5 ล้าน" or "2.5M"
    m = re.search(r'([\d,]+\.?\d*)\s*(?:ล้าน|M\b)', text)
    if m:
        price = int(float(m.group(1).replace(',', '')) * 1_000_000)
    if not price:
        # Plain numbers ≥ 6 digits (≥ 100,000 THB) — take the first plausible one
        for raw in re.findall(r'[\d,]+', text):
            val = int(raw.replace(',', ''))
            if 500_000 <= val <= 50_000_000:
                price = val
                break
    if not price:
        return None

    # ── Size ───────────────────────────────────────────────────
    size = parse_size(text)

    # ── Floor ──────────────────────────────────────────────────
    fm = re.search(r'(?:floor|ชั้น|fl\.?)[^\d]*(\d+)', text, re.I)
    floor = int(fm.group(1)) if fm else None

    # ── Bedrooms / Bathrooms ───────────────────────────────────
    bm  = re.search(r'(\d+)\s*(?:bed(?:room)?s?|br\b|ห้องนอน)', text, re.I)
    bam = re.search(r'(\d+)\s*(?:bath(?:room)?s?|ba\b|ห้องน้ำ)',  text, re.I)
    beds  = int(bm.group(1))  if bm  else None
    baths = int(bam.group(1)) if bam else None

    # ── Furnishing ─────────────────────────────────────────────
    furnishing = parse_furnishing(text)

    # ── Listing date ───────────────────────────────────────────
    dm = re.search(
        r'(\d+\s*(?:day|week|month|year)[s]?\s*ago'
        r'|\d{1,2}\s+\w+\s+20\d{2}'
        r'|20\d{2}[/-]\d{2})',
        text, re.I,
    )
    listed = parse_relative_date(dm.group(0) if dm else "")

    # ── Title: first non-trivial line ──────────────────────────
    title = ""
    for line in text.splitlines():
        line = line.strip()
        if len(line) > 15:
            title = line[:120]
            break

    return {
        "first_seen":        run_date,
        "last_seen":         run_date,
        "listed_month_year": listed,
        "price_thb":         price,
        "price_per_sqm":     int(price / size) if size else None,
        "size_sqm":          size,
        "floor":             floor,
        "building":          None,
        "bedrooms":          beds,
        "bathrooms":         baths,
        "furnishing":        furnishing,
        "view":              None,
        "title":             title,
        "source":            source,
        "url":               href,
        "fingerprint":       make_fingerprint(size, floor, beds, baths),
    }


def _scrape(page: Page, name: str, urls: list[str], run_date: str,
            relevance_filter: bool = True) -> list[dict]:
    """
    Generic scraper used by every site.
    Tries each URL in order; stops when cards are found.
    Uses JS-based card finder (works on React/Next.js with hashed class names).
    """
    print(f"  → {name} …")
    cards: list[dict] = []

    for url in urls:
        try:
            _load_page(page, url)
            cards = _js_cards(page)
            if cards:
                print(f"    {len(cards)} cards at {url}")
                break
        except Exception as e:
            print(f"    {name} ({url}) error: {e}")

    if not cards:
        _save_debug_html(page, name.lower())
        print(f"    0 cards found — debug HTML saved for inspection")
        return []

    records = []
    for card in cards:
        # Relevance filter: skip if the card text doesn't mention our condo
        # (some sites return nearby results mixed in)
        if relevance_filter and not is_relevant(card.get("text", "")):
            continue
        rec = _parse_card(card, name, run_date)
        if rec:
            records.append(rec)

    print(f"    → {len(records)} listings extracted")
    return records


# ─────────────────────────────────────────
#  SITE-SPECIFIC SCRAPERS  (thin wrappers over _scrape)
#
#  NOTE: DDproperty.com is protected by Cloudflare Turnstile.
#  GitHub Actions IPs (Azure/AWS datacenters) are automatically blocked.
#  It is kept here as a best-effort attempt; it will return 0 results if
#  Cloudflare fires. PropertyHub and LivingInsider are smaller Thai sites
#  that do not use Cloudflare and should work reliably.
# ─────────────────────────────────────────

def scrape_ddproperty(page: Page, run_date: str) -> list[dict]:
    """DDproperty — likely blocked by Cloudflare from CI, kept as best-effort."""
    return _scrape(page, "DDproperty", [
        "https://www.ddproperty.com/en/property-for-sale"
        "?freetext=Supalai+City+Resort+Rama+8&property_type_code%5B%5D=CONDO",
    ], run_date, relevance_filter=False)


def scrape_propertyhub(page: Page, run_date: str) -> list[dict]:
    """PropertyHub.in.th — Thai portal, no Cloudflare."""
    return _scrape(page, "PropertyHub", [
        "https://www.propertyhub.in.th/th/sale/condo"
        "?keyword=%E0%B8%A8%E0%B8%B8%E0%B8%A0%E0%B8%B2%E0%B8%A5%E0%B8%B1%E0%B8%A2+%E0%B8%8B%E0%B8%B4%E0%B8%95%E0%B8%B5%E0%B9%89+%E0%B8%A3%E0%B8%B5%E0%B8%AA%E0%B8%AD%E0%B8%A3%E0%B9%8C%E0%B8%97+%E0%B8%9E%E0%B8%A3%E0%B8%B0%E0%B8%A3%E0%B8%B2%E0%B8%A1+8",
        "https://www.propertyhub.in.th/th/sale/condo?keyword=supalai+city+resort+rama+8",
        "https://www.propertyhub.in.th/property/sale?type=condo&keyword=supalai+city+resort+rama+8",
    ], run_date, relevance_filter=True)


def scrape_livinginsider(page: Page, run_date: str) -> list[dict]:
    """LivingInsider.com — Thai portal, no Cloudflare."""
    return _scrape(page, "LivingInsider", [
        "https://www.livinginsider.com/searchword/%E0%B8%A8%E0%B8%B8%E0%B8%A0%E0%B8%B2%E0%B8%A5%E0%B8%B1%E0%B8%A2-%E0%B8%8B%E0%B8%B4%E0%B8%95%E0%B8%B5%E0%B9%89-%E0%B8%A3%E0%B8%B5%E0%B8%AA%E0%B8%AD%E0%B8%A3%E0%B9%8C%E0%B8%97-%E0%B8%9E%E0%B8%A3%E0%B8%B0%E0%B8%A3%E0%B8%B2%E0%B8%A1-8/Condo/sell.html",
        "https://www.livinginsider.com/searchword/Supalai-City-Resort-Rama-8/Condo/sell.html",
        "https://www.livinginsider.com/search?keyword=supalai+city+resort+rama+8&property_type=condo&listing_type=sell",
    ], run_date, relevance_filter=True)



# ─────────────────────────────────────────
#  HISTORY PERSISTENCE
# ─────────────────────────────────────────

def load_history() -> list[dict]:
    if HISTORY_FILE.exists():
        with open(HISTORY_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


def save_history(records: list[dict]):
    DATA_DIR.mkdir(exist_ok=True)
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)


def save_raw(records: list[dict], run_ts: datetime):
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    fname = RAW_DIR / f"{run_ts.strftime('%Y-%m-%d_%H-%M')}_raw.json"
    with open(fname, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)


# ─────────────────────────────────────────
#  EXCEL WRITER
# ─────────────────────────────────────────

HEADER_LABELS = {
    "first_seen":        "First Seen",
    "last_seen":         "Last Seen",
    "listed_month_year": "Listed (YYYY-MM)",
    "price_thb":         "Price (THB)",
    "price_per_sqm":     "Price/sqm (THB)",
    "size_sqm":          "Size (sqm)",
    "floor":             "Floor",
    "building":          "Building",
    "bedrooms":          "Bedrooms",
    "bathrooms":         "Bathrooms",
    "furnishing":        "Furnishing",
    "view":              "View",
    "title":             "Listing Title",
    "source":            "Source(s)",
    "url":               "URL(s)",
    "fingerprint":       "Unit Fingerprint",
}


def style_sheet(ws, header_hex: str = "1F4E79"):
    header_fill = PatternFill(start_color=header_hex, end_color=header_hex, fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=10)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center

    # Auto-width (cap at 50)
    for col in ws.columns:
        max_len = max((len(str(cell.value or "")) for cell in col), default=8)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 3, 50)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    # Number format for price columns (C and E = indices 3 and 5)
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            if cell.column in (4, 5):   # price_thb, price_per_sqm (1-indexed)
                if isinstance(cell.value, (int, float)):
                    cell.number_format = "#,##0"


def write_excel(history: list[dict]):
    all_cols = COLUMNS
    human_cols = [HEADER_LABELS.get(c, c) for c in all_cols]

    df_all = pd.DataFrame(history, columns=all_cols)
    df_all.sort_values(["last_seen", "price_thb"], ascending=[False, True], inplace=True)
    df_all.columns = human_cols

    # My unit: 30–40 sqm, 1 BR
    raw_mine = [
        r for r in history
        if r.get("size_sqm") is not None
        and abs(float(r["size_sqm"]) - MY_SIZE_SQM) <= MY_SIZE_TOL
        and r.get("bedrooms") == MY_BEDROOMS
    ]
    df_mine = pd.DataFrame(raw_mine, columns=all_cols)
    df_mine.sort_values(["last_seen", "price_thb"], ascending=[False, True], inplace=True)
    df_mine.columns = human_cols

    # Run summary
    run_dates = sorted({r["last_seen"] for r in history}, reverse=True)
    summary_rows = []
    for rd in run_dates:
        subset  = [r for r in history if r["last_seen"] == rd]
        mine_s  = [r for r in subset
                   if r.get("size_sqm") is not None
                   and abs(float(r["size_sqm"]) - MY_SIZE_SQM) <= MY_SIZE_TOL
                   and r.get("bedrooms") == MY_BEDROOMS]
        prices  = [r["price_thb"] for r in subset  if r.get("price_thb")]
        m_prices= [r["price_thb"] for r in mine_s  if r.get("price_thb")]
        summary_rows.append({
            "Run Date":               rd,
            "Total Listings":         len(subset),
            "My Unit Matches":        len(mine_s),
            "Overall Min Price":      min(prices)  if prices  else None,
            "Overall Max Price":      max(prices)  if prices  else None,
            "Overall Avg Price":      int(sum(prices)/len(prices))  if prices  else None,
            "My Unit Min Price":      min(m_prices) if m_prices else None,
            "My Unit Avg Price":      int(sum(m_prices)/len(m_prices)) if m_prices else None,
            "My Unit Max Price":      max(m_prices) if m_prices else None,
            "Sources Scraped":        ", ".join(sorted({r["source"] for r in subset})),
        })
    df_summary = pd.DataFrame(summary_rows)

    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        df_all.to_excel(writer,     sheet_name="All Listings",               index=False)
        df_mine.to_excel(writer,    sheet_name="My Unit (~35sqm 1BR)",       index=False)
        df_summary.to_excel(writer, sheet_name="Run Summary",                index=False)

    wb = load_workbook(EXCEL_FILE)
    style_sheet(wb["All Listings"],            "1F4E79")   # dark blue
    style_sheet(wb["My Unit (~35sqm 1BR)"],    "1A5E20")   # dark green
    style_sheet(wb["Run Summary"],             "4A235A")   # purple
    wb.save(EXCEL_FILE)
    print(f"  Excel saved → {EXCEL_FILE}")


# ─────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────

def main():
    run_ts   = datetime.now()
    run_date = run_ts.strftime("%Y-%m-%d")

    print(f"\n{'='*60}")
    print(f"  Supalai City Resort Rama 8 — Price Tracker")
    print(f"  Run: {run_ts.strftime('%Y-%m-%d %H:%M')}")
    print(f"{'='*60}\n")

    DATA_DIR.mkdir(exist_ok=True)
    RAW_DIR.mkdir(parents=True, exist_ok=True)

    scrapers = [
        ("PropertyHub",   scrape_propertyhub),    # propertyhub.in.th
        ("LivingInsider", scrape_livinginsider),  # livinginsider.com
        ("DDproperty",    scrape_ddproperty),      # ddproperty.com (Cloudflare — best-effort)
    ]

    all_raw: list[dict] = []

    # Locate the Playwright-managed Chromium binary; fall back to known installed paths
    def _find_chromium() -> str | None:
        import glob as _glob
        candidates = sorted(
            _glob.glob("/opt/pw-browsers/chromium-*/chrome-linux/chrome"),
            reverse=True,   # newest revision first
        )
        return candidates[0] if candidates else None

    launch_kwargs: dict = dict(
        headless=True,
        args=[
            "--no-sandbox",
            "--disable-blink-features=AutomationControlled",
            "--disable-dev-shm-usage",
            "--disable-gpu",
        ],
    )
    chrome_path = _find_chromium()
    if chrome_path:
        launch_kwargs["executable_path"] = chrome_path
        print(f"Using Chromium: {chrome_path}")

    with sync_playwright() as pw:
        browser = pw.chromium.launch(**launch_kwargs)
        ctx = browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/123.0.0.0 Safari/537.36"
            ),
            viewport={"width": 1366, "height": 768},
            locale="th-TH",
            timezone_id="Asia/Bangkok",
        )
        page = ctx.new_page()
        # Hide webdriver property
        page.add_init_script(
            "Object.defineProperty(navigator,'webdriver',{get:()=>undefined})"
        )

        for name, fn in scrapers:
            print(f"Scraping {name} …")
            try:
                results = fn(page, run_date)
                all_raw.extend(results)
            except Exception as e:
                print(f"  {name} failed: {e}")
            time.sleep(2)

        browser.close()

    # Save raw snapshot
    save_raw(all_raw, run_ts)

    # Dedup within this run (same unit on multiple sites)
    unique_run = deduplicate_within_run(all_raw)
    print(f"\nThis run: {len(all_raw)} raw → {len(unique_run)} after within-run dedup")

    # Load history and merge
    history = load_history()
    added = 0
    updated = 0
    for rec in unique_run:
        if is_new_or_changed(rec, history):
            history.append(rec)
            added += 1
        else:
            # Update last_seen for matching record
            fp    = rec["fingerprint"]
            price = rec["price_thb"]
            for h in history:
                if (h["fingerprint"] == fp
                        and h.get("price_thb") == price):
                    h["last_seen"] = run_date
                    updated += 1
                    break

    print(f"History: +{added} new, {updated} updated last_seen")
    save_history(history)

    # Excel
    print("\nWriting Excel …")
    write_excel(history)

    # Summary
    mine = [
        r for r in unique_run
        if r.get("size_sqm") is not None
        and abs(float(r["size_sqm"]) - MY_SIZE_SQM) <= MY_SIZE_TOL
        and r.get("bedrooms") == MY_BEDROOMS
    ]
    prices = [r["price_thb"] for r in unique_run if r.get("price_thb")]
    mine_prices = [r["price_thb"] for r in mine if r.get("price_thb")]

    print(f"\n{'─'*60}")
    print(f"  Run complete")
    print(f"  Total unique listings this run : {len(unique_run)}")
    if prices:
        print(f"  Overall price range           : ฿{min(prices):,.0f} – ฿{max(prices):,.0f}")
    print(f"  Matches for your unit (≈35sqm 1BR): {len(mine)}")
    if mine_prices:
        print(f"  Your unit price range         : ฿{min(mine_prices):,.0f} – ฿{max(mine_prices):,.0f}")
        print(f"  Your unit avg price           : ฿{sum(mine_prices)//len(mine_prices):,.0f}")
    print(f"{'─'*60}\n")


if __name__ == "__main__":
    main()
