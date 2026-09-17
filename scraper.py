"""
Flipkart Mobile Assortment Scraper — v5 (v4 core + Apple supplemental pass)
============================================================================
v4 recap (parity pass w/ Fridge/TV scraper core):
  • find_product_link() -> (url, pid, anchor) 3-tuple
  • extract_pid() prefers ?pid= query param over /p/itmXXXX path slug
  • extract_price_block(): strips Exchange/Bank/EMI offer text before
    parsing ₹ values, reads price/MRP in DOM order
  • is_partial_page() + fetch_and_parse_page(): escalates to Playwright
    on ANY under-rendered page (not just zero-product pages)
  • Shared Playwright browser singleton, requests.Session() reuse
  • extract_special_tag() / extract_badges()
  • extract_mobile_attributes(): RAM/ROM/Expandable/Display/Camera/
    Battery/Processor/Warranty parsed from spec bullets
  • Resumable via PageSummary sheet + fixed OUTPUT_FILE

NEW IN v5 — Apple Supplemental Pass:
  Problem: the generic "mobiles" search buries older iPhones / SE /
  storage variants behind every other brand competing for the same 24
  slots per page, so they silently fall off before Flipkart's relevance
  ranking surfaces them. Fix: after the normal full scrape finishes, hit
  Flipkart's brand-facet URL directly:

      https://www.flipkart.com/search?q=mobiles&...&p[]=facets.brand[]=Apple

  and page through it with the SAME parse/fallback machinery, writing
  ONLY rows whose PID isn't already in the Products sheet. No dupes,
  fully resumable via its own ApplePageSummary sheet.

PATCH 1 — Adaptive requests circuit breaker:
  Problem: when Flipkart is fingerprint-blocking the requests.Session()
  (persistent 403s, not transient rate-limiting), the old code still
  burned 3 retries x sleep-backoff PER PAGE before falling back to
  Playwright — wasting 25-40s/page x hundreds of pages for a network
  path that was already proven dead for this run. Fix: a circuit
  breaker trips after N consecutive full-page failures and skips the
  requests path entirely (going straight to Playwright) until a
  periodic recheck interval passes, in case the block is temporary/IP
  based. A single 403 no longer triggers a sleep-and-retry loop either
  — it's treated as a hard fingerprint block, not a soft rate limit.

PATCH 2 — Early-exit on consecutive completely-empty pages:
  Problem: Flipkart's "X products ≈ Y pages" header estimate is
  frequently WAY higher than the number of pages the search paginator
  actually serves — past some depth (observed here: ~page 42 of a
  claimed 330) every subsequent page returns a genuinely empty result
  grid (0 raw cards, confirmed even after the anchor-climb fallback and
  a full Playwright scroll-render). Retrying and Playwright-escalating
  each of those ~290 dead pages one at a time wastes ~25s/page for no
  data. Fix: track CONSECUTIVE complete failures (0 products, not just
  "few products") in both the main loop and the Apple supplemental
  loop. Once that streak hits MAX_CONSECUTIVE_EMPTY_PAGES, assume we've
  hit the real end of Flipkart's served results, log it clearly, and
  break out of that pagination loop early — the main loop then falls
  through into the Apple supplemental pass exactly as if it had
  finished all "total_pages" normally, so nothing downstream needs to
  change or gets skipped.

SETUP:
    pip install requests beautifulsoup4 openpyxl playwright lxml
    playwright install chromium
"""

import re
import time
import random
import logging
from datetime import datetime
from typing import Optional, Dict, List, Tuple, Any

import requests
from bs4 import BeautifulSoup, Tag
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

# ============================================================
#                  ⚙️ CONFIGURATION
# ============================================================
BASE_URL = (
    "https://www.flipkart.com/search?q=mobiles&otracker=search"
    "&otracker1=search&marketplace=FLIPKART&as-show=on&as=off"
)

# ---- Apple supplemental pass ----
# Flipkart's brand-facet filter, discovered via inspector. Same page
# pattern as the main search: page 1 has no &page= param, page 2+ does.
APPLE_BASE_URL = (
    "https://www.flipkart.com/search?q=mobiles&otracker=search"
    "&otracker1=search&marketplace=FLIPKART&as-show=on&as=off"
    "&p%5B%5D=facets.brand%255B%255D%3DApple"
)
RUN_APPLE_SUPPLEMENTAL_PASS = True   # set False to skip this pass entirely

# Fixed filename (not timestamped) so resume-by-PageSummary works
# across multiple runs.
OUTPUT_FILE = "7th_flipkart_mobile_assortment_v5.xlsx"

MAX_PAGES = 400                    # hard cap, mobiles is a huge category
PRODUCTS_PER_PAGE_DEFAULT = 24     # used for partial-page ratio check
MIN_ACCEPTABLE_PRODUCTS_RATIO = 0.70

REQUEST_DELAY_MIN = 2.0
REQUEST_DELAY_MAX = 4.5
SAVE_EVERY_N_PAGES = 5
MAX_RETRIES_PER_URL = 3
TIMEOUT_SECONDS = 20

DIAGNOSTIC_MODE = False            # True = page 1 only, print, no write
ENABLE_PLAYWRIGHT_FALLBACK = True

# ---- Circuit breaker for the requests path ----
# Once requests.Session() gets fingerprint-blocked (persistent 403s),
# retrying it page after page is pure wasted time. Trip after this many
# CONSECUTIVE whole-page failures, then skip straight to Playwright for
# every subsequent page until the recheck interval comes around (in
# case the block is IP-based / temporary and clears up later in a long
# run).
CIRCUIT_BREAKER_TRIP_THRESHOLD = 2
CIRCUIT_BREAKER_RECHECK_EVERY_N_PAGES = 50

# ---- Early-exit on consecutive completely-empty pages ----
# Flipkart's declared total-product-count / total-page-count is often
# wildly optimistic — the paginator itself stops serving real results
# well before that point and just returns empty grids. Once we see this
# many CONSECUTIVE pages with 0 products (even after Playwright +
# scroll + anchor-climb fallback all agree on 0), stop paginating that
# section rather than grinding through hundreds of guaranteed-dead
# pages. Applies independently to the main "mobiles" loop and the Apple
# brand-facet loop.
MAX_CONSECUTIVE_EMPTY_PAGES = 5

LOG_LEVEL = logging.INFO

# ============================================================
#                    LOGGING SETUP
# ============================================================
logging.basicConfig(
    level=LOG_LEVEL,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-IN,en;q=0.9,en-US;q=0.8",
    "Accept-Encoding": "gzip, deflate, br",
    "Connection": "keep-alive",
}

SESSION = requests.Session()
SESSION.headers.update(HEADERS)

# ============================================================
#                    REGEX PATTERNS
# ============================================================
PATTERNS = {
    "discount": re.compile(r"(\d+)\s*%\s*off", re.IGNORECASE),
    "rating_count": re.compile(r"([\d,]+)\s*Ratings?\b", re.IGNORECASE),
    "review_count": re.compile(r"([\d,]+)\s*Reviews?\b", re.IGNORECASE),
    "rating_value": re.compile(r"\b([1-5](?:\.\d)?)\b"),
    "total_products": re.compile(r"of\s+([\d,]+)\s+(?:results|products|items)", re.IGNORECASE),
    "page_of_total": re.compile(r"Page\s+\d+\s+of\s+([\d,]+)", re.IGNORECASE),
    "junk_prefix": re.compile(r"^(add\s+to\s+compare|compare)\s*", re.IGNORECASE),
    "product_href": re.compile(r"/p/(itm[A-Za-z0-9]+|[A-Za-z0-9\-]+)"),

    "pid_query_param": re.compile(r"[?&]pid=([A-Z0-9]{10,20})", re.IGNORECASE),
    "itm_path_id": re.compile(r"/p/(itm[A-Za-z0-9]+)"),

    # ---- Mobile-specific structured-attribute parsers ----
    "mobile_ram": re.compile(r"(\d+)\s*GB\s*RAM", re.IGNORECASE),
    "mobile_rom": re.compile(r"(\d+)\s*(GB|TB)\s*ROM", re.IGNORECASE),
    "mobile_expandable": re.compile(r"Expandable\s*Upto\s*(\d+\s*(?:GB|TB))", re.IGNORECASE),
    "mobile_display_inch": re.compile(r"\(([\d.]+)\s*inch\)", re.IGNORECASE),
    "mobile_battery": re.compile(r"(\d+)\s*mAh", re.IGNORECASE),
}

KNOWN_SPECIAL_TAGS = [
    "Bestseller", "Flipkart's Choice", "Trending", "New Launch",
    "Great Discount", "Limited Deal", "Deal of the Day",
]


# ============================================================
#                 SHARED / GENERIC HELPERS
# ============================================================

def clean_text(text: Optional[str]) -> str:
    if not text:
        return ""
    return " ".join(text.split()).strip()


def extract_number(raw: Optional[str]) -> Optional[int]:
    if not raw:
        return None
    cleaned = str(raw).replace(",", "").replace("₹", "").strip()
    try:
        return int(float(cleaned))
    except ValueError:
        return None


def extract_float(raw: Optional[str]) -> Optional[float]:
    if not raw:
        return None
    cleaned = str(raw).replace(",", "").strip()
    try:
        return round(float(cleaned), 2)
    except ValueError:
        return None


def extract_pid(href: str) -> str:
    """Flipkart URLs carry TWO ids: the /p/itmXXXX SEO slug (path) and
    the real ?pid= query param. The query param is canonical — always
    check it first regardless of which appears earlier in the string."""
    m = PATTERNS["pid_query_param"].search(href)
    if m:
        return m.group(1).upper()
    m = PATTERNS["itm_path_id"].search(href)
    if m:
        return m.group(1)
    return href


# ============================================================
#                 CARD DISCOVERY (LEAF NODES)
# ============================================================

def get_leaf_cards(soup: BeautifulSoup) -> List[Tag]:
    """
    Strategy A: div[data-tkid] — this is what Flipkart's SEARCH RESULT
    template (q=mobiles) uses. Must de-dupe nested wrappers (outer
    wrapper parses into junk titles like "Add to CompareMOTOROLA...").
    Strategy B: div[data-id] — category/browse-page template, kept as
    a fallback in case the query ever redirects to a /pr/ style page.
    Strategy C: anchor-climb — last resort, keyword-scored for mobiles.
    """
    all_tkid = soup.select("div[data-tkid]")
    leaves = [c for c in all_tkid if c.find("div", attrs={"data-tkid": True}) is None]
    if leaves:
        logger.debug(f"{len(leaves)} leaf cards via data-tkid (dropped {len(all_tkid) - len(leaves)} nested)")
        return leaves

    all_id = soup.select("div[data-id]")
    leaves = [c for c in all_id if c.find("div", attrs={"data-id": True}) is None]
    if leaves:
        logger.debug(f"{len(leaves)} leaf cards via data-id")
        return leaves

    logger.warning("No data-tkid/data-id leaves found! Using anchor-climb fallback.")
    mobile_keywords = ["ram", "rom", "battery", "camera", "processor",
                       "display", "mah", "inch", "smartphone", "mobile"]
    seen_pids = set()
    candidates = []
    for a in soup.find_all("a", href=PATTERNS["product_href"]):
        href = a.get("href", "")
        pid = extract_pid(href)
        if not pid or pid == href or pid in seen_pids:
            continue
        seen_pids.add(pid)
        parent = a.parent
        for _ in range(6):
            if not parent:
                break
            txt = parent.get_text(" ").lower()
            if "₹" in txt and any(w in txt for w in mobile_keywords):
                candidates.append(parent)
                break
            parent = parent.parent

    logger.debug(f"{len(candidates)} leaf cards via anchor-climb")
    return candidates


def find_product_link(card: Tag) -> Optional[Tuple[str, str, Tag]]:
    """Returns (full_url, pid, anchor_tag) or None."""
    link = card.find("a", href=PATTERNS["product_href"])
    if link:
        href = link.get("href", "")
        url = href if href.startswith("http") else f"https://www.flipkart.com{href}"
        return url, extract_pid(href), link

    for a in card.find_all("a", attrs={"target": "_blank"}):
        href = a.get("href", "")
        if "/p/" in href:
            url = href if href.startswith("http") else f"https://www.flipkart.com{href}"
            return url, extract_pid(href), a

    for a in card.find_all("a", href=True):
        href = a.get("href", "")
        if "/p/" in href:
            url = href if href.startswith("http") else f"https://www.flipkart.com{href}"
            return url, extract_pid(href), a

    return None


def extract_title(card: Tag, link: Tag) -> Optional[str]:
    if link is not None and hasattr(link, "get"):
        t = link.get("title")
        if t and len(t.strip()) > 5:
            return clean_text(t)

    div = card.select_one("div.RG5Slk")
    if div:
        t = div.get_text(strip=True)
        if t and len(t) > 5:
            return clean_text(t)

    mobile_keywords = ["ram", "rom", "battery", "camera", "processor", "display", "mah"]
    for col in card.find_all("div", class_=re.compile(r"col-\d+")):
        txt = col.get_text(" ", strip=True)
        if len(txt) > 15 and any(w in txt.lower() for w in mobile_keywords):
            lines = [l.strip() for l in txt.split("\n") if l.strip()]
            if lines and lines[0].lower() not in ["add to compare", "compare"]:
                return clean_text(lines[0])

    texts = [t for t in card.stripped_strings if len(t) > 10 and "http" not in t]
    if texts:
        return clean_text(PATTERNS["junk_prefix"].sub("", texts[0]))

    return None


def validate_and_clean(title: Optional[str]) -> bool:
    if not title:
        return False
    t_lower = title.strip().lower()
    if t_lower in ["add to compare", "", "compare"]:
        return False
    if t_lower.startswith(("add to compare", "compare")):
        return False
    if len(title.strip()) < 3:
        return False
    if re.match(r'^[\d.,\s]+$', title):
        return False
    if re.match(r'^\d+\.\s', title):   # "1. " numbered junk
        return False
    return True


# ============================================================
#                    RATING / PRICE / SPECS
# ============================================================

def extract_rating_info(card: Tag) -> Dict[str, Any]:
    result = {"rating": None, "rating_count": None, "review_count": None}
    combined_text = card.get_text(" ", strip=True)

    rating_span = card.select_one("span[id^='productRating_']")
    if rating_span:
        val_text = "".join(
            child.strip() for child in rating_span.descendants
            if isinstance(child, str) and child.strip()
        )
        m = PATTERNS["rating_value"].search(val_text)
        if m:
            result["rating"] = m.group(1)

        parent_span = rating_span.find_parent("span") or rating_span.find_parent("div")
        if parent_span:
            next_sib = parent_span.find_next_sibling(["span", "div"])
            if next_sib:
                sib_txt = next_sib.get_text(" ", strip=True)
                rc = PATTERNS["rating_count"].search(sib_txt)
                rv = PATTERNS["review_count"].search(sib_txt)
                if rc:
                    result["rating_count"] = extract_number(rc.group(1))
                if rv:
                    result["review_count"] = extract_number(rv.group(1))

    if not result["rating"]:
        m = PATTERNS["rating_value"].search(combined_text)
        if m:
            result["rating"] = m.group(1)

    if not result["rating_count"]:
        pvbnmb = card.select_one("span.PvbNMB")
        rr_text = pvbnmb.get_text(" ", strip=True) if pvbnmb else combined_text
        rc = PATTERNS["rating_count"].search(rr_text)
        if rc:
            result["rating_count"] = extract_number(rc.group(1))
        elif pvbnmb:
            m2 = re.search(r"\(?([\d,]+)\)?", pvbnmb.get_text(strip=True))
            if m2:
                result["rating_count"] = extract_number(m2.group(1))

    if not result["review_count"]:
        rv = PATTERNS["review_count"].search(combined_text)
        if rv:
            result["review_count"] = extract_number(rv.group(1))

    return result


def extract_price_block(card: Tag) -> Dict[str, Any]:
    """Strips Exchange/Bank/EMI offer text BEFORE parsing ₹ values, and
    reads prices in DOM order (not numeric-sorted) — an "Upto ₹3,000 Off
    on Exchange" line sits ahead of the real price/MRP in raw text and
    was previously mis-parsed as the MRP."""
    result = {"price": None, "mrp": None, "discount_pct": None}

    price_container = (
        card.select_one("div[class*='QiMO5r']") or
        card.select_one("div[class*='oFEPlD']") or
        card.select_one("div[class*='col-5-12']") or
        card
    )

    raw_text = price_container.get_text(" ", strip=True) if price_container else card.get_text(" ", strip=True)
    if "₹" not in raw_text:
        raw_text = card.get_text(" ", strip=True)

    for marker in ["Off on Exchange", "Exchange", "Bank Offer", "No Cost EMI", "EMI starts"]:
        idx = raw_text.find(marker)
        if idx != -1:
            upto_idx = raw_text.rfind("Upto", 0, idx)
            cut_point = upto_idx if upto_idx != -1 else idx
            raw_text = raw_text[:cut_point]

    price_matches = re.findall(r"₹\s?([\d,]+)", raw_text)
    nums = [n for n in (extract_number(p) for p in price_matches) if n and n > 100]

    if nums:
        result["price"] = nums[0]
        if len(nums) > 1:
            result["mrp"] = nums[1] if nums[1] >= nums[0] else max(nums[1:])

    dm = PATTERNS["discount"].search(raw_text)
    if dm:
        result["discount_pct"] = int(dm.group(1))

    if result["discount_pct"] and result["price"] and not result["mrp"]:
        result["mrp"] = int(result["price"] / (1 - result["discount_pct"] / 100))

    return result


def extract_specs(card: Tag) -> List[str]:
    specs = []
    ul = card.select_one("ul.HwRTzP")
    if ul:
        for li in ul.find_all("li"):
            t = clean_text(li.get_text())
            if t and len(t) > 3:
                specs.append(t)
        if specs:
            return specs

    for ul in card.find_all("ul"):
        items = [clean_text(li.get_text()) for li in ul.find_all("li")]
        items = [t for t in items if t and len(t) > 3]
        if len(items) >= 2:
            specs.extend(items)
            return specs

    potential = card.select("div.CMXw7N") or [card]
    for div in potential:
        children = div.find_all(["li", "div", "span"])
        short_texts = [c.get_text(strip=True) for c in children
                       if 3 < len(c.get_text(strip=True)) < 120 and c.name not in ("a", "img")]
        if len(short_texts) >= 2:
            specs.extend(short_texts[:10])
            break

    return specs


def extract_badges(card: Tag) -> List[str]:
    badges = []
    badge_keywords = ["hot deal", "assured", "bank offer", "exchange", "emi", "free delivery"]
    lower_text = card.get_text(" ").lower()
    for kw in badge_keywords:
        if kw in lower_text:
            badges.append(kw.title())
    return badges


def extract_special_tag(card: Tag) -> Optional[str]:
    """Ribbon tags like 'Bestseller' — same structure across categories:
    a small div with an inline background-color style, sitting outside
    the normal title/price block."""
    tag_div = card.select_one("div[class*='o2uEoz']")
    if tag_div:
        text = clean_text(tag_div.get_text())
        if text and 2 < len(text) < 40:
            return text

    for div in card.find_all("div", style=re.compile(r"background", re.IGNORECASE)):
        text = clean_text(div.get_text())
        if text and 2 < len(text) < 30 and not re.search(r"[₹\d%]", text):
            return text

    combined = card.get_text(" ", strip=True)
    for tag in KNOWN_SPECIAL_TAGS:
        if tag.lower() in combined.lower():
            return tag

    return None


# ============================================================
#      MOBILE-SPECIFIC STRUCTURED ATTRIBUTES (spec parsing)
# ============================================================

def extract_mobile_attributes(specs: List[str], title: Optional[str]) -> Dict[str, Any]:
    """
    Equivalent of extract_fridge_attributes(), but for mobiles the
    reliable source is the bullet SPEC list, not the title (fridge
    titles are highly templated; mobile titles are mostly just brand +
    model + color + storage, e.g. "REDMI Note 12 Pro (Stardust, 128 GB)").

    Spec bullets are near-universally formatted like:
        "8 GB RAM | 128 GB ROM | Expandable Upto 1 TB"
        "16.94 cm (6.67 inch) Full HD+ Display"
        "50MP + 2MP | 8MP Front Camera"
        "5000 mAh Battery"
        "Snapdragon 4 Gen 2 Processor"
        "1 Year Warranty on Handset..."
    """
    result = {
        "ram_gb": None, "rom_gb": None, "expandable_storage": None,
        "display_size_inch": None, "rear_camera": None, "front_camera": None,
        "battery_mah": None, "processor": None, "warranty": None,
    }

    search_blob = " | ".join(specs) if specs else (title or "")

    ram_m = PATTERNS["mobile_ram"].search(search_blob)
    if ram_m:
        result["ram_gb"] = extract_number(ram_m.group(1))

    rom_m = PATTERNS["mobile_rom"].search(search_blob)
    if rom_m:
        val = extract_number(rom_m.group(1))
        if val and rom_m.group(2).upper() == "TB":
            val *= 1024
        result["rom_gb"] = val

    exp_m = PATTERNS["mobile_expandable"].search(search_blob)
    if exp_m:
        result["expandable_storage"] = exp_m.group(1).strip().upper()

    disp_m = PATTERNS["mobile_display_inch"].search(search_blob)
    if disp_m:
        result["display_size_inch"] = extract_float(disp_m.group(1))

    batt_m = PATTERNS["mobile_battery"].search(search_blob)
    if batt_m:
        result["battery_mah"] = extract_number(batt_m.group(1))

    # Per-line scan for free-text fields (camera/processor/warranty are
    # too variable for a single regex to be reliable across brands)
    for line in specs:
        low = line.lower()
        if "camera" in low:
            if "front" in low and not result["front_camera"]:
                result["front_camera"] = line
            elif "front" not in low and not result["rear_camera"]:
                result["rear_camera"] = line
        elif "processor" in low and not result["processor"]:
            result["processor"] = line
        elif "warranty" in low and not result["warranty"]:
            result["warranty"] = line

    return result


# ============================================================
#                      MAIN CARD PARSER
# ============================================================

def parse_single_card(card: Tag) -> Optional[Dict[str, Any]]:
    link_info = find_product_link(card)
    if not link_info:
        logger.debug("Skipping card: no product link found")
        return None
    url, pid, anchor_tag = link_info

    raw_title = extract_title(card, anchor_tag)
    if not validate_and_clean(raw_title):
        logger.debug(f"Skipping card {pid}: invalid title '{raw_title}'")
        return None
    title = PATTERNS["junk_prefix"].sub("", raw_title).strip()

    rating_data = extract_rating_info(card)
    price_data = extract_price_block(card)
    specs = extract_specs(card)
    badges = extract_badges(card)
    special_tag = extract_special_tag(card)
    mobile_attrs = extract_mobile_attributes(specs, title)

    has_content = any([
        price_data["price"], rating_data["rating"],
        rating_data["rating_count"], specs,
    ])
    if not has_content:
        logger.warning(f"Card {pid} ('{title[:30]}...') had almost no extracted data. Skipping.")
        return None

    return {
        "pid": pid, "title": title, "url": url,
        "selling_price": price_data["price"],
        "mrp": price_data["mrp"],
        "discount_pct": price_data["discount_pct"],
        "avg_rating": rating_data["rating"],
        "total_ratings": rating_data["rating_count"],
        "total_reviews": rating_data["review_count"],
        "ram_gb": mobile_attrs["ram_gb"],
        "rom_gb": mobile_attrs["rom_gb"],
        "expandable_storage": mobile_attrs["expandable_storage"],
        "display_size_inch": mobile_attrs["display_size_inch"],
        "rear_camera": mobile_attrs["rear_camera"],
        "front_camera": mobile_attrs["front_camera"],
        "battery_mah": mobile_attrs["battery_mah"],
        "processor": mobile_attrs["processor"],
        "warranty": mobile_attrs["warranty"],
        "key_features": " | ".join(specs) if specs else "",
        "special_tag": special_tag or "",
        "badges_offers": ", ".join(badges) if badges else "",
        "scraped_at": datetime.now().isoformat(),
    }


def parse_page(html: str) -> List[Dict[str, Any]]:
    soup = BeautifulSoup(html, "lxml")
    cards = get_leaf_cards(soup)

    if 0 < len(cards) < PRODUCTS_PER_PAGE_DEFAULT * MIN_ACCEPTABLE_PRODUCTS_RATIO:
        logger.debug(f"⚠️ Only {len(cards)} raw cards found (expected ~{PRODUCTS_PER_PAGE_DEFAULT}) — possibly under-rendered.")

    results, pids_seen = [], set()
    for idx, card in enumerate(cards, 1):
        try:
            parsed = parse_single_card(card)
            if parsed and parsed["pid"] not in pids_seen:
                results.append(parsed)
                pids_seen.add(parsed["pid"])
        except Exception as e:
            logger.error(f"Error parsing card #{idx}: {e}", exc_info=True)
            continue

    logger.info(f"Successfully parsed {len(results)} products from page (raw cards: {len(cards)}).")
    return results


def is_partial_page(products: List[Dict], expected: int = PRODUCTS_PER_PAGE_DEFAULT) -> bool:
    """Escalates to Playwright on ANY under-rendered page, not just
    zero-product pages — a 9-of-24 lazy-load render must not be
    accepted as final."""
    if not products:
        return True
    return len(products) < (expected * MIN_ACCEPTABLE_PRODUCTS_RATIO)


# ============================================================
#            ADAPTIVE CIRCUIT BREAKER FOR requests PATH
# ============================================================
# Once Flipkart starts fingerprint-blocking requests.Session() (persistent
# 403s across every retry), continuing to retry it page after page is pure
# wasted time — a 403 here means "your client is identified as a bot", not
# "try again in a few seconds". This breaker trips after N consecutive
# whole-page failures and then skips the requests path entirely (going
# straight to Playwright) for the next `recheck_every` pages, at which
# point it gives requests one more chance in case the block was IP-based
# / temporary and has since cleared.
REQUESTS_CIRCUIT = {
    "consecutive_failures": 0,
    "tripped": False,
    "trip_threshold": CIRCUIT_BREAKER_TRIP_THRESHOLD,
    "recheck_every": CIRCUIT_BREAKER_RECHECK_EVERY_N_PAGES,
    "pages_since_trip": 0,
}


# ============================================================
#                     NETWORK LAYER
# ============================================================

def fetch_page_requests(url: str) -> Optional[str]:
    """
    Fetches via requests.Session(). Behavior changes from earlier
    versions:
      • A 403 is treated as a hard fingerprint block, not a transient
        rate limit — no sleep-and-retry loop, we bail immediately so
        fetch_and_parse_page() can escalate to Playwright right away.
      • A circuit breaker skips this function's network call entirely
        once it's proven dead for `trip_threshold` consecutive pages,
        until `recheck_every` pages have passed.
    """
    if REQUESTS_CIRCUIT["tripped"]:
        REQUESTS_CIRCUIT["pages_since_trip"] += 1
        if REQUESTS_CIRCUIT["pages_since_trip"] < REQUESTS_CIRCUIT["recheck_every"]:
            logger.debug("Circuit breaker OPEN — skipping requests path, going straight to Playwright.")
            return None
        logger.info("Circuit breaker: periodic recheck of requests path...")
        REQUESTS_CIRCUIT["pages_since_trip"] = 0

    for attempt in range(MAX_RETRIES_PER_URL):
        try:
            resp = SESSION.get(url, timeout=TIMEOUT_SECONDS)
            if resp.status_code == 200:
                REQUESTS_CIRCUIT["consecutive_failures"] = 0
                if REQUESTS_CIRCUIT["tripped"]:
                    logger.info("✅ Circuit breaker CLOSED — requests path working again.")
                REQUESTS_CIRCUIT["tripped"] = False
                return resp.text
            elif resp.status_code == 403:
                logger.warning(f"403 Forbidden (attempt {attempt+1}) — fingerprint block, not retrying this URL further via requests.")
                break
            elif resp.status_code == 429:
                wait = (attempt + 1) * 10
                logger.warning(f"429 Rate limited. Waiting {wait}s...")
                time.sleep(wait)
            else:
                logger.warning(f"HTTP {resp.status_code}. Retrying...")
                time.sleep(3)
        except Exception as e:
            logger.error(f"Request failed attempt {attempt+1}: {e}")
            time.sleep(2)

    # Whole page failed via requests.
    REQUESTS_CIRCUIT["consecutive_failures"] += 1
    if REQUESTS_CIRCUIT["consecutive_failures"] >= REQUESTS_CIRCUIT["trip_threshold"]:
        if not REQUESTS_CIRCUIT["tripped"]:
            logger.warning(
                f"⚡ Circuit breaker TRIPPED — requests failed "
                f"{REQUESTS_CIRCUIT['consecutive_failures']} consecutive pages. "
                f"Skipping requests path for the next {REQUESTS_CIRCUIT['recheck_every']} pages "
                f"(going straight to Playwright)."
            )
        REQUESTS_CIRCUIT["tripped"] = True
        REQUESTS_CIRCUIT["pages_since_trip"] = 0
    return None


_playwright_ctx = None
_browser_instance = None


def get_or_create_browser():
    global _playwright_ctx, _browser_instance
    if _browser_instance is not None:
        return _browser_instance
    from playwright.sync_api import sync_playwright
    _playwright_ctx = sync_playwright().start()
    _browser_instance = _playwright_ctx.chromium.launch(
        headless=True,
        args=["--no-sandbox", "--disable-dev-shm-usage", "--disable-gpu",
              "--disable-extensions", "--disable-background-networking",
              "--disable-default-apps", "--disable-sync",
              "--js-flags=--max-old-space-size=256"],
    )
    logger.info("Playwright browser launched (shared for entire run).")
    return _browser_instance


def close_shared_browser():
    global _playwright_ctx, _browser_instance
    if _browser_instance is not None:
        try:
            _browser_instance.close()
            logger.info("Playwright browser closed.")
        except Exception:
            pass
        _browser_instance = None
    if _playwright_ctx is not None:
        try:
            _playwright_ctx.stop()
        except Exception:
            pass
        _playwright_ctx = None


def fetch_page_playwright(url: str, target_card_count: int = PRODUCTS_PER_PAGE_DEFAULT) -> Optional[str]:
    """Uses the SHARED browser. Polls the LIVE max(data-tkid, data-id)
    count while scrolling, since scrollHeight plateaus early on
    Flipkart's virtualized grid."""
    if not ENABLE_PLAYWRIGHT_FALLBACK:
        return None

    context = None
    try:
        browser = get_or_create_browser()
        context = browser.new_context(user_agent=HEADERS["User-Agent"],
                                       viewport={"width": 1280, "height": 800})
        page = context.new_page()
        page.goto(url, wait_until="domcontentloaded", timeout=30000)
        page.wait_for_timeout(1500)

        previous_count = 0
        for attempt in range(12):
            page.evaluate("window.scrollBy(0, document.body.scrollHeight)")
            page.wait_for_timeout(700)
            current_count = page.evaluate(
                "Math.max(document.querySelectorAll('div[data-tkid]').length,"
                " document.querySelectorAll('div[data-id]').length)"
            )
            logger.debug(f"  [Playwright] scroll {attempt+1}: {current_count} cards")
            if current_count >= target_card_count:
                break
            if current_count == previous_count:
                page.wait_for_timeout(1000)
            previous_count = current_count

        page.evaluate("window.scrollTo(0, 0)")
        page.wait_for_timeout(500)
        html = page.content()
        return html
    except Exception as e:
        logger.error(f"Playwright error: {e}")
        return None
    finally:
        if context is not None:
            try:
                context.close()
            except Exception:
                pass


def fetch_and_parse_page(url: str, page_label: str = "") -> List[Dict[str, Any]]:
    html = fetch_page_requests(url)
    products = parse_page(html) if html else []

    if is_partial_page(products):
        if html is None:
            logger.info(f"{page_label}: requests path unavailable/blocked — going straight to Playwright...")
        else:
            logger.warning(f"{page_label}: got {len(products)} products via requests "
                            f"(expected ~{PRODUCTS_PER_PAGE_DEFAULT}) — forcing Playwright + scroll...")
        pw_html = fetch_page_playwright(url)
        if pw_html:
            pw_products = parse_page(pw_html)
            if len(pw_products) > len(products):
                logger.info(f"{page_label}: Playwright recovered {len(pw_products)} "
                            f"(vs {len(products)}). Using Playwright result.")
                products = pw_products
            else:
                logger.warning(f"{page_label}: Playwright did not improve "
                                f"({len(pw_products)} vs {len(products)}). Keeping original.")

    return products


def determine_total_pages(first_page_html: str) -> int:
    m = PATTERNS["total_products"].search(first_page_html)
    if m:
        total_prods = extract_number(m.group(1))
        if total_prods:
            calc_pages = -(-total_prods // PRODUCTS_PER_PAGE_DEFAULT)  # ceil
            final = min(calc_pages, MAX_PAGES)
            logger.info(f"Detected {total_prods} products ≈ {calc_pages} pages. Capping at {final}.")
            return final

    m = PATTERNS["page_of_total"].search(first_page_html)
    if m:
        total_pages = extract_number(m.group(1))
        if total_pages:
            final = min(total_pages, MAX_PAGES)
            logger.info(f"Detected 'Page X of {total_pages}'. Capping at {final}.")
            return final

    logger.warning("Could not detect total pages/products. Defaulting to MAX_PAGES.")
    return MAX_PAGES


def build_page_url(page_num: int) -> str:
    return f"{BASE_URL}&page={page_num}"


def build_apple_page_url(page_num: int) -> str:
    """Apple brand-facet search. Page 1 has no &page= param, same as
    the generic search."""
    if page_num <= 1:
        return APPLE_BASE_URL
    return f"{APPLE_BASE_URL}&page={page_num}"


# ============================================================
#                      EXCEL I/O
# ============================================================

COLUMN_HEADERS = [
    "Page No.", "PID", "Product Name", "Product URL",
    "Selling Price (₹)", "MRP (₹)", "Discount (%)",
    "Avg Rating", "Total Ratings", "Total Reviews",
    "RAM (GB)", "ROM (GB)", "Expandable Storage", "Display (inch)",
    "Rear Camera", "Front Camera", "Battery (mAh)", "Processor", "Warranty",
    "Key Features (Specs)", "Special Tag", "Badges/Offers", "Scraped At",
]


def init_workbook():
    """Load existing file (for resume) or create a fresh styled one.
    Returns (wb, ws, summary_ws, apple_summary_ws)."""
    try:
        wb = openpyxl.load_workbook(OUTPUT_FILE)
        ws = wb["Products"]
        summary_ws = wb["PageSummary"]
        try:
            apple_summary_ws = wb["ApplePageSummary"]
        except KeyError:
            apple_summary_ws = wb.create_sheet("ApplePageSummary")
            apple_summary_ws.append(["Apple Page", "Product Count", "New Rows Added", "Timestamp"])
        logger.info(f"Resuming from existing workbook: {OUTPUT_FILE}")
    except (FileNotFoundError, KeyError):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Products"

        header_fill = PatternFill(start_color="0066CC", end_color="0066CC", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")
        ws.append(COLUMN_HEADERS)
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        widths = [8, 18, 45, 50, 12, 12, 11, 10, 12, 12, 9, 9, 14, 12, 40, 30, 12, 30, 30, 60, 16, 20, 20]
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

        summary_ws = wb.create_sheet("PageSummary")
        summary_ws.append(["Page", "Product Count", "With Rating", "Timestamp"])

        apple_summary_ws = wb.create_sheet("ApplePageSummary")
        apple_summary_ws.append(["Apple Page", "Product Count", "New Rows Added", "Timestamp"])

    return wb, ws, summary_ws, apple_summary_ws


def get_scraped_pages(summary_ws) -> set:
    pages = set()
    for row in summary_ws.iter_rows(min_row=2, values_only=True):
        if row and row[0] is not None:
            try:
                pages.add(int(row[0]))
            except (ValueError, TypeError):
                continue
    return pages


def get_existing_pids(ws) -> set:
    """All PIDs already written to the Products sheet — used so the
    Apple supplemental pass only ADDS genuinely missing rows instead of
    re-inserting iPhones the generic 'mobiles' search already caught."""
    pids = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row and row[1]:
            pids.add(str(row[1]))
    return pids


def append_rows(ws, products: List[Dict], page_num):
    for p in products:
        ws.append([
            page_num, p["pid"], p["title"], p["url"],
            p["selling_price"], p["mrp"], p["discount_pct"],
            p["avg_rating"], p["total_ratings"], p["total_reviews"],
            p["ram_gb"], p["rom_gb"], p["expandable_storage"], p["display_size_inch"],
            p["rear_camera"], p["front_camera"], p["battery_mah"], p["processor"], p["warranty"],
            p["key_features"], p["special_tag"], p["badges_offers"], p["scraped_at"],
        ])


# ============================================================
#                       DIAGNOSTIC
# ============================================================

def run_diagnostic_mode():
    logger.info("=" * 60)
    logger.info("DIAGNOSTIC MODE: page 1 only, nothing written")
    logger.info("=" * 60)

    url = build_page_url(1)
    products = fetch_and_parse_page(url, page_label="Page 1 (diagnostic)")

    if not products:
        logger.error("CRITICAL: could not extract any products from page 1.")
        close_shared_browser()
        return

    with_rating = sum(1 for p in products if p["avg_rating"])
    with_ram = sum(1 for p in products if p["ram_gb"])
    logger.info(f"Parsed {len(products)} products | with rating: {with_rating} | with RAM parsed: {with_ram}\n")

    for i, p in enumerate(products[:15], 1):
        tag_str = f"[{p['special_tag']}] " if p["special_tag"] else ""
        print(f"{i:3d} | PID:{p['pid']:>18} | {str(p['avg_rating']):>3}★ | "
              f"₹{str(p['selling_price']):>8} | {tag_str}{p['title'][:45]}")
        print(f"     RAM:{p['ram_gb']} ROM:{p['rom_gb']} Display:{p['display_size_inch']}in "
              f"Battery:{p['battery_mah']}mAh")
        print(f"     Rear Cam: {p['rear_camera']}")
        print(f"     Processor: {p['processor']}")
        print("-" * 100)

    close_shared_browser()


# ============================================================
#            APPLE SUPPLEMENTAL PASS (runs after main scrape)
# ============================================================

def run_apple_supplemental_scrape(wb, ws, apple_summary_ws, existing_pids: set) -> int:
    """
    Runs AFTER the normal 'mobiles' scrape finishes. Hits Flipkart's
    Apple-brand-facet URL page by page — same fetch/parse/Playwright-
    fallback machinery as the main loop — but only WRITES rows for PIDs
    not already present in the Products sheet (existing_pids), so
    iPhones already caught by the generic search aren't duplicated.

    Also carries the same early-exit protection as the main loop: if
    Flipkart's declared page count for the Apple facet is likewise
    inflated beyond what it actually serves, this stops after
    MAX_CONSECUTIVE_EMPTY_PAGES consecutive completely-empty pages
    instead of grinding through the rest.
    """
    logger.info("=" * 60)
    logger.info("SUPPLEMENTAL PASS: Apple brand-filter (fills gaps missed by generic search)")
    logger.info("=" * 60)

    scraped_apple_pages = get_scraped_pages(apple_summary_ws)

    first_url = build_apple_page_url(1)
    first_html = fetch_page_requests(first_url)
    products_p1 = parse_page(first_html) if first_html else []

    if is_partial_page(products_p1):
        logger.warning("Apple page 1: partial/empty via requests — forcing Playwright...")
        pw_html = fetch_page_playwright(first_url)
        if pw_html:
            pw_products = parse_page(pw_html)
            if len(pw_products) > len(products_p1):
                products_p1 = pw_products
                first_html = pw_html

    if not first_html and not products_p1:
        logger.error("Apple pass: could not reach Flipkart. Skipping supplemental scrape.")
        return 0

    total_apple_pages = determine_total_pages(first_html or "")
    logger.info(f"Apple filter: {total_apple_pages} pages detected.")

    new_rows_added = 0
    pages_since_save = 0
    consecutive_empty_pages = 0   # <-- early-exit tracker (Apple loop)

    def add_new_products(products: List[Dict], page_label) -> Tuple[List[Dict], int]:
        nonlocal new_rows_added
        fresh = [p for p in products if p["pid"] not in existing_pids]
        dupes = len(products) - len(fresh)
        for p in fresh:
            existing_pids.add(p["pid"])
        append_rows(ws, fresh, page_label)
        new_rows_added += len(fresh)
        return fresh, dupes

    if 1 not in scraped_apple_pages:
        fresh, dupes = add_new_products(products_p1, "Apple-p1")
        apple_summary_ws.append([1, len(products_p1), len(fresh), datetime.now().isoformat()])
        logger.info(f"Apple page 1: {len(products_p1)} found | {len(fresh)} NEW | {dupes} already captured")
        pages_since_save += 1
        consecutive_empty_pages = 0 if products_p1 else 1
    else:
        logger.info("Apple page 1 already scraped — skipping (resume).")

    for pg in range(2, total_apple_pages + 1):
        if pg in scraped_apple_pages:
            logger.info(f"[Apple page {pg}] already scraped, skipping")
            continue

        url = build_apple_page_url(pg)
        delay = random.uniform(REQUEST_DELAY_MIN, REQUEST_DELAY_MAX)
        logger.info(f"--- Apple page {pg}/{total_apple_pages} ({delay:.1f}s delay) ---")
        time.sleep(delay)

        products = fetch_and_parse_page(url, page_label=f"Apple page {pg}")

        if not products:
            consecutive_empty_pages += 1
            logger.error(f"Apple page {pg}: FAILED — no products extracted. "
                         f"({consecutive_empty_pages}/{MAX_CONSECUTIVE_EMPTY_PAGES} consecutive empty)")
            apple_summary_ws.append([pg, 0, 0, datetime.now().isoformat()])

            if consecutive_empty_pages >= MAX_CONSECUTIVE_EMPTY_PAGES:
                logger.warning(
                    f"⛔ {consecutive_empty_pages} consecutive completely empty Apple pages "
                    f"(pages {pg - consecutive_empty_pages + 1}-{pg}). Flipkart's Apple facet "
                    f"likely has fewer real result pages than the estimated {total_apple_pages}. "
                    f"Stopping Apple supplemental pass early."
                )
                break
        else:
            consecutive_empty_pages = 0
            fresh, dupes = add_new_products(products, f"Apple-p{pg}")
            apple_summary_ws.append([pg, len(products), len(fresh), datetime.now().isoformat()])
            logger.info(f"Apple page {pg}: {len(products)} found | {len(fresh)} NEW | {dupes} dupes | "
                        f"running new-row total: {new_rows_added}")

        pages_since_save += 1
        if pages_since_save >= SAVE_EVERY_N_PAGES:
            wb.save(OUTPUT_FILE)
            pages_since_save = 0
            logger.info(f"⚡️ Progress saved: {OUTPUT_FILE}")

    logger.info(f"Apple supplemental pass complete — {new_rows_added} new rows added.")
    return new_rows_added


# ============================================================
#                         MAIN LOOP
# ============================================================

def main():
    if DIAGNOSTIC_MODE:
        run_diagnostic_mode()
        return

    logger.info("Starting Flipkart Mobile Scraper v5")
    logger.info(f"Output: {OUTPUT_FILE}")

    wb, ws, summary_ws, apple_summary_ws = init_workbook()
    scraped_pages = get_scraped_pages(summary_ws)

    first_url = build_page_url(1)
    first_html_raw = fetch_page_requests(first_url)
    products_batch_1 = parse_page(first_html_raw) if first_html_raw else []

    if is_partial_page(products_batch_1):
        logger.warning("Page 1: partial/empty via requests — forcing Playwright...")
        pw_html = fetch_page_playwright(first_url)
        if pw_html:
            pw_products = parse_page(pw_html)
            if len(pw_products) > len(products_batch_1):
                products_batch_1 = pw_products
                first_html_raw = pw_html

    if not first_html_raw and not products_batch_1:
        logger.critical("Cannot reach Flipkart. Aborting.")
        close_shared_browser()
        return

    total_pages = determine_total_pages(first_html_raw or "")
    logger.info(f"Total pages to process: {total_pages}")

    processed_count = 0
    pages_since_save = 0
    consecutive_empty_pages = 0   # <-- early-exit tracker (main loop)

    try:
        if 1 not in scraped_pages:
            with_rating = sum(1 for p in products_batch_1 if p["avg_rating"])
            append_rows(ws, products_batch_1, 1)
            summary_ws.append([1, len(products_batch_1), with_rating, datetime.now().isoformat()])
            processed_count += len(products_batch_1)
            logger.info(f"Page 1 done: {len(products_batch_1)} products scraped.")
            pages_since_save += 1
            consecutive_empty_pages = 0 if products_batch_1 else 1
        else:
            logger.info("Page 1 already scraped — skipping (resume).")

        for pg in range(2, total_pages + 1):
            if pg in scraped_pages:
                logger.info(f"[page {pg}] already scraped, skipping")
                continue

            url = build_page_url(pg)
            delay = random.uniform(REQUEST_DELAY_MIN, REQUEST_DELAY_MAX)
            logger.info(f"--- Processing page {pg}/{total_pages} ({delay:.1f}s delay) ---")
            time.sleep(delay)

            products = fetch_and_parse_page(url, page_label=f"Page {pg}")

            if not products:
                consecutive_empty_pages += 1
                logger.error(f"Page {pg}: COMPLETE FAILURE. (Progress preserved) "
                             f"[{consecutive_empty_pages}/{MAX_CONSECUTIVE_EMPTY_PAGES} consecutive empty]")
                ws.append([pg, "[ERROR]", f"Failed to fetch page {pg}", url,
                           "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                           "", "", "", datetime.now().isoformat()])
                summary_ws.append([pg, 0, 0, datetime.now().isoformat()])

                if consecutive_empty_pages >= MAX_CONSECUTIVE_EMPTY_PAGES:
                    logger.warning(
                        f"⛔ {consecutive_empty_pages} consecutive completely empty pages "
                        f"(pages {pg - consecutive_empty_pages + 1}-{pg}) despite Flipkart's "
                        f"header claiming {total_pages} total pages. This almost always means "
                        f"the paginator has actually run out of real results (Flipkart's "
                        f"declared product/page count is frequently inflated well beyond what "
                        f"it actually serves). Stopping the main 'mobiles' pagination early — "
                        f"proceeding straight to the Apple supplemental pass."
                    )
                    break
            else:
                consecutive_empty_pages = 0
                with_rating = sum(1 for p in products if p["avg_rating"])
                append_rows(ws, products, pg)
                summary_ws.append([pg, len(products), with_rating, datetime.now().isoformat()])
                processed_count += len(products)
                logger.info(f"Page {pg} success: {len(products)} added "
                            f"({with_rating} with rating). Running total: {processed_count}")

            pages_since_save += 1
            if pages_since_save >= SAVE_EVERY_N_PAGES:
                wb.save(OUTPUT_FILE)
                pages_since_save = 0
                logger.info(f"⚡️ Progress saved: {OUTPUT_FILE}")

        # --- Supplemental Apple-only pass, runs after the main scrape ---
        # NOTE: this runs regardless of WHY the main loop's for-range ended
        # (natural completion at total_pages, OR the early-exit break above)
        # — so the Apple pass is never skipped just because the mobiles
        # pagination gave up early.
        if RUN_APPLE_SUPPLEMENTAL_PASS:
            existing_pids = get_existing_pids(ws)
            logger.info(f"Products sheet currently has {len(existing_pids)} unique PIDs "
                        f"before Apple supplemental pass.")
            apple_new = run_apple_supplemental_scrape(wb, ws, apple_summary_ws, existing_pids)
            processed_count += apple_new

    except KeyboardInterrupt:
        logger.warning("Interrupted by user — saving progress before exit...")

    finally:
        wb.save(OUTPUT_FILE)
        close_shared_browser()
        logger.info("=" * 60)
        logger.info(f"DONE! Total products collected this run: {processed_count}")
        logger.info(f"File saved: {OUTPUT_FILE}")
        logger.info("=" * 60)


if __name__ == "__main__":
    main()