"""
Business Directory Scraper
==========================
A structured data-extraction pipeline that turns paginated directory listings
into a formatted Excel workbook. Extracts business name, phone, address,
website, rating and category, walks pagination automatically, and throttles
requests with randomised delays.

The parsing and export layers are source-agnostic: parse_listing,
parse_results_page and export_to_excel operate on HTML and dataclasses rather
than on any particular site, and are unit-tested against saved fixtures.

Responsible use
---------------
The live-fetch path targets Yellow Pages. Check a site's Terms of Service and
robots.txt before pointing this at it, and prefer an official API or a
licensed data provider where one exists. --demo runs the entire pipeline
against generated data with no network access, and is the intended path for
evaluating this project.

Usage:
    python directory_scraper.py --demo    # full pipeline, no network requests
    python directory_scraper.py --category "plumbers" --city "Austin, TX"
    python directory_scraper.py --category "dentists" --city "Chicago, IL" --max-pages 5

Output: A formatted Excel file with all results.
"""

import re
import time
import random
import logging
import argparse
import datetime
from dataclasses import dataclass, asdict
from pathlib import Path

from curl_cffi import requests as cffi_requests
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(), logging.FileHandler("directory_scraper.log", encoding="utf-8")],
)
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------
@dataclass
class Business:
    name: str = ""
    phone: str = ""
    address: str = ""
    city_state: str = ""
    website: str = ""
    category: str = ""
    rating: str = ""
    review_count: str = ""
    date_scraped: str = ""

    def __post_init__(self):
        if not self.date_scraped:
            self.date_scraped = datetime.date.today().isoformat()


# ---------------------------------------------------------------------------
# HTTP fetch
# ---------------------------------------------------------------------------
def fetch_page(url: str, session: cffi_requests.Session, retries: int = 3) -> BeautifulSoup | None:
    """
    Fetch a URL and return a BeautifulSoup object, or None on failure.

    Retries up to `retries` times with linearly increasing backoff, and treats
    a suspiciously small response body as a failure rather than parsing it --
    an error or interstitial page is usually far shorter than a real results
    page, so length is a cheap sanity check before handing bytes to the parser.

    Returns None rather than raising: the caller decides whether a failed page
    ends the run, so one bad page cannot abort a long pagination walk.
    """
    for attempt in range(1, retries + 1):
        try:
            resp = session.get(
                url,
                headers={
                    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
                    "Accept-Language": "en-US,en;q=0.9",
                    "Referer": "https://www.google.com/",
                    "Upgrade-Insecure-Requests": "1",
                },
                timeout=20,
            )
            if resp.status_code == 200 and len(resp.content) > 10_000:
                return BeautifulSoup(resp.text, "lxml")
            logger.warning(f"Unexpected response on attempt {attempt}: status={resp.status_code} size={len(resp.content)}")
        except Exception as e:
            logger.warning(f"Request error on attempt {attempt}: {e}")

        if attempt < retries:
            wait = 5 * attempt
            logger.info(f"Waiting {wait}s before retry...")
            time.sleep(wait)

    return None


# ---------------------------------------------------------------------------
# Parser — uses selectors confirmed against live Yellow Pages HTML
# ---------------------------------------------------------------------------
def parse_listing(card: BeautifulSoup) -> Business | None:
    """Parse a single Yellow Pages organic listing card."""

    def safe_text(tag_name: str, attrs: dict) -> str:
        el = card.find(tag_name, attrs)
        return el.get_text(strip=True) if el else ""

    name = safe_text("a", {"class": "business-name"})
    if not name:
        return None

    # Phone is in a div with classes "phones phone primary"
    phone_el = card.find("div", class_="phones")
    phone = phone_el.get_text(strip=True) if phone_el else ""

    # Address: street is div.street-address, city/state is div.locality
    street = safe_text("div", {"class": "street-address"})
    locality = safe_text("div", {"class": "locality"})

    # Website — the actual external URL is in the href attribute
    website = ""
    ws_tag = card.find("a", class_="track-visit-website")
    if ws_tag:
        href = ws_tag.get("href", "")
        # YP internal redirect URLs start with /; real external URLs start with http
        if href.startswith("http"):
            website = href.split("?")[0]  # strip YP tracking params

    # Categories
    cats_el = card.find("div", class_="categories")
    category = cats_el.get_text(separator=", ", strip=True) if cats_el else ""

    # Rating — star level is encoded as a CSS class like "result-rating three"
    rating = ""
    rating_el = card.find("div", class_="result-rating")
    if rating_el:
        classes = rating_el.get("class", [])
        word_map = {"one": "1", "two": "2", "three": "3", "four": "4", "five": "5"}
        for cls in classes:
            if cls in word_map:
                rating = word_map[cls]
                break

    # Review count is in span.count inside the ratings div
    count_el = card.find("span", class_="count")
    review_count = re.sub(r"[^\d]", "", count_el.get_text()) if count_el else ""

    return Business(
        name=name,
        phone=phone,
        address=street,
        city_state=locality,
        website=website,
        category=category,
        rating=f"{rating}/5" if rating else "",
        review_count=review_count,
    )


def parse_results_page(soup: BeautifulSoup) -> tuple[list[Business], str | None]:
    """
    Parse all organic listings from a page.
    Returns (listings, next_page_url).
    """
    # Organic results don't have the 'flash-endt' ad class
    all_cards = soup.find_all("div", class_="result")
    organic_cards = [c for c in all_cards if "flash-endt" not in c.get("class", [])]

    listings = []
    for card in organic_cards:
        biz = parse_listing(card)
        if biz and biz.name:
            listings.append(biz)

    # Next page link
    next_link = soup.find("a", class_="next")
    next_url = None
    if next_link and next_link.get("href"):
        href = next_link["href"]
        next_url = f"https://www.yellowpages.com{href}" if href.startswith("/") else href

    return listings, next_url


# ---------------------------------------------------------------------------
# Main scraper
# ---------------------------------------------------------------------------
class DirectoryScraper:
    SEARCH_URL = "https://www.yellowpages.com/search"

    def __init__(self, category: str, city: str):
        self.category = category
        self.city = city
        # One persistent session shares cookies and connection pool
        self.session = cffi_requests.Session(impersonate="chrome124")

    def _first_url(self) -> str:
        from urllib.parse import urlencode
        params = {"search_terms": self.category, "geo_location_terms": self.city}
        return f"{self.SEARCH_URL}?{urlencode(params)}"

    def scrape(self, max_pages: int = 10) -> list[Business]:
        all_results: list[Business] = []
        url = self._first_url()

        for page_num in range(1, max_pages + 1):
            logger.info(f"Scraping page {page_num}: {url}")

            soup = fetch_page(url, self.session)
            if soup is None:
                logger.warning(f"Could not load page {page_num} - stopping")
                break

            page_results, next_url = parse_results_page(soup)
            logger.info(f"Page {page_num}: {len(page_results)} organic listings")
            all_results.extend(page_results)

            if not next_url or not page_results:
                logger.info("No more pages")
                break

            url = next_url

            if page_num < max_pages:
                delay = random.uniform(3, 7)
                logger.info(f"Waiting {delay:.1f}s...")
                time.sleep(delay)

        logger.info(f"Total collected: {len(all_results)}")
        return all_results


# ---------------------------------------------------------------------------
# Demo data generator (no live requests)
# ---------------------------------------------------------------------------
def generate_demo_results(category: str, city: str, count: int = 35) -> list[Business]:
    """Generate realistic-looking demo business listings."""
    city_name = city.split(",")[0].strip()
    state = city.split(",")[1].strip() if "," in city else "TX"
    streets = [
        "Main St", "Oak Ave", "Maple Dr", "Cedar Ln", "Pine St",
        "Commerce Blvd", "Industrial Pkwy", "Business Park Dr", "Market St",
        "First Ave", "Second St", "Third Ave", "Riverside Dr", "Airport Blvd",
        "University Ave", "Technology Dr", "Center St", "Heritage Way",
    ]
    zip_codes = ["78701", "78702", "78703", "78704", "78741", "78745", "78748"]
    prefixes = [city_name, "Texas", "Lone Star", "Premier", "Ace", "Pro",
                "All-Star", "Quality", "Expert", "Master", "Elite",
                "Reliable", "Quick", "24-Hour", "Family", "Local"]
    suffix_map = {
        "plumbers": ["Plumbing", "Plumbing & Drain", "Pipe Works", "Drain Services",
                     "Water Works", "Plumbing Solutions", "Sewer & Drain"],
        "dentists": ["Dental", "Dentistry", "Dental Care", "Smile Center",
                     "Family Dental", "Dental Group", "Oral Health"],
        "electricians": ["Electric", "Electrical", "Electrical Services",
                         "Power Solutions", "Wiring Experts", "Electric Co"],
    }
    suffixes = suffix_map.get(category.lower(), [f"{category.title()} Services", f"{category.title()} Pros"])
    area_codes = ["512", "737", "254", "210"]
    ratings_pool = ["3/5", "4/5", "5/5", "4/5", "5/5", "4/5", "3/5", ""]

    results = []
    for i in range(count):
        name = f"{random.choice(prefixes)} {random.choice(suffixes)}"
        if i % 5 == 4:
            first = random.choice(["John", "Mike", "Sarah", "David", "Lisa", "Carlos"])
            last = random.choice(["Smith", "Johnson", "Williams", "Brown", "Davis"])
            name = f"{first} {last}'s {random.choice(suffixes)}"

        area = random.choice(area_codes)
        phone = f"({area}) {random.randint(200,999)}-{random.randint(1000,9999)}"
        zipcode = random.choice(zip_codes)
        address = f"{random.randint(100,9999)} {random.choice(streets)}"
        locality = f"{city_name}, {state} {zipcode}"
        clean = name.lower().replace(" ", "").replace("'", "")
        website = f"https://www.{clean}.com" if random.random() > 0.35 else ""
        review_count = str(random.randint(3, 280)) if random.random() > 0.2 else ""

        results.append(Business(
            name=name, phone=phone, address=address, city_state=locality,
            website=website, category=category.title(),
            rating=random.choice(ratings_pool), review_count=review_count,
        ))
    return results


# ---------------------------------------------------------------------------
# Excel export
# ---------------------------------------------------------------------------
def export_to_excel(businesses: list[Business], output_path: str) -> None:
    """Write business results to a formatted Excel file."""
    if not businesses:
        logger.warning("No results to export")
        return

    rows = [{
        "Business Name": b.name,
        "Phone": b.phone,
        "Address": b.address,
        "City / State / ZIP": b.city_state,
        "Website": b.website,
        "Category": b.category,
        "Rating": b.rating,
        "Review Count": b.review_count,
        "Date Scraped": b.date_scraped,
    } for b in businesses]

    df = pd.DataFrame(rows)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Results")
        ws = writer.sheets["Results"]

        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="C0392B", end_color="C0392B", fill_type="solid")
        alt_fill = PatternFill(start_color="FDF2F2", end_color="FDF2F2", fill_type="solid")

        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center")

        for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
            if row_idx % 2 == 0:
                for cell in row:
                    cell.fill = alt_fill

        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col) + 2
            ws.column_dimensions[col[0].column_letter].width = min(max_len, 55)

    logger.info(f"Exported {len(businesses)} businesses to {output_path}")


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description="Business Directory Scraper (Yellow Pages)")
    parser.add_argument("--category", default="plumbers", help="Business type to search")
    parser.add_argument("--city", default="Austin, TX", help="City to search (e.g. 'Austin, TX')")
    parser.add_argument("--max-pages", type=int, default=5, help="Max pages to scrape")
    parser.add_argument("--output", default=None, help="Output Excel filename")
    parser.add_argument("--demo", action="store_true", help="Generate demo data (no live requests)")
    args = parser.parse_args()

    city_slug = args.city.split(",")[0].strip().replace(" ", "_")
    output = args.output or f"{args.category.replace(' ', '_')}_{city_slug}.xlsx"

    logger.info(f"Searching: '{args.category}' in '{args.city}'")

    if args.demo:
        logger.info("Demo mode: generating sample results...")
        results = generate_demo_results(args.category, args.city)
    else:
        scraper = DirectoryScraper(category=args.category, city=args.city)
        results = scraper.scrape(max_pages=args.max_pages)

    if results:
        export_to_excel(results, output)
        print(f"\n[OK] {len(results)} businesses saved to: {output}")
    else:
        print("No results found. Check the log for details.")


if __name__ == "__main__":
    main()
