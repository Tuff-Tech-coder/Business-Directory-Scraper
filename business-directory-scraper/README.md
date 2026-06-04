# Business Directory Scraper

A Python web scraper that extracts business contact information from Yellow Pages given a city and category. Automatically handles pagination, applies rate limiting with random delays, and outputs a clean, formatted Excel file.

## Features

- Searches Yellow Pages by **business category** and **city**
- Extracts: business name, phone, address, city/state/ZIP, website, category, rating, review count
- **Auto-pagination**: keeps scraping until no more pages (configurable max)
- **Rate limiting**: random 2–6 second delays between page requests
- Uses **`curl_cffi`** to impersonate Chrome's TLS fingerprint — bypasses Cloudflare protection that blocks standard Python `requests` (verified working as of May 2026)
- Retry logic with exponential backoff on HTTP errors
- **Demo mode**: generates 35 realistic sample listings without any live requests
- Formatted Excel output with alternating row colors and auto-fit columns

## Project Structure

```
business-directory-scraper/
├── directory_scraper.py           # Main scraper
├── requirements.txt               # Python dependencies
├── plumbers_austin_live.xlsx      # Live-scraped output (90 real listings, 3 pages)
├── plumbers_austin_sample.xlsx    # Demo-generated output (35 listings)
└── directory_scraper.log          # Application log (created on run)
```

## Setup

```bash
python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate
pip install -r requirements.txt
```

## Usage

```bash
# Live scraping — plumbers in Austin, TX
python directory_scraper.py --category "plumbers" --city "Austin, TX"

# Different category and city
python directory_scraper.py --category "dentists" --city "Chicago, IL"

# Limit to 3 pages
python directory_scraper.py --category "electricians" --city "Denver, CO" --max-pages 3

# Demo mode (no live requests)
python directory_scraper.py --demo --category "plumbers" --city "Austin, TX"

# Custom output filename
python directory_scraper.py --category "restaurants" --city "New York, NY" --output nyc_restaurants.xlsx
```

## Command-Line Options

| Option | Default | Description |
|---|---|---|
| `--category` | `plumbers` | Business type to search (e.g. dentists, electricians, restaurants) |
| `--city` | `Austin, TX` | City and state |
| `--max-pages` | `5` | Maximum result pages to scrape |
| `--output` | auto-generated | Output Excel filename |
| `--demo` | off | Generate sample data without live requests |

## Sample Output (`plumbers_austin_sample.xlsx`)

| Business Name | Phone | Address | City / State / ZIP | Website | Category | Rating | Review Count |
|---|---|---|---|---|---|---|---|
| Austin Plumbing | (512) 555-0123 | 4821 Main St | Austin, TX 78704 | https://... | Plumbers | 4.5 | 127 |
| Premier Pipe Works | (737) 555-0456 | 1156 Commerce Blvd | Austin, TX 78741 | | Plumbers | 4.0 | 43 |
| ... | ... | ... | ... | ... | ... | ... | ... |

## Notes

- Yellow Pages is protected by Cloudflare. Standard `requests` gets a 403; `curl_cffi` with `impersonate="chrome124"` passes cleanly by matching Chrome's TLS fingerprint.
- Yellow Pages selectors may change over time. If results come back empty, use browser dev tools to inspect the current HTML structure and update the selector logic in `parse_listing()`.
- For high-volume scraping (many cities or categories), add proxy rotation to avoid IP rate limiting. Services like Bright Data or ScraperAPI work well for this.
- The demo mode is a reliable way to show clients the output format and code structure without depending on live site access.
