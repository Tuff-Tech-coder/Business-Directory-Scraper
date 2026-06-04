# Business Directory Scraper

![Python](https://img.shields.io/badge/Python-3.10+-3776AB?logo=python&logoColor=white)
![curl_cffi](https://img.shields.io/badge/curl__cffi-TLS%20impersonation-FF6C37)
![pandas](https://img.shields.io/badge/pandas-2.x-150458?logo=pandas&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green)

A lead-generation scraper that extracts business contact information from Yellow Pages for any category and city — business name, phone, address, website, rating, and review count — and exports it to a formatted Excel file ready for outreach.

Its standout feature: it uses **`curl_cffi` to impersonate Chrome's TLS fingerprint**, which passes the Cloudflare bot protection that returns a `403` to standard Python HTTP clients. This is a practical demonstration of getting past modern, fingerprint-based anti-bot defenses cleanly.

---

## Why it's useful

Building a prospect list by copying entries off a directory site is tedious and incomplete. This tool collects an entire category across a city in minutes, deduplicated into a tidy spreadsheet — and it does so politely, with rate limiting and retries, so it behaves like a considerate client rather than a hammer.

---

## Features

- **TLS-fingerprint impersonation** — `curl_cffi` with `chrome124` passes Cloudflare's fingerprint check that blocks plain `requests`/`urllib`.
- **Search by category + city** with results parsed from confirmed live Yellow Pages selectors.
- **Rich extraction** — name, phone, street address, city/state/ZIP, external website (with tracking params stripped), category, star rating, and review count.
- **Automatic pagination** — follows the "next" link until results end or a page cap is reached.
- **Polite & resilient** — randomized 3–7s delays between pages, a persistent session for cookies/connection reuse, and retry logic with backoff.
- **Formatted Excel output** — styled header, alternating row shading, and auto-fit columns.
- **Demo mode** generates 35 realistic sample listings with no live requests.

---

## Tech stack

`Python` · `curl_cffi` · `BeautifulSoup` · `pandas` · `openpyxl` · `dataclasses` · `argparse` · `logging`

---

## Project structure

```
business-directory-scraper/
├── directory_scraper.py           # Main scraper
├── requirements.txt               # Python dependencies
├── plumbers_austin_live.xlsx      # Live-scraped sample output
├── plumbers_austin_sample.xlsx    # Demo-generated sample output
└── directory_scraper.log          # Generated: application log
```

---

## Setup

```bash
python -m venv venv
source venv/bin/activate          # Windows: venv\Scripts\activate
pip install -r requirements.txt
```

---

## Usage

```bash
# Live scrape
python directory_scraper.py --category "plumbers" --city "Austin, TX"
python directory_scraper.py --category "dentists" --city "Chicago, IL" --max-pages 5

# Sample data, no network
python directory_scraper.py --demo
```

**Options:** `--category` · `--city` · `--max-pages` · `--output` · `--demo`

Output defaults to `<category>_<city>.xlsx` (e.g. `plumbers_Austin.xlsx`).

---

## How it works

1. Opens a persistent `curl_cffi` session impersonating Chrome's TLS fingerprint.
2. Builds the Yellow Pages search URL from the category and city.
3. Fetches each page (validating status and response size), parses organic listings while skipping ads, and follows the "next" link.
4. Throttles between pages and retries with backoff on failures.
5. Exports all results to a styled Excel workbook.

---

## Note on responsible use

This project is for educational and legitimate lead-generation purposes. Review and respect each target site's Terms of Service and `robots.txt`, and keep the built-in rate limiting in place. Anti-bot defenses change over time, so live selectors may need occasional maintenance.

---

## Possible extensions

Add email-address discovery, CRM export (HubSpot/Salesforce), per-record deduplication across runs, or a queue of categories/cities for batch collection.
