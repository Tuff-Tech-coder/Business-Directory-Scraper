# Business Directory Scraper

![Python](https://img.shields.io/badge/Python-3.11+-3776AB?logo=python&logoColor=white)
![pandas](https://img.shields.io/badge/pandas-3.x-150458?logo=pandas&logoColor=white)
![Tests](https://img.shields.io/badge/tests-21%20passing-brightgreen)
![License](https://img.shields.io/badge/License-MIT-green)

A structured data-extraction pipeline that turns paginated business directory listings into a clean, formatted Excel workbook — business name, phone, address, website, category, rating and review count.

Built as a study in **resilient HTTP clients**: session reuse, request throttling, retry with backoff, defensive HTML parsing, and a full offline mode so the entire pipeline can be exercised without a single network call.

```bash
pip install -r requirements.txt
python directory_scraper.py --demo          # full pipeline, zero network calls
```

---

## ⚖️ Responsible use

This tool was written to practice HTTP client engineering and HTML parsing. Automated collection from any site is governed by that site's Terms of Service and `robots.txt`, and personal contact data is subject to privacy law (GDPR, CCPA, and similar).

**Before pointing this at any live site:**

1. Read that site's Terms of Service and `robots.txt`, and honour them.
2. Prefer an official API or a licensed data provider where one exists.
3. Keep the built-in rate limiting on. Do not raise `--max-pages` to hammer a host.
4. Treat any collected personal data as regulated — have a lawful basis before storing or contacting.

`--demo` mode is the default path for evaluating this project and generates realistic synthetic data locally. **The maintainer does not endorse using this against any site whose terms prohibit automated access.**

---

## Engineering notes

The interesting problems here were not the parsing — they were everything around it.

**Transport layer.** Uses [`curl_cffi`](https://github.com/lexiforest/curl_cffi) rather than `requests`. Directory sites commonly sit behind CDN edge protection that rejects Python's default TLS handshake, so a request that a browser completes fine returns `403` from `requests`. `curl_cffi` performs a browser-equivalent TLS handshake, which makes the client behave like the browser a human would use. Understanding *why* a request fails at the transport layer rather than the application layer is the transferable skill.

**Politeness by construction.** Randomized 3–7 second delays between pages, one persistent session for cookie and connection reuse, and a hard page cap. The client is designed to place less load on a host than a person clicking through the same pages.

**Defensive parsing.** Every field extraction tolerates a missing element and falls back to an empty string rather than raising, so one malformed card cannot abort a run. Ad placements are filtered from organic results by CSS class.

**Structured output.** A `Business` dataclass is the single record type; `dataclasses.asdict` feeds the DataFrame, so adding a field is a one-line change. Excel export applies header styling, alternating row shading and auto-fit column widths.

**Observability.** Structured logging to both console and file, so a long run can be diagnosed after the fact.

---

## Usage

```bash
# Offline — generates 35 realistic synthetic listings
python directory_scraper.py --demo

# Live (only against a source whose terms permit it)
python directory_scraper.py --category "plumbers" --city "Austin, TX" --max-pages 3
```

| Flag | Default | Description |
|---|---|---|
| `--category` | `plumbers` | Business type to search |
| `--city` | `Austin, TX` | City, e.g. `"Chicago, IL"` |
| `--max-pages` | `5` | Hard cap on pages fetched |
| `--output` | auto | Output `.xlsx` filename |
| `--demo` | off | Generate synthetic data, no network |

## Output

| Business Name | Phone | Address | City / State / ZIP | Website | Category | Rating | Reviews | Date Scraped |
|---|---|---|---|---|---|---|---|---|
| Ace Plumbing & Drain | (512) 555-0142 | 1820 Commerce Blvd | Austin, TX 78701 | https://… | Plumbers | 4/5 | 87 | 2026-07-30 |

## Development

```bash
pip install -r requirements.txt
pip install pytest ruff

pytest -q          # 21 tests
ruff check .
```

The suite runs entirely against saved HTML fixtures in `tests/fixtures.py`, so
it needs no network access and is deterministic in CI. It covers the three
parts that carry the logic:

- **`parse_listing`** — defensive extraction. A card missing every optional
  element must return empty strings rather than raise, an unrated listing must
  render as `""` and not `"0/5"`, and an internal redirect `href` must not be
  recorded as the business's website.
- **`parse_results_page`** — paid placements are filtered out by their
  `flash-endt` class, unparseable cards are dropped, and relative next-page
  links are resolved to absolute URLs while absolute ones are left alone.
- **`export_to_excel`** — the workbook reads back with the expected header row
  and row count, partial records export as blanks, and an empty result set
  writes no file at all rather than an empty workbook.

## Tech stack

`Python` · `curl_cffi` · `BeautifulSoup` · `pandas` · `openpyxl` · `dataclasses` · `argparse` · `logging`

## License

MIT — see [LICENSE](LICENSE).
