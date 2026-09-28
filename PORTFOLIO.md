# Portfolio presentation

## Project title

**Directory Lab — a testable business-data extraction pipeline**

## Short description

A Python application that transforms paginated business-directory HTML into structured records and formatted Excel workbooks. The project combines defensive parsing, retry logic, duplicate removal, and explicit failure reporting with an interactive local dashboard. A reproducible offline demo makes the parsing and export workflow easy to evaluate.

## LinkedIn post draft

Project spotlight: Directory Lab, a Python business-directory extraction project.

The interesting work was making the data trustworthy: preserving half-star ratings and functional website parameters, resolving pagination links, removing duplicate records, and keeping untrusted text from becoming Excel formulas.

The project now includes a local dashboard for searching and filtering businesses, inspecting page-by-page processing, and exporting the matching results. Its offline demonstration feeds generated HTML through the same parser and pagination logic used by the scraper.

The current build passes 60 automated tests covering parsing, transport failure handling, pagination, CLI behavior, and Excel/API exports.

The screenshots use clearly labeled synthetic data. Live-site compatibility is a separate validation step.

Stack: Python, BeautifulSoup, curl_cffi, pandas, openpyxl, and vanilla JavaScript.

#Python #DataEngineering #SoftwareDevelopment #Portfolio

## Resume bullet

Developed a Python directory-data pipeline with defensive HTML parsing, bounded retries, pagination safeguards, duplicate removal, and formatted Excel exports; added a reproducible offline dashboard and a 60-test regression/integration suite.

## Screenshot sequence and captions

1. **Results overview** — 35 synthetic businesses collected from three generated HTML pages. Summary metrics are calculated from the resulting records.
2. **Filtered results** — Combine a 4+ star threshold with website availability to narrow the dataset to 11 businesses. Search and pagination work on the filtered records.
3. **Pipeline activity** — Inspect how each page contributes to the final dataset, including two repeated listings removed and six sponsored cards excluded.
4. **Excel export** — Preview the 11 matching records before downloading a workbook with typed ratings/review counts, filters, frozen headers, and source notes. This screenshot is the application's workbook preview, not Microsoft Excel.

Lead with screenshot 1, followed by 3 and 4 for an engineering-focused portfolio. Screenshot 2 is useful when emphasizing the interface and interaction design.

## Interview demonstration (about 90 seconds)

1. Launch the local dashboard and explain that the demo uses generated HTML with real parser execution.
2. Show the 35-record overview and describe how missing values are represented.
3. Select 4+ stars and Has website; show the 11 matching results.
4. Open Pipeline activity and explain duplicate detection, ad filtering, and page boundaries.
5. Open Excel export and download the matching workbook.
6. Discuss one reproduced defect and its regression test, such as preserving a functional `location=2` website parameter while removing `utm_source`.

Avoid claims about production scale, time savings, live scraping success, or customers unless you separately measure or establish them. The verified strengths are the working local pipeline, code quality, test coverage, and inspectable outputs.
