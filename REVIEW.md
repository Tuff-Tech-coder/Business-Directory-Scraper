# Debugging and verification record

Reviewed on 2026-09-27. Baseline: the original 21 tests and lint check passed. Twelve additional regression cases were then run against the original code, and all twelve failed before the fixes.

## Reproduced defects fixed

| Defect | Result after repair |
|---|---|
| All website query parameters were discarded | Known tracking parameters are removed while functional parameters survive |
| Half-star ratings were rounded down | `four half` parses as 4.5/5 |
| Adjacent text elements merged words | Names preserve spaces between text nodes |
| Query-relative next links stayed relative | Links such as `?page=2` resolve against the current page |
| Protocol-relative links were prefixed incorrectly | `//www.yellowpages.com/...` resolves correctly |
| Zero page limits silently did nothing | API and CLI validation reject nonpositive limits |
| Exports failed when a parent folder was missing | Parent directories are created |
| Formula-like business names became Excel formulas | Scraped text is explicitly stored as literal strings |
| Illegal control characters crashed Excel export | Unsupported control characters are removed |
| Importing the module created a log file | Logging configuration occurs only at the CLI boundary |
| Demo URLs could point to real businesses or contain invalid characters | Demo URLs use normalized slugs and reserved `.example` domains |
| Chicago demo records used Austin ZIP codes | Supported cities have matching location labels; unknown locations omit ZIP codes |

## Other repairs and additions

- Replaced the arbitrary 10 KB HTTP body threshold with recognizable-markup validation.
- Distinguished permanent HTTP failures from retryable failures and unrecognized interstitials.
- Added repeated-pagination detection, same-host traversal, duplicate removal, and session cleanup.
- Preserved partial data while returning a nonzero exit code and marking the workbook as partial.
- Added argument validation, portable default filenames, typed Excel values, frozen headers, filters, source notes, and atomic writes.
- Reworked demo execution so generated HTML goes through production parsing and pagination.
- Added a loopback-only dashboard with search, combined filters, pagination, actual workbook generation, and a mobile layout.
- Replaced the bundled sample workbook with a freshly generated, labeled synthetic export.

## Verification

- **60 tests passed** locally on Windows / Python 3.12.
- **Ruff lint passed**; Python files formatted with Ruff.
- **JavaScript syntax check passed** using Node.
- CLI demo: 35 Austin plumbers, 3 pages, 2 duplicates removed, 6 sponsored cards excluded.
- CLI and browser alternate scenario: 75 Chicago dentists, 7 pages, 6 duplicates removed, 14 sponsored cards excluded.
- Browser: next/previous page controls, 4+ star and website filters, empty search results, disabled empty export, and filter reset after a new run.
- Filtered Excel API response saved and read back: **11 data rows**, with the same names as the matching records.
- Desktop and mobile layouts visually inspected. Mobile has no page-level horizontal overflow; the results table scrolls within its container.
- No browser console errors or warnings were observed in the checked run.

## Limits of verification

Live Yellow Pages requests were not made. HTTP behavior is covered with mocked responses, so current external access and selector compatibility are not established. Remote CI and Python 3.11 execution were not run locally.

The in-app browser reported that the workbook was generated, and the API returned a valid Excel attachment. Its download-event observer timed out, so delivery into the browser's Downloads location was not independently confirmed. The same API response was saved directly as an output and verified with openpyxl. Native Excel inspection was unavailable because computer-use access to Excel was not approved; screenshot 4 shows the application's workbook preview.

The Windows convenience launcher is provided with readable setup steps; the program itself was launched and tested directly with the isolated Python environment used for this review.
