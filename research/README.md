# Tariq El-Jdideh directory research

The user-facing Excel workbook is in `deliverables/Tariq_El_Jdideh_Shop_Directory_2026-10-16.xlsx`.

This is a dated **public-source research compilation**, not a fully verified current business directory. No businesses were called, messaged or visited. The workbook explicitly separates eight source-backed listings from four unconfirmed leads, retains unresolved operating-hours conflicts, and includes field-level references and a source access/freshness ledger. Geographic coverage includes source-identified Kaskas/El-Horge commercial streets and is not an independently surveyed boundary.

The source data and short evidence transcriptions are embedded in `build_shop_directory.py`. They were manually compiled using web search/page retrieval on 2026-10-16. Re-running this script only regenerates the workbook; it does **not** perform scraping, renew the source review date, or verify business details.

## Rebuild

```sh
python -m venv .venv
.venv/bin/pip install -r research/requirements.txt
.venv/bin/python research/build_shop_directory.py
```

The generator validates unique business/source IDs, source references, phone formatting, worksheet row counts, phone text types, date formatting and the absence of formulas. These are file/data integrity checks, not real-world verification. Complete the workbook's Confirmation log after actual business outreach to establish current information.

## Arabic edition

The Arabic workbook is `deliverables/Tariq_El_Jdideh_Shop_Directory_AR_2026-10-16.xlsx`.
Regenerate it from the existing English workbook with:

```sh
.venv/bin/python research/build_arabic_directory.py
```

It translates sheet names, headings, shop details, evidence notes, instructions and the confirmation dropdown; enables right-to-left layout; and retains source excerpts in their original language for auditing. Phone numbers, source IDs, dates and URLs are unchanged. The script validates preservation of contact details, links, dates, verbatim excerpts and table metadata. Translation is not new verification.

## Compact Excel and offline browser edition

```sh
.venv/bin/python research/build_lightweight_directory.py
```

This reads the existing Arabic workbook and creates:

- `deliverables/shops_ar_small.xlsx`: a real Excel workbook with two worksheets. All 12 shop records, original contact data, cautions, source references and the complete 27-record source ledger are retained. Duplicate field-evidence rows, the empty follow-up log and the separate long instructions sheet are omitted; an explicit verification warning and record-tier column remain.
- `deliverables/index.html`: a self-contained right-to-left Arabic page. Open it directly in any modern browser, including offline. It has name/phone/address search, activity filters, expandable cautions and source links. No library, font, image, server or network request is required to display or filter the directory. Following an external source link requires internet access. Verbatim source excerpts are retained in Excel, not duplicated in the browser version.

Generation validates exact preservation of shop/source rows in Excel, file integrity, HTML source targets and all shop names, primary phones, locations, hours and cautions. The compact Excel file is 19,806 bytes (47% smaller than the full Arabic edition); the standalone HTML page is approximately 41 KB. Sizes can vary slightly on regeneration.

Optional browser DOM smoke test, with jsdom installed outside the Python dependencies:

```sh
JSDOM_MODULE=/path/to/node_modules/jsdom node research/test_browser_directory.cjs
```

The requested list of 667 stores has not been provided. These editions contain the same 12 original researched listings and do not add or reverify businesses.

The repository's original `scrape.py`, `urls.txt` and scraped content are unchanged.
