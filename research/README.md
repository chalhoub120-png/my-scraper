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

The repository's original `scrape.py`, `urls.txt` and scraped content are unchanged.
