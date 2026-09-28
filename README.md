# Payer Policy Document Discovery

A lightweight Python crawler that discovers publicly available provider/medical-policy documents for the 10 payer seeds supplied with the assessment.

## Features

- Reads all payers from `data/payer_seed_list.csv`.
- Checks `robots.txt` and respects `Allow`, `Disallow`, `Crawl-delay`, and declared sitemaps.
- Discovers documents through public HTML links, sitemaps, and URL patterns.
- Follows allowed subdomains and verifies candidate documents before output.
- Extracts title, document type, state/region, line of business, policy number and explicitly labelled dates.
- Records `source_page_url`, `discovery_path`, and the actual `extraction_method`.
- Supports `medical_policy`, `pa_list`, and `drug_list` classification.
- Uses retries/backoff for temporary HTTP failures.
- Does not log in, use credentials, or bypass CAPTCHA/WAF/bot protection.
- Creates a new timestamped output set for every run.
- Resumes interrupted runs using a simple `scraper_state.json` file.

## Project structure

```text
Scrape-Projects/
├── scraper.py
├── data/payer_seed_list.csv
├── requirements.txt
├── README.md
├── NOTES.md
└── scraper_state.json       # created during a run
```

Each run creates:

```text
PayerPolicyDocumentDiscovery_YYYYMMDD_HHMMSS.csv
PayerPolicyDocumentDiscovery_YYYYMMDD_HHMMSS.xlsx
PayerPolicyDocumentDiscovery_YYYYMMDD_HHMMSS_run_log.jsonl
PayerPolicyDocumentDiscovery_YYYYMMDD_HHMMSS_summary.json
```

## Setup

Python 3.10+ is recommended.

```powershell
python -m venv .venv
Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Dependencies:

- `requests` — HTTP crawling
- `pypdf` — PDF extraction
- `openpyxl` — XLSX output

## Run

Run all 10 payers:

```powershell
python scraper.py --fresh
```

Test one payer:

```powershell
python scraper.py --payer UHC --fresh
```

Resume an interrupted run:

```powershell
python scraper.py
```

Use `--fresh` to start a completely new run.

## Output

The CSV/XLSX contains the required 22 columns in the specified order:

```text
payer_name, payer_alias, state_or_region, line_of_business,
document_title, document_type, document_url, source_page_url,
discovery_path, file_type, policy_number, effective_date,
last_updated_date, http_status, content_hash_sha256,
file_size_bytes, requires_auth, render_mode, extraction_method,
confidence_score, scrape_timestamp_utc, notes
```

`document_url` values are deduplicated.

`extraction_method` records the techniques actually used, for example:

```text
sitemap|url_pattern|pdf_metadata|pdf_text
html_anchor|url_pattern|pdf_text
html_title|html_text|labeled_text
```

### Run log

`*_run_log.jsonl` is one JSON object per line and is both machine-readable and easy to inspect. Typical events include:

```text
run.start
payer.start
robots.loaded
robots.blocked
fetch.failed
doc.found
payer.done
run.done
```

Important events are also printed clearly in the terminal.

### Summary

`*_summary.json` contains run status, timestamps, payer count, per-payer `attempted/found/failed/skipped`, overall totals, and the names of the generated files.

## Safety and crawl limits

The crawler only uses public, permitted access. It does not log in or bypass CAPTCHA/WAF/bot protection. Robots rules are checked before requests; crawl delays are respected; temporary `429`/`5xx` failures are retried with backoff; and crawling/document discovery is bounded.

## Resumability

The assessment requires resumability. Instead of a database, `scraper_state.json` records the current run and completed payers. If a run stops, `python scraper.py` continues the unfinished run. `--fresh` starts a new run.

## Architecture

```text
seed CSV → robots.txt → sitemap/HTML discovery
         → candidate document → verification
         → metadata extraction → CSV/XLSX + log + summary
```

The implementation is intentionally kept in one Python file so the full flow is easy to review and explain.
