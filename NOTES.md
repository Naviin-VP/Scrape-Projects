# Notes

## 1. Approach

The scraper uses a small, single-file Python design so the crawl flow is easy to review and explain.

For each payer it:

1. Reads the seed host from `data/payer_seed_list.csv`.
2. Fetches and parses `robots.txt`.
3. Applies `Allow`, `Disallow`, and `Crawl-delay` rules.
4. Uses declared sitemaps and public HTML pages as discovery sources.
5. Follows allowed subdomains and resolves document links.
6. Identifies candidate policy/document URLs using URL and page signals.
7. Fetches candidates with polite delays and retry/backoff.
8. Verifies the returned content before creating a row.
9. Extracts metadata from HTML or PDF content.
10. Writes the exact 22-column dataset plus structured run log and summary.

No login, credentials, CAPTCHA bypass, WAF bypass, or other access-control circumvention is used.

## 2. Discovery and extraction

Discovery is deliberately varied because payer websites publish documents differently.

The crawler can discover documents through:

- HTML links/anchors
- Declared sitemap URLs
- URL patterns/signals

The `discovery_path` records the route used to find a document.

The `extraction_method` records the actual techniques used to obtain or verify information. Examples include:

```text
sitemap|url_pattern|pdf_metadata|pdf_text
html_anchor|url_pattern|pdf_text
html_title|html_text|labeled_text
```

The implementation does not claim to use a technique that it did not actually use.

## 3. Metadata rules

Document type is classified from strong URL/title/text signals, including:

- `medical_policy`
- `pa_list`
- `drug_list`

Effective and updated dates are extracted conservatively. A date is not invented from a plan year or a filename alone. An optional date remains empty when the document does not explicitly provide the required label.

Policy numbers are taken from labelled policy/document-number text when available.

State/region and line of business are extracted from explicit or strong URL/title signals rather than assuming values from the payer name.

`content_hash_sha256` is calculated from the raw downloaded bytes.

## 4. Resumability and output files

The assessment requires the crawler to be resumable. This implementation uses `scraper_state.json` rather than SQLite.

The state file records the current run and completed payers. If a run stops, running the scraper again without `--fresh` continues the unfinished run.

Every new run receives a UTC timestamp and creates a separate output set:

```text
PayerPolicyDocumentDiscovery_YYYYMMDD_HHMMSS.csv
PayerPolicyDocumentDiscovery_YYYYMMDD_HHMMSS.xlsx
PayerPolicyDocumentDiscovery_YYYYMMDD_HHMMSS_run_log.jsonl
PayerPolicyDocumentDiscovery_YYYYMMDD_HHMMSS_summary.json
```

This prevents one run from overwriting another.

## 5. Logging and observability

The run log is JSON Lines: one JSON object per event.

Important events include:

- `run.start`
- `payer.start`
- `robots.loaded`
- `robots.blocked`
- `fetch.failed`
- `doc.found`
- `payer.done`
- `run.done`

The terminal also prints short human-readable progress messages.

The summary contains per-payer:

- attempted
- found
- failed
- skipped

and overall totals, run status, timestamps, and generated file names.

This is intended to make it clear whether a payer produced no documents, was blocked, or experienced fetch failures.

## 6. Ethical and operational constraints

Only publicly accessible content is considered.

The crawler:

- respects robots.txt;
- applies a minimum polite request delay;
- respects declared crawl delays;
- retries temporary failures with backoff;
- keeps discovery bounded;
- does not use credentials;
- does not bypass CAPTCHA/WAF/bot protection.

If access is blocked, the event is recorded rather than bypassed.

## 7. Known limitations / coverage gaps

This is a bounded crawler rather than an unrestricted search engine.

Potential gaps include:

- Documents exposed only through JavaScript applications may not be discovered.
- Documents behind authentication are intentionally excluded.
- Robots-disallowed discovery endpoints are not crawled.
- A payer may publish documents in a directory or URL pattern that is not reachable from the bounded discovery frontier.
- Metadata that is not explicitly published may remain empty.
- Some sites may return WAF/bot-protection responses, which are recorded as failures or blocks rather than bypassed.
- URL-level discovery does not guarantee that every historical or regional document is found.

A lower document count therefore does not by itself mean that the payer publishes fewer documents; the run log and summary should be reviewed for blocks, failures, and coverage gaps.

## 8. What I would improve with more time

With additional development time, I would:

- add more payer-specific discovery strategies where the public site structure warrants them;
- support more document formats and metadata sources;
- improve duplicate/near-duplicate detection while keeping the exact `document_url` rule;
- add stronger automated tests for robots rules, date extraction, document classification, and discovery paths;
- improve handling of JavaScript-rendered public indexes where permitted;
- record more detailed timing and coverage statistics;
- add a small validation report for schema, duplicate URLs, hashes, dates, and confidence scores.

The current implementation intentionally favors a transparent, bounded, reproducible crawler over aggressive discovery or bypass techniques.
