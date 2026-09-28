"""
Payer Policy Document Discovery Crawler

Author: Navin V P

Purpose:
    Discover public payer medical policies, prior authorization,
    formulary, and drug-list documents.

Technology:
    Python, Requests, PyPDF, OpenPyXL, Regex, JSON/JSONL, CSV.

Features:
    Robots-aware crawling, sitemap discovery, metadata extraction,
    document classification, SHA-256 hashing, resumability,
    structured logging, validation, and timestamped outputs.

Principles:
    Public pages only. No authentication or bypassing CAPTCHA,
    WAF, bot protection, or access controls. Uses polite,
    bounded crawling with retries and backoff.
"""

import csv
import hashlib
import json
import os
import re
import time
from email.utils import parsedate_to_datetime
from datetime import datetime, timezone
from urllib.parse import urljoin, urlparse, urldefrag

import requests
from openpyxl import Workbook
from pypdf import PdfReader

BASE = os.path.dirname(os.path.abspath(__file__))
SEED_FILE = os.path.join(BASE, "data", "payer_seed_list.csv")
PROJECT_NAME = "PayerPolicyDocumentDiscovery"
RUN_STATE_FILE = os.path.join(BASE, "scraper_state.json")

# Example summary structure:
# {
#   "run": {"started_at_utc": "...", "finished_at_utc": "...", "status": "completed"},
#   "payers": [{"payer_name": "UHC", "attempted": 10, "found": 5, "failed": 1, "skipped": 0}],
#   "totals": {"attempted": 10, "found": 5, "failed": 1, "skipped": 0, "rows_written": 5}
# }
COLUMNS = [
    "payer_name", "payer_alias", "state_or_region", "line_of_business",
    "document_title", "document_type", "document_url", "source_page_url",
    "discovery_path", "file_type", "policy_number", "effective_date",
    "last_updated_date", "http_status", "content_hash_sha256",
    "file_size_bytes", "requires_auth", "render_mode", "extraction_method",
    "confidence_score", "scrape_timestamp_utc", "notes"
]

HEADERS = {
    "User-Agent": "PublicPolicyCrawler/1.0 (assessment; contact not provided)"
}
DOC_RE = re.compile(r"\.(?:pdf|docx?|xlsx?|csv)(?:$|[?#])", re.I)
TAG_RE = re.compile(r"<[^>]+>")
DATE_RE = re.compile(
    r"\b(?:20\d{2}-\d{2}-\d{2}|"
    r"(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|"
    r"Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|"
    r"Dec(?:ember)?)\s+\d{1,2},?\s+20\d{2})\b", re.I
)
POLICY_RE = re.compile(
    r"(?:policy\s*(?:number|#|no\.?)|document\s*(?:number|#|no\.?))"
    r"\s*[:#-]?\s*([A-Za-z0-9./_-]{2,40})", re.I
)

TYPE_RULES = [
    ("drug_list", r"\b(?:pdl|preferred drug list|prescription drug list|drug list|formulary)\b"),
    ("pa_list", r"\b(?:prior authorization|prior auth|precertification|precert list)\b"),
    ("medical_policy", r"\b(?:medical policy|clinical policy|coverage policy|medical necessity|clinical guideline|coverage guideline)\b"),
]

LOB_RULES = [
    ("Medicare Advantage", r"\bmedicare advantage\b"),
    ("Medicaid", r"\bmedicaid\b"),
    ("Community Plan", r"\bcommunity plan\b"),
    ("Marketplace", r"\bmarketplace\b"),
    ("Exchange", r"\bexchange\b"),
    ("Commercial", r"\bcommercial\b"),
    ("Medicare", r"\bmedicare\b"),
]

session = requests.Session()
session.headers.update(HEADERS)


def now():
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def run_stamp():
    return datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")


def make_run_files(stamp):
    prefix = f"{PROJECT_NAME}_{stamp}"
    return {
        "csv": os.path.join(BASE, f"{prefix}.csv"),
        "xlsx": os.path.join(BASE, f"{prefix}.xlsx"),
        "log": os.path.join(BASE, f"{prefix}_run_log.jsonl"),
        "summary": os.path.join(BASE, f"{prefix}_summary.json"),
    }


def log(event, payer, **data):
    """Write one clean JSON object per line for easy debugging/review."""
    row = {"timestamp_utc": now(), "event": event, "payer": payer, **data}
    with open(RUN_FILES["log"], "a", encoding="utf-8") as f:
        f.write(json.dumps(row, ensure_ascii=False) + "\n")


def print_log(event, payer, **data):
    """Show a short human-readable message while the scraper runs."""
    message = f"[{payer}] {event}"
    if data:
        details = " | ".join(f"{k}={v}" for k, v in data.items())
        message += f" | {details}"
    print(message)
    log(event, payer, **data)


def clean_url(url):
    return urldefrag(url)[0]


def same_site(url, seed):
    a, b = urlparse(url), urlparse(seed)
    return a.scheme in ("http", "https") and (
        a.netloc == b.netloc or a.netloc.endswith("." + b.netloc)
    )


def text_from_html(html):
    return re.sub(r"\s+", " ", TAG_RE.sub(" ", html)).strip()


def title_from_html(html):
    m = re.search(r"<title[^>]*>(.*?)</title>", html, re.I | re.S)
    return re.sub(r"\s+", " ", TAG_RE.sub(" ", m.group(1))).strip() if m else ""


def links_from_html(html, base_url):
    found = []
    for href, anchor in re.findall(
        r'<a\b[^>]*href=["\']([^"\']+)["\'][^>]*>(.*?)</a>',
        html, re.I | re.S
    ):
        url = normalize_url(urljoin(base_url, href))
        if url.startswith(("http://", "https://")):
            found.append((url, text_from_html(anchor)))
    return found


def classify(text):
    low = text.lower()
    for kind, pattern in TYPE_RULES:
        if re.search(pattern, low):
            return kind
    return ""


def url_type(url):
    path = urlparse(url).path.lower()
    if path.endswith(".pdf"):
        return "pdf"
    if path.endswith(".doc"):
        return "doc"
    if path.endswith(".docx"):
        return "docx"
    if path.endswith(".xls"):
        return "xls"
    if path.endswith(".xlsx"):
        return "xlsx"
    if path.endswith(".csv"):
        return "csv"
    return "html"


def normalize_url(url):
    url = clean_url(url)
    parsed = urlparse(url)
    path = parsed.path.rstrip("/") or "/"
    return parsed._replace(path=path).geturl()


def server_last_modified(response):
    value = response.headers.get("Last-Modified")
    if not value:
        return ""
    try:
        return parsedate_to_datetime(value).date().isoformat()
    except (TypeError, ValueError, OverflowError):
        return ""


def line_of_business(text):
    for value, pattern in LOB_RULES:
        if re.search(pattern, text, re.I):
            return value
    return ""


def state_or_region(text, url, title=""):
    states = {
        "Alabama":"AL","Alaska":"AK","Arizona":"AZ","Arkansas":"AR","California":"CA",
        "Colorado":"CO","Connecticut":"CT","Delaware":"DE","Florida":"FL","Georgia":"GA",
        "Illinois":"IL","Indiana":"IN","Iowa":"IA","Kansas":"KS","Kentucky":"KY",
        "Louisiana":"LA","Maine":"ME","Maryland":"MD","Massachusetts":"MA","Michigan":"MI",
        "Minnesota":"MN","Mississippi":"MS","Missouri":"MO","Montana":"MT","Nebraska":"NE",
        "Nevada":"NV","New Hampshire":"NH","New Jersey":"NJ","New Mexico":"NM","New York":"NY",
        "North Carolina":"NC","North Dakota":"ND","Ohio":"OH","Oklahoma":"OK","Oregon":"OR",
        "Pennsylvania":"PA","Rhode Island":"RI","South Carolina":"SC","South Dakota":"SD",
        "Tennessee":"TN","Texas":"TX","Utah":"UT","Vermont":"VT","Virginia":"VA",
        "Washington":"WA","West Virginia":"WV","Wisconsin":"WI","Wyoming":"WY"
    }
    explicit = re.search(r"\b(?:state|region)\s*[:=-]\s*([A-Z]{2})\b", text, re.I)
    if explicit:
        return explicit.group(1).upper(), "labeled_text"
    for source, value in (("url_pattern", url), ("html_title", title)):
        for name, code in states.items():
            if re.search(rf"\b{re.escape(name)}\b", value, re.I) or re.search(rf"(?:^|[-_/]){code}(?:[-_/]|$)", value, re.I):
                return code, source
    return "", ""


def line_of_business_with_method(text, url, title):
    for value, pattern in LOB_RULES:
        if re.search(pattern, title, re.I):
            return value, "html_title"
    for value, pattern in LOB_RULES:
        if re.search(pattern, url, re.I):
            return value, "url_pattern"
    labeled = re.search(r"(?:line of business|lob|product|plan)\s*[:=-]\s*([^|,;\n]{2,50})", text, re.I)
    if labeled:
        value = labeled.group(1).strip()
        for known, pattern in LOB_RULES:
            if re.search(pattern, value, re.I):
                return known, "labeled_text"
    return "", ""

def labeled_date(text, labels):
    for label in labels:
        m = re.search(
            rf"{label}\s*[:\-]?\s*"
            r"((?:20\d{2}-\d{2}-\d{2})|"
            r"(?:[A-Za-z]+\s+\d{1,2},?\s+20\d{2}))",
            text, re.I
        )
        if m:
            return m.group(1)
    return ""


def policy_number(text):
    m = POLICY_RE.search(text)
    return m.group(1) if m else ""


def parse_pdf(content):
    try:
        reader = PdfReader(__import__("io").BytesIO(content))
        meta = reader.metadata
        title = str(meta.title).strip() if meta and meta.title else ""
        text = " ".join((p.extract_text() or "") for p in reader.pages[:8])
        return title, text, bool(title)
    except Exception:
        return "", "", False

def get_robots(seed):
    url = urljoin(seed, "/robots.txt")
    try:
        r = session.get(url, timeout=15)
        if r.status_code != 200:
            return None, [], 1.0
        groups, sitemaps = {}, []
        agent = None
        delay = 1.0
        for line in r.text.splitlines():
            line = line.split("#", 1)[0].strip()
            if not line:
                continue
            key, _, value = line.partition(":")
            key, value = key.strip().lower(), value.strip()
            if key == "user-agent":
                agent = value
                groups.setdefault(agent, {"allow": [], "deny": []})
            elif agent in ("*", HEADERS["User-Agent"]) and key in ("allow", "disallow"):
                groups.setdefault(agent, {"allow": [], "deny": []})["allow" if key == "allow" else "deny"].append(value)
            elif key == "crawl-delay" and agent in ("*", HEADERS["User-Agent"]):
                try:
                    delay = max(1.0, float(value))
                except ValueError:
                    pass
            elif key == "sitemap":
                sitemaps.append(value)
        rules = groups.get(HEADERS["User-Agent"], groups.get("*", {"allow": [], "deny": []}))
        return rules, sitemaps, delay
    except requests.RequestException:
        return None, [], 1.0


def allowed(url, rules):
    if rules is None:
        return False
    path = urlparse(url).path or "/"
    for rule in rules["allow"]:
        if rule and path.startswith(rule):
            return True
    for rule in rules["deny"]:
        if rule and path.startswith(rule):
            return False
    return True


def fetch(url, delay, retries=2):
    time.sleep(delay)
    for attempt in range(retries + 1):
        try:
            r = session.get(url, timeout=20, allow_redirects=True)
            if r.status_code in (429, 500, 502, 503, 504) and attempt < retries:
                time.sleep(2 ** attempt)
                continue
            return r
        except requests.RequestException:
            if attempt == retries:
                return None
            time.sleep(2 ** attempt)
    return None


def sitemap_urls(text):
    return re.findall(r"<loc>\s*(.*?)\s*</loc>", text, re.I | re.S)


def discover(payer, seed, rules, delay):
    pages, docs = [], []
    queue = [seed]
    seen = set()

    for _ in range(30):
        if not queue:
            break
        url = clean_url(queue.pop(0))
        if url in seen or not same_site(url, seed):
            continue
        seen.add(url)
        if not allowed(url, rules):
            log("robots.blocked", payer, url=url)
            continue

        r = fetch(url, delay)
        if not r:
            log("fetch.failed", payer, url=url, reason="request_error")
            continue

        log("fetch.done", payer, url=url, status=r.status_code)
        if r.status_code != 200:
            log("fetch.failed", payer, url=url, status=r.status_code)
            continue

        pages.append(url)
        for link, anchor in links_from_html(r.text, r.url):
            if DOC_RE.search(link):
                docs.append((link, url, f"seed -> page -> anchor -> document ({anchor[:80]})"))
            elif same_site(link, seed) and link not in seen and len(queue) < 60:
                queue.append(link)

    return pages, docs


def discover_sitemaps(payer, seed, rules, delay):
    robots, maps, _ = get_robots(seed)
    urls = []
    for sm in maps[:10]:
        if not allowed(sm, rules):
            log("robots.blocked", payer, url=sm)
            continue
        r = fetch(sm, delay)
        if not r or r.status_code != 200:
            continue
        for u in sitemap_urls(r.text)[:1000]:
            urls.append(clean_url(u))
    return list(dict.fromkeys(urls))


def document_row(payer, alias, url, source, path, response):
    """Create one verified document row from a public response."""
    content = response.content
    final_url = normalize_url(response.url)
    content_type = response.headers.get("content-type", "").lower()
    is_pdf = "pdf" in content_type or final_url.lower().endswith(".pdf")

    methods = []
    if "sitemap" in path:
        methods.append("sitemap")
    if "anchor" in path:
        methods.append("html_anchor")
    if re.search(r"policy|guideline|clinical|precert|prior.?auth|pdl|drug", final_url, re.I):
        methods.append("url_pattern")

    if is_pdf:
        pdf_result = parse_pdf(content)
        title, body = pdf_result[0], pdf_result[1]
        if len(pdf_result) > 2 and pdf_result[2]:
            methods.append(str(pdf_result[2]))
        methods.extend(["pdf_metadata", "pdf_text"])
    else:
        title = title_from_html(response.text)
        body = text_from_html(response.text)[:15000]
        methods.extend(["html_title", "html_text"])

    kind = classify(" ".join([final_url, title])) or classify(body)
    if not kind:
        return None

    effective = labeled_date(body, ["effective date", "effective"])
    updated = labeled_date(body, ["last updated", "updated date", "revision date", "revised"])
    policy = policy_number(body)

    if effective or updated or policy:
        methods.append("labeled_text")

    if title:
        methods.append("title")
    else:
        title = os.path.basename(urlparse(final_url).path) or final_url
        methods.append("filename")

    methods = list(dict.fromkeys(methods))
    strong_signal = classify(" ".join([final_url, title]))
    confidence = 0.95 if strong_signal else 0.80

    region_value, region_method = state_or_region(body[:8000], final_url, title)
    if region_method:
        methods.append(region_method)

    return {
        "payer_name": payer,
        "payer_alias": alias,
        "state_or_region": region_value,
        "line_of_business": line_of_business(" ".join([final_url, title, body[:3000]])),
        "document_title": title,
        "document_type": kind,
        "document_url": final_url,
        "source_page_url": normalize_url(source),
        "discovery_path": path,
        "file_type": "pdf" if is_pdf else url_type(final_url),
        "policy_number": policy,
        "effective_date": effective,
        "last_updated_date": updated or server_last_modified(response),
        "http_status": response.status_code,
        "content_hash_sha256": hashlib.sha256(content).hexdigest(),
        "file_size_bytes": len(content),
        "requires_auth": "N",
        "render_mode": "static",
        "extraction_method": "|".join(methods),
        "confidence_score": f"{confidence:.2f}",
        "scrape_timestamp_utc": now(),
        "notes": ""
    }


def validate_rows(rows):
    """Run simple checks before final output is written."""
    errors = []

    if any(list(row.keys()) != COLUMNS for row in rows):
        errors.append("schema_or_column_order")

    urls = [row.get("document_url", "") for row in rows]
    if len(urls) != len(set(urls)):
        errors.append("duplicate_document_url")

    for row in rows:
        if not re.fullmatch(r"[0-9a-f]{64}", row.get("content_hash_sha256", "")):
            errors.append("invalid_sha256")
        try:
            score = float(row.get("confidence_score", ""))
            if not 0 <= score <= 1:
                errors.append("invalid_confidence")
        except ValueError:
            errors.append("invalid_confidence")
        if row.get("requires_auth") not in ("Y", "N"):
            errors.append("invalid_requires_auth")

    return list(dict.fromkeys(errors))



def load_payers():
    with open(SEED_FILE, newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def load_existing(csv_file):
    if not os.path.exists(csv_file):
        return {}
    try:
        with open(csv_file, newline="", encoding="utf-8") as f:
            return {
                r["document_url"]: r
                for r in csv.DictReader(f)
                if r.get("document_url")
            }
    except (OSError, csv.Error, KeyError):
        return {}


def atomic_replace(temp_file, target_file):
    """Replace a completed temporary file without exposing a partial file."""
    os.replace(temp_file, target_file)


def save_csv(rows, csv_file):
    csv_tmp = csv_file + ".tmp"
    with open(csv_tmp, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=COLUMNS)
        writer.writeheader()
        writer.writerows(rows)
        f.flush()
        os.fsync(f.fileno())
    atomic_replace(csv_tmp, csv_file)


def excel_value(value):
    """Convert any unexpected structured value into an Excel-safe scalar."""
    if value is None:
        return ""
    if isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, (tuple, list, set)):
        return " | ".join(str(excel_value(v)) for v in value if v is not None)
    if isinstance(value, dict):
        return " | ".join(
            f"{key}={excel_value(val)}" for key, val in value.items()
        )
    return str(value)



def save(rows, csv_file, xlsx_file):
    """Write CSV/XLSX atomically so an interruption does not corrupt the last checkpoint."""
    xlsx_tmp = xlsx_file + ".tmp"

    save_csv(rows, csv_file)

    wb = Workbook()
    ws = wb.active
    ws.title = "documents"
    ws.append(COLUMNS)
    for row in rows:
        ws.append([excel_value(row.get(column, "")) for column in COLUMNS])
    wb.save(xlsx_tmp)
    atomic_replace(xlsx_tmp, xlsx_file)


def load_run_state():
    if not os.path.exists(RUN_STATE_FILE):
        return None
    try:
        with open(RUN_STATE_FILE, encoding="utf-8") as f:
            state = json.load(f)
        if state.get("status") in ("running", "interrupted"):
            return state
    except (OSError, json.JSONDecodeError):
        pass
    return None


def save_run_state(state):
    temp = RUN_STATE_FILE + ".tmp"
    with open(temp, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2)
        f.flush()
        os.fsync(f.fileno())
    atomic_replace(temp, RUN_STATE_FILE)


def load_summary(summary_file):
    if not os.path.exists(summary_file):
        return None
    try:
        with open(summary_file, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, json.JSONDecodeError):
        return None


def save_summary(summary, summary_file):
    temp = summary_file + ".tmp"
    with open(temp, "w", encoding="utf-8") as f:
        json.dump(summary, f, indent=2, ensure_ascii=False)
        f.flush()
        os.fsync(f.fileno())
    atomic_replace(temp, summary_file)


def checkpoint(rows, summary, stamp, completed_payers, current_payer=None,
               state_status="running", write_xlsx=False):
    """Persist collected data safely at every checkpoint."""
    rows_sorted = list(rows.values()) if isinstance(rows, dict) else list(rows)

    # Guarantee every schema value is a scalar before any file writer sees it.
    rows_sorted = [
        {column: excel_value(row.get(column, "")) for column in COLUMNS}
        for row in rows_sorted
    ]
    rows_sorted.sort(key=lambda x: (x["payer_name"], x["document_url"]))

    # CSV is the primary durable checkpoint.
    save_csv(rows_sorted, RUN_FILES["csv"])

    # XLSX is secondary. Never let an XLSX problem destroy the CSV/state/summary.
    if write_xlsx or not os.path.exists(RUN_FILES["xlsx"]):
        try:
            save(rows_sorted, RUN_FILES["csv"], RUN_FILES["xlsx"])
        except Exception as exc:
            print_log(
                "output.warning",
                "ALL",
                output="xlsx",
                error=f"{type(exc).__name__}: {exc}"
            )

    summary["totals"]["rows_written"] = len(rows_sorted)
    summary["run"]["status"] = state_status
    summary["run"]["last_checkpoint_utc"] = now()
    summary["run"]["files"] = {
        "csv": os.path.basename(RUN_FILES["csv"]),
        "xlsx": os.path.basename(RUN_FILES["xlsx"]),
        "log": os.path.basename(RUN_FILES["log"]),
        "summary": os.path.basename(RUN_FILES["summary"])
    }
    if current_payer:
        summary["run"]["current_payer"] = current_payer

    save_summary(summary, RUN_FILES["summary"])

    save_run_state({
        "run_stamp": stamp,
        "status": state_status,
        "completed_payers": sorted(completed_payers),
        "current_payer": current_payer or "",
        "rows_written": len(rows_sorted),
        "last_checkpoint_utc": now()
    })




RUN_FILES = {}
RUNTIME = {"stop_requested": False}



def main():
    import argparse
    import signal

    p = argparse.ArgumentParser()
    p.add_argument("--payer", help="Run one payer only")
    p.add_argument("--fresh", action="store_true",
                   help="Start a new run instead of resuming an interrupted run")
    args = p.parse_args()

    previous = None if args.fresh else load_run_state()

    if previous:
        stamp = previous["run_stamp"]
        RUN_FILES.update(make_run_files(stamp))
        existing = load_existing(RUN_FILES["csv"])
        completed_payers = set(previous.get("completed_payers", []))
        print(f"Resuming run: {stamp}")
    else:
        stamp = run_stamp()
        RUN_FILES.update(make_run_files(stamp))
        existing = {}
        completed_payers = set()

    payers = load_payers()
    if args.payer:
        payers = [p for p in payers if p["payer_name"].lower() == args.payer.lower()]

    # Resume the existing summary so interrupted runs retain the statistics
    # already collected before the interruption.
    summary = load_summary(RUN_FILES["summary"]) if previous else None
    if not summary:
        summary = {
            "run": {
                "started_at_utc": now(),
                "status": "running",
                "payer_count": len(payers)
            },
            "payers": [],
            "totals": {
                "attempted": 0,
                "found": 0,
                "failed": 0,
                "skipped": 0,
                "rows_written": len(existing)
            }
        }

    # Keep the current command's payer count accurate.
    summary["run"]["payer_count"] = len(payers)

    def stats_for(payer_name):
        for item in summary["payers"]:
            if item.get("payer_name") == payer_name:
                return item
        item = {
            "payer_name": payer_name,
            "attempted": 0,
            "found": 0,
            "failed": 0,
            "skipped": 0,
            "pages_discovered": 0,
            "candidate_documents": 0,
            "robots_blocked": 0,
            "elapsed_seconds": 0.0
        }
        summary["payers"].append(item)
        return item

    def refresh_totals():
        summary["totals"] = {
            "attempted": sum(x.get("attempted", 0) for x in summary["payers"]),
            "found": sum(x.get("found", 0) for x in summary["payers"]),
            "failed": sum(x.get("failed", 0) for x in summary["payers"]),
            "skipped": sum(x.get("skipped", 0) for x in summary["payers"]),
            "rows_written": len(existing)
        }

    def request_stop(signum, frame):
        RUNTIME["stop_requested"] = True
        raise KeyboardInterrupt(f"Signal {signum} received")

    # Ctrl+C and normal process termination both get a final checkpoint.
    signal.signal(signal.SIGINT, request_stop)
    if hasattr(signal, "SIGTERM"):
        signal.signal(signal.SIGTERM, request_stop)

    # Create CSV/XLSX/summary immediately. This guarantees visible output
    # even if the crawler stops before discovering its first document.
    refresh_totals()
    checkpoint(
        existing, summary, stamp, completed_payers,
        current_payer=previous.get("current_payer") if previous else None,
        state_status="running"
    )
    print_log("run.start", "ALL", payer_count=len(payers))

    current_payer = None

    try:
        for payer in payers:
            if payer["payer_name"] in completed_payers:
                log("payer.resume.skip", payer["payer_name"],
                    reason="already_completed")
                continue

            name = payer["payer_name"]
            alias = payer.get("payer_alias") or name
            seed = payer.get("seed_url") or payer.get("hint_host", "")
            if not seed.startswith(("http://", "https://")):
                seed = "https://" + seed

            current_payer = name
            stats = stats_for(name)
            payer_started = time.monotonic()

            print(f"\n[{name}] starting")
            log("payer.start", name, seed_url=seed)

            rules, _, delay = get_robots(seed)
            if rules is None:
                stats["skipped"] += 1
                print_log("robots.unavailable", name,
                          url=urljoin(seed, "/robots.txt"))
                stats["elapsed_seconds"] = round(time.monotonic() - payer_started, 2)
                refresh_totals()
                checkpoint(existing, summary, stamp, completed_payers,
                           current_payer=name)
                continue

            print_log("robots.loaded", name, crawl_delay=delay)
            pages, docs = discover(name, seed, rules, delay)
            stats["pages_discovered"] = len(pages)

            for u in discover_sitemaps(name, seed, rules, delay):
                if DOC_RE.search(u):
                    docs.append((u, seed, "seed -> robots sitemap -> document"))

            unique_docs = {}
            for doc in docs:
                unique_docs.setdefault(normalize_url(doc[0]), doc)

            stats["candidate_documents"] = len(unique_docs)
            refresh_totals()
            checkpoint(existing, summary, stamp, completed_payers,
                       current_payer=name)

            for url, source, path in unique_docs.values():
                if url in existing:
                    stats["skipped"] += 1
                    refresh_totals()
                    checkpoint(existing, summary, stamp, completed_payers,
                               current_payer=name)
                    continue

                if not allowed(url, rules):
                    stats["skipped"] += 1
                    stats["robots_blocked"] += 1
                    log("robots.blocked", name, url=url)
                    refresh_totals()
                    checkpoint(existing, summary, stamp, completed_payers,
                               current_payer=name)
                    continue

                stats["attempted"] += 1
                response = fetch(url, delay)

                if not response:
                    stats["failed"] += 1
                    log("fetch.failed", name, url=url, reason="request_error")
                    refresh_totals()
                    checkpoint(existing, summary, stamp, completed_payers,
                               current_payer=name)
                    continue

                if response.status_code != 200:
                    stats["failed"] += 1
                    log("fetch.failed", name, url=url,
                        status=response.status_code)
                    refresh_totals()
                    checkpoint(existing, summary, stamp, completed_payers,
                               current_payer=name)
                    continue

                row = document_row(name, alias, url, source, path, response)
                if not row:
                    stats["failed"] += 1
                    log("document.rejected", name, url=url,
                        reason="not_a_target_document")
                    refresh_totals()
                    checkpoint(existing, summary, stamp, completed_payers,
                               current_payer=name)
                    continue

                existing[row["document_url"]] = row
                stats["found"] += 1
                print_log("doc.found", name,
                          url=row["document_url"],
                          document_type=row["document_type"])
                print(f"  found: {row['document_title'][:80]}")

                # CRITICAL: persist immediately after every successful document.
                refresh_totals()
                checkpoint(existing, summary, stamp, completed_payers,
                           current_payer=name, write_xlsx=True)

            stats["elapsed_seconds"] = round(time.monotonic() - payer_started, 2)
            completed_payers.add(name)
            log("payer.done", name, **stats)

            refresh_totals()
            checkpoint(existing, summary, stamp, completed_payers,
                       current_payer=None, write_xlsx=True)

            print(
                f"[{name}] completed | "
                f"attempted={stats['attempted']} | found={stats['found']} | "
                f"failed={stats['failed']} | skipped={stats['skipped']}"
            )

        rows = list(existing.values())
        rows.sort(key=lambda x: (x["payer_name"], x["document_url"]))

        validation_errors = validate_rows(rows)
        summary["validation"] = {
            "status": "PASS" if not validation_errors else "FAIL",
            "errors": validation_errors
        }
        summary["totals"]["rows_written"] = len(rows)
        summary["run"]["finished_at_utc"] = now()
        summary["run"]["status"] = "completed"
        summary["run"].pop("current_payer", None)

        save(rows, RUN_FILES["csv"], RUN_FILES["xlsx"])
        save_summary(summary, RUN_FILES["summary"])
        save_run_state({
            "run_stamp": stamp,
            "status": "completed",
            "completed_payers": sorted(completed_payers),
            "rows_written": len(rows),
            "last_checkpoint_utc": now()
        })

        print_log("validation", "ALL",
                   status=summary["validation"]["status"],
                   errors=len(summary["validation"]["errors"]))
        print_log("run.done", "ALL", rows_written=len(rows))
        print(f"\nDone. {len(rows)} rows written.")

    except (KeyboardInterrupt, SystemExit) as exc:
        # Ctrl+C/SIGTERM: preserve everything collected so far.
        refresh_totals()
        summary["run"]["status"] = "interrupted"
        summary["run"]["interrupted_at_utc"] = now()
        summary["run"]["current_payer"] = current_payer or ""

        checkpoint(existing, summary, stamp, completed_payers,
                   current_payer=current_payer,
                   state_status="interrupted", write_xlsx=True)

        log("run.interrupted", "ALL",
            rows_written=len(existing),
            current_payer=current_payer or "",
            reason=type(exc).__name__)

        print(
            f"\nRun interrupted safely. {len(existing)} rows are saved.\n"
            f"Resume with: python scraper.py"
        )

    except Exception as exc:
        # Unexpected application errors: preserve the partial dataset too.
        refresh_totals()
        summary["run"]["status"] = "failed"
        summary["run"]["failed_at_utc"] = now()
        summary["run"]["current_payer"] = current_payer or ""
        summary["run"]["error"] = f"{type(exc).__name__}: {exc}"

        checkpoint(existing, summary, stamp, completed_payers,
                   current_payer=current_payer,
                   state_status="interrupted", write_xlsx=True)

        log("run.failed", "ALL",
            rows_written=len(existing),
            current_payer=current_payer or "",
            error=f"{type(exc).__name__}: {exc}")

        print(
            f"\nRun stopped because of an unexpected error. "
            f"{len(existing)} rows were saved.\n"
            f"Resume with: python scraper.py"
        )
        raise


if __name__ == "__main__":
    main()
