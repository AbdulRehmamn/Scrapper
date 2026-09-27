# ScrapeFlow Pro v2.0

ScrapeFlow Pro is a high-performance local web-scraping dashboard built with Python
and vanilla JavaScript. It downloads public webpages directly and does not require a paid
data API, API key, account, or database.

## What it extracts

- Repeated directory/listing cards
- Schema.org / JSON-LD and Microdata (`itemscope`/`itemprop`) entities
- OpenGraph and Twitter Card social metadata
- HTML definition lists (`<dl>`, `<dt>`, `<dd>`) and HTML tables
- Email addresses and normalized phone numbers with US State & Timezone mapping
- Social profile links across 30+ major networks (LinkedIn, X, Facebook, Instagram, WhatsApp, Discord, GitHub, Reddit, Threads, etc.)
- Page titles, descriptions, headings, status codes and source URLs
- Exact elements selected with a custom CSS selector

Exports are available as formatted Excel workbooks, UTF-8 CSV, and complete
JSON files.

## Version 2.0 Supercharged Features

- **Concurrent Async Crawling**: Up to 3 parallel requests using semaphore controls.
- **Rotating User-Agent Pool**: Anti-detection rotation of realistic modern browser headers.
- **Auto-Retry with Exponential Backoff**: Automatic recovery on rate limits (429) and server errors (500, 502, 503, 504).
- **Anti-Bot Stealth Engine**: Removes headless automation signatures (`navigator.webdriver`, webdriver prototypes, touchpoints).
- **Content Hash Deduplication**: Avoids scraping or storing duplicate content across URLs.
- **Proxy Support**: Connect via HTTP or SOCKS5 proxies directly from the dashboard.
- **Smart Pagination Detection**: Detects `<link rel="next">` and query offsets automatically.

## Start on Windows

1. Extract the entire ZIP into a normal folder.
2. Double-click `start_scraper.bat`.
3. On the first run, allow time for Python packages and Chromium to install.
4. The dashboard opens automatically at `http://127.0.0.1:8765`.
5. Keep the black command window open while using the scraper.

Python 3.10 or newer is required. Direct HTTP mode works even if the optional
Chromium download fails.

## Start on macOS or Linux

```bash
chmod +x start_scraper.sh install_browser.sh
./start_scraper.sh
```

The launch script creates an isolated `.venv`, installs the requirements, and
checks the Chromium browser package. On some Linux distributions Playwright may
also request system browser libraries.

## Extraction modes

- **Smart records:** Detects structured data, repeated listings, contacts and
  table rows automatically.
- **All contacts:** Collects emails, phone numbers, social links and likely
  addresses with minimal processing.
- **Email + phone numbers:** Keeps names, emails and phone numbers together
  when they appear in the same listing, and excludes social/address-only rows.
- **Tables:** Converts HTML tables into rows and exportable records.
- **CSS selector:** Creates one record per element matching a selector such as
  `.listing-card`, `article.company`, or `table tbody tr`.

Use **Render JavaScript** when a page loads its content after the initial HTML.
This launches a local headless Chromium browser; it is slower but works on many
modern websites.

## Crawl controls

- Maximum 5,000 pages per job
- Up to 5,000 crawl levels for long next-page pagination chains
- Same-domain links only
- Configurable 200–5,000 ms request delay
- Optional URL include and exclude text filters
- `robots.txt` respected by default
- 10 MB response limit per page
- Local, private and link-local network destinations blocked
- Up to 250,000 extracted records, 250,000 contacts and 250,000 table rows
- Completed jobs remain available locally for 24 hours
- At most two large scrape jobs can run concurrently

For 1,000–5,000-page jobs, use an include filter when possible, keep the
computer awake, and leave the command window open. Even at the minimum 200 ms
delay, 5,000 requests require at least about 17 minutes plus website response
time. Responsible delays may make the job take substantially longer.

## Accuracy and limitations

No generic scraper can guarantee perfect results on every website because page
structures differ and can change. Some websites block automated browsers,
require login, use CAPTCHAs, prohibit scraping in their terms, or expose data
only after user interaction. ScrapeFlow Pro does not bypass those controls.

For the cleanest data, use Custom CSS Selector mode on a consistent directory
page and review the exported workbook before using the records.

## Responsible use

Only collect data you are authorized to access. Follow the website's terms,
copyright requirements, privacy rules, and applicable laws. Do not use this
tool to collect protected personal data, overload a website, bypass access
controls, or send unsolicited messages.

## Run tests

After the first launch:

Windows:

```text
.venv\Scripts\python.exe -m unittest discover -s tests -v
```

macOS/Linux:

```bash
.venv/bin/python -m unittest discover -s tests -v
```

## Local API

Interactive API documentation is available at `http://127.0.0.1:8765/docs`.
The dashboard uses these endpoints:

- `POST /api/scrape` — start a job
- `GET /api/jobs/{job_id}` — retrieve progress and the completion summary
- `GET /api/jobs/{job_id}/data` — paginate records, contacts, tables, pages or errors
- `POST /api/jobs/{job_id}/cancel` — cancel a running job
- `GET /api/jobs/{job_id}/export?format=xlsx` — export data

No third-party extraction API is called. The only outbound requests are direct
requests to the public webpages you choose and the initial package/browser
downloads during setup.
