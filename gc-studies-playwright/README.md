# GC Studies Playwright scraper

Drives the real General Commons Studies UI in Chromium (no GraphQL/API shortcuts),
visits every study listed on the page, and writes an HTML inventory report.

Source page: https://general.datacommons.cancer.gov/#/studies

## What it captures

For each study (expected: **61**):

- PHS accession / study ID (from the study detail page)
- Number of participants
- Number of files
- Study name, program, data types, access (from the listing table)

## Install

Python 3.10+ recommended.

```bash
cd gc-studies-playwright
python3 -m pip install -r requirements.txt
python3 -m playwright install chromium
```

On first run you may also need OS deps:

```bash
python3 -m playwright install --with-deps chromium
```

## Run

```bash
cd gc-studies-playwright
python3 scrape_studies.py
```

Optional env vars:

| Variable | Default | Meaning |
| --- | --- | --- |
| `GC_STUDIES_URL` | `https://general.datacommons.cancer.gov/#/studies` | Studies listing URL |
| `PLAYWRIGHT_HEADLESS` | `1` | Set `0` to show the browser |
| `GC_OUTPUT_DIR` | `./output` | Where JSON + HTML are written |

Outputs (under `output/` by default):

- `gc-studies-playwright-report.html` — browsable report
- `studies.json` — raw scraped records

## How the UI is driven

1. Open the Studies page
2. Dismiss the government warning dialog (`Continue`)
3. Set **Rows per page** to `100` so all studies appear in the table
4. Read every listing row (PHS link, participants, files, metadata)
5. Click/navigate to each study’s `#/study/<id>` detail page and re-read PHS, participants, and files from that page
6. Emit the HTML report from the scraped UI data
