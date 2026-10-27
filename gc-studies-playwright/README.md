# GC Studies Playwright scraper

Drives the real General Commons Studies UI in Chromium (no GraphQL/API shortcuts),
reads the page **statbar** totals (Studies / Participants / Samples / Files),
visits every study listed on the page, compares PHS / participants / file counts to
the baseline Excel workbook, and writes an HTML pass/fail report.

Source page: https://general.datacommons.cancer.gov/#/studies  

**Baseline Excel (primary):** `GC Studypage_Basecount.xlsx` at the package root  
(same location as local: `…\gc-studies-playwright\GC Studypage_Basecount.xlsx`)  
Fallback: `data/GC_Studypage_Basecount.xlsx`

## What it captures

- **Statbar** (page header widgets `#statsbar_title_*`): Studies, Participants, Samples, Files
- For each study (expected: **61**):
  - PHS accession / study ID (from the study detail page)
  - Number of participants
  - Number of files
  - Study name, program, data types, access (from the listing table)
  - Pass / fail vs the Excel baseline (participants + files)

Primary report totals are the **statbar** values. Table column sums are secondary cross-checks only.

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

## Run (headed — default for local)

The browser window is shown by default:

```bash
cd gc-studies-playwright
python3 scrape_studies.py
```

### Headless override (cloud / CI / no display)

```bash
HEADLESS=1 python3 scrape_studies.py
# or
PLAYWRIGHT_HEADLESS=1 python3 scrape_studies.py
```

Optional env vars:

| Variable | Default | Meaning |
| --- | --- | --- |
| `GC_STUDIES_URL` | `https://general.datacommons.cancer.gov/#/studies` | Studies listing URL |
| `GC_BASECOUNT_XLSX` | auto: package-root `GC Studypage_Basecount.xlsx`, then `data/` | Excel comparison baseline |
| `HEADLESS` / `PLAYWRIGHT_HEADLESS` | unset → **headed** | Set `1` / `true` for headless |
| `GC_OUTPUT_DIR` | `./output` | Where JSON + HTML are written |

Outputs (under `output/` by default):

- `gc-studies-playwright-report.html` — browsable pass/fail report (statbar primary)
- `studies.json` — scraped + comparison records

## Comparison rules

1. Load baseline rows that have a Study / PHS id (version-only continuation rows are ignored).
2. Match UI studies to Excel by normalized PHS / study ID (case-insensitive).
3. Compare **Number of Participants** and **Number of Files** (Excel column may be spelled `Nubmer of Files`).
4. Status per study:
   - **PASS** — both counts match
   - **FAIL** — PHS found in Excel but one or both counts differ
   - **MISSING_IN_BASE** — scraped from UI but not in Excel
   - **MISSING_IN_UI** — present in Excel but not scraped from UI

## How the UI is driven

1. Open the Studies page (headed Chromium by default)
2. Dismiss the government warning dialog (`Continue`)
3. Read the **statbar** totals (Studies / Participants / Samples / Files)
4. Set **Rows per page** to `100` so all studies appear in the table
5. Read every listing row (PHS link, participants, files, metadata)
6. Navigate to each study’s `#/study/<id>` detail page and re-read PHS, participants, and files
7. Compare against package-root `GC Studypage_Basecount.xlsx` (or `data/` fallback) and emit the HTML report
