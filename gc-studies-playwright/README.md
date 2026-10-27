# GC Studies Playwright scraper

Compact Chromium scrape of the General Commons Studies page (headed by default).

- Reads **statbar** Studies / Participants / Samples / Files
- Scrapes each Study’s participants + files from the table
- Compares to package-root `GC Studypage_Basecount.xlsx` (fallback: `data/`)
- Writes `output/gc-studies-playwright-report.html`

## Install

```bash
cd gc-studies-playwright
python3 -m pip install -r requirements.txt
python3 -m playwright install chromium
```

## Run

Default is **production** + **headed** (browser visible):

```bash
python3 scrape_studies.py
```

### Headed / headless

| Mode | How |
| --- | --- |
| Headed (default) | `python3 scrape_studies.py` or `--headed` |
| Headless | `HEADLESS=1 python3 scrape_studies.py` or `--headless` |

```bash
# see the browser
python3 scrape_studies.py --headed

# no browser window (CI / remote)
HEADLESS=1 python3 scrape_studies.py
python3 scrape_studies.py --headless
```

**PowerShell:**

```powershell
# headed (default) — watch the browser
python scrape_studies.py
python scrape_studies.py --headed

# headless
$env:HEADLESS = "1"; python scrape_studies.py
python scrape_studies.py --headless

# QA2 + headed
$env:GC_STUDIES_URL = "https://general-qa2.datacommons.cancer.gov/#/studies"
python scrape_studies.py --headed

# QA2 + headless
$env:HEADLESS = "1"
python scrape_studies.py --url "https://general-qa2.datacommons.cancer.gov/#/studies"
```

Also accepts `PLAYWRIGHT_HEADLESS=1` (same as `HEADLESS`).

### Switch environment (no code edits)

| Env | Studies URL |
| --- | --- |
| Prod (default) | `https://general.datacommons.cancer.gov/#/studies` |
| QA2 | `https://general-qa2.datacommons.cancer.gov/#/studies` |
| QA | `https://general-qa.datacommons.cancer.gov/#/studies` |
| Stage | `https://general-stage.datacommons.cancer.gov/#/studies` |

```bash
GC_STUDIES_URL='https://general-qa2.datacommons.cancer.gov/#/studies' python3 scrape_studies.py
BASE_URL='https://general-qa2.datacommons.cancer.gov/#/studies' python3 scrape_studies.py
python3 scrape_studies.py --url 'https://general-qa2.datacommons.cancer.gov/#/studies'
python3 scrape_studies.py 'https://general-stage.datacommons.cancer.gov/#/studies'
```

Optional: `GC_BASECOUNT_XLSX`, `GC_OUTPUT_DIR`

## Sorting verification

```bash
# default headless
python3 verify_sorting.py
HEADLESS=0 python3 verify_sorting.py --headed
GC_STUDIES_URL='https://general-qa2.datacommons.cancer.gov/#/studies' python3 verify_sorting.py
```

Writes `output/gc-studies-sorting-report.html` and `output/sorting-results.json`.
Checks default sort (Study Name ASC) and asc/desc order for every column.
