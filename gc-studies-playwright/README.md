# GC Studies Playwright scraper

Compact headed Chromium scrape of the General Commons Studies page.

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

Default is **production** (headed):

```bash
python3 scrape_studies.py
```

### Switch environment (no code edits)

| Env | Studies URL |
| --- | --- |
| Prod (default) | `https://general.datacommons.cancer.gov/#/studies` |
| QA2 | `https://general-qa2.datacommons.cancer.gov/#/studies` |
| QA | `https://general-qa.datacommons.cancer.gov/#/studies` |
| Stage | `https://general-stage.datacommons.cancer.gov/#/studies` |

```bash
# env var (either name works)
GC_STUDIES_URL='https://general-qa2.datacommons.cancer.gov/#/studies' python3 scrape_studies.py
BASE_URL='https://general-qa2.datacommons.cancer.gov/#/studies' python3 scrape_studies.py

# CLI
python3 scrape_studies.py --url 'https://general-qa2.datacommons.cancer.gov/#/studies'
python3 scrape_studies.py 'https://general-stage.datacommons.cancer.gov/#/studies'

# prod explicitly
python3 scrape_studies.py --url 'https://general.datacommons.cancer.gov/#/studies'
```

Headless (cloud/CI): `HEADLESS=1 python3 scrape_studies.py`

Optional: `GC_BASECOUNT_XLSX`, `GC_OUTPUT_DIR`
