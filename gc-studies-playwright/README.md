# GC Studies Playwright scraper

One Chromium run of the General Commons Studies page (headed by default).

- Reads **statbar** Studies / Participants / Samples / Files
- Scrapes each Study’s participants + files; compares to Excel baseline
- Verifies **default sort** + per-column asc/desc against `sorting_baseline.json`
- Writes one report: `output/gc-studies-playwright-report.html`

## Install

```bash
cd gc-studies-playwright
python3 -m pip install -r requirements.txt
python3 -m playwright install chromium
```

## Run

```bash
python3 scrape_studies.py
```

### Headed / headless

| Mode | How |
| --- | --- |
| Headed (default) | `python3 scrape_studies.py` or `--headed` |
| Headless | `HEADLESS=1 python3 scrape_studies.py` or `--headless` |

```bash
python3 scrape_studies.py --headed
HEADLESS=1 python3 scrape_studies.py
python3 scrape_studies.py --headless
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
python3 scrape_studies.py --url 'https://general-qa2.datacommons.cancer.gov/#/studies'
```

Optional: `GC_BASECOUNT_XLSX`, `GC_SORTING_BASELINE`, `GC_OUTPUT_DIR`

Baselines:
- Excel counts: package-root `GC Studypage_Basecount.xlsx` (fallback: `data/`)
- Sorting: package-root `sorting_baseline.json` (default Study Name ASC; all 8 columns sortable PASS)

`verify_sorting.py` is an alias that runs the same combined flow.
