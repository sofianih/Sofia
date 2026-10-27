# GC Studies Playwright scraper

Compact headed Chromium scrape of https://general.datacommons.cancer.gov/#/studies

- Reads **statbar** Studies / Participants / Samples / Files
- Scrapes each Study’s participants + files from the table
- Compares to package-root `GC Studypage_Basecount.xlsx` (fallback: `data/`)
- Writes `output/gc-studies-playwright-report.html`

```bash
cd gc-studies-playwright
python3 -m pip install -r requirements.txt
python3 -m playwright install chromium
python3 scrape_studies.py          # headed (default)
HEADLESS=1 python3 scrape_studies.py
```
