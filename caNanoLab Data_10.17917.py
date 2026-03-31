"""
caNanoLab checks for dataset / publication DOI prefix 10.17917.

Extend CONFIG (URL, Excel path, selectors) to match your QA workflow.
Production portal: https://cananolab.cancer.gov
"""

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
import sys

if sys.platform == "win32" and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

# -----------------------------------
# CONFIG
# -----------------------------------
# Full DOI often looks like 10.17917/xxxx — set the complete string when known.
DATASET_DOI = "10.17917"

# TODO: set your test spreadsheet path when you add Excel-driven checks.
# excel_path = r"C:\Sofia\CRDC_CLI\TestData QA Env\...\your_workbook.xlsx"

# caNanoLab public site (adjust if you use QA/stage).
URL = "https://cananolab.cancer.gov"


def run() -> None:
    with sync_playwright() as p:
        browser = p.chromium.launch(channel="chrome", headless=False)
        page = browser.new_page()
        try:
            page.goto(URL, timeout=120_000)
            try:
                page.get_by_text("Continue").click(timeout=8000)
            except PlaywrightTimeoutError:
                pass
            page.wait_for_load_state("networkidle")
            print(f"Opened {URL} (DOI ref: {DATASET_DOI})")
            print("Add search steps or Excel comparisons here.")
        finally:
            browser.close()


if __name__ == "__main__":
    run()
