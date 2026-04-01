"""
General QA Data Commons — Gabriella Miller / Kids First studies (five PHS + study pairs).

Expected Participants / Samples / Files are **embedded in this file** by default (snapshot
from Kidsfirst basecounts.xlsx) — no local path required.

Optional: set env **KIDS_FIRST_EXCEL_PATH** to a workbook path to load expectations from Excel
instead (same layout: PHS Accession, Study Name, Participants, Samples, Files).
"""

from __future__ import annotations

import html
import os
import re
import sys
import time
from pathlib import Path

import pandas as pd
from playwright.sync_api import (
    Error as PlaywrightError,
    sync_playwright,
    TimeoutError as PlaywrightTimeoutError,
)

if sys.platform == "win32" and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

# -----------------------------------
# CONFIG
# -----------------------------------
URL = "https://general-qa.datacommons.cancer.gov/#/data"

REPORT_PATH = "GC_kidsfirst_Pgm_Report.html"

# Embedded base counts (same order as SCENARIOS). Source: Kidsfirst basecounts.xlsx
# PHS | Participants | Samples | Files
EMBEDDED_BASE_COUNTS: list[tuple[int, int, int]] = [
    (1287, 1460, 49518),  # phs001228
    (1776, 1776, 26788),  # phs001846
    (60, 119, 3342),  # phs002187
    (84, 311, 12727),  # phs001714
    (185, 192, 4728),  # phs001738
]

# (PHS accession, exact study name as in the Study Name filter list)
SCENARIOS: list[tuple[str, str]] = [
    (
        "phs001228",
        "Gabriella Miller Kids First (GMKF) Pediatric Research Program in Susceptibility to "
        "Ewing Sarcoma Based on Germline Risk and Familial History of Cancer",
    ),
    (
        "phs001846",
        "Gabriella Miller Kids First Pediatric Research Program in Genetics at the "
        "Intersection of Childhood Cancer and Birth Defects",
    ),
    (
        "phs002187",
        "Gabriella Miller Kids First Pediatric Research Program in Germline and Somatic "
        "Variants in Myeloid Malignancies in Children",
    ),
    (
        "phs001714",
        "Gabriella Miller Kids First Pediatric Research Program: An Integrated Clinical and "
        "Genomic Analysis of Treatment Failure in Pediatric Osteosarcoma",
    ),
    (
        "phs001738",
        "Kids First Pediatric Research Study in Familial Predisposition to "
        "Hematopoietic Malignancies",
    ),
]

assert len(EMBEDDED_BASE_COUNTS) == len(SCENARIOS), "Embedded counts must match SCENARIOS"


def _find_phs_column(df: pd.DataFrame) -> str | None:
    for c in df.columns:
        if "phs" in str(c).lower():
            return str(c)
    return None


def _norm_phs(val: str) -> str:
    return re.sub(r"\s+", "", str(val).strip().lower())


def load_expectations_workbook(excel_path: str) -> pd.DataFrame:
    path = Path(excel_path)
    if not path.is_file():
        raise FileNotFoundError(
            f"Base counts Excel not found: {path}\n"
            "Set env KIDS_FIRST_EXCEL_PATH or place a workbook at the default path.\n"
            "Expected columns include: Study Name, Number of Participants, Samples, "
            "Number of Files (and optionally a PHS / PHS ACCESSION column)."
        )

    raw_df = pd.read_excel(path, engine="openpyxl", header=None)
    try:
        header_row = next(
            i
            for i in range(min(15, len(raw_df)))
            if "study name" in " ".join(raw_df.iloc[i].astype(str).str.lower())
        )
    except StopIteration as e:
        raise RuntimeError(
            "Could not find a header row containing 'study name' in the first 15 rows."
        ) from e

    df = pd.read_excel(path, engine="openpyxl", header=header_row)
    df.columns = df.columns.map(str).str.strip()
    df = df.loc[:, ~df.columns.str.contains("^Unnamed", na=False)]
    if "Study Name" not in df.columns:
        raise RuntimeError("Excel must contain a 'Study Name' column.")
    df["Study Clean"] = (
        df["Study Name"].astype(str).str.replace("\n", " ").str.strip()
    )
    return df


def lookup_expected_counts(
    df: pd.DataFrame, phs: str, study_display: str
) -> tuple[int, int, int]:
    """Return (participants, samples, files) from the workbook for this PHS + study."""
    study_key = study_display.replace("\n", " ").strip()
    mask_s = df["Study Clean"].str.strip() == study_key
    if not mask_s.any():
        mask_s = df["Study Clean"].str.lower().str.strip() == study_key.lower()

    phs_col = _find_phs_column(df)
    if phs_col:
        mask_p = df[phs_col].map(_norm_phs) == _norm_phs(phs)
        subset = df.loc[mask_s & mask_p]
    else:
        subset = df.loc[mask_s]

    if subset.empty:
        raise ValueError(
            f"No Excel row for PHS={phs!r} and study matching {study_key[:80]!r}…"
        )
    if len(subset) > 1:
        raise ValueError(
            f"Multiple Excel rows ({len(subset)}) for PHS={phs!r}; "
            "add or fix a PHS column in the workbook."
        )

    row = subset.iloc[0]

    def _col(*names: str) -> int:
        for n in names:
            for c in df.columns:
                if str(c).strip().lower() == n.lower():
                    v = row[c]
                    if pd.isna(v):
                        raise ValueError(f"Empty cell in column {c!r}")
                    return int(float(v))
        raise KeyError(f"Missing column; tried {names!r}")

    p = _col("Number of Participants", "Participants", "# Participants")
    s = _col("Samples", "Number of Samples")
    f = _col("Number of Files", "Files", "Number of files")
    return p, s, f


def get_expected_counts(
    scenario_index: int,
    phs: str,
    study: str,
    df: pd.DataFrame | None,
) -> tuple[int, int, int]:
    """Use Excel when `df` is provided; otherwise embedded constants for this index."""
    if df is not None:
        return lookup_expected_counts(df, phs, study)
    return EMBEDDED_BASE_COUNTS[scenario_index]


# -----------------------------------
# UI helpers (aligned with GC_phs004225.py)
# -----------------------------------
def dismiss_continue_if_present(page) -> None:
    try:
        page.get_by_role("button", name=re.compile(r"^Continue$", re.I)).click(
            timeout=12000
        )
    except PlaywrightTimeoutError:
        try:
            page.get_by_text("Continue").click(timeout=5000)
        except PlaywrightTimeoutError:
            pass


def get_tab_count(page, tab_name: str) -> int:
    tab = page.get_by_role("tab", name=re.compile(tab_name, re.I))
    tab.wait_for(timeout=15000)
    text = tab.inner_text()
    match = re.search(r"\(([\d,]+)\)", text)
    if not match:
        raise ValueError(
            f"No count in parentheses for tab {tab_name!r}; text was: {text!r}"
        )
    return int(match.group(1).replace(",", ""))


def _click_phs_facet_header(page) -> None:
    """Open PHS facet — avoid get_by_text strict duplicate (header div + chip span)."""
    hdr = page.locator('[id="PHS Accession"]')
    if hdr.count() > 0:
        hdr.first.click()
    else:
        page.get_by_role(
            "button", name=re.compile(r"PHS\s+Accession", re.I)
        ).first.click()


def _phs_search_input(page):
    """
    PHS facet search box inside the expanded accordion that contains #PHS Accession.
    general-qa uses placeholder e.g. phs000000 for PHS (Study uses e.g. Study Name).
    """
    return page.locator(
        'xpath=//div[contains(@class,"MuiAccordion-root") and contains(@class,"Mui-expanded")]'
        '[.//*[@id="PHS Accession"]]//input[@type="text"]'
        '[contains(@placeholder,"phs000000")][1]'
    )


def _click_phs_suggestion(page, phs_value: str, *, pause: float = 1.5) -> None:
    """Pick PHS accession from autocomplete / list (general-qa)."""
    time.sleep(pause)
    try:
        page.locator(
            "[role='listbox'] [role='option'], "
            ".MuiAutocomplete-listbox li, "
            ".MuiMenu-list li"
        ).filter(has_text=re.compile(re.escape(phs_value), re.I)).first.click(
            timeout=15000
        )
    except PlaywrightTimeoutError:
        page.get_by_text(phs_value, exact=False).first.click(timeout=15000)


def apply_phs_filter(page, phs_value: str) -> None:
    print(f"🔍 PHS ACCESSION → {phs_value}")
    _click_phs_facet_header(page)
    time.sleep(0.7)
    pinp = _phs_search_input(page)
    pinp.scroll_into_view_if_needed()
    pinp.fill(phs_value, force=True)
    _click_phs_suggestion(page, phs_value, pause=2.2)
    time.sleep(3)


def clear_phs_filter(page, phs_value: str) -> None:
    """Open PHS filter and toggle the same accession off (second click)."""
    print(f"🔍 Clear PHS ACCESSION → {phs_value}")
    _click_phs_facet_header(page)
    time.sleep(0.7)
    inp = _phs_search_input(page)
    try:
        inp.scroll_into_view_if_needed(timeout=5000)
    except PlaywrightTimeoutError:
        pass
    inp.fill("", force=True)
    time.sleep(0.25)
    inp.fill(phs_value, force=True)
    try:
        _click_phs_suggestion(page, phs_value, pause=1.2)
    except PlaywrightError:
        pass
    time.sleep(2)
    page.keyboard.press("Escape")
    time.sleep(1)


def open_study_filter(page) -> None:
    page.locator("//div[@id='Study Name']").first.click()
    page.locator("//p[contains(@class,'MuiTypography-body1')]").first.wait_for(
        timeout=15000
    )
    time.sleep(2)


def build_xpath_text(text: str) -> str:
    if "'" in text:
        parts = text.split("'")
        return "concat(" + ", \"'\", ".join([f"'{p}'" for p in parts]) + ")"
    return f"'{text}'"


def toggle_study_checkbox(page, study_name: str) -> None:
    name = study_name.strip()
    safe_text = build_xpath_text(name)

    def find_and_click() -> bool:
        study = page.locator(f"xpath=//p[contains(normalize-space(), {safe_text})]")
        n = study.count()
        if n == 0:
            return False
        if n > 1:
            exact = page.locator(f"xpath=//p[normalize-space()={safe_text}]")
            if exact.count() >= 1:
                study = exact.first
            else:
                study = study.filter(has_text=name).first
        else:
            study = study.first

        study.scroll_into_view_if_needed()
        time.sleep(0.5)
        box = study.bounding_box()
        if box:
            page.mouse.click(box["x"] - 15, box["y"] + box["height"] / 2)
            return True
        return False

    for _ in range(5):
        if find_and_click():
            return
        page.mouse.wheel(0, 300)
        time.sleep(1)

    page.mouse.wheel(0, -10000)
    time.sleep(1.5)
    for i in range(35):
        if find_and_click():
            return
        print(f"↘️ Scrolling for study (step {i + 1})")
        page.mouse.wheel(0, 400)
        time.sleep(1.5)

    raise RuntimeError(f"Study row not found: {study_name[:100]}…")


def select_study(page, study_name: str) -> None:
    print(f"🔍 Study Name → {study_name[:70]}…")
    open_study_filter(page)
    toggle_study_checkbox(page, study_name)
    page.keyboard.press("Escape")
    time.sleep(2.5)


def deselect_study(page, study_name: str) -> None:
    print(f"🔍 Unselect study → {study_name[:70]}…")
    open_study_filter(page)
    toggle_study_checkbox(page, study_name)
    time.sleep(1.5)
    page.keyboard.press("Escape")
    time.sleep(2)


def generate_html_report(rows: list[dict], path: str = REPORT_PATH) -> None:
    doc = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8"/>
  <title>Kids First Program — GC QA Report</title>
  <style>
    body {{ font-family: Arial, sans-serif; margin: 24px; }}
    h1 {{ font-size: 1.25rem; }}
    table {{ border-collapse: collapse; width: 100%; max-width: 1400px; }}
    th, td {{ border: 1px solid #ccc; padding: 8px 10px; text-align: left; vertical-align: top; }}
    th {{ background: #333; color: #fff; text-align: center; }}
    td.count {{ text-align: center; font-weight: 600; }}
    td.pass {{ background: #d4edda; }}
    td.fail {{ background: #f8d7da; }}
    .meta {{ color: #444; margin-bottom: 16px; }}
  </style>
</head>
<body>
  <h1>Gabriella Miller / Kids First — count check report</h1>
  <div class="meta"><strong>URL:</strong> {html.escape(URL)}</div>
  <table>
    <thead>
      <tr>
        <th>PHS accession</th>
        <th>Study name</th>
        <th>Participants counts</th>
        <th>Samples count</th>
        <th>Files count</th>
        <th>Status</th>
      </tr>
    </thead>
    <tbody>
"""
    for r in rows:
        st = r["status"]
        cls = "pass" if st == "PASS" else "fail"
        doc += f"""      <tr>
        <td>{html.escape(r["phs"])}</td>
        <td>{html.escape(r["study"])}</td>
        <td class="count {cls}">{html.escape(r["participants_cell"])}</td>
        <td class="count {cls}">{html.escape(r["samples_cell"])}</td>
        <td class="count {cls}">{html.escape(r["files_cell"])}</td>
        <td class="count {cls}"><strong>{html.escape(st)}</strong></td>
      </tr>
"""
    doc += """    </tbody>
  </table>
</body>
</html>
"""
    with open(path, "w", encoding="utf-8") as f:
        f.write(doc)


def _fmt_cell(ui: int, exp: int) -> str:
    if ui == exp:
        return str(ui)
    return f"{ui} (expected {exp})"


def run() -> None:
    excel_path = os.environ.get("KIDS_FIRST_EXCEL_PATH", "").strip()
    df: pd.DataFrame | None = None
    if excel_path:
        print(f"📎 Loading expectations from Excel: {excel_path}")
        df = load_expectations_workbook(excel_path)
    else:
        print("📎 Using embedded base counts (no KIDS_FIRST_EXCEL_PATH).")

    results: list[dict] = []

    with sync_playwright() as p:
        browser = p.chromium.launch(channel="chrome", headless=False)
        page = browser.new_page()
        try:
            page.goto(URL, timeout=120_000, wait_until="domcontentloaded")
            dismiss_continue_if_present(page)
            try:
                page.wait_for_load_state("networkidle", timeout=90_000)
            except PlaywrightTimeoutError:
                print("ℹ️ networkidle timeout — continuing once DOM is usable.")
            time.sleep(2)

            for i, (phs, study) in enumerate(SCENARIOS):
                exp_p, exp_s, exp_f = get_expected_counts(i, phs, study, df)
                apply_phs_filter(page, phs)
                select_study(page, study)

                try:
                    p_ui = get_tab_count(page, "Participants")
                    s_ui = get_tab_count(page, "Samples")
                    f_ui = get_tab_count(page, "Files")
                except Exception as e:
                    results.append(
                        {
                            "phs": phs,
                            "study": study,
                            "participants_cell": f"— ({e})",
                            "samples_cell": "—",
                            "files_cell": "—",
                            "status": "FAIL",
                        }
                    )
                    deselect_study(page, study)
                    clear_phs_filter(page, phs)
                    continue

                ok = p_ui == exp_p and s_ui == exp_s and f_ui == exp_f
                status = "PASS" if ok else "FAIL"
                print(
                    f"UI P/S/F = {p_ui}/{s_ui}/{f_ui} | Base = {exp_p}/{exp_s}/{exp_f} → {status}"
                )

                results.append(
                    {
                        "phs": phs,
                        "study": study,
                        "participants_cell": _fmt_cell(p_ui, exp_p),
                        "samples_cell": _fmt_cell(s_ui, exp_s),
                        "files_cell": _fmt_cell(f_ui, exp_f),
                        "status": status,
                    }
                )

                deselect_study(page, study)
                clear_phs_filter(page, phs)

            generate_html_report(results)
            print(f"\n📄 HTML report written: {REPORT_PATH}")
        finally:
            browser.close()


if __name__ == "__main__":
    run()
