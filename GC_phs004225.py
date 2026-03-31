from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
import pandas as pd
import time
import re
import html
import sys

if sys.platform == "win32" and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

# -----------------------------------
# CONFIG
# -----------------------------------
excel_path = r"C:\Sofia\CRDC_CLI\TestData QA Env\GC March release\phs004225\Base counts for studies.xlsx"
URL = "https://general-qa.datacommons.cancer.gov/#/data"
PHS = "phs004225"

results = []

# -----------------------------------
# LOAD EXCEL
# -----------------------------------
raw_df = pd.read_excel(excel_path, engine="openpyxl", header=None)

try:
    header_row = next(
        i
        for i in range(10)
        if "study name" in " ".join(raw_df.iloc[i].astype(str).str.lower())
    )
except StopIteration as e:
    raise RuntimeError(
        "Could not find a header row containing 'study name' in the first 10 rows."
    ) from e

df = pd.read_excel(excel_path, engine="openpyxl", header=header_row)
df.columns = df.columns.map(str).str.strip()
df = df.loc[:, ~df.columns.str.contains("^Unnamed", na=False)]
df["Study Clean"] = df["Study Name"].astype(str).str.replace("\n", " ").str.strip()

# -----------------------------------
# HELPERS
# -----------------------------------
def get_tab_count(page, tab_name):
    tab = page.get_by_role("tab", name=re.compile(tab_name))
    tab.wait_for(timeout=10000)
    text = tab.inner_text()
    match = re.search(r"\(([\d,]+)\)", text)
    if not match:
        raise ValueError(
            f"No count in parentheses found for tab {tab_name!r}; text was: {text!r}"
        )
    return int(match.group(1).replace(",", ""))


def open_study_filter(page):
    page.locator("//div[@id='Study Name']").first.click()
    page.locator("//p[contains(@class,'MuiTypography-body1')]").first.wait_for(timeout=10000)
    time.sleep(2)


def build_xpath_text(text):
    if "'" in text:
        parts = text.split("'")
        return "concat(" + ", \"'\", ".join([f"'{p}'" for p in parts]) + ")"
    return f"'{text}'"


def toggle_checkbox(page, study_name):
    # Full title required: many studies share the same first 40 characters (e.g. CPTAC
    # Glioblastoma vs Head and Neck), and matching a short prefix clicks the wrong row.
    name = study_name.strip()
    safe_text = build_xpath_text(name)

    def find_and_click():
        study = page.locator(f"xpath=//p[contains(normalize-space(), {safe_text})]")
        n = study.count()
        if n == 0:
            return False
        # If multiple <p> still match, use exact normalized text (substring collisions).
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

    print(f"🔄 Reset scroll for: {study_name}")
    page.mouse.wheel(0, -10000)
    time.sleep(1.5)

    for i in range(35):
        if find_and_click():
            return
        print(f"↘️ Scrolling for: {study_name} (step {i+1})")
        page.mouse.wheel(0, 400)
        time.sleep(1.5)

    raise Exception(f"❌ Study not found: {study_name}")


def deselect_previous_study(page, prev_study):
    if prev_study:
        open_study_filter(page)
        toggle_checkbox(page, prev_study)
        time.sleep(1.5)
        page.keyboard.press("Escape")
        time.sleep(1)


# -----------------------------------
# HTML REPORT
# -----------------------------------
def generate_html_report(results):
    html_doc = """
    <html>
    <head>
        <title>QA Report</title>
        <style>
            body { font-family: Arial; }
            table { border-collapse: collapse; width: 100%; }
            th, td { border: 1px solid #ddd; padding: 8px; text-align: center; }
            th { background-color: #333; color: white; }
            .pass { background-color: #d4edda; }
            .fail { background-color: #f8d7da; }
        </style>
    </head>
    <body>
        <h2>Study QA Report</h2>
        <table>
            <tr>
                <th>Study Name</th>
                <th>UI Participants</th>
                <th>Excel Participants</th>
                <th>UI Samples</th>
                <th>Excel Samples</th>
                <th>UI Files</th>
                <th>Excel Files</th>
                <th>UI Protocols</th>
                <th>Excel Protocols</th>
                <th>Status</th>
            </tr>
    """

    for r in results:
        row_class = "pass" if r["status"] == "PASS" else "fail"

        html_doc += f"""
        <tr class="{row_class}">
            <td>{html.escape(str(r['study']), quote=True)}</td>
            <td>{r['ui_p']}</td>
            <td>{r['exp_p']}</td>
            <td>{r['ui_s']}</td>
            <td>{r['exp_s']}</td>
            <td>{r['ui_f']}</td>
            <td>{r['exp_f']}</td>
            <td>{r['ui_pr']}</td>
            <td>{r['exp_pr']}</td>
            <td>{html.escape(str(r['status']), quote=True)}</td>
        </tr>
        """

    html_doc += "</table></body></html>"

    with open("QA_Report.html", "w", encoding="utf-8") as f:
        f.write(html_doc)


# -----------------------------------
# PLAYWRIGHT
# -----------------------------------
with sync_playwright() as p:

    browser = p.chromium.launch(channel="chrome", headless=False)
    page = browser.new_page()

    try:
        page.goto(URL)

        try:
            page.get_by_text("Continue").click(timeout=8000)
        except PlaywrightTimeoutError:
            pass

        page.wait_for_load_state("networkidle")

        print("🔍 Selecting PHS...")
        page.get_by_text("PHS ACCESSION").click()
        page.locator("input[placeholder*='phs']").fill(PHS)
        time.sleep(2)

        page.get_by_text(PHS).first.click()
        time.sleep(3)

        for i, row in df.iterrows():

            study_name = row["Study Clean"]

            try:
                print(f"\n🔍 Study {i+1}: {study_name}")

                prev_study = results[-1]["study"] if results else None
                deselect_previous_study(page, prev_study)

                time.sleep(1.5)

                open_study_filter(page)
                toggle_checkbox(page, study_name)

                page.keyboard.press("Escape")
                time.sleep(2.5)

                p_val = get_tab_count(page, "Participants")
                s_val = get_tab_count(page, "Samples")
                f_val = get_tab_count(page, "Files")
                pr_val = get_tab_count(page, "Protocols")

                exp_p = int(row["Number of Participants"])
                exp_s = int(row["Samples"])
                exp_f = int(row["Number of Files"])
                exp_pr = int(row["Protocols"])

                status = (
                    "PASS"
                    if (
                        p_val == exp_p
                        and s_val == exp_s
                        and f_val == exp_f
                        and pr_val == exp_pr
                    )
                    else "FAIL"
                )

                print(f"UI → P:{p_val}, S:{s_val}, F:{f_val}, PR:{pr_val}")
                print(f"Excel → P:{exp_p}, S:{exp_s}, F:{exp_f}, PR:{exp_pr}")
                print("✅ PASS" if status == "PASS" else "❌ FAIL")

                results.append(
                    {
                        "study": study_name,
                        "ui_p": p_val,
                        "ui_s": s_val,
                        "ui_f": f_val,
                        "ui_pr": pr_val,
                        "exp_p": exp_p,
                        "exp_s": exp_s,
                        "exp_f": exp_f,
                        "exp_pr": exp_pr,
                        "status": status,
                    }
                )

            except Exception as e:
                print("❌ ERROR:", e)

        generate_html_report(results)
        print("\n📄 HTML report generated: QA_Report.html")

        print("\n🟢 Done")
    finally:
        browser.close()
