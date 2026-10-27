#!/usr/bin/env python3
"""
Drive the NCI General Commons Studies UI with Playwright and emit an HTML report.

Does NOT call GraphQL/API snapshot endpoints for study inventory data — all
PHS / participant / file counts come from what the browser renders.

Compares each scraped study against ``data/GC_Studypage_Basecount.xlsx`` using
the Excel **Study** column as the key, checking Number of Participants and
Number of Files (Excel typo ``Nubmer of Files`` is accepted).
"""

from __future__ import annotations

import html
import json
import os
import re
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import pandas as pd
from playwright.sync_api import Page, TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

PACKAGE_DIR = Path(__file__).resolve().parent

STUDIES_URL = os.environ.get(
    "GC_STUDIES_URL",
    "https://general.datacommons.cancer.gov/#/studies",
).strip() or "https://general.datacommons.cancer.gov/#/studies"

DEFAULT_EXCEL = PACKAGE_DIR / "data" / "GC_Studypage_Basecount.xlsx"
EXCEL_PATH = Path(
    os.environ.get("GC_BASECOUNT_XLSX", str(DEFAULT_EXCEL)).strip() or str(DEFAULT_EXCEL)
)

OUTPUT_DIR = Path(os.environ.get("GC_OUTPUT_DIR", str(PACKAGE_DIR / "output")))
EXPECTED_STUDIES = 61


def env_truthy(name: str) -> bool | None:
    """Return True/False if env var is set to a known truthy/falsey value, else None."""
    raw = os.environ.get(name)
    if raw is None or str(raw).strip() == "":
        return None
    return str(raw).strip().lower() in {"1", "true", "yes", "on", "headless"}


def resolve_headless() -> bool:
    """
    Local default is headed (headless=False).

    Override with HEADLESS=1 or PLAYWRIGHT_HEADLESS=1 for cloud/CI.
    Explicit HEADLESS=0 / PLAYWRIGHT_HEADLESS=0 forces headed.
    """
    for key in ("HEADLESS", "PLAYWRIGHT_HEADLESS"):
        val = env_truthy(key)
        if val is not None:
            # HEADLESS=1 → headless True; HEADLESS=0 → headed
            # For PLAYWRIGHT_HEADLESS the same convention applies.
            raw = os.environ.get(key, "").strip().lower()
            if raw in {"0", "false", "no", "off", "headed"}:
                return False
            return True
    return False  # default: headed


HEADLESS = resolve_headless()


def log(msg: str) -> None:
    print(msg, flush=True)


def parse_int(text: Any) -> int | None:
    if text is None:
        return None
    if isinstance(text, bool):
        return None
    if isinstance(text, (int, float)) and not pd.isna(text):
        return int(text)
    s = str(text).strip()
    if s.lower() in {"", "nan", "none", "—", "-"}:
        return None
    cleaned = re.sub(r"[^\d-]", "", s)
    if cleaned in {"", "-"}:
        return None
    try:
        return int(cleaned)
    except ValueError:
        return None


def normalize_phs(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip()


def dismiss_warning_dialog(page: Page) -> None:
    """Close the US government system-use warning if it is present."""
    continue_btn = page.get_by_role("button", name=re.compile(r"^Continue$", re.I))
    try:
        continue_btn.first.wait_for(state="visible", timeout=5000)
    except PlaywrightTimeoutError:
        return
    try:
        continue_btn.first.click(timeout=5000, force=True)
    except PlaywrightTimeoutError:
        try:
            continue_btn.first.evaluate("el => el.click()")
        except Exception:  # noqa: BLE001
            return
    for _ in range(20):
        page.wait_for_timeout(200)
        remaining = page.locator('[role="dialog"]').filter(
            has_text=re.compile(r"Warning|Government system|authorized use", re.I)
        )
        if remaining.count() == 0:
            break
        if continue_btn.count():
            try:
                continue_btn.first.click(timeout=1000, force=True)
            except PlaywrightTimeoutError:
                pass
    page.wait_for_timeout(300)


def wait_for_studies_table(page: Page) -> None:
    page.locator("table tbody tr").first.wait_for(state="visible", timeout=60000)


def set_rows_per_page(page: Page, value: str = "100") -> None:
    dismiss_warning_dialog(page)
    selector = page.locator('[aria-label="rows per page dropdown selector"]').first
    selector.wait_for(state="visible", timeout=15000)
    try:
        selector.click(timeout=10000)
    except PlaywrightTimeoutError:
        dismiss_warning_dialog(page)
        selector.click(timeout=10000, force=True)
    page.get_by_role("option", name=value, exact=True).click()
    page.wait_for_timeout(800)
    wait_for_studies_table(page)


def read_listing_rows(page: Page) -> list[dict[str, Any]]:
    """Parse every visible row from the Studies listing table."""
    rows = page.evaluate(
        """() => {
          const headers = [...document.querySelectorAll('table thead th')].map(
            (th) => (th.innerText || '').trim().toLowerCase()
          );
          const find = (candidates) => {
            for (const c of candidates) {
              const i = headers.findIndex((h) => h === c);
              if (i >= 0) return i;
            }
            for (const c of candidates) {
              const i = headers.findIndex((h) => h.includes(c));
              if (i >= 0) return i;
            }
            return -1;
          };
          const iStudy = find(['study']) === -1 ? 0 : find(['study']);
          const iVersion = find(['study version']);
          const iName = find(['study name']);
          const iProgram = find(['program name']);
          const iTypes = find(['data types']);
          const iAccess = find(['study access']);
          const iPart = find(['number of participants', 'participants']);
          const iFiles = find(['number of files', 'files']);

          return [...document.querySelectorAll('table tbody tr')].map((tr) => {
            const cells = [...tr.querySelectorAll('td')];
            const cellText = (i) => (i >= 0 && cells[i] ? (cells[i].innerText || '').trim() : '');
            const link = cells[iStudy]?.querySelector('a');
            return {
              phs: cellText(iStudy),
              study_version: cellText(iVersion),
              study_name: cellText(iName),
              program_name: cellText(iProgram),
              data_types: cellText(iTypes),
              study_access: cellText(iAccess),
              participants_text: cellText(iPart),
              files_text: cellText(iFiles),
              detail_href: link ? link.getAttribute('href') || '' : '',
            };
          });
        }"""
    )
    studies: list[dict[str, Any]] = []
    for row in rows:
        phs = normalize_phs(row.get("phs"))
        if not phs:
            continue
        href = (row.get("detail_href") or "").strip()
        if href.startswith("#/"):
            detail_url = STUDIES_URL.split("#")[0] + href
        elif href.startswith("http"):
            detail_url = href
        else:
            detail_url = f"https://general.datacommons.cancer.gov/#/study/{phs}"
        studies.append(
            {
                "phs": phs,
                "study_version": row.get("study_version") or "",
                "study_name": row.get("study_name") or "",
                "program_name": row.get("program_name") or "",
                "data_types": row.get("data_types") or "",
                "study_access": row.get("study_access") or "",
                "listing_participants": parse_int(row.get("participants_text")),
                "listing_files": parse_int(row.get("files_text")),
                "detail_url": detail_url,
            }
        )
    return studies


def scrape_study_detail(page: Page, study: dict[str, Any]) -> dict[str, Any]:
    """Open a study detail page and read PHS / participants / files from the UI."""
    result = {
        **study,
        "detail_phs": None,
        "detail_participants": None,
        "detail_files": None,
        "detail_study_name": None,
        "detail_ok": False,
        "error": None,
    }
    try:
        page.goto(study["detail_url"], wait_until="domcontentloaded", timeout=90000)
        dismiss_warning_dialog(page)
        page.wait_for_timeout(1500)
        try:
            page.get_by_text(re.compile(r"STUDY ACCESSION|NUMBER OF PARTICIPANTS", re.I)).first.wait_for(
                timeout=45000
            )
        except PlaywrightTimeoutError:
            page.wait_for_timeout(2000)

        extracted = page.evaluate(
            """(expectedPhs) => {
              const body = (document.body && document.body.innerText) || '';
              const grabAfter = (label) => {
                const re = new RegExp(label + '\\\\s*\\\\n?\\\\s*([^\\\\n]+)', 'i');
                const m = body.match(re);
                return m ? m[1].trim() : null;
              };
              let phs = grabAfter('STUDY ACCESSION');
              let participants = grabAfter('NUMBER OF PARTICIPANTS');
              let files = grabAfter('NUMBER OF FILES');
              let name = null;
              const nameMatch = body.match(/Study\\s*:\\s*([^\\n]+)/i);
              if (nameMatch) name = nameMatch[1].trim();
              if (!participants) {
                const m = body.match(/([\\d,]+)\\s*\\n\\s*PARTICIPANTS/i);
                if (m) participants = m[1];
              }
              if (!files) {
                const m = body.match(/([\\d,]+)\\s*\\n\\s*FILES/i);
                if (m) files = m[1];
              }
              if (!phs && expectedPhs && body.includes(expectedPhs)) {
                phs = expectedPhs;
              }
              return { phs, participants, files, name, url: location.href };
            }""",
            study["phs"],
        )
        result["detail_phs"] = normalize_phs(extracted.get("phs")) or study["phs"]
        result["detail_participants"] = parse_int(extracted.get("participants"))
        result["detail_files"] = parse_int(extracted.get("files"))
        result["detail_study_name"] = extracted.get("name")
        result["detail_ok"] = (
            bool(result["detail_phs"])
            and result["detail_participants"] is not None
            and result["detail_files"] is not None
        )
        if not result["detail_ok"]:
            result["error"] = (
                f"Incomplete detail scrape: phs={result['detail_phs']!r} "
                f"participants={result['detail_participants']!r} files={result['detail_files']!r}"
            )
    except Exception as exc:  # noqa: BLE001
        result["error"] = f"{type(exc).__name__}: {exc}"
        result["detail_ok"] = False
    return result


def finalize_record(rec: dict[str, Any]) -> dict[str, Any]:
    """Prefer detail-page values; fall back to listing table values."""
    phs = normalize_phs(rec.get("detail_phs") or rec.get("phs"))
    participants = (
        rec.get("detail_participants")
        if rec.get("detail_participants") is not None
        else rec.get("listing_participants")
    )
    files = rec.get("detail_files") if rec.get("detail_files") is not None else rec.get("listing_files")
    return {
        "phs": phs,
        "study_name": rec.get("detail_study_name") or rec.get("study_name") or "",
        "study_version": rec.get("study_version") or "",
        "program_name": rec.get("program_name") or "",
        "data_types": rec.get("data_types") or "",
        "study_access": rec.get("study_access") or "",
        "number_of_participants": participants,
        "number_of_files": files,
        "detail_url": rec.get("detail_url") or "",
        "source": "detail" if rec.get("detail_ok") else "listing_fallback",
        "detail_ok": bool(rec.get("detail_ok")),
        "error": rec.get("error"),
        "listing_participants": rec.get("listing_participants"),
        "listing_files": rec.get("listing_files"),
        "detail_participants": rec.get("detail_participants"),
        "detail_files": rec.get("detail_files"),
    }


def load_excel_baseline(path: Path) -> dict[str, dict[str, Any]]:
    """
    Load baseline keyed by Excel **Study** column.

    Version-only continuation rows (blank Study) are skipped.
    Accepts the misspelled ``Nubmer of Files`` column.
    """
    if not path.is_file():
        raise FileNotFoundError(f"Baseline Excel not found: {path}")

    df = pd.read_excel(path, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    study_col = next((c for c in df.columns if c.lower() == "study"), None)
    if not study_col:
        raise RuntimeError(f"Excel missing Study column; got {list(df.columns)}")

    part_col = next(
        (c for c in df.columns if "participant" in c.lower()),
        None,
    )
    files_col = next(
        (
            c
            for c in df.columns
            if "file" in c.lower() or c.lower().startswith("nubmer")
        ),
        None,
    )
    if not part_col or not files_col:
        raise RuntimeError(
            f"Excel missing participants/files columns; got {list(df.columns)}"
        )

    name_col = next((c for c in df.columns if c.lower() == "study name"), None)

    baseline: dict[str, dict[str, Any]] = {}
    for _, row in df.iterrows():
        phs = normalize_phs(row.get(study_col))
        if not phs or phs.lower() == "nan":
            continue
        key = phs.lower()
        baseline[key] = {
            "phs": phs,
            "study_name": normalize_phs(row.get(name_col)) if name_col else "",
            "excel_participants": parse_int(row.get(part_col)),
            "excel_files": parse_int(row.get(files_col)),
        }
    return baseline


def compare_to_baseline(
    studies: list[dict[str, Any]], baseline: dict[str, dict[str, Any]]
) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    """
    Compare each UI study to Excel by Study/PHS id.

    Status values: PASS, FAIL, MISSING_IN_BASE, MISSING_IN_UI
    """
    compared: list[dict[str, Any]] = []
    seen_keys: set[str] = set()

    for s in studies:
        key = (s.get("phs") or "").lower()
        seen_keys.add(key)
        base = baseline.get(key)
        ui_p = s.get("number_of_participants")
        ui_f = s.get("number_of_files")
        if base is None:
            status = "MISSING_IN_BASE"
            excel_p = None
            excel_f = None
            participants_match = False
            files_match = False
            mismatch_detail = "Study not found in Excel baseline"
        else:
            excel_p = base.get("excel_participants")
            excel_f = base.get("excel_files")
            participants_match = ui_p == excel_p
            files_match = ui_f == excel_f
            if participants_match and files_match:
                status = "PASS"
                mismatch_detail = ""
            else:
                status = "FAIL"
                parts = []
                if not participants_match:
                    parts.append(f"participants UI={ui_p} Excel={excel_p}")
                if not files_match:
                    parts.append(f"files UI={ui_f} Excel={excel_f}")
                mismatch_detail = "; ".join(parts)

        compared.append(
            {
                **s,
                "excel_participants": excel_p,
                "excel_files": excel_f,
                "participants_match": participants_match,
                "files_match": files_match,
                "compare_status": status,
                "mismatch_detail": mismatch_detail,
            }
        )

    missing_in_ui: list[dict[str, Any]] = []
    for key, base in baseline.items():
        if key in seen_keys:
            continue
        missing_in_ui.append(
            {
                "phs": base["phs"],
                "study_name": base.get("study_name") or "",
                "study_version": "",
                "program_name": "",
                "data_types": "",
                "study_access": "",
                "number_of_participants": None,
                "number_of_files": None,
                "detail_url": "",
                "source": "excel_only",
                "detail_ok": False,
                "error": "Present in Excel baseline but not scraped from UI",
                "excel_participants": base.get("excel_participants"),
                "excel_files": base.get("excel_files"),
                "participants_match": False,
                "files_match": False,
                "compare_status": "MISSING_IN_UI",
                "mismatch_detail": "Study in Excel but not found on live Studies page",
            }
        )

    all_rows = compared + missing_in_ui
    all_rows.sort(key=lambda r: (r.get("phs") or "").lower())

    summary = {
        "baseline_path": str(EXCEL_PATH),
        "baseline_studies": len(baseline),
        "ui_studies": len(studies),
        "pass": sum(1 for r in all_rows if r["compare_status"] == "PASS"),
        "fail": sum(1 for r in all_rows if r["compare_status"] == "FAIL"),
        "missing_in_base": sum(1 for r in all_rows if r["compare_status"] == "MISSING_IN_BASE"),
        "missing_in_ui": sum(1 for r in all_rows if r["compare_status"] == "MISSING_IN_UI"),
    }
    return all_rows, summary


def fmt_num(n: int | None) -> str:
    if n is None:
        return "—"
    return f"{n:,}"


def generate_html(studies: list[dict[str, Any]], meta: dict[str, Any]) -> str:
    scrape_failures = [s for s in studies if s.get("error") and s.get("compare_status") != "MISSING_IN_UI"]
    part_sum = sum(s["number_of_participants"] or 0 for s in studies if s.get("number_of_participants") is not None)
    file_sum = sum(s["number_of_files"] or 0 for s in studies if s.get("number_of_files") is not None)
    cmp = meta.get("comparison") or {}

    by_part = sorted(
        [s for s in studies if s.get("number_of_participants") is not None],
        key=lambda s: s.get("number_of_participants") or -1,
        reverse=True,
    )[:5]
    by_files = sorted(
        [s for s in studies if s.get("number_of_files") is not None],
        key=lambda s: s.get("number_of_files") or -1,
        reverse=True,
    )[:5]

    def top_rows(items: list[dict[str, Any]], field: str) -> str:
        out = []
        for s in items:
            name = html.escape((s.get("study_name") or "")[:70])
            phs = html.escape(s.get("phs") or "")
            out.append(
                f"<tr><td class='phs'>{phs}</td><td>{name}</td>"
                f"<td class='num'>{fmt_num(s.get(field))}</td></tr>"
            )
        return "".join(out)

    def access_pill(access: str) -> str:
        a = (access or "").strip()
        cls = "access-open" if a.lower() == "open" else "access-controlled"
        return f"<span class='pill {cls}'>{html.escape(a or '—')}</span>"

    def status_cell(s: dict[str, Any]) -> str:
        st = s.get("compare_status") or "—"
        cls = {
            "PASS": "ok",
            "FAIL": "err",
            "MISSING_IN_BASE": "warn",
            "MISSING_IN_UI": "warn",
        }.get(st, "warn")
        detail = html.escape(s.get("mismatch_detail") or "")
        return (
            f"<span class='status {cls}'>{html.escape(st)}</span>"
            + (f"<div class='muted small'>{detail}</div>" if detail else "")
        )

    body_rows = []
    for i, s in enumerate(studies, 1):
        row_cls = {
            "PASS": "row-pass",
            "FAIL": "row-fail",
            "MISSING_IN_BASE": "row-warn",
            "MISSING_IN_UI": "row-warn",
        }.get(s.get("compare_status") or "", "")
        p_match = s.get("participants_match")
        f_match = s.get("files_match")
        p_cls = "match" if p_match else "mismatch"
        f_cls = "match" if f_match else "mismatch"
        body_rows.append(
            f"""<tr class="{row_cls}">
              <td class="num">{i}</td>
              <td class="phs"><a href="{html.escape(s.get('detail_url') or '#', quote=True)}" target="_blank" rel="noopener">{html.escape(s.get('phs') or '')}</a></td>
              <td class="name">{html.escape(s.get('study_name') or '')}</td>
              <td>{access_pill(s.get('study_access') or '')}</td>
              <td class="num {p_cls}">{fmt_num(s.get('number_of_participants'))}</td>
              <td class="num">{fmt_num(s.get('excel_participants'))}</td>
              <td class="num {f_cls}">{fmt_num(s.get('number_of_files'))}</td>
              <td class="num">{fmt_num(s.get('excel_files'))}</td>
              <td>{status_cell(s)}</td>
            </tr>"""
        )

    fail_rows = [s for s in studies if s.get("compare_status") != "PASS"]
    fail_section = ""
    if fail_rows:
        fr = []
        for s in fail_rows:
            fr.append(
                f"<tr class='row-fail'><td class='phs'>{html.escape(s.get('phs') or '')}</td>"
                f"<td>{html.escape(s.get('study_name') or '')}</td>"
                f"<td class='num'>{fmt_num(s.get('number_of_participants'))}</td>"
                f"<td class='num'>{fmt_num(s.get('excel_participants'))}</td>"
                f"<td class='num'>{fmt_num(s.get('number_of_files'))}</td>"
                f"<td class='num'>{fmt_num(s.get('excel_files'))}</td>"
                f"<td>{status_cell(s)}</td></tr>"
            )
        fail_section = f"""
    <section class="card">
      <h2>Mismatches &amp; gaps</h2>
      <p class="sub">{len(fail_rows)} of {len(studies)} studies did not fully match the Excel baseline</p>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Study (PHS)</th><th>Name</th>
              <th class="num">UI Part.</th><th class="num">Excel Part.</th>
              <th class="num">UI Files</th><th class="num">Excel Files</th>
              <th>Status</th>
            </tr>
          </thead>
          <tbody>{''.join(fr)}</tbody>
        </table>
      </div>
    </section>"""
    else:
        fail_section = """
    <div class="note">
      <strong>All match:</strong> Every Study from the Excel baseline matches the live UI
      participant and file counts.
    </div>"""

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>NCI General Commons — Playwright Studies Report</title>
<style>
  :root {{
    --ink: #0f1c2e;
    --muted: #5a6a7e;
    --line: #d7e0ea;
    --surface: #ffffff;
    --bg0: #e8f1f7;
    --bg1: #f4f8fb;
    --accent: #0b6e99;
    --accent-deep: #084d6b;
    --open: #0f7a4a;
    --open-bg: #e6f6ee;
    --ctrl: #8a4b12;
    --ctrl-bg: #fff1e3;
    --shadow: 0 10px 30px rgba(15, 28, 46, 0.08);
  }}
  * {{ box-sizing: border-box; }}
  body {{
    margin: 0;
    font-family: "Segoe UI", "Helvetica Neue", Arial, sans-serif;
    color: var(--ink);
    background:
      radial-gradient(1200px 500px at 10% -10%, #cfe6f5 0%, transparent 55%),
      radial-gradient(900px 400px at 100% 0%, #d9efe4 0%, transparent 50%),
      linear-gradient(180deg, var(--bg0), var(--bg1) 240px, #eef3f7);
    line-height: 1.45;
  }}
  .wrap {{ max-width: 1360px; margin: 0 auto; padding: 32px 20px 64px; }}
  header.hero {{
    background: linear-gradient(135deg, #0b3d5c 0%, #0b6e99 55%, #1a8f7a 100%);
    color: #fff;
    border-radius: 18px;
    padding: 36px 40px;
    box-shadow: var(--shadow);
    margin-bottom: 28px;
  }}
  header.hero .eyebrow {{
    letter-spacing: 0.12em; text-transform: uppercase; font-size: 12px; opacity: 0.85; margin-bottom: 10px;
  }}
  header.hero h1 {{ margin: 0 0 10px; font-size: clamp(1.6rem, 2.5vw, 2.2rem); font-weight: 700; }}
  header.hero p {{ margin: 0; max-width: 72ch; opacity: 0.92; }}
  header.hero .meta {{
    margin-top: 18px; display: flex; flex-wrap: wrap; gap: 10px 18px; font-size: 0.92rem; opacity: 0.9;
  }}
  header.hero a {{ color: #dff3ff; }}
  .stats {{
    display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr));
    gap: 14px; margin-bottom: 28px;
  }}
  .stat {{
    background: var(--surface); border: 1px solid var(--line); border-radius: 14px;
    padding: 18px; box-shadow: var(--shadow);
  }}
  .stat .label {{ color: var(--muted); font-size: 0.78rem; text-transform: uppercase; letter-spacing: 0.06em; }}
  .stat .value {{ font-size: 1.6rem; font-weight: 700; margin-top: 6px; color: var(--accent-deep); }}
  .stat .hint {{ font-size: 0.8rem; color: var(--muted); margin-top: 4px; }}
  .stat.pass .value {{ color: #0f7a4a; }}
  .stat.fail .value {{ color: #a11; }}
  section.card {{
    background: var(--surface); border: 1px solid var(--line); border-radius: 16px;
    padding: 22px 24px; box-shadow: var(--shadow); margin-bottom: 22px;
  }}
  section.card h2 {{ margin: 0 0 6px; font-size: 1.2rem; }}
  section.card .sub {{ margin: 0 0 16px; color: var(--muted); font-size: 0.95rem; }}
  .grid-2 {{ display: grid; grid-template-columns: 1fr 1fr; gap: 22px; }}
  @media (max-width: 900px) {{
    .grid-2 {{ grid-template-columns: 1fr; }}
    header.hero {{ padding: 28px 22px; }}
  }}
  table {{ width: 100%; border-collapse: collapse; font-size: 0.9rem; }}
  th, td {{ padding: 9px 10px; border-bottom: 1px solid var(--line); text-align: left; vertical-align: top; }}
  th {{
    font-size: 0.72rem; text-transform: uppercase; letter-spacing: 0.05em; color: var(--muted);
    background: #f7fafc; position: sticky; top: 0; z-index: 1;
  }}
  tbody tr:hover {{ background: #f3f9fc; }}
  tr.row-fail {{ background: #fff5f5; }}
  tr.row-warn {{ background: #fffaf0; }}
  tr.row-pass {{ }}
  .num {{ text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }}
  .num.match {{ color: #0f7a4a; font-weight: 600; }}
  .num.mismatch {{ color: #a11; font-weight: 700; }}
  .phs {{ font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace; font-size: 0.86rem; white-space: nowrap; }}
  .phs a {{ color: var(--accent); text-decoration: none; font-weight: 600; }}
  .phs a:hover {{ text-decoration: underline; }}
  .name {{ max-width: 280px; }}
  .muted {{ color: var(--muted); }}
  .small {{ font-size: 0.8rem; margin-top: 4px; }}
  .pill {{ display: inline-block; padding: 2px 10px; border-radius: 999px; font-size: 0.78rem; font-weight: 600; }}
  .access-open {{ background: var(--open-bg); color: var(--open); }}
  .access-controlled {{ background: var(--ctrl-bg); color: var(--ctrl); }}
  .table-wrap {{ max-height: 70vh; overflow: auto; border: 1px solid var(--line); border-radius: 12px; }}
  .note {{
    background: #eef8f3; border: 1px solid #b7e0cc; border-radius: 12px;
    padding: 12px 14px; color: #1e5a3c; font-size: 0.9rem; margin-bottom: 18px;
  }}
  .status.ok {{ color: #0f7a4a; font-weight: 700; }}
  .status.warn {{ color: #8a4b12; font-weight: 700; }}
  .status.err {{ color: #a11; font-weight: 700; }}
  footer {{ margin-top: 8px; color: var(--muted); font-size: 0.85rem; }}
</style>
</head>
<body>
  <div class="wrap">
    <header class="hero">
      <div class="eyebrow">NCI CRDC · General Commons · Playwright UI vs Excel</div>
      <h1>Studies Inventory Report</h1>
      <p>
        Live UI scrape of every study detail page, compared to the Excel
        <strong>Study</strong> column baseline for participants and file counts.
        Headed Chromium by default (no GraphQL inventory shortcut).
      </p>
      <div class="meta">
        <span>Source: <a href="{html.escape(meta['source_url'], quote=True)}" target="_blank" rel="noopener">{html.escape(meta['source_url'])}</a></span>
        <span>Baseline: <code>{html.escape(Path(cmp.get('baseline_path') or '').name)}</code></span>
        <span>Collected: {html.escape(meta['collected_at'])}</span>
        <span>Mode: {html.escape('headed' if not meta.get('headless') else 'headless')}</span>
      </div>
    </header>

    <div class="stats">
      <div class="stat">
        <div class="label">Studies (UI)</div>
        <div class="value">{meta.get('studies_collected', len(studies))}</div>
        <div class="hint">Expected {EXPECTED_STUDIES}</div>
      </div>
      <div class="stat">
        <div class="label">Excel baseline</div>
        <div class="value">{cmp.get('baseline_studies', '—')}</div>
        <div class="hint">Study column rows</div>
      </div>
      <div class="stat pass">
        <div class="label">PASS</div>
        <div class="value">{cmp.get('pass', 0)}</div>
        <div class="hint">Participants + files match</div>
      </div>
      <div class="stat fail">
        <div class="label">FAIL / gaps</div>
        <div class="value">{(cmp.get('fail') or 0) + (cmp.get('missing_in_base') or 0) + (cmp.get('missing_in_ui') or 0)}</div>
        <div class="hint">Fail {cmp.get('fail', 0)} · missing base {cmp.get('missing_in_base', 0)} · missing UI {cmp.get('missing_in_ui', 0)}</div>
      </div>
      <div class="stat">
        <div class="label">Participants (UI sum)</div>
        <div class="value">{fmt_num(part_sum)}</div>
      </div>
      <div class="stat">
        <div class="label">Files (UI sum)</div>
        <div class="value">{fmt_num(file_sum)}</div>
      </div>
    </div>

    <div class="note">
      <strong>Comparison:</strong> For each Excel <code>Study</code> id, UI participants and files
      are matched to Excel <code>Number of Participants</code> and <code>Nubmer of Files</code>.
      Report covers all {len(studies)} study rows (UI ∪ baseline).
      Scrape detail issues: {len(scrape_failures)}.
    </div>

    {fail_section}

    <div class="grid-2">
      <section class="card">
        <h2>Largest by participants</h2>
        <p class="sub">Top 5 from live UI scrape</p>
        <table>
          <thead><tr><th>PHS</th><th>Study</th><th class="num">Participants</th></tr></thead>
          <tbody>{top_rows(by_part, 'number_of_participants')}</tbody>
        </table>
      </section>
      <section class="card">
        <h2>Largest by files</h2>
        <p class="sub">Top 5 from live UI scrape</p>
        <table>
          <thead><tr><th>PHS</th><th>Study</th><th class="num">Files</th></tr></thead>
          <tbody>{top_rows(by_files, 'number_of_files')}</tbody>
        </table>
      </section>
    </div>

    <section class="card">
      <h2>All studies — UI vs Excel</h2>
      <p class="sub">Green UI counts match Excel; red counts mismatch. Status is PASS only when both match.</p>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th class="num">#</th>
              <th>Study (PHS)</th>
              <th>Study name</th>
              <th>Access</th>
              <th class="num">UI Part.</th>
              <th class="num">Excel Part.</th>
              <th class="num">UI Files</th>
              <th class="num">Excel Files</th>
              <th>Match</th>
            </tr>
          </thead>
          <tbody>
            {''.join(body_rows)}
          </tbody>
        </table>
      </div>
    </section>

    <footer>
      Generated by <code>gc-studies-playwright/scrape_studies.py</code>.
      Duration: {html.escape(str(meta.get('duration_sec', '—')))}s.
      Baseline: {html.escape(str(cmp.get('baseline_path') or ''))}.
    </footer>
  </div>
</body>
</html>
"""


def launch_browser(p):
    """Launch Chromium; default headed. Fall back to headless if display unavailable."""
    headless = HEADLESS
    try:
        browser = p.chromium.launch(headless=headless)
        log(f"Browser launched (headless={headless})")
        return browser, headless
    except Exception as exc:  # noqa: BLE001
        if headless:
            raise
        log(f"Headed launch failed ({exc}); falling back to headless. Use HEADLESS=1 to skip headed.")
        browser = p.chromium.launch(headless=True)
        return browser, True


def main() -> int:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    started = datetime.now(timezone.utc)
    log(f"Studies URL: {STUDIES_URL}")
    log(f"Headless default resolved: {HEADLESS} (override with HEADLESS=1)")
    log(f"Excel baseline: {EXCEL_PATH}")
    log(f"Output: {OUTPUT_DIR}")

    baseline = load_excel_baseline(EXCEL_PATH)
    log(f"Excel Study rows loaded: {len(baseline)}")

    with sync_playwright() as p:
        browser, actual_headless = launch_browser(p)
        context = browser.new_context(viewport={"width": 1440, "height": 900})
        page = context.new_page()

        log("Opening Studies listing…")
        page.goto(STUDIES_URL, wait_until="domcontentloaded", timeout=120000)
        dismiss_warning_dialog(page)
        wait_for_studies_table(page)
        set_rows_per_page(page, "100")

        listing = read_listing_rows(page)
        log(f"Listing rows collected: {len(listing)}")
        if len(listing) < EXPECTED_STUDIES:
            log("Fewer than expected rows; paginating with Next page…")
            seen = {s["phs"]: s for s in listing}
            while True:
                next_btn = page.get_by_role("button", name=re.compile(r"Next page", re.I))
                if not next_btn.count() or not next_btn.first.is_enabled():
                    break
                first_phs = listing[0]["phs"] if listing else None
                next_btn.first.click()
                page.wait_for_timeout(800)
                wait_for_studies_table(page)
                page_rows = read_listing_rows(page)
                if page_rows and page_rows[0]["phs"] == first_phs and len(seen) >= len(listing):
                    break
                for row in page_rows:
                    seen[row["phs"]] = row
                listing = list(seen.values())
                log(f"  after page: {len(listing)} unique studies")
                if len(listing) >= EXPECTED_STUDIES:
                    break

        if not listing:
            log("ERROR: no studies found in the listing table.")
            browser.close()
            return 1

        scraped_raw: list[dict[str, Any]] = []
        for i, study in enumerate(listing, 1):
            log(f"[{i}/{len(listing)}] Visiting {study['phs']} …")
            scraped_raw.append(scrape_study_detail(page, study))

        browser.close()

    studies = [finalize_record(r) for r in scraped_raw]
    studies, comparison = compare_to_baseline(studies, baseline)

    ended = datetime.now(timezone.utc)
    meta = {
        "source_url": STUDIES_URL,
        "collected_at": ended.strftime("%Y-%m-%d %H:%M UTC"),
        "browser": "Chromium (Playwright)",
        "headless": actual_headless,
        "listing_count": len(listing),
        "studies_collected": sum(1 for s in studies if s.get("compare_status") != "MISSING_IN_UI"),
        "detail_ok_count": sum(1 for s in studies if s.get("detail_ok")),
        "failures": [
            {"phs": s.get("phs"), "error": s.get("error"), "compare_status": s.get("compare_status")}
            for s in studies
            if s.get("compare_status") != "PASS"
        ],
        "duration_sec": round((ended - started).total_seconds(), 1),
        "method": "playwright-ui-vs-excel",
        "comparison": comparison,
    }

    json_path = OUTPUT_DIR / "studies.json"
    html_path = OUTPUT_DIR / "gc-studies-playwright-report.html"
    json_path.write_text(json.dumps({"meta": meta, "studies": studies}, indent=2), encoding="utf-8")
    html_path.write_text(generate_html(studies, meta), encoding="utf-8")

    log(f"Wrote {json_path}")
    log(f"Wrote {html_path}")
    log(
        f"Done: UI={comparison['ui_studies']} Excel={comparison['baseline_studies']} "
        f"PASS={comparison['pass']} FAIL={comparison['fail']} "
        f"missing_base={comparison['missing_in_base']} missing_ui={comparison['missing_in_ui']} "
        f"in {meta['duration_sec']}s (headless={actual_headless})"
    )
    return 0 if comparison["ui_studies"] else 1


if __name__ == "__main__":
    sys.exit(main())
