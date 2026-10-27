#!/usr/bin/env python3
"""
Drive the NCI General Commons Studies UI with Playwright and emit an HTML report.

Does NOT call GraphQL/API snapshot endpoints for study inventory data — all
PHS / participant / file counts come from what the browser renders.
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

from playwright.sync_api import Page, TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

STUDIES_URL = os.environ.get(
    "GC_STUDIES_URL",
    "https://general.datacommons.cancer.gov/#/studies",
).strip() or "https://general.datacommons.cancer.gov/#/studies"

HEADLESS = os.environ.get("PLAYWRIGHT_HEADLESS", "1").strip() not in {"0", "false", "False", "no"}
OUTPUT_DIR = Path(os.environ.get("GC_OUTPUT_DIR", str(Path(__file__).resolve().parent / "output")))
EXPECTED_STUDIES = 61


def log(msg: str) -> None:
    print(msg, flush=True)


def parse_int(text: str | None) -> int | None:
    if text is None:
        return None
    cleaned = re.sub(r"[^\d-]", "", str(text))
    if cleaned in {"", "-"}:
        return None
    try:
        return int(cleaned)
    except ValueError:
        return None


def dismiss_warning_dialog(page: Page) -> None:
    """Close the US government system-use warning if it is present.

    The portal can mount more than one ``role=dialog`` node; the warning itself
    is the one that contains a visible **Continue** button.
    """
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
    # Wait until the warning overlay stops intercepting clicks.
    for _ in range(20):
        page.wait_for_timeout(200)
        remaining = page.locator('[role="dialog"]').filter(
            has_text=re.compile(r"Warning|Government system|authorized use", re.I)
        )
        if remaining.count() == 0:
            break
        # Overlay may linger visually; try Continue again.
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
    # Confirm the table grew (or at least settled).
    wait_for_studies_table(page)


def read_listing_rows(page: Page) -> list[dict[str, Any]]:
    """Parse every visible row from the Studies listing table."""
    rows = page.evaluate(
        """() => {
          const headers = [...document.querySelectorAll('table thead th')].map(
            (th) => (th.innerText || '').trim().toLowerCase()
          );
          const idx = (name) => headers.findIndex((h) => h === name || h.includes(name));
          const iStudy = idx('study') === -1 ? 0 : idx('study');
          // Prefer exact header matches when possible.
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
        phs = (row.get("phs") or "").strip()
        if not phs:
            continue
        href = (row.get("detail_href") or "").strip()
        if href.startswith("#/"):
            detail_url = STUDIES_URL.split("#")[0].rstrip("/") + "/" + href.lstrip("/")
            # Prefer absolute hash URL on same origin
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
                "listing_participants_text": row.get("participants_text") or "",
                "listing_files_text": row.get("files_text") or "",
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
        # Wait for study accession / participants label text.
        page.wait_for_timeout(1500)
        try:
            page.get_by_text(re.compile(r"STUDY ACCESSION|NUMBER OF PARTICIPANTS", re.I)).first.wait_for(
                timeout=45000
            )
        except PlaywrightTimeoutError:
            # Fall back to any body content that includes the PHS id.
            page.wait_for_timeout(2000)

        extracted = page.evaluate(
            """(expectedPhs) => {
              const body = (document.body && document.body.innerText) || '';
              const grabAfter = (label) => {
                const re = new RegExp(label + '\\\\s*\\\\n?\\\\s*([^\\\\n]+)', 'i');
                const m = body.match(re);
                return m ? m[1].trim() : null;
              };
              // Prefer structured labels on the study page.
              let phs = grabAfter('STUDY ACCESSION');
              let participants = grabAfter('NUMBER OF PARTICIPANTS');
              let files = grabAfter('NUMBER OF FILES');
              let name = null;
              const nameMatch = body.match(/Study\\s*:\\s*([^\\n]+)/i);
              if (nameMatch) name = nameMatch[1].trim();

              // Banner counters often appear as "123\\nPARTICIPANTS" / "7,629\\nFILES"
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
        result["detail_phs"] = (extracted.get("phs") or "").strip() or study["phs"]
        result["detail_participants"] = parse_int(extracted.get("participants"))
        result["detail_files"] = parse_int(extracted.get("files"))
        result["detail_study_name"] = extracted.get("name")
        result["detail_ok"] = (
            result["detail_phs"] is not None
            and result["detail_participants"] is not None
            and result["detail_files"] is not None
        )
        if not result["detail_ok"]:
            result["error"] = (
                f"Incomplete detail scrape: phs={result['detail_phs']!r} "
                f"participants={result['detail_participants']!r} files={result['detail_files']!r}"
            )
    except Exception as exc:  # noqa: BLE001 - capture per-study failures into report
        result["error"] = f"{type(exc).__name__}: {exc}"
        result["detail_ok"] = False
    return result


def finalize_record(rec: dict[str, Any]) -> dict[str, Any]:
    """Prefer detail-page values; fall back to listing table values."""
    phs = rec.get("detail_phs") or rec.get("phs")
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


def fmt_num(n: int | None) -> str:
    if n is None:
        return "—"
    return f"{n:,}"


def generate_html(studies: list[dict[str, Any]], meta: dict[str, Any]) -> str:
    failures = [s for s in studies if s.get("error") or not s.get("detail_ok")]
    part_sum = sum(s["number_of_participants"] or 0 for s in studies)
    file_sum = sum(s["number_of_files"] or 0 for s in studies)
    controlled = sum(1 for s in studies if (s.get("study_access") or "").lower() == "controlled")
    open_access = sum(1 for s in studies if (s.get("study_access") or "").lower() == "open")

    by_part = sorted(
        studies,
        key=lambda s: (s.get("number_of_participants") is not None, s.get("number_of_participants") or -1),
        reverse=True,
    )[:5]
    by_files = sorted(
        studies,
        key=lambda s: (s.get("number_of_files") is not None, s.get("number_of_files") or -1),
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

    body_rows = []
    for i, s in enumerate(studies, 1):
        err = html.escape(s.get("error") or "")
        status = "OK" if s.get("detail_ok") else ("FALLBACK" if not s.get("error") else "ERROR")
        status_cls = "ok" if status == "OK" else ("warn" if status == "FALLBACK" else "err")
        body_rows.append(
            f"""<tr>
              <td class="num">{i}</td>
              <td class="phs"><a href="{html.escape(s.get('detail_url') or '#', quote=True)}" target="_blank" rel="noopener">{html.escape(s.get('phs') or '')}</a></td>
              <td class="name">{html.escape(s.get('study_name') or '')}</td>
              <td>{html.escape(s.get('program_name') or '')}</td>
              <td>{html.escape(s.get('data_types') or '')}</td>
              <td>{access_pill(s.get('study_access') or '')}</td>
              <td class="num">{fmt_num(s.get('number_of_participants'))}</td>
              <td class="num">{fmt_num(s.get('number_of_files'))}</td>
              <td><span class="status {status_cls}">{status}</span>{(' <span class="muted">' + err + '</span>') if err else ''}</td>
            </tr>"""
        )

    failure_note = (
        f"{len(failures)} study detail page(s) had issues; listing values used where needed."
        if failures
        else "All study detail pages returned PHS, participants, and file counts."
    )

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
  .wrap {{ max-width: 1280px; margin: 0 auto; padding: 32px 20px 64px; }}
  header.hero {{
    background: linear-gradient(135deg, #0b3d5c 0%, #0b6e99 55%, #1a8f7a 100%);
    color: #fff;
    border-radius: 18px;
    padding: 36px 40px;
    box-shadow: var(--shadow);
    margin-bottom: 28px;
  }}
  header.hero .eyebrow {{
    letter-spacing: 0.12em;
    text-transform: uppercase;
    font-size: 12px;
    opacity: 0.85;
    margin-bottom: 10px;
  }}
  header.hero h1 {{ margin: 0 0 10px; font-size: clamp(1.6rem, 2.5vw, 2.2rem); font-weight: 700; }}
  header.hero p {{ margin: 0; max-width: 70ch; opacity: 0.92; }}
  header.hero .meta {{
    margin-top: 18px; display: flex; flex-wrap: wrap; gap: 10px 18px; font-size: 0.92rem; opacity: 0.9;
  }}
  header.hero a {{ color: #dff3ff; }}
  .stats {{
    display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
    gap: 14px; margin-bottom: 28px;
  }}
  .stat {{
    background: var(--surface); border: 1px solid var(--line); border-radius: 14px;
    padding: 18px; box-shadow: var(--shadow);
  }}
  .stat .label {{ color: var(--muted); font-size: 0.82rem; text-transform: uppercase; letter-spacing: 0.06em; }}
  .stat .value {{ font-size: 1.7rem; font-weight: 700; margin-top: 6px; color: var(--accent-deep); }}
  .stat .hint {{ font-size: 0.8rem; color: var(--muted); margin-top: 4px; }}
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
  table {{ width: 100%; border-collapse: collapse; font-size: 0.92rem; }}
  th, td {{ padding: 10px 12px; border-bottom: 1px solid var(--line); text-align: left; vertical-align: top; }}
  th {{
    font-size: 0.78rem; text-transform: uppercase; letter-spacing: 0.05em; color: var(--muted);
    background: #f7fafc; position: sticky; top: 0; z-index: 1;
  }}
  tbody tr:hover {{ background: #f3f9fc; }}
  .num {{ text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }}
  .phs {{ font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace; font-size: 0.88rem; white-space: nowrap; }}
  .phs a {{ color: var(--accent); text-decoration: none; font-weight: 600; }}
  .phs a:hover {{ text-decoration: underline; }}
  .name {{ max-width: 320px; }}
  .muted {{ color: var(--muted); }}
  .pill {{ display: inline-block; padding: 2px 10px; border-radius: 999px; font-size: 0.78rem; font-weight: 600; }}
  .access-open {{ background: var(--open-bg); color: var(--open); }}
  .access-controlled {{ background: var(--ctrl-bg); color: var(--ctrl); }}
  .table-wrap {{
    max-height: 70vh; overflow: auto; border: 1px solid var(--line); border-radius: 12px;
  }}
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
      <div class="eyebrow">NCI CRDC · General Commons · Playwright UI scrape</div>
      <h1>Studies Inventory Report</h1>
      <p>
        Collected by driving the live Studies listing and each study detail page in Chromium.
        No GraphQL inventory shortcut was used for these counts.
      </p>
      <div class="meta">
        <span>Source: <a href="{html.escape(meta['source_url'], quote=True)}" target="_blank" rel="noopener">{html.escape(meta['source_url'])}</a></span>
        <span>Collected: {html.escape(meta['collected_at'])}</span>
        <span>Method: Playwright UI ({html.escape(meta['browser'])})</span>
      </div>
    </header>

    <div class="stats">
      <div class="stat">
        <div class="label">Studies collected</div>
        <div class="value">{len(studies)}</div>
        <div class="hint">Expected {EXPECTED_STUDIES}</div>
      </div>
      <div class="stat">
        <div class="label">Participants (sum)</div>
        <div class="value">{fmt_num(part_sum)}</div>
        <div class="hint">From scraped UI values</div>
      </div>
      <div class="stat">
        <div class="label">Files (sum)</div>
        <div class="value">{fmt_num(file_sum)}</div>
        <div class="hint">From scraped UI values</div>
      </div>
      <div class="stat">
        <div class="label">Access mix</div>
        <div class="value">{controlled}/{open_access}</div>
        <div class="hint">Controlled / Open</div>
      </div>
    </div>

    <div class="note">
      <strong>Coverage:</strong> {html.escape(failure_note)}
      Listing page showed <strong>{meta.get('listing_count', len(studies))}</strong> rows after setting Rows per page to 100;
      each linked study detail page was opened to verify PHS, participants, and files.
    </div>

    <div class="grid-2">
      <section class="card">
        <h2>Largest by participants</h2>
        <p class="sub">Top 5 studies by number of participants</p>
        <table>
          <thead><tr><th>PHS</th><th>Study</th><th class="num">Participants</th></tr></thead>
          <tbody>{top_rows(by_part, 'number_of_participants')}</tbody>
        </table>
      </section>
      <section class="card">
        <h2>Largest by files</h2>
        <p class="sub">Top 5 studies by number of files</p>
        <table>
          <thead><tr><th>PHS</th><th>Study</th><th class="num">Files</th></tr></thead>
          <tbody>{top_rows(by_files, 'number_of_files')}</tbody>
        </table>
      </section>
    </div>

    <section class="card">
      <h2>All studies</h2>
      <p class="sub">PHS / study ID, participants, and file counts as shown in the UI</p>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th class="num">#</th>
              <th>PHS / ID</th>
              <th>Study name</th>
              <th>Program</th>
              <th>Data types</th>
              <th>Access</th>
              <th class="num">Participants</th>
              <th class="num">Files</th>
              <th>Detail visit</th>
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
      Failures: {len(failures)}. Duration: {html.escape(str(meta.get('duration_sec', '—')))}s.
    </footer>
  </div>
</body>
</html>
"""


def main() -> int:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    started = datetime.now(timezone.utc)
    log(f"Studies URL: {STUDIES_URL}")
    log(f"Headless: {HEADLESS}")
    log(f"Output: {OUTPUT_DIR}")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS)
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
            # Fallback: paginate with Next page if rows-per-page did not expand fully.
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
                    # Page did not advance
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
    # Stable sort by PHS/id
    studies.sort(key=lambda s: (s.get("phs") or "").lower())

    ended = datetime.now(timezone.utc)
    meta = {
        "source_url": STUDIES_URL,
        "collected_at": ended.strftime("%Y-%m-%d %H:%M UTC"),
        "browser": "Chromium (Playwright)",
        "listing_count": len(listing),
        "studies_collected": len(studies),
        "detail_ok_count": sum(1 for s in studies if s.get("detail_ok")),
        "failures": [
            {"phs": s.get("phs"), "error": s.get("error")}
            for s in studies
            if s.get("error") or not s.get("detail_ok")
        ],
        "duration_sec": round((ended - started).total_seconds(), 1),
        "method": "playwright-ui",
    }

    json_path = OUTPUT_DIR / "studies.json"
    html_path = OUTPUT_DIR / "gc-studies-playwright-report.html"
    json_path.write_text(json.dumps({"meta": meta, "studies": studies}, indent=2), encoding="utf-8")
    html_path.write_text(generate_html(studies, meta), encoding="utf-8")

    log(f"Wrote {json_path}")
    log(f"Wrote {html_path}")
    log(
        f"Done: {len(studies)} studies "
        f"({meta['detail_ok_count']} detail OK, {len(meta['failures'])} failures) "
        f"in {meta['duration_sec']}s"
    )
    # Non-zero only if we collected nothing; partial failures stay in the report.
    return 0 if studies else 1


if __name__ == "__main__":
    sys.exit(main())
