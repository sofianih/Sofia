#!/usr/bin/env python3
"""Headed GC Studies scrape → Excel compare → HTML report."""

from __future__ import annotations

import html
import json
import os
import re
import sys
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd
from playwright.sync_api import sync_playwright

DIR = Path(__file__).resolve().parent
URL = os.environ.get("GC_STUDIES_URL", "https://general.datacommons.cancer.gov/#/studies")
OUT = Path(os.environ.get("GC_OUTPUT_DIR", DIR / "output"))
HEADLESS = os.environ.get("HEADLESS", os.environ.get("PLAYWRIGHT_HEADLESS", "")).lower() in {
    "1", "true", "yes", "on",
}


def num(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, (int, float)):
        return int(v)
    s = re.sub(r"[^\d-]", "", str(v))
    return int(s) if s not in ("", "-") else None


def fmt(n):
    return f"{n:,}" if n is not None else "—"


def excel_path() -> Path:
    env = os.environ.get("GC_BASECOUNT_XLSX", "").strip()
    if env:
        return Path(env)
    for p in (DIR / "GC Studypage_Basecount.xlsx", DIR / "data" / "GC_Studypage_Basecount.xlsx"):
        if p.is_file():
            return p
    return DIR / "GC Studypage_Basecount.xlsx"


def load_excel(path: Path) -> dict[str, dict]:
    df = pd.read_excel(path, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    study = next(c for c in df.columns if c.lower() == "study")
    part = next(c for c in df.columns if "participant" in c.lower())
    files = next(c for c in df.columns if "file" in c.lower() or c.lower().startswith("nubmer"))
    out = {}
    for _, row in df.iterrows():
        phs = str(row[study]).strip() if pd.notna(row[study]) else ""
        if not phs or phs.lower() == "nan":
            continue
        out[phs.lower()] = {"phs": phs, "participants": num(row[part]), "files": num(row[files])}
    return out


def dismiss(page):
    btn = page.get_by_role("button", name=re.compile(r"^Continue$", re.I))
    if btn.count():
        try:
            btn.first.click(timeout=3000, force=True)
            page.wait_for_timeout(400)
        except Exception:
            pass


def scrape(page):
    page.goto(URL, wait_until="domcontentloaded", timeout=120000)
    dismiss(page)
    page.locator("table tbody tr").first.wait_for(timeout=60000)
    page.locator("[id^=statsbar_title_]").first.wait_for(timeout=30000)

    statbar = page.evaluate(
        """() => {
          const o = {};
          for (const el of document.querySelectorAll('[id^=statsbar_title_]')) {
            const label = (el.innerText || '').trim().toLowerCase();
            let count = null;
            const parent = el.parentElement;
            if (parent) {
              for (const s of parent.children) {
                const t = (s.innerText || '').trim();
                if (/^[\\d,]+$/.test(t)) { count = t; break; }
              }
            }
            o[label] = count;
          }
          return o;
        }"""
    )
    statbar = {k: num(v) for k, v in statbar.items()}

    # Show all rows
    page.locator('[aria-label="rows per page dropdown selector"]').first.click()
    page.get_by_role("option", name="100", exact=True).click()
    page.wait_for_timeout(800)

    rows = page.evaluate(
        """() => [...document.querySelectorAll('table tbody tr')].map(tr => {
          const c = [...tr.querySelectorAll('td')].map(td => (td.innerText || '').trim());
          const a = tr.querySelector('td a');
          return {
            phs: c[0] || '',
            study_name: c[2] || '',
            participants: c[6] || '',
            files: c[7] || '',
            href: a ? a.getAttribute('href') || '' : '',
          };
        })"""
    )
    studies = []
    for r in rows:
        phs = (r.get("phs") or "").strip()
        if not phs:
            continue
        href = r.get("href") or ""
        detail = (URL.split("#")[0] + href) if href.startswith("#/") else f"{URL.split('#')[0]}#/study/{phs}"
        studies.append(
            {
                "phs": phs,
                "study_name": r.get("study_name") or "",
                "participants": num(r.get("participants")),
                "files": num(r.get("files")),
                "detail_url": detail,
            }
        )
    return statbar, studies


def compare(studies, baseline):
    seen = set()
    out = []
    for s in studies:
        key = s["phs"].lower()
        seen.add(key)
        b = baseline.get(key)
        if not b:
            status, detail = "MISSING_IN_BASE", "not in Excel"
            ep = ef = None
            pm = fm = False
        else:
            ep, ef = b["participants"], b["files"]
            pm, fm = s["participants"] == ep, s["files"] == ef
            status = "PASS" if pm and fm else "FAIL"
            detail = "" if status == "PASS" else f"part UI={s['participants']} Excel={ep}; files UI={s['files']} Excel={ef}"
        out.append({**s, "excel_participants": ep, "excel_files": ef, "status": status, "detail": detail})
    for key, b in baseline.items():
        if key in seen:
            continue
        out.append(
            {
                "phs": b["phs"],
                "study_name": "",
                "participants": None,
                "files": None,
                "detail_url": "",
                "excel_participants": b["participants"],
                "excel_files": b["files"],
                "status": "MISSING_IN_UI",
                "detail": "in Excel only",
            }
        )
    out.sort(key=lambda x: x["phs"].lower())
    return out


def report_html(statbar, studies, meta):
    passes = sum(1 for s in studies if s["status"] == "PASS")
    fails = len(studies) - passes
    tpart = sum(s["participants"] or 0 for s in studies if s["participants"] is not None)
    tfiles = sum(s["files"] or 0 for s in studies if s["files"] is not None)
    rows = []
    for i, s in enumerate(studies, 1):
        cls = {"PASS": "ok", "FAIL": "bad", "MISSING_IN_BASE": "warn", "MISSING_IN_UI": "warn"}.get(s["status"], "")
        rows.append(
            f"<tr class='{cls}'><td>{i}</td>"
            f"<td><a href='{html.escape(s['detail_url'] or '#', quote=True)}'>{html.escape(s['phs'])}</a></td>"
            f"<td>{html.escape(s.get('study_name') or '')}</td>"
            f"<td class='n'>{fmt(s['participants'])}</td><td class='n'>{fmt(s['excel_participants'])}</td>"
            f"<td class='n'>{fmt(s['files'])}</td><td class='n'>{fmt(s['excel_files'])}</td>"
            f"<td class='{cls}'>{html.escape(s['status'])}"
            f"{('<div class=m>' + html.escape(s['detail']) + '</div>') if s.get('detail') else ''}</td></tr>"
        )
    return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"/><title>GC Studies Report</title>
<style>
body{{font-family:Segoe UI,Arial,sans-serif;margin:24px;background:#f4f8fb;color:#0f1c2e}}
.bar{{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;background:#d7ebf5;border:2px solid #7eb6d0;border-radius:12px;padding:16px;margin:16px 0}}
.card{{background:#fff;border:1px solid #d7e0ea;border-radius:10px;padding:14px}}
.card b{{display:block;font-size:1.6rem;color:#084d6b}}
.sec{{display:flex;gap:12px;flex-wrap:wrap;margin:12px 0 20px;color:#5a6a7e;font-size:.9rem}}
table{{width:100%;border-collapse:collapse;background:#fff}}
th,td{{border-bottom:1px solid #d7e0ea;padding:8px;text-align:left;font-size:.9rem;vertical-align:top}}
th{{background:#f7fafc;position:sticky;top:0}} .n{{text-align:right;font-variant-numeric:tabular-nums}}
.ok{{color:#0f7a4a;font-weight:700}} .bad{{color:#a11;font-weight:700;background:#fff5f5}}
.warn{{color:#8a4b12;font-weight:700}} .m{{font-weight:400;color:#5a6a7e;font-size:.8rem}}
a{{color:#0b6e99;text-decoration:none}}
</style></head><body>
<h1>GC Studies — Playwright report</h1>
<p>Source: <a href="{html.escape(URL)}">{html.escape(URL)}</a> ·
Baseline: <code>{html.escape(Path(meta['excel']).name)}</code> ·
{html.escape(meta['collected_at'])} · {'headed' if not meta['headless'] else 'headless'}</p>
<div class="bar">
  <div class="card">Statbar · Studies<b>{fmt(statbar.get('studies'))}</b></div>
  <div class="card">Statbar · Participants<b>{fmt(statbar.get('participants'))}</b></div>
  <div class="card">Statbar · Samples<b>{fmt(statbar.get('samples'))}</b></div>
  <div class="card">Statbar · Files<b>{fmt(statbar.get('files'))}</b></div>
</div>
<p>Primary totals from page statbar widgets (<code>#statsbar_title_*</code>), not table sums.</p>
<div class="sec">
  <span>PASS: <b class="ok">{passes}</b></span>
  <span>FAIL/gaps: <b class="bad">{fails}</b></span>
  <span>Table sum · Participants: {fmt(tpart)} (secondary)</span>
  <span>Table sum · Files: {fmt(tfiles)} (secondary)</span>
</div>
<table><thead><tr>
<th>#</th><th>Study</th><th>Name</th><th class="n">UI Part.</th><th class="n">Excel Part.</th>
<th class="n">UI Files</th><th class="n">Excel Files</th><th>Status</th>
</tr></thead><tbody>{''.join(rows)}</tbody></table>
</body></html>"""


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    xlsx = excel_path()
    print(f"URL={URL}\nExcel={xlsx}\nHeadless={HEADLESS}", flush=True)
    baseline = load_excel(xlsx)
    print(f"Excel studies={len(baseline)}", flush=True)

    with sync_playwright() as p:
        try:
            browser = p.chromium.launch(headless=HEADLESS)
            actual_headless = HEADLESS
        except Exception as e:
            if HEADLESS:
                raise
            print(f"Headed failed ({e}); using headless", flush=True)
            browser = p.chromium.launch(headless=True)
            actual_headless = True
        page = browser.new_page(viewport={"width": 1440, "height": 900})
        statbar, listing = scrape(page)
        browser.close()

    print(
        f"Statbar Studies={statbar.get('studies')} Participants={statbar.get('participants')} "
        f"Samples={statbar.get('samples')} Files={statbar.get('files')}",
        flush=True,
    )
    studies = compare(listing, baseline)
    meta = {
        "source_url": URL,
        "excel": str(xlsx),
        "collected_at": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC"),
        "headless": actual_headless,
        "statbar": statbar,
        "pass": sum(1 for s in studies if s["status"] == "PASS"),
        "fail": sum(1 for s in studies if s["status"] != "PASS"),
    }
    (OUT / "studies.json").write_text(json.dumps({"meta": meta, "studies": studies}, indent=2), encoding="utf-8")
    html_path = OUT / "gc-studies-playwright-report.html"
    html_path.write_text(report_html(statbar, studies, meta), encoding="utf-8")
    print(f"Wrote {html_path}\nPASS={meta['pass']} FAIL={meta['fail']}", flush=True)
    return 0 if listing else 1


if __name__ == "__main__":
    sys.exit(main())
