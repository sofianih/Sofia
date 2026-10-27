#!/usr/bin/env python3
"""Headed GC Studies scrape → Excel compare → HTML report."""
import html, json, os, re, sys
from datetime import datetime, timezone
from pathlib import Path
import pandas as pd
from playwright.sync_api import sync_playwright

DIR = Path(__file__).resolve().parent
DEFAULT_URL = "https://general.datacommons.cancer.gov/#/studies"
OUT = Path(os.environ.get("GC_OUTPUT_DIR", DIR / "output"))
HEADLESS = os.environ.get("HEADLESS", os.environ.get("PLAYWRIGHT_HEADLESS", "")).lower() in {"1", "true", "yes", "on"}


def resolve_url():
    """CLI (--url=… / --url … / bare https…) then GC_STUDIES_URL / BASE_URL, else prod."""
    args = sys.argv[1:]
    for i, a in enumerate(args):
        if a.startswith("--url="):
            return a.split("=", 1)[1].strip() or DEFAULT_URL
        if a in ("--url", "-u") and i + 1 < len(args):
            return args[i + 1].strip() or DEFAULT_URL
        if a.startswith("http"):
            return a.strip()
    return (os.environ.get("GC_STUDIES_URL") or os.environ.get("BASE_URL") or DEFAULT_URL).strip() or DEFAULT_URL


URL = resolve_url()


def num(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, (int, float)):
        return int(v)
    s = re.sub(r"[^\d-]", "", str(v))
    return int(s) if s not in ("", "-") else None


def fmt(n):
    return f"{n:,}" if n is not None else "—"


def excel_path():
    env = os.environ.get("GC_BASECOUNT_XLSX", "").strip()
    if env:
        return Path(env)
    for p in (DIR / "GC Studypage_Basecount.xlsx", DIR / "data" / "GC_Studypage_Basecount.xlsx"):
        if p.is_file():
            return p
    return DIR / "GC Studypage_Basecount.xlsx"


def load_excel(path):
    df = pd.read_excel(path, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    study = next(c for c in df.columns if c.lower() == "study")
    part = next(c for c in df.columns if "participant" in c.lower())
    files = next(c for c in df.columns if "file" in c.lower() or c.lower().startswith("nubmer"))
    out = {}
    for _, row in df.iterrows():
        phs = str(row[study]).strip() if pd.notna(row[study]) else ""
        if phs and phs.lower() != "nan":
            out[phs.lower()] = {"phs": phs, "participants": num(row[part]), "files": num(row[files])}
    return out


def scrape(page):
    page.goto(URL, wait_until="domcontentloaded", timeout=120000)
    btn = page.get_by_role("button", name=re.compile(r"^Continue$", re.I))
    if btn.count():
        try:
            btn.first.click(timeout=3000, force=True); page.wait_for_timeout(400)
        except Exception:
            pass
    page.locator("table tbody tr").first.wait_for(timeout=60000)
    page.locator("[id^=statsbar_title_]").first.wait_for(timeout=30000)
    js_bar = """() => { const o={}; for (const el of document.querySelectorAll('[id^=statsbar_title_]')) {
      const label=(el.innerText||'').trim().toLowerCase(); let count=null;
      if (el.parentElement) for (const s of el.parentElement.children) {
        const t=(s.innerText||'').trim(); if (/^[\\d,]+$/.test(t)) { count=t; break; } }
      o[label]=count; } return o; }"""
    statbar = {k: num(v) for k, v in page.evaluate(js_bar).items()}
    page.locator('[aria-label="rows per page dropdown selector"]').first.click()
    page.get_by_role("option", name="100", exact=True).click(); page.wait_for_timeout(800)
    base = URL.split("#")[0]
    js_rows = """() => [...document.querySelectorAll('table tbody tr')].map(tr => {
      const c=[...tr.querySelectorAll('td')].map(td => (td.innerText||'').trim());
      const a=tr.querySelector('td a');
      return {phs:c[0]||'', name:c[2]||'', participants:c[6]||'', files:c[7]||'', href:a?(a.getAttribute('href')||''):''}; })"""
    studies = []
    for r in page.evaluate(js_rows):
        phs = (r.get("phs") or "").strip()
        if not phs:
            continue
        href = r.get("href") or ""
        studies.append({"phs": phs, "study_name": r.get("name") or "", "participants": num(r.get("participants")),
                        "files": num(r.get("files")),
                        "detail_url": (base + href) if href.startswith("#/") else f"{base}#/study/{phs}"})
    return statbar, studies


def compare(studies, baseline):
    seen, out = set(), []
    for s in studies:
        key = s["phs"].lower(); seen.add(key); b = baseline.get(key)
        if not b:
            ep = ef = None; status, detail = "MISSING_IN_BASE", "not in Excel"
        else:
            ep, ef = b["participants"], b["files"]
            status = "PASS" if s["participants"] == ep and s["files"] == ef else "FAIL"
            detail = "" if status == "PASS" else f"part {s['participants']}/{ep}; files {s['files']}/{ef}"
        out.append({**s, "excel_participants": ep, "excel_files": ef, "status": status, "detail": detail})
    for key, b in baseline.items():
        if key not in seen:
            out.append({"phs": b["phs"], "study_name": "", "participants": None, "files": None, "detail_url": "",
                        "excel_participants": b["participants"], "excel_files": b["files"],
                        "status": "MISSING_IN_UI", "detail": "in Excel only"})
    return sorted(out, key=lambda x: x["phs"].lower())


CSS = """body{font-family:Segoe UI,Arial,sans-serif;margin:24px;background:#f4f8fb;color:#0f1c2e}
.bar{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;background:#d7ebf5;border:2px solid #7eb6d0;border-radius:12px;padding:16px;margin:16px 0}
.card{background:#fff;border:1px solid #d7e0ea;border-radius:10px;padding:14px}.card b{display:block;font-size:1.6rem;color:#084d6b}
.sec{display:flex;gap:12px;flex-wrap:wrap;margin:12px 0 20px;color:#5a6a7e;font-size:.9rem}
table{width:100%;border-collapse:collapse;background:#fff}th,td{border-bottom:1px solid #d7e0ea;padding:8px;text-align:left;font-size:.9rem;vertical-align:top}
th{background:#f7fafc;position:sticky;top:0}.n{text-align:right;font-variant-numeric:tabular-nums}
.ok{color:#0f7a4a;font-weight:700}.bad{color:#a11;font-weight:700;background:#fff5f5}.warn{color:#8a4b12;font-weight:700}
.m{font-weight:400;color:#5a6a7e;font-size:.8rem}a{color:#0b6e99;text-decoration:none}"""


def report_html(statbar, studies, meta):
    passes = sum(1 for s in studies if s["status"] == "PASS")
    tpart = sum(s["participants"] or 0 for s in studies if s["participants"] is not None)
    tfiles = sum(s["files"] or 0 for s in studies if s["files"] is not None)
    rows = []
    for i, s in enumerate(studies, 1):
        cls = {"PASS": "ok", "FAIL": "bad"}.get(s["status"], "warn")
        d = f"<div class=m>{html.escape(s['detail'])}</div>" if s.get("detail") else ""
        rows.append(
            f"<tr class={cls}><td>{i}</td><td><a href='{html.escape(s['detail_url'] or '#', True)}'>{html.escape(s['phs'])}</a></td>"
            f"<td>{html.escape(s.get('study_name') or '')}</td><td class=n>{fmt(s['participants'])}</td>"
            f"<td class=n>{fmt(s['excel_participants'])}</td><td class=n>{fmt(s['files'])}</td>"
            f"<td class=n>{fmt(s['excel_files'])}</td><td class={cls}>{s['status']}{d}</td></tr>")
    cards = "".join(f"<div class=card>Statbar · {lab}<b>{fmt(statbar.get(k))}</b></div>"
                    for lab, k in (("Studies", "studies"), ("Participants", "participants"),
                                   ("Samples", "samples"), ("Files", "files")))
    return f"""<!DOCTYPE html><html><head><meta charset=utf-8><title>GC Studies Report</title>
<style>{CSS}</style></head><body>
<h1>GC Studies — Playwright report</h1>
<p>Source: <a href="{html.escape(URL)}">{html.escape(URL)}</a> · Baseline: <code>{html.escape(Path(meta['excel']).name)}</code> ·
{html.escape(meta['collected_at'])} · {'headed' if not meta['headless'] else 'headless'}</p>
<div class=bar>{cards}</div>
<p>Primary totals from page statbar (<code>#statsbar_title_*</code>), not table sums.</p>
<div class=sec><span>PASS: <b class=ok>{passes}</b></span><span>FAIL/gaps: <b class=bad>{len(studies)-passes}</b></span>
<span>Table sum · Participants: {fmt(tpart)} (secondary)</span><span>Table sum · Files: {fmt(tfiles)} (secondary)</span></div>
<table><thead><tr><th>#</th><th>Study</th><th>Name</th><th class=n>UI Part.</th><th class=n>Excel Part.</th>
<th class=n>UI Files</th><th class=n>Excel Files</th><th>Status</th></tr></thead>
<tbody>{''.join(rows)}</tbody></table></body></html>"""


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    xlsx = excel_path()
    print(f"URL={URL}\nExcel={xlsx}\nHeadless={HEADLESS}", flush=True)
    baseline = load_excel(xlsx)
    print(f"Excel studies={len(baseline)}", flush=True)
    with sync_playwright() as p:
        try:
            browser = p.chromium.launch(headless=HEADLESS); actual_headless = HEADLESS
        except Exception as e:
            if HEADLESS:
                raise
            print(f"Headed failed ({e}); using headless", flush=True)
            browser = p.chromium.launch(headless=True); actual_headless = True
        page = browser.new_page(viewport={"width": 1440, "height": 900})
        statbar, listing = scrape(page); browser.close()
    print(f"Statbar {statbar}", flush=True)
    studies = compare(listing, baseline)
    meta = {"excel": str(xlsx), "collected_at": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC"),
            "headless": actual_headless, "statbar": statbar,
            "pass": sum(1 for s in studies if s["status"] == "PASS"),
            "fail": sum(1 for s in studies if s["status"] != "PASS")}
    (OUT / "studies.json").write_text(json.dumps({"meta": meta, "studies": studies}, indent=2), encoding="utf-8")
    path = OUT / "gc-studies-playwright-report.html"
    path.write_text(report_html(statbar, studies, meta), encoding="utf-8")
    print(f"Wrote {path}\nPASS={meta['pass']} FAIL={meta['fail']}", flush=True)
    return 0 if listing else 1


if __name__ == "__main__":
    sys.exit(main())
