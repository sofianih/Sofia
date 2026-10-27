#!/usr/bin/env python3
"""GC Studies scrape → Excel compare + column sorting → one HTML report."""
import html, json, os, re, sys
from datetime import datetime, timezone
from pathlib import Path
import pandas as pd
from playwright.sync_api import sync_playwright

DIR = Path(__file__).resolve().parent
DEFAULT_URL = "https://general.datacommons.cancer.gov/#/studies"
OUT = Path(os.environ.get("GC_OUTPUT_DIR", DIR / "output"))

COLUMN_TYPES = {
    0: "text", 1: "text", 2: "text", 3: "text", 4: "text", 5: "text",
    6: "numeric", 7: "numeric",
}


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


def resolve_headless():
    """Default headed. CLI --headless / --headed wins; else HEADLESS / PLAYWRIGHT_HEADLESS."""
    args = {a.lower() for a in sys.argv[1:]}
    if "--headed" in args or "-headed" in args:
        return False
    if "--headless" in args or "-headless" in args:
        return True
    raw = (os.environ.get("HEADLESS") or os.environ.get("PLAYWRIGHT_HEADLESS") or "").strip().lower()
    if raw in {"0", "false", "no", "off", "headed"}:
        return False
    if raw in {"1", "true", "yes", "on", "headless"}:
        return True
    return False  # local default: headed


URL = resolve_url()
HEADLESS = resolve_headless()


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
    for p in (DIR / "GC Studypage_Basecount.xlsx", DIR / "data" / "GC Studypage_Basecount.xlsx"):
        if p.is_file():
            return p
    return DIR / "GC Studypage_Basecount.xlsx"


def sorting_baseline_path():
    env = os.environ.get("GC_SORTING_BASELINE", "").strip()
    if env:
        return Path(env)
    for p in (DIR / "sorting_baseline.json", DIR / "data" / "sorting_baseline.json"):
        if p.is_file():
            return p
    return DIR / "sorting_baseline.json"


def load_sorting_baseline(path):
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    default = data.get("default") or {}
    cols = {}
    for c in data.get("columns") or []:
        name = (c.get("name") or "").strip()
        if name:
            cols[name] = {
                "name": name,
                "sortable": bool(c.get("sortable", True)),
                "ascending": str(c.get("ascending", "PASS")).upper(),
                "descending": str(c.get("descending", "PASS")).upper(),
            }
    return {
        "default_name": (default.get("name") or "").strip(),
        "default_direction": (default.get("direction") or "asc").strip().lower(),
        "columns": cols,
        "path": str(path),
    }


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


def open_page(page):
    page.goto(URL, wait_until="domcontentloaded", timeout=120000)
    btn = page.get_by_role("button", name=re.compile(r"^Continue$", re.I))
    if btn.count():
        try:
            btn.first.click(timeout=3000, force=True)
            page.wait_for_timeout(400)
        except Exception:
            pass
    page.locator("table tbody tr").first.wait_for(timeout=60000)
    page.locator("[id^=statsbar_title_]").first.wait_for(timeout=30000)
    page.wait_for_timeout(500)


def set_rows_100(page):
    dropdown = page.locator('[aria-label="rows per page dropdown selector"]')
    if dropdown.count():
        dropdown.first.click()
        page.get_by_role("option", name="100", exact=True).click()
        page.wait_for_timeout(800)


def read_statbar(page):
    js_bar = """() => { const o={}; for (const el of document.querySelectorAll('[id^=statsbar_title_]')) {
      const label=(el.innerText||'').trim().toLowerCase(); let count=null;
      if (el.parentElement) for (const s of el.parentElement.children) {
        const t=(s.innerText||'').trim(); if (/^[\\d,]+$/.test(t)) { count=t; break; } }
      o[label]=count; } return o; }"""
    return {k: num(v) for k, v in page.evaluate(js_bar).items()}


def scrape_listing(page):
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
        studies.append({
            "phs": phs, "study_name": r.get("name") or "",
            "participants": num(r.get("participants")), "files": num(r.get("files")),
            "detail_url": (base + href) if href.startswith("#/") else f"{base}#/study/{phs}",
        })
    return studies


def compare(studies, baseline):
    seen, out = set(), []
    for s in studies:
        key = s["phs"].lower()
        seen.add(key)
        b = baseline.get(key)
        if not b:
            ep = ef = None
            status, detail = "MISSING_IN_BASE", "not in Excel"
        else:
            ep, ef = b["participants"], b["files"]
            status = "PASS" if s["participants"] == ep and s["files"] == ef else "FAIL"
            detail = "" if status == "PASS" else f"part {s['participants']}/{ep}; files {s['files']}/{ef}"
        out.append({**s, "excel_participants": ep, "excel_files": ef, "status": status, "detail": detail})
    for key, b in baseline.items():
        if key not in seen:
            out.append({
                "phs": b["phs"], "study_name": "", "participants": None, "files": None, "detail_url": "",
                "excel_participants": b["participants"], "excel_files": b["files"],
                "status": "MISSING_IN_UI", "detail": "in Excel only",
            })
    return sorted(out, key=lambda x: x["phs"].lower())


# --- sorting ---

def header_state(page):
    return page.evaluate("""() => {
      const ths = [...document.querySelectorAll('table thead th')];
      return ths.map((th, i) => {
        const label = th.querySelector('[class*="MuiTableSortLabel-root"]') || th.querySelector('[role="button"]');
        const cls = label ? (label.className || '') : '';
        const svg = label ? label.querySelector('svg') : null;
        const svgCls = svg ? (svg.getAttribute('class') || '') : '';
        let direction = null;
        const active = /MuiTableSortLabel-active/.test(cls);
        if (active) {
          if (/iconDirectionDesc/.test(svgCls) || /DirectionDesc/.test(svgCls)) direction = 'desc';
          else if (/iconDirectionAsc/.test(svgCls) || /DirectionAsc/.test(svgCls)) direction = 'asc';
          else direction = (th.getAttribute('direction') || 'asc').toLowerCase();
        }
        const text = ((label && label.childNodes[0] && label.childNodes[0].textContent) || th.innerText || '').trim();
        return {
          index: i, name: text,
          fieldClass: (th.className || '').split(/\\s+/).find(c => !c.startsWith('Mui')) || '',
          sortable: !!(label && /Sort by/i.test(label.getAttribute('title') || '')),
          active, direction,
        };
      });
    }""")


def row_values(page, col_index):
    return page.evaluate("""(col) => [...document.querySelectorAll('table tbody tr')].map(tr => {
      const td = tr.querySelectorAll('td')[col];
      if (!td) return '';
      const raw = (td.innerText || '').trim();
      return raw.split(/\\n+/).map(s => s.trim()).filter(Boolean)[0] || '';
    }).filter((_, i) => {
      const tds = document.querySelectorAll('table tbody tr')[i]?.querySelectorAll('td');
      return tds && (tds[0].innerText || '').trim();
    })""", col_index)


def parse_numeric(s):
    if s is None:
        return None
    cleaned = re.sub(r"[^\d.-]", "", str(s))
    if cleaned in ("", "-", "."):
        return None
    try:
        return float(cleaned)
    except ValueError:
        return None


def is_sorted(page, values, col_type, direction):
    if not values:
        return False
    if col_type == "numeric":
        nums = [parse_numeric(v) for v in values]
        if any(n is None for n in nums):
            return False
        ordered = sorted(nums)
        if direction == "desc":
            ordered = list(reversed(ordered))
        return nums == ordered
    expected = page.evaluate(
        """([arr, dir]) => {
          const copy = [...arr].sort((a, b) => String(a).localeCompare(String(b), undefined, {sensitivity: 'base'}));
          return dir === 'desc' ? copy.reverse() : copy;
        }""",
        [values, direction],
    )
    return values == expected


def click_sort(page, col_index):
    label = page.locator("table thead th").nth(col_index).locator(
        '[role="button"], [class*="MuiTableSortLabel-root"]'
    )
    prev = header_state(page)
    prev_vals = row_values(page, col_index)
    prev_sig = (
        prev[col_index].get("active"), prev[col_index].get("direction"),
        tuple(prev_vals[:5]), tuple(prev_vals[-3:]),
    )
    label.first.click()
    for _ in range(40):
        page.wait_for_timeout(150)
        state = header_state(page)
        vals = row_values(page, col_index)
        sig = (
            state[col_index].get("active"), state[col_index].get("direction"),
            tuple(vals[:5]), tuple(vals[-3:]),
        )
        if sig != prev_sig:
            page.wait_for_timeout(200)
            return
    page.wait_for_timeout(300)


def verify_column(page, col, col_type):
    result = {
        "index": col["index"], "name": col["name"], "type": col_type,
        "sortable_ui": col.get("sortable", False), "clicks": [], "status": "PASS",
    }
    if not col.get("sortable"):
        before = row_values(page, col["index"])
        try:
            click_sort(page, col["index"])
        except Exception:
            result["status"] = "NOT_SORTABLE"
            return result
        after_state = header_state(page)[col["index"]]
        after = row_values(page, col["index"])
        if not after_state.get("active") and after == before:
            result["status"] = "NOT_SORTABLE"
            return result

    ok_all = True
    for click_n in (1, 2):
        before_state = header_state(page)[col["index"]]
        before_vals = row_values(page, col["index"])
        click_sort(page, col["index"])
        state = header_state(page)[col["index"]]
        vals = row_values(page, col["index"])
        direction = state.get("direction")
        order_ok = bool(state.get("active") and direction in ("asc", "desc")
                        and is_sorted(page, vals, col_type, direction))
        if click_n == 2 and before_state.get("active") and before_state.get("direction"):
            expected_toggle = "desc" if before_state["direction"] == "asc" else "asc"
            if direction != expected_toggle:
                order_ok = False
            if (
                len(set(str(v).casefold() if col_type == "text" else v for v in vals)) > 1
                and before_vals == vals
            ):
                order_ok = False
        result["clicks"].append({
            "click": click_n, "ui_direction": direction,
            "sample_first3": vals[:3], "order_ok": order_ok,
        })
        if not order_ok:
            ok_all = False
    result["status"] = "PASS" if ok_all else "FAIL"
    return result


def verify_sorting(page):
    """Capture default sort, then check each column asc/desc. Call after rows=100."""
    columns = header_state(page)
    default = next((c for c in columns if c.get("active")), None) or {}
    results = []
    for col in columns:
        ctype = COLUMN_TYPES.get(col["index"], "text")
        print(f"  Sorting [{col['index']}] {col['name']}…", flush=True)
        r = verify_column(page, col, ctype)
        print(f"    → {r['status']}", flush=True)
        results.append(r)
    return {"default": default, "results": results}


def _live_dir_status(result, direction):
    if result.get("status") == "NOT_SORTABLE" or not result.get("sortable_ui"):
        return "—"
    click = next((c for c in result.get("clicks") or [] if c.get("ui_direction") == direction), None)
    if not click:
        return "FAIL"
    return "PASS" if click.get("order_ok") else "FAIL"


def compare_sorting(live, baseline):
    """Compare live sort results to sorting_baseline.json (like Excel basecount)."""
    default = live.get("default") or {}
    live_name = (default.get("name") or "").strip()
    live_dir = (default.get("direction") or "").strip().lower()
    exp_name = baseline["default_name"]
    exp_dir = baseline["default_direction"]
    default_ok = (
        live_name.casefold() == exp_name.casefold()
        and live_dir == exp_dir
    )
    default_cmp = {
        "live_name": live_name or "—",
        "live_direction": live_dir or "—",
        "expected_name": exp_name or "—",
        "expected_direction": exp_dir or "—",
        "status": "PASS" if default_ok else "FAIL",
    }

    by_name = {r["name"]: r for r in live.get("results") or []}
    rows = []
    seen = set()
    for name, exp in baseline["columns"].items():
        seen.add(name.casefold())
        live_r = by_name.get(name)
        if not live_r:
            # try case-insensitive match
            live_r = next((r for r in by_name.values() if r["name"].casefold() == name.casefold()), None)
        if not live_r:
            rows.append({
                "name": name,
                "sortable_live": "—", "sortable_expected": "yes" if exp["sortable"] else "no",
                "asc_live": "—", "desc_live": "—",
                "asc_expected": exp["ascending"], "desc_expected": exp["descending"],
                "sample_asc": [], "sample_desc": [],
                "status": "MISSING_IN_UI", "detail": "in baseline only",
            })
            continue
        sortable_live = live_r.get("sortable_ui") and live_r.get("status") != "NOT_SORTABLE"
        asc_live = _live_dir_status(live_r, "asc")
        desc_live = _live_dir_status(live_r, "desc")
        asc_click = next((c for c in live_r.get("clicks") or [] if c.get("ui_direction") == "asc"), None)
        desc_click = next((c for c in live_r.get("clicks") or [] if c.get("ui_direction") == "desc"), None)
        ok = (
            bool(sortable_live) == bool(exp["sortable"])
            and asc_live == exp["ascending"]
            and desc_live == exp["descending"]
        )
        rows.append({
            "name": name,
            "sortable_live": "yes" if sortable_live else "no",
            "sortable_expected": "yes" if exp["sortable"] else "no",
            "asc_live": asc_live, "desc_live": desc_live,
            "asc_expected": exp["ascending"], "desc_expected": exp["descending"],
            "sample_asc": (asc_click or {}).get("sample_first3") or [],
            "sample_desc": (desc_click or {}).get("sample_first3") or [],
            "status": "PASS" if ok else "FAIL",
            "detail": "" if ok else (
                f"live sortable={('yes' if sortable_live else 'no')}/"
                f"{'yes' if exp['sortable'] else 'no'}; "
                f"asc {asc_live}/{exp['ascending']}; desc {desc_live}/{exp['descending']}"
            ),
        })

    for r in live.get("results") or []:
        if r["name"].casefold() in seen:
            continue
        sortable_live = r.get("sortable_ui") and r.get("status") != "NOT_SORTABLE"
        rows.append({
            "name": r["name"],
            "sortable_live": "yes" if sortable_live else "no",
            "sortable_expected": "—",
            "asc_live": _live_dir_status(r, "asc"),
            "desc_live": _live_dir_status(r, "desc"),
            "asc_expected": "—", "desc_expected": "—",
            "sample_asc": [], "sample_desc": [],
            "status": "MISSING_IN_BASE", "detail": "not in sorting baseline",
        })

    return {"default": default_cmp, "columns": rows}


# --- report ---

CSS = """body{font-family:Segoe UI,Arial,sans-serif;margin:24px;background:#f4f8fb;color:#0f1c2e}
.bar{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;background:#d7ebf5;border:2px solid #7eb6d0;border-radius:12px;padding:16px;margin:16px 0}
.card{background:#fff;border:1px solid #d7e0ea;border-radius:10px;padding:14px}.card b{display:block;font-size:1.6rem;color:#084d6b}
.sec{display:flex;gap:12px;flex-wrap:wrap;margin:12px 0 20px;color:#5a6a7e;font-size:.9rem}
h2{margin:28px 0 8px;font-size:1.15rem}
.default{margin:0 0 12px;font-size:.95rem}
table{width:100%;border-collapse:collapse;background:#fff;margin-bottom:8px}
th,td{border-bottom:1px solid #d7e0ea;padding:8px;text-align:left;font-size:.9rem;vertical-align:top}
th{background:#f7fafc;position:sticky;top:0}.n{text-align:right;font-variant-numeric:tabular-nums}
.ok{color:#0f7a4a;font-weight:700}.bad{color:#a11;font-weight:700;background:#fff5f5}.warn{color:#8a4b12;font-weight:700}.na{color:#5a6a7e}
.m{font-weight:400;color:#5a6a7e;font-size:.8rem}
.sample{display:block;margin-top:3px;color:#5a6a7e;font-size:.8rem;font-weight:400}
a{color:#0b6e99;text-decoration:none}"""


def _status_span(status, sample=None):
    if status in (None, "—"):
        return "<span class=na>—</span>"
    cls = {"PASS": "ok", "FAIL": "bad"}.get(status, "warn")
    sample_html = ""
    if sample:
        shown = ", ".join(html.escape(str(v)) for v in sample[:3])
        sample_html = f"<span class=sample>{shown}</span>"
    return f"<span class={cls}>{html.escape(status)}</span>{sample_html}"


def sorting_section(sort_cmp, meta):
    d = sort_cmp.get("default") or {}
    dcls = "ok" if d.get("status") == "PASS" else "bad"
    live = f"{html.escape(d.get('live_name') or '—')} · {html.escape((d.get('live_direction') or '—').upper())}"
    exp = f"{html.escape(d.get('expected_name') or '—')} · {html.escape((d.get('expected_direction') or '—').upper())}"
    rows = []
    for r in sort_cmp.get("columns") or []:
        cls = {"PASS": "ok", "FAIL": "bad"}.get(r["status"], "warn")
        detail = f"<div class=m>{html.escape(r['detail'])}</div>" if r.get("detail") else ""
        rows.append(
            f"<tr class={cls}><td>{html.escape(r['name'])}</td>"
            f"<td>{html.escape(r['sortable_live'])}</td>"
            f"<td>{_status_span(r['asc_live'], r.get('sample_asc'))}</td>"
            f"<td>{_status_span(r['desc_live'], r.get('sample_desc'))}</td>"
            f"<td class={cls}>{r['status']}{detail}</td></tr>"
        )
    sp = meta.get("sort_pass", 0)
    sf = meta.get("sort_fail", 0)
    return f"""<h2>Sorting vs baseline</h2>
<p class=default><b>Default sort:</b> {live}
<span class={dcls}> {d.get('status', '—')}</span>
<span class=m> (expected {exp} · <code>{html.escape(Path(meta.get('sorting_baseline') or '').name)}</code>)</span></p>
<div class=sec><span>Sort PASS: <b class=ok>{sp}</b></span><span>Sort FAIL/gaps: <b class=bad>{sf}</b></span></div>
<table><thead><tr><th>Column</th><th>Sortable?</th><th>Ascending</th><th>Descending</th><th>vs baseline</th></tr></thead>
<tbody>{''.join(rows)}</tbody></table>"""


def report_html(statbar, studies, sort_cmp, meta):
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
            f"<td class=n>{fmt(s['excel_files'])}</td><td class={cls}>{s['status']}{d}</td></tr>"
        )
    cards = "".join(
        f"<div class=card>Statbar · {lab}<b>{fmt(statbar.get(k))}</b></div>"
        for lab, k in (("Studies", "studies"), ("Participants", "participants"),
                       ("Samples", "samples"), ("Files", "files"))
    )
    return f"""<!DOCTYPE html><html><head><meta charset=utf-8><title>GC Studies Report</title>
<style>{CSS}</style></head><body>
<h1>GC Studies — Playwright report</h1>
<p>Source: <a href="{html.escape(URL)}">{html.escape(URL)}</a> ·
Excel: <code>{html.escape(Path(meta['excel']).name)}</code> ·
Sort baseline: <code>{html.escape(Path(meta.get('sorting_baseline') or '').name)}</code> ·
{html.escape(meta['collected_at'])} · {'headed' if not meta['headless'] else 'headless'}</p>
<div class=bar>{cards}</div>
<div class=sec><span>Excel PASS: <b class=ok>{passes}</b></span><span>Excel FAIL/gaps: <b class=bad>{len(studies)-passes}</b></span>
<span>Table sum · Participants: {fmt(tpart)} (secondary)</span><span>Table sum · Files: {fmt(tfiles)} (secondary)</span></div>
{sorting_section(sort_cmp, meta)}
<h2>Studies vs Excel</h2>
<table><thead><tr><th>#</th><th>Study</th><th>Name</th><th class=n>UI Part.</th><th class=n>Excel Part.</th>
<th class=n>UI Files</th><th class=n>Excel Files</th><th>Status</th></tr></thead>
<tbody>{''.join(rows)}</tbody></table></body></html>"""


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    xlsx = excel_path()
    sort_base_path = sorting_baseline_path()
    print(f"URL={URL}\nExcel={xlsx}\nSortBaseline={sort_base_path}\nHeadless={HEADLESS}", flush=True)
    baseline = load_excel(xlsx)
    sort_baseline = load_sorting_baseline(sort_base_path)
    print(f"Excel studies={len(baseline)} · sort baseline columns={len(sort_baseline['columns'])}", flush=True)
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
        page = browser.new_page(viewport={"width": 1600, "height": 1000})
        open_page(page)
        # Default sort before any header clicks
        default_cols = header_state(page)
        default_sort = next((c for c in default_cols if c.get("active")), None) or {}
        print(f"Default sort: {default_sort.get('name')} {default_sort.get('direction')}", flush=True)
        set_rows_100(page)
        statbar = read_statbar(page)
        listing = scrape_listing(page)
        print(f"Statbar {statbar} · listing={len(listing)}", flush=True)
        sorting = verify_sorting(page)
        sorting["default"] = default_sort  # preserve pre-click default
        browser.close()

    studies = compare(listing, baseline)
    sort_cmp = compare_sorting(sorting, sort_baseline)
    sort_col_pass = sum(1 for r in sort_cmp["columns"] if r["status"] == "PASS")
    sort_col_fail = sum(1 for r in sort_cmp["columns"] if r["status"] != "PASS")
    if sort_cmp["default"]["status"] != "PASS":
        sort_col_fail += 1
    meta = {
        "excel": str(xlsx),
        "sorting_baseline": str(sort_base_path),
        "collected_at": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC"),
        "headless": actual_headless,
        "statbar": statbar,
        "pass": sum(1 for s in studies if s["status"] == "PASS"),
        "fail": sum(1 for s in studies if s["status"] != "PASS"),
        "sort_pass": sort_col_pass,
        "sort_fail": sort_col_fail,
        "default_sort": sort_cmp["default"],
    }
    (OUT / "studies.json").write_text(
        json.dumps({
            "meta": meta, "studies": studies,
            "sorting_live": sorting, "sorting_vs_baseline": sort_cmp,
        }, indent=2),
        encoding="utf-8",
    )
    path = OUT / "gc-studies-playwright-report.html"
    path.write_text(report_html(statbar, studies, sort_cmp, meta), encoding="utf-8")
    print(
        f"Wrote {path}\nExcel PASS={meta['pass']} FAIL={meta['fail']} · "
        f"Sort vs baseline PASS={meta['sort_pass']} FAIL={meta['sort_fail']} · "
        f"default={sort_cmp['default']['status']}",
        flush=True,
    )
    return 0 if listing and meta["sort_fail"] == 0 and meta["fail"] == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
