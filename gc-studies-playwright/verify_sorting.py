#!/usr/bin/env python3
"""Verify GC Studies table default + per-column sorting → HTML report."""
import html, json, os, re, sys
from datetime import datetime, timezone
from pathlib import Path
from playwright.sync_api import sync_playwright

DIR = Path(__file__).resolve().parent
DEFAULT_URL = "https://general.datacommons.cancer.gov/#/studies"
OUT = Path(os.environ.get("GC_OUTPUT_DIR", DIR / "output"))

# Known column semantic types (index → text|numeric)
COLUMN_TYPES = {
    0: "text",   # Study
    1: "text",   # Study Version
    2: "text",   # Study Name
    3: "text",   # Program Name
    4: "text",   # Data Types
    5: "text",   # Study Access
    6: "numeric",  # Number of Participants
    7: "numeric",  # Number of Files
}


def resolve_url():
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
    return True  # CI-friendly default for sorting verify


URL = resolve_url()
HEADLESS = resolve_headless()


def open_studies(page):
    page.goto(URL, wait_until="domcontentloaded", timeout=120000)
    btn = page.get_by_role("button", name=re.compile(r"^Continue$", re.I))
    if btn.count():
        try:
            btn.first.click(timeout=3000, force=True)
            page.wait_for_timeout(400)
        except Exception:
            pass
    page.locator("table tbody tr").first.wait_for(timeout=60000)
    page.wait_for_timeout(800)
    # Show all rows so sort verification sees the full set
    dropdown = page.locator('[aria-label="rows per page dropdown selector"]')
    if dropdown.count():
        dropdown.first.click()
        page.get_by_role("option", name="100", exact=True).click()
        page.wait_for_timeout(900)


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
          index: i,
          name: text,
          fieldClass: (th.className || '').split(/\\s+/).find(c => !c.startsWith('Mui')) || '',
          sortable: !!(label && /Sort by/i.test(label.getAttribute('title') || '')),
          title: label ? (label.getAttribute('title') || '') : '',
          active,
          direction,
          thDirectionAttr: th.getAttribute('direction'),
          labelClass: cls,
          iconClass: svgCls,
        };
      });
    }""")


def row_values(page, col_index):
    return page.evaluate("""(col) => [...document.querySelectorAll('table tbody tr')].map(tr => {
      const td = tr.querySelectorAll('td')[col];
      if (!td) return '';
      // Prefer first non-empty line for multi-version cells
      const raw = (td.innerText || '').trim();
      const first = raw.split(/\\n+/).map(s => s.trim()).filter(Boolean)[0] || '';
      return first;
    }).filter((_, i, arr) => {
      // keep rows that have a Study id in col 0
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


def expected_text_order(page, values, direction):
    """Browser localeCompare — matches MUI/DataGrid string sort on this site."""
    ordered = page.evaluate(
        """([arr, dir]) => {
          const copy = [...arr].sort((a, b) => String(a).localeCompare(String(b), undefined, {sensitivity: 'base'}));
          return dir === 'desc' ? copy.reverse() : copy;
        }""",
        [values, direction],
    )
    return ordered


def is_sorted(page, values, col_type, direction):
    """Return (ok, detail). Numeric: parse commas. Text: JS localeCompare (site behavior)."""
    if not values:
        return False, "no values"
    if col_type == "numeric":
        nums = [parse_numeric(v) for v in values]
        if any(n is None for n in nums):
            return False, f"non-numeric values present: {[v for v,n in zip(values,nums) if n is None][:3]}"
        ordered = sorted(nums)
        if direction == "desc":
            ordered = list(reversed(ordered))
        ok = nums == ordered
        return ok, (
            f"first3={nums[:3]} last3={nums[-3:]}" if not ok
            else f"n={len(nums)} first={nums[0]} last={nums[-1]}"
        )
    expected = expected_text_order(page, values, direction)
    ok = values == expected
    if ok:
        return True, f"n={len(values)} first3={values[:3]}"
    mism = next((i for i in range(len(values)) if values[i] != expected[i]), None)
    return False, (
        f"mismatch@{mism}: got={values[mism]!r} expected={expected[mism]!r}; "
        f"first3={values[:3]}"
    )


def wait_sort_settle(page, col_index, prev_signature):
    """Wait until column values or active header change after click."""
    for _ in range(40):
        page.wait_for_timeout(150)
        state = header_state(page)
        vals = row_values(page, col_index)
        sig = (state[col_index].get("active"), state[col_index].get("direction"), tuple(vals[:5]), tuple(vals[-3:]))
        if sig != prev_signature:
            page.wait_for_timeout(200)
            return
    page.wait_for_timeout(300)


def click_sort(page, col_index):
    label = page.locator("table thead th").nth(col_index).locator('[role="button"], [class*="MuiTableSortLabel-root"]')
    prev = header_state(page)
    prev_vals = row_values(page, col_index)
    prev_sig = (prev[col_index].get("active"), prev[col_index].get("direction"), tuple(prev_vals[:5]), tuple(prev_vals[-3:]))
    label.first.click()
    wait_sort_settle(page, col_index, prev_sig)


def verify_column(page, col, col_type):
    """Click twice (asc/desc toggle) and verify order + UI indicator each time."""
    result = {
        "index": col["index"],
        "name": col["name"],
        "field": col.get("fieldClass") or "",
        "type": col_type,
        "sortable_ui": col.get("sortable", False),
        "clicks": [],
        "status": "PASS",
        "notes": [],
    }
    if not col.get("sortable"):
        # Attempt click anyway; if nothing happens, mark not sortable
        before = row_values(page, col["index"])
        try:
            click_sort(page, col["index"])
        except Exception as e:
            result["status"] = "NOT_SORTABLE"
            result["notes"].append(f"no sort control / click failed: {e}")
            return result
        after_state = header_state(page)[col["index"]]
        after = row_values(page, col["index"])
        if not after_state.get("active") and after == before:
            result["status"] = "NOT_SORTABLE"
            result["notes"].append("header present but sorting did not activate")
            return result

    # First click — becomes active (direction depends on previous active column)
    for click_n in (1, 2):
        before_state = header_state(page)[col["index"]]
        before_vals = row_values(page, col["index"])
        click_sort(page, col["index"])
        state = header_state(page)[col["index"]]
        vals = row_values(page, col["index"])
        direction = state.get("direction")
        click_rec = {
            "click": click_n,
            "ui_active": bool(state.get("active")),
            "ui_direction": direction,
            "row_count": len(vals),
            "sample_first3": vals[:3],
            "sample_last3": vals[-3:],
        }
        if not state.get("active") or direction not in ("asc", "desc"):
            click_rec["order_ok"] = False
            click_rec["detail"] = (
                f"expected active sort with asc/desc; "
                f"got active={state.get('active')} dir={direction}"
            )
            result["status"] = "FAIL"
        else:
            ok, detail = is_sorted(page, vals, col_type, direction)
            click_rec["order_ok"] = ok
            click_rec["detail"] = detail
            if not ok:
                result["status"] = "FAIL"
            if click_n == 2 and before_state.get("active") and before_state.get("direction"):
                expected_toggle = "desc" if before_state["direction"] == "asc" else "asc"
                click_rec["toggled"] = direction == expected_toggle
                if not click_rec["toggled"]:
                    result["notes"].append(
                        f"expected toggle to {expected_toggle}, got {direction}"
                    )
                    result["status"] = "FAIL"
                click_rec["order_changed"] = before_vals != vals
                if (
                    len(set(str(v).casefold() if col_type == "text" else v for v in vals)) > 1
                    and before_vals == vals
                ):
                    result["notes"].append("row order did not change on reverse click")
                    result["status"] = "FAIL"
        result["clicks"].append(click_rec)

    if result["clicks"] and all(c.get("order_ok") for c in result["clicks"]):
        toggle_ok = len(result["clicks"]) < 2 or result["clicks"][1].get("toggled", True)
        if toggle_ok and not any("did not change" in n for n in result["notes"]):
            result["status"] = "PASS"
            result["notes"] = [n for n in result["notes"] if "toggle" not in n]
    return result


CSS = """
body{font-family:Segoe UI,Arial,sans-serif;margin:24px;background:#f4f8fb;color:#0f1c2e}
h1{margin:0 0 6px;font-size:1.35rem}
.meta{color:#5a6a7e;font-size:.9rem;margin:0 0 12px}
.default{margin:0 0 18px;font-size:.95rem}
table{width:100%;border-collapse:collapse;background:#fff}
th,td{border-bottom:1px solid #d7e0ea;padding:10px 12px;text-align:left;font-size:.95rem;vertical-align:top}
th{background:#f7fafc}
.ok{color:#0f7a4a;font-weight:700}.bad{color:#a11;font-weight:700}.na{color:#5a6a7e}
.sample{display:block;margin-top:4px;color:#5a6a7e;font-size:.8rem;font-weight:400}
a{color:#0b6e99}
"""


def _dir_cell(result, direction):
    """PASS/FAIL (or —) for one direction, with a short first-values sample."""
    if result.get("status") == "NOT_SORTABLE" or not result.get("sortable_ui"):
        return "<span class=na>—</span>"
    click = next(
        (c for c in result.get("clicks") or [] if c.get("ui_direction") == direction),
        None,
    )
    if not click:
        return "<span class=bad>FAIL</span>"
    label = "PASS" if click.get("order_ok") else "FAIL"
    cls = "ok" if click.get("order_ok") else "bad"
    sample = click.get("sample_first3") or []
    sample_html = ""
    if sample:
        shown = ", ".join(html.escape(str(v)) for v in sample[:3])
        sample_html = f"<span class=sample>{shown}</span>"
    return f"<span class={cls}>{label}</span>{sample_html}"


def report_html(default_sort, columns_meta, results, meta):
    del columns_meta  # unused in simple report
    def_name = html.escape(default_sort.get("name") or "—")
    def_dir = html.escape((default_sort.get("direction") or "—").upper())
    rows = []
    for r in results:
        sortable = "yes" if r.get("sortable_ui") and r.get("status") != "NOT_SORTABLE" else "no"
        rows.append(
            f"<tr><td>{html.escape(r['name'])}</td>"
            f"<td>{sortable}</td>"
            f"<td>{_dir_cell(r, 'asc')}</td>"
            f"<td>{_dir_cell(r, 'desc')}</td></tr>"
        )
    return f"""<!DOCTYPE html><html><head><meta charset=utf-8>
<title>GC Studies — Sorting</title><style>{CSS}</style></head><body>
<h1>GC Studies — Sorting</h1>
<p class=meta><a href="{html.escape(URL)}">{html.escape(URL)}</a> ·
{html.escape(meta['collected_at'])} · {meta.get('row_count', '—')} rows</p>
<p class=default><b>Default sort:</b> {def_name} · {def_dir}</p>
<table>
<thead><tr><th>Column</th><th>Sortable?</th><th>Ascending</th><th>Descending</th></tr></thead>
<tbody>{''.join(rows)}</tbody>
</table>
</body></html>"""


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    print(f"URL={URL}\nHeadless={HEADLESS}", flush=True)
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
        open_studies(page)
        columns = header_state(page)
        default = next((c for c in columns if c.get("active")), None) or {}
        # Capture default first rows (all cells)
        default_first = page.evaluate("""() => [...document.querySelectorAll('table tbody tr')].slice(0,5).map(tr =>
          [...tr.querySelectorAll('td')].map(td => {
            const raw=(td.innerText||'').trim();
            return raw.split(/\\n+/).map(s=>s.trim()).filter(Boolean)[0]||'';
          }))""")
        row_count = len(row_values(page, 0))
        print(f"Columns={len(columns)} default={default.get('name')} {default.get('direction')} rows={row_count}", flush=True)

        results = []
        for col in columns:
            ctype = COLUMN_TYPES.get(col["index"], "text")
            print(f"  Checking [{col['index']}] {col['name']} ({ctype})…", flush=True)
            # Reset toward a known state: if this isn't the default column, still fine —
            # verify_column clicks twice from whatever state we're in.
            r = verify_column(page, col, ctype)
            print(f"    → {r['status']}", flush=True)
            results.append(r)

        browser.close()

    meta = {
        "collected_at": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC"),
        "headless": actual_headless,
        "url": URL,
        "row_count": row_count,
        "default_first_rows": default_first,
        "pass": sum(1 for r in results if r["status"] == "PASS"),
        "fail": sum(1 for r in results if r["status"] == "FAIL"),
        "not_sortable": sum(1 for r in results if r["status"] == "NOT_SORTABLE"),
    }
    payload = {"meta": meta, "default_sort": default, "columns": columns, "results": results}
    (OUT / "sorting-results.json").write_text(json.dumps(payload, indent=2), encoding="utf-8")
    html_path = OUT / "gc-studies-sorting-report.html"
    html_path.write_text(report_html(default, columns, results, meta), encoding="utf-8")
    print(f"Wrote {html_path}\nPASS={meta['pass']} FAIL={meta['fail']} NOT_SORTABLE={meta['not_sortable']}", flush=True)
    return 0 if meta["fail"] == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
