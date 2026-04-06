"""
Self-contained script: Data Commons — PHS + study caNanoLab Data (Samples / Files / Protocols),
then Cancer Nanotechnology + Protocol Name CPMV-S100A9 Production, then uncheck CPMV and filter
Protocol Type → Endotoxin, then DOI, Publication Title, Nanomaterial Entity → aptamer,
Functionalizing Entity → Antibody, Characterization Type → Clinical trial, Characterization Name →
Anti-Tumor Efficacy In Vivo. Expected counts come from **caNanoLabData_Basecounts.xlsx** next to
this script unless **CANANOLAB_BASECOUNTS_PATH** points elsewhere.

Set **CANANOLAB_DATA_COMMONS_URL** to override the portal URL (default: general-qa2).

Study Name uses a popover checklist; CN subfacets use in-sidebar accordions. HTML report: nine sections.
"""

from playwright.sync_api import (
    Locator,
    sync_playwright,
    Error as PlaywrightError,
    TimeoutError as PlaywrightTimeoutError,
)
import html
import os
import re
import sys
import time
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd

if sys.platform == "win32" and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

# -----------------------------------
# CONFIG
# -----------------------------------
_DEFAULT_DATA_COMMONS_URL = "https://general-qa2.datacommons.cancer.gov/#/data"

PHS_ACCESSION = "10.17917"
STUDY_NAME = "caNanoLab Data"

_SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_BASECOUNTS_WORKBOOK = _SCRIPT_DIR / "caNanoLabData_Basecounts.xlsx"

# Scenario IDs must match the Scenario column in caNanoLabData_Basecounts.xlsx (case-insensitive).
BC_STUDY = "study"
BC_CPMV = "cpmv"
BC_ENDOTOXIN = "endotoxin"
BC_DOI = "doi"
BC_PUBLICATION_TITLE = "publication_title"
BC_NM_APTAMER = "nm_aptamer"
BC_FE_ANTIBODY = "fe_antibody"
BC_CLINICAL_TRIAL = "clinical_trial"
BC_CNAME_ANTITUMOR = "cname_antitumor"

_REQUIRED_BASECOUNT_SCENARIOS: tuple[str, ...] = (
    BC_STUDY,
    BC_CPMV,
    BC_ENDOTOXIN,
    BC_DOI,
    BC_PUBLICATION_TITLE,
    BC_NM_APTAMER,
    BC_FE_ANTIBODY,
    BC_CLINICAL_TRIAL,
    BC_CNAME_ANTITUMOR,
)


def _data_commons_url() -> str:
    v = os.environ.get("CANANOLAB_DATA_COMMONS_URL", "").strip()
    return v if v else _DEFAULT_DATA_COMMONS_URL


def _norm_basecount_scenario(val: str) -> str:
    return re.sub(r"\s+", "_", str(val).strip().lower())


def _find_df_column(df: pd.DataFrame, *names: str) -> str | None:
    lower = {str(c).strip().lower(): str(c) for c in df.columns}
    for n in names:
        key = n.strip().lower()
        if key in lower:
            return lower[key]
    return None


def load_cananolab_basecounts_workbook(xlsx_path: str) -> dict[str, dict[str, int]]:
    """
    Load expected P/S/F/Pr counts per scenario from Excel.
    Required columns: Scenario (or Section), Participants, Samples, Files, Protocols.
    """
    path = Path(xlsx_path)
    if not path.is_file():
        raise FileNotFoundError(
            f"Base counts workbook not found: {path}\n"
            f"Expected **caNanoLabData_Basecounts.xlsx** in {_SCRIPT_DIR}, or set "
            "env CANANOLAB_BASECOUNTS_PATH."
        )

    df = pd.read_excel(path, engine="openpyxl")
    df.columns = df.columns.map(lambda c: str(c).strip())

    scen_col = _find_df_column(df, "scenario", "scenario_id", "section")
    if not scen_col:
        raise RuntimeError(
            "Workbook must have a Scenario column (or Scenario ID / Section)."
        )
    p_col = _find_df_column(df, "participants", "participant")
    s_col = _find_df_column(df, "samples", "sample")
    f_col = _find_df_column(df, "files", "file")
    pr_col = _find_df_column(df, "protocols", "protocol")
    if not all((p_col, s_col, f_col, pr_col)):
        raise RuntimeError(
            "Workbook must have Participants, Samples, Files, and Protocols columns."
        )

    out: dict[str, dict[str, int]] = {}
    for _, row in df.iterrows():
        raw = row[scen_col]
        if pd.isna(raw) or str(raw).strip() == "":
            continue
        key = _norm_basecount_scenario(str(raw))
        if not key or key.startswith("#"):
            continue

        def _cell(col: str) -> int:
            v = row[col]
            if pd.isna(v):
                raise ValueError(f"Empty cell for scenario {key!r} column {col!r}")
            return int(float(v))

        out[key] = {
            "participants": _cell(p_col),
            "samples": _cell(s_col),
            "files": _cell(f_col),
            "protocols": _cell(pr_col),
        }

    missing = [s for s in _REQUIRED_BASECOUNT_SCENARIOS if s not in out]
    if missing:
        raise ValueError(
            "Workbook is missing required scenario row(s): "
            f"{', '.join(missing)}. Found: {sorted(out.keys())}."
        )

    return out


# Cancer Nanotechnology + protocol filter scenario
CANCER_NANO_FACET = "Cancer Nanotechnology"
# QA2 header (class jss* changes): <div class="jss491">Cancer Nanotechnology</div>
# We match by exact text / id only, not jss classes.

# Nested under Cancer Nanotechnology — never click to open this header div (QA2 DOM):
#   <div id="Nanomaterial Entity" class="jss488">Nanomaterial Entity</div>
NANOMATERIAL_ENTITY_DIV_ID = "Nanomaterial Entity"
NANOMATERIAL_ENTITY_HEADER_IDS = (
    NANOMATERIAL_ENTITY_DIV_ID,
    "Nanomaterial entity",
    "NANOMATERIAL ENTITY",
    "Nanomaterial_entity",
    "nanomaterial entity",
)


def _blocked_nanomaterial_header_id(attr_id: str | None) -> bool:
    """True if this element id is the Nanomaterial Entity facet header (any known casing)."""
    if not attr_id or not attr_id.strip():
        return False
    key = attr_id.strip().lower()
    return key in {x.strip().lower() for x in NANOMATERIAL_ENTITY_HEADER_IDS}


PROTOCOL_NAME_FACET = "Protocol Name"
# QA2 DOM example (jss* class changes per build — do not rely on it):
#   <div id="Protocol Name" class="jss488">Protocol Name</div>
PROTOCOL_NAME_VALUE = "CPMV-S100A9 Production"
# Protocol option row (QA2; jss* / numeric MUI suffixes change — use partial MuiTypography-body1):
#   <p class="MuiTypography-root-23704 jss23685 MuiTypography-body1-23706">CPMV-S100A9 Production</p>

# After CPMV scenario: uncheck protocol name, filter Protocol Type → Endotoxin (under CN)
PROTOCOL_TYPE_FACET = "Protocol Type"
PROTOCOL_TYPE_VALUE = "Endotoxin"

# Scenario 4: uncheck Endotoxin, then DOI → single record (under Cancer Nanotechnology).
DOI_FACET = "DOI"
DOI_VALUE = "10.17917/0165-RC67"

# Scenario 5: uncheck DOI, then Publication Title (under CN).
PUBLICATION_TITLE_FACET = "Publication Title"
PUBLICATION_TITLE_VALUE = (
    "3D In Vitro Model (R)evolution: Unveiling Tumor-Stroma Interactions"
)

# Scenario 6: uncheck Publication Title, open Nanomaterial Entity, check entity type aptamer.
NANOMATERIAL_ENTITY_TYPE_VALUE = "aptamer"

# Scenario 7: uncheck aptamer (NM entity), Functionalizing Entity → Antibody.
FUNCTIONALIZING_ENTITY_FACET = "Functionalizing Entity"
FUNCTIONALIZING_ENTITY_VALUE = "Antibody"

# Scenario 8: uncheck Antibody (FE), Characterization Type → clinical trial (QA2 list label).
CHARACTERIZATION_TYPE_FACET = "Characterization Type"
CHARACTERIZATION_TYPE_VALUE = "Clinical trial"

# Scenario 9: uncheck Characterization Type (clinical trial), Characterization Name → Anti-Tumor….
CHARACTERIZATION_NAME_FACET = "Characterization Name"
CHARACTERIZATION_NAME_VALUE = "Anti-Tumor Efficacy In Vivo"

REPORT_PATH = "caNanoLab_QA_Report.html"


@dataclass
class CancerNanotechProtocolFilterLocators:
    """Resolved Playwright locators: collect first, then .check() the checkbox inputs."""

    sidebar: Locator | None
    cn_header: Locator
    protocol_name_header: Locator


def collect_cancer_nanotech_protocol_locators(page) -> CancerNanotechProtocolFilterLocators:
    """
    Gather locators for the Cancer Nanotechnology facet header and Protocol Name header
    scoped to that section. Call after the data view and filters are visible.
    """
    sidebar = _filter_sidebar_root(page)
    cn_header = cancer_nanotechnology_filter_trigger(page)
    protocol_name_header = _protocol_name_locator_in_cn_section(cn_header)
    return CancerNanotechProtocolFilterLocators(
        sidebar=sidebar,
        cn_header=cn_header,
        protocol_name_header=protocol_name_header,
    )


def locate_filter_checkbox(scope: Locator, label_text: str) -> Locator | None:
    """
    Return the checkbox Locator for an option label inside a filter list/popover, or None.
    Order: role=checkbox → MuiFormControlLabel → list row → MuiTypography row host.
    """
    t = label_text.strip()
    if not t:
        return None
    name_re = re.compile(re.escape(t), re.I)

    try:
        cb = scope.get_by_role("checkbox", name=name_re)
        if cb.count() > 0:
            return cb.first
    except PlaywrightError:
        pass

    try:
        labels = scope.locator(".MuiFormControlLabel-root, label.MuiFormControlLabel-root")
        match = labels.filter(has_text=name_re)
        if match.count() > 0:
            inp = match.first.locator("input[type='checkbox']")
            if inp.count() > 0:
                return inp.first
    except PlaywrightError:
        pass

    try:
        items = scope.locator(
            "[role='option'], li.MuiMenuItem-root, .MuiListItemButton-root, .MuiListItem-root"
        )
        match = items.filter(has_text=name_re)
        if match.count() > 0:
            row = match.first
            inp = row.locator("input[type='checkbox']")
            if inp.count() > 0:
                return inp.first
            cb2 = row.get_by_role("checkbox")
            if cb2.count() > 0:
                return cb2.first
    except PlaywrightError:
        pass

    try:
        paras = scope.locator(
            "p[class*='MuiTypography-body1'], div[class*='MuiTypography-body1'], "
            "span[class*='MuiTypography-body1']"
        ).filter(has_text=name_re)
        if paras.count() > 0:
            target = paras.first
            for ax in (
                "xpath=ancestor::li[1]",
                "xpath=ancestor::div[contains(@class,'MuiListItem')][1]",
                "xpath=ancestor::div[contains(@class,'MuiButtonBase')][1]",
            ):
                host = target.locator(ax)
                if host.count() == 0:
                    continue
                inp = host.first.locator("input[type='checkbox']")
                if inp.count() > 0:
                    return inp.first
                cb3 = host.first.get_by_role("checkbox")
                if cb3.count() > 0:
                    return cb3.first
    except PlaywrightError:
        pass

    return None


def _ensure_filter_checkbox_unchecked(page, scope: Locator, cb: Locator, label: str) -> None:
    """Counterpart to _ensure_filter_checkbox_checked for MUI hidden inputs."""
    cb.scroll_into_view_if_needed()
    time.sleep(0.12)
    try:
        if not cb.is_checked():
            return
    except PlaywrightError:
        pass
    try:
        cb.uncheck(force=True)
        return
    except PlaywrightError as e:
        if "did not change" not in str(e).lower():
            pass
    for ax in (
        "xpath=ancestor::label[1]",
        "xpath=ancestor::*[contains(@class,'MuiFormControlLabel')][1]",
        "xpath=ancestor::li[1]",
    ):
        try:
            w = cb.locator(ax)
            if w.count() > 0:
                w.first.click(force=True)
                time.sleep(0.25)
                return
        except PlaywrightError:
            continue
    sb = _filter_sidebar_root(page)
    roots: tuple = (scope,)
    if sb is not None:
        roots = roots + (sb,)
    roots = roots + (page,)
    for root in roots:
        if _click_mui_typography_filter_row_toggle(page, root, label):
            return
    for root in roots:
        try:
            _toggle_filter_option_paragraph_click(page, label, root=root)
            return
        except Exception:
            continue
    raise RuntimeError(f"Could not uncheck {label!r}")


def _ensure_filter_checkbox_checked(page, scope: Locator, cb: Locator, label: str) -> None:
    """
    MUI often uses aria-hidden native inputs; Playwright check() may not flip visible state.
    Fall back to label / FormControlLabel / list row / typography click.
    """
    cb.scroll_into_view_if_needed()
    time.sleep(0.12)
    try:
        if cb.is_checked():
            return
    except PlaywrightError:
        pass
    try:
        cb.check(force=True)
        return
    except PlaywrightError as e:
        if "did not change" not in str(e).lower():
            pass
    for ax in (
        "xpath=ancestor::label[1]",
        "xpath=ancestor::*[contains(@class,'MuiFormControlLabel')][1]",
        "xpath=ancestor::li[1]",
    ):
        try:
            w = cb.locator(ax)
            if w.count() > 0:
                w.first.click(force=True)
                time.sleep(0.25)
                return
        except PlaywrightError:
            continue
    if _try_mui_typography_filter_row(page, scope, label):
        return
    _toggle_filter_option_paragraph_click(page, label, root=scope)


def check_filter_option_checkbox(page, scope: Locator, label: str) -> None:
    """
    Scroll the filter list, resolve the checkbox locator for `label`, then .check() it.
    Falls back to typography row click if no checkbox is found.
    """
    for _ in range(7):
        cb = locate_filter_checkbox(scope, label)
        if cb is not None:
            _ensure_filter_checkbox_checked(page, scope, cb, label)
            return
        _scroll_filter_list_down(scope)
        time.sleep(0.3)

    _scroll_filter_list_top(scope)
    time.sleep(0.35)
    for _ in range(40):
        cb = locate_filter_checkbox(scope, label)
        if cb is not None:
            _ensure_filter_checkbox_checked(page, scope, cb, label)
            return
        _scroll_filter_list_down(scope)
        time.sleep(0.3)

    print(f"⚠️ No checkbox locator for {label!r}; using label click fallback.")
    _toggle_filter_option_paragraph_click(page, label, root=scope)


def _facet_checkbox_input_id(facet_label: str, option_value: str) -> str:
    """QA2 encodes filter checkboxes as id=\"checkbox_<facet>_<value>\" (spaces in facet kept)."""
    return f"checkbox_{facet_label}_{option_value}"


def toggle_facet_checkbox_by_input_id(
    page, facet_label: str, option_value: str, *, want_checked: bool
) -> bool:
    """
    Click the MUI label wrapping the checkbox input with the stable id, to check or uncheck.
    Returns True if the input node existed and was clicked.
    """
    cid = _facet_checkbox_input_id(facet_label, option_value)
    inp = page.locator(f'[id="{cid}"]')
    try:
        if inp.count() == 0:
            return False
        inp.first.scroll_into_view_if_needed()
        time.sleep(0.15)
        try:
            if inp.first.is_checked() == want_checked:
                return True
        except PlaywrightError:
            pass
        for ax in (
            "xpath=ancestor::label[1]",
            "xpath=ancestor::*[contains(@class,'MuiFormControlLabel')][1]",
        ):
            w = inp.first.locator(ax)
            if w.count() > 0:
                w.first.click(force=True)
                time.sleep(0.25)
                return True
        inp.first.click(force=True)
        time.sleep(0.25)
        return True
    except PlaywrightError:
        return False


def uncheck_filter_option_checkbox(page, scope: Locator, label: str) -> None:
    """Scroll the filter list, resolve the checkbox for `label`, then uncheck it."""
    for _ in range(7):
        cb = locate_filter_checkbox(scope, label)
        if cb is not None:
            _ensure_filter_checkbox_unchecked(page, scope, cb, label)
            return
        _scroll_filter_list_down(scope)
        time.sleep(0.3)

    _scroll_filter_list_top(scope)
    time.sleep(0.35)
    for _ in range(40):
        cb = locate_filter_checkbox(scope, label)
        if cb is not None:
            _ensure_filter_checkbox_unchecked(page, scope, cb, label)
            return
        _scroll_filter_list_down(scope)
        time.sleep(0.3)

    print(f"⚠️ No checkbox locator to uncheck {label!r}; toggling row click.")
    roots: list[Locator] = [scope]
    sb = _filter_sidebar_root(page)
    if sb is not None:
        roots.append(sb)
    last_err: Exception | None = None
    for r in roots:
        try:
            _toggle_filter_option_paragraph_click(page, label, root=r)
            return
        except Exception as e:
            last_err = e
            continue
    if last_err is not None:
        raise last_err
    raise RuntimeError(f"Could not uncheck filter option {label!r}")


def _filter_sidebar_root(page):
    """Prefer the left filter column so we do not match stray text in main content."""
    for anchor in ("PHS ACCESSION", "Study Name"):
        try:
            hit = page.get_by_text(anchor, exact=False).first
            if hit.count() == 0:
                continue
            hit.wait_for(state="visible", timeout=3000)
            for ax in (
                "xpath=ancestor::div[contains(@class,'MuiDrawer-paper')][1]",
                "xpath=ancestor::aside[1]",
                "xpath=ancestor::div[contains(@class,'MuiPaper-root')][1]",
            ):
                dr = hit.locator(ax)
                if dr.count() == 0:
                    continue
                cand = dr.first
                if cand.is_visible():
                    return cand
        except PlaywrightError:
            continue

    for sel in (
        "aside",
        "[role='complementary']",
        ".MuiDrawer-paper",
        "[class*='Filter']",
        "[class*='filter']",
        "[class*='Drawer']",
        "[class*='drawer']",
    ):
        loc = page.locator(sel).first
        try:
            if loc.count() > 0 and loc.is_visible():
                return loc
        except PlaywrightError:
            continue
    return None


def _accordion_direct_summary(accordion_root: Locator) -> Locator | None:
    """
    The summary element for this accordion root only — not nested child accordions.
    Using .first anywhere under the root wrongly picks Nanomaterial’s summary inside CN.
    """
    for xp in (
        "xpath=child::*[contains(@class,'MuiAccordionSummary')][1]",
        "xpath=child::*[contains(@class,'AccordionSummary')][1]",
    ):
        s = accordion_root.locator(xp)
        if s.count() > 0:
            return s.first
    return None


def _accordion_summary_suggests_nanomaterial(accordion_root: Locator) -> bool:
    """True when this accordion root’s own summary is the Nanomaterial entity row."""
    try:
        summ = _accordion_direct_summary(accordion_root)
        if summ is None:
            return False
        parts = (
            summ.inner_text(),
            summ.get_attribute("aria-label") or "",
            summ.get_attribute("id") or "",
        )
        blob = " ".join(parts).lower()
        return "nanomaterial" in blob
    except PlaywrightError:
        return False


def _locator_in_nanomaterial_accordion_row(loc: Locator) -> bool:
    """True if the nearest ancestor MuiAccordion-root is the nanomaterial entity row."""
    acc = loc.locator("xpath=ancestor::div[contains(@class,'MuiAccordion-root')][1]")
    if acc.count() == 0:
        return False
    return _accordion_summary_suggests_nanomaterial(acc.first)


def _collapse_accordion_summary_if_expanded(summ: Locator) -> bool:
    try:
        if summ.get_attribute("aria-expanded") != "true":
            return False
        summ.scroll_into_view_if_needed()
        time.sleep(0.12)
        summ.click(force=True)
        time.sleep(0.4)
        return True
    except PlaywrightError:
        return False


def collapse_nanomaterial_entity_subfacet_under_cn(page, cn_header: Locator) -> None:
    """
    Collapse an expanded “Nanomaterial entity” row under the CN accordion subtree when needed.
    """
    acc = cn_header.locator(
        "xpath=ancestor::div[contains(@class,'MuiAccordion-root')][1]"
    )
    if acc.count() > 0:
        root = acc.first
        if not _accordion_summary_suggests_nanomaterial(root):
            for mid in NANOMATERIAL_ENTITY_HEADER_IDS:
                header = root.locator(f'[id="{mid}"]')
                if header.count() == 0:
                    continue
                try:
                    summ = header.first.locator(
                        "xpath=ancestor::*[contains(@class,'MuiAccordionSummary')][1]"
                    )
                    if summ.count() > 0 and _collapse_accordion_summary_if_expanded(
                        summ.first
                    ):
                        print(
                            "ℹ️ Collapsed Nanomaterial entity (id "
                            f"{mid!r} under CN subtree)."
                        )
                except PlaywrightError:
                    continue

            summaries = root.locator("[class*='MuiAccordionSummary']")
            for i in range(summaries.count()):
                s = summaries.nth(i)
                try:
                    blob = (
                        s.inner_text() + " " + (s.get_attribute("aria-label") or "")
                    ).lower()
                    if "nanomaterial" not in blob:
                        continue
                    if _collapse_accordion_summary_if_expanded(s):
                        print(
                            "ℹ️ Collapsed Nanomaterial entity (summary under CN subtree)."
                        )
                except PlaywrightError:
                    continue


def cancer_nanotechnology_filter_trigger(page):
    """
    Click target for the Cancer Nanotechnology facet header (div with label text; id may be absent).
    Prefers the filter sidebar; skips hits tied to the Nanomaterial entity accordion row.
    """
    side = _filter_sidebar_root(page)
    scope = side if side is not None else page

    def _skip_nm(el: Locator) -> bool:
        if _blocked_nanomaterial_header_id(el.get_attribute("id")):
            return True
        return _locator_in_nanomaterial_accordion_row(el)

    by_id = page.locator(f"//*[@id='{CANCER_NANO_FACET}']")
    for i in range(by_id.count()):
        el = by_id.nth(i)
        if not _skip_nm(el):
            return el

    hits = scope.get_by_text(CANCER_NANO_FACET, exact=True)
    for i in range(hits.count()):
        el = hits.nth(i)
        if not _skip_nm(el):
            return el

    divs = scope.locator(
        f"//div[normalize-space(.)='{CANCER_NANO_FACET}']"
        f"[not(.//div[normalize-space(.)='{CANCER_NANO_FACET}'])]"
    )
    for i in range(divs.count()):
        el = divs.nth(i)
        if not _skip_nm(el):
            return el

    raise RuntimeError(
        f"Could not find header {CANCER_NANO_FACET!r} outside a nanomaterial accordion row."
    )


def _protocol_name_associated_with_cn_header(cn_header: Locator, proto: Locator) -> bool:
    """
    True if the facet node is inside the CN header’s subtree or follows the CN header
    in document order (so we don’t grab another column’s div#Protocol Name, etc.).
    """
    page = cn_header.page
    cn_h = pr_h = None
    try:
        cn_h = cn_header.element_handle()
        pr_h = proto.element_handle()
        if cn_h is None or pr_h is None:
            return False
        return page.evaluate(
            """([cn, pr]) => {
                if (!cn || !pr) return false;
                const pos = cn.compareDocumentPosition(pr);
                if (pos & Node.DOCUMENT_POSITION_DISCONNECTED) return false;
                const CONTAINS = Node.DOCUMENT_POSITION_CONTAINS;
                const FOLLOWING = Node.DOCUMENT_POSITION_FOLLOWING;
                return Boolean(pos & CONTAINS) || Boolean(pos & FOLLOWING);
            }""",
            [cn_h, pr_h],
        )
    except PlaywrightError:
        return False
    finally:
        for h in (cn_h, pr_h):
            if h is not None:
                try:
                    h.dispose()
                except PlaywrightError:
                    pass


def _cn_section_facet_header_locator(cn_header: Locator, facet_id: str) -> Locator:
    """
    Facet div#`facet_id` that belongs to the Cancer Nanotechnology block only.

    Prefer the nearest MuiAccordion-root that contains the CN header, then the first
    matching id inside that root. That avoids opening another entity’s “Protocol Name”
    (e.g. nanomaterial) when a global id search would hit the wrong row first.

    Fallback: direct following-sibling, then first matching id in the sidebar that
    follows the CN header in document order.

    When facet_id is the Nanomaterial Entity header id, the match lives inside that
    accordion row — do not skip it via _locator_in_nanomaterial_accordion_row.
    """
    page = cn_header.page
    resolving_nm_header = _blocked_nanomaterial_header_id(facet_id)

    for depth in range(1, 12):
        acc = cn_header.locator(
            f"xpath=ancestor::div[contains(@class,'MuiAccordion-root')][{depth}]"
        )
        if acc.count() == 0:
            break
        if _accordion_summary_suggests_nanomaterial(acc.first):
            continue
        inner = acc.first.locator(f'[id="{facet_id}"]')
        for j in range(inner.count()):
            cand = inner.nth(j)
            try:
                if not resolving_nm_header and _locator_in_nanomaterial_accordion_row(
                    cand
                ):
                    continue
            except PlaywrightError:
                continue
            return cand

    sib = cn_header.locator(
        f"xpath=following-sibling::div[@id='{facet_id}'][1]"
    )
    if sib.count() > 0 and (
        resolving_nm_header or not _locator_in_nanomaterial_accordion_row(sib.first)
    ):
        return sib.first

    def _first_matching_after_cn(scope: Locator) -> Locator | None:
        nodes = scope.locator(f'[id="{facet_id}"]')
        try:
            n = nodes.count()
        except PlaywrightError:
            return None
        for i in range(n):
            cand = nodes.nth(i)
            try:
                if not resolving_nm_header and _locator_in_nanomaterial_accordion_row(
                    cand
                ):
                    continue
                if _protocol_name_associated_with_cn_header(cn_header, cand):
                    return cand
            except PlaywrightError:
                continue
        return None

    side = _filter_sidebar_root(page)
    for scope in ([side] if side is not None else []) + [page]:
        picked = _first_matching_after_cn(scope)
        if picked is not None:
            return picked

    raise RuntimeError(
        f"Could not find div#\"{facet_id}\" under the {CANCER_NANO_FACET!r} accordion."
    )


def _protocol_name_locator_in_cn_section(cn_header: Locator) -> Locator:
    return _cn_section_facet_header_locator(cn_header, PROTOCOL_NAME_FACET)


def _protocol_type_locator_in_cn_section(cn_header: Locator) -> Locator:
    return _cn_section_facet_header_locator(cn_header, PROTOCOL_TYPE_FACET)


def _doi_locator_in_cn_section(cn_header: Locator) -> Locator:
    return _cn_section_facet_header_locator(cn_header, DOI_FACET)


def _publication_title_locator_in_cn_section(cn_header: Locator) -> Locator:
    return _cn_section_facet_header_locator(cn_header, PUBLICATION_TITLE_FACET)


def _nanomaterial_entity_facet_header_in_cn_section(cn_header: Locator) -> Locator:
    return _cn_section_facet_header_locator(cn_header, NANOMATERIAL_ENTITY_DIV_ID)


def _functionalizing_entity_locator_in_cn_section(cn_header: Locator) -> Locator:
    return _cn_section_facet_header_locator(cn_header, FUNCTIONALIZING_ENTITY_FACET)


def _characterization_type_locator_in_cn_section(cn_header: Locator) -> Locator:
    return _cn_section_facet_header_locator(cn_header, CHARACTERIZATION_TYPE_FACET)


def _characterization_name_locator_in_cn_section(cn_header: Locator) -> Locator:
    return _cn_section_facet_header_locator(cn_header, CHARACTERIZATION_NAME_FACET)


def _wait_after_facet_header_click(page) -> None:
    """After clicking a filter facet header, wait for list / popover."""
    try:
        page.locator("//p[contains(@class,'MuiTypography-body1')]").first.wait_for(
            timeout=10000
        )
    except PlaywrightTimeoutError:
        pass
    deadline = time.time() + 10
    while time.time() < deadline:
        papers = page.locator(".MuiPopover-paper, .MuiMenu-paper")
        if papers.count() == 0:
            time.sleep(0.2)
            continue
        try:
            if papers.last.is_visible():
                break
        except PlaywrightError:
            pass
        time.sleep(0.2)
    time.sleep(1.5)


def _facet_header_targets_nanomaterial_entity(header: Locator) -> bool:
    """Never expand the Nanomaterial entity row — only Protocol Name / Type / CN, etc."""
    try:
        hid = header.get_attribute("id") or ""
        if _blocked_nanomaterial_header_id(hid) or "nanomaterial" in hid.lower():
            return True
        summ = header.locator(
            "xpath=ancestor::*[contains(@class,'MuiAccordionSummary')][1]"
        )
        if summ.count() > 0:
            sid = summ.first.get_attribute("id") or ""
            if _blocked_nanomaterial_header_id(sid):
                return True
            blob = (
                summ.first.inner_text()
                + " "
                + (summ.first.get_attribute("aria-label") or "")
            ).lower()
            if "nanomaterial" in blob:
                return True
    except PlaywrightError:
        pass
    return False


def expand_sidebar_facet_header(
    page, header: Locator, *, allow_nanomaterial_entity: bool = False
) -> None:
    """
    QA2 filter sidebar: facet rows are often MuiAccordion — one click expands, a second
    collapses. If the summary already has aria-expanded=true, do not click again.
    (Popover-style filters still get a single click when no AccordionSummary wraps the node.)

    By default, never expand the Nanomaterial Entity header (wrong-clicks elsewhere).
    Set allow_nanomaterial_entity=True only when intentionally opening that facet.
    """
    header.wait_for(state="visible", timeout=15000)
    header.scroll_into_view_if_needed()
    if not allow_nanomaterial_entity and _facet_header_targets_nanomaterial_entity(
        header
    ):
        return
    try:
        summary = header.locator(
            "xpath=ancestor::*[contains(@class,'MuiAccordionSummary')][1]"
        )
        if summary.count() > 0 and summary.first.get_attribute("aria-expanded") == "true":
            time.sleep(0.25)
            return
    except PlaywrightError:
        pass
    if not allow_nanomaterial_entity and _blocked_nanomaterial_header_id(
        header.get_attribute("id")
    ):
        return
    header.click()
    _wait_after_facet_header_click(page)


def open_cancer_from_locator(page, cn_header: Locator) -> None:
    """Expand the Cancer Nanotechnology in-sidebar section (no separate checkbox row on QA2)."""
    expand_sidebar_facet_header(page, cn_header)


def open_cancer_nanotechnology_facet(page):
    """Open the Cancer Nanotechnology filter using the real header div, then wait for options UI."""
    open_cancer_from_locator(page, cancer_nanotechnology_filter_trigger(page))


def open_filter_panel_by_id(page, div_id):
    if _blocked_nanomaterial_header_id(div_id):
        raise RuntimeError(
            f"Refusing to click blocked Nanomaterial Entity header id {div_id!r}"
        )
    loc = page.locator(f"//div[@id='{div_id}']")
    if loc.count() > 0:
        loc.first.click()
    else:
        page.get_by_text(div_id, exact=True).first.click()
    page.locator("//p[contains(@class,'MuiTypography-body1')]").first.wait_for(timeout=10000)
    deadline = time.time() + 10
    while time.time() < deadline:
        papers = page.locator(".MuiPopover-paper, .MuiMenu-paper")
        if papers.count() == 0:
            time.sleep(0.2)
            continue
        try:
            if papers.last.is_visible():
                break
        except PlaywrightError:
            pass
        time.sleep(0.2)
    time.sleep(1.5)


def open_study_filter(page):
    open_filter_panel_by_id(page, "Study Name")


def build_xpath_text(text):
    if "'" in text:
        parts = text.split("'")
        return "concat(" + ", \"'\", ".join([f"'{p}'" for p in parts]) + ")"
    return f"'{text}'"


def active_filter_popover(page):
    """Visible MUI filter menu / popover (DOM) that lists checkbox options."""
    for sel in (
        ".MuiPopover-paper.MuiPaper-elevation8",
        ".MuiPopover-paper",
        ".MuiMenu-paper",
        "[role='presentation'] .MuiPaper-root",
        "div[role='presentation'] .MuiPopover-paper",
        ".MuiModal-root .MuiPaper-root",
    ):
        loc = page.locator(sel)
        if loc.count() == 0:
            continue
        cand = loc.last
        try:
            if cand.is_visible():
                return cand
        except PlaywrightError:
            continue
    return None


def filter_options_container(page):
    """Where filter checkbox / option rows live after opening a facet (popover, menu, or drawer)."""
    pop = active_filter_popover(page)
    if pop is not None:
        return pop
    for sel in (
        ".MuiAccordionDetails-root",
        ".MuiCollapse-root .MuiList-root",
        ".MuiDrawer-paper",
    ):
        loc = page.locator(sel)
        if loc.count() == 0:
            continue
        cand = loc.last
        try:
            if cand.is_visible():
                return cand
        except PlaywrightError:
            continue
    return None


def _accordion_expanded_panel_for_facet_header(facet_header: Locator) -> Locator | None:
    """
    In-place expanded facet: options live under the same MuiAccordion-root as the header.
    Do not use document-wide .last AccordionDetails — that is often another entity’s panel.
    """
    for ax in ("xpath=ancestor::div[contains(@class,'MuiAccordion-root')][1]",):
        root = facet_header.locator(ax)
        if root.count() == 0:
            continue
        acc = root.first
        for sel in (".MuiAccordionDetails-root", "[class*='AccordionDetails']"):
            det = acc.locator(sel)
            n = det.count()
            for i in range(n):
                d = det.nth(i)
                try:
                    if d.is_visible():
                        return d
                except PlaywrightError:
                    continue
    return None


def resolve_filter_list_scope(
    page, facet_header: Locator, wait_seconds: float = 2.8
) -> Locator:
    """
    Prefer MUI popover/menu (often portaled) over in-place accordion details.
    Poll briefly: the menu can mount a few hundred ms after the facet header click.
    """
    deadline = time.time() + wait_seconds
    while time.time() < deadline:
        pop = active_filter_popover(page)
        if pop is not None:
            return pop
        panel = _accordion_expanded_panel_for_facet_header(facet_header)
        if panel is not None:
            return panel
        time.sleep(0.12)

    alt = filter_options_container(page)
    if alt is not None:
        return alt
    return page


def resolve_filter_list_scope_prefer_accordion(page, facet_header: Locator) -> Locator:
    """In-sidebar option lists: prefer this facet’s AccordionDetails over a stray Popover."""
    panel = _accordion_expanded_panel_for_facet_header(facet_header)
    if panel is not None:
        return panel
    return resolve_filter_list_scope(page, facet_header)


def _scroll_filter_list_down(scope):
    try:
        scope.evaluate(
            """(el) => {
                const sc = el.querySelector('.MuiList-root')
                    || el.querySelector('[role="listbox"]')
                    || el.querySelector('.MuiMenu-list')
                    || el;
                sc.scrollTop += 300;
            }"""
        )
    except PlaywrightError:
        pass


def _scroll_filter_list_top(scope):
    try:
        scope.evaluate(
            """(el) => {
                const sc = el.querySelector('.MuiList-root')
                    || el.querySelector('[role="listbox"]')
                    || el.querySelector('.MuiMenu-list')
                    || el;
                sc.scrollTop = 0;
            }"""
        )
    except PlaywrightError:
        pass


def _try_check_checkbox_in_scope(scope, label_text: str) -> bool:
    """Resolve checkbox via locate_filter_checkbox, then check (single code path)."""
    cb = locate_filter_checkbox(scope, label_text)
    if cb is None:
        return False
    try:
        cb.scroll_into_view_if_needed()
        time.sleep(0.15)
        if not cb.is_checked():
            cb.check(force=True)
        return True
    except PlaywrightError:
        return False


def _try_mui_typography_filter_row(page, scope, label_text: str) -> bool:
    """
    Select option rows built as MuiTypography body1 <p> (stable substring: MuiTypography-body1).
    Prefer an ancestor row checkbox; else click left of the label like the legacy path.
    """
    t = label_text.strip()
    if not t:
        return False
    name_re = re.compile(re.escape(t), re.I)

    paras = scope.locator(
        "p[class*='MuiTypography-body1'], div[class*='MuiTypography-body1'], "
        "span[class*='MuiTypography-body1']"
    ).filter(has_text=name_re)
    if paras.count() == 0:
        return False

    target = paras.first
    target.scroll_into_view_if_needed()
    time.sleep(0.15)

    for ax in (
        "xpath=ancestor::li[1]",
        "xpath=ancestor::div[contains(@class,'MuiListItem')][1]",
        "xpath=ancestor::div[contains(@class,'MuiButtonBase')][1]",
    ):
        try:
            host = target.locator(ax)
            if host.count() == 0:
                continue
            inp = host.first.locator("input[type='checkbox']")
            if inp.count() > 0:
                if not inp.first.is_checked():
                    inp.first.check(force=True)
                return True
            cb = host.first.get_by_role("checkbox")
            if cb.count() > 0:
                if not cb.first.is_checked():
                    cb.first.check()
                return True
        except PlaywrightError:
            continue

    box = target.bounding_box()
    if box:
        page.mouse.click(box["x"] - 15, box["y"] + box["height"] / 2)
        return True
    target.click()
    return True


def _click_mui_typography_filter_row_toggle(page, scope, label_text: str) -> bool:
    """Click left of the MUI body1 label to toggle checkbox (check or uncheck)."""
    t = label_text.strip()
    if not t:
        return False
    name_re = re.compile(re.escape(t), re.I)
    paras = scope.locator(
        "p[class*='MuiTypography-body1'], div[class*='MuiTypography-body1'], "
        "span[class*='MuiTypography-body1']"
    ).filter(has_text=name_re)
    if paras.count() == 0:
        return False
    target = paras.first
    target.scroll_into_view_if_needed()
    time.sleep(0.15)
    box = target.bounding_box()
    if box:
        page.mouse.click(box["x"] - 15, box["y"] + box["height"] / 2)
        return True
    target.click()
    return True


def _try_select_filter_row(page, scope, label_text: str) -> bool:
    return _try_check_checkbox_in_scope(scope, label_text) or _try_mui_typography_filter_row(
        page, scope, label_text
    )


def _toggle_filter_option_paragraph_click(page, option_label, root=None):
    """Legacy: click left of MuiTypography label when no checkbox is exposed."""
    scope = root or page
    name = option_label.strip()
    safe_text = build_xpath_text(name)

    typo_row = "p[class*='MuiTypography-body1'], div[class*='MuiTypography-body1'], span[class*='MuiTypography-body1']"
    name_re = re.compile(re.escape(name), re.I)

    def find_and_click():
        for role in ("menuitem", "option"):
            try:
                r = scope.get_by_role(role, name=name_re)
                if r.count() > 0:
                    r.first.scroll_into_view_if_needed()
                    time.sleep(0.2)
                    r.first.click()
                    return True
            except PlaywrightError:
                pass

        for list_sel in (
            ".MuiList-root",
            "[role='listbox']",
            "ul.MuiMenu-list",
            ".MuiMenu-list",
        ):
            inner = scope.locator(list_sel).first
            try:
                if inner.count() == 0:
                    continue
                t = inner.get_by_text(name, exact=True)
                if t.count() > 0:
                    t.first.scroll_into_view_if_needed()
                    time.sleep(0.2)
                    t.first.click()
                    return True
            except PlaywrightError:
                pass

        typed = scope.locator(typo_row).filter(has_text=name_re)
        if typed.count() > 0:
            study = typed.first
            study.scroll_into_view_if_needed()
            time.sleep(0.5)
            box = study.bounding_box()
            if box:
                page.mouse.click(box["x"] - 15, box["y"] + box["height"] / 2)
                return True
            return False

        study = scope.locator(
            f"xpath=.//*[self::p or self::div or self::span]"
            f"[contains(@class,'MuiTypography-body1')][contains(normalize-space(), {safe_text})]"
        )
        n = study.count()
        if n == 0:
            return False
        if n > 1:
            exact = scope.locator(
                f"xpath=.//*[self::p or self::div or self::span]"
                f"[contains(@class,'MuiTypography-body1')][normalize-space()={safe_text}]"
            )
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
        _scroll_filter_list_down(scope)
        time.sleep(0.5)

    print(f"🔄 Reset scroll for: {name}")
    _scroll_filter_list_top(scope)
    time.sleep(1)

    for i in range(35):
        if find_and_click():
            return
        print(f"↘️ Scrolling for: {name} (step {i+1})")
        _scroll_filter_list_down(scope)
        time.sleep(0.5)

    raise Exception(f"❌ Filter option not found: {name}")


def select_filter_checkbox(page, option_label, root=None):
    """
    Select a filter value via checkboxes and/or MuiTypography-body1 option rows in the popover.
    root: optional locator for the popover/list; otherwise uses active_filter_popover(page).
    """
    text = option_label.strip()
    scope = root if root is not None else active_filter_popover(page)
    if scope is None:
        scope = page

    if _try_select_filter_row(page, scope, text):
        return

    for _ in range(6):
        _scroll_filter_list_down(scope)
        time.sleep(0.35)
        if _try_select_filter_row(page, scope, text):
            return

    print(f"🔄 Reset scroll for checkbox: {text}")
    _scroll_filter_list_top(scope)
    time.sleep(0.4)

    for i in range(40):
        if _try_select_filter_row(page, scope, text):
            return
        print(f"↘️ Scrolling for: {text} (step {i+1})")
        _scroll_filter_list_down(scope)
        time.sleep(0.35)

    print("⚠️ Filter row not selected; falling back to full typography scroll search.")
    _toggle_filter_option_paragraph_click(page, text, root=scope)


def toggle_study_checkbox(page, study_name, root=None):
    """Alias: prefer DOM checkboxes in filter popovers."""
    select_filter_checkbox(page, study_name, root=root)


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


def click_tab(page, tab_name):
    tab = page.get_by_role("tab", name=re.compile(tab_name))
    tab.click()
    time.sleep(2)


def get_table_total_count(page):
    """
    Read total rows from the active grid’s pagination/footer (MUI-style “1–10 of 1,672”).
    Returns None if no recognizable total is found.
    """
    time.sleep(0.5)
    footer_selectors = [
        ".MuiTablePagination-displayedRows",
        "p.MuiTablePagination-displayedRows",
        ".MuiTablePagination-root",
        "[class*='TablePagination']",
        ".MuiDataGrid-footerContainer",
    ]
    for sel in footer_selectors:
        loc = page.locator(sel)
        n = loc.count()
        if n == 0:
            continue
        for i in range(min(n, 8)):
            el = loc.nth(i)
            try:
                if not el.is_visible():
                    continue
                text = el.inner_text()
            except PlaywrightError:
                continue
            norm = text.replace("\u2013", "-").replace("–", "-")
            m = re.search(r"of\s+([\d,]+)", norm, re.I)
            if m:
                return int(m.group(1).replace(",", ""))
    return None


def verify_tab_table_and_base(page, tab_label, base_value):
    """
    Compare base file expectation vs tab badge count vs table pagination total.
    Returns a dict for console + HTML reporting.
    """
    tab_count = get_tab_count(page, tab_label)
    click_tab(page, tab_label)
    table_count = get_table_total_count(page)

    print(f"\n── {tab_label} ──")
    print(f"  Base file: {base_value} | Tab: {tab_count} | Table total: {table_count}")

    base_vs_tab = tab_count == base_value
    if table_count is None and base_value == 0 and tab_count == 0:
        # Empty tabs often omit pagination; treat as 0 rows.
        table_count = 0
        table_display = "0"
        print("  (Table total inferred as 0 — no pagination text.)")
        base_vs_table = True
        tab_vs_table = True
    elif table_count is None:
        print("  ⚠️ Could not read table total from pagination/footer.")
        base_vs_table = False
        tab_vs_table = False
        table_display = "—"
    else:
        base_vs_table = table_count == base_value
        tab_vs_table = tab_count == table_count
        table_display = str(table_count)

    check_rows = [
        ("base vs tab", base_vs_tab),
        ("base vs table", base_vs_table),
        ("tab vs table", tab_vs_table),
    ]
    for label, ok in check_rows:
        sym = "✅" if ok else "❌"
        print(f"  {sym} {label}")

    row_ok = all(ok for _, ok in check_rows)

    return {
        "entity": tab_label,
        "base": base_value,
        "tab": tab_count,
        "table": table_count,
        "table_display": table_display,
        "base_vs_tab": base_vs_tab,
        "base_vs_table": base_vs_table,
        "tab_vs_table": tab_vs_table,
        "row_ok": row_ok,
    }


def generate_html_report(
    sections,
    overall_ok,
    *,
    url,
    phs,
    study_name,
    scenario2_summary="",
    scenario3_summary="",
    scenario4_summary="",
    scenario5_summary="",
    scenario6_summary="",
    scenario7_summary="",
    scenario8_summary="",
    scenario9_summary="",
    report_path=REPORT_PATH,
):
    """
    sections: list of (section_title: str, results: list[dict])
    """
    generated = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")

    doc = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="utf-8"/>
    <title>caNanoLab Data QA Report</title>
    <style>
        body {{ font-family: Arial, sans-serif; margin: 24px; }}
        h1 {{ font-size: 1.35rem; }}
        h2 {{ font-size: 1.1rem; margin-top: 28px; margin-bottom: 10px; }}
        .meta {{ color: #444; margin-bottom: 20px; line-height: 1.6; }}
        table {{ border-collapse: collapse; width: 100%; max-width: 1100px; margin-bottom: 8px; }}
        th, td {{ border: 1px solid #ccc; padding: 8px 10px; text-align: center; }}
        th {{ background: #333; color: #fff; }}
        td.pass {{ background: #d4edda; font-weight: 600; }}
        td.fail {{ background: #f8d7da; font-weight: 600; }}
        tr.section-fail td:not(.pass):not(.fail) {{ background: #fff8f8; }}
        .summary {{ margin-top: 20px; font-size: 1.1rem; }}
        .summary.pass {{ color: #155724; }}
        .summary.fail {{ color: #721c24; }}
    </style>
</head>
<body>
    <h1>caNanoLab Data — count QA report</h1>
    <div class="meta">
        <div><strong>Generated:</strong> {html.escape(generated)}</div>
        <div><strong>URL:</strong> {html.escape(url)}</div>
        <div><strong>PHS ACCESSION:</strong> {html.escape(phs)}</div>
        <div><strong>Study:</strong> {html.escape(study_name)}</div>
"""
    if scenario2_summary:
        doc += f"        <div><strong>Scenario 2:</strong> {html.escape(scenario2_summary)}</div>\n"
    if scenario3_summary:
        doc += f"        <div><strong>Scenario 3:</strong> {html.escape(scenario3_summary)}</div>\n"
    if scenario4_summary:
        doc += f"        <div><strong>Scenario 4:</strong> {html.escape(scenario4_summary)}</div>\n"
    if scenario5_summary:
        doc += f"        <div><strong>Scenario 5:</strong> {html.escape(scenario5_summary)}</div>\n"
    if scenario6_summary:
        doc += f"        <div><strong>Scenario 6:</strong> {html.escape(scenario6_summary)}</div>\n"
    if scenario7_summary:
        doc += f"        <div><strong>Scenario 7:</strong> {html.escape(scenario7_summary)}</div>\n"
    if scenario8_summary:
        doc += f"        <div><strong>Scenario 8:</strong> {html.escape(scenario8_summary)}</div>\n"
    if scenario9_summary:
        doc += f"        <div><strong>Scenario 9:</strong> {html.escape(scenario9_summary)}</div>\n"
    doc += """    </div>
"""

    for section_title, results in sections:
        doc += f"    <h2>{html.escape(section_title)}</h2>\n"
        doc += """    <table>
        <thead>
            <tr>
                <th>Entity</th>
                <th>Base file</th>
                <th>Tab count</th>
                <th>Table total</th>
                <th>Status</th>
            </tr>
        </thead>
        <tbody>
"""
        for r in results:
            row_cls = "" if r["row_ok"] else ' class="section-fail"'
            status = "PASS" if r["row_ok"] else "FAIL"
            status_cls = "pass" if r["row_ok"] else "fail"
            doc += f"""            <tr{row_cls}>
                <td>{html.escape(r["entity"])}</td>
                <td>{r["base"]}</td>
                <td>{r["tab"]}</td>
                <td>{html.escape(r["table_display"])}</td>
                <td class="{status_cls}">{html.escape(status)}</td>
            </tr>
"""
        doc += "        </tbody>\n    </table>\n"

    overall_text = "PASS" if overall_ok else "FAIL"
    overall_cls = "summary pass" if overall_ok else "summary fail"
    doc += f"""    <p class="{overall_cls}"><strong>Overall: {html.escape(overall_text)}</strong></p>
</body>
</html>
"""

    with open(report_path, "w", encoding="utf-8") as f:
        f.write(doc)


def dismiss_continue_if_present(page):
    try:
        page.get_by_role("button", name=re.compile(r"^Continue$", re.I)).click(
            timeout=15000
        )
    except PlaywrightTimeoutError:
        try:
            page.get_by_text("Continue").click(timeout=5000)
        except PlaywrightTimeoutError:
            pass


def apply_phs_filter(page, phs_value):
    print(f"🔍 PHS ACCESSION → {phs_value}")
    page.get_by_text("PHS ACCESSION").click()
    page.locator("input[placeholder*='phs']").fill(phs_value)
    time.sleep(2)
    page.get_by_text(phs_value).first.click()
    time.sleep(3)


def select_study_name(page, study_name):
    print(f"🔍 Study Name → {study_name}")
    open_study_filter(page)
    pop = active_filter_popover(page)
    check_filter_option_checkbox(page, pop or page, study_name)
    page.keyboard.press("Escape")
    time.sleep(2.5)


def open_cancer_nanotech_subfacet(
    page,
    facet_header: Locator,
    facet_id_for_error: str,
    *,
    allow_nanomaterial_entity: bool = False,
) -> None:
    """Expand a subfacet (e.g. Protocol Name, Protocol Type) under Cancer Nanotechnology."""
    try:
        facet_header.wait_for(state="visible", timeout=12000)
    except PlaywrightTimeoutError as e:
        raise RuntimeError(
            f"Could not find facet {facet_id_for_error!r} in the {CANCER_NANO_FACET} section."
        ) from e
    facet_header.scroll_into_view_if_needed()
    expand_sidebar_facet_header(
        page, facet_header, allow_nanomaterial_entity=allow_nanomaterial_entity
    )
    try:
        page.locator("//p[contains(@class,'MuiTypography-body1')]").first.wait_for(
            timeout=10000
        )
    except PlaywrightTimeoutError:
        pass
    time.sleep(1.0)


def open_protocol_name_facet_from_locator(page, protocol_header: Locator) -> None:
    open_cancer_nanotech_subfacet(page, protocol_header, PROTOCOL_NAME_FACET)


def _open_protocol_name_facet_under_cancer_nanotechnology(page):
    """
    Click the <div id="Protocol Name"> in the same UI section as Cancer Nanotechnology
    (scoped by document order relative to the CN header, not a global following:: sweep).
    """
    cn = cancer_nanotechnology_filter_trigger(page)
    proto = _protocol_name_locator_in_cn_section(cn)
    open_protocol_name_facet_from_locator(page, proto)


def select_cancer_nanotechnology_and_protocol(page, protocol_label):
    """
    1) Expand Cancer Nanotechnology in the left sidebar (section header only — QA2 has no
       separate checkbox row for that label; children e.g. PROTOCOL NAME appear underneath).
    2) Expand PROTOCOL NAME in that section, then check the protocol value row.
    """
    print("📍 Locators: Cancer Nanotechnology + Protocol Name (under CN header)")
    L = collect_cancer_nanotech_protocol_locators(page)
    print(
        f"   • filter sidebar: {'yes' if L.sidebar is not None else 'no (full-page scope)'}"
    )
    print(f"   • cn_header: {CANCER_NANO_FACET!r}")
    print(f"   • protocol_name_header: div#{PROTOCOL_NAME_FACET!r} under CN")

    print(f"🔍 {CANCER_NANO_FACET} (expand sidebar section)")
    open_cancer_from_locator(page, L.cn_header)
    pop = active_filter_popover(page)
    try:
        if pop is not None and pop.is_visible():
            page.keyboard.press("Escape")
            time.sleep(0.35)
    except PlaywrightError:
        pass
    time.sleep(0.6)

    collapse_nanomaterial_entity_subfacet_under_cn(page, L.cn_header)

    print(f"🔍 {PROTOCOL_NAME_FACET} (under {CANCER_NANO_FACET}) → {protocol_label}")
    open_protocol_name_facet_from_locator(page, L.protocol_name_header)
    collapse_nanomaterial_entity_subfacet_under_cn(page, L.cn_header)
    scope_pr = resolve_filter_list_scope_prefer_accordion(page, L.protocol_name_header)
    _scroll_filter_list_top(scope_pr)
    time.sleep(0.35)
    check_filter_option_checkbox(page, scope_pr, protocol_label)
    pop2 = active_filter_popover(page)
    try:
        if pop2 is not None and pop2.is_visible():
            page.keyboard.press("Escape")
            time.sleep(0.35)
    except PlaywrightError:
        pass
    time.sleep(2.0)


def uncheck_cpmv_and_select_protocol_type_endotoxin(page) -> None:
    """
    After CPMV scenario: open Protocol Name, uncheck CPMV row; open Protocol Type, check Endotoxin.
    """
    print(f"🔍 Unselect {PROTOCOL_NAME_FACET} → {PROTOCOL_NAME_VALUE!r}")
    L = collect_cancer_nanotech_protocol_locators(page)
    collapse_nanomaterial_entity_subfacet_under_cn(page, L.cn_header)
    open_cancer_nanotech_subfacet(page, L.protocol_name_header, PROTOCOL_NAME_FACET)
    collapse_nanomaterial_entity_subfacet_under_cn(page, L.cn_header)
    scope_pn = resolve_filter_list_scope_prefer_accordion(page, L.protocol_name_header)
    _scroll_filter_list_top(scope_pn)
    time.sleep(0.4)
    if not toggle_facet_checkbox_by_input_id(
        page, PROTOCOL_NAME_FACET, PROTOCOL_NAME_VALUE, want_checked=False
    ):
        uncheck_filter_option_checkbox(page, scope_pn, PROTOCOL_NAME_VALUE)
    else:
        print("   (Unchecked CPMV via checkbox id.)")
    pop = active_filter_popover(page)
    try:
        if pop is not None and pop.is_visible():
            page.keyboard.press("Escape")
            time.sleep(0.35)
    except PlaywrightError:
        pass
    time.sleep(0.6)

    print(f"🔍 {PROTOCOL_TYPE_FACET} → {PROTOCOL_TYPE_VALUE!r}")
    cn = cancer_nanotechnology_filter_trigger(page)
    pt_header = _protocol_type_locator_in_cn_section(cn)
    open_cancer_nanotech_subfacet(page, pt_header, PROTOCOL_TYPE_FACET)
    collapse_nanomaterial_entity_subfacet_under_cn(page, cn)
    scope_pt = resolve_filter_list_scope_prefer_accordion(page, pt_header)
    _scroll_filter_list_top(scope_pt)
    time.sleep(0.4)
    if not toggle_facet_checkbox_by_input_id(
        page, PROTOCOL_TYPE_FACET, PROTOCOL_TYPE_VALUE, want_checked=True
    ):
        check_filter_option_checkbox(page, scope_pt, PROTOCOL_TYPE_VALUE)
    else:
        print(f"   (Checked {PROTOCOL_TYPE_VALUE!r} via checkbox id.)")
    pop2 = active_filter_popover(page)
    try:
        if pop2 is not None and pop2.is_visible():
            page.keyboard.press("Escape")
            time.sleep(0.35)
    except PlaywrightError:
        pass
    time.sleep(2.0)


def uncheck_protocol_type_endotoxin(page) -> None:
    """Clear Protocol Type → Endotoxin under Cancer Nanotechnology (no other CN checkbox filters)."""
    print(f"🔍 Unselect {PROTOCOL_TYPE_FACET} → {PROTOCOL_TYPE_VALUE!r}")
    L = collect_cancer_nanotech_protocol_locators(page)
    collapse_nanomaterial_entity_subfacet_under_cn(page, L.cn_header)
    cn = cancer_nanotechnology_filter_trigger(page)
    pt_header = _protocol_type_locator_in_cn_section(cn)
    open_cancer_nanotech_subfacet(page, pt_header, PROTOCOL_TYPE_FACET)
    collapse_nanomaterial_entity_subfacet_under_cn(page, cn)
    scope_pt = resolve_filter_list_scope_prefer_accordion(page, pt_header)
    _scroll_filter_list_top(scope_pt)
    time.sleep(0.4)
    if not toggle_facet_checkbox_by_input_id(
        page, PROTOCOL_TYPE_FACET, PROTOCOL_TYPE_VALUE, want_checked=False
    ):
        uncheck_filter_option_checkbox(page, scope_pt, PROTOCOL_TYPE_VALUE)
    else:
        print(f"   (Unchecked {PROTOCOL_TYPE_VALUE!r} via checkbox id.)")
    pop = active_filter_popover(page)
    try:
        if pop is not None and pop.is_visible():
            page.keyboard.press("Escape")
            time.sleep(0.35)
    except PlaywrightError:
        pass
    time.sleep(2.0)


def select_doi_under_cancer_nanotechnology(page, doi_value: str) -> None:
    """Expand DOI under CN and check the given DOI option (stable id or checkbox fallback)."""
    print(f"🔍 {DOI_FACET} (under {CANCER_NANO_FACET}) → {doi_value!r}")
    L = collect_cancer_nanotech_protocol_locators(page)
    collapse_nanomaterial_entity_subfacet_under_cn(page, L.cn_header)
    cn = cancer_nanotechnology_filter_trigger(page)
    doi_header = _doi_locator_in_cn_section(cn)
    open_cancer_nanotech_subfacet(page, doi_header, DOI_FACET)
    collapse_nanomaterial_entity_subfacet_under_cn(page, cn)
    scope_doi = resolve_filter_list_scope_prefer_accordion(page, doi_header)
    _scroll_filter_list_top(scope_doi)
    time.sleep(0.4)
    if not toggle_facet_checkbox_by_input_id(
        page, DOI_FACET, doi_value, want_checked=True
    ):
        check_filter_option_checkbox(page, scope_doi, doi_value)
    else:
        print(f"   (Checked {doi_value!r} via checkbox id.)")
    pop = active_filter_popover(page)
    try:
        if pop is not None and pop.is_visible():
            page.keyboard.press("Escape")
            time.sleep(0.35)
    except PlaywrightError:
        pass
    time.sleep(2.0)


def uncheck_doi_under_cancer_nanotechnology(page, doi_value: str) -> None:
    """Open DOI under CN and uncheck the given option."""
    print(f"🔍 Unselect {DOI_FACET} → {doi_value!r}")
    L = collect_cancer_nanotech_protocol_locators(page)
    collapse_nanomaterial_entity_subfacet_under_cn(page, L.cn_header)
    cn = cancer_nanotechnology_filter_trigger(page)
    doi_header = _doi_locator_in_cn_section(cn)
    open_cancer_nanotech_subfacet(page, doi_header, DOI_FACET)
    collapse_nanomaterial_entity_subfacet_under_cn(page, cn)
    scope_doi = resolve_filter_list_scope_prefer_accordion(page, doi_header)
    _scroll_filter_list_top(scope_doi)
    time.sleep(0.4)
    if not toggle_facet_checkbox_by_input_id(
        page, DOI_FACET, doi_value, want_checked=False
    ):
        uncheck_filter_option_checkbox(page, scope_doi, doi_value)
    else:
        print(f"   (Unchecked {doi_value!r} via checkbox id.)")
    pop = active_filter_popover(page)
    try:
        if pop is not None and pop.is_visible():
            page.keyboard.press("Escape")
            time.sleep(0.35)
    except PlaywrightError:
        pass
    time.sleep(2.0)


def select_publication_title_under_cancer_nanotechnology(
    page, title_value: str
) -> None:
    """Expand Publication Title under CN and check the given option."""
    print(f"🔍 {PUBLICATION_TITLE_FACET} (under {CANCER_NANO_FACET}) → {title_value!r}")
    L = collect_cancer_nanotech_protocol_locators(page)
    collapse_nanomaterial_entity_subfacet_under_cn(page, L.cn_header)
    cn = cancer_nanotechnology_filter_trigger(page)
    pt_hdr = _publication_title_locator_in_cn_section(cn)
    open_cancer_nanotech_subfacet(page, pt_hdr, PUBLICATION_TITLE_FACET)
    collapse_nanomaterial_entity_subfacet_under_cn(page, cn)
    scope_pt = resolve_filter_list_scope_prefer_accordion(page, pt_hdr)
    _scroll_filter_list_top(scope_pt)
    time.sleep(0.4)
    if not toggle_facet_checkbox_by_input_id(
        page, PUBLICATION_TITLE_FACET, title_value, want_checked=True
    ):
        check_filter_option_checkbox(page, scope_pt, title_value)
    else:
        print("   (Checked publication title via checkbox id.)")
    pop = active_filter_popover(page)
    try:
        if pop is not None and pop.is_visible():
            page.keyboard.press("Escape")
            time.sleep(0.35)
    except PlaywrightError:
        pass
    time.sleep(2.0)


def uncheck_publication_title_under_cancer_nanotechnology(
    page, title_value: str
) -> None:
    """Open Publication Title under CN and uncheck the given option."""
    print(f"🔍 Unselect {PUBLICATION_TITLE_FACET} → {title_value!r}")
    L = collect_cancer_nanotech_protocol_locators(page)
    collapse_nanomaterial_entity_subfacet_under_cn(page, L.cn_header)
    cn = cancer_nanotechnology_filter_trigger(page)
    pt_hdr = _publication_title_locator_in_cn_section(cn)
    open_cancer_nanotech_subfacet(page, pt_hdr, PUBLICATION_TITLE_FACET)
    collapse_nanomaterial_entity_subfacet_under_cn(page, cn)
    scope_pt = resolve_filter_list_scope_prefer_accordion(page, pt_hdr)
    _scroll_filter_list_top(scope_pt)
    time.sleep(0.4)
    if not toggle_facet_checkbox_by_input_id(
        page, PUBLICATION_TITLE_FACET, title_value, want_checked=False
    ):
        uncheck_filter_option_checkbox(page, scope_pt, title_value)
    else:
        print("   (Unchecked publication title via checkbox id.)")
    pop = active_filter_popover(page)
    try:
        if pop is not None and pop.is_visible():
            page.keyboard.press("Escape")
            time.sleep(0.35)
    except PlaywrightError:
        pass
    time.sleep(2.0)


def select_nanomaterial_entity_type_under_cancer_nanotechnology(
    page, entity_type_label: str
) -> None:
    """
    Expand Nanomaterial Entity under CN and check the given entity type (e.g. aptamer).

    Uses allow_nanomaterial_entity on the header click only for this scenario.
    Does not collapse the NM facet after opening — that would close the option list.
    """
    print(
        f"🔍 {NANOMATERIAL_ENTITY_DIV_ID} (under {CANCER_NANO_FACET}) → "
        f"{entity_type_label!r}"
    )
    L = collect_cancer_nanotech_protocol_locators(page)
    collapse_nanomaterial_entity_subfacet_under_cn(page, L.cn_header)
    cn = cancer_nanotechnology_filter_trigger(page)
    nm_hdr = _nanomaterial_entity_facet_header_in_cn_section(cn)
    open_cancer_nanotech_subfacet(
        page,
        nm_hdr,
        NANOMATERIAL_ENTITY_DIV_ID,
        allow_nanomaterial_entity=True,
    )
    scope_nm = resolve_filter_list_scope_prefer_accordion(page, nm_hdr)
    _scroll_filter_list_top(scope_nm)
    time.sleep(0.4)
    toggled = False
    for fk in (NANOMATERIAL_ENTITY_DIV_ID, "Nanomaterial entity"):
        if toggle_facet_checkbox_by_input_id(
            page, fk, entity_type_label, want_checked=True
        ):
            toggled = True
            print(f"   (Checked {entity_type_label!r} via checkbox id, facet {fk!r}.)")
            break
    if not toggled:
        check_filter_option_checkbox(page, scope_nm, entity_type_label)
    pop = active_filter_popover(page)
    try:
        if pop is not None and pop.is_visible():
            page.keyboard.press("Escape")
            time.sleep(0.35)
    except PlaywrightError:
        pass
    time.sleep(2.0)


def uncheck_nanomaterial_entity_type_under_cancer_nanotechnology(
    page, entity_type_label: str
) -> None:
    """Expand Nanomaterial Entity under CN and uncheck the given entity type."""
    print(
        f"🔍 Unselect {NANOMATERIAL_ENTITY_DIV_ID} → {entity_type_label!r}"
    )
    L = collect_cancer_nanotech_protocol_locators(page)
    collapse_nanomaterial_entity_subfacet_under_cn(page, L.cn_header)
    cn = cancer_nanotechnology_filter_trigger(page)
    nm_hdr = _nanomaterial_entity_facet_header_in_cn_section(cn)
    open_cancer_nanotech_subfacet(
        page,
        nm_hdr,
        NANOMATERIAL_ENTITY_DIV_ID,
        allow_nanomaterial_entity=True,
    )
    scope_nm = resolve_filter_list_scope_prefer_accordion(page, nm_hdr)
    _scroll_filter_list_top(scope_nm)
    time.sleep(0.4)
    toggled = False
    for fk in (NANOMATERIAL_ENTITY_DIV_ID, "Nanomaterial entity"):
        if toggle_facet_checkbox_by_input_id(
            page, fk, entity_type_label, want_checked=False
        ):
            toggled = True
            print(
                f"   (Unchecked {entity_type_label!r} via checkbox id, facet {fk!r}.)"
            )
            break
    if not toggled:
        uncheck_filter_option_checkbox(page, scope_nm, entity_type_label)
    pop = active_filter_popover(page)
    try:
        if pop is not None and pop.is_visible():
            page.keyboard.press("Escape")
            time.sleep(0.35)
    except PlaywrightError:
        pass
    time.sleep(2.0)


def select_functionalizing_entity_under_cancer_nanotechnology(
    page, option_label: str
) -> None:
    """Expand Functionalizing Entity under CN and check the given option (e.g. Antibody)."""
    print(
        f"🔍 {FUNCTIONALIZING_ENTITY_FACET} (under {CANCER_NANO_FACET}) → "
        f"{option_label!r}"
    )
    L = collect_cancer_nanotech_protocol_locators(page)
    collapse_nanomaterial_entity_subfacet_under_cn(page, L.cn_header)
    cn = cancer_nanotechnology_filter_trigger(page)
    fe_hdr = _functionalizing_entity_locator_in_cn_section(cn)
    open_cancer_nanotech_subfacet(page, fe_hdr, FUNCTIONALIZING_ENTITY_FACET)
    collapse_nanomaterial_entity_subfacet_under_cn(page, cn)
    scope_fe = resolve_filter_list_scope_prefer_accordion(page, fe_hdr)
    _scroll_filter_list_top(scope_fe)
    time.sleep(0.4)
    toggled = False
    for fk in (FUNCTIONALIZING_ENTITY_FACET, "Functionalizing entity"):
        if toggle_facet_checkbox_by_input_id(
            page, fk, option_label, want_checked=True
        ):
            toggled = True
            print(
                f"   (Checked {option_label!r} via checkbox id, facet {fk!r}.)"
            )
            break
    if not toggled:
        check_filter_option_checkbox(page, scope_fe, option_label)
    pop = active_filter_popover(page)
    try:
        if pop is not None and pop.is_visible():
            page.keyboard.press("Escape")
            time.sleep(0.35)
    except PlaywrightError:
        pass
    time.sleep(2.0)


def uncheck_functionalizing_entity_under_cancer_nanotechnology(
    page, option_label: str
) -> None:
    """Open Functionalizing Entity under CN and uncheck the given option."""
    print(f"🔍 Unselect {FUNCTIONALIZING_ENTITY_FACET} → {option_label!r}")
    L = collect_cancer_nanotech_protocol_locators(page)
    collapse_nanomaterial_entity_subfacet_under_cn(page, L.cn_header)
    cn = cancer_nanotechnology_filter_trigger(page)
    fe_hdr = _functionalizing_entity_locator_in_cn_section(cn)
    open_cancer_nanotech_subfacet(page, fe_hdr, FUNCTIONALIZING_ENTITY_FACET)
    collapse_nanomaterial_entity_subfacet_under_cn(page, cn)
    scope_fe = resolve_filter_list_scope_prefer_accordion(page, fe_hdr)
    _scroll_filter_list_top(scope_fe)
    time.sleep(0.4)
    toggled = False
    for fk in (FUNCTIONALIZING_ENTITY_FACET, "Functionalizing entity"):
        if toggle_facet_checkbox_by_input_id(
            page, fk, option_label, want_checked=False
        ):
            toggled = True
            print(
                f"   (Unchecked {option_label!r} via checkbox id, facet {fk!r}.)"
            )
            break
    if not toggled:
        uncheck_filter_option_checkbox(page, scope_fe, option_label)
    pop = active_filter_popover(page)
    try:
        if pop is not None and pop.is_visible():
            page.keyboard.press("Escape")
            time.sleep(0.35)
    except PlaywrightError:
        pass
    time.sleep(2.0)


def select_characterization_type_under_cancer_nanotechnology(
    page, option_label: str
) -> None:
    """Expand Characterization Type under CN and check the given option."""
    print(
        f"🔍 {CHARACTERIZATION_TYPE_FACET} (under {CANCER_NANO_FACET}) → "
        f"{option_label!r}"
    )
    L = collect_cancer_nanotech_protocol_locators(page)
    collapse_nanomaterial_entity_subfacet_under_cn(page, L.cn_header)
    cn = cancer_nanotechnology_filter_trigger(page)
    ct_hdr = _characterization_type_locator_in_cn_section(cn)
    open_cancer_nanotech_subfacet(page, ct_hdr, CHARACTERIZATION_TYPE_FACET)
    collapse_nanomaterial_entity_subfacet_under_cn(page, cn)
    scope_ct = resolve_filter_list_scope_prefer_accordion(page, ct_hdr)
    _scroll_filter_list_top(scope_ct)
    time.sleep(0.4)
    toggled = False
    for val in dict.fromkeys(
        (option_label, "clinical trial", "Clinical Trial")
    ):
        for fk in (CHARACTERIZATION_TYPE_FACET, "Characterization type"):
            if toggle_facet_checkbox_by_input_id(
                page, fk, val, want_checked=True
            ):
                toggled = True
                print(
                    f"   (Checked {val!r} via checkbox id, facet {fk!r}.)"
                )
                break
        if toggled:
            break
    if not toggled:
        check_filter_option_checkbox(page, scope_ct, option_label)
    pop = active_filter_popover(page)
    try:
        if pop is not None and pop.is_visible():
            page.keyboard.press("Escape")
            time.sleep(0.35)
    except PlaywrightError:
        pass
    time.sleep(2.0)


def uncheck_characterization_type_under_cancer_nanotechnology(
    page, option_label: str
) -> None:
    """Open Characterization Type under CN and uncheck the given option (id may be lower/title case)."""
    print(f"🔍 Unselect {CHARACTERIZATION_TYPE_FACET} → {option_label!r}")
    L = collect_cancer_nanotech_protocol_locators(page)
    collapse_nanomaterial_entity_subfacet_under_cn(page, L.cn_header)
    cn = cancer_nanotechnology_filter_trigger(page)
    ct_hdr = _characterization_type_locator_in_cn_section(cn)
    open_cancer_nanotech_subfacet(page, ct_hdr, CHARACTERIZATION_TYPE_FACET)
    collapse_nanomaterial_entity_subfacet_under_cn(page, cn)
    scope_ct = resolve_filter_list_scope_prefer_accordion(page, ct_hdr)
    _scroll_filter_list_top(scope_ct)
    time.sleep(0.4)
    toggled = False
    for val in dict.fromkeys(
        (option_label, "clinical trial", "Clinical Trial")
    ):
        for fk in (CHARACTERIZATION_TYPE_FACET, "Characterization type"):
            if toggle_facet_checkbox_by_input_id(
                page, fk, val, want_checked=False
            ):
                toggled = True
                print(
                    f"   (Unchecked {val!r} via checkbox id, facet {fk!r}.)"
                )
                break
        if toggled:
            break
    if not toggled:
        uncheck_filter_option_checkbox(page, scope_ct, option_label)
    pop = active_filter_popover(page)
    try:
        if pop is not None and pop.is_visible():
            page.keyboard.press("Escape")
            time.sleep(0.35)
    except PlaywrightError:
        pass
    time.sleep(2.0)


def select_characterization_name_under_cancer_nanotechnology(
    page, option_label: str
) -> None:
    """Expand Characterization Name under CN and check the given option."""
    print(
        f"🔍 {CHARACTERIZATION_NAME_FACET} (under {CANCER_NANO_FACET}) → "
        f"{option_label!r}"
    )
    L = collect_cancer_nanotech_protocol_locators(page)
    collapse_nanomaterial_entity_subfacet_under_cn(page, L.cn_header)
    cn = cancer_nanotechnology_filter_trigger(page)
    cname_hdr = _characterization_name_locator_in_cn_section(cn)
    open_cancer_nanotech_subfacet(page, cname_hdr, CHARACTERIZATION_NAME_FACET)
    collapse_nanomaterial_entity_subfacet_under_cn(page, cn)
    scope_cname = resolve_filter_list_scope_prefer_accordion(page, cname_hdr)
    _scroll_filter_list_top(scope_cname)
    time.sleep(0.4)
    toggled = False
    for fk in (CHARACTERIZATION_NAME_FACET, "Characterization name"):
        if toggle_facet_checkbox_by_input_id(
            page, fk, option_label, want_checked=True
        ):
            toggled = True
            print(
                f"   (Checked {option_label!r} via checkbox id, facet {fk!r}.)"
            )
            break
    if not toggled:
        check_filter_option_checkbox(page, scope_cname, option_label)
    pop = active_filter_popover(page)
    try:
        if pop is not None and pop.is_visible():
            page.keyboard.press("Escape")
            time.sleep(0.35)
    except PlaywrightError:
        pass
    time.sleep(2.0)


def run() -> None:
    wb_path = os.environ.get("CANANOLAB_BASECOUNTS_PATH", "").strip()
    workbook = Path(wb_path).expanduser() if wb_path else DEFAULT_BASECOUNTS_WORKBOOK
    print(f"📎 Loading base counts from: {workbook}")
    bc = load_cananolab_basecounts_workbook(str(workbook))

    url = _data_commons_url()
    print(f"🌐 Data Commons URL: {url}")

    with sync_playwright() as p:
        browser = p.chromium.launch(channel="chrome", headless=False)
        page = browser.new_page()
        try:
            page.goto(url, timeout=120_000)
            dismiss_continue_if_present(page)
            page.wait_for_load_state("networkidle")

            print(f"Opened {url}")
            apply_phs_filter(page, PHS_ACCESSION)
            select_study_name(page, STUDY_NAME)
            print("✅ Filters applied: PHS and study name.")

            study_checks = [
                ("Samples", bc[BC_STUDY]["samples"]),
                ("Files", bc[BC_STUDY]["files"]),
                ("Protocols", bc[BC_STUDY]["protocols"]),
            ]
            study_results = []
            for label, base in study_checks:
                study_results.append(verify_tab_table_and_base(page, label, base))

            select_cancer_nanotechnology_and_protocol(page, PROTOCOL_NAME_VALUE)
            print(
                "✅ Filters applied: CN sidebar expanded + PROTOCOL NAME + protocol checkbox."
            )

            cpmv_checks = [
                ("Participants", bc[BC_CPMV]["participants"]),
                ("Samples", bc[BC_CPMV]["samples"]),
                ("Files", bc[BC_CPMV]["files"]),
                ("Protocols", bc[BC_CPMV]["protocols"]),
            ]
            cpmv_results = []
            for label, base in cpmv_checks:
                cpmv_results.append(verify_tab_table_and_base(page, label, base))

            uncheck_cpmv_and_select_protocol_type_endotoxin(page)
            print(
                "✅ Filters applied: CPMV unchecked; "
                f"{PROTOCOL_TYPE_FACET} → {PROTOCOL_TYPE_VALUE}."
            )

            endotoxin_checks = [
                ("Participants", bc[BC_ENDOTOXIN]["participants"]),
                ("Samples", bc[BC_ENDOTOXIN]["samples"]),
                ("Files", bc[BC_ENDOTOXIN]["files"]),
                ("Protocols", bc[BC_ENDOTOXIN]["protocols"]),
            ]
            endotoxin_results = []
            for label, base in endotoxin_checks:
                endotoxin_results.append(verify_tab_table_and_base(page, label, base))

            uncheck_protocol_type_endotoxin(page)
            print(
                f"✅ Filters applied: {PROTOCOL_TYPE_FACET} → {PROTOCOL_TYPE_VALUE} cleared."
            )

            select_doi_under_cancer_nanotechnology(page, DOI_VALUE)
            print(
                f"✅ Filters applied: {CANCER_NANO_FACET} + {DOI_FACET} → {DOI_VALUE}."
            )

            scenario4_checks = [
                ("Participants", bc[BC_DOI]["participants"]),
                ("Samples", bc[BC_DOI]["samples"]),
                ("Files", bc[BC_DOI]["files"]),
                ("Protocols", bc[BC_DOI]["protocols"]),
            ]
            scenario4_results = []
            for label, base in scenario4_checks:
                scenario4_results.append(verify_tab_table_and_base(page, label, base))

            uncheck_doi_under_cancer_nanotechnology(page, DOI_VALUE)
            print(f"✅ Filters applied: {DOI_FACET} → {DOI_VALUE} cleared.")

            select_publication_title_under_cancer_nanotechnology(
                page, PUBLICATION_TITLE_VALUE
            )
            print(
                f"✅ Filters applied: {CANCER_NANO_FACET} + "
                f"{PUBLICATION_TITLE_FACET} → (selected title)."
            )

            scenario5_checks = [
                ("Participants", bc[BC_PUBLICATION_TITLE]["participants"]),
                ("Samples", bc[BC_PUBLICATION_TITLE]["samples"]),
                ("Files", bc[BC_PUBLICATION_TITLE]["files"]),
                ("Protocols", bc[BC_PUBLICATION_TITLE]["protocols"]),
            ]
            scenario5_results = []
            for label, base in scenario5_checks:
                scenario5_results.append(verify_tab_table_and_base(page, label, base))

            uncheck_publication_title_under_cancer_nanotechnology(
                page, PUBLICATION_TITLE_VALUE
            )
            print(
                f"✅ Filters applied: {PUBLICATION_TITLE_FACET} → (title) cleared."
            )

            select_nanomaterial_entity_type_under_cancer_nanotechnology(
                page, NANOMATERIAL_ENTITY_TYPE_VALUE
            )
            print(
                f"✅ Filters applied: {CANCER_NANO_FACET} + "
                f"{NANOMATERIAL_ENTITY_DIV_ID} → {NANOMATERIAL_ENTITY_TYPE_VALUE!r}."
            )

            scenario6_checks = [
                ("Participants", bc[BC_NM_APTAMER]["participants"]),
                ("Samples", bc[BC_NM_APTAMER]["samples"]),
                ("Files", bc[BC_NM_APTAMER]["files"]),
                ("Protocols", bc[BC_NM_APTAMER]["protocols"]),
            ]
            scenario6_results = []
            for label, base in scenario6_checks:
                scenario6_results.append(verify_tab_table_and_base(page, label, base))

            uncheck_nanomaterial_entity_type_under_cancer_nanotechnology(
                page, NANOMATERIAL_ENTITY_TYPE_VALUE
            )
            print(
                f"✅ Filters applied: {NANOMATERIAL_ENTITY_DIV_ID} → "
                f"{NANOMATERIAL_ENTITY_TYPE_VALUE!r} cleared."
            )

            select_functionalizing_entity_under_cancer_nanotechnology(
                page, FUNCTIONALIZING_ENTITY_VALUE
            )
            print(
                f"✅ Filters applied: {CANCER_NANO_FACET} + "
                f"{FUNCTIONALIZING_ENTITY_FACET} → {FUNCTIONALIZING_ENTITY_VALUE!r}."
            )

            scenario7_checks = [
                ("Participants", bc[BC_FE_ANTIBODY]["participants"]),
                ("Samples", bc[BC_FE_ANTIBODY]["samples"]),
                ("Files", bc[BC_FE_ANTIBODY]["files"]),
                ("Protocols", bc[BC_FE_ANTIBODY]["protocols"]),
            ]
            scenario7_results = []
            for label, base in scenario7_checks:
                scenario7_results.append(verify_tab_table_and_base(page, label, base))

            uncheck_functionalizing_entity_under_cancer_nanotechnology(
                page, FUNCTIONALIZING_ENTITY_VALUE
            )
            print(
                f"✅ Filters applied: {FUNCTIONALIZING_ENTITY_FACET} → "
                f"{FUNCTIONALIZING_ENTITY_VALUE!r} cleared."
            )

            select_characterization_type_under_cancer_nanotechnology(
                page, CHARACTERIZATION_TYPE_VALUE
            )
            print(
                f"✅ Filters applied: {CANCER_NANO_FACET} + "
                f"{CHARACTERIZATION_TYPE_FACET} → {CHARACTERIZATION_TYPE_VALUE!r}."
            )

            scenario8_checks = [
                ("Participants", bc[BC_CLINICAL_TRIAL]["participants"]),
                ("Samples", bc[BC_CLINICAL_TRIAL]["samples"]),
                ("Files", bc[BC_CLINICAL_TRIAL]["files"]),
                ("Protocols", bc[BC_CLINICAL_TRIAL]["protocols"]),
            ]
            scenario8_results = []
            for label, base in scenario8_checks:
                scenario8_results.append(verify_tab_table_and_base(page, label, base))

            uncheck_characterization_type_under_cancer_nanotechnology(
                page, CHARACTERIZATION_TYPE_VALUE
            )
            print(
                f"✅ Filters applied: {CHARACTERIZATION_TYPE_FACET} → "
                f"{CHARACTERIZATION_TYPE_VALUE!r} cleared."
            )

            select_characterization_name_under_cancer_nanotechnology(
                page, CHARACTERIZATION_NAME_VALUE
            )
            print(
                f"✅ Filters applied: {CANCER_NANO_FACET} + "
                f"{CHARACTERIZATION_NAME_FACET} → {CHARACTERIZATION_NAME_VALUE!r}."
            )

            scenario9_checks = [
                ("Participants", bc[BC_CNAME_ANTITUMOR]["participants"]),
                ("Samples", bc[BC_CNAME_ANTITUMOR]["samples"]),
                ("Files", bc[BC_CNAME_ANTITUMOR]["files"]),
                ("Protocols", bc[BC_CNAME_ANTITUMOR]["protocols"]),
            ]
            scenario9_results = []
            for label, base in scenario9_checks:
                scenario9_results.append(verify_tab_table_and_base(page, label, base))

            all_ok = (
                all(r["row_ok"] for r in study_results)
                and all(r["row_ok"] for r in cpmv_results)
                and all(r["row_ok"] for r in endotoxin_results)
                and all(r["row_ok"] for r in scenario4_results)
                and all(r["row_ok"] for r in scenario5_results)
                and all(r["row_ok"] for r in scenario6_results)
                and all(r["row_ok"] for r in scenario7_results)
                and all(r["row_ok"] for r in scenario8_results)
                and all(r["row_ok"] for r in scenario9_results)
            )

            print("\n── Summary ──")
            print("✅ All checks passed." if all_ok else "❌ One or more checks failed.")

            report_sections = [
                (f"Study: {STUDY_NAME}", study_results),
                (
                    f"{CANCER_NANO_FACET} + {PROTOCOL_NAME_VALUE}",
                    cpmv_results,
                ),
                (
                    f"{CANCER_NANO_FACET} + {PROTOCOL_TYPE_VALUE} (after unchecking CPMV)",
                    endotoxin_results,
                ),
                (
                    f"{CANCER_NANO_FACET} + {DOI_FACET} → {DOI_VALUE} "
                    f"(after uncheck {PROTOCOL_TYPE_VALUE})",
                    scenario4_results,
                ),
                (
                    f"{CANCER_NANO_FACET} + {PUBLICATION_TITLE_FACET} "
                    f"(after uncheck {DOI_VALUE})",
                    scenario5_results,
                ),
                (
                    f"{CANCER_NANO_FACET} + {NANOMATERIAL_ENTITY_DIV_ID} → "
                    f"{NANOMATERIAL_ENTITY_TYPE_VALUE} (after uncheck publication title)",
                    scenario6_results,
                ),
                (
                    f"{CANCER_NANO_FACET} + {FUNCTIONALIZING_ENTITY_FACET} → "
                    f"{FUNCTIONALIZING_ENTITY_VALUE} (after uncheck aptamer)",
                    scenario7_results,
                ),
                (
                    f"{CANCER_NANO_FACET} + {CHARACTERIZATION_TYPE_FACET} → "
                    f"{CHARACTERIZATION_TYPE_VALUE} (after uncheck Antibody)",
                    scenario8_results,
                ),
                (
                    f"{CANCER_NANO_FACET} + {CHARACTERIZATION_NAME_FACET} → "
                    f"{CHARACTERIZATION_NAME_VALUE} (after uncheck clinical trial)",
                    scenario9_results,
                ),
            ]
            generate_html_report(
                report_sections,
                all_ok,
                url=url,
                phs=PHS_ACCESSION,
                study_name=STUDY_NAME,
                scenario2_summary=f"{CANCER_NANO_FACET} + {PROTOCOL_NAME_VALUE}",
                scenario3_summary=(
                    f"Uncheck {PROTOCOL_NAME_VALUE}; "
                    f"{PROTOCOL_TYPE_FACET} → {PROTOCOL_TYPE_VALUE}"
                ),
                scenario4_summary=(
                    f"Uncheck {PROTOCOL_TYPE_VALUE}; {DOI_FACET} → {DOI_VALUE}; "
                    f"expect Participants / Samples / Files / Protocols = "
                    f"{bc[BC_DOI]['participants']} / {bc[BC_DOI]['samples']} / "
                    f"{bc[BC_DOI]['files']} / {bc[BC_DOI]['protocols']}"
                ),
                scenario5_summary=(
                    f"Uncheck {DOI_VALUE}; {PUBLICATION_TITLE_FACET} → "
                    f"{PUBLICATION_TITLE_VALUE}; expect "
                    f"{bc[BC_PUBLICATION_TITLE]['participants']} / "
                    f"{bc[BC_PUBLICATION_TITLE]['samples']} / "
                    f"{bc[BC_PUBLICATION_TITLE]['files']} / "
                    f"{bc[BC_PUBLICATION_TITLE]['protocols']} "
                    f"(Participants / Samples / Files / Protocols)"
                ),
                scenario6_summary=(
                    f"Uncheck publication title; {NANOMATERIAL_ENTITY_DIV_ID} → "
                    f"{NANOMATERIAL_ENTITY_TYPE_VALUE}; expect "
                    f"{bc[BC_NM_APTAMER]['participants']} / "
                    f"{bc[BC_NM_APTAMER]['samples']} / "
                    f"{bc[BC_NM_APTAMER]['files']} / "
                    f"{bc[BC_NM_APTAMER]['protocols']} "
                    f"(Participants / Samples / Files / Protocols)"
                ),
                scenario7_summary=(
                    f"Uncheck {NANOMATERIAL_ENTITY_TYPE_VALUE}; "
                    f"{FUNCTIONALIZING_ENTITY_FACET} → {FUNCTIONALIZING_ENTITY_VALUE}; "
                    f"expect {bc[BC_FE_ANTIBODY]['participants']} / "
                    f"{bc[BC_FE_ANTIBODY]['samples']} / "
                    f"{bc[BC_FE_ANTIBODY]['files']} / "
                    f"{bc[BC_FE_ANTIBODY]['protocols']} "
                    f"(Participants / Samples / Files / Protocols)"
                ),
                scenario8_summary=(
                    f"Uncheck {FUNCTIONALIZING_ENTITY_VALUE}; "
                    f"{CHARACTERIZATION_TYPE_FACET} → clinical trial "
                    f"({CHARACTERIZATION_TYPE_VALUE!r} in UI); "
                    f"expect {bc[BC_CLINICAL_TRIAL]['participants']} / "
                    f"{bc[BC_CLINICAL_TRIAL]['samples']} / "
                    f"{bc[BC_CLINICAL_TRIAL]['files']} / "
                    f"{bc[BC_CLINICAL_TRIAL]['protocols']} "
                    f"(Participants / Samples / Files / Protocols)"
                ),
                scenario9_summary=(
                    f"Uncheck clinical trial; {CHARACTERIZATION_NAME_FACET} → "
                    f"{CHARACTERIZATION_NAME_VALUE}; expect "
                    f"{bc[BC_CNAME_ANTITUMOR]['participants']} / "
                    f"{bc[BC_CNAME_ANTITUMOR]['samples']} / "
                    f"{bc[BC_CNAME_ANTITUMOR]['files']} / "
                    f"{bc[BC_CNAME_ANTITUMOR]['protocols']} "
                    f"(Participants / Samples / Files / Protocols)"
                ),
            )
            print(f"\n📄 HTML report written: {REPORT_PATH}")
        finally:
            browser.close()


if __name__ == "__main__":
    run()
