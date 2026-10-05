#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
VCP Presentation Generator
==========================
Deterministically builds the VCP Management View PowerPoint for ONE function
from the VCP Execution Excel tracker, using the approved template as the only
design source.  No value is invented: every string written to the deck is
either a verbatim Excel value, a template label, or a documented derivation
(counts, averages, signals) that is listed in the validation summary.

Sub-commands
------------
  inspect  <xlsx> [--function F]         list sheets / parse a function sheet and report issues (no deck)
  build    <xlsx> --function F [opts]    build the deck + validation summary (+ model JSON)
  render   <pptx> [--outdir D]           export slide PNGs through PowerPoint (Windows COM) for visual QA

Exit codes: 0 ok | 2 usage/parse error | 3 blocking questions for the user (nothing written)
"""
import argparse
import copy
import datetime as dt
import json
import math
import os
import re
import subprocess
import sys
import warnings
from collections import OrderedDict
from pathlib import Path

warnings.filterwarnings("ignore")
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:  # pragma: no cover
    pass

import openpyxl
from lxml import etree
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches

HERE = Path(__file__).resolve().parent
TEMPLATE_V1 = HERE / "template" / "Service_Transformation_VCP_Management_View.pptx"   # original Management View layout
DEFAULT_TEMPLATE = HERE / "template" / "CS_VCP_Monthly_Report.pptx"                      # 'Monthly VCP report' layout (current)
_TEMPLATE_KIND = {}


def _is_v2_template(template):
    """True when the template is the 'Monthly VCP report' layout (slide 2 carries a native chart named 'Chart')."""
    key = str(Path(template).resolve())
    if key not in _TEMPLATE_KIND:
        import vcp_render2
        _TEMPLATE_KIND[key] = vcp_render2.is_v2(Presentation(key))
    return _TEMPLATE_KIND[key]

ZWSP = "​"
YEAR_RE = re.compile(r"^\s*-?(20\d\d)\s*$")
NUM_RE = re.compile(r"[-+]?\d+(?:[.,]\d+)?")
CURRENCY_RE = re.compile(r"€\s*[~<>+(]*\s*\d")
STATUS_CANON = {
    "completed": "Completed", "complete": "Completed", "done": "Completed", "closed": "Completed",
    "in progress": "In progress", "inprogress": "In progress", "in-progress": "In progress",
    "ongoing": "In progress", "on going": "In progress", "wip": "In progress",
    "not started": "Not started", "notstarted": "Not started", "not yet started": "Not started",
    "delayed": "Delayed", "delay": "Delayed", "late": "Delayed",
    "on hold": "On hold", "onhold": "On hold", "hold": "On hold",
}
LOWER_BETTER_HINTS = ("reduction", "reduce", "time", "days", "cost", "inventory", "complaint",
                      "lead", "mat ", "mat%", "copq", "conq", "qn", "nc ", "failure", "mcos",
                      "icos", "cycle", "lost", "hold", "delay", "warranty", "fco", "3rd")

# Template palette (taken from the approved deck, never changed)
PALETTE = {
    "Below target":      {"fill": "F8D7DA", "font": "C7372F"},
    "Delayed":           {"fill": "F8D7DA", "font": "C7372F"},
    "On / above target": {"fill": "DCEEDC", "font": "0E7A3D"},
    "Completed":         {"fill": "DCEEDC", "font": "0E7A3D"},
    "In progress":       {"fill": "FFF2CC", "font": "916400"},
    "No data":           {"fill": "ECEFF3", "font": "525866"},
    "Direction to confirm": {"fill": "FFF2CC", "font": "916400"},
    "Not started":       {"fill": "ECEFF3", "font": "525866"},
    "On hold":           {"fill": "ECEFF3", "font": "525866"},
}
COMPACT_ROW_H = 436245  # EMU, body-row height of the themes table; applied to the impact table beyond 5 rows
NOT_AVAILABLE = "Not available"


# --------------------------------------------------------------------------- helpers
def clean(v):
    """Normalise a cell value: strip zero-width spaces, collapse whitespace. Numbers/dates untouched."""
    if v is None:
        return None
    if isinstance(v, str):
        s = " ".join(v.replace(ZWSP, "").split())
        return s if s else None
    return v


def canon_status(v, issues, where):
    s = clean(v)
    if s is None:
        return None
    if not isinstance(s, str):
        issues.append(("ambiguous", f"{where}: status cell holds a non-text value {s!r}; kept verbatim"))
        return str(s)
    c = STATUS_CANON.get(s.lower())
    if c is None:
        issues.append(("ambiguous", f"{where}: unrecognised status {s!r}; kept verbatim (no colour rule)"))
        return s
    return c


def is_year(v):
    if v is None:
        return None
    m = YEAR_RE.match(str(v))
    return int(m.group(1)) if m else None


def a1(ws, r, c):
    return ws.cell(r, c).coordinate


def parse_number(v):
    """Extract a comparable number from a cell (numeric or text like '€8.4M', '23%'). None if not parseable."""
    if v is None:
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        # only a single plain number, optionally with € / % / M decoration, is comparable
        # (ranges, '~', '>', slashes or free text are not)
        s = v.replace(ZWSP, "").strip()
        m = re.fullmatch(r"€?\s*([-+]?\d+(?:[.,]\d+)?)\s*(?:%|M|M€|€M|pts)?", s)
        if not m:
            return None
        try:
            return float(m.group(1).replace(",", "."))
        except ValueError:
            return None
    return None


def trim_num(x, max_dec=2):
    s = f"{x:.{max_dec}f}".rstrip("0").rstrip(".")
    return s if s not in ("", "-0") else "0"


def fmt_date(v):
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime("%d %b %Y")
    return str(v)


def json_default(o):
    if isinstance(o, (dt.datetime, dt.date)):
        return o.isoformat()
    return str(o)


# --------------------------------------------------------------------------- Excel extraction
def _fast_workbook_copy(path):
    """Copy of the workbook with all <mergeCells> removed. The generator only reads cell values
    (a merged range keeps its value in the top-left cell either way), but openpyxl re-styles every
    cell of every merge on load - ~10,000 merges in the tracker turn a 2 s load into 1-2 minutes."""
    import hashlib
    import tempfile
    import zipfile
    src = Path(path)
    st = src.stat()
    key = hashlib.md5(f"{src.resolve()}|{st.st_size}|{st.st_mtime_ns}".encode("utf-8")).hexdigest()[:16]
    cache = Path(tempfile.gettempdir()) / "vcp_deck_cache"
    cache.mkdir(exist_ok=True)
    dst = cache / f"{key}.xlsx"
    if dst.exists():
        return dst
    merges = re.compile(r"<mergeCells[^>]*>.*?</mergeCells>|<mergeCells[^>]*/>", re.S)
    tmp = dst.with_suffix(".tmp")
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("xl/worksheets/sheet") and b"<mergeCells" in data:
                data = merges.sub("", data.decode("utf-8")).encode("utf-8")
            zout.writestr(item, data)
    tmp.replace(dst)
    return dst


class Tracker:
    def __init__(self, path, include_hidden=False):
        self.path = Path(path)
        try:
            fast = _fast_workbook_copy(path)
        except Exception:  # noqa: BLE001 - any problem: fall back to the slow but safe load
            fast = path
        self.wb = openpyxl.load_workbook(str(fast), data_only=True)
        self.include_hidden = include_hidden

    def sheets(self):
        return [(ws.title, ws.sheet_state) for ws in self.wb.worksheets]

    def resolve_sheet(self, wanted):
        cands = [ws.title for ws in self.wb.worksheets if self.include_hidden or ws.sheet_state == "visible"]
        norm = lambda s: re.sub(r"[^a-z0-9]", "", s.lower())
        w = norm(wanted)
        exact = [c for c in cands if norm(c) == w]
        if len(exact) == 1:
            return exact[0], []
        part = [c for c in cands if w in norm(c) or norm(c) in w]
        if len(part) == 1:
            return part[0], []
        if len(part) > 1:
            return None, part
        return None, []


def extract(tracker, sheet_name, lower_better=(), lenient=False):
    """lenient=True also captures KPI blocks whose section header is not named
    'Critical ...' / 'Enabling ...' (older tracker layouts). Used for the comparison
    (previous month) workbook only, so month-on-month deltas can still be computed."""
    ws = tracker.wb[sheet_name]
    issues = []          # (severity, text)  severity: blocking | ambiguous | missing | note
    maxr = ws.max_row
    A = lambda r: clean(ws.cell(r, 1).value)

    # ---- locate markers
    year_blocks = []       # (year, header_row)
    kpi_headers = []       # (row_of_Metric_header, kind)
    kpi_markers = []       # rows of the 'Critical ...' / 'Enabling ...' section headings
    dep_marker = None
    r = 1
    while r <= maxr:
        a = A(r)
        y = is_year(a)
        if y and A(r + 1) == "Key Themes":
            year_blocks.append((y, r + 1))
        elif isinstance(a, str) and re.search(r"(?i)interdependenc", a) and re.search(r"(?i)across", a):
            dep_marker = r
        elif a == "Metric" or (lenient and isinstance(a, str) and clean(ws.cell(r, 8).value) == "Quaterly"):
            kind, marker = None, r
            for back in range(1, 4):
                p = A(r - back)
                if isinstance(p, str) and re.search(r"(?i)enabling", p):
                    kind, marker = "enabling", r - back; break
                if isinstance(p, str) and re.search(r"(?i)critical", p):
                    kind, marker = "critical", r - back; break
            if kind is None and lenient and any(clean(ws.cell(rr, 8).value) == "T/A" for rr in range(r + 1, min(r + 40, maxr + 1))):
                kind = "critical" if not any(k == "critical" for _, k in kpi_headers) else "enabling"
                marker = r
            if kind:
                kpi_headers.append((r, kind))
                kpi_markers.append(marker)
        r += 1
    if not year_blocks:
        issues.append(("blocking", "No year blocks ('2026' followed by a 'Key Themes' header) found in this sheet; it does not follow the tracker layout."))
    years = [y for y, _ in year_blocks]

    # ---- theme / action rows
    stop_rows = sorted([h - 1 for _, h in year_blocks] + kpi_markers + [h for h, _ in kpi_headers] + ([dep_marker] if dep_marker else []) + [maxr + 1])
    themes = []            # list of dict(year, theme, rows=[...])
    for (year, hdr) in year_blocks:
        end = min([s for s in stop_rows if s > hdr + 1] + [maxr + 1])
        cur = None
        for rr in range(hdr + 2, end):
            a = A(rr)
            if a is not None and not is_year(a):
                cur = {"year": year, "theme": str(a), "theme_cell": a1(ws, rr, 1), "rows": []}
                themes.append(cur)
            action = clean(ws.cell(rr, 3).value)
            if action is None:
                continue
            if cur is None:
                cur = {"year": year, "theme": None, "theme_cell": None, "rows": []}
                themes.append(cur)
                issues.append(("ambiguous", f"{sheet_name}!C{rr}: action without a Key Theme in column A"))
            row = {
                "row": rr,
                "owner": clean(ws.cell(rr, 2).value),
                "action": str(action),
                "start": clean(ws.cell(rr, 6).value),
                "planned": clean(ws.cell(rr, 7).value),
                "progress": ws.cell(rr, 8).value if isinstance(ws.cell(rr, 8).value, (int, float)) and not isinstance(ws.cell(rr, 8).value, bool) else None,
                "status": canon_status(ws.cell(rr, 9).value, issues, f"{sheet_name}!I{rr}"),
                "theme_progress": ws.cell(rr, 10).value if isinstance(ws.cell(rr, 10).value, (int, float)) else None,
                "units": clean(ws.cell(rr, 11).value),
                "impact": clean(ws.cell(rr, 12).value),
            }
            if ws.cell(rr, 8).value is not None and row["progress"] is None:
                issues.append(("ambiguous", f"{sheet_name}!H{rr}: progress {ws.cell(rr, 8).value!r} is not numeric; ignored in averages"))
            cur["rows"].append(row)
    for t in themes:
        if not t["rows"]:
            issues.append(("missing", f"{sheet_name}!{t['theme_cell']}: key theme '{t['theme']}' ({t['year']}) has no action rows"))

    # ---- KPI blocks
    kpis = {"critical": [], "enabling": []}
    kpi_block_years = None
    for idx, (mh, kind) in enumerate(kpi_headers):
        yrs = [is_year(ws.cell(mh + 1, c).value) for c in (4, 5, 6)]
        if all(yrs):
            kpi_block_years = yrs
        else:
            issues.append(("ambiguous", f"{sheet_name}!D{mh+1}:F{mh+1}: target-year labels not found; using {years or 'unknown'}"))
            yrs = (years + [None, None, None])[:3]
        nxt = min([h for h, _ in kpi_headers if h > mh] + ([dep_marker] if dep_marker and dep_marker > mh else []) + [maxr + 1])
        rr = mh + 2
        while rr < nxt:
            if clean(ws.cell(rr, 8).value) == "T/A":
                name = clean(ws.cell(rr, 1).value)
                trow, arow = rr + 1, None
                if clean(ws.cell(trow, 8).value) != "Target":
                    issues.append(("ambiguous", f"{sheet_name}!H{trow}: expected 'Target' label under T/A anchor; row skipped"))
                    rr += 1; continue
                for k in range(trow + 1, trow + 4):
                    if clean(ws.cell(k, 8).value) == "Actual":
                        arow = k; break
                if arow is None:
                    issues.append(("ambiguous", f"{sheet_name}!H{trow+1}:H{trow+3}: 'Actual' label not found under T/A anchor at row {rr}; row skipped"))
                    rr += 1; continue
                if name is None:
                    # empty KPI slot (template rows with no metric) - ignore but note once per block
                    issues.append(("note", f"{sheet_name}!A{rr}: empty KPI slot (no metric name) ignored"))
                    rr = arow + 1; continue
                qlabels = [clean(ws.cell(rr, c).value) for c in (9, 10, 11, 12)]
                if qlabels != ["Q1", "Q2", "Q3", "Q4"]:
                    issues.append(("ambiguous", f"{sheet_name}!I{rr}:L{rr}: quarter labels are {qlabels}; expected Q1..Q4"))
                kpi = {
                    "kind": kind, "row": rr, "metric": str(name), "owner": clean(ws.cell(rr, 2).value),
                    "baseline": clean(ws.cell(rr, 3).value), "baseline_fmt": ws.cell(rr, 3).number_format,
                    "targets": OrderedDict(), "committed": clean(ws.cell(rr, 7).value),
                    "quarters": OrderedDict(), "cells": {"baseline": a1(ws, rr, 3), "committed": a1(ws, rr, 7), "owner": a1(ws, rr, 2)},
                }
                fmts = [ws.cell(rr, 3).number_format]
                for y, c in zip(yrs, (4, 5, 6)):
                    kpi["targets"][str(y)] = clean(ws.cell(rr, c).value)
                    kpi["cells"][f"target_{y}"] = a1(ws, rr, c)
                    fmts.append(ws.cell(rr, c).number_format)
                for qi, c in enumerate((9, 10, 11, 12), 1):
                    q = f"Q{qi}"
                    kpi["quarters"][q] = {"target": clean(ws.cell(trow, c).value), "actual": clean(ws.cell(arow, c).value)}
                    kpi["cells"][f"{q}_target"] = a1(ws, trow, c)
                    kpi["cells"][f"{q}_actual"] = a1(ws, arow, c)
                    fmts.append(ws.cell(trow, c).number_format); fmts.append(ws.cell(arow, c).number_format)
                kpi["kind_of_value"] = infer_value_kind(kpi, fmts, issues, sheet_name)
                kpi["lower_is_better"] = any(name.lower() == lb.lower() for lb in lower_better)
                kpis[kind].append(kpi)
                rr = arow + 1
            else:
                rr += 1
    if not kpi_headers:
        issues.append(("missing", "No 'Critical' / 'Enabling' KPI tables (rows with a 'Metric' header) found in this sheet"))

    # ---- interdependencies (asks placed on this function by other functions)
    deps = []
    if dep_marker:
        cur_year = None
        rr = dep_marker + 1
        empty_run = 0
        while rr <= maxr and empty_run < 40:
            a = A(rr)
            y = is_year(a)
            if y and A(rr + 1) == "Functions":
                cur_year = y; rr += 2; empty_run = 0; continue
            if a is None:
                empty_run += 1; rr += 1; continue
            empty_run = 0
            if a in ("Functions",):
                rr += 1; continue
            if cur_year is None:
                issues.append(("ambiguous", f"{sheet_name}!A{rr}: dependency row before any year header; skipped"))
                rr += 1; continue
            prog = ws.cell(rr, 5).value
            deps.append({
                "year": cur_year, "row": rr, "function": str(a),
                "theme": clean(ws.cell(rr, 2).value), "action": clean(ws.cell(rr, 3).value),
                "planned": clean(ws.cell(rr, 4).value),
                "progress": prog if isinstance(prog, (int, float)) and not isinstance(prog, bool) else None,
                "status": canon_status(ws.cell(rr, 6).value, issues, f"{sheet_name}!F{rr}"),
            })
            if deps[-1]["action"] is None:
                issues.append(("missing", f"{sheet_name}!C{rr}: dependency row for '{a}' has no action text"))
            rr += 1
    else:
        issues.append(("missing", "No 'Interdependencies across all function' section found in this sheet"))

    # ---- VCP outcome table ('Metric | Year | Q1..Q4 | FY [| %]'), e.g. Sales / iGM / Adj. EBITA by year
    outcomes, outcomes_rows = [], None
    for rr in range(1, maxr + 1):
        if A(rr) == "Metric" and clean(ws.cell(rr, 2).value) == "Year":
            hdr = [clean(ws.cell(rr, c).value) for c in range(1, 12)]
            fy_col = next((i + 1 for i, h in enumerate(hdr) if isinstance(h, str) and h.upper() == "FY"), None)
            if not fy_col:
                continue
            cur, r2 = None, rr + 1
            while r2 <= maxr:
                name, yr = A(r2), is_year(ws.cell(r2, 2).value)
                if name is None and yr is None:
                    break
                if name is not None:
                    cur = {"metric": str(name), "row": r2, "years": OrderedDict(), "pct": OrderedDict()}
                    outcomes.append(cur)
                if cur is not None and yr:
                    cur["years"][str(yr)] = clean(ws.cell(r2, fy_col).value)
                    extra = ws.cell(r2, fy_col + 1).value
                    if isinstance(extra, (int, float)) and not isinstance(extra, bool):
                        cur["pct"][str(yr)] = extra
                r2 += 1
            outcomes_rows = (rr, r2 - 1)
            break
    return {
        "outcomes": outcomes, "outcomes_rows": outcomes_rows,
        "source_file": str(tracker.path), "sheet": sheet_name,
        "function": str(A(1) or sheet_name),
        "years": years, "kpi_years": kpi_block_years,
        "themes": themes, "kpis": kpis, "dependencies": deps, "issues": issues,
    }


def infer_value_kind(kpi, fmts, issues, sheet):
    """percent | currency | number, inferred from Excel number formats and sibling text cells of the same KPI row."""
    vals = [kpi["baseline"]] + list(kpi["targets"].values()) + [q[k] for q in kpi["quarters"].values() for k in ("target", "actual")]
    pct_fmt = any(isinstance(f, str) and "%" in f for f in fmts)
    pct_txt = any(isinstance(v, str) and "%" in v for v in vals) or "(%)" in kpi["metric"] or kpi["metric"].strip().endswith("%")
    cur_txt = any(isinstance(v, str) and CURRENCY_RE.search(v) for v in vals) or bool(re.search(r"\(M\s*€\)|\(M€\)|MEur|€M|M €", kpi["metric"]))
    numeric = [v for v in vals if isinstance(v, (int, float)) and not isinstance(v, bool)]
    if (pct_fmt or pct_txt) and cur_txt:
        issues.append(("ambiguous", f"{sheet}!A{kpi['row']} '{kpi['metric']}': row mixes % and € values; numeric cells shown unformatted"))
        return "number"
    if pct_fmt or pct_txt:
        return "percent"
    if cur_txt:
        return "currency"
    if numeric:
        issues.append(("note", f"{sheet}!A{kpi['row']} '{kpi['metric']}': unit not inferable (no % format, no € text); numbers shown as stored"))
    return "number"


# --------------------------------------------------------------------------- derivations
def fmt_value(v, kind, number_format=None, issues=None, where=""):
    if v is None:
        return NOT_AVAILABLE
    if isinstance(v, (dt.datetime, dt.date)):
        return fmt_date(v)
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, (int, float)):
        if kind == "percent":
            if number_format and "%" in str(number_format):
                return f"{v*100:.1f}%"
            if abs(v) <= 1:
                return f"{v*100:.1f}%"
            if issues is not None:
                issues.append(("note", f"{where}: value {v!r} in a %-typed row with General format shown as percentage points"))
            return f"{trim_num(v)}%"
        if kind == "currency":
            return f"€{trim_num(v)}M"
        return trim_num(v, 4)
    return str(v)


def kpi_latest_quarter(kpi):
    latest = None
    for q, d in kpi["quarters"].items():
        if d["target"] is not None and d["actual"] is not None:
            latest = q
    return latest


def kpi_has_any_quarter_data(kpi):
    return any(d["target"] is not None or d["actual"] is not None for d in kpi["quarters"].values())


def signal_for(kpi, q, issues, sheet):
    d = kpi["quarters"].get(q, {"target": None, "actual": None})
    t, a = parse_number(d["target"]), parse_number(d["actual"])
    if d["target"] is None or d["actual"] is None:
        return "No data", None
    if t is None or a is None:
        issues.append(("ambiguous", f"{sheet}: '{kpi['metric']}' {q} target/actual ({d['target']!r} / {d['actual']!r}) not numerically comparable; signal = No data"))
        return "No data", None
    if kpi.get("direction_unconfirmed"):
        return "Direction to confirm", (a - t)
    if kpi["kind_of_value"] == "percent":
        # normalise fractions vs points for the delta only
        pass
    good = (a <= t) if kpi["lower_is_better"] else (a >= t)
    return ("On / above target" if good else "Below target"), (a - t)


def delta_text(kpi, q):
    d = kpi["quarters"][q]
    t, a = parse_number(d["target"]), parse_number(d["actual"])
    kind = kpi["kind_of_value"]
    if kind == "percent":
        scale = 100 if (abs(t) <= 1 and abs(a) <= 1) else 1
        return f"({(a-t)*scale:+.1f} pts)"
    if kind == "currency":
        return f"({a-t:+.2f}M)"
    return f"({trim_num(a-t, 4) if a-t < 0 else '+'+trim_num(a-t, 4)})"


def derive_status(statuses):
    s = [x for x in statuses if x]
    if not s:
        return "Not started"
    if all(x == "Completed" for x in s):
        return "Completed"
    if any(x == "Delayed" for x in s):
        return "Delayed"
    if all(x == "Not started" for x in s):
        return "Not started"
    if all(x == "On hold" for x in s):
        return "On hold"
    return "In progress"


def pct(values):
    v = [x for x in values if x is not None]
    return round(100 * sum(v) / len(v)) if v else 0


def build_view(model, quarter=None, truncate_dep_actions=None):
    """Turn the extracted model into the exact strings for each slide. Returns (view, questions)."""
    issues = model["issues"]
    questions = []
    sheet = model["sheet"]
    crit, enab = model["kpis"]["critical"], model["kpis"]["enabling"]
    all_kpis = crit + enab
    years = model["kpi_years"] or model["years"] or []
    last_year = years[-1] if years else None
    fn = model["function"]

    # reporting quarter
    if quarter is None:
        latest = {kpi_latest_quarter(k) for k in crit} - {None}
        if not latest:
            latest = {kpi_latest_quarter(k) for k in enab} - {None}
        if len(latest) == 1:
            quarter = latest.pop()
        elif len(latest) > 1:
            questions.append(f"Reporting quarter is ambiguous: KPIs have different latest quarters with both target and actual ({sorted(latest)}). Choose the reporting quarter (CLI: --quarter Qn).")
        else:
            questions.append("No KPI has both a quarterly target and actual, so the reporting quarter cannot be detected. Choose the reporting quarter (CLI: --quarter Qn); the KPI slides will show 'No data' for it.")
    q = quarter or "Q?"

    # lower-is-better candidates that actually get compared
    hints = []
    for k in all_kpis:
        if not k["lower_is_better"] and kpi_latest_quarter(k) and any(h in k["metric"].lower() for h in LOWER_BETTER_HINTS):
            hints.append(k["metric"])
            k["direction_unconfirmed"] = True          # shown as 'Direction to confirm', never guessed green/red
    if hints:
        issues.append(("ambiguous", "KPI direction not declared for: " + "; ".join(hints)
                       + ". Their names suggest a lower actual is better, so the quarter signal is shown as 'Direction to confirm' "
                         "instead of red/green. Declare them with --lower-better (or confirm they are higher-is-better) to get a signal."))

    # ---- slide 2 critical table
    def kpi_row(k, cols):
        sig, _ = signal_for(k, q, issues, sheet)
        vals = {
            "metric": k["metric"],
            "baseline": fmt_value(k["baseline"], k["kind_of_value"], k["baseline_fmt"], issues, f"{sheet}!{k['cells']['baseline']}"),
            "q_target": fmt_value(k["quarters"].get(q, {}).get("target"), k["kind_of_value"], None, issues, f"{sheet}!{k['cells'].get(q+'_target','')}"),
            "q_actual": fmt_value(k["quarters"].get(q, {}).get("actual"), k["kind_of_value"], None, issues, f"{sheet}!{k['cells'].get(q+'_actual','')}"),
            "signal": sig,
        }
        for y in years:
            vals[f"target_{y}"] = fmt_value(k["targets"].get(str(y)), k["kind_of_value"], None, issues, f"{sheet}!{k['cells'].get('target_'+str(y),'')}")
        if k["committed"]:
            vals["note"] = k["committed"]
        else:
            missing = []
            if k["baseline"] is None: missing.append("baseline")
            if any(v is None for v in k["targets"].values()): missing.append("targets")
            missing.append("committed action")
            if not kpi_has_any_quarter_data(k): missing.append("quarterly data")
            vals["note"] = ", ".join(missing[:-1]) + (" and " if len(missing) > 1 else "") + missing[-1] + " need completion"
            vals["note"] = vals["note"][0].upper() + vals["note"][1:]
        return [vals[c] for c in cols], sig

    crit_cols = ["metric", "baseline"] + [f"target_{y}" for y in years] + ["q_target", "q_actual", "signal"]
    crit_rows, crit_sigs = [], []
    for k in crit:
        r, s = kpi_row(k, crit_cols); crit_rows.append(r); crit_sigs.append(s)
    enab_cols = ["metric", "baseline", f"target_{last_year}", "q_target", "q_actual", "signal", "note"]
    enab_rows, enab_sigs = [], []
    for k in enab:
        r, s = kpi_row(k, enab_cols); enab_rows.append(r); enab_sigs.append(s)

    # ---- slide 4 themes
    theme_rows = []
    for t in model["themes"]:
        rows = t["rows"]
        tp = next((r["theme_progress"] for r in rows if r["theme_progress"] is not None), None)
        p = round(100 * tp) if tp is not None else pct([r["progress"] for r in rows])
        st = derive_status([r["status"] for r in rows])
        actions = "; ".join(r["action"] for r in rows) if rows else NOT_AVAILABLE
        theme_rows.append({"cells": [str(t["year"]), t["theme"] or NOT_AVAILABLE, actions, f"{p}% | {st}"], "status": st,
                           "derived": "theme_progress column" if tp is not None else f"average of {len([r for r in rows if r['progress'] is not None])} action progress values"})

    # ---- slide 5 dependencies grouped by function (alphabetical, as in the template)
    deps = model["dependencies"]
    groups = OrderedDict()
    for d in sorted(deps, key=lambda d: (d["function"].lower(), d["year"], d["row"])):
        groups.setdefault(d["function"], []).append(d)
    dep_rows = []
    for fname, ds in groups.items():
        themes_u = []
        for d in ds:
            if d["theme"] and d["theme"] not in themes_u:
                themes_u.append(d["theme"])
        acts = [d["action"] or NOT_AVAILABLE for d in ds]
        text = f"{len(ds)}: " + "; ".join(acts)
        if truncate_dep_actions and len(text) > truncate_dep_actions:
            text = text[:truncate_dep_actions].rstrip() + "..."
        st = derive_status([d["status"] for d in ds])
        dep_rows.append({"cells": [fname, "; ".join(themes_u) or NOT_AVAILABLE, text, f"{pct([d['progress'] for d in ds])}% | {st}"], "status": st})
    n_dep = len(deps)
    n_dep_ns = sum(1 for d in deps if d["status"] == "Not started")
    n_dep_last = sum(1 for d in deps if d["year"] == last_year)

    # ---- slide 6 impact (source: column L 'VCP Commitment KPI Impact')
    impact_rows = []
    any_impact = False
    for t in model["themes"]:
        imp = []
        for r in t["rows"]:
            if r["impact"] and r["impact"] not in imp:
                imp.append(r["impact"])
        if imp:
            any_impact = True
        impact_rows.append([t["theme"] or NOT_AVAILABLE, "; ".join(imp) if imp else NOT_AVAILABLE, ""])

    # ---- slide 7 RCA + data readiness
    red = []
    for k, s in zip(crit + enab, crit_sigs + enab_sigs):
        if s == "Below target":
            d = k["quarters"][q]
            red.append([k["metric"], f"{fmt_value(d['actual'], k['kind_of_value'])} vs {fmt_value(d['target'], k['kind_of_value'])} {delta_text(k, q)}", "", ""])
    n_k = len(all_kpis)
    owner_blank = sum(1 for k in all_kpis if not k["owner"])
    tgt_missing = {y: sum(1 for k in all_kpis if k["targets"].get(str(y)) is None) for y in years}
    q_cells_total = n_k * 8
    q_cells_missing = sum(1 for k in all_kpis for d in k["quarters"].values() for key in ("target", "actual") if d[key] is None)
    no_q = sum(1 for k in all_kpis if not kpi_has_any_quarter_data(k))
    action_rows = [r for t in model["themes"] for r in t["rows"]]
    n_a = len(action_rows)
    blanks = {
        "Impact": sum(1 for r in action_rows if not r["impact"]),
        "Owner": sum(1 for r in action_rows if not r["owner"]),
        "Start Date": sum(1 for r in action_rows if not r["start"]),
        "Theme Progress": sum(1 for r in action_rows if r["theme_progress"] is None),
    }
    readiness = [
        ["Owner fields", f"{owner_blank}/{n_k} blank across commitment/KPI rows"],
        ["Target fields", "; ".join(f"{y} target missing {m}/{n_k}" for y, m in tgt_missing.items()) or NOT_AVAILABLE],
        ["Quarterly data", f"{q_cells_missing}/{q_cells_total} target/actual cells missing; {no_q} KPIs have no quarterly data"],
        ["Theme tracker blanks", " | ".join(f"{k}: {v}/{n_a}" for k, v in blanks.items())],
    ]

    def rng(rows):
        return f"rows {min(rows)}-{max(rows)}" if rows else "no rows"
    src = f"{Path(model['source_file']).name}, sheet '{sheet}'"
    action_rows_n = [r["row"] for t in model["themes"] for r in t["rows"]]
    sources = {
        "cover": src,
        "critical": f"{src}, critical KPI table {rng([k['row'] for k in crit])} (T/A anchor rows)",
        "enabling": f"{src}, enabling KPI table {rng([k['row'] for k in enab])} (T/A anchor rows)",
        "themes": f"{src}, year-block theme/action rows {rng(action_rows_n)}",
        "dependencies": f"{src}, 'Interdependencies across all function' section {rng([d['row'] for d in deps])}",
        "impact": f"{src}, column L 'VCP Commitment KPI Impact' of the theme/action rows {rng(action_rows_n)}",
        "rca": f"{src}, KPI rows with a '{q} signal' of Below target; data-readiness counts computed from the KPI and theme/action tables",
        "delta": f"{src} compared with the previous tracker (theme/action rows, KPI tables, dependencies); matching by name, see validation summary",
    }
    def pop_pct(ks):
        tot = len(ks) * 8
        filled = sum(1 for k in ks for d in k["quarters"].values() for key in ("target", "actual") if d[key] is not None)
        return round(100 * filled / tot) if tot else 0
    readiness_raw = {
        "n_kpis": n_k, "owner_blank": owner_blank, "targets_missing": {str(y): m for y, m in tgt_missing.items()},
        "q_missing": q_cells_missing, "q_total": q_cells_total, "no_q": no_q,
        "crit_pop_pct": pop_pct(crit), "enab_pop_pct": pop_pct(enab),
        "baseline_missing": sum(1 for k in all_kpis if k["baseline"] is None),
        "committed_missing": sum(1 for k in all_kpis if not k["committed"]),
        "theme_cols_empty": [name for name, cnt in blanks.items() if n_a and cnt == n_a] + (["Units"] if n_a and all(not r["units"] for r in action_rows) else []),
        "planned_blank": sum(1 for r in action_rows if not r["planned"]),
        "progress_blank": sum(1 for r in action_rows if r["progress"] is None),
        "n_actions": n_a, "ready_total": sum(1 for k in all_kpis if kpi_latest_quarter(k) == q),
    }
    dep_register = [{"year": d["year"], "function": d["function"], "action": d["action"], "status": d["status"]}
                    for d in sorted(deps, key=lambda d: (d["year"], d["row"]))]
    orows = model.get("outcomes_rows")
    sources["outcomes"] = f"{src}, outcome table (Metric/Year/FY) " + (f"rows {orows[0]}-{orows[1]}" if orows else "not present in the sheet")
    sources["impact"] = sources["impact"] + "; VCP outcome table " + (f"rows {orows[0]}-{orows[1]}" if orows else "not present")
    view = {
        "sources": sources,
        "outcomes": model.get("outcomes", []), "dep_register": dep_register, "readiness_raw": readiness_raw,
        "function": fn, "quarter": q, "years": years, "last_year": last_year,
        "title": f"{fn} VCP Program Execution {years[0]}-{years[-1]}" if years else f"{fn} VCP Program Execution",
        "kicker": f"{fn.upper()} VCP",
        "critical": {"cols": crit_cols, "rows": crit_rows, "signals": crit_sigs,
                     "cards": [str(len(crit)), str(sum(1 for s in crit_sigs if s == "Below target")), str(sum(1 for s in crit_sigs if s == "On / above target"))]},
        "enabling": {"cols": enab_cols, "rows": enab_rows, "signals": enab_sigs,
                     "cards": [str(sum(1 for k in enab if kpi_latest_quarter(k) == q)), str(sum(1 for s in enab_sigs if s == "Below target")), str(sum(1 for k in enab if not kpi_has_any_quarter_data(k)))]},
        "themes": theme_rows,
        "dependencies": {"rows": dep_rows, "cards": [str(n_dep), str(n_dep_ns), str(n_dep_last)]},
        "impact": {"rows": impact_rows, "has_source_data": any_impact},
        "rca": {"rows": red, "readiness": readiness},
        "counts": {"themes": len(model["themes"]), "actions": n_a, "critical_kpis": len(crit), "enabling_kpis": len(enab), "dependencies": n_dep},
    }
    return view, questions


def _table_geometry(graphic_frame):
    tbl = graphic_frame.table._tbl
    trs = tbl.findall(qn("a:tr"))
    widths = [int(g.get("w")) / 914400 for g in tbl.find(qn("a:tblGrid")).findall(qn("a:gridCol"))]
    head_h = int(trs[0].get("h", 0)) / 914400
    proto = trs[1] if len(trs) > 1 else trs[0]
    m = re.search(r'sz="(\d+)"', etree.tostring(proto, encoding="unicode"))
    pt = int(m.group(1)) / 100 if m else 12
    return {"widths": widths, "head_h": head_h, "pt": pt, "min_h": int(proto.get("h", 0)) / 914400, "top": graphic_frame.top / 914400}


_FONTS = {}


def _font(pt):
    """Calibri metrics (what PowerPoint substitutes for the template's Helvetica Neue on Windows); None if unavailable."""
    if pt not in _FONTS:
        f = None
        try:
            from PIL import ImageFont
            for cand in (r"C:\Windows\Fonts\calibri.ttf", "calibri.ttf", "Calibri.ttf"):
                try:
                    f = ImageFont.truetype(cand, size=int(round(pt * 10)))   # 10x for precision
                    break
                except OSError:
                    continue
        except ImportError:
            f = None
        _FONTS[pt] = f
    return _FONTS[pt]


def _text_lines(text, width_in, pt):
    """Number of lines a cell needs: word-wrap simulation with real glyph widths (fallback: character count)."""
    font = _font(pt)
    usable_pt = max(width_in - 0.2, 0.3) * 72
    total = 0
    for para in str(text or "").split("\n"):
        words = para.split(" ")
        if font is None:
            cpi = 12.0 * (12.0 / pt)
            total += max(1, math.ceil(len(para) / max(int(max(width_in - 0.2, 0.3) * cpi), 1)))
            continue
        lines, cur = 1, 0.0
        space = font.getlength(" ") / 10
        for w in words:
            wl = font.getlength(w) / 10
            if cur == 0:
                cur = wl
            elif cur + space + wl <= usable_pt:
                cur += space + wl
            else:
                lines += 1
                cur = wl
            while cur > usable_pt:              # a single word longer than the cell
                lines += 1
                cur -= usable_pt
        total += lines
    return max(total, 1)


def est_row_height(cells, geo, min_h=None):
    """Estimate (inches) of a rendered table row: the tallest wrapped cell at the body font size."""
    pt = geo["pt"]
    lines_max = max(_text_lines(txt, w, pt) for txt, w in zip(cells, geo["widths"]))
    return max(min_h if min_h is not None else geo["min_h"], lines_max * pt * 1.2 / 72 + 0.1)


MIN_BODY_PT = 10


def pack_rows(rows_cells, geo, min_h=None, slide_h=7.5, bottom_margin=0.15, pt=None):
    """Group row indices so each group fits between the table top and the slide bottom."""
    if pt:
        geo = dict(geo, pt=pt)
    avail = slide_h - bottom_margin - geo["top"] - geo["head_h"]
    groups, cur, used = [], [], 0.0
    for i, cells in enumerate(rows_cells):
        h = est_row_height(cells, geo, min_h)
        if cur and used + h > avail:
            groups.append(cur)
            cur, used = [], 0.0
        cur.append(i)
        used += h
    groups.append(cur)
    return groups


TABLE_SHAPES = {"critical": (1, "Table 28"), "enabling": (2, "Table 28"), "themes": (3, "Table 10"),
                "dependencies": (4, "Table 28"), "impact": (5, "Table 10"), "rca": (6, "Table 13")}
TITLES = {"critical": "Critical VCP Commitments", "enabling": "VCP Commitment Enabling KPIs",
          "dependencies": "Cross-functional dependencies", "impact": "Impacting KPIs towards VCP", "rca": "RCA for red KPIs"}
PLACEHOLDER = {
    "critical": "No critical KPI rows recorded in the sheet",
    "enabling": "No enabling KPI rows recorded in the sheet",
    "themes": "No key themes / actions recorded in the sheet",
    "dependencies": "No interdependency rows recorded in the sheet",
}


def rows_by_key(view):
    return {"critical": view["critical"]["rows"], "enabling": view["enabling"]["rows"],
            "themes": [t["cells"] for t in view["themes"]], "dependencies": [d["cells"] for d in view["dependencies"]["rows"]],
            "impact": view["impact"]["rows"], "rca": view["rca"]["rows"]}


def plan_tables(view, template, impact_mode=None):
    """Decide how each table is laid out: first the template font size, then 1-2pt smaller
    (never below MIN_BODY_PT) to keep it on one slide, and only then split across slides."""
    prs = Presentation(str(template))
    slides = list(prs.slides)
    plan, over = {}, []
    for key, (si, name) in TABLE_SHAPES.items():
        rows = rows_by_key(view)[key]
        if key == "impact" and impact_mode == "drop" and not view["impact"]["has_source_data"]:
            plan[key] = {"groups": [list(range(len(rows)))], "pt": None, "orig_pt": None}
            continue
        geo = _table_geometry(shape_by_name(slides[si], name))
        min_h = COMPACT_ROW_H / 914400 if key == "impact" and len(rows) > 5 else None
        chosen, groups = geo["pt"], [[]]
        if rows:
            for pt in (geo["pt"], geo["pt"] - 1, geo["pt"] - 2):
                if pt < MIN_BODY_PT:
                    break
                groups, chosen = pack_rows(rows, geo, min_h, pt=pt), pt
                if len(groups) == 1:
                    break
            if len(groups) > 1:            # shrinking did not help: split at the template size
                groups, chosen = pack_rows(rows, geo, min_h), geo["pt"]
        plan[key] = {"groups": groups, "pt": chosen, "orig_pt": geo["pt"]}
        if len(groups) > 1:
            over.append(f"{key}: {len(rows)} rows need {len(groups)} slides")
    return plan, over


def pre_checks(view, template, allow_extra=False, impact_mode=None):
    """Questions that must be answered by the user before a deck is written, plus the slide plan."""
    qs = []
    if _is_v2_template(template):
        import vcp_render2
        plan, over, qs2 = vcp_render2.plan_v2(view, template, impact_mode)
        if over and not allow_extra:
            qs.append("Excel content exceeds the template's slide capacity (" + "; ".join(over) + "). Choose 'add continuation slides' (CLI: --allow-extra-slides) or reduce the scope in the tracker.")
        return qs + qs2, plan
    plan, over = plan_tables(view, template, impact_mode)
    if over and not allow_extra:
        qs.append("Excel content exceeds the template's slide capacity (" + "; ".join(over) + "). Choose 'add continuation slides' (CLI: --allow-extra-slides) or reduce the scope in the tracker.")
    if not view["impact"]["has_source_data"] and impact_mode is None:
        qs.append("The 'VCP Commitment KPI Impact' column (L) is blank for every action, so the 'Impacting KPIs towards VCP' slide has no source data. Choose 'keep' (theme names only, KPI columns left empty) or 'drop' (remove the slide) for the Impact slide (CLI: --impact-slide keep|drop).")
    return qs, plan


# --------------------------------------------------------------------------- PPTX writing
def shape_by_name(slide, name):
    for sh in slide.shapes:
        if sh.name == name:
            return sh
    raise KeyError(f"shape '{name}' not found on slide")


def set_shape_text(shape, text):
    """Replace the text of a shape keeping the first run's formatting."""
    tf = shape.text_frame
    p0 = tf.paragraphs[0]
    for p in list(tf.paragraphs[1:]):
        p._p.getparent().remove(p._p)
    runs = p0.runs
    if runs:
        runs[0].text = text
        for r in runs[1:]:
            r._r.getparent().remove(r._r)
    else:
        p0.add_run().text = text


def _ensure_run(p):
    runs = p.findall(qn("a:r"))
    if runs:
        for extra in runs[1:]:
            p.remove(extra)
        return runs[0]
    r = etree.Element(qn("a:r"))
    rPr = etree.SubElement(r, qn("a:rPr"))
    end = p.find(qn("a:endParaRPr"))
    if end is not None:
        for k, v in end.attrib.items():
            rPr.set(k, v)
        for child in end:
            rPr.append(copy.deepcopy(child))
        p.insert(list(p).index(end), r)
    else:
        p.append(r)
    etree.SubElement(r, qn("a:t"))
    return r


def _set_solid(parent, hexval):
    for tag in ("a:solidFill", "a:noFill", "a:gradFill", "a:pattFill", "a:blipFill"):
        for el in parent.findall(qn(tag)):
            parent.remove(el)
    sf = etree.Element(qn("a:solidFill"))
    etree.SubElement(sf, qn("a:srgbClr")).set("val", hexval)
    idx = 0
    for i, ch in enumerate(list(parent)):
        if etree.QName(ch).localname in ("ln", "lnL", "lnR", "lnT", "lnB", "lnTlToBr", "lnBlToTr", "cell3D"):
            idx = i + 1
    parent.insert(idx, sf)


def _cell_fill(tc):
    tcPr = tc.find(qn("a:tcPr"))
    if tcPr is not None:
        sf = tcPr.find(qn("a:solidFill"))
        if sf is not None and sf.find(qn("a:srgbClr")) is not None:
            return sf.find(qn("a:srgbClr")).get("val")
    return "FFFFFF"


def set_tc(tc, text, font_color=None, fill=None, bold=None):
    txBody = tc.find(qn("a:txBody"))
    ps = txBody.findall(qn("a:p"))
    for extra in ps[1:]:
        txBody.remove(extra)
    p = ps[0]
    r = _ensure_run(p)
    rPr = r.find(qn("a:rPr"))
    if rPr is None:
        rPr = etree.Element(qn("a:rPr"))
        r.insert(0, rPr)
    for att in ("err", "dirty"):
        rPr.attrib.pop(att, None)
    if bold is not None:
        rPr.set("b", "1" if bold else "0")
    if font_color:
        _set_solid(rPr, font_color)
    t = r.find(qn("a:t"))
    t.text = text
    if text != text.strip():
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    if fill:
        tcPr = tc.find(qn("a:tcPr"))
        if tcPr is None:
            tcPr = etree.SubElement(tc, qn("a:tcPr"))
        _set_solid(tcPr, fill)


def fill_table(graphic_frame, rows, status_col=None, status_values=None, header=None, palette=None):
    """Rebuild the body rows of a template table from its prototype rows (keeps zebra styling)."""
    tbl = graphic_frame.table._tbl
    trs = tbl.findall(qn("a:tr"))
    head, protos = trs[0], trs[1:3] if len(trs) > 2 else trs[1:2]
    if header:
        for tc, txt in zip(head.findall(qn("a:tc")), header):
            if txt is not None:
                set_tc(tc, txt)
    for tr in trs[1:]:
        tbl.remove(tr)
    for i, row in enumerate(rows):
        tr = copy.deepcopy(protos[i % len(protos)])
        for ext in tr.findall(qn("a:extLst")):
            tr.remove(ext)
        tcs = tr.findall(qn("a:tc"))
        for ci, (tc, txt) in enumerate(zip(tcs, row)):
            if status_col is not None and ci == status_col:
                st = status_values[i] if status_values else None
                pal = (palette or PALETTE).get(st) if st else None
                if pal:
                    set_tc(tc, txt, font_color=pal["font"], fill=pal["fill"], bold=True)
                else:   # no status (placeholder row): plain cell in the row's own fill
                    set_tc(tc, txt, font_color="000000", fill=_cell_fill(tcs[0]), bold=False)
            else:
                set_tc(tc, txt)
        tbl.append(tr)
    total_h = sum(int(tr.get("h", 0)) for tr in tbl.findall(qn("a:tr")))
    if total_h:
        graphic_frame.height = Emu(total_h)


def _normalise_row_style(tr):
    """Give every cell of the row the first cell's fill and run styling.

    Template columns carry per-column styling (the 'Qn signal' cell is red); once a column means
    something else that styling must not travel with it. fill_table re-colours the status column after.
    """
    tcs = tr.findall(qn("a:tc"))
    if len(tcs) < 2:
        return
    src = tcs[0]
    src_tcPr, src_rPr = src.find(qn("a:tcPr")), src.find(".//" + qn("a:rPr"))
    for tc in tcs[1:]:
        old = tc.find(qn("a:tcPr"))
        if old is not None:
            tc.remove(old)
        if src_tcPr is not None:
            tc.append(copy.deepcopy(src_tcPr))
        for rPr in tc.findall(".//" + qn("a:rPr")) + tc.findall(".//" + qn("a:endParaRPr")):
            for tag in ("a:solidFill", "a:noFill"):
                for el in rPr.findall(qn(tag)):
                    rPr.remove(el)
            if src_rPr is not None:
                f = src_rPr.find(qn("a:solidFill"))
                if f is not None:
                    rPr.append(copy.deepcopy(f))
                for att in ("b", "i", "u", "sz"):
                    if src_rPr.get(att) is not None:
                        rPr.set(att, src_rPr.get(att))
                    else:
                        rPr.attrib.pop(att, None)


def set_table_columns(graphic_frame, widths_in):
    """Rebuild a template table to len(widths_in) columns, cloning the template's own cell XML.

    New cells go after the last <a:tc> and before the row's <a:extLst>; appending after extLst is
    invalid and renders as an unstyled (white-on-white) cell.
    """
    tbl = graphic_frame.table._tbl
    grid = tbl.find(qn("a:tblGrid"))
    proto_col = copy.deepcopy(grid.findall(qn("a:gridCol"))[-1])
    while len(grid.findall(qn("a:gridCol"))) > len(widths_in):
        grid.remove(grid.findall(qn("a:gridCol"))[-1])
    while len(grid.findall(qn("a:gridCol"))) < len(widths_in):
        grid.append(copy.deepcopy(proto_col))
    for gc, w in zip(grid.findall(qn("a:gridCol")), widths_in):
        gc.set("w", str(int(Inches(w))))
    for tr in tbl.findall(qn("a:tr")):
        proto_tc = copy.deepcopy(tr.findall(qn("a:tc"))[-1])
        while len(tr.findall(qn("a:tc"))) > len(widths_in):
            tr.remove(tr.findall(qn("a:tc"))[-1])
        while len(tr.findall(qn("a:tc"))) < len(widths_in):
            pos = list(tr).index(tr.findall(qn("a:tc"))[-1]) + 1
            tr.insert(pos, copy.deepcopy(proto_tc))
        _normalise_row_style(tr)
    graphic_frame.width = Emu(int(Inches(sum(widths_in))))


def cell_lines(tc, lines):
    """Write several paragraphs into one cell: lines = [(text, size_pt|None, colour|None, bold), ...]."""
    txBody = tc.find(qn("a:txBody"))
    ps = txBody.findall(qn("a:p"))
    proto = copy.deepcopy(ps[0])
    for extra in ps[1:]:
        txBody.remove(extra)
    for i, (text, size, colour, bold) in enumerate(lines):
        p = txBody.findall(qn("a:p"))[0] if i == 0 else copy.deepcopy(proto)
        if i:
            txBody.append(p)
        run = _ensure_run(p)
        rPr = run.find(qn("a:rPr"))
        if rPr is None:
            rPr = etree.Element(qn("a:rPr"))
            run.insert(0, rPr)
        rPr.attrib.pop("dirty", None)
        rPr.attrib.pop("err", None)
        rPr.set("b", "1" if bold else "0")
        if size:
            rPr.set("sz", str(int(round(size * 100))))
        if colour:
            _set_solid(rPr, colour)
        run.find(qn("a:t")).text = text


def colour_column(graphic_frame, col, tones):
    """Colour one body column per row: tones[i] = {'fill':..,'font':..} or None."""
    trs = graphic_frame.table._tbl.findall(qn("a:tr"))
    for tr, tone in zip(trs[1:], tones):
        if not tone:
            continue
        tcs = tr.findall(qn("a:tc"))
        if col >= len(tcs):
            continue
        tc = tcs[col]
        tcPr = tc.find(qn("a:tcPr"))
        if tcPr is None:
            tcPr = etree.SubElement(tc, qn("a:tcPr"))
        _set_solid(tcPr, tone["fill"])
        for rPr in tc.findall(".//" + qn("a:rPr")):
            _set_solid(rPr, tone["font"])
            rPr.set("b", "1")


def set_body_font(graphic_frame, pt):
    """Set the font size of every body-row run (used only to keep a table on one slide)."""
    tbl = graphic_frame.table._tbl
    for tr in tbl.findall(qn("a:tr"))[1:]:
        for el in tr.iter(qn("a:rPr"), qn("a:endParaRPr")):
            el.set("sz", str(int(round(pt * 100))))


def set_notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def compact_rows(graphic_frame, h=COMPACT_ROW_H):
    tbl = graphic_frame.table._tbl
    trs = tbl.findall(qn("a:tr"))
    for tr in trs[1:]:
        tr.set("h", str(h))
    graphic_frame.height = Emu(sum(int(tr.get("h", 0)) for tr in trs))


def duplicate_slide(prs, src, after_index):
    new = prs.slides.add_slide(src.slide_layout)
    for shp in list(new.shapes):
        shp._element.getparent().remove(shp._element)
    # carry the source slide's relationships (charts, images, notes) and remap the r:id references,
    # otherwise a copied chart or picture points at a relationship the new slide does not have
    from pptx.opc.constants import RELATIONSHIP_TYPE as _RT
    rid_map = {}
    for rId, rel in src.part.rels.items():
        if rel.reltype in (_RT.SLIDE_LAYOUT, _RT.NOTES_SLIDE):
            continue
        try:
            rid_map[rId] = (new.part.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref) if rel.is_external
                            else new.part.relate_to(rel.target_part, rel.reltype))
        except Exception:  # noqa: BLE001 - a relationship we cannot carry is dropped, not fatal
            continue
    for shp in src.shapes:
        el = copy.deepcopy(shp._element)
        for node in el.iter():
            for att in (qn("r:id"), qn("r:embed"), qn("r:link")):
                v = node.get(att)
                if v in rid_map:
                    node.set(att, rid_map[v])
        new.shapes._spTree.append(el)
    lst = prs.slides._sldIdLst
    el = list(lst)[-1]
    lst.remove(el)
    lst.insert(after_index + 1, el)
    return new


def delete_slide(prs, slide):
    lst = prs.slides._sldIdLst
    for sldId in list(lst):
        if prs.slides.part.related_part(sldId.rId) is slide.part:
            prs.part.drop_rel(sldId.rId)
            lst.remove(sldId)
            return


def placeholder_row(key, geo, extra_added):
    ncols = len(geo["widths"])
    col = 0 if key in ("critical", "enabling") else max(range(ncols), key=lambda i: geo["widths"][i])
    row = [""] * ncols
    row[col] = PLACEHOLDER[key]
    extra_added.append(f"placeholder text on empty table: '{PLACEHOLDER[key]}'")
    return row


def fill_section(prs, slide, table_name, rows, statuses, status_col, header, groups, title_text, extra_added, after_fill=None, pt=None, note=None,
                 palette=None, title_shape="Rectangle 1"):
    """Fill one template table; add continuation slides for further row groups. after_fill(slide, table) runs per slide."""
    parts = [[rows[i] for i in g] for g in groups]
    sparts = [[statuses[i] for i in g] for g in groups] if statuses else [None] * len(parts)
    tbl = shape_by_name(slide, table_name)
    fill_table(tbl, parts[0], status_col, sparts[0], header, palette=palette)
    if pt:
        set_body_font(tbl, pt)
    if note:
        set_notes(slide, note)
    if after_fill:
        after_fill(slide, tbl)
    for i, (p, sp) in enumerate(zip(parts[1:], sparts[1:]), 1):
        ns = duplicate_slide(prs, slide, prs.slides.index(slide) + i - 1)
        extra_added.append(f"{title_text} (cont.)")
        set_shape_text(shape_by_name(ns, title_shape), f"{title_text} (cont.)")
        t2 = shape_by_name(ns, table_name)
        fill_table(t2, p, status_col, sp, header, palette=palette)
        if pt:
            set_body_font(t2, pt)
        if note:
            set_notes(ns, note)
        if after_fill:
            after_fill(ns, t2)


def write_deck(view, template, out_path, allow_extra=False, impact_mode=None, questions=None):
    qs, plan = pre_checks(view, template, allow_extra, impact_mode)
    questions.extend(qs)
    if questions:
        return None
    if _is_v2_template(template):
        import vcp_render2
        stamp2 = dt.date.today().strftime("%d %b %Y")

        def note_v2(key):
            return ("[Sources]\n- " + view["sources"].get(key, "")
                    + f"\n- Generated {stamp2} by the VCP Presentation Generator; every value verbatim from the Excel, derivations listed in the validation summary.")
        return vcp_render2.write_deck_v2(view, view["_model"], template, out_path, plan, allow_extra, impact_mode, [], note_v2)
    prs = Presentation(str(template))
    slides = list(prs.slides)
    s1, s2, s3, s4, s5, s6, s7 = slides[:7]
    q, ly = view["quarter"], view["last_year"]
    extra_added = []
    rk = rows_by_key(view)

    def rows_or_placeholder(key, slide, table_name):
        rows = rk[key]
        if rows:
            return rows, plan[key]["groups"]
        geo = _table_geometry(shape_by_name(slide, table_name))
        return [placeholder_row(key, geo, extra_added)], [[0]]

    def pt_for(key):
        p = plan[key]
        if p["pt"] and p["orig_pt"] and p["pt"] != p["orig_pt"]:
            extra_added.append(f"body font reduced from {p['orig_pt']:g}pt to {p['pt']:g}pt on the {key} table to keep it on one slide")
            return p["pt"]
        return None

    stamp = dt.date.today().strftime("%d %b %Y")

    def note_for(key):
        return ("[Sources]\n- " + view["sources"].get(key, "")
                + f"\n- Generated {stamp} by the VCP Presentation Generator; every value verbatim from the Excel, derivations listed in the validation summary.")

    # ---- slide 1 cover
    set_shape_text(shape_by_name(s1, "Text Placeholder 10"), view["title"])
    set_notes(s1, note_for("cover"))

    # ---- slide 2 critical commitments
    set_shape_text(shape_by_name(s2, "Rectangle 15"), view["kicker"])
    for name, val in zip(("Rectangle 5", "Rectangle 8", "Rectangle 11"), view["critical"]["cards"]):
        set_shape_text(shape_by_name(s2, name), val)
    set_shape_text(shape_by_name(s2, "Rectangle 9"), f"{q} commitments below target")
    set_shape_text(shape_by_name(s2, "Rectangle 12"), f"{q} commitments met or above target")
    hdr2 = ["Metric", "2025 baseline"] + [f"{y} target" for y in view["years"]] + [f"{q} target", f"{q} actual", f"{q} signal"]
    ncol = len(shape_by_name(s2, "Table 28").table.columns)
    if len(hdr2) != ncol:
        raise SystemExit(f"Critical table needs {len(hdr2)} columns but template has {ncol} (years found: {view['years']}).")
    rows, groups = rows_or_placeholder("critical", s2, "Table 28")
    fill_section(prs, s2, "Table 28", rows, view["critical"]["signals"] or [None], ncol - 1, hdr2, groups, TITLES["critical"], extra_added, pt=pt_for("critical"), note=note_for("critical"))

    # ---- slide 3 enabling KPIs
    for name, val in zip(("Rectangle 5", "Rectangle 8", "Rectangle 11"), view["enabling"]["cards"]):
        set_shape_text(shape_by_name(s3, name), val)
    set_shape_text(shape_by_name(s3, "Rectangle 6"), f"Enabling KPIs with {q} target and actual")
    set_shape_text(shape_by_name(s3, "Rectangle 9"), f"{q} commitments below target")
    hdr3 = ["KPI", "2025 baseline", f"{ly} target", f"{q} target", f"{q} actual", f"{q} signal", "Committed action / data note"]
    rows, groups = rows_or_placeholder("enabling", s3, "Table 28")
    fill_section(prs, s3, "Table 28", rows, view["enabling"]["signals"] or [None], 5, hdr3, groups, TITLES["enabling"], extra_added, pt=pt_for("enabling"), note=note_for("enabling"))

    # ---- slide 4 themes & actions
    y0, y1 = (view["years"][0], str(view["years"][-1])[-2:]) if view["years"] else ("", "")
    title4 = f"Execution of Year-wise Key themes & Action Items ({y0}-{y1})"
    set_shape_text(shape_by_name(s4, "Rectangle 1"), title4)
    rows, groups = rows_or_placeholder("themes", s4, "Table 10")
    fill_section(prs, s4, "Table 10", rows, [t["status"] for t in view["themes"]] or [None], 3, None, groups, title4, extra_added, pt=pt_for("themes"), note=note_for("themes"))

    # ---- month-on-month delta slides (only when a previous tracker was given)
    if view.get("delta"):
        import vcp_delta
        vcp_delta.add_delta_slides(prs, view, s2, s4, extra_added, note_for("delta"))

    # ---- slide 5 dependencies
    for name, val in zip(("Rectangle 5", "Rectangle 8", "Rectangle 11"), view["dependencies"]["cards"]):
        set_shape_text(shape_by_name(s5, name), val)
    set_shape_text(shape_by_name(s5, "Rectangle 12"), f"{ly} dependency rows")
    rows, groups = rows_or_placeholder("dependencies", s5, "Table 28")
    fill_section(prs, s5, "Table 28", rows, [d["status"] for d in view["dependencies"]["rows"]] or [None], 3, None, groups, TITLES["dependencies"], extra_added, pt=pt_for("dependencies"), note=note_for("dependencies"))

    # ---- slide 6 impact
    if impact_mode == "drop" and not view["impact"]["has_source_data"]:
        delete_slide(prs, s6)
    else:
        lbl = shape_by_name(s6, "Rectangle 3")
        lbl._element.getparent().remove(lbl._element)          # '*DUMMY DATA' label
        rows = rk["impact"]
        after = (lambda sl, t: compact_rows(t)) if len(rows) > 5 else None
        fill_section(prs, s6, "Table 10", rows, None, None, None, plan["impact"]["groups"], TITLES["impact"], extra_added, after_fill=after, pt=pt_for("impact"), note=note_for("impact"))

    # ---- slide 7 RCA
    hdr7 = ["Red metric", f"{q} signal", None, None]
    red, groups = rk["rca"], plan["rca"]["groups"]
    statuses = ["Below target"] * len(red)
    if not red:
        red = [[f"No KPI below target for {q}", "", "", ""]]
        groups = [[0]]
        statuses = ["No data"]
        extra_added.append(f"placeholder text on empty RCA table: 'No KPI below target for {q}'")
    fill_section(prs, s7, "Table 13", red, statuses, 0, hdr7, groups, TITLES["rca"], extra_added, pt=pt_for("rca"), note=note_for("rca"),
                 after_fill=lambda sl, t: fill_table(shape_by_name(sl, "Table 14"), view["rca"]["readiness"]))

    prs.save(str(out_path))
    return extra_added


# --------------------------------------------------------------------------- reconciliation
def deck_texts(pptx_path):
    prs = Presentation(str(pptx_path))
    out = []
    for i, s in enumerate(prs.slides, 1):
        for sh in s.shapes:
            if sh.has_text_frame:
                for p in sh.text_frame.paragraphs:
                    out.append((i, "".join(r.text for r in p.runs)))
            if sh.has_table:
                for row in sh.table.rows:
                    for c in row.cells:
                        out.append((i, c.text))
    return out


def reconcile(model, view, pptx_path):
    """Check every Excel-sourced string is present verbatim in the deck; check every deck table cell is expected."""
    texts = deck_texts(pptx_path)
    deck_set = set(t for _, t in texts)
    deck_blob = "\n".join(t for _, t in texts)
    missing = []
    def need(label, s):
        if s is None:
            return
        s = str(s)
        if s not in deck_blob:
            missing.append(f"{label}: {s[:80]!r}")
    for t in model["themes"]:
        need(f"theme {t['theme_cell']}", t["theme"])
        for r in t["rows"]:
            need(f"action C{r['row']}", r["action"])
    for kind in ("critical", "enabling"):
        for k in model["kpis"][kind]:
            need(f"{kind} metric A{k['row']}", k["metric"])
            shown = kind == "enabling" or view.get("template_version") != 2   # the Monthly report charts critical KPIs instead of tabling them
            if isinstance(k["baseline"], str) and shown: need(f"baseline {k['cells']['baseline']}", k["baseline"])
            for y, v in k["targets"].items():
                if isinstance(v, str) and kind == "critical" and shown: need(f"target {y} row {k['row']}", v)
            if k["committed"] and kind == "enabling": need(f"committed {k['cells']['committed']}", k["committed"])
    for d in model["dependencies"]:
        need(f"dependency function A{d['row']}", d["function"])
        need(f"dependency action C{d['row']}", d["action"])
    # expected strings from the view
    expected = []
    for sec in ("critical", "enabling"):
        for row in view[sec]["rows"]: expected += row
    for t in view["themes"]: expected += t["cells"]
    for d in view["dependencies"]["rows"]: expected += d["cells"]
    for r in view["rca"]["rows"]: expected += [x for x in r if x]
    for r in view["rca"]["readiness"]: expected += r
    if view.get("expected_v2") is not None:
        # the Monthly report composes values into sentences (remarks, driver map): check verbatim presence within the deck text
        expected = view["expected_v2"]
        not_written = [e for e in expected if e and e not in deck_blob]
    else:
        not_written = [e for e in expected if e and e not in deck_set]
    return {"excel_strings_checked": len([1 for _ in _iter_needs(model)]), "excel_strings_missing": missing,
            "view_strings_expected": len([e for e in expected if e]), "view_strings_not_written": not_written}


def _iter_needs(model):
    for t in model["themes"]:
        yield t["theme"]
        for r in t["rows"]: yield r["action"]
    for kind in ("critical", "enabling"):
        for k in model["kpis"][kind]:
            yield k["metric"]
            if k["committed"] and kind == "enabling": yield k["committed"]
    for d in model["dependencies"]:
        yield d["function"]; yield d["action"]


# --------------------------------------------------------------------------- report
def write_summary(path, model, view, extra_added, recon, args, template):
    tr = model
    iss = tr["issues"]
    by = lambda sev: [t for s, t in iss if s == sev]
    used = [tr["sheet"]]
    L = []
    L.append(f"# Validation summary - {view['function']} VCP Management View\n")
    L.append(f"- Source Excel: `{tr['source_file']}`")
    L.append(f"- Excel sheets used: {', '.join(used)} (all other sheets untouched)")
    L.append(f"- Template: `{template}`")
    L.append(f"- Reporting quarter: {view['quarter']}  |  Target years: {', '.join(map(str, view['years']))}")
    c = view["counts"]
    L.append(f"- Records processed: {c['themes']} key themes, {c['actions']} action rows, {c['critical_kpis']} critical KPIs, {c['enabling_kpis']} enabling KPIs, {c['dependencies']} interdependency rows")
    conts = [e for e in extra_added if e.endswith("(cont.)")]
    if conts:
        L.append(f"- Continuation slides added on request: {', '.join(conts)}")
    for e in extra_added:
        if e.startswith("placeholder"):
            L.append(f"- Generated {e}")
        elif not e.endswith("(cont.)"):
            L.append(f"- Layout: {e}")
    L.append("")
    L.append("## Reconciliation against the source Excel")
    L.append(f"- Excel strings (themes, actions, metrics, committed actions, dependencies) checked for verbatim presence: {recon['excel_strings_checked']}; missing: {len(recon['excel_strings_missing'])}")
    for m in recon["excel_strings_missing"]:
        L.append(f"  - MISSING: {m}")
    L.append(f"- Deck table strings expected from the mapping: {recon['view_strings_expected']}; not written: {len(recon['view_strings_not_written'])}")
    for m in recon["view_strings_not_written"]:
        L.append(f"  - NOT WRITTEN: {m!r}")
    ok = not recon["excel_strings_missing"] and not recon["view_strings_not_written"]
    L.append(f"- Result: {'RECONCILED - no business value changed, omitted, duplicated or added' if ok else 'DISCREPANCIES FOUND - see above'}")
    L.append("")
    L.append("## Derived values (documented, not business data)")
    L.append("- Quarter signal = actual vs target of the reporting quarter (higher is better unless the KPI was passed in --lower-better); 'No data' when either value is missing or not numeric.")
    L.append("- Theme progress = 'Key Theme Progress' column if filled, otherwise the average of the theme's numeric action progress values; theme status = Completed if all actions completed, Delayed if any delayed, Not started if none started, otherwise In progress.")
    L.append("- Dependency rows are grouped per function (alphabetical); progress = average of the group's numeric progress values; status derived as above. Action count prefix 'N:' = number of rows in the group.")
    L.append("- KPI cards, dependency cards and the 'Data readiness gap' table are counts computed from the sheet.")
    L.append("- Numbers are displayed with the unit inferred from the Excel number format (% cells) or sibling text cells (€...M) of the same KPI row; percentages 1 decimal.")
    L.append("- Text normalisation only: zero-width spaces removed, runs of whitespace/newlines collapsed to one space. Spelling is kept exactly as in Excel.")
    L.append("- Empty cells are shown as 'Not available' (template convention). RCA / Management follow-up columns have no Excel source and are left empty.")
    L.append("")
    L.append("## Excel fields the template design does not display")
    L.append("- Critical KPI 'Committed Actions' (column G) - the template's critical table has no such column" + (": " + "; ".join(f"{k['metric']}: {k['committed']}" for k in tr["kpis"]["critical"] if k["committed"]) if any(k["committed"] for k in tr["kpis"]["critical"]) else "."))
    L.append("- Per-action Owner, Start Date, Planned Completion Date, Progress, Units (theme tables show the aggregated progress/status only)")
    L.append("- Dependency 'Planned Completion Date' and per-row progress (dependency table shows the aggregated progress/status per function)")
    L.append("- Interdependency columns inside the year blocks (asks this function places on other functions, columns M onwards) - the template only shows the 'Interdependencies across all function' section")
    L.append("- Quarterly values other than the reporting quarter, and KPI owner names")
    L.append("")
    L.append("## Missing or ambiguous items")
    for sev in ("ambiguous", "missing", "note"):
        items = by(sev)
        if items:
            L.append(f"### {sev.title()} ({len(items)})")
            for t in items:
                L.append(f"- {t}")
    if not iss:
        L.append("- None")
    if view.get("delta"):
        import vcp_delta
        L.append("")
        L.extend(vcp_delta.summary_lines(view["delta"]))
    if view.get("governance"):
        import vcp_governance
        L.append("")
        L.extend(vcp_governance.summary_lines(view, view["_model"], view["governance"], view["governance_plan"], view.get("decisions", [])))
    # theme derivation detail
    L.append("")
    L.append("## Theme progress derivation detail")
    for t in view["themes"]:
        L.append(f"- {t['cells'][0]} | {t['cells'][1][:60]} -> {t['cells'][3]} ({t['derived']})")
    Path(path).write_text("\n".join(L) + "\n", encoding="utf-8")
    return ok


# --------------------------------------------------------------------------- render (Windows / PowerPoint)
def render(pptx_path, outdir, width=1600, height=900):
    pptx_path = Path(pptx_path).resolve()
    outdir = Path(outdir).resolve()
    outdir.mkdir(parents=True, exist_ok=True)
    stem = pptx_path.stem
    ps = f"""
$a = New-Object -ComObject PowerPoint.Application
$pres = $a.Presentations.Open("{pptx_path}", $true, $false, $false)
$i = 1
foreach ($s in $pres.Slides) {{ $s.Export("{outdir}\\{stem}-slide-$i.png", "PNG", {width}, {height}); $i++ }}
$pres.Close(); $a.Quit(); "exported $($i-1) slides"
"""
    res = subprocess.run(["powershell", "-NoProfile", "-Command", ps], capture_output=True, text=True)
    print(res.stdout.strip() or res.stderr.strip())
    return sorted(outdir.glob(f"{stem}-slide-*.png"), key=lambda p: int(p.stem.rsplit("-", 1)[1]))


# --------------------------------------------------------------------------- one-call API (CLI + web app)
def inspection_report(model, view, questions, plan):
    L = [f"Sheet: {model['sheet']}   Function: {model['function']}   Years: {model['years']}   Quarter: {view['quarter']}"]
    c = view["counts"]
    L.append(f"Themes {c['themes']} | Actions {c['actions']} | Critical KPIs {c['critical_kpis']} | Enabling KPIs {c['enabling_kpis']} | Dependencies {c['dependencies']}")
    for key, p in plan.items():
        groups = p["groups"]
        n = sum(len(g) for g in groups)
        flag = f"  <-- needs {len(groups)} slides (template has 1)" if len(groups) > 1 else ""
        if p["pt"] and p["orig_pt"] and p["pt"] != p["orig_pt"]:
            flag += f"  (body font {p['orig_pt']:g}pt -> {p['pt']:g}pt to fit one slide)"
        L.append(f"  table {key:13s}: {n:3d} rows{flag}")
    if view.get("outcomes") is not None:
        L.append("  VCP outcome table (Metric/Year/FY): " + (f"{len(view['outcomes'])} metrics found" if view["outcomes"] else "not present -> keep/drop decision for the impact linkage slide (Monthly report template)"))
    if not view["impact"]["has_source_data"]:
        L.append("  Column L 'VCP Commitment KPI Impact' is blank for all actions (Management View template: keep/drop decision for the Impact slide)")
    L.append("\nQUESTIONS FOR THE USER:" if questions else "\nNo blocking questions.")
    for qq in questions:
        L.append(f"  - {qq}")
    for sev in ("ambiguous", "missing", "note"):
        items = [t for s, t in model["issues"] if s == sev]
        if items:
            L.append(f"\n{sev.upper()} ({len(items)}):")
            for t in items:
                L.append(f"  - {t}")
    return "\n".join(L)


def generate(xlsx, function, outdir="output", quarter=None, lower_better=(), allow_extra=False, impact_mode=None,
             title=None, template=None, truncate_dep_actions=None, include_hidden=False, dry_run=False, previous=None,
             governance=None, decisions=(), governance_function=None, with_change_table=False,
             with_governance_slides=False):
    """Parse + (optionally) build in one call. Returns a dict:
       questions (list, non-empty => nothing written), report (text), pptx/summary/model (paths), reconciled (bool)."""
    template = template or DEFAULT_TEMPLATE
    tracker = Tracker(xlsx, include_hidden=include_hidden)
    sheet, cands = tracker.resolve_sheet(function)
    if sheet is None:
        q = (f"'{function}' matches several sheets: {cands}. Which one should be used?" if cands
             else f"No visible sheet matches '{function}'. Sheets: {[n for n, s in tracker.sheets() if s == 'visible']}")
        return {"questions": [q], "report": q, "pptx": None, "summary": None, "model": None, "reconciled": False}
    model = extract(tracker, sheet, lower_better=list(lower_better))
    questions = [t for s, t in model["issues"] if s == "blocking"]
    view, qs = build_view(model, quarter=quarter, truncate_dep_actions=truncate_dep_actions)
    view["_model"] = model
    view["with_change_table"] = bool(with_change_table)
    view["with_governance_slides"] = bool(with_governance_slides)
    questions += qs
    if title:
        view["title"] = title
    delta = None
    if previous:
        import vcp_delta
        prev_tracker = Tracker(previous, include_hidden=include_hidden)
        prev_sheet, pc = prev_tracker.resolve_sheet(sheet)
        if prev_sheet is None:
            prev_sheet, pc = prev_tracker.resolve_sheet(function)
        if prev_sheet is None:
            questions.append(f"The previous tracker '{Path(previous).name}' has no sheet matching '{sheet}' (candidates: {pc or 'none'}). Which sheet should be compared?")
        else:
            prev_model = extract(prev_tracker, prev_sheet, lower_better=list(lower_better), lenient=True)
            delta = vcp_delta.compare(prev_model, model, quarter=view["quarter"])
            view["delta"] = delta
    gov = None
    if governance:
        import vcp_governance
        if not _is_v2_template(template):
            questions.append("The governance / decision slides and the interactive KPI cockpit are built for the Monthly VCP report "
                             "template; re-run without --governance or with the Monthly VCP report template.")
        else:
            gov = vcp_governance.read_governance(governance)
            view["governance"] = gov
            view["decisions"] = [str(d) for d in decisions if str(d).strip()]
            gsheet = governance_function or vcp_governance.gov_sheet_for(model["function"], gov) or "*"
            view["governance_scope"] = gsheet
            view["governance_plan"] = vcp_governance.plan_slides(view, model, gov, view["decisions"], gsheet)
            gsrc = f"{Path(governance).name} (sections '1. Action Items' / '2. Decisions Taken' of each function sheet)"
            view["sources"].update({
                "decisions": "Decisions as provided verbatim by the meeting owner; implementation mapping derived from this deck's own slides",
                "governance": gsrc,
                "linkage": view["sources"]["impact"] + "; theme/action rows and KPI committed actions of the same sheet",
                "cockpit": view["sources"]["critical"] + "; " + view["sources"]["enabling"],
                "rca_drill": view["sources"]["rca"] + "; " + view["sources"]["dependencies"],
            })
            model["issues"].extend(gov.get("issues", []))
    pre_qs, plan = pre_checks(view, template, allow_extra=allow_extra, impact_mode=impact_mode)
    all_qs = questions + pre_qs
    out = {"questions": all_qs, "sheet": sheet, "function": model["function"], "pptx": None, "summary": None, "model": None, "reconciled": False, "delta": delta}
    if dry_run or all_qs:
        out["report"] = inspection_report(model, view, all_qs, plan)
        return out
    outdir = Path(outdir); outdir.mkdir(parents=True, exist_ok=True)
    safe = re.sub(r"[^A-Za-z0-9]+", "_", model["function"]).strip("_")
    stem = f"{safe}_VCP_Monthly_Review" if _is_v2_template(template) else f"{safe}_VCP_Management_View"
    out_pptx, out_md, out_json = (outdir / f"{stem}.pptx", outdir / f"{safe}_validation_summary.md", outdir / f"{safe}_model.json")
    extra = write_deck(view, template, out_pptx, allow_extra=allow_extra, impact_mode=impact_mode, questions=questions)
    recon = reconcile(model, view, out_pptx)
    ok = write_summary(out_md, model, view, extra, recon, None, template)
    view_json = {k: v for k, v in view.items() if not k.startswith("_")}
    out_json.write_text(json.dumps({"model": model, "view": view_json, "reconciliation": recon, "delta": delta}, indent=2, default=json_default, ensure_ascii=False), encoding="utf-8")
    out.update({"pptx": out_pptx, "summary": out_md, "model": out_json, "reconciled": ok, "report": out_md.read_text(encoding="utf-8")})
    return out


# --------------------------------------------------------------------------- auto mode (Momo / one-click)
AUTO_DEFAULTS = {
    "defaultFunction": None,            # None = ask which tab (sheet) to build
    "quarter": None,                    # None = auto-detect
    "impactSlide": "drop",              # what to do when column L is blank
    "allowExtraSlides": True,
    "lowerBetterKpis": [],
    "filePattern": "*VCP*Execution*.xls*",
    "searchDirs": ["~/Downloads", "~/Desktop", "~/OneDrive - Philips/Desktop", "~/OneDrive - Philips/Documents", "~/Documents"],
}


def load_defaults(path=None):
    """Merge AUTO_DEFAULTS with a JSON file (either a bare dict or a config with a 'vcp' key)."""
    d = dict(AUTO_DEFAULTS)
    if path:
        try:
            cfg = json.loads(Path(path).read_text(encoding="utf-8"))
            cfg = cfg.get("vcp", cfg) if isinstance(cfg, dict) else {}
            d.update({k: v for k, v in cfg.items() if k in AUTO_DEFAULTS})
        except (OSError, ValueError) as e:
            d["_warning"] = f"defaults file {path} not used: {e}"
    return d


def find_tracker(pattern=None, dirs=None, depth=2):
    """Newest Excel tracker matching the pattern in the usual folders (Downloads, Desktop, OneDrive...)."""
    pattern = pattern or AUTO_DEFAULTS["filePattern"]
    dirs = dirs or AUTO_DEFAULTS["searchDirs"]
    hits = []
    for d in dirs:
        base = Path(os.path.expanduser(os.path.expandvars(d)))
        if not base.is_dir():
            continue
        for lvl in range(depth + 1):
            hits += [p for p in base.glob("/".join(["*"] * lvl + [pattern])) if p.is_file() and not p.name.startswith("~$")]
    if not hits:
        return None, []
    hits = sorted(set(hits), key=lambda p: p.stat().st_mtime, reverse=True)
    return hits[0], hits[:5]


def auto(xlsx=None, function=None, outdir="output", defaults=None, open_after=False, template=None, previous=None):
    """One call, no questions where a default exists. Returns generate()'s dict plus 'spoken' and 'tracker'."""
    d = dict(defaults or AUTO_DEFAULTS)
    candidates = []
    if not xlsx:
        xlsx, candidates = find_tracker(d.get("filePattern"), d.get("searchDirs"))
        if not xlsx:
            return {"questions": ["I could not find a VCP execution tracker (pattern " + str(d.get("filePattern")) + ") in Downloads, Desktop or OneDrive. Tell me the file path."],
                    "spoken": "I couldn't find the VCP tracker Excel in your usual folders. Where is it?", "pptx": None, "tracker": None}
    function = function or d.get("defaultFunction")
    if not function:
        sheets = [n for n, st in Tracker(str(xlsx)).sheets() if st == "visible"]
        q = "Which tab (function) should I build? The tracker has: " + ", ".join(sheets) + "."
        return {"questions": [q], "spoken": q, "pptx": None, "tracker": str(xlsx), "sheets": sheets, "reconciled": False}
    res = generate(str(xlsx), function, outdir=outdir, quarter=d.get("quarter") or None,
                   lower_better=d.get("lowerBetterKpis") or [], allow_extra=bool(d.get("allowExtraSlides", True)),
                   impact_mode=d.get("impactSlide") or "drop", template=template, previous=previous)
    res["tracker"] = str(xlsx)
    res["other_trackers"] = [str(p) for p in candidates[1:]]
    if res["questions"]:
        res["spoken"] = "I need one answer before I build the " + str(function) + " deck: " + res["questions"][0]
        return res
    # counts for the spoken summary come from the validation summary header line
    m = re.search(r"Records processed: (.*)", res.get("report", ""))
    counts = m.group(1) if m else "the tracker data"
    name = Path(xlsx).name
    res["spoken"] = (f"Done. The {res['function']} VCP deck is ready from {name}: {counts}. "
                     + ("Every value reconciled with the Excel." if res["reconciled"] else "Some values did not reconcile, please check the summary.")
                     + (f" I used the newest of {len(candidates)} trackers I found, the one in {Path(xlsx).parent.name}." if len(candidates) > 1 else ""))
    if res.get("delta"):
        ts = res["delta"]["theme_summary"]
        res["spoken"] += (f" Compared with the previous tracker: {ts['moved']} themes moved forward, {ts['unchanged']} unchanged, "
                          f"{ts['declined']} declined, and {len(res['delta']['changes'])} changes are on the delta slides.")
    if open_after and res.get("pptx"):
        try:
            os.startfile(str(res["pptx"]))  # noqa: S606 - opens PowerPoint on Windows
        except OSError as e:
            res["spoken"] += f" I could not open it automatically ({e})."
    return res


# --------------------------------------------------------------------------- CLI
def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)

    p_ins = sub.add_parser("inspect", help="list sheets or parse one function sheet and report issues")
    p_ins.add_argument("xlsx"); p_ins.add_argument("--function"); p_ins.add_argument("--include-hidden", action="store_true")
    p_ins.add_argument("--quarter"); p_ins.add_argument("--lower-better", default="")
    p_ins.add_argument("--template", default=str(DEFAULT_TEMPLATE))

    p_b = sub.add_parser("build", help="build the deck for one function")
    p_b.add_argument("xlsx"); p_b.add_argument("--function", required=True)
    p_b.add_argument("--template", default=str(DEFAULT_TEMPLATE)); p_b.add_argument("--outdir", default="output")
    p_b.add_argument("--quarter", help="reporting quarter label, e.g. Q2 (auto-detected when unambiguous)")
    p_b.add_argument("--lower-better", default="", help="comma-separated KPI names where a lower actual is good")
    p_b.add_argument("--allow-extra-slides", action="store_true", help="add continuation slides when rows exceed template capacity")
    p_b.add_argument("--impact-slide", choices=["keep", "drop"], help="what to do with the Impact slide when column L is blank")
    p_b.add_argument("--truncate-dep-actions", type=int, help="max characters of the dependency 'Actions' cell (template style '...'); default: full text")
    p_b.add_argument("--title", help="override the cover title"); p_b.add_argument("--include-hidden", action="store_true")
    p_b.add_argument("--previous", help="previous month's tracker: adds the month-on-month delta slides (key theme movement chart + what changed)")
    p_b.add_argument("--governance", help="VCP Governance Tracker workbook: adds the decisions slide, the previous-meeting action tracker, "
                                          "the theme/KPI linkage slide and the interactive KPI cockpit with RCA drill-downs")
    p_b.add_argument("--decision", action="append", default=[], metavar="TEXT",
                     help="a decision taken in the meeting, verbatim (repeat for several)")
    p_b.add_argument("--with-change-table", action="store_true",
                     help="also add the 'What changed since the previous tracker' table slides (default: only the theme-movement chart)")
    p_b.add_argument("--with-governance-slides", action="store_true",
                     help="also add the decisions slide and the previous-meeting action tracker (needs --governance)")
    p_b.add_argument("--governance-function", metavar="SHEET",
                     help="scope the action tracker to one governance sheet (default: every function, that sheet listed first)")
    p_b.add_argument("--render", action="store_true", help="also export PNGs through PowerPoint for visual QA")

    p_r = sub.add_parser("render", help="export slide PNGs through PowerPoint (Windows)")
    p_r.add_argument("pptx"); p_r.add_argument("--outdir", default="output/render")

    p_a = sub.add_parser("auto", help="find the tracker, apply configured defaults, build without questions (Momo / one-click)")
    p_a.add_argument("xlsx", nargs="?", help="tracker path; omitted = newest tracker in Downloads/Desktop/OneDrive")
    p_a.add_argument("--function", help="sheet name; omitted = defaultFunction from --defaults")
    p_a.add_argument("--defaults", help="JSON file with defaults (bare dict or a config with a 'vcp' key)")
    p_a.add_argument("--outdir", default="output"); p_a.add_argument("--template", default=str(DEFAULT_TEMPLATE))
    p_a.add_argument("--open", action="store_true", help="open the deck in PowerPoint when done")
    p_a.add_argument("--previous", help="previous month's tracker for the month-on-month delta slides")
    p_a.add_argument("--json", action="store_true", help="print the result as JSON (paths, questions, spoken summary)")

    a = ap.parse_args(argv)

    if a.cmd == "render":
        for p in render(a.pptx, a.outdir):
            print(p)
        return 0

    if a.cmd == "auto":
        res = auto(a.xlsx, a.function, outdir=a.outdir, defaults=load_defaults(a.defaults), open_after=a.open, template=a.template, previous=a.previous)
        if a.json:
            print(json.dumps({k: (str(v) if isinstance(v, Path) else v) for k, v in res.items() if k != "report"}, indent=2, ensure_ascii=False))
        else:
            print("SPOKEN: " + res["spoken"])
            for qq in res["questions"]:
                print("QUESTION: " + qq)
            if res.get("pptx"):
                print(f"Tracker: {res['tracker']}\nDeck:    {res['pptx']}\nSummary: {res['summary']}")
        return 3 if res["questions"] else (0 if res.get("reconciled") else 1)

    tracker = Tracker(a.xlsx, include_hidden=a.include_hidden)
    if a.cmd == "inspect" and not a.function:
        print("Sheets in workbook:")
        for name, state in tracker.sheets():
            print(f"  {'[hidden] ' if state != 'visible' else '         '}{name}")
        return 0

    lower = [s.strip() for s in a.lower_better.split(",") if s.strip()]
    if a.cmd == "inspect":
        res = generate(a.xlsx, a.function, quarter=a.quarter, lower_better=lower, template=a.template,
                       include_hidden=a.include_hidden, dry_run=True)
        print(res["report"])
        return 3 if res["questions"] else 0

    res = generate(a.xlsx, a.function, outdir=a.outdir, quarter=a.quarter, lower_better=lower,
                   allow_extra=a.allow_extra_slides, impact_mode=a.impact_slide, title=a.title, template=a.template,
                   truncate_dep_actions=a.truncate_dep_actions, include_hidden=a.include_hidden, previous=a.previous,
                   governance=a.governance, decisions=a.decision, governance_function=a.governance_function,
                   with_change_table=a.with_change_table, with_governance_slides=a.with_governance_slides)
    if res["questions"]:
        print("NOT BUILT - the following must be resolved with the user first:")
        for qq in res["questions"]:
            print(f"  - {qq}")
        return 3
    print(f"Deck:    {res['pptx']}\nSummary: {res['summary']}\nModel:   {res['model']}\nReconciled: {'YES' if res['reconciled'] else 'NO - see summary'}")
    if a.render:
        for p in render(res["pptx"], Path(a.outdir) / "render"):
            print(p)
    return 0 if res["reconciled"] else 1


if __name__ == "__main__":
    sys.exit(main())
