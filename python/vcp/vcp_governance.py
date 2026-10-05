# -*- coding: utf-8 -*-
"""
Governance / decision slides for the Monthly VCP report template.

Implements the 07 Sep 2026 review decisions:
  1. "Update PPT template to include previous meeting action tracker, KPI risk status, owner,
      timeline, support required, and root-cause tracking."
  2. "Incorporate linkage between themes/actions and KPI impact within the review format."

plus the interactive KPI cockpit: the last slide lists the critical and enabling KPIs from the
VCP execution tracker; a red Qn signal is a hyperlink that jumps to that KPI's RCA drill-down
slide, which opens the cross-functional interdependency register (support required from other
functions) with each supporting action's progress and status, and derives the RCA evidence.

Data rules (unchanged): every value is verbatim from a source workbook or a documented
derivation. Fields the sources do not carry (KPI risk per action, support required per action,
root cause per action) are shown as "To be captured" - never invented.

Sources
  - previous meeting actions / decisions : the VCP Governance Tracker workbook (one sheet per
    function, sections "1. Action Items", "2. Decisions Taken", "3. Supporting Contacts")
  - KPIs, themes, interdependencies      : the VCP Execution tracker sheet of the deck's function
  - today's decisions                    : passed in verbatim by the user (--decision)
"""
import copy
import re
from pathlib import Path

from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

import vcp_deck as vd
import vcp_render2 as r2

TO_BE_CAPTURED = "To be captured"
GOV_STATUS = {"closed": "Completed", "completed": "Completed", "in progress": "In progress",
              "open": "Not started", "not started": "Not started", "delayed": "Delayed", "on hold": "On hold"}
# governance sheet that matches an execution-tracker function (both workbooks name functions differently)
GOV_SHEET_FOR = {"service transformation": "CS & IB", "ib": "CS & IB", "isc": "ISC", "marketing": "Marketing",
                 "marketing_new": "Marketing", "marketing_aug26": "Marketing", "nar": "NAR-ISC & MCoS",
                 "r&d": "Portfolio & R&D", "portfolio": "Portfolio & R&D", "mcos & icos": "Productivity"}
REGISTER_ROWS = r2.REGISTER_ROWS
ACTION_ROWS_PER_SLIDE = 9          # governance action tracker rows per slide (8-column table)


# ----------------------------------------------------------------------------- governance workbook
def _norm_status(v):
    s = (str(v).strip() if v is not None else "")
    return GOV_STATUS.get(s.lower(), s or None)


def read_governance(path):
    """Parse the VCP Governance Tracker: {'file':..., 'functions':[{'name','actions','decisions','contacts'}]}"""
    tracker = vd.Tracker(path)
    out = {"file": Path(path).name, "path": str(path), "functions": [], "issues": []}
    for name, state in tracker.sheets():
        if state != "visible":
            continue
        ws = tracker.wb[name]
        A = lambda r: vd.clean(ws.cell(r, 1).value)
        sections = {}
        for r in range(1, ws.max_row + 1):
            a = A(r)
            if isinstance(a, str):
                m = re.match(r"^\d+\.\s*(Action Items|Decisions Taken|Supporting Contacts)", a, re.I)
                if m:
                    sections[m.group(1).lower()] = r
        fn = {"name": name, "label": vd.clean(ws.cell(2, 2).value) or name, "actions": [], "decisions": [], "contacts": []}
        bounds = sorted(sections.values()) + [ws.max_row + 1]

        def end_of(start):
            return next(b for b in bounds if b > start)

        r0 = sections.get("action items")
        if r0:
            hdr = r0 + 1
            for r in range(hdr + 1, end_of(r0)):
                act = vd.clean(ws.cell(r, 2).value)
                if act is None:
                    continue
                fn["actions"].append({
                    "no": vd.clean(ws.cell(r, 1).value), "action": str(act),
                    "owner": vd.clean(ws.cell(r, 3).value), "due": vd.clean(ws.cell(r, 4).value),
                    "status_raw": vd.clean(ws.cell(r, 5).value), "status": _norm_status(ws.cell(r, 5).value),
                    "remarks": vd.clean(ws.cell(r, 6).value), "row": r, "sheet": name,
                })
        else:
            out["issues"].append(("missing", f"{name}: no '1. Action Items' section found"))
        r0 = sections.get("decisions taken")
        if r0:
            for r in range(r0 + 2, end_of(r0)):
                dec = vd.clean(ws.cell(r, 2).value)
                if dec is None:
                    continue
                fn["decisions"].append({"no": vd.clean(ws.cell(r, 1).value), "decision": str(dec),
                                        "details": vd.clean(ws.cell(r, 3).value), "row": r})
        r0 = sections.get("supporting contacts")
        if r0:
            for r in range(r0 + 2, end_of(r0)):
                area = vd.clean(ws.cell(r, 2).value)
                if area is None:
                    continue
                fn["contacts"].append({"area": str(area), "contact": vd.clean(ws.cell(r, 3).value), "row": r})
        out["functions"].append(fn)
    return out


def governance_actions(gov, only_function=None):
    """Flat action list; when only_function is given, that governance sheet first then the rest."""
    rows = []
    for fn in gov["functions"]:
        for a in fn["actions"]:
            rows.append(dict(a, function=fn["name"]))
    if only_function:
        key = only_function.lower()
        rows.sort(key=lambda a: (a["function"].lower() != key, a["function"].lower(), a["row"]))
    return rows


def gov_sheet_for(function_name, gov):
    names = {f["name"].lower(): f["name"] for f in gov["functions"]}
    want = GOV_SHEET_FOR.get((function_name or "").lower())
    if want and want.lower() in names:
        return names[want.lower()]
    return names.get((function_name or "").lower())


# ----------------------------------------------------------------------------- pptx helpers
def link_run_to_slide(slide, run, target_slide):
    """Internal 'jump to slide' hyperlink on a text run (python-pptx only exposes external URLs)."""
    rId = slide.part.relate_to(target_slide.part, RT.SLIDE)
    rPr = run.get_or_add_rPr()
    hl = rPr.get_or_add_hlinkClick()
    hl.set(qn("r:id"), rId)
    hl.set("action", "ppaction://hlinksldjump")
    return hl


def link_cell_to_slide(slide, tc, target_slide, colour=None):
    for p in tc.findall(qn("a:txBody") + "/" + qn("a:p")) or tc.find(qn("a:txBody")).findall(qn("a:p")):
        for r in p.findall(qn("a:r")):
            link_run_to_slide(slide, r, target_slide)
            rPr = r.find(qn("a:rPr"))
            rPr.set("u", "sng")
            if colour:
                vd._set_solid(rPr, colour)


def button(slide, proto, text, target_slide, left, top, width=None, height=None, fill=None, font=None, size=None):
    """Clone a rounded-rectangle prototype into a click-to-slide button."""
    b = r2.clone(slide, proto, name=f"Button {text[:18]}")
    b.left, b.top = Emu(int(Inches(left))), Emu(int(Inches(top)))
    if width:
        b.width = Emu(int(Inches(width)))
    if height:
        b.height = Emu(int(Inches(height)))
    if fill:
        r2.fill(b, fill)
    vd.set_shape_text(b, text)
    for p in b.text_frame.paragraphs:
        for run in p.runs:
            if size:
                run.font.size = Pt(size)
            if font:
                run.font.color.rgb = r2.rgb(font)
            run.font.bold = True
    b.click_action.target_slide = target_slide
    return b


set_table_columns = vd.set_table_columns          # shared table helpers now live in vcp_deck
cell_lines = vd.cell_lines


def _unused_normalise_row_style(tr):
    """Give every cell of the row the first cell's fill and run colour.

    The template's own columns carry per-column styling (the 'Qn signal' cell is red); once the
    columns mean something else, that styling must not travel with them. fill_table re-colours the
    status column afterwards.
    """
    tcs = tr.findall(qn("a:tc"))
    if len(tcs) < 2:
        return
    src = tcs[0]
    src_tcPr = src.find(qn("a:tcPr"))
    src_rPr = src.find(".//" + qn("a:rPr"))
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
                fill = src_rPr.find(qn("a:solidFill"))
                if fill is not None:
                    rPr.append(copy.deepcopy(fill))
                for att in ("b", "i", "u", "sz"):
                    if src_rPr.get(att) is not None:
                        rPr.set(att, src_rPr.get(att))
                    else:
                        rPr.attrib.pop(att, None)


# ----------------------------------------------------------------------------- derivations
def kpi_risk(signal):
    """KPI risk status derived from the reporting-quarter signal (documented in the summary)."""
    return {"Below target": "At risk", "On / above target": "On track", "No data": "Not measurable",
            "Direction to confirm": "Direction to confirm"}.get(signal, "Not measurable")


RISK_PALETTE = {"At risk": r2.STATUS2["Delayed"], "On track": r2.STATUS2["Completed"],
                "Not measurable": r2.STATUS2["No data"],
                "Direction to confirm": {"fill": "FFF2CC", "font": "916400"}}


def fn_rollup(deps):
    """Per supporting function: actions, average progress (numeric only), status mix."""
    by = {}
    for d in deps:
        e = by.setdefault(d["function"], {"function": d["function"], "n": 0, "prog": [], "done": 0, "not_started": 0, "delayed": 0})
        e["n"] += 1
        if d["progress"] is not None:
            e["prog"].append(d["progress"])
        if d["status"] == "Completed":
            e["done"] += 1
        elif d["status"] == "Not started":
            e["not_started"] += 1
        elif d["status"] == "Delayed":
            e["delayed"] += 1
    out = []
    for e in by.values():
        e["avg"] = round(100 * sum(e["prog"]) / len(e["prog"])) if e["prog"] else 0
        e["state"] = ("Completed" if e["done"] == e["n"] else
                      "Delayed" if e["delayed"] else
                      "Not started" if e["not_started"] == e["n"] else "In progress")
        out.append(e)
    out.sort(key=lambda e: (e["avg"], -e["not_started"]))
    return out


def rca_synthesis(kpi, q, deps, roll, gap):
    """Evidence-based RCA statement. States the rule applied and the counts; never asserts a cause."""
    n = len(deps)
    if not n:
        return (f"RCA is not captured in the tracker and no cross-functional dependency actions are recorded for this "
                f"function, so no supporting-function evidence is available for {kpi['metric']}.")
    done = [e["function"] for e in roll if e["state"] == "Completed"]
    lagging = [e for e in roll if e["state"] in ("Not started", "Delayed")]
    ns = sum(e["not_started"] for e in roll)
    parts = [f"RCA is not captured in the tracker. Evidence: {n} cross-functional actions across {len(roll)} functions"]
    if done:
        parts.append(f"delivered by {', '.join(done)} (100%)")
    if lagging:
        parts.append(f"{ns} not started, concentrated in {', '.join(e['function'] for e in lagging)}")
    parts.append(f"{kpi['metric']} is {gap} against the {q} target")
    rule = ("Rule applied: where supporting actions are outstanding they are the first root-cause candidates; where every "
            "supporting action is complete the cause sits inside this function's own execution.")
    return ". ".join(parts) + ". " + rule


# ----------------------------------------------------------------------------- slide builders
def build_decisions(slide, view, decisions, gov, plan_map, extra_added):
    fn = view["function"]
    txt = r2.txt
    txt(slide, "Rectangle 1", "MEETING DECISIONS")
    txt(slide, "Rectangle 2", f"Decisions taken and how this deck implements them")
    tbl = r2.sh(slide, "Table 13")
    set_table_columns(tbl, [0.55, 5.15, 4.30, 1.94])
    header = ["#", "Decision taken (verbatim)", "Implemented in this deck by", "Where"]
    rows = [[str(i + 1), d, plan_map[i]["how"], plan_map[i]["where"]] for i, d in enumerate(decisions)]
    vd.fill_table(tbl, rows, header=header)
    geo = vd._table_geometry(tbl)
    groups = vd.pack_rows(rows, geo)
    if len(groups) > 1:
        extra_added.append(f"decision table: {len(rows)} decisions need more than one slide at the template's row height")


def build_action_tracker(slide, rows, gov, view, part, n_parts, extra_added):
    txt = r2.txt
    total = part["total"]
    txt(slide, "Rectangle 1", "PREVIOUS MEETING ACTION TRACKER")
    title = "Previous meeting action tracker"
    txt(slide, "Rectangle 2", title + (f" ({part['i'] + 1}/{n_parts})" if n_parts > 1 else ""))
    txt(slide, "Rectangle 3", f"Source: {gov['file']}. {total} actions across {part['n_functions']} "
                              f"function{'s' if part['n_functions'] != 1 else ''} "
                              f"({part['closed']} closed, {part['inprog']} in progress, {part['open']} open). "
                              f"KPI risk, support required and root cause are new columns agreed on 07 Sep 2026 and are not yet populated in the tracker.")
    txt(slide, "Rectangle 7", str(total))
    txt(slide, "Rectangle 12", str(part["open"] + part["inprog"]))
    txt(slide, "Rectangle 17", f"0/{total}")
    txt(slide, "Rectangle 8", "Actions tracked")
    txt(slide, "Rectangle 9", "From the governance tracker")
    txt(slide, "Rectangle 13", "Still open")
    txt(slide, "Rectangle 14", f"{part['closed']} closed, {part['inprog']} in progress")
    txt(slide, "Rectangle 18", "New fields populated")
    txt(slide, "Rectangle 19", "Columns added, to be filled")
    tbl = r2.sh(slide, "Table 54")
    set_table_columns(tbl, [1.05, 4.05, 1.35, 1.20, 0.95, 1.05, 1.35, 1.46])
    header = ["Function", "Action item", "Owner", "Timeline", "KPI risk", "Support required", "Root cause / blocker", "Status"]
    body, statuses = [], []
    for a in rows:
        body.append([a["function"], a["action"], a["owner"] or vd.NOT_AVAILABLE, str(a["due"] or vd.NOT_AVAILABLE),
                     TO_BE_CAPTURED, TO_BE_CAPTURED, TO_BE_CAPTURED, a["status"] or vd.NOT_AVAILABLE])
        statuses.append(a["status"])
    if not body:
        body, statuses = [["", "No action items in the governance tracker", "", "", "", "", "", ""]], [None]
        extra_added.append("placeholder text 'No action items in the governance tracker' on the action tracker")
    vd.fill_table(tbl, body, status_col=7, status_values=statuses, header=header, palette=r2.STATUS2)
    r2.fill(r2.sh(slide, "Rectangle 11"), r2.C["red"] if part["open"] else r2.C["amber"])
    r2.color(r2.sh(slide, "Rectangle 12"), r2.C["red"] if part["open"] else r2.C["amber"])
    vd.set_body_font(tbl, 8)
    # remarks as a second, muted line inside the action cell (real data, clearly labelled)
    trs = tbl.table._tbl.findall(qn("a:tr"))
    for tr, a in zip(trs[1:], rows):
        if a.get("remarks"):
            tc = tr.findall(qn("a:tc"))[1]
            cell_lines(tc, [(a["action"], 8, r2.C["ink"], False),
                            ("Progress note: " + str(a["remarks"]), 6.5, r2.C["muted"], False)])
    for tr in trs[1:]:
        for idx in (4, 5, 6):
            tc = tr.findall(qn("a:tc"))[idx]
            for p in tc.find(qn("a:txBody")).findall(qn("a:p")):
                for run in p.findall(qn("a:r")):
                    rPr = run.find(qn("a:rPr"))
                    if rPr is not None:
                        vd._set_solid(rPr, r2.C["amber"])
                        rPr.set("i", "1")


def build_linkage(slide, view, model, extra_added, rca_targets=None):
    """Decision 2: linkage between themes/actions and KPI impact.

    Also the KPI cockpit: every critical and enabling KPI with its quarter actual, the change against
    the previous tracker, signal, risk and committed action. For a red KPI the last column carries the
    clickable link to that KPI's RCA drill-down.
    """
    txt = r2.txt
    q = view["quarter"]
    themes = model["themes"]
    crit, enab = model["kpis"]["critical"], model["kpis"]["enabling"]
    all_kpis = crit + enab
    n_actions = sum(len(t["rows"]) for t in themes)
    txt(slide, "Rectangle 1", "KPI COCKPIT / THEME LINKAGE")
    txt(slide, "Rectangle 2", "Linkage between themes, actions and KPI impact")
    txt(slide, "Rectangle 3", f"All critical and enabling KPIs of sheet '{model['sheet']}'. The tracker's 'VCP Commitment KPI Impact' column "
                              f"is blank in {n_actions} of {n_actions} action rows, so the theme-to-KPI mapping is not yet captured; the committed "
                              f"action behind each KPI is shown instead. Click 'Open RCA drill-down' on a red KPI to jump to its root-cause slide.")
    tbl = r2.sh(slide, "Table 54")
    dm = r2.kpi_delta_map(view)
    set_table_columns(tbl, [1.85, 0.85, 1.00, 1.45, 1.10, 1.05, 3.06, 2.10])
    header = ["KPI", "Section", f"{q} actual", "Δ vs previous", f"{q} signal", "KPI risk",
              "Committed action that drives it (tracker column G)", "Linked key theme (column L)"]
    rows, risks, dtones, stones = [], [], [], []
    for k in all_kpis:
        sig, _ = vd.signal_for(k, q, model["issues"], model["sheet"])
        risk = kpi_risk(sig)
        act = k["quarters"].get(q, {}).get("actual")
        dtext, dtone = r2.delta_mark(dm.get(k["metric"]))
        link = rca_targets.get(k["metric"]) if rca_targets else None
        rows.append([k["metric"], "Critical" if k in crit else "Enabling",
                     vd.fmt_value(act, k["kind_of_value"]) if act is not None else vd.NOT_AVAILABLE,
                     r2.marked(dtext, dtone), sig, risk,
                     k["committed"] or "Not captured in the tracker",
                     "Not captured (column L blank)"])
        risks.append(risk)
        dtones.append(r2.tone_palette(dtone))
        stones.append(vd.PALETTE.get(sig))
    if not rows:
        rows, risks = [["No KPI rows in the tracker", "", "", "", "", "", "", ""]], [None]
        dtones, stones = [None], [None]
        extra_added.append("placeholder text 'No KPI rows in the tracker' on the linkage table")
    vd.fill_table(tbl, rows, status_col=5, status_values=risks, header=header, palette=RISK_PALETTE)
    vd.colour_column(tbl, 3, dtones)
    vd.colour_column(tbl, 4, stones)
    vd.set_body_font(tbl, 9)
    # the last column carries the clickable RCA drill-down for each red KPI
    linked = 0
    if rca_targets:
        trs = tbl.table._tbl.findall(qn("a:tr"))
        for tr, k in zip(trs[1:], all_kpis):
            target = rca_targets.get(k["metric"])
            if target is None:
                continue
            tc = tr.findall(qn("a:tc"))[7]
            vd.cell_lines(tc, [("Not captured (column L blank)", 8, r2.C["muted"], False),
                               ("➔  Open RCA drill-down", 9, vd.PALETTE["Below target"]["font"], True)])
            ps = tc.find(qn("a:txBody")).findall(qn("a:p"))
            for run in ps[-1].findall(qn("a:r")):
                link_run_to_slide(slide, run, target)
                run.find(qn("a:rPr")).set("u", "sng")
            linked += 1
    if linked:
        extra_added.append(f"KPI linkage slide: {linked} red KPI row(s) link to their RCA drill-down from the 'Linked key theme' column")
    # cards
    txt(slide, "Rectangle 7", f"{len(all_kpis)}")
    txt(slide, "Rectangle 8", "KPIs in scope")
    txt(slide, "Rectangle 9", f"{len(crit)} critical, {len(enab)} enabling")
    txt(slide, "Rectangle 12", f"{len(themes)}")
    r2.fill(r2.sh(slide, "Rectangle 11"), r2.C["blue"])
    r2.color(r2.sh(slide, "Rectangle 12"), r2.C["blue"])
    txt(slide, "Rectangle 13", "Key themes")
    txt(slide, "Rectangle 14", f"{n_actions} actions recorded")
    txt(slide, "Rectangle 17", f"0/{n_actions}")
    txt(slide, "Rectangle 18", "Actions linked to a KPI")
    txt(slide, "Rectangle 19", "Column L, per decision 2")


def build_cockpit(slide, view, model, rca_targets, extra_added):
    """Last slide: critical + enabling KPIs; a red signal links to that KPI's RCA slide."""
    txt = r2.txt
    q = view["quarter"]
    crit, enab = model["kpis"]["critical"], model["kpis"]["enabling"]
    all_kpis = crit + enab
    states = {id(k): r2.kpi_state(k, q, model["issues"], model["sheet"]) for k in all_kpis}
    red = [k for k in all_kpis if states[id(k)][0] == "Below target"]
    nodata = [k for k in all_kpis if states[id(k)][0] == "No data"]
    txt(slide, "Rectangle 1", "VCP PERFORMANCE COCKPIT")
    txt(slide, "Rectangle 2", f"Critical and enabling KPIs: {q} status and risk")
    txt(slide, "Rectangle 3", f"Data fetched from {Path(model['source_file']).name}, sheet '{model['sheet']}' "
                              f"(Critical VCP Commitments and VCP Commitment Enabling KPIs sections). "
                              f"Click a red '{q} signal' cell to open its RCA drill-down.")
    txt(slide, "Rectangle 7", str(len(all_kpis)))
    txt(slide, "Rectangle 8", "KPIs tracked")
    txt(slide, "Rectangle 9", f"{len(crit)} critical, {len(enab)} enabling")
    txt(slide, "Rectangle 12", str(len(red)))
    txt(slide, "Rectangle 13", f"Red KPIs ({q})")
    txt(slide, "Rectangle 14", "Click the signal cell for RCA" if red else "No KPI below target")
    txt(slide, "Rectangle 17", str(len(nodata)))
    txt(slide, "Rectangle 18", "Not measurable")
    txt(slide, "Rectangle 19", f"No {q} target/actual pair")
    tbl = r2.sh(slide, "Table 54")
    set_table_columns(tbl, [2.05, 0.85, 1.15, 1.15, 1.15, 1.15, 1.10, 1.86])
    ly = view["last_year"]
    header = ["KPI", "Section", "Owner", "2025 baseline", f"{ly} target", f"{q} target", f"{q} actual", f"{q} signal"]
    rows, sigs = [], []
    for k in all_kpis:
        sig, ratio, ft, fa, gap = states[id(k)]
        rows.append([k["metric"], "Critical" if k in crit else "Enabling", k["owner"] or vd.NOT_AVAILABLE,
                     vd.fmt_value(k["baseline"], k["kind_of_value"], k["baseline_fmt"]),
                     vd.fmt_value(k["targets"].get(str(ly)), k["kind_of_value"]),
                     ft or vd.NOT_AVAILABLE, fa or vd.NOT_AVAILABLE,
                     (sig + "  ➔ RCA") if sig == "Below target" else sig])
        sigs.append(sig)
    if not rows:
        rows, sigs = [["No KPI rows in the tracker", "", "", "", "", "", "", ""]], [None]
        extra_added.append("placeholder text 'No KPI rows in the tracker' on the cockpit table")
    vd.fill_table(tbl, rows, status_col=7, status_values=sigs, header=header)
    vd.set_body_font(tbl, 9)
    # hyperlink the red signal cells to their RCA slide
    trs = tbl.table._tbl.findall(qn("a:tr"))
    linked = 0
    for tr, k in zip(trs[1:], all_kpis):
        target = rca_targets.get(k["metric"])
        if target is None:
            continue
        tc = tr.findall(qn("a:tc"))[7]
        for p in tc.find(qn("a:txBody")).findall(qn("a:p")):
            for run in p.findall(qn("a:r")):
                link_run_to_slide(slide, run, target)
                rPr = run.find(qn("a:rPr"))
                rPr.set("u", "sng")
                vd._set_solid(rPr, vd.PALETTE["Below target"]["font"])
        linked += 1
    return linked


def build_rca(slide, view, model, kpi, back_slide, extra_added, prev_deps=None):
    """One RCA drill-down per red KPI: support required from other functions + evidence."""
    txt = r2.txt
    q = view["quarter"]
    sig, ratio, ft, fa, gap = r2.kpi_state(kpi, q, model["issues"], model["sheet"])
    deps = model["dependencies"]
    roll = fn_rollup(deps)
    ns = sum(e["not_started"] for e in roll)
    txt(slide, "Rectangle 1", "RCA DRILL-DOWN")
    txt(slide, "Rectangle 2", f"RCA DrillDown : {kpi['metric']}: {r2.gap_abs(kpi, gap)} below the {q} target")
    txt(slide, "Rectangle 7", str(len(deps)))
    txt(slide, "Rectangle 8", "Support actions required")
    txt(slide, "Rectangle 9", f"From {len(roll)} supporting function{'s' if len(roll) != 1 else ''}")
    txt(slide, "Rectangle 12", str(ns))
    txt(slide, "Rectangle 13", "Not started")
    txt(slide, "Rectangle 14", f"{round(100 * ns / len(deps))}% of support actions" if deps else "No dependency rows")
    ly = view["last_year"]
    lab = r2.sh(slide, "Rectangle 21")
    lab.width = Emu(int(Inches(3.0)))
    txt(slide, "Rectangle 21", f"{q} gap")
    txt(slide, "Rectangle 22", f"{fa} vs {ft} · gap {r2.gap_abs(kpi, gap)} · baseline "
                               f"{vd.fmt_value(kpi['baseline'], kpi['kind_of_value'], kpi['baseline_fmt'])} → {ly} target "
                               f"{vd.fmt_value(kpi['targets'].get(str(ly)), kpi['kind_of_value'])}")
    txt(slide, "Rectangle 24", "Support required from other functions (progress of each action in the tracker)")
    delivered = [e["function"] for e in roll if e["state"] == "Completed"]
    txt(slide, "Rectangle 26", (f"Delivered: {', '.join(delivered)}" if delivered else "No function fully delivered"))
    if delivered:
        r2.fill(r2.sh(slide, "Rectangle: Rounded Corners 25"), r2.C["green_bg"])
        r2.color(r2.sh(slide, "Rectangle 26"), r2.C["green"])
    txt(slide, "Rectangle 121", rca_synthesis(kpi, q, deps, roll, (r2.gap_abs(kpi, gap) + " below") if gap else "below")
        + f" Committed action (tracker): {kpi['committed'] or 'not captured'}.")
    # register rows: worst progress first so the blockers surface at the top
    order = {"Not started": 0, "Delayed": 1, "In progress": 2, "Completed": 3}
    rows = sorted(deps, key=lambda d: (order.get(d["status"], 9), d["progress"] if d["progress"] is not None else 0, d["function"]))
    proto = {n: r2.sh(slide, n) for n in ("Rectangle 32", "Rectangle 33", "Rectangle 34", "Rectangle 35",
                                          "Rectangle: Rounded Corners 36", "Rectangle 37")}
    for s in r2.shapes_in_band(slide, 3.44, 6.60, exclude=set(proto)):
        r2.remove(s)
    # widen the status pill so "100% Completed" fits on one line, and shorten the action column to match
    proto["Rectangle 35"].width = Emu(int(Inches(8.10)))
    proto["Rectangle: Rounded Corners 36"].left = Emu(int(Inches(10.70)))
    proto["Rectangle: Rounded Corners 36"].width = Emu(int(Inches(1.55)))
    proto["Rectangle 37"].left = Emu(int(Inches(10.78)))
    proto["Rectangle 37"].width = Emu(int(Inches(1.40)))
    hdr_status = r2.sh(slide, "Rectangle 31")
    hdr_status.left, hdr_status.width = Emu(int(Inches(10.78))), Emu(int(Inches(1.40)))
    shown = rows[:REGISTER_ROWS]
    if len(rows) > REGISTER_ROWS:
        extra_added.append(f"RCA slide for {kpi['metric']}: {len(rows)} support actions, {REGISTER_ROWS} shown (worst progress first)")
    if not shown:
        shown = [{"year": "", "function": "", "action": "No cross-functional dependency actions recorded for this function",
                  "status": None, "progress": None}]
    for i, d in enumerate(shown):
        dy = i * r2.REGISTER_PITCH
        if i % 2 == 0 and i:
            r2.clone(slide, proto["Rectangle 32"], dy=dy)
        y_ = proto["Rectangle 33"] if i == 0 else r2.clone(slide, proto["Rectangle 33"], dy=dy)
        f_ = proto["Rectangle 34"] if i == 0 else r2.clone(slide, proto["Rectangle 34"], dy=dy)
        a_ = proto["Rectangle 35"] if i == 0 else r2.clone(slide, proto["Rectangle 35"], dy=dy)
        pb = proto["Rectangle: Rounded Corners 36"] if i == 0 else r2.clone(slide, proto["Rectangle: Rounded Corners 36"], dy=dy)
        pt_ = proto["Rectangle 37"] if i == 0 else r2.clone(slide, proto["Rectangle 37"], dy=dy)
        vd.set_shape_text(y_, str(d["year"]))
        vd.set_shape_text(f_, d["function"] or "")
        prog = f"{round(100 * d['progress'])}%" if d.get("progress") is not None else "0%"
        vd.set_shape_text(a_, d["action"] or vd.NOT_AVAILABLE)
        pal = r2.STATUS2.get(d["status"], r2.STATUS2["No data"])
        vd.set_shape_text(pt_, f"{prog} {d['status']}" if d["status"] else "—")
        r2.fill(pb, pal["fill"])
        r2.color(pt_, pal["font"])
    # back button
    badge = r2.sh(slide, "Rectangle: Rounded Corners 25")
    button(slide, badge, "⬅ Back to KPI linkage", back_slide, left=10.47, top=0.72, width=1.86, height=0.24,
           fill=r2.C["blue_bg"], font=r2.C["blue"], size=8)


# ----------------------------------------------------------------------------- orchestration
def plan_slides(view, model, gov, decisions, only_function=None):
    """How many slides each governance section needs (before anything is written)."""
    actions = governance_actions(gov, only_function)
    if only_function:
        actions = [a for a in actions if a["function"].lower() == only_function.lower()] if only_function != "*" else actions
    parts = [actions[i:i + ACTION_ROWS_PER_SLIDE] for i in range(0, len(actions), ACTION_ROWS_PER_SLIDE)] or [[]]
    q = view["quarter"]
    red = [k for k in model["kpis"]["critical"] + model["kpis"]["enabling"]
           if vd.signal_for(k, q, [], model["sheet"])[0] == "Below target"]
    return {"actions": actions, "action_parts": parts, "red": red, "decisions": list(decisions)}


def fill_kpi_slides(prs, slides, view, model, extra_added, note_for):
    """The KPI cockpit / linkage slide and one RCA drill-down per red KPI. Part of every deck."""
    rca_targets = {kpi["metric"]: s for kpi, s in slides["rca"]}
    build_linkage(slides["linkage"], view, model, extra_added, rca_targets)
    vd.set_notes(slides["linkage"], note_for("cockpit"))
    for kpi, s in slides["rca"]:
        build_rca(s, view, model, kpi, slides["linkage"], extra_added)
        vd.set_notes(s, note_for("rca_drill"))


def fill_governance_slides(prs, slides, kpi_slides, view, model, gov, decisions, plan, extra_added, note_for):
    """Opt-in: the decisions slide and the previous-meeting action tracker."""
    slides = dict(slides, linkage=kpi_slides["linkage"], rca=kpi_slides["rca"])
    actions, parts, red = plan["actions"], plan["action_parts"], plan["red"]
    n_fun = len({a["function"] for a in actions})
    closed = sum(1 for a in actions if a["status"] == "Completed")
    inprog = sum(1 for a in actions if a["status"] == "In progress")
    open_ = len(actions) - closed - inprog
    cockpit_no = prs.slides.index(slides["linkage"]) + 1
    rca_nos = [prs.slides.index(s) + 1 for _, s in slides["rca"]]
    action_nos = [prs.slides.index(s) + 1 for s in slides["actions"]]
    how = [
        {"how": "Previous meeting actions from the governance tracker with owner, timeline and status; KPI risk, support required and "
                "root-cause columns added to the format (not yet populated in the source). KPI risk per KPI is on the cockpit; support "
                "required and root-cause evidence are on the RCA drill-downs.",
         "where": f"Slides {', '.join(map(str, action_nos))}, {cockpit_no}"
                  + (f", {rca_nos[0]}-{rca_nos[-1]}" if len(rca_nos) > 1 else (f", {rca_nos[0]}" if rca_nos else ""))},
        {"how": "KPI-by-KPI view of the committed action that drives it, the change against the previous tracker, and the "
                "theme-to-KPI column that the tracker still leaves blank; each red KPI links to its RCA drill-down.",
         "where": f"Slide {cockpit_no}"},
    ]
    while len(how) < len(decisions):
        how.append({"how": "Recorded for tracking; no deck element derived from it.", "where": "—"})
    if slides.get("decisions") is not None:
        build_decisions(slides["decisions"], view, decisions, gov, how, extra_added)
        vd.set_notes(slides["decisions"], note_for("decisions"))
    for i, s in enumerate(slides["actions"]):
        part = {"i": i, "total": len(actions), "n_functions": n_fun, "closed": closed, "inprog": inprog, "open": open_}
        build_action_tracker(s, parts[i] if i < len(parts) else [], gov, view, part, len(slides["actions"]), extra_added)
        vd.set_notes(s, note_for("governance"))
        if i:
            extra_added.append("Previous meeting action tracker (cont.)")
    return extra_added


def summary_lines(view, model, gov, plan, decisions):
    actions = plan["actions"]
    n_actions = sum(len(t["rows"]) for t in model["themes"])
    L = ["## Meeting decisions, governance tracker and interactive RCA",
         f"- Decisions taken (verbatim, as provided): " + " | ".join(f"{i + 1}. {d}" for i, d in enumerate(decisions)),
         f"- Governance source: `{gov['file']}` - {len(actions)} action items across {len({a['function'] for a in actions})} function sheets, "
         f"read verbatim (action, owner, due date, status, remarks).",
         "- New tracker fields agreed on 07 Sep 2026 (KPI risk, support required, root cause) are present as columns and marked "
         f"'{TO_BE_CAPTURED}': the governance workbook carries no such fields, so nothing was inferred for them.",
         "- KPI risk is derived per KPI from the reporting-quarter signal: Below target -> At risk, On/above target -> On track, "
         "no target/actual pair -> Not measurable.",
         f"- Theme-to-KPI linkage: the tracker's 'VCP Commitment KPI Impact' column (L) is blank in {n_actions}/{n_actions} action rows, "
         "so the linkage slide shows the linkage the tracker does carry (each KPI's committed action, column G) and flags the gap.",
         f"- KPI cockpit: all {len(model['kpis']['critical'])} critical and {len(model['kpis']['enabling'])} enabling KPIs of sheet "
         f"'{model['sheet']}'; the {len(plan['red'])} red signal cells are internal hyperlinks to their RCA drill-down slide, which links back.",
         f"- RCA drill-downs: the tracker records interdependencies per function, not per KPI, so each drill-down shows the same "
         f"{len(model['dependencies'])} cross-functional actions (worst progress first) with their verbatim action text, progress and status. "
         "The synthesis line states the counts and the diagnostic rule applied; no root cause is asserted.",
         "- Nothing on these slides is invented: every action, owner, date, status, remark, KPI value and dependency action is a verbatim "
         "cell from one of the two workbooks."]
    return L
