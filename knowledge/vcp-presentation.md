# VCP deck from the Excel tracker (data-exact, template-driven)

Trigger phrases: "take the excel and make me the presentation", "build the VCP deck",
"create the ISC presentation from the tracker", "VCP pre-read for Service", "refresh the
VCP ppt", "excel se presentation bana do".

This is NOT the creative deck workflow in presentations.md. Here every number and
sentence comes verbatim from the Excel tracker and the design is the approved VCP
template. Nothing is invented, reworded or summarised. A deterministic script does the
whole job; you only run it, read its result, and speak.

## One command does everything
```powershell
$env:PYTHONUTF8 = "1"
python "python\vcp\vcp_deck.py" auto --defaults config.json --outdir "data\presentations" --open --json
```
- No path given -> it picks the NEWEST file matching `*VCP*Execution*.xls*` in Downloads,
  Desktop, OneDrive Desktop/Documents (the `vcp.searchDirs` list in config.json).
- Function (tab) not named -> the script returns a question listing the tabs; ASK IT
  ALOUD, short: "Which tab should I build - Service Transformation, ISC, IB, Quality,
  Regulatory, Finance, NAR, Europe, Growth, R and D, or Marketing?" Then re-run with
  `--function <answer>`. Never pick a tab yourself (unless `vcp.defaultFunction` is set
  in config.json). If the user names one ("ISC deck", "for Quality"), pass `--function ISC`.
- The same questions the web UI asks are the ones the script returns: which tab, which
  reporting quarter (only when no KPI has both target and actual), which file when
  several trackers match. Ask exactly those, one at a time, and nothing else.
- Layout is automatic: tables that nearly fit are kept on one slide by trimming the body
  font by 1-2pt; only genuinely long tables get a "(cont.)" slide. Speaker notes of each
  slide list the sheet and rows used.
  Sheet names: Service Transformation, IB, ISC, Quality, Regulatory, Finance,
  MCos & ICos, Marketing_New, NAR, Europe, Growth, R&D. Fuzzy names work ("service",
  "r and d" -> say "R&D").
- A spoken path ("the file on my desktop called VCP execution v3") -> find it with the
  Windows Search recipe in find-files.md and pass the path as the first argument.
- `--open` opens the deck in PowerPoint when it is done. `--json` returns:
  `spoken` (say this), `questions` (ask this), `pptx`, `summary`, `tracker`, `reconciled`.

## Defaults Pratibha has approved for automatic runs (config.json -> "vcp")
- impact slide dropped when the tracker's "VCP Commitment KPI Impact" column is blank
- continuation slides added when a table does not fit the template
- reporting quarter auto-detected from the latest quarter with both target and actual
- every KPI treated as higher-is-better unless listed in `lowerBetterKpis`
Change them in config.json, never in the script.

## What to say
- Success: say the `spoken` text from the JSON, naturally. Example: "Done, the Service
  Transformation VCP deck is ready from the version 2 tracker - 9 themes, 31 actions,
  9 KPIs, all reconciled with the Excel. I've opened it for you."
- `questions` non-empty (exit code 3) -> nothing was built. Ask the ONE thing needed,
  e.g. "Which quarter should I report - the tracker has no target and actual pairs yet?"
  Then re-run with `--quarter Q2` style extra flags via `build` (see below) or with the
  answer folded into config.json if it is permanent.
- Several trackers found: the newest was used; mention it in one clause if the names
  differ ("I used the version 2 file from Downloads").
- If PowerPoint refuses to open the file or the script errors: say so plainly and what
  you tried. Never hand-build the deck instead.

## When the user wants something other than the defaults
Use `build` with explicit flags (all documented in `python\vcp\MAPPING.md`):
```powershell
python "python\vcp\vcp_deck.py" build "<xlsx>" --function ISC --quarter Q2 --impact-slide keep --allow-extra-slides --lower-better "Inventory MAT %" --outdir "data\presentations"
```
`inspect` (same arguments, no deck) lists every question, missing and ambiguous item.

## Outputs (data\presentations)
`<Function>_VCP_Management_View.pptx`, `<Function>_validation_summary.md` (sheets used,
counts, reconciliation, derived values, missing/ambiguous items) and
`<Function>_model.json` (every value with its Excel cell). If the user asks "what did
you change" or "is it accurate", read the summary and answer from it.

## Other Excel files (Mode C: not the VCP tracker)
If `auto`/`inspect` says the sheet "does not follow the tracker layout", or the user
points at some other workbook (a sales sheet, a budget, a survey), build a DESIGNED deck
with the presentations.md method but with these data rules:
1. Read the workbook with openpyxl (`data_only=True`) — every sheet, headers, merged
   cells, number formats. Print the tables to yourself before designing anything.
2. Every number, percentage, date, name and label on a slide is copied verbatim from a
   cell (use the cell's number format for display: 0.493 formatted `0.0%` -> 49.3%).
   No rounding beyond the cell format, no invented totals — if you compute a total or
   an average, label it "(calculated)" on the slide.
3. Charts: build them with python-pptx native charts from the exact cell ranges, never
   as drawn images. Titles say which sheet/range they come from (small caption).
4. Never summarise free text from cells into your own words; quote or shorten with "..."
   only if a cell is longer than a slide can hold, and say so in the notes.
5. Structure the deck around what the sheets contain (one section per sheet/table),
   not around what you imagine the topic should include.
6. Finish with a QA pass: render slides via COM, check every figure on the slides
   against the printed tables, fix any mismatch, then list in the speaker notes of slide 1
   which sheets and ranges were used.
7. Save to data\presentations\<Workbook name>_deck.pptx and open it. Tell the user the
   sheets used and that all figures are verbatim from the file.

## Memory
After a run, note in data/memory.md which tracker file and function were used last, so
"make it again" or "same for ISC" needs no clarification.

## Template update (2026-09-07)
The deck now follows the "Monthly VCP report" layout (python/vcp/template/CS_VCP_Monthly_Report.pptx): a target-vs-actual chart
with remarks for the critical KPIs, the enabling table, the themes table, a dependency register, a VCP impact linkage slide and an
RCA + sheet performance slide. Nothing changes in how you run it; the only extra question the script may ask is keep/drop for the
impact linkage slide when the sheet has no VCP outcome table (Metric | Year | FY block) - answer with `--impact-slide keep` or `drop`
(config.json vcp.impactSlide already answers it as `drop`). If Pratibha asks for "the old format", add
`--template python\vcp\template\Service_Transformation_VCP_Management_View.pptx`.
