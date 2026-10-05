# Excel tracker -> VCP Management View mapping

Source: one function sheet of `VCP - Execution sheet ... .xlsx` (e.g. `Service Transformation`, `ISC`, `IB`, `Quality`, `Regulatory`, `Finance`, `MCos & ICos`, `NAR`, `Europe`, `Growth`, `R&D`, `Marketing_New`).
Design: `template/Service_Transformation_VCP_Management_View.pptx` (7 slides). Nothing else is used.

## Sheet layout the parser expects (all function sheets share it)

| Section | How it is found | Columns read |
|---|---|---|
| Year blocks | column A = `2026` / `2027` / `2028`, next row A = `Key Themes` | A Key Theme (merged, carried down), B Owner, C Actions Description, F Start Date, G Planned Completion Date, H Progress (0-1), I Status, J Key Theme Progress, K Units, L VCP Commitment KPI Impact |
| Critical KPIs | heading containing `Critical`, then a `Metric` header row; year labels in D/E/F of the next row (`-2026` = 2026) | per metric row: A Metric, B Owner, C Baseline (2025), D/E/F targets, G Committed Actions, H `T/A` anchor with I-L = Q1..Q4, row+1 `Target`, first `Actual` row below |
| Enabling KPIs | heading containing `Enabling`, same structure | same |
| Interdependencies | heading `Interdependencies across all function`; per year: A = year, next row A = `Functions` | A Function, B Key Themes, C Interdependency Actions, D Planned Completion Date, E Progress, F Status |

Not used: the interdependency columns inside the year blocks (M onwards - asks this function places on others), helper tables after the KPI blocks (e.g. Growth OIT tables, Service Sales/iGM/EBITA rows), narrative text rows, hidden sheets (`Master`, `Combined Deck`, `MoS End to End`, `List Format`), `Portfolio`, `KEY RISKS`.

## Slide mapping

| Slide | Template shapes | Content |
|---|---|---|
| 1 Cover | `Text Placeholder 10` | `<Function> VCP Program Execution <first year>-<last year>` (override with `--title`); subtitle unchanged |
| 2 Critical VCP Commitments | kicker `Rectangle 15` = `<FUNCTION> VCP`; cards `Rectangle 5/8/11` (numbers) and `Rectangle 6/9/12` (labels); `Table 28` | Cards: critical KPIs tracked / `Qn commitments below target` / `Qn commitments met or above target`. Table: Metric, 2025 baseline, `<y> target` x3, `Qn target`, `Qn actual`, `Qn signal` |
| 3 Enabling KPIs | cards, `Table 28` | Cards: enabling KPIs with Qn target and actual / below target / KPIs with no quarterly data. Table: KPI, 2025 baseline, `<last year> target`, Qn target, Qn actual, Qn signal, Committed action / data note |
| 4 Themes & actions | `Rectangle 1` title, `Table 10` | One row per key theme per year: Year, Key theme, Actions (all action descriptions joined with `; `), `NN% \| Status` |
| 5 Cross-functional dependencies | cards, `Table 28` | Cards: rows captured / not started / rows in last year. One row per function (alphabetical): Function, Dependency focus (unique key themes joined `; `), `N: action; action...`, `NN% \| Status` |
| 6 Impacting KPIs towards VCP | `Table 10`, `Rectangle 3` (`*DUMMY DATA` label is removed) | Key theme, Impacting KPIs = unique values of column L for the theme's actions, VCP Impact = no Excel source (left empty). If column L is blank everywhere the user must choose `--impact-slide keep` or `drop` |
| 7 RCA for red KPIs | `Table 13`, `Table 14` | Left: one row per KPI with signal `Below target`: metric, `actual vs target (delta)`, RCA and Management follow-up empty (no source). Right: data-readiness counts |

## Derivations (the only non-verbatim content; all listed in the validation summary)

- **Reporting quarter**: the latest quarter in which a critical KPI has both target and actual; must be unique across KPIs, else `--quarter Qn` is required (falls back to enabling KPIs).
- **Signal**: numeric actual vs numeric target of the reporting quarter. `actual >= target` -> `On / above target`, else `Below target`; direction is reversed for KPIs listed in `--lower-better`. Missing or non-numeric (`>36%`, `~€68M`, ranges) -> `No data`. The script lists KPIs whose names suggest lower-is-better so the user can confirm.
- **Display format**: cells with a `%` number format -> `xx.x%`; General-format numbers in a %-typed row -> fraction if <= 1 else shown as points (flagged); rows whose text cells contain `€...M` -> `€<value>M`; otherwise the number as stored. Text values are verbatim.
- **Theme progress**: column J if filled, else mean of the theme's numeric H values (Python `round`, half-to-even, as in the template: 62.5 -> 62%). No values -> 0%.
- **Theme / dependency status**: all Completed -> Completed; any Delayed -> Delayed; all Not started (or blank) -> Not started; all On hold -> On hold; otherwise In progress.
- **Status canonicalisation**: case-insensitive mapping of common spellings (`in progress`, `ongoing`, `completed`, `done`, `not started`, `delayed`, `on hold`); anything else is kept verbatim and flagged.
- **Delta on the RCA slide**: `(-0.30M)` for currency (2 decimals), `(-1.6 pts)` for percentages (1 decimal), plain difference otherwise.
- **Counts**: cards and the data-readiness table are counts of the parsed rows (Owner blanks / n KPIs; targets missing per year; quarterly target+actual cells missing out of n x 8; Impact/Owner/Start Date/Theme Progress blanks out of the action rows).
- **Enabling data note**: if Committed Actions is blank the note lists what is missing (`Baseline, targets, committed action and quarterly data need completion`).
- **Empty cells** -> `Not available`. Empty tables -> one placeholder row stating that nothing is recorded (listed in the summary).
- **Text normalisation**: zero-width spaces removed; runs of spaces/newlines collapsed to one space. Spelling untouched.

## Palette (from the template, never changed)

| Meaning | Fill | Font |
|---|---|---|
| Below target / Delayed | `F8D7DA` | `C7372F` |
| On / above target / Completed | `DCEEDC` | `0E7A3D` |
| In progress | `FFF2CC` | `916400` |
| No data / Not started / On hold | `ECEFF3` | `525866` |

## Capacity

Row heights are measured with real Calibri glyph widths (word-wrap simulation) against the template column widths. A table may reach 0.15" from the slide bottom. If it does not fit at the template body size, the body font is reduced by 1pt, then 2pt (never below 10pt) to keep it on one slide; only if that still does not fit are the rows split across continuation slides (title + ` (cont.)`, cards repeated), which requires `--allow-extra-slides`, otherwise the script stops with a question. Any font reduction is listed in the validation summary as a layout note. Nothing is truncated unless the user asks for `--truncate-dep-actions N`.

## Speaker notes
Every generated slide carries `[Sources]` notes naming the workbook, the sheet and the row range the slide was built from, plus the generation date. The template's own notes (which referred to an older file) are replaced.

## Loading
The tracker has ~10,000 merged ranges, which makes openpyxl slow (about 40 s per load). The generator works from a cached copy with the merge definitions removed (values are unaffected; verified identical extraction), stored under the system temp folder and keyed by file path, size and modification time.

---

## Template 2: "Monthly VCP report" (`vcp_agent/template/CS_VCP_Monthly_Report.pptx`) — current standard

The generator detects this template automatically (slide 2 carries a native chart named `Chart`) and renders it
with `vcp_render2.py`. The Management View template above remains selectable (`--template`, or the Template
drop-down in the web app). Same extraction, same derivations, same reconciliation; only the presentation differs.

| Template slide | Element | Source / derivation |
|---|---|---|
| 1 Cover | title | `{Function} VCP Program Execution {first year}-{last year}` |
| 2 Critical VCP commitments | kicker | `{FUNCTION} · CRITICAL VCP COMMITMENTS` |
| | title / subtitle | `{Qn} target vs actual for critical KPIs`; subtitle names KPIs without a Qn target/actual pair (not charted) |
| | card 1 | number of critical KPI rows |
| | card 2 | `{below}/{total}` critical KPIs below target, names below the card (green when none) |
| | chart | one category per critical KPI with a numeric Qn target+actual; series 1 = target normalised to 100 %, series 2 = actual ÷ target; actual bar green when on/above target (lower-is-better KPIs respected), red when below; labels show the verbatim formatted target and actual |
| | Monthly remarks | one block per critical KPI: `{gap} below the Qn target (actual vs target)` / `At the Qn target (...)` / `No Qn pair in the sheet` + `Committed action: {column G}`; label colour by signal |
| 3 Enabling KPIs | title | `{ready} of {n} enabling KPIs are Qn-ready; {red} below target` (ready = target and actual present) |
| | cards | ready/total, red count, no-quarterly-data count |
| | table | as in template 1 (KPI, 2025 baseline, last-year target, Qn target, Qn actual, Qn signal, Committed action) |
| 4 Themes and actions | table | as in template 1; status cell colours: In progress blue, Completed green, Not started grey, Delayed red |
| 5 Interdependencies | cards | dependency rows; not-started rows and their share |
| | Leadership focus | `{F1} and {F2} hold {k} of the {m} not-started actions` (two functions with most not-started rows) |
| | badge | `{last year}: no actions captured` or `{last year}: k actions captured` |
| | register | one shape row per dependency row in sheet order (year, function, action verbatim, status pill); 16 rows per slide, continuation slides beyond |
| | source line | file, sheet, rows of the "Interdependencies across all function" section |
| 6 VCP impact linkage | cards | last-year FY value of the first three metrics of the sheet's outcome table (`Metric | Year | Q1..Q4 | FY [| %]` block, e.g. Sales / iGM / Adj. EBITA), with the first-year value; `Not quantified` when absent |
| | driver map | one row per critical KPI (max 4 per slide): committed levers (column G) → KPI attainment `x% of Qn target; gap …` → outcome metric i of the outcome table (`Not quantified in the sheet` when absent) |
| | chart | FY values of every outcome metric by year (€M); chip = percentage column of the outcome table for the last year (e.g. `2028 iGM plan: 56.7%`) |
| | keep/drop | when the sheet has no outcome table the slide needs an explicit keep/drop decision (`--impact-slide`) |
| 7 RCA and sheet performance | RCA blocks | one per KPI below target (critical then enabling): gap pill, `RCA status: not captured in the sheet. Qn actual vs target: …`, `Committed action (sheet): …`; 4 per slide at template size, up to 6 compressed, continuation beyond |
| | sheet performance | blank quarterly cells x/y and populated %; critical / enabling populated %; missing-field bars (Owner, Baseline, Target per year, Committed actions) with counts; fully empty theme-tracker columns and planned-completion / progress blanks |
| footer | page number | `Monthly VCP report | i/N` recomputed after continuation or dropped slides |

Narrative fields of the template (management actions, "engine" names, remarks prose) are replaced by data-derived
statements only; they are never invented. Month-on-month delta slides (`--previous`) use the same design language:
a copy of slide 2 with the theme movement bar chart, and a copy of slide 4 with the change table.

---

## Decision slides, governance tracker and interactive RCA (added 07 Sep 2026)

Enabled with `--governance <VCP Governance Tracker.xlsx>` plus one `--decision "<text>"` per decision taken.
Built by `vcp_governance.py`; the deck is written as `<Function>_VCP_Review_with_Decisions.pptx` and the approved
template file is never modified. Slides are appended after the standard report, the cockpit last before the drill-downs.

| Slide | Element | Source / derivation |
|---|---|---|
| Decisions taken | decision text | verbatim as supplied (`--decision`) |
| | "Implemented in this deck by" / "Where" | mapping to the slides this generator produced (deck-internal, not business data) |
| Previous meeting action tracker | Function, Action item, Owner, Timeline, Status, progress note | verbatim from the governance workbook's "1. Action Items" section of each function sheet (columns B, C, D, E, F) |
| | KPI risk, Support required, Root cause / blocker | columns added per decision 1; the governance workbook has no such fields, so every cell reads "To be captured" |
| | cards | counts of actions, still-open actions and populated new fields |
| | scope | default: the governance sheet matching the deck's function (`GOV_SHEET_FOR`); `--governance-function "*"` includes every function, `--governance-function NAME` picks one |
| Theme / action to KPI linkage | KPI, Section, Qn signal, Committed action | KPI sections of the execution tracker (metric, column G) |
| | KPI risk | derived from the Qn signal: Below target → At risk, On/above target → On track, no pair → Not measurable |
| | Linked key theme (column L) | tracker column "VCP Commitment KPI Impact"; blank in every action row, so shown as "Not captured (column L blank)" |
| KPI cockpit (last slide) | table | every critical and enabling KPI of the sheet: metric, section, owner, 2025 baseline, last-year target, Qn target, Qn actual, Qn signal |
| | red signal cell | internal hyperlink (`ppaction://hlinksldjump`) to that KPI's RCA drill-down slide |
| RCA drill-down (one per red KPI) | title / gap panel | Qn actual vs target, gap, baseline → last-year target, all verbatim/derived |
| | support register | the sheet's "Interdependencies across all function" rows: year, function, action verbatim, progress %, status; sorted worst progress first |
| | cards | number of support actions, supporting functions, not-started count and share |
| | synthesis line | counts plus the stated diagnostic rule; no root cause is asserted, and the line says RCA is not captured in the tracker |
| | back button | click action to the cockpit slide |

The tracker records interdependencies per function, not per KPI, so every drill-down shows the same support register.

### Month-on-month delta elements (added 07 Sep 2026, `--previous <older tracker>`)

| Slide | Element | Source / derivation |
|---|---|---|
| Critical VCP commitments | delta chips under the chart | one chip per critical KPI: change in the reporting-quarter **actual** against the previous tracker. "No change", "New this cycle" or "No previous value" when the comparison is not possible. Tone: green when the KPI moved the good way (direction respects `--lower-better`), red the other way, grey when flat |
| Enabling KPIs | extra column "Δ Qn actual vs previous" | same per-KPI comparison, inserted before the signal column |
| Themes and actions | extra column "Δ progress vs previous" | key-theme progress against the previous tracker, in points, with the previous value quoted |
| Interdependencies | status chips above the register | count of dependency actions per status, previous → current with the difference. Green when Completed rises, red when Not started rises |
| Which supporting function progressed *(new slide)* | grouped bar chart + read-out | average of the progress values recorded for each supporting function's actions, previous versus current, one bar pair per function, with a per-function sentence |

KPIs are matched across the critical/enabling sections (an older tracker may file the same metric under a
differently-named section), and the comparison workbook is parsed leniently so older section headings still yield
KPI blocks. Everything unmatched is reported as new or removed rather than silently dropped.

Every delta cell and chip carries a direction mark and a red/green tone: ▲ green when the value moved the good way
(direction respects `--lower-better`), ▼ red the other way, – grey for no change, ✦ blue for a metric that is new this
cycle. The numeric change is always shown next to the mark.

### KPI cockpit merged into the linkage slide (07 Sep 2026)
The separate cockpit slide was dropped. The linkage slide now carries the full KPI table - metric, section, Qn actual,
delta vs previous, Qn signal, KPI risk and committed action - and its "Linked key theme (column L)" cell holds the
clickable `➔ Open RCA drill-down` link for every red KPI. The RCA slides link back to it.

---

## The standard deck (07 Sep 2026): same 14 slide types for every function

`vcp_deck.py build` on the Monthly VCP report template produces one fixed slide set per function,
written as `<Function>_VCP_Monthly_Review.pptx`:

1. Cover
2. Critical VCP commitments (chart + monthly remarks + month-on-month delta chips)
3. VCP commitment enabling KPIs (table incl. the "delta vs previous" column)
4. Themes and actions (table incl. the "delta progress vs previous" column)
5. Key theme movement vs previous tracker (chart)
6. Interdependencies (register + month-on-month status chips)
7. Which supporting function progressed (chart + per-function read-out)
8. VCP impact linkage
9. RCA and sheet performance
10. KPI cockpit / theme linkage - every critical and enabling KPI; a red KPI's "Linked key theme" cell
    links to its RCA drill-down
11+. One RCA drill-down per red KPI, each linking back to slide 10

Slides 5, 6 (chips) and 7 need `--previous <older tracker>`. The number of RCA drill-downs equals the
number of KPIs below target, and long tables add "(cont.)" slides - those are the only differences
between functions.

Opt-in extras (off by default, they were removed from the standard set on 07 Sep 2026):
- `--with-change-table` : the "What changed since the previous tracker" table slides
- `--with-governance-slides` (with `--governance` and `--decision`) : the decisions slide and the
  previous-meeting action tracker

`python vcp_agent/build_all.py --xlsx <current> --previous <older>` builds every function sheet in one run.

### KPI direction
A KPI whose name suggests a lower actual is better (MCoS, lead time, cost, ...) but which is not declared
with `--lower-better` is shown as "Direction to confirm" in amber instead of being coloured red or green,
and the reason is recorded in the validation summary. Declaring it with `--lower-better` gives it a signal.
