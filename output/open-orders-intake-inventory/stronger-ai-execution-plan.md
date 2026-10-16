# Open Orders Report Automation: Execution Brief for the Stronger AI Review

This file is an **architect's brief**. It was written by one AI session (2026-10-09) so that a later **Stronger AI review** session (the "orchestrator") can take the proof of concept in this folder to a production-ready monthly tool without redesigning it. The orchestrator executes, delegates the mechanical and parallel work to **smaller AI** subagents, verifies, and pushes. The user runs it with the prompt at the very end (Section 11).

---

## 1. Mission

Preparing the "001. Open Orders - Order Intake - Inventory Analysis By MSA" workbook by hand takes an analyst most of a day each month: for every city, open three Snowflake exports, filter, add a size class, sort, shade, copy an inventory month across, repeat. A proof-of-concept Excel macro that does all of it in one click already exists here and its logic has been validated against the real sample. Your mission is to make it something the team trusts and runs every month:

1. Harden the macro so it runs first time on a real PC, on every city, on messy exports.
2. Prove it with test fixtures and a repeatable check, not by eye.
3. Give the team a one-page run guide and a one-page business case.
4. Leave the repo in the state the repo rules file demands.

Deliverables (one or more commits, all to `main`):

| Deliverable | Path | Notes |
|---|---|---|
| Hardened macro | `source/open-orders-intake-inventory/build-open-orders-report.bas` | Same module name, same entry points. Every behaviour change mirrored in the Python reference. |
| Python reference | `source/open-orders-intake-inventory/build_report_reference.py` | Stays the executable spec of the rules; the macro must agree with it. |
| Fixture generator | `source/open-orders-intake-inventory/make_test_fixtures.py` | Writes 3 synthetic cities with edge cases (Section 6, Phase 2). |
| Report checker | `source/open-orders-intake-inventory/check_report.py` | Asserts an output workbook obeys the data contract. Independent of the builder. |
| Excel test protocol | `output/open-orders-intake-inventory/excel-test-protocol.md` | Step list and expected results for the first run on a real PC (only if you cannot run Excel yourself). |
| Run guide | `output/open-orders-intake-inventory/macro-instructions.md` | Update, do not rewrite. Keep Sections 2 and 3 to one screen each. |
| Business case | `output/open-orders-intake-inventory/business-case.md` | One page. Before/after process, time per month, risk reduction, next steps. |
| Map rows | `MAP.md` | One row per new or changed file, same commit. |

---

## 2. Standing rules

1. Read the repo rules file and `MAP.md` first and obey them: push directly to `main`, no pull request, update `MAP.md` in the same commit as any file change, outputs in `output/open-orders-intake-inventory/`, sources in `source/open-orders-intake-inventory/`.
2. **Do not redesign the data contract** in Section 4. It was reverse-engineered from the hand-built Austin and San Antonio tabs and confirmed row for row. Change a rule only when the user answers one of the open questions in Section 5, and then change it in the macro, the Python reference and the instructions together.
3. **Delegate aggressively to a smaller AI.** It costs a small fraction of the orchestrator per token. Everything that is mechanical, parallel or verbose belongs to the smaller AI: writing fixture scripts, drafting documents, checking a table against spec pages, reviewing one section of VBA against a checklist, summarising a long file. You keep integration, judgment calls and the final word. Use **a second AI reviewer** (a mid-size AI) only where a worker needs real code judgment, such as an independent VBA review. Run smaller-AI agents in waves of three to five in one message so they work in parallel.
4. Keep every smaller-AI agent's total context under **100,000 tokens**: a self-contained brief, the exact file paths it may read (never "look around the repo"), the exact output you want and where to write it. Verify a worker's output with a script or a second worker, not by reading it all yourself.
5. **Keep your own context lean.** Never dump the sample workbooks cell by cell (the Austin tab alone is 55 KB of text). Use `build_report_reference.py` and short pandas or openpyxl scripts to answer questions about the data, and ask a smaller AI to summarise anything long.
6. **There is no Excel in a cloud session.** The VBA module has never been executed, only reviewed. Rules of engagement:
   - Every VBA change is mirrored in the Python reference and tested there on the sample and the fixtures.
   - If this session runs on a Windows machine with the AI coding tool installed locally (ask the user at kickoff), run the macro for real through PowerShell COM automation (Section 6, Phase 4). That is the gold standard.
   - Otherwise write `excel-test-protocol.md` so the user can run the acceptance test in ten minutes and paste the Run Log back to you.
7. Keep the macro a single `.bas` module with `ShowOpenOrdersReportConfig` as the only macro in the Macros list (Alt+F8); `BuildOpenOrdersReport` and the other Config button macros take an unused Optional argument so they are reachable from the Config buttons (and `Application.Run`) only. No add-ins, no external references, no `Scripting.Dictionary` (Mac users), no features newer than Excel 2010. The file is stored with Windows (CRLF) line endings and `.gitattributes` marks `*.bas -text` so git never converts them; keep it that way, the VBA editor's Import File needs CRLF.
8. Report faithfully. If a check fails, say so with the output. If something was not run, say it was not run.

---

## 3. What already exists (read in this order, then stop reading)

| Read | Why | How much |
|---|---|---|
| `output/open-orders-intake-inventory/macro-instructions.md` | The rules, the Config settings, the validation results, the open questions. This is the contract in prose. | All of it. |
| `source/open-orders-intake-inventory/build-open-orders-report.bas` | The macro (about 900 lines, heavily commented). | All of it, once. |
| `source/open-orders-intake-inventory/build_report_reference.py` | The executable spec. | All of it. |
| `output/open-orders-intake-inventory/size-class-reference.md` | MODEL_GROUP to size class with published specs and a list of boundary models. | Skim the rules and the "needs a human check" list. |
| `source/open-orders-intake-inventory/sample-data/` | Anonymised Austin inputs and the hand-built target. | Do not open by hand. Run the reference script on the folder and look at its log. |

State of verification when this brief was written:

- Python reference vs hand-built Austin tab: Order Intake block 170/170 rows as a set, size-class order identical, 167/170 rows in identical position (the 3 others are hand-ordering inside excavator classes); inventory block 29/29 rows in order with 3 TOTAL values off by 1 or 2 in the sample itself (anonymisation noise).
- VBA: statically reviewed by a second AI reviewer (no compile errors found; three runtime findings and several edge-case divergences fixed). Never compiled or run.
- The `CTL=` / `MX=` header numbers in the sample (21/2 Austin, 15/15 San Antonio) could not be derived from the inputs. The macro offers three definitions behind `Config!HeaderCountMode`.

---

## 4. The data contract (fixed unless the user says otherwise)

### 4.1 Inputs
One folder per month holding, for each city, files named `Order Intake <City>*.xlsx` (required), `Inventory by Month <City>*.xlsx` (optional), `Open Order <City>*.xlsx` (optional). The city is the text after the report name up to the first underscore. Headers are matched by name, not position:

- Order Intake: `ORDER_DATE, REQUESTED_DATE, ORDER_STATUS, IS_OPEN_ORDER, SUM(F.ORDERED_QUANTITY), PRODUCT_TYPE_ABBR, MODEL_GROUP, MODEL_NUMBER, ORGANIZATION_NAME`
- Inventory by Month: `ORGANIZATION_NAME, ORGANIZATION_CUSTOMER_NUMBER, PRODUCT_TYPE, MODEL_GROUP, MODEL_NUMBER, INVENTORY_DATE, TOTAL`
- Open Order: the Order Intake columns plus `SALES_ORDER_NUMBER, ORACLE_ORDER, ORGANIZATION_CUSTOMER_NUMBER` (it is the open subset of Order Intake)

### 4.2 Output workbook
Sheet `File Format` (legend), then one sheet per city in file-name order, then `Run Log`. Every city sheet:

| Range | Content |
|---|---|
| A1 | "Size Class Sheet", orange FFC000, centred |
| B1:J1 merged | "This Data Comes off the Order Intake Report From Snowflake", black fill, white text, centred |
| L1 | "This Data Comes from Inventory by Month From Snowflake", black fill, white text |
| A2:J2 | `Size Class` + the nine Order Intake headers, yellow FFFF00, centred |
| L2 | "Inventory By Month for <Month> <day><ordinal>" from the snapshot date, yellow |
| M2 / N2 | "CTL=" (yellow) / count (green 92D050) |
| O2 / P2 | "MX=" (yellow) / count (light blue 00B0F0) |
| A3:J(n) | Order Intake rows after filtering, sorted, shaded per size class, dates `mm-dd-yy`, AutoFilter on A2:J(n) |
| Q7:U7 | `PRODUCT_TYPE, MODEL_GROUP, MODEL_NUMBER, INVENTORY_DATE, TOTAL`, green, centred (row is `Config!InventoryHeaderRow`) |
| Q8:U(m) | Inventory rows for the snapshot date, CTL (green) then MX (light blue) then SSL (orange, only if included), centred |

Column widths A..U: 21, 19.57, 23.29, 20, 21.71, 36.29, 23, 18.57, 24.71, 57.57, 8.43, 33.57, 9, 9.43, 9.57, 9, 25.86, 14.71, 16.71, 17.71, 12.

### 4.3 Rules
- Keep `PRODUCT_TYPE_ABBR` in {CTL, MX}; add SSL only when `IncludeSSL` = TRUE.
- Drop `ORDER_STATUS` = CANCELLED when `DropCancelled` = TRUE (default).
- Size class from `MODEL_GROUP` via the SizeClassMap sheet. Loaders by rated operating capacity in lb, `0<2000, 2000<2300, 2300<2600, 2600<3000, 3000<3500, 3500&Over` (upper bound exclusive). Excavators by nominal tonnes = model number / 10, `>00-01, >01-02, >02-03, >03-04, >04-05, >05-06, >06-10, Over10` (upper bound inclusive). Unknown E-models fall back to the tonnes rule; anything else is `UNMAPPED`, shaded red, listed in the Run Log.
- Sort by size class in the order above, then ORDER_DATE descending; ties keep source order.
- Size-class shading: theme Accent 6 with tints 0.80 / 0.60 / 0.40 / -0.25 for the first four loader classes, RGB 00B050 for `3000<3500`, Accent 6 tint -0.50 for `3500&Over`; Text 2 tint 0.90 for `>00-01`; Accent 4 tints 0.60 / 0.40 for `>01-02` / `>02-03`; Accent 1 tint 0.40 for `>03-04`; Accent 4 tints -0.25 / -0.50 for `>04-05` / `>05-06`; RGB 0070C0 for `>06-10`.
- Inventory snapshot = `Config!InventorySnapshotDate`, or the latest INVENTORY_DATE in that city's file when blank. Compare on the date part only.
- Header counts per `HeaderCountMode`: `InventoryUnits` = live SUMIF of TOTAL by type over the block (default); `InventoryRows` = COUNTIF; `OpenOrdersDueInMonth` = Open Order rows per type with REQUESTED_DATE in the snapshot month.
- Run Log: one line per city with rows read, kept, dropped by reason, snapshot date, unmapped groups; first two lines hold the folder, timestamp and options.

---

## 5. Open questions for the user

Ask all of these in **one message at kickoff**, then proceed with the defaults while you wait. Do not block on them.

| # | Question | Default if unanswered |
|---|---|---|
| 1 | What do `CTL=` / `MX=` in N2/P2 mean? The sample's 21/2 and 15/15 match no count on the inputs. | `InventoryUnits` (sum of inventory units by type). |
| 2 | Should CANCELLED intake rows be dropped (Austin tab) or kept (San Antonio tab)? | Drop. |
| 3 | Which inventory month: latest in the file, or the previous complete month? The Austin tab used the second-latest. | Latest; `InventorySnapshotDate` overrides. |
| 4 | Does the full workbook use the Open Order file anywhere beyond N2/P2 (an Open Orders tab, a pivot)? If yes, get a screenshot or the tab. | Not used except in `OpenOrdersDueInMonth`. |
| 5 | Are Skid-Steer Loaders ever wanted? | No. |
| 6 | Who owns the size-class definition, and are the boundary models in `size-class-reference.md` placed correctly (T550, T64/S64, T770, S86-2, E40)? | Keep the table as is, rows marked "verify". |
| 7 | Exact file names as they come out of Snowflake today, and whether they are `.xlsx` or `.csv`. | The convention in Section 4.1. |
| 8 | Is the AI coding tool running on a Windows PC with Excel installed (so the macro can be executed for real), or in the cloud? | Cloud: write the test protocol instead. |
| 9 | Is the "all cities on one tab with a City column" variant wanted now? The user said "not for now". | Build it behind `Config!AllCitiesTab` = FALSE only if time allows; otherwise leave a note. |

---

## 6. Work plan

### Phase 0: Orient (you, 15 minutes)
Read the files in Section 3. Run the reference on the sample and keep the log line in your notes:

```
cd source/open-orders-intake-inventory
python3 build_report_reference.py sample-data /tmp/poc.xlsx --snapshot 2026-11-08
```

Confirm the two log lines say 170 kept and 29 kept. If they do not, stop and find out why before anything else.

### Phase 1: Kickoff message (you, 5 minutes)
Send the user the Section 5 questions in one message. Continue immediately.

### Phase 2: Fixtures and checker (smaller-AI wave 1, three agents in parallel)
- **H-FIX** writes `make_test_fixtures.py`: creates a folder with three synthetic cities. City A: clean, two orgs, CTL+MX+SSL, two inventory months, some CANCELLED rows, one `UNMAPPED` model group ("T999"), an E-model not in the map ("E25" must land in `>02-03` by the tonnes rule). City B: Order Intake only (no Inventory, no Open Order), one data row, a blank row in the middle, headers in a different column order. City C: Inventory file with the Config snapshot date missing, Open Order file present, a 35-character city name with a "/" in it, dates stored as text in one column. Also writes `expected.json` with the counts the checker should see.
- **H-CHECK** writes `check_report.py <output.xlsx> [expected.json]`: opens the workbook with openpyxl and asserts the contract in Section 4.2 and 4.3 (sheet set and order, header cells and fills, size-class order monotone, no CANCELLED when configured, only CTL/MX, inventory block single date, CTL before MX, header count formulas present, Run Log line per city, red fill on UNMAPPED). Exit code 0/1 and a readable report. It must pass on the Python reference's output for the sample before you accept it.
- **H-SPEC** takes `size-class-reference.md`, searches Bobcat's current NA spec pages for every row marked Med or Low or BOUNDARY, and returns a table of rows whose class should change, with the source URL. You then update `SizeClassMap` defaults in both the `.bas` and the `.py`.

Acceptance for this phase: `python3 build_report_reference.py fixtures/ out.xlsx && python3 check_report.py out.xlsx fixtures/expected.json` passes; and it passes on the sample folder too.

### Phase 3: Harden the macro (you, with a second AI reviewer)
Work through this list, mirroring each change in the Python reference and re-running Phase 2's check after each:
1. Apply any open findings from the earlier second-reviewer pass (they are applied already; re-review after your changes with a fresh second AI reviewer, using the prompt in Section 7, Template E).
2. Dates stored as text (City C): convert with `CDate` when `IsDate`, otherwise log the row and leave it.
3. CSV inputs are deliberately not supported (locale-dependent date parsing). If the user's exports are CSV (question 7), add an explicit `Workbooks.OpenText` path with a fixed date format rather than relying on `Workbooks.Open`.
4. One-row and zero-row inputs: `UsedRange.Value` returns a scalar for one cell; the code must not call `UBound` on it.
5. Progress in the status bar (`Application.StatusBar = "Building Austin..."`) and cleared at the end.
6. A "Build report" button on the Config sheet (a Forms button assigned to `BuildOpenOrdersReport`) created by `ShowOpenOrdersReportConfig` if missing.
7. Optional, only if the user said yes to question 9: `Config!AllCitiesTab` = TRUE adds one sheet "All Cities" with column K = City (header "CITY"), same blocks, inventory block omitted.
8. Keep the module importable: `Attribute VB_Name` first line, `Option Explicit`, CRLF line endings, no line longer than 1,000 characters, no statement with more than 20 line continuations or 50 arguments (the three `Array(...)` literals in `EnsureMapSheet` are the ones to watch; add a fourth array rather than growing one).

### Phase 4: Run it in Excel
- **If Windows with Excel is available** (question 8), run headless through COM from PowerShell: open the builder `.xlsm`, set `Config!InputFolder` to the fixtures folder, call `Application.Run "BuildOpenOrdersReport"`, close, then run `check_report.py` on the result. Repeat for `sample-data`. Paste the Run Log into your report. This is the acceptance test; nothing else substitutes for it.
- **Otherwise** have a smaller-AI agent draft `excel-test-protocol.md` from the fixture definitions: numbered steps, the Config values to set, and an expected-results table per fixture (tab names, rows kept, UNMAPPED groups, N2/P2 values, red rows). You edit it to be exact. Ask the user to run it and paste the Run Log back; when it arrives, fix whatever it shows and repeat.

### Phase 5: Documents (smaller-AI wave 2, two agents in parallel)
- **H-DOC-GUIDE** updates `macro-instructions.md`: Section 2 and 3 to one screen each, Section 6 with the fixture results, Section 7 with the user's answers to Section 5. Give it the current file and the answers; forbid it from touching Sections 4 and 5 except where an answer changes a default.
- **H-DOC-CASE** drafts `business-case.md` (one page): the manual process as the user described it, the automated process, a time table with blanks for the user's own numbers (cities x minutes per city per month, review time after), error classes removed (missed cancellations, inconsistent filters between tabs, mis-shaded rows), what it took to build, and three next steps (all-cities tab, scheduled export from Snowflake, Power Query alternative). Plain prose, no marketing tone.
If the user wants the printed one-pager style used elsewhere in this repo, compile the business case with `Specs/exam-style-latex-spec.md`; otherwise markdown is the deliverable.

### Phase 6: QA and ship (you)
Run the checklist in Section 8, add the `MAP.md` rows, commit with a descriptive message, push to `main`, and report: what changed, what was verified and how, what was not run, what the user still owes (answers, the Excel run).

---

## 7. Smaller-AI delegation: rules and prompt templates

Rules for every worker brief:
- State the goal in two sentences, then the exact inputs (file paths it may read, data pasted inline if small), then the exact output (path, format, length), then a "do not" list (no repo browsing, no changes to other files, no questions back; state assumptions in the report instead).
- Ask for a short structured report at the end (what was produced, what was assumed, what could not be done).
- Give a token budget hint: "keep your total reading under 60K tokens".
- Run independent workers in the same message so they execute in parallel.

**Template A, fixture generator (smaller AI):**
> Write `source/open-orders-intake-inventory/make_test_fixtures.py` (Python 3, openpyxl only). It creates `<out>/` with Excel files for three synthetic cities exactly as specified below [paste Phase 2 H-FIX spec]. Column headers must be spelled exactly: [paste Section 4.1]. Dates are real datetimes except where the spec says text. Also write `<out>/expected.json` with, per city: rows_in, rows_kept (after dropping SSL and CANCELLED), unmapped_groups, inventory_rows_at_latest_date, latest_inventory_date. Do not read any other repository file. Report: the counts you wrote to expected.json and any spec ambiguity you resolved.

**Template B, checker (smaller AI):**
> Write `source/open-orders-intake-inventory/check_report.py`. Input: an output workbook and an optional expected.json. Assert the contract below [paste Section 4.2 and 4.3], print PASS/FAIL per rule with the offending cell, exit 1 on any failure. Use openpyxl only; fills are read from `cell.fill.fgColor` (theme+tint or rgb). Do not read any other repository file. Report: the list of rules you check and any rule you could not express.

**Template C, document drafter (smaller AI):**
> Draft `<path>` in plain prose, about <n> words, for a dealer-operations audience. Content: [bullets]. Constraints: no marketing adjectives, no em-dashes, numbers in tables not sentences, headings only if over 500 words. Inputs: [paste or path]. Do not read other repository files. Report: assumptions you made.

**Template D, spec verifier (smaller AI, needs web search):**
> For each row in the table below [paste the Med/Low/BOUNDARY rows], find Bobcat's current North American spec page and return MODEL_GROUP, published rated operating capacity (lb) or nominal tonnes, the URL, and the size class under these rules [paste Section 4.3 rule]. Mark CHANGE where the class differs from the table. Do not guess; write NOT FOUND.

**Template E, VBA reviewer (second AI reviewer):**
> Review `/path/build-open-orders-report.bas` (Excel 2016/365 VBA, never compiled) against `/path/build_report_reference.py` (validated behaviour). Find compile errors, runtime errors and logic differences; for each give severity, line number, the problem and the corrected line. Check in particular: Option Explicit coverage, UDT returns, line continuation limits in `Array(...)`, Dir re-entrancy, UsedRange scalar case, Sort object usage, AutoFilter toggling, sheet-name limits, snapshot date comparison, SUMIF ranges. Report only; do not edit files.

---

## 8. Acceptance checklist

- [ ] `python3 build_report_reference.py sample-data out.xlsx --snapshot 2026-11-08` logs 170 kept / 29 kept and `check_report.py` passes on the result.
- [ ] `make_test_fixtures.py` + reference + `check_report.py` pass on the fixtures, including the UNMAPPED group, the one-row city, the missing-inventory city and the odd sheet name.
- [ ] The `.bas` imports into a fresh `.xlsm` without a compile error (Debug > Compile VBAProject) and `BuildOpenOrdersReport` runs on `sample-data` and on the fixtures with the Run Log matching the Python log line for line (counts may differ only where the protocol says so). Done by you via COM, or by the user via the protocol.
- [ ] Opening the macro's Austin tab next to the hand-built Austin tab shows the same rows, order, shading and widths.
- [ ] `macro-instructions.md` Sections 2 and 3 fit on one screen each and match the current Config keys.
- [ ] `business-case.md` exists, one page, no invented numbers (blanks for the user's figures).
- [ ] Every new or changed file has a `MAP.md` row. `git status` is clean after the push to `main`.
- [ ] Your final message lists what was not run and what the user still has to do.

---

## 9. Budget expectations

| Work | Model | Tokens (rough) | Cost at API rates |
|---|---|---|---|
| 5 to 7 smaller-AI workers (fixtures, checker, spec check, two documents, protocol) | Smaller AI | 50K to 80K each | under $0.50 total |
| 1 to 2 VBA reviews | Second AI reviewer | 150K each | about $1 |
| Orchestration, integration, macro edits | Stronger AI review | 400K to 700K | $5 to $10 |

Well within one working session. If Phase 3 item 7 (all-cities tab) is in scope, add one more smaller-AI worker to extend the fixtures and checker.

---

## 10. Things that went wrong once already

- Reading the sample workbooks cell by cell cost 90 KB of context for no gain. Use scripts.
- The header counts (N2/P2) were chased for a while: no combination of the three inputs produces 21/2. Do not repeat that search; ask the user (question 1).
- The analyst's hand ordering inside a size class does not follow any single key. The macro's rule (ORDER_DATE descending) is a decision, not a discovery; keep it unless the user objects.
- The anonymised inventory TOTAL values differ from the hand-built tab by 1 or 2 in three rows. That is noise in the sample, not a bug to fix.

---

## 11. The prompt to paste into the Stronger AI review

> Read `output/open-orders-intake-inventory/stronger-ai-execution-plan.md` in full and execute it as the orchestrator. Follow its standing rules exactly: push to `main`, keep `MAP.md` current, do not redesign the data contract. Send me the Section 5 questions in your first message, then keep going with the defaults until I answer. Delegate every mechanical, parallel or document-writing task to smaller-AI subagents, in parallel waves, with self-contained briefs under 100K tokens each; use a second AI reviewer for the independent VBA review. I am running this from [a Windows PC with Excel installed / the cloud] — pick the matching path in Phase 4. Finish with the acceptance checklist in Section 8 filled in honestly, including what was not run.
