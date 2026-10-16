# Open Orders Report v2: Implementation Plan for the Stronger AI Review

This file is an **architect's brief**. It was written by one AI session (2026-10-12) from the owner's notes of a call with their manager about the proof-of-concept macro in this folder. A later **Stronger AI review** session (the "implementer") executes it: it turns macro version 1.3 into version 2.0, the matching Python reference, fixtures, checker and documents, and pushes everything to `main`. The owner runs it with the prompt at the very end (Section 12).

The manager's feedback changes the inputs (all cities in one set of three files, a separate size-class workbook), the output (all cities on one tab, inventory on its own tab, no colours, a professional look) and three rules (cancelled orders and skid-steer loaders are kept but hidden, rows are grouped by size class then order status then date). It also answered most of the questionnaire. Everything else about the tool, above all the Config control panel and the setup-check-first behaviour, stays as it is.

---

## 1. Mission and deliverables

Make the report the manager described, keep the tool as trustworthy as version 1.3 (every rule mirrored in the Python reference, proven on fixtures and on the sample, run in LibreOffice), and leave the repo in the state the repo rules file demands.

| Deliverable | Path | Notes |
|---|---|---|
| Macro v2.0 | `source/open-orders-intake-inventory/build-open-orders-report.bas` | Same module name, same single entry point in Alt+F8, one more Config button. |
| Python reference v2.0 | `source/open-orders-intake-inventory/build_report_reference.py` | Stays the executable spec; the macro must agree with it. |
| Size-class starter | `source/open-orders-intake-inventory/make_size_class_master.py` writes `output/open-orders-intake-inventory/size-class-master.xlsx` | The first version of the master list (Section 5.1.3). |
| Fixtures v2 | `source/open-orders-intake-inventory/make_test_fixtures.py` and `test-fixtures/` | All-cities layout plus one per-city file, a fixture size-class file, `expected.json` for two runs (Section 7.1). |
| Checker v2 | `source/open-orders-intake-inventory/check_report.py` | The v2 contract (Section 7.2). |
| Other scripts | `check_bas.py`, `run_macro_libreoffice.py`, `compare_with_handbuilt.py` | Updated for v2 (Section 7.3). |
| Sample output | `output/open-orders-intake-inventory/001. Open Orders - Order Intake - Inventory Analysis By MSA_POC.xlsx` | Rebuilt by the v2 reference from `sample-data`, so the repo shows the v2 look. |
| Run guide | `output/open-orders-intake-inventory/macro-instructions.md` | Updated, not rewritten (Section 8). |
| Excel test protocol | `output/open-orders-intake-inventory/excel-test-protocol.md` | Rewritten for the v2 runs. |
| Business case | `output/open-orders-intake-inventory/business-case.md` | Light touch: yearly run, v2 scope. |
| Manual PDF, Rev. 4 | `source/.../open-orders-macro-manual.tex` and `output/.../open-orders-macro-manual.pdf` | One page, same format (Section 8.2). |
| Questions PDF, Rev. 2 | `source/.../open-orders-report-questions.tex` and `output/.../open-orders-report-questions.pdf` | One page: answers recorded, open questions left (Section 8.3). |
| Map rows | `MAP.md` | One row per new, changed, renamed or deleted file, same commit. |

---

## 2. Standing rules

1. Read the repo rules file and `MAP.md` first and obey them: push directly to `main`, no pull request, update `MAP.md` in the same commit as any file change, sources in `source/open-orders-intake-inventory/`, outputs in `output/open-orders-intake-inventory/`, keep the macro manual current and one page whenever the macro changes anything a user sees.
2. **Do not spawn agents of the most expensive tier.** Doing everything yourself is fine. If you delegate, use smaller-AI subagents for mechanical, parallel or verbose work (fixture tables, document drafts, the HP research in Section 5.1.3) and one second AI reviewer for an independent review of the finished VBA. Every brief is self-contained: the goal, the exact file paths it may read, the exact output, a "do not" list; never "look around the repo". Verify a worker's output with a script, not by reading it all.
3. **Build to the contract in Section 5.** It records the manager's decisions. Where a detail is left to you, pick the simplest option, write it down in `macro-instructions.md`, and keep the macro, the reference and the checker in agreement. Change a contract rule only when the owner tells you to.
4. **There is no Excel in a cloud session.** Every VBA change is mirrored in the Python reference and tested there first; then the macro itself is run in headless LibreOffice through `run_macro_libreoffice.py` and compared with the reference's output (`check_report.py --compare`). Features LibreOffice lacks are substituted in the test copy, listed in the script's table, and called out as "reviewed, not run" in the documents. The real acceptance test is the owner's Excel run from `excel-test-protocol.md`.
5. Keep the macro a single `.bas` module, Excel 2010 compatible, no add-ins, no external references, no `Scripting.Dictionary`, `ShowOpenOrdersReportConfig` the only macro in Alt+F8, the button macros hidden by their unused Optional argument, CRLF line endings, ASCII only, and `check_bas.py` passing (it enforces the line, continuation and `Array(...)` limits that bit this project before).
6. Keep the Config control panel's design: a settings table with checked Value cells and drop-downs, buttons on one row, the Last run box, the help box, protection without a password, the setup check that stops before anything is created. New settings join the table in the same style; nothing is silently replaced.
7. Keep `build_report_reference.py` usable as the no-macro fallback (same inputs, same options, same output).
8. Report faithfully. If a check fails, say so with the output. If something was not run, say it was not run.

---

## 3. Read first, then stop reading

| Read | Why | How much |
|---|---|---|
| `output/open-orders-intake-inventory/macro-instructions.md` | The v1.3 rules, Config, validation history. | All of it. |
| `source/open-orders-intake-inventory/build-open-orders-report.bas` | The macro you are changing (about 1,900 lines, commented). | All of it, once. |
| `source/open-orders-intake-inventory/build_report_reference.py` | The executable spec you are changing. | All of it. |
| `source/open-orders-intake-inventory/make_test_fixtures.py` and `check_report.py` | The fixture tables and the rule ids you are replacing. | Skim the tables and the rule list in the docstring. |
| `source/open-orders-intake-inventory/run_macro_libreoffice.py` | How the macro is executed without Excel and what is substituted. | The docstring and `test_copy`. |
| `output/open-orders-intake-inventory/size-class-reference.md` | Published ROC and tonnes per model, the boundary models to flag. | All of it (input for the starter file). |
| `source/open-orders-intake-inventory/open-orders-macro-manual.tex` and `open-orders-report-questions.tex` | The two PDFs you rebuild. | All of both (short). |
| `Specs/exam-style-latex-spec.md` | The one-page format both PDFs follow. | Skim. |
| `source/open-orders-intake-inventory/sample-data/` | Austin inputs and the hand-built target. | Do not open by hand; use the scripts. |

Baseline before you change anything (from `source/open-orders-intake-inventory`):

```
python3 build_report_reference.py sample-data /tmp/v1-sample.xlsx --snapshot 2026-11-08
python3 check_report.py /tmp/v1-sample.xlsx
python3 build_report_reference.py test-fixtures /tmp/v1-fix.xlsx && python3 check_report.py /tmp/v1-fix.xlsx --expected test-fixtures/expected.json
python3 check_bas.py
```

All four must pass (170 kept / 29 kept on the sample; 61 checks on the fixtures). If not, stop and find out why.

Facts about the sample you will need: Order Intake Austin has 278 rows (CTL 160, SSL 68, MX 50), with ORDER_STATUS values ORDERED 104, Not yet processed 59, INVOICED 52, CANCELLED 44, Completely processed 19. Inventory by Month Austin has 299 rows over nine INVENTORY_DATE values (the 8th of each month in the anonymised sample; the 1st in real exports), column F; for November 2026 it holds 18 CTL rows (57 units), 11 MX rows (43 units), 7 SSL rows (7 units). The Open Order file holds the 167 open rows of the intake.

---

## 4. Decision log: what the feedback changes

| # | Feedback (call notes, 2026-10-12) | Decision | Where it lands |
|---|---|---|---|
| 1 | The report will probably run once a year, not monthly. | Nothing in the code depends on the cadence. Documents say "each run" and "yearly" instead of "monthly". | Manual, instructions, business case. |
| 2 | All cities come in one file, still the same three files, with an extra city column whose name is not known yet. | Files are `Order Intake*.xlsx`, `Inventory by Month*.xlsx`, `Open Order*.xlsx` (plural accepted). The city is read from a city column found by `CityColumn` or auto-detected; a file without one gets its city from the file name as in v1. Every file of a kind is read and merged. | 5.1.1, 5.1.2. |
| 3 | Colours and most formatting only helped the manager keep track while pasting. Remove the colours; free rein to make it nicer; keep the structure familiar; professional, minimal, subtle. | No size-class shading, no yellow/black/orange banners, no theme colours. One small palette, bold header rows, thin rules, frozen headers, hidden gridlines. Same tab order (legend, data, Run Log), same row 1 and row 2 roles, same column order. | 5.2, 5.5. |
| 4 | The Config sheet's design is right; keep that philosophy. | The control panel stays. New settings join the table; one new button; the Last run and help boxes are updated. | 5.4. |
| 5 | Keep cancelled orders in the report but filtered out automatically; the data may be needed later. | Cancelled rows are written and hidden by an AutoFilter criterion on ORDER_STATUS. `ShowCancelled` (default FALSE) replaces `DropCancelled`. | 5.3.1. |
| 6 | Include skid-steer loaders, defaulted to filtered off. | SSL rows are written on both data tabs and hidden by the AutoFilter on the product type column. `ShowSSL` (default FALSE) replaces `IncludeSSL`. | 5.3.1. |
| 7 | Group by size class, then order status, then date. | Sort keys: size class (legend order), order status (Section 5.3.3 order), ORDER_DATE newest first, city, export order. | 5.3.3. |
| 8 | Q10: INVENTORY_DATE (column F) is a dump on the first of the month and names the month ("8/1/26" is "August"). Not a focus. | The inventory is handled by month, not by exact date: `InventoryMonth` (any day in the month; blank = latest month in the file) picks the month the summary shows and the default filter; a `Month` column ("August 2026") is added to the inventory tab. | 5.3.4. |
| 9 | A separate size-class Excel sheet is coming, with HP, Rated Operator Capacity, Name of machine, Size class. Generate a starter from the rules and the input data. It is a separate Config input (may live on SharePoint; a local copy also works). | `SizeClassFile` setting with a "Choose size class file..." button; the file is the master list; the `SizeClassMap` sheet is retired. A generator script writes the starter workbook. | 5.1.3, 5.4. |
| 10 | The green and blue data at the side fed the CTL= and MX= figures: they are counts of the totals from Inventory by Month (could not be reconciled only because the sample was anonymised). | CTL= and MX= become a "Units by city" summary on the inventory tab: live SUMIFS of TOTAL per city and product type for the chosen month. `HeaderCountMode` is retired; the Open Order file no longer feeds anything. | 5.2.3. |
| 11 | Output: all cities on one tab (the old analysis), and the Inventory by Month numbers on a separate tab filterable by city. | Tabs `File Format`, `Order Intake`, `Inventory by Month`, `Run Log`. | 5.2. |
| 12 | Q4: Austin may simply have had no cancelled orders that month. | Consistent with 5; the per-tab difference in the sample needs no rule. | Instructions Section 7. |
| 13 | Q6: the master list rarely updates (not in two years) but establishing one is good practice. | The size-class file is the master; UNMAPPED and "placed by the tonnes rule" lines in the Run Log say what to add to it. | 5.1.3, 5.3.2. |
| 14 | Q7: not sure about the boundary models; flag them in the master list. | The starter's Notes column flags T550, T64, S64, T770, S86, E40 and the other boundary rows from `size-class-reference.md`. | 5.1.3. |
| 15 | Q8: answered broadly, do what makes sense for the actual values. | Status order and date direction as in 5.3.3, shown on the File Format tab so the manager can object. | 5.3.3, 8.3. |
| 16 | Q9: nothing is added by hand after a build. | No extra columns, notes or totals on the Order Intake tab; the only computed cells are the summary formulas. | 5.2. |
| 17 | Update the .tex files and regenerate the PDFs of the questions document and the one-page manual, same formats, questions document down to one page. | Phase 5. | 8.2, 8.3. |

---

## 5. The v2 data contract

### 5.1 Inputs

#### 5.1.1 Files
One folder holding the Snowflake exports for all cities. A file's kind is the start of its name, case-insensitive: `Order Intake` (required), `Inventory by Month` (optional), `Open Order` or `Open Orders` (optional). Anything after the kind is free text (`Order Intake 2026.xlsx`, `Order Intake.xlsx`, `Order Intake Austin_2026-09.xlsx` all work). `.xlsx`, `.xlsm`, `.xls` are read; CSV is not (as in v1). Files starting with `001`, containing `Analysis By MSA`, or starting with `~$` are ignored. Every file of a kind is read and its rows merged; the Run Log lists each file with its row count, and the setup check warns when a kind has more than one file ("2 Order Intake files found, all are read: ...") so a stray old export is noticed.

Headers are matched by name, not position, as in v1:

- Order Intake: `ORDER_DATE, REQUESTED_DATE, ORDER_STATUS, IS_OPEN_ORDER, SUM(F.ORDERED_QUANTITY), PRODUCT_TYPE_ABBR, MODEL_GROUP, MODEL_NUMBER, ORGANIZATION_NAME`, plus the city column.
- Inventory by Month: `ORGANIZATION_NAME, ORGANIZATION_CUSTOMER_NUMBER, PRODUCT_TYPE, MODEL_GROUP, MODEL_NUMBER, INVENTORY_DATE, TOTAL`, plus the city column.
- Open Order: the Order Intake columns plus `SALES_ORDER_NUMBER, ORACLE_ORDER, ORGANIZATION_CUSTOMER_NUMBER`. Read only to log its row count; it feeds nothing in v2 (open question 3 in Section 10).

#### 5.1.2 The city of a row
1. If `Config!CityColumn` is set, that header (trimmed, case-insensitive) must exist in every export that is read, else the setup check reports a problem naming the headers it found.
2. If `CityColumn` is blank, the column is auto-detected per file: the first header, in this order, that equals `CITY`, equals `MSA`, contains `CITY`, contains `MSA`, contains `METRO`, contains `MARKET` (ignoring case, spaces and underscores). The setup check and the Run Log say which column was used for each file ("city column MSA_NAME, 3 cities").
3. If no city column is found, the city is taken from the file name as in v1 (the text after the kind up to the first underscore, trimmed). If that is empty too, the setup check reports a problem: "no city column found in Order Intake.xlsx (headers: ...). Type the header name into CityColumn."
4. City values are trimmed, runs of spaces collapsed, compared case-insensitively; the first spelling seen wins. Blank city cells become `(no city)` and are counted in the Run Log.

The setup check reads the files (not only their names) so it can list the cities before a build; expect it to take a few seconds on big exports.

#### 5.1.3 The size-class file (the master list)
`Config!SizeClassFile` is the full path of an Excel workbook, local or a synced SharePoint/OneDrive path; required (a problem if blank or missing). Its first worksheet is read with the v1 `ReadSource` rule (the header row is the first row with at least three non-blank cells). Columns are matched by name with these aliases, ignoring case, spaces and underscores:

| Column | Accepted headers | Used for |
|---|---|---|
| HP | `HP`, `Horsepower` | Information only. |
| Rated operating capacity | `Rated Operating Capacity`, `Rated Operator Capacity`, `ROC`, `Operating Capacity` | Information and a consistency warning (below). |
| Name of machine | `Name of machine`, `Machine`, `Model`, `MODEL_GROUP` | The lookup key. |
| Size class | `Size Class`, `SIZE_CLASS`, `Class` | The class written to column A. |

Any other column is ignored. A row is usable when the name is non-blank and the class is one of the 13 classes or `Over10`; other rows are skipped with a warning naming the row. A name listed twice is a warning; the first row wins. Lookup: the data's `MODEL_GROUP`, trimmed and upper-cased, is matched against the name exactly, then against the name's first word (so `T650 Compact Track Loader` still matches `T650`). Consistency warning, loaders only (names starting with T or S, numeric ROC): when the ROC puts the model in a different class than the file says under the loader rule, warn "row 12 T64: ROC 2300 is class 2300<2600 by the rule, file says 2000<2300". The file's class is still used.

The starter (`make_size_class_master.py` writes `output/open-orders-intake-inventory/size-class-master.xlsx`): one sheet named `Size Class`, header row 1 with exactly `HP`, `Rated Operating Capacity`, `Name of machine`, `Size Class`, `Notes`; one row per model in v1's `SizeClassMap` (83 rows: loaders, skid-steers, excavators, from the `EnsureMapSheet` arrays and `size-class-reference.md`). Values: Name = the MODEL_GROUP as Snowflake emits it; ROC in lb for loaders (the reference table's figure); for excavators the operating weight in lb (reference table kg times 2.2046, rounded) with a note that the class follows nominal tonnes = model number / 10; HP from published Bobcat specs where you can find it with reasonable confidence (bounded effort: one smaller-AI research pass or an hour; blanks are acceptable, the owner will replace the list); Size Class as in v1. Notes: `sample` / `spec <lb>` / `rule <t>` provenance as in v1, plus `CHECK: on a class boundary` for T550, T64, S64, T770, S86, E40 and every row the reference table marks BOUNDARY or CONFLICT. Style the sheet like the report (Section 5.5), freeze the header, AutoFilter on. The fixture generator writes its own small size-class file (Section 7.1); it does not import the starter.

The builder's `SizeClassMap` sheet is retired: `ShowOpenOrdersReportConfig` no longer creates it, a build never reads it, and if one exists the setup check warns once "SizeClassMap is no longer used; the master list is SizeClassFile. You can delete the sheet." Remove `EnsureMapSheet`, `CheckMapSheet`, `ProtectMapSheet` and the three model arrays.

### 5.2 Output workbook
File name from `OutputFileName` (default unchanged). Tabs in this order: `File Format`, `Order Intake`, `Inventory by Month`, `Run Log`. No per-city tabs, so `SanitizeSheetName` and `UniqueSheetName` go.

#### 5.2.1 File Format (the legend, now text)
Gridlines off. A1 title "File Format"; D1 note "How the report is laid out and which rules built it". Then four small tables, each with a bold caption row and a header row in the header style:

1. **Size classes**: Class, Family, Rule. The 13 classes in order, then `Over10`, then `UNMAPPED` ("model not in the size class file; add it there and build again"). Loader rule text: "rated operating capacity in lb, lower bound inclusive, upper bound exclusive"; excavator rule: "nominal tonnes = model number / 10, upper bound inclusive".
2. **Order status order**: the Section 5.3.3 list with rank.
3. **Default filters**: tab, column, what is shown, the Config setting, and the sentence "To see hidden rows, clear the filter on that column."
4. **This build**: built when, builder version, input folder, each file read with its rows and city column, size class file, inventory month shown, cities.

Also a plain single-column list **Months in this build** (chronological, one cell per month, e.g. `August 2026`) that the inventory tab's month drop-down points at.

#### 5.2.2 Order Intake (all cities)
Gridlines off, freeze panes at A3, AutoFilter on A2:K(n+2).

| Cell or range | Content |
|---|---|
| A1 | "Order Intake", title style |
| D1 | note: "Source: Snowflake Order Intake export, <n> files, built <date>. Cancelled orders and skid-steer loaders are hidden by the filters on ORDER_STATUS and PRODUCT_TYPE_ABBR; clear a filter to see them." (adapted to the settings) |
| A2:K2 | headers: `Size Class, City, ORDER_DATE, REQUESTED_DATE, ORDER_STATUS, IS_OPEN_ORDER, SUM(F.ORDERED_QUANTITY), PRODUCT_TYPE_ABBR, MODEL_GROUP, MODEL_NUMBER, ORGANIZATION_NAME` |
| A3:K(n+2) | every kept row (Section 5.3.1), sorted (Section 5.3.3); dates `m/d/yyyy`; PRODUCT_TYPE_ABBR written trimmed and upper-cased; other cells as exported, text protected as in v1 (`SafeCellText`) |

Size Class stays in column A so the sheet reads as before; City is the one new column, next to it. Column widths: v1's widths for the export columns, City 18.

#### 5.2.3 Inventory by Month (all cities, filter by city)
Gridlines off. Layout, with `n` = number of cities in the build (union of the cities seen in Order Intake and Inventory files, alphabetical):

| Cell or range | Content |
|---|---|
| A1 | "Inventory by Month", title style |
| D1 | note: "Source: Snowflake Inventory by Month export, <n> files, built <date>. The table below shows <Month>; change the month cell to see another. Rows are filtered to that month and to CTL and MX; clear a filter to see the rest." |
| A3 | "Units in", bold; B3 = the month cell: the month text (e.g. `August 2026`), input tint (the Config yellow `FFFFCC`), data validation list pointing at `'File Format'` Months in this build |
| A4:E4 | headers `City, CTL, MX, SSL, Total` |
| A5:E(4+n) | one row per city: `=SUMIFS(TOTAL range, City range, $A5, PRODUCT_TYPE range, "Compact Track Loaders", Month range, $B$3)` and the same for `Mini Excavators` and `Skid-Steer Loaders`; Total `=SUM(B5:D5)` |
| A(5+n):E(5+n) | "Total" row, bold, `=SUM()` down each column |
| row h = 7+n | detail headers `City, PRODUCT_TYPE, MODEL_GROUP, MODEL_NUMBER, INVENTORY_DATE, Month, TOTAL`; AutoFilter on A(h):G(last); freeze panes below it |
| A(h+1):G(last) | every inventory row for CTL, MX and SSL (other product types dropped and counted), sorted by city, product type (CTL, MX, SSL), INVENTORY_DATE newest first, export order; INVENTORY_DATE `m/d/yyyy`; Month = "<Month name> <yyyy>" from INVENTORY_DATE |

The summary ranges cover the whole detail block, so the figures do not change when the detail is filtered; say so in the note. ORGANIZATION_NAME and ORGANIZATION_CUSTOMER_NUMBER stay out, as in v1 (open question 5). SSL columns are written even when hidden by the filter; the summary shows SSL units regardless. The `CTL=` and `MX=` figures of the old report are the CTL and MX columns of this table (units, the sum of TOTAL); open question 2 asks the manager to confirm units rather than line counts on the first real run.

#### 5.2.4 Run Log
As in v1 (A1 "Built ... from ... in ... s", A2 the options line, then one line per item from row 4), gridlines off, column A wide. Lines, in this order: one per file read (name, rows, city column or "city from file name"), the size class file (path, usable rows, warnings), then per city in alphabetical order: the Order Intake line, the Inventory line; then a line for each Open Order file ("not used by this report"), then the text-date lines and warnings as in v1. Exact wording is yours; `expected.json` pins it. Counts must add up: rows read = written + dropped (other product types) + blank rows ignored.

### 5.3 Rules

#### 5.3.1 Keep, hide, drop
- Order Intake: keep every row whose `PRODUCT_TYPE_ABBR` (trimmed, upper-cased) is `CTL`, `MX` or `SSL`; drop and count the rest by type. Nothing is dropped for its status.
- Default filters on the Order Intake tab: `ORDER_STATUS` not equal to `CANCELLED` (case-insensitive) unless `ShowCancelled`; `PRODUCT_TYPE_ABBR` in {CTL, MX} plus SSL when `ShowSSL`. Hidden rows are exactly the rows those criteria exclude. In VBA: `rng.AutoFilter Field:=5, Criteria1:="<>CANCELLED"` and `rng.AutoFilter Field:=8, Criteria1:=Array("CTL", "MX"), Operator:=xlFilterValues`. In openpyxl: `auto_filter.add_filter_column` for the value list, a `FilterColumn` with a `CustomFilter(operator="notEqual", val="CANCELLED")` for the status, and `row_dimensions[r].hidden = True` for every excluded row, because Excel shows the hidden flags as stored and does not re-evaluate a filter on open.
- Inventory: keep rows whose `PRODUCT_TYPE` is `Compact Track Loaders`, `Mini Excavators` or `Skid-Steer Loaders` (trimmed, case-insensitive); drop and count the rest. Default filters: `Month` equals the chosen month; `PRODUCT_TYPE` in the two (or three) visible types.
- Blank rows are ignored and counted, as in v1. Dates stored as text are converted exactly as in v1 (`TextToDate`, the Run Log lines, the non-ISO warning); keep that code.

#### 5.3.2 Size class
From `MODEL_GROUP` via the size-class file (Section 5.1.3). A model not in the file: if its `PRODUCT_TYPE_ABBR` is `MX` and the name is `E` plus digits, the tonnes rule applies (number / 10, upper bound inclusive, `Over10` above 10) and the Run Log says "placed by the tonnes rule (not in the size class file): E25"; otherwise `UNMAPPED`, listed in the Run Log and the Last run box, and the Size Class cell gets the flag tint (the only coloured cell in the data). `UNMAPPED` sorts last, `Over10` after `>06-10`.

#### 5.3.3 Sort order (Order Intake)
1. Size class rank: the 13 classes in legend order, then `Over10`, then `UNMAPPED`.
2. Order status rank, by the trimmed upper-cased status: `ORDERED` 1, `NOT YET PROCESSED` 2, `INVOICED` 3, `COMPLETELY PROCESSED` 4, any other status 5, `CANCELLED` 6. Open orders first because the report is about open orders; cancelled last because they are hidden by default. The list is one constant in the module and in the reference, printed on the File Format tab.
3. `ORDER_DATE` newest first, date part only; rows without a date after the dated rows of their group (the hand-built Austin tab matched newest-first in 167 of 170 rows, so this stays).
4. City, alphabetical.
5. Export order (file order, then row order): the sort is stable, as v1's `StableSortIndex` is.

A single Double key still works: class rank, status rank, (3000000 minus the day serial), city rank (its index in the alphabetical city list) packed in that order of significance; or sort in passes from the least significant key up. Either way, mirror it exactly in the reference.

#### 5.3.4 Inventory month
`Config!InventoryMonth` (renamed from `InventorySnapshotDate`): blank = the latest month present across the Inventory files; a date = that date's month, whatever its day. The month text is `<English month> <yyyy>`. It fills B3 of the inventory tab, the default filter on `Month`, the File Format "This build" table and the Run Log options line (`InventoryMonth=latest in file` or `=August 2026`). A build with no Inventory file still writes the inventory tab with the summary (all zeros) and an empty detail block, and logs it.

#### 5.3.5 Run Log and Last run box
As in v1, plus the new lines above. `RunLogLineToCheck` also catches "placed by the tonnes rule", "not in the size class file", "dropped", "(no city)" and "all are read".

### 5.4 Config v2.0 and the upgrade

Nine settings again, rows 2 to 10 of the Config sheet, so `BUTTON_ROW` 12, `STATUS_ROW` 14 and `HELP_ROW` 21 keep their places:

| Key | Default | Check in the cell | Meaning |
|---|---|---|---|
| `InputFolder` | blank | hint only | Folder with the exports (all cities). Blank = folder picker at build. |
| `OutputFolder` | blank | hint only | Blank = InputFolder. |
| `OutputFileName` | `001. Open Orders - Order Intake - Inventory Analysis By MSA.xlsx` | hint only | Unchanged. |
| `SizeClassFile` | blank | hint only | Full path of the size-class workbook (master list). Required. The button "Choose size class file..." fills it. |
| `CityColumn` | blank | hint only | Header of the city column in the exports. Blank = found automatically (CITY, MSA, METRO, MARKET). |
| `InventoryMonth` | blank | date | Any day in the month to show, e.g. 8/1/2026. Blank = latest month in the Inventory file. |
| `ShowCancelled` | FALSE | TRUE/FALSE list | FALSE hides cancelled orders behind the filter; they stay in the file. |
| `ShowSSL` | FALSE | TRUE/FALSE list | FALSE hides skid-steer loaders behind the filters on both tabs; they stay in the file. |
| `OpenAfterBuild` | TRUE | TRUE/FALSE list | Unchanged. |

Buttons on row 12: Build report, Check setup, Choose input folder..., Choose output folder..., Choose size class file... (a file picker, `msoFileDialogFilePicker` = 3, filtered to Excel files, falling back to InputBox as `PickFolder` does). New public `ChooseOpenOrdersSizeClassFile(Optional ByVal notInMacroList As Variant)`; add it to `BUTTON_MACROS` in `check_bas.py`.

Upgrade in place (`ShowOpenOrdersReportConfig` on a v1.x Config): read the old values, rebuild rows 2 to 10 with the v2 keys, carry values over: `ShowCancelled = NOT DropCancelled`, `ShowSSL = IncludeSSL`, `InventoryMonth = InventorySnapshotDate`; drop `HeaderCountMode` and `InventoryHeaderRow` and say so once in the Last run box ("Upgraded to v2.0: HeaderCountMode and InventoryHeaderRow are no longer used; set SizeClassFile."). Help box rows: blocked macros (as now); build stops with a list (as now); "Size Class is UNMAPPED" now says add the model to the size-class file; new row "A city is missing or wrong" pointing at CityColumn; editing the sheet (as now).

Setup check additions: SizeClassFile exists, is readable, has usable rows (problem otherwise) plus its warnings; the city column per file (problem when none and no file-name city); models in the Order Intake data that are not in the size-class file (warning listing them, so the manager can fix the master before building); more than one file of a kind (warning); a leftover SizeClassMap sheet (warning). The Check setup message lists the cities and the inventory months found.

### 5.5 Style sheet (both implementations, exactly)

| Element | Specification |
|---|---|
| Font | The workbook default; do not set a font name. Title 14 bold `262626`; notes 9 `7F7F7F`; headers bold white on `44546A`; body default size. |
| Colours | RGB only, never theme colours (they shift with the theme and bit this project once). Palette: `262626` ink, `7F7F7F` note, `44546A` header fill, `D9D9D9` row rule, `808080` group rule, `FFFFCC` input cell, `FCE4D6` flag (UNMAPPED cell only). Nothing else. |
| Gridlines | Off on all four tabs (`ws.sheet_view.showGridLines = False`; in VBA `ActiveWindow.DisplayGridlines = False` with the sheet active). |
| Data blocks | Thin `D9D9D9` inside-horizontal borders across the block; a thin `808080` top border across the first row of each size-class group (Order Intake) and each city group (Inventory detail); no vertical lines. |
| Headers | Bold white on `44546A`, left-aligned for text columns, right-aligned over numbers and dates, row height 20. |
| Numbers | Quantities `0`, units and totals `#,##0`, dates `m/d/yyyy`. |
| Summary table | Header style; Total row bold with a thin `808080` top border; month cell `FFFFCC` with a thin `808080` border. |
| Freeze panes | Order Intake at A3; Inventory below the detail header; none elsewhere. |
| Widths | Order Intake A 20.29, B 18, then v1's nine (18.86, 22.57, 19.29, 21, 35.57, 22.29, 17.86, 24, 56.86). Inventory A 18, B 25.14, C 14, D 16, E 17, F 16, G 11.29. Keep v1's convention: VBA sets the UI width, openpyxl stores UI width plus 0.71, the checker tolerates 0.3. |

In LibreOffice `ActiveWindow.FreezePanes`, `DisplayGridlines` and AutoFilter criteria may need substitutions in the test copy; find out in Phase 3, substitute only what fails, and list it in the script's table.

---

## 6. Work plan

### Phase 0: Orient (you)
Read Section 3, run the baseline, note the numbers. Write nothing yet.

### Phase 1: Size-class master
1. Write `make_size_class_master.py` and run it; commit the starter workbook. Check it with a five-line openpyxl script: 83 rows, four headers spelled exactly, every class valid, every boundary model flagged.
2. Add the size-class reader to `build_report_reference.py` (aliases, usable rows, duplicates, first-word match, the ROC consistency warning) as a function the checker and the fixtures can import.

Acceptance: the reader loads the starter with zero warnings, and loads the fixture file (Phase 2) with exactly the warnings `expected.json` lists.

### Phase 2: Reference, fixtures, checker (the executable spec)
1. Rewrite `make_test_fixtures.py` to Section 7.1 and regenerate `test-fixtures/` (git rm the v1 per-city fixture files that no longer exist; the script itself never deletes).
2. Rewrite `build_report_reference.py` to Section 5: discovery and merge, city column, size-class file, keep/hide, sort, the four tabs, styles, filters with hidden flags, Run Log. Options: `--size-class-file PATH` (required), `--city-column NAME`, `--inventory-month YYYY-MM-DD`, `--show-cancelled`, `--show-ssl`.
3. Rewrite `check_report.py` to Section 7.2. It must fail on the v1 reference's output and pass on the v2 output.
4. Update `compare_with_handbuilt.py` to Section 7.3 and run it on the sample build.
5. Rebuild the sample output in `output/` with `--inventory-month 2026-11-08`.

Acceptance: `check_report.py` passes on both fixture runs with `--expected`, and on the sample build; `compare_with_handbuilt.py` passes; the sample build shows 278 Order Intake rows written with 108 hidden (68 SSL plus 40 cancelled CTL/MX), 170 visible, and the November summary Austin CTL 57, MX 43, SSL 7.

### Phase 3: Macro v2.0
1. Mirror Phase 2 in the `.bas`: discovery and merge, city column, `LoadSizeClassFile` replacing `LoadMap`, the keep/hide rules, the sort, the four writers, styles, AutoFilter criteria, freeze panes, data validation for the month cell, Run Log, Config v2 and its upgrade, the new button, the setup-check additions, the help box. Remove dead code (per-city tabs, SizeClassMap, HeaderCountMode, InventoryHeaderRow, theme colours). Bump the header comment to Version 2.0.
2. `check_bas.py`: add the new button macro; keep every other check; it must pass.
3. `run_macro_libreoffice.py`: pass `SizeClassFile` (and the other v2 settings) into Config; add substitutions only for what LibreOffice cannot run; normalise the result; run both fixture runs and the sample, then `check_report.py --expected` on the fixtures and `--compare` against the Phase 2 reference builds on all three. Run `--check-only` on a good setup, on a missing SizeClassFile, on six bad settings at once, and on an empty folder, and print the Last run box each time.
4. Independent review of the finished `.bas` against the reference (a second AI reviewer with the Template E brief from `stronger-ai-execution-plan.md` Section 7, updated for v2), then fix what it finds and re-run step 3.

Acceptance: `check_bas.py` passes; LibreOffice builds pass `check_report.py` and are identical to the reference builds in content (`--compare`); the substitution table lists everything the Excel protocol must still cover.

### Phase 4: Documents
`macro-instructions.md`, `excel-test-protocol.md`, `business-case.md` per Section 8.1.

### Phase 5: The two PDFs
Per Sections 8.2 and 8.3. Build with `./build.sh source/open-orders-intake-inventory/open-orders-macro-manual.tex source/open-orders-intake-inventory/open-orders-report-questions.tex`. If `latexmk` is missing in the session, install it first (`sudo apt-get install -y latexmk texlive-latex-extra texlive-pictures texlive-fonts-recommended`; `tcolorbox` and `tikz` come from the extra and pictures packages). Check each PDF is exactly one page with no overfull boxes in the log.

### Phase 6: QA and ship
Run the Section 9 checklist, update every affected `MAP.md` row, commit with a descriptive message, push to `main`, and report: what changed, what was verified and how, what was not run, what the owner still has to do (the Excel run, the real size-class list, the open questions).

---

## 7. Fixtures and checks

### 7.1 Fixtures v2 (`make_test_fixtures.py` writes `test-fixtures/`)
- `Order Intake 2026.xlsx`: one sheet, header row 1, city column headed `MSA` (exercises auto-detect), rows: the 20 Amarillo rows of v1 (keep the edge cases: `Cancelled` in mixed case, ` mx ` padded, T999 UNMAPPED, E25 by the tonnes rule, ties on 9/15), the 7 Dallas-Fort Worth [North] Metroplex rows (ORDER_DATE as text incl. `TBD`), one row with an unknown `PRODUCT_TYPE_ABBR` (`ATT`, dropped and counted), one blank row in the middle, one row with a blank city.
- `Order Intake Brownsville.xlsx`: no city column (city from the file name), shuffled headers, an empty row 2, one data row. Proves merging two files of a kind and the file-name fallback.
- `Inventory by Month 2026.xlsx`: city column headed `CITY` (a different header from the intake file), the v1 Amarillo rows (two months, one date with a time of day) and Dallas rows (two months), one row of an unknown product type, no Brownsville rows.
- `Open Orders 2026.xlsx`: a few rows with a city column; present only to be logged.
- `Size Class.xlsx`: headers spelled `HP`, `Rated Operator Capacity`, `Name of machine`, `Size Class` (the owner's spelling, proves the aliases); rows for the models used plus: a row with class `2000-2300` (invalid, warning), `T650` listed twice (warning), a loader whose ROC disagrees with its class (warning), one name written `T740 Compact Track Loader` (first-word match); T999 and E25 absent.
- Decoys that must be ignored: `001. Open Orders - Order Intake - Inventory Analysis By MSA_2025.xlsx` and the zero-byte `~$Order Intake 2026.xlsx`.
- `expected.json`: two runs, `default` (ShowCancelled FALSE, ShowSSL FALSE, InventoryMonth latest) and `variant` (ShowCancelled TRUE, ShowSSL TRUE, InventoryMonth 2026-08-15, i.e. August 2026). Per run: sheet order; the options line; the Order Intake rows in order as `[class, city, group, status, order date, hidden]`; the summary table (city, CTL, MX, SSL units) and the month cell; the inventory detail rows in order as `[city, type, group, month, total, hidden]`; the File Format "This build" facts worth pinning (files and city columns); the exact Run Log lines. The script asserts hand-computed numbers against what it generates, as v1 did; keep that habit.

### 7.2 Checker v2 (`check_report.py`)
Rule ids to implement (keep the PASS/FAIL per rule, the RESULT line and the exit codes):

| Id | Rule |
|---|---|
| R1 | Sheet names and order: File Format, Order Intake, Inventory by Month, Run Log. |
| R2 | Style on the data tabs: title cells, header fills `44546A` with white bold font, gridlines off, freeze panes where specified, widths within 0.3, no theme colours anywhere, no fill in the data blocks except `FCE4D6` on UNMAPPED Size Class cells. |
| R2F | File Format: the four tables present with the 13 classes in order, the status list, the filter table, the build table, the months list. |
| R3a-b | Order Intake: headers exact; every class known; class rank never decreases. |
| R3c | Inside a class, status rank never decreases; inside a class and status, ORDER_DATE never increases and no date follows a non-date. |
| R3d | PRODUCT_TYPE_ABBR is CTL, MX or SSL, trimmed and upper-cased. |
| R3e | Hidden rows are exactly the CANCELLED rows (when ShowCancelled FALSE) and the SSL rows (when ShowSSL FALSE); the AutoFilter ref and its filter columns match. |
| R3f | Date cells in C and D have a date number format; row rules present; group rule on each first row of a class. |
| R4a | Inventory summary: month cell text is a month in the months list and has the list validation; one row per city, alphabetical; every cell a SUMIFS over the full detail ranges with the right criteria; Total column and row formulas. |
| R4b | Inventory detail: headers exact; sorted by city, type order, date descending; Month equals the month of INVENTORY_DATE; hidden rows exactly the other months and the hidden types. |
| R6 | Run Log: A1 starts "Built ", A2 options line parses (InventoryMonth, ShowCancelled, ShowSSL, CityColumn, SizeClassFile name), every city named. |
| E0-E6 | `--expected`: run selected by the options line; sheet order; Order Intake rows; summary; detail rows; build facts; Run Log lines (unified diff on mismatch). |
| C1-C4 | `--compare`: same sheets; Order Intake A..K and hidden flags equal; inventory summary values (evaluate the formulas with the data, since neither builder caches values the same way) and detail rows and hidden flags equal; Run Log A2 and rows 4.. equal. |

### 7.3 Other scripts
- `compare_with_handbuilt.py`: take the visible Austin rows of the v2 Order Intake tab (not hidden, City = Austin) and compare with the hand-built Austin tab: 170 of 170 rows equal as a set (columns ORDER_DATE..ORGANIZATION_NAME), the Size Class column identical in order; positions may now differ because of the status key, so report the count in the same position but do not fail on it. Inventory: the November rows of the v2 detail block for Austin, 29 of 29 equal in order with at most 3 TOTAL differences. The summary CTL/MX cells are not compared with the hand-built 21/2 (anonymisation).
- `run_macro_libreoffice.py`: v2 settings, substitutions, normalisation; keep `--check-only` and `--set`.
- `check_bas.py`: the new button macro; nothing else changes.

---

## 8. Documents

### 8.1 Markdown
- `macro-instructions.md`: v2.0 at the top; Section 1 file list (new script, starter workbook); Section 2 setup (no SizeClassMap; "Choose size class file..."; where to keep the master list); Section 3 "Each run" instead of "Every month" (file names without city, the city column, the month); Section 4 the v2 Config table and the upgrade conversions; Section 5 the v2 rules (keep/hide, sort, month, summary, size-class file, aliases, warnings); Section 6 validation results of v2 (fixtures, sample, LibreOffice, what was substituted); Section 7 decisions: record the answers from the call and keep only the open questions of Section 10; Section 8 adding a city (nothing to do now: it is a value in the city column), changing the layout; Section 9 troubleshooting rows for the new messages (SizeClassFile missing, no city column, model not in the master, more than one file of a kind, the hidden rows); Section 11 the v2 fallback command line.
- `excel-test-protocol.md`: rewrite for v2: CHECK 1 compile, CHECK 2 Config (nine settings, five buttons, no SizeClassMap), CHECK 3 style (header fills, no colours, hidden rows: the filter buttons on ORDER_STATUS and PRODUCT_TYPE_ABBR show a filter icon, clearing them reveals the cancelled and SSL rows), Run A fixtures default, Run B fixtures variant, Run C sample with `InventoryMonth` 2026-11-08 (summary Austin CTL 57, MX 43, SSL 7; 170 visible rows; side-by-side with the hand-built tab: same rows, same class order, positions differ by the status key), Run D rebuild with the report open, Run F the control panel including the new button, a missing SizeClassFile (problem, nothing built) and the upgrade from a v1.3 Config. Expected Run Log text verbatim from `expected.json`.
- `business-case.md`: "monthly" becomes "each run (the report is expected yearly)"; the process paragraph mentions all cities in one build and the size-class master; the open decisions paragraph shrinks to Section 10.

### 8.2 Manual, Rev. 4 (`open-orders-macro-manual.tex`, one page, same format)
Header `Rev.~4` and the date of the change. Scope: three exports for all cities from one folder plus the size-class workbook; one workbook with File Format, Order Intake (all cities), Inventory by Month (units by city and the detail) and Run Log. Table 1 key settings: `InputFolder`, `SizeClassFile`, `CityColumn`, `InventoryMonth`, `ShowCancelled`, `ShowSSL` (one line each). Section 2 one-time setup: the warning box stays; steps: save as .xlsm, import and compile, run `ShowOpenOrdersReportConfig`, click Choose size class file..., Enable Content. Section 3 "Running the report" (not "Monthly run"): export the three files into one folder, check Table 1, Build report, read the Run Log and the Last run box, add any UNMAPPED model to the size-class file and build again, clear a filter to see cancelled or skid-steer rows. Section 4 road map: step 1 "Confirm the rules" is done (2026-10-12 call) except the items in Section 10; step 2 "Fix the macro" becomes "Build v2" (done when this plan ships); step 3 Test in Excel (the protocol); step 4 "Add more cities" is replaced by "Adopt the master list" (the owner's real size-class list replaces the starter; done when a build reports no UNMAPPED models); step 5 Feed PowerPoint; step 6 Hand over. Keep the table to one page; shorten wording before cutting rows.

### 8.3 Questions document, Rev. 2 (`open-orders-report-questions.tex`, one page)
Keep the title, the exam-style format and the `questions` environment, remove the `\clearpage`, bump to `Rev.~2` with the date. Opening paragraph: the questions were answered on a call on 2026-10-12; this revision records the answers the macro now applies and leaves the few still open. Section 1 "Answered" as a three-column table (Topic, Answer, What the macro does), one row per decision in Section 4 of this plan that came from a question (exports and cities, cancelled orders, skid-steer loaders, size-class master and boundary models, order inside a class, nothing added after a build, inventory month, the CTL= and MX= figures). Section 2 "Still open" as a `questions` table with the answer column, holding the Section 10 questions. Must fit one page; tighten wording, do not drop a question.

---

## 9. Acceptance checklist

- [ ] Baseline (Section 3) passed before any change.
- [ ] `make_size_class_master.py` writes the starter; 83 rows, four headers exact, every class valid, boundary rows flagged; the reader loads it with no warnings.
- [ ] `make_test_fixtures.py` regenerates `test-fixtures/` and its self-check passes; old per-city fixture files are removed from git.
- [ ] `build_report_reference.py` + `check_report.py --expected` pass on both fixture runs; `check_report.py` passes on the sample build; `compare_with_handbuilt.py` passes; the sample numbers in Phase 2's acceptance line hold.
- [ ] `check_bas.py` passes on the v2 `.bas`; CRLF and ASCII intact.
- [ ] `run_macro_libreoffice.py` builds of both fixture runs and the sample pass `check_report.py` and are identical to the reference builds under `--compare`; the four `--check-only` cases behave as described; the substitution table is current.
- [ ] Independent VBA review done and its findings fixed or rejected in writing.
- [ ] The sample output in `output/` is rebuilt by v2 and opens in LibreOffice without repair prompts.
- [ ] `macro-instructions.md`, `excel-test-protocol.md`, `business-case.md` updated; Config keys in the documents match the module.
- [ ] Both PDFs rebuilt, exactly one page each, no overfull boxes; Rev. 4 and Rev. 2 in their headers.
- [ ] Every new, changed, renamed or deleted file has a correct `MAP.md` row. `git status` clean after the push to `main`.
- [ ] Final message lists what was not run and what the owner still has to do.

---

## 10. Open questions after this round (defaults apply until answered)

| # | Question | Default |
|---|---|---|
| 1 | What is the city column called in the three exports? | Auto-detect; `CityColumn` overrides. |
| 2 | Are CTL= and MX= the units in stock (sum of TOTAL) or the number of inventory lines? Compare the summary with the manager's last hand-built figures on the first real run. | Units (sum of TOTAL). |
| 3 | Is the Open Order export used for anything in the report? | Read and logged only. |
| 4 | Does the proposed order status order (ORDERED, Not yet processed, INVOICED, Completely processed, other, CANCELLED) read right? | As proposed. |
| 5 | Should ORGANIZATION_NAME (and the customer number) appear on the inventory tab? What does TOTAL count (on hand only, or in transit and demo too)? | Left out; TOTAL as exported. |
| 6 | Who owns the size-class file and where does it live (SharePoint path)? | The starter in `output/` until the owner's list arrives. |

---

## 11. Things that went wrong once already (carry-over from `stronger-ai-execution-plan.md`)

- Reading the sample workbooks cell by cell wasted context. Use scripts.
- Theme colours: Excel names the first four theme slots backwards and LibreOffice has none; v2 uses RGB only, so this trap is gone, but do not reintroduce `ThemeColor`.
- The `.bas` must stay CRLF and ASCII, with no line over 1,000 characters, no statement over 20 continuations and no `Array(...)` over 50 arguments; `check_bas.py` catches it, run it after every edit.
- A variable named `cDate` (a reserved word) once broke compilation; `check_bas.py` checks reserved names, keep it that way.
- The hand-built CTL=/MX= figures (21/2 and 15/15) will never match the anonymised sample. Do not chase them; the first real run settles question 2.
- openpyxl filter criteria do not hide rows by themselves; set the hidden flags too, or Excel shows every row with a filter icon that lies.

---

## 12. The prompt to paste into the Stronger AI review

> Read `output/open-orders-intake-inventory/stronger-ai-implementation-plan.md` in full and execute it as the implementer. Follow its standing rules exactly: push to `main`, keep `MAP.md` current, build to the Section 5 contract, mirror every rule in the macro, the Python reference and the checker, run the macro in LibreOffice and compare it with the reference, and keep the one-page manual and the questions PDF current. Do not spawn agents of the most expensive tier; if you delegate, use a smaller AI for mechanical work and a second AI reviewer for the independent VBA review, with self-contained briefs. Work through Phases 0 to 6 in order and finish with the Section 9 checklist filled in honestly, including what was not run and what I still owe (the Excel run, the real size-class list, the Section 10 answers).
