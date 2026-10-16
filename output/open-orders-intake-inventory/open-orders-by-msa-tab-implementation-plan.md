# Open Orders Macro: "Open Orders by MSA" Tab, Implementation Plan for Claude Opus (high effort)

This file is an **architect's brief**. It was written by one AI session (2026-10-14) for a later session, Claude Opus run at high reasoning effort (the "implementer"), that adds one tab to the workbook the Open Orders report macro builds. The owner runs the implementer with the prompt at the very end (Section 12).

The feature request, in the owner's words: *Add a tab to the output file. The purpose of this tab is to display in a chart orders that have not yet been processed and orders that have been ordered. In the "Order intake by MSA" file, include the items in column D that read "Not yet processed" and "Ordered". Group the items into the categories in column H, "Product Type Abbreviation". Then group each category with a total count of Ordered Quantity (column G) by MSA (column K).*

Everything below turns that into a contract the macro, the Python reference, the checker, the fixtures, the LibreOffice run and the documents all follow. Nothing in the existing tabs changes except where Section 5 says so.

---

## 1. Mission and deliverables

Macro **v2.5** builds the same workbook as v2.4 plus one tab, **Open Orders by MSA**, placed between Inventory by Month and Run Log. The tab holds one table (ordered quantity of open orders, one row per MSA, one column per product type, with totals) and one clustered column chart of that table (one bar group per product type, one bar per MSA). The table is live SUMIFS formulas over the Order Intake tab, so it always agrees with the rows the report holds; the chart reads the table.

| Deliverable | Path | Notes |
|---|---|---|
| Macro v2.5 | `source/open-orders-intake-inventory/build-open-orders-report.bas` | The new tab, the File Format and Run Log additions, the version bump. CRLF, ASCII; `check_bas.py` passes. |
| Reference v2.5 | `source/open-orders-intake-inventory/build_report_reference.py` | Same tab, same formulas, same chart (openpyxl), same log lines. |
| Checker v2.5 | `source/open-orders-intake-inventory/check_report.py` | New rules R5a, R5b, E7, C5 (Section 7.2); R1, R2, R2F, E1, E6 updated. |
| Fixtures | `source/open-orders-intake-inventory/make_test_fixtures.py`, `test-fixtures/expected.json` | `open_orders` per run, new log lines, sheet order; the hand-computed figures in Section 7.1 asserted in `self_check`. |
| LibreOffice run | `source/open-orders-intake-inventory/run_macro_libreoffice.py` | One new substitution (`TestChart`), the chart drawn by the normaliser (Section 7.3). |
| Documents | `output/open-orders-intake-inventory/macro-instructions.md`, `excel-test-protocol.md`, `source/.../open-orders-macro-manual.tex` + `output/.../open-orders-macro-manual.pdf` (Rev. 9) | Section 8. |
| Map rows | `MAP.md` | Every edited file's row says v2.5 and names the new tab where it lists tabs; this plan's row marked "Executed". Same commit. |

Not deliverables: no Config setting, no change to the Order Intake or Inventory by Month rows, formulas, filters or Run Log lines that exist today (new lines are added, existing ones are not reworded), no PowerPoint, no second chart, no new PDF besides the rebuilt manual. The flow drawings, the team review deck and the questions PDFs stay at v2.4 (they describe that revision); say so in the final report.

---

## 2. Standing rules

1. Read the repo rules file (`CLAUDE.md`) and `MAP.md` first and obey them: push directly to `main` (a cloud session also pushes the same commit to its `claude/<name>` branch), no pull request, update `MAP.md` in the same commit, source under `source/open-orders-intake-inventory/`, output under `output/open-orders-intake-inventory/`.
2. **One rule, three places.** Every rule in Section 5 is implemented in the macro, in `build_report_reference.py` and in `check_report.py`; the fixtures assert it. A rule you cannot implement in one of the three is not implemented: stop and say so in the report.
3. **Do the work yourself.** If you delegate at all, use a smaller AI for mechanical edits (document wording, MAP rows) with a self-contained brief, and at most one second AI reviewer for an independent read of the VBA diff (Section 6, Phase 5). Never let a worker change the contract.
4. **Every change is validated before it is committed**: `check_bas.py` on the module, the reference on the fixtures and the Austin sample, `check_report.py --expected` on every build, the macro in LibreOffice compared with the reference (`--compare`). What you could not run (real Excel), you say.
5. Facts beat decoration. The tab uses the v2.4 style sheet as it is (Section 5.6); add no colour, font or border the style sheet does not already have, except the chart's series palette in Section 5.2.
6. Report faithfully at the end: what passed, what was not run, what you decided that this brief left open, and what the owner still owes.

---

## 3. Read first, then stop reading

| Read | Why | How much |
|---|---|---|
| `CLAUDE.md`, `MAP.md` | Repo rules; where things are. | All (short). |
| This brief | The contract. | All. |
| `source/open-orders-intake-inventory/build-open-orders-report.bas` | Header comment (lines 1 to 60): the rules and the tab list. `BuildOpenOrdersReport` (from line 215): where sheets are created, written and paged. `PrepareData` (from 1578): how `mOI` is filled (column 5 = ORDER_STATUS as exported, 7 = SUM(F.ORDERED_QUANTITY), 8 = PRODUCT_TYPE_ABBR upper-cased, `mOICity` = MSA registry index, `mOIStatus` = status rank). `BuildLogLines` (from 1926): the per-MSA log lines and their order. `WriteFileFormat` (from 2055): Tables 3 and 4. `WriteInventory` (from 2206): the "Units in" table is the pattern for the new table (SUMIFS by MSA, Total column and row, `SUM_ALIGN`, `StyleTable ... unbanded 1`, double rule above Total). The style helpers (from 2370). | Those places. |
| `source/open-orders-intake-inventory/build_report_reference.py` | `Build.write` (line 716) creates the sheets; `Build.inventory` (836) is the pattern; `run_log_lines` (659) the log. | Those three methods and the helpers they call. |
| `source/open-orders-intake-inventory/check_report.py` | The docstring (rule list), `SHEETS`, `check_style`, `check_file_format`, `check_inventory` (pattern for the new table's checks), `run_expected`, `run_compare`, `main`. | Those. |
| `source/open-orders-intake-inventory/make_test_fixtures.py` | `compute_run` (229): how expected facts are computed from the tables, and `self_check` (410): the hand-computed assertions. The Order Intake table (from line 68) to hand-compute Section 7.1. | Those. |
| `source/open-orders-intake-inventory/run_macro_libreoffice.py` | The docstring table of substitutions, `EXTRA`, `module_text` (114), `normalise` (358). | Those. |
| `output/open-orders-intake-inventory/macro-instructions.md` | Sections 1, 3, 5, 6, 7, 8: what to update. | Those. |
| `output/open-orders-intake-inventory/excel-test-protocol.md` | Runs A, B, C and Table A: where the new checks go. | Those. |
| `source/open-orders-intake-inventory/open-orders-macro-manual.tex`, `build.sh` | The one-page manual and how it is built. | All (short). |

Do not open the sample workbooks or the fixture `.xlsx` files cell by cell; use scripts. Do not read the other plans in `output/` beyond Section 11 of `stronger-ai-implementation-plan.md` (the traps list, repeated in Section 11 here).

---

## 4. The request in the report's terms

The owner's column letters are those of the Snowflake "Order Intake" export for all MSAs. The macro reads exports by **header name**, not by letter (the fixtures and the Austin sample have the same columns in other positions), so the rule is stated by header:

| Owner said | Export header | Where it is in the report's Order Intake tab |
|---|---|---|
| Column D, "Not yet processed" and "Ordered" | `ORDER_STATUS` | Column E, written as exported (`ORDERED`, `Not yet processed` in the exports seen so far) |
| Column H, Product Type Abbreviation | `PRODUCT_TYPE_ABBR` | Column H, trimmed and upper-cased: `CTL`, `MX`, `SSL` (other types are dropped by v2.4 and never reach the report) |
| Column G, Ordered Quantity | `SUM(F.ORDERED_QUANTITY)` | Column G, number format `0` |
| Column K, MSA | the MSA column (Config `MSAColumn`, else auto-detected) | Column B, first spelling seen, `(no MSA)` when blank |

**The rule.** An *open order row* is an Order Intake row (any size class, CTL, MX or SSL) whose `ORDER_STATUS`, compared the way Excel's SUMIFS compares text (whole text, ignoring case, not ignoring stray spaces), is `ORDERED` or `Not yet processed`. The tab shows, for each MSA and product type, the **sum of `SUM(F.ORDERED_QUANTITY)`** over the open order rows ("total count of Ordered Quantity" is read as that sum; an order line's quantity, not a count of lines; see Section 10, question 1). Non-numeric quantities add nothing (SUMIFS ignores text; the macro's own tally uses `IsNumber`).

Two facts worth knowing, from the data in the repo: in the Austin sample the two statuses together count 163 rows and `IS_OPEN_ORDER` TRUE counts 163 rows, but they are not the same rows (one `Completely processed` row is flagged TRUE and one `Not yet processed` row FALSE; the totals coincide by chance), which is why Section 5.4 logs the disagreement count rather than using the flag; and in the Order Intake export the `CANCELLED` rows and the `SSL` rows the report hides are excluded by status and by the chart's series choice respectively, so the hidden rows never need to be consulted.

---

## 5. The contract, fixed

### 5.1 Sheet order

`File Format`, `Order Intake`, `Inventory by Month`, `Open Orders by MSA`, `Run Log`. The new sheet is created with `Worksheets.Add(After:=wsInv)` before the Run Log sheet; `SetUpPage` is called for it with no title rows and `pagesTall` 1 (it fits one page); gridlines hidden; no freeze panes; no AutoFilter.

### 5.2 The Open Orders by MSA tab

Row and column positions are fixed so the checker can address them. `nc` = number of MSAs (`mNCities`), in the alphabetical order of the inventory summary (`SortedCity`), `(no MSA)` included; `nt` = 3 if Config `ShowSSL` is TRUE, else 2 (the product types the chart plots).

| Cell(s) | Content |
|---|---|
| A1 | Title `Open Orders by MSA` (StyleTitle: 14 bold, medium rule under row 1 across A:E). |
| D1 | Note (italic): `Source: the Order Intake tab, built yyyy-mm-dd. Ordered quantity of orders with ORDER_STATUS ORDERED or Not yet processed, by MSA and product type.` followed by ` Chart: CTL and MX.` when ShowSSL is FALSE, else ` Chart: CTL, MX and SSL.` |
| A3 | Caption (CaptionCell): `Table 1.` bold, then ` Open orders by MSA: sum of SUM(F.ORDERED_QUANTITY) over the Order Intake rows whose ORDER_STATUS is ORDERED or Not yet processed.` |
| A4:E4 | Header `MSA`, `CTL`, `MX`, `SSL`, `Total` (PutRow; StyleTable with `SUM_ALIGN` = `LRRRR`, `unbanded` 1). |
| A5:A(4+nc) | The MSA names (SafeCellText), one per row, alphabetical. |
| B..D of each MSA row r | `=SUMIFS(OI!G,OI!B,$A r,OI!H,B$4,OI!E,"ORDERED")+SUMIFS(OI!G,OI!B,$A r,OI!H,B$4,OI!E,"Not yet processed")` where `OI!X` stands for `'Order Intake'!$X$3:$X$last`, `last` = 2 + max(number of Order Intake rows, 1). Written with `.Formula` (comma separators, as the inventory SUMIFS are). Number format `#,##0`. |
| E of each MSA row r | `=SUM(B r:D r)`, `#,##0`. |
| Row 5+nc | `Total` row: A = `Total`; B..E = `=SUM(X5:X(4+nc))` (when nc = 0: the value 0); bold; double rule above; not banded; `#,##0`. |
| Column widths | `Array(36, 12, 12, 12, 12)` (the MSA names run to 35 characters). |
| Chart | One `ChartObject` named `OpenOrdersChart`, drawn only when nc > 0, anchored at the top-left of cell A(7+nc) (two empty rows under the Total row), 700 pt wide, 320 pt high. Source `A4:(B+nt-1)(4+nc)`, i.e. `A4:C(4+nc)` or `A4:D(4+nc)`, `PlotBy:=xlRows`: one series per MSA named from column A, categories `CTL`, `MX` (and `SSL`) from row 4. `xlColumnClustered`, gap width 80. Title `Open orders by product type and MSA (ordered quantity)`. Legend at the bottom. Value axis title `Ordered quantity`, number format `#,##0`, horizontal major gridlines on, no vertical gridlines. Data labels on every series, `#,##0`. Every chart text in Cambria 10 (title 12 bold). Series fills from the palette below, in MSA order, wrapping after eight. Chart area and plot area without fill colour (white), chart border thin grey (`D9D9D9`). When nc = 0, no chart; A(7+nc) holds the italic note `No MSAs: no chart.` |

The Total column and Total row are **not** in the chart's source range.

Palette (an array function `ChartPalette()` of eight Longs, built with `RGB`; the first is the report green `HEAD_FILL`):

| # | Hex | Use |
|---|---|---|
| 1 | 548235 | report green |
| 2 | 7F7F7F | grey |
| 3 | 2F5597 | blue |
| 4 | C55A11 | rust |
| 5 | 7030A0 | purple |
| 6 | 70AD47 | light green |
| 7 | A5A5A5 | light grey |
| 8 | 264478 | navy |

The macro's chart code lives in one procedure, `DrawOpenOrdersChart(ws, nc, nt, anchorRow)`, called once from `WriteOpenOrders`; nothing else in the module touches charts. (The LibreOffice harness replaces that one call; Section 7.3.)

### 5.3 File Format additions

- **Table 3 (Default filters)** gets a fifth data row: `Open Orders by MSA` | `PRODUCT_TYPE_ABBR` | `CTL, MX on the chart` (ShowSSL FALSE) or `CTL, MX, SSL on the chart` (ShowSSL TRUE) | `ShowSSL = TRUE|FALSE`. The italic "Note. Clear a filter to see the rows it hides." line moves down one row. `FF_BUILD_ROW` becomes **39** in the macro and the reference (Table 4's caption moves from row 38 to 39, one empty row kept above it); `check_report.py` finds captions by text and needs no row constant, but its Table 3 check reads **five** rows and the note at `r + 7`.
- **Table 4 (This build)** gets one pair after `MSAs`: `Open orders` | `ORDERED and Not yet processed: <n> rows, <q> units (CTL <a>, MX <b>, SSL <c>), on the Open Orders by MSA tab`, where the figures are the macro's own tally over the same rows the SUMIFS count (Section 4), all three product types regardless of ShowSSL. `CountText` wording (`1 row`, `2 rows`); `units` always plural, as the inventory log lines do.

### 5.4 Run Log additions

The lines are added, in this order, without changing any existing line:

1. After each MSA's inventory line (the pair of lines per MSA in `BuildLogLines`), one line: `<MSA>: Open orders <n rows> (ORDERED <a>, Not yet processed <b>): CTL <x>, MX <y>, SSL <z> units`, or `<MSA>: no open orders (ORDERED or Not yet processed)` when the MSA has none. Counts over all three product types regardless of ShowSSL.
2. After the `Inventory: ...` lines (the "shows zeros" block), when every MSA has none: `Open orders: no rows with ORDER_STATUS ORDERED or Not yet processed - the Open Orders by MSA tab shows zeros` (it contains "shows zeros", so the Last run box lists it).
3. When some kept rows have a status that is `ORDERED` or `NOT YET PROCESSED` after `Trim` and `UCase` but not as written (stray spaces, so SUMIFS will not count them): `WARNING: <n rows> have ORDER_STATUS ORDERED or Not yet processed with extra spaces and are not counted on the Open Orders by MSA tab: "<text as exported>", ...` (distinct texts, in quotes, in the order first seen; at most `MAX_ROWS_LISTED` of them, then `...`). This mirrors the v2.4 decision that hidden-row matching follows Excel's filter (ignores case, not stray spaces).
4. When some kept rows have `IS_OPEN_ORDER` disagreeing with the rule (TRUE with another status, or FALSE with one of the two statuses): `Open orders: <n rows> where IS_OPEN_ORDER disagrees with the status (TRUE with another status, or FALSE with ORDERED / Not yet processed)`. Information only (it matches no `RunLogLineToCheck` pattern). A value that is not a Boolean counts as neither TRUE nor FALSE and never disagrees.

The macro's final message box and the Last run box gain nothing new beyond what these lines produce through the existing triage.

### 5.5 Config

No new setting. ShowSSL already decides whether SSL is on the chart (Section 5.2); the table always has the SSL column, exactly as the inventory "Units in" table does. The Config help text (`ConfigHelp`) for ShowSSL gains the words `and off the Open Orders by MSA chart`; the upgrade path is unchanged (a v2.4 Config sheet is a v2.5 Config sheet).

### 5.6 Style

The v2.4 style sheet, unchanged: Cambria 10, green header row (`HEAD_FILL`, white bold, medium rule above, double rule under), horizontal rules only, banding by conditional format on the MSA rows (not the Total row), right-aligned numbers indented one step, landscape one page, page header and footer as the other tabs. The chart's fonts and palette are the only additions (Section 5.2). No theme colours anywhere (`check_bas.py` and the LibreOffice runner both refuse `ThemeColor`).

### 5.7 Version

`2.4` becomes `2.5` in every file the implementer edits where it names the macro version: the module header and `BUILDER_TEXT`, `VERSION` in the reference, the checker's docstring, the runner's docstring, `'version'` in `make_test_fixtures.py` (and therefore `expected.json`), the instructions, the test protocol, the manual (Rev. 9, dated the build day), the MAP rows of those files. Files not edited keep `2.4`.

---

## 6. Work plan

### Phase 0: Orient
Read Section 3. Run the baseline so you know it is green before you touch anything:

```
cd source/open-orders-intake-inventory
python3 check_bas.py
python3 make_test_fixtures.py test-fixtures      # regenerates the fixtures and expected.json; self_check passes
python3 build_report_reference.py test-fixtures /tmp/ref-default.xlsx --size-class-file "test-fixtures/Size Class.xlsx"
python3 check_report.py /tmp/ref-default.xlsx --expected test-fixtures/expected.json
python3 run_macro_libreoffice.py test-fixtures /tmp/lo-default.xlsx --size-class-file "test-fixtures/Size Class.xlsx"
python3 check_report.py /tmp/lo-default.xlsx --expected test-fixtures/expected.json --compare /tmp/ref-default.xlsx
```

If the baseline is not green, stop and report; do not build on a red baseline.

### Phase 1: The executable spec first (fixtures and checker)
1. `make_test_fixtures.py`: add `open_orders` to each run (Section 7.1), the new log lines in their places, the sheet order, `version` 2.5; extend `self_check` with the hand-computed figures of Section 7.1 (they were computed independently for this brief; if your computation disagrees, find out which is wrong before going on and say so in the report).
2. `check_report.py`: Section 7.2. Run it against the v2.4 reference output: R1, R2F, E1, E6 and the new rules must now **fail** for the right reasons (no tab, four filter rows, missing lines). That failing run is evidence the checks bite; keep its output for the report.

### Phase 2: Reference v2.5
`build_report_reference.py`: `Build.write` creates the sheet in order; a new `open_orders(self, ws)` method on the pattern of `inventory`; the chart with openpyxl (`BarChart`, `type='col'`, `grouping='clustered'`, `add_data(Reference(ws, min_col=1, max_col=1+nt, min_row=4, max_row=4+nc), from_rows=True, titles_from_data=True)`, `set_categories(Reference(ws, min_col=2, max_col=1+nt, min_row=4))`, anchored at `A{7+nc}`, width 24.7 cm, height 11.3 cm, title, legend bottom, axis title, data labels, series fills from the palette); Table 3 row, Table 4 pair, `FF_BUILD_ROW` 39, the log lines. Build the fixtures (default and variant) and the Austin sample; `check_report.py --expected` passes on both fixture runs; open the Austin build's `xl/charts/chart1.xml` and confirm the refs by eye once.

### Phase 3: Macro v2.5
`build-open-orders-report.bas`: the header comment (tab list, rule bullet for the new tab, version), `BUILDER_TEXT`, `FF_BUILD_ROW`, a new `WriteOpenOrders(ws, opt, builtAt)` on the pattern of the "Units in" part of `WriteInventory`, `DrawOpenOrdersChart`, `ChartPalette`, the tally for Table 4 and the log lines (a small `Private Type` or module arrays filled in `BuildLogLines`, as the inventory units are), `ConfigHelp` text. `check_bas.py` after every edit. Then the LibreOffice run (Section 7.3) on both fixture runs and the Austin sample, each compared with the reference build (`--compare`) and checked with `--expected`.

### Phase 4: Documents
Section 8. Rebuild the manual with `./build.sh source/open-orders-intake-inventory/open-orders-macro-manual.tex`; it must stay one page (`pdfinfo` or `pdftotext` page count).

### Phase 5: Review and ship
Optional second AI reviewer: give it the VBA diff only and this brief's Section 5, ask for compile risks (reserved names, `Array` sizes, `Chart` members that need Excel 2010, continuations) and rule mismatches; fix what is real. Fill in Section 9 honestly. One commit (or a few logical ones), `MAP.md` in the same commit, push to `main` (and to the session branch).

---

## 7. Fixtures and checks

### 7.1 Expected facts (hand-computed for this brief from the fixture exports and the Austin sample)

`open_orders` per run: `[[MSA, CTL, MX, SSL, Total], ...]` in MSA order, then `['Total', ...]`, plus `open_orders_chart` = `{'series': nc, 'categories': ['CTL', 'MX'] or ['CTL', 'MX', 'SSL']}`.

Fixtures, both runs (`default`: ShowSSL FALSE; `variant`: InventoryMonth 2026-08-15, ShowCancelled TRUE, ShowSSL TRUE; the table is the same in both, the chart's categories differ):

| MSA | CTL | MX | SSL | Total | open rows |
|---|---|---|---|---|---|
| (no MSA) | 1 | 0 | 0 | 1 | 1 |
| Amarillo | 8 | 5 | 1 | 14 | 13 |
| Brownsville | 1 | 0 | 0 | 1 | 1 |
| Dallas-Fort Worth [North] Metroplex | 4 | 1 | 0 | 5 | 5 |
| Total | 14 | 6 | 1 | 21 | 20 |

Over the 20 open rows: 18 `ORDERED`, 2 `Not yet processed`; no stray-space statuses (log line 3 of Section 5.4 does not appear); `IS_OPEN_ORDER` agrees with the status rule on every kept fixture row (log line 4 does not appear either). The ` mx ` row of the fixture counts as MX (the tab reads column H, already upper-cased). Compute the per-MSA ORDERED / Not yet processed split from the table in `make_test_fixtures.py` and assert it in `self_check` too.

Austin sample (one MSA from the file name, `ShowSSL` FALSE): Austin 86 | 34 | 43 | 163, 163 open rows (104 `ORDERED`, 59 `Not yet processed`). `IS_OPEN_ORDER` disagrees on 2 rows (a `Completely processed` row flagged TRUE, a `Not yet processed` row flagged FALSE), so the Austin Run Log carries `Open orders: 2 rows where IS_OPEN_ORDER disagrees with the status (TRUE with another status, or FALSE with ORDERED / Not yet processed)`. Put these in the test protocol's Run C.

### 7.2 Checker v2.5 (`check_report.py`)

Changed: `SHEETS` gains `Open Orders by MSA` before `Run Log` (R1, E1, C1); `check_style` learns kind `oo` (title and note style, header A4:E4 in `SUM_ALIGN`, widths `(36, 12, 12, 12, 12)`, no freeze panes, gridlines off, `#,##0` on B5:E(5+nc), banding on rows 5..4+nc only, print setup one page tall, no title rows); `check_file_format` reads five filter rows (the fifth as Section 5.3) and the note at `r + 7`, and requires the `Open orders` item in Table 4; E6 compares the new log lines like any other.

New rules:

| Rule | Sheet | Checks |
|---|---|---|
| R5a | Open Orders by MSA | A1 title, A3 caption text, A4:E4 headers exact; MSA rows = the MSAs of the Order Intake tab, alphabetical, one each; every B..D cell is the two-SUMIFS formula of Section 5.2 with the exact ranges (first row 3, last row 2 + max(n, 1)) and criteria; E cells `=SUM(B r:D r)`; Total row formulas; double rule above Total; Total bold; when evaluated against the Order Intake rows (the checker's own `evaluate`, extended to a `+` of two SUMIFS if needed) the values equal the checker's own tally of the open rows. |
| R5b | Open Orders by MSA | Read from the `.xlsx` zip (openpyxl drops charts on load): the sheet's rels name exactly one drawing, the drawing exactly one chart part when nc > 0 and none when nc = 0; `c:barChart` with `c:barDir val="col"` and `c:grouping val="clustered"`; `c:ser` count = nc; series i has `c:tx` ref `'Open Orders by MSA'!$A$(4+i)`, `c:cat` ref `$B$4:$C$4` (or `$D$4` with ShowSSL), `c:val` ref `$B$(4+i):$C$(4+i)` (or `$D$`); a `c:title`; the anchor's `xdr:from/xdr:row` >= 6 + nc (below the table). Compare refs after stripping `$` and the sheet quotes, so Excel's and openpyxl's spellings both pass. |
| E7 | Open Orders by MSA | `--expected`: the table values (formulas evaluated) equal `open_orders`; series count and categories equal `open_orders_chart`. |
| C5 | Open Orders by MSA | `--compare`: A4:E(5+nc) equal cell for cell (formula text and evaluated value); same series count, categories and anchor row. |

### 7.3 LibreOffice run (`run_macro_libreoffice.py`)

Add to `EXTRA` (count 1): the call `DrawOpenOrdersChart ws, nc, nt, anchorRow` (exact text as written in `WriteOpenOrders`) becomes `TestChart ws, nc, nt, anchorRow`; the harness records `CHART|<sheet>|nc|nt|anchorRow`; `normalise` draws the chart with the same openpyxl code the reference uses (factor it into a small function both import, or copy it verbatim and say so). Add the row to the docstring's substitution table: *DrawOpenOrdersChart -> TestChart (records it): LibreOffice's VBA chart support is partial, and openpyxl drops any chart when it loads the file in the normaliser, so the normaliser draws the recorded chart.* Everything else about the tab (title, table, formulas, rules, page setup) runs as written through the existing substitutions. Expect the `SetUpPage` property count (16) and `AUTOFILTER_CALLS` (4) to be unchanged; if an edit changes them, the runner stops and tells you.

### 7.4 Static check (`check_bas.py`)

No change expected. Watch: no `Dim` named `chart`, `series` or `title` colliding with the reserved list is not an issue, but `Name` and `Type` are reserved words (use `co.Name` as a member, never as a variable); keep `Array(...)` under 50 arguments and statements under 20 continuations; the chart procedure must not use `CreateObject`.

---

## 8. Documents

### 8.1 `macro-instructions.md` (v2.5)
Section 1 (the folder) and the header: five tabs. Section 3 (each run): one sentence, the Open Orders by MSA tab shows the open orders chart; SSL is on it only with ShowSSL TRUE. Section 5: a new 5.5a *Open Orders by MSA tab* stating the rule of Section 4 and the table and chart of 5.2 in prose; 5.7 the new File Format rows and the four log lines. Section 6 (validated): the figures of Section 7.1 and what was run. Section 7 (decisions): the decisions of Section 10 of this brief as "decided during the build". Section 8: how to change the two statuses (one place in each of the three files) and the palette. Grep the file for `four tabs`, `four-tab` and `4 tabs` and fix each.

### 8.2 `excel-test-protocol.md` (v2.5)
Setup step: five tabs in order. Run A: a **CHECK 6 (Open Orders by MSA, Excel only)**: the table of Section 7.1, the formula in B6 shown exactly as Excel displays it, the chart (four series in the legend, two bar groups CTL and MX, data labels, title), print preview one page; Table A gains a row. Run B: the chart has three groups (SSL added), the table unchanged. Run C: Austin 86 / 34 / 43 / 163, one series. Run D and E unchanged. The Run Log text blocks in Runs A and B gain the new lines in their places (copy them from `expected.json` after Phase 1, never from memory).

### 8.3 Manual, Rev. 9 (`open-orders-macro-manual.tex`, one page)
Header: `Rev.~9 \quad <build date>`; Scope: v2.5, the tab list gains "Open Orders by MSA (ordered quantity of orders ORDERED or Not yet processed, by MSA and product type, with a chart)"; Running the report step (c): "...read the Run Log and Last run box; the Open Orders by MSA tab charts the open orders"; Table 1, ShowSSL row: "FALSE hides skid-steer loaders behind a filter and off the chart." Road map step 5 (Feed the deck) unchanged in wording; its "Main actions" may add "the Open Orders chart is the first deck figure". Rebuild with `./build.sh`; confirm one page; commit `.tex` and `.pdf` together.

### 8.4 `MAP.md`
Rows of every edited file updated (v2.5, the new tab named where tabs are listed); this plan's row ends with "Executed <date>."

---

## 9. Acceptance checklist (fill in, honestly, in the final report)

- [ ] `check_bas.py`: PASS.
- [ ] `make_test_fixtures.py`: self_check passes with the Section 7.1 figures asserted.
- [ ] Reference, fixtures default and variant: `check_report.py --expected` PASS (every rule, R5a, R5b, E7 included).
- [ ] Reference, Austin sample: `check_report.py` PASS (no `--expected`), Austin 86 / 34 / 43 / 163.
- [ ] Macro in LibreOffice, fixtures default and variant and Austin: `check_report.py --expected` (fixtures) and `--compare` against the reference builds PASS.
- [ ] A v2.4 build fails the v2.5 checker for the expected reasons (output kept).
- [ ] The chart XML of one Excel-format build inspected by eye: refs, type, anchor.
- [ ] Manual Rev. 9: one page, PDF rebuilt and committed with the `.tex`.
- [ ] Instructions and test protocol updated; no "four tabs" left.
- [ ] `MAP.md` rows updated in the same commit; this plan marked Executed.
- [ ] Not run: real Excel (the owner runs the test protocol); the chart's look in Excel (fonts, palette, label overlap with many MSAs) is unverified until then.
- [ ] Section 10 decisions and anything the implementer decided beyond them listed for the owner.

---

## 10. Decisions taken in this brief, and open questions (defaults apply until answered)

| # | Question | Default taken here |
|---|---|---|
| 1 | "Total count of Ordered Quantity": sum of `SUM(F.ORDERED_QUANTITY)`, or number of order lines? | Sum of the quantity column. The row counts are on the Run Log and in File Format Table 4, so both figures exist. |
| 2 | Table shape: MSA rows with product type columns (chosen; same shape as the inventory "Units in" table, so the two tables read side by side), or product type rows with MSA columns? | MSA rows. The chart groups by product type either way. |
| 3 | Chart shape: one bar per MSA inside each product type group (chosen), or a stacked column per product type, or a multi-level category axis with one series? | Clustered columns, series per MSA. With more than about twelve MSAs the legend and labels get crowded; say so in the report if the real export has that many, and propose the multi-level axis then. |
| 4 | SSL on the chart: follow ShowSSL (chosen), always, or never? | Follows ShowSSL; the table always has the SSL column. |
| 5 | Statuses matched as written (SUMIFS semantics: case-insensitive, spaces matter), with a WARNING for stray-space rows (chosen), or trimmed by writing a trimmed status into Order Intake column E? | As written; the Order Intake contract (E2) does not change. |
| 6 | Tab name `Open Orders by MSA`. The Open Order export (read and logged only, open question 3 of v2.4) is **not** the source; the note on the tab says the source is the Order Intake tab. | Keep the name unless the owner objects. |
| 7 | A split by status (ORDERED vs Not yet processed) on the tab? | Not built; it is on the Run Log per MSA. Add a second table only if asked. |
| 8 | Should the chart also feed the PowerPoint deck (road map step 5)? | Out of scope; this tab is the first figure the deck step can copy. |

---

## 11. Things that went wrong once already, and new traps

Carried over from the v2 plan:
- Reading the sample workbooks cell by cell wasted context. Use scripts.
- No theme colours (`ThemeColor`); RGB only.
- The `.bas` must stay CRLF and ASCII, no line over 1,000 characters, no statement over 20 continuations, no `Array(...)` over 50 arguments; `check_bas.py` after every edit.
- Reserved words as variable names (`cDate` once) break compilation; `Name`, `Type`, `Date`, `Len` are in the list.
- openpyxl filter criteria do not hide rows; the hidden flags must be set too (not needed on the new tab, which has no filter).

New for this feature:
- **openpyxl drops charts when it loads a workbook.** Any step that opens and re-saves a build (the LibreOffice normaliser, `compare_with_handbuilt.py` if ever extended) loses the chart; draw it after loading, never before. The checker therefore reads the chart from the zip, not through openpyxl.
- **`SetSourceData` on a header-only range** (nc = 0) raises an error in Excel; the contract draws no chart then (Section 5.2).
- **`PlotBy`**: with MSA rows and product-type columns the series are rows (`xlRows`); the wrong constant gives series per product type and fails R5b.
- **`.Formula` strings always use commas** and English function names; only conditional-format formulas need `ListSeparator()`. Do not "fix" the SUMIFS separators.
- **The chart anchor** must be placed after the table is sized (`7 + nc`), and the sheet must be active for `ChartObjects.Add` only if you select anything; prefer not to `Select` at all inside `DrawOpenOrdersChart` (the module activates sheets elsewhere; keep the chart code free of `Select`, `ActiveChart` and `ActiveSheet`).
- **Excel 2010**: `SeriesCollection(i).Format.Fill.ForeColor.RGB` and `DataLabels.NumberFormat` exist from 2007; `ChartObject.Chart.ChartTitle.Text` too. Avoid `FullSeriesCollection` (2013+).
- **`FF_BUILD_ROW` moves to 39**: the macro and the reference both define it; change both or Table 4 lands on different rows and C-rules fail.

---

## 12. The prompt to paste into Claude Opus (high reasoning effort)

> Read `output/open-orders-intake-inventory/open-orders-by-msa-tab-implementation-plan.md` in full and execute it as the implementer, at high reasoning effort. Follow its standing rules exactly: push to `main` (and to this session's branch), keep `MAP.md` current, build to the Section 5 contract, implement every rule in the macro, the Python reference and the checker, assert the Section 7.1 figures in the fixtures, run the macro in LibreOffice and compare it with the reference, and keep the one-page manual (Rev. 9), the instructions and the test protocol current. Do the work yourself; if you delegate, use a smaller AI for mechanical edits and at most one second AI reviewer for the VBA diff, with self-contained briefs. Work through Phases 0 to 5 in order and finish with the Section 9 checklist filled in honestly, including what was not run (real Excel) and the Section 10 decisions I should confirm.
