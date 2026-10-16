# Copilot Prompts: Checking the Real Exports and the Excel Run

Twelve short prompts for a work AI (Copilot or similar) that can read an attached file. P1 to P10 each ask a single question about one file and want a fixed reply format, so the answers can be pasted back for the Stronger AI review and compared with the macro's own figures. P11 is different and comes first: it builds the size class file itself, entirely in Copilot, from the PowerPoint slide that holds the latest size classes. P12 is for the next goal, automated PowerPoint slides: it has Copilot describe last year's deck, slide by slide and without any figures, so the description can be sent to Claude to write the step that fills the deck from the report. They go with macro v2.4 (2026-10-12).

> **Start here: P11 is the most important prompt.** The macro will not build anything without a size class file, and every model missing from it shows up as UNMAPPED. P11 turns the PowerPoint slide (the most up-to-date source) into that file, checks it against the models in your Order Intake export, and gives you a table to save as the workbook Config `SizeClassFile` points at. Do P11 before anything else.

## How to use them

1. Use the Copilot your company approves for work files. Start a **new chat for each prompt**, attach **only the file named**, and paste the text in the grey box as it is. P11 is the exception: its steps run **in one chat**, one after the other, adding a file at each step.
2. Fill in a blank only where the prompt shows `<...>` and tells you to (P6 and P9).
3. Check the row count. Most replies end with `TOTAL ROWS:`. In Excel, click the column letter A of the data sheet: the status bar shows `Count`. Count minus 1 (the header) should equal `TOTAL ROWS`. If it doesn't, Copilot did not read the whole file. Say so when you paste, or use the **Excel way** line under the prompt instead.
4. Copy each reply in full. Paste them all back for the Stronger AI review in one message, each under a line naming the prompt and the file, e.g. `P2 - Inventory by Month 2026.xlsx`.
5. The replies hold only counts, sums, MSA names, model names and dates. No customer names or numbers come back.

| ID | When | Attach | What it settles |
|---|---|---|---|
| **P11** | **First, before anything else** | **the PowerPoint slide, then Order Intake (one chat)** | **The size class file itself (Config `SizeClassFile`), built from the slide and checked against your models** |
| P1 | Before the run | each export, one at a time | The exact headers: the MSA column and where TOTAL is |
| P2 | Before the run | each export, one at a time | MSA spellings agree across the files |
| P3 | Before the run | Order Intake | Statuses, product types, open orders per MSA |
| P4 | Before the run | Inventory by Month | Months present, product type spelling, real dates or text |
| P5 | Before the run | the latest hand-built report | Whether CTL= and MX= equal the sum of TOTAL on each tab |
| P6 | Before the run | Inventory by Month | Units per MSA for one month, worked out independently |
| P7 | Before the run | Order Intake | Model groups to check against the size class list |
| P8 | Before the run | Open Order | Open question 3: is Open Order just the open part of Order Intake? |
| P9 | After the run | the report the macro built | The macro's "Units in" table |
| P10 | After the run | the report the macro built | The Run Log, word for word |
| P12 | When starting the PowerPoint step | last year's deck (one chat, two steps) | A slide-by-slide description of the deck, with no figures, to send to Claude (questions 14 to 16 of the questions PDF) |

Short on time: **P11 first, always.** Then P1 (Inventory by Month), P5, P6, P9 and P10 matter most.

---

## P11. Build the size class file from the PowerPoint slide (most important: do this first)

**Why it matters.** The size class file is the only place the report gets its size classes from. Without it the macro stops before it builds anything; with a model missing, that model's orders land in UNMAPPED. The PowerPoint slide has the latest classes, but not in the layout the macro reads, and its model names may not be spelt the way Snowflake writes MODEL_GROUP. P11 fixes both in Copilot: Step 1 reads the slide, Step 2 checks it against the models you actually sell, Step 3 hands you the finished table. Step 4 is optional.

**Run all the steps in one Copilot chat**, in order, so each step can use the answer before it. If Copilot loses track (it says it cannot see the earlier table), start a new chat, attach the file for that step and paste the last table under the prompt.

### Step 1. Read the slide (attach the PowerPoint file)

```text
Attached is a PowerPoint file. One slide in it lists Bobcat machine models and the size class each belongs to. The layout may be a table, groups of models under class headings, or columns per class.

Task: build one table with one row per machine model on that slide.

How to fill each row:
- Name of machine: the model code only, as written on the slide, without the word Bobcat or the machine type. Examples: T76, S650, E35, T76-2, E50Z. If the slide writes a series as a suffix (for example "T76 R2" or "T76 R-Series"), keep it exactly as written; Step 2 fixes the spelling.
- Slide class: the class label exactly as the slide shows it for that model.
- Size Class: the slide class translated to exactly one of these names, nothing else:
  Track loaders (T...), by rated operating capacity in lb, as the business defines them:
    0<2000, 2000<2300, 2300<2600, 2600<3000, 3000<3500, 3500&Over
  Mini excavators (E...), by nominal tonnes, as the business defines them:
    >00-01, >01-02, >02-03, >03-04, >04-05, >05-06, >06-10
  Skid-steer loaders (S...), by frame size, kept exactly as the slide names them:
    EX Small, Small Frame, Med Frame, Large Frame
- Rated Operating Capacity: the capacity in lb if the slide shows it (for an excavator, its operating weight in lb if shown); otherwise leave it blank.
- HP: the horsepower if the slide shows it; otherwise leave it blank.
- Notes: write "CHECK: slide class does not match" if the slide label does not translate to exactly one of the names above (for example its range crosses a boundary), "CHECK: on a class boundary" if a capacity shown is within 100 lb of a track loader class edge, and "CHECK: listed twice" if a model appears more than once on the slide. Otherwise leave it blank.

Rules:
- Use only what is on the slide. Do not add models from memory. Do not look anything up.
- Do not explain. Reply only in the format below.
- If you cannot find such a slide, reply with one line: CANNOT: <reason>

Reply in exactly this format:
SLIDE: <slide number and title>
CLASS LABELS
<slide class label> -> <Size Class name, or NONE>
TABLE
Name of machine | Slide class | Size Class | Rated Operating Capacity | HP | Notes
<one line per model>
MODELS ON SLIDE: <number>
```

Check before going on: count the models on the slide yourself and compare with `MODELS ON SLIDE`. Read the `CLASS LABELS` lines: each slide label should point to the class you expect. If one is wrong, tell Copilot in the chat (for example `"Under 2,000" is 0<2000`) and ask it to redo the table.

### Step 2. Check it against your models (attach the Order Intake export, same chat)

```text
Attached is the Order Intake export. Use its first sheet. Keep the table from the slide.

Task: compare the slide table with the models in this export.
1. Take the export rows whose PRODUCT_TYPE_ABBR is CTL, MX or SSL (ignore capitals and extra spaces). List every different MODEL_GROUP among them.
2. For each MODEL_GROUP, find its row in the slide table: the same name ignoring capitals and spaces, or the same model written another way (for example "T76 R2" on the slide and "T76-2" in the export). When it is the same model written another way, change Name of machine in the slide table to the export's spelling and add "renamed from <slide name>" to Notes.
3. A MODEL_GROUP with no row in the slide table: add a row with Name of machine = the MODEL_GROUP exactly as in the export.
   - E followed by digits (a mini excavator): leave Size Class blank; Notes "CHECK: not on slide, class unknown". There is no rule for it: the class comes from the slide or the spec sheet.
   - Any other model: give your best Size Class only if you are sure of its rated operating capacity; Notes "CHECK: not on slide". If you are not sure, leave Size Class blank and write "CHECK: not on slide, class unknown".
4. Keep every slide row, including models the export does not have.

Rules:
- Read every row of the export. Do not explain. Reply only in the format below.
- If a column is missing or you cannot read the file, reply with one line: CANNOT: <reason>

Reply in exactly this format:
RENAMED
<slide name> -> <MODEL_GROUP>
ADDED (not on slide)
<PRODUCT_TYPE_ABBR> | <MODEL_GROUP> | <rows in export> | <Size Class or blank>
ON SLIDE, NOT IN EXPORT
<Name of machine>
TOTAL MODEL_GROUPS IN EXPORT: <number>
```

Every model with a blank Size Class must get one before Step 3. Find its class (from its capacity, weight or frame on the spec sheet) and tell Copilot in the chat, for example `T870 is 3500&Over` or `E55 is >05-06`. A model left blank is skipped by the macro, and its orders show as UNMAPPED.

### Step 3. The finished size class table (same chat)

```text
Now give the final size class table: the slide table after Step 2, with the corrections I gave you in this chat.

Rules:
- One row per Name of machine. No blank Size Class. Size Class must be exactly one of the 17 names from Step 1.
- Sort: track loaders (T...), then skid-steer loaders (S...), then mini excavators (E...), then anything else; inside each group by Size Class in the order listed in Step 1, then by Name of machine.
- Reply only with the table as tab-separated text, no other words, starting with this header row:
Name of machine	Rated Operating Capacity	HP	Size Class	Notes
- After the table, one line: ROWS: <number of model rows>
```

If your Copilot can create files, you can instead ask: `Create an Excel workbook with this table on the first sheet, named Size Class, and Size Class formatted as text.` Then download it and go to point 4 below.

Save it as the size class file:

1. Open a new Excel workbook. Select column D (Size Class) and set it to **Text** (Home > Number Format > Text) before pasting, so values like `>00-01` stay as they are.
2. Click cell A1 of the **first sheet** and paste. The macro reads only the first sheet; rename it `Size Class` if you like.
3. Check that `ROWS` matches the number of rows under the header, and that no Size Class cell is empty.
4. Save it as a local `.xlsx`, for example `C:\Temp\oo-test\size-class-master.xlsx`, replacing the starter list.
5. On the Config sheet, click **Choose size class file...** and pick it, then click **Check setup**. It lists any row with a class name it does not know, any model listed twice, and any Order Intake model still missing. If it lists something, fix it in Copilot (same chat) and paste the table again.

### Step 4 (optional). What changed from the starter list (attach `size-class-master.xlsx`, same chat)

Useful to show your manager what the slide changed. Attach the starter list the macro came with (or its printed copy, `size-class-master.pdf`).

```text
Attached is the previous size class list (first sheet; columns Name of machine and Size Class). Compare it with the final table from Step 3.

Rules: compare names ignoring capitals and spaces. Do not explain. Reply only in the format below.

Reply in exactly this format:
CLASS CHANGED
<Name of machine> | <previous Size Class> -> <new Size Class>
NEW (not in the previous list)
<Name of machine> | <Size Class>
DROPPED (in the previous list only)
<Name of machine> | <Size Class>
UNCHANGED: <number>
```

Excel way (for the whole of P11): copy the slide's models and classes into Excel by hand in the four columns `Name of machine`, `Rated Operating Capacity`, `Size Class`, `Notes`, and compare MODEL_GROUP with a PivotTable of the Order Intake export (P7).

## P1. Headers (run once for each export)

```text
Attached is one Excel file. Look at its first sheet only.

Task: find the header row. It is the first row with text in at least three cells. List every header in that row, left to right, with its column letter, exactly as written (same spelling, capitals and spaces). Then count the data rows below the header row (rows that have any value).

Rules:
- Do not guess. Do not explain. Reply only in the format below.
- If you cannot open the file or read the sheet, reply with one line: CANNOT: <reason>

Reply in exactly this format:
FILE: <file name>
SHEET: <first sheet name>
HEADER ROW: <row number>
A: <header>
B: <header>
(one line per column, up to the last header)
TOTAL ROWS: <number of data rows>
```

Excel way: look at the first row of the sheet; the status bar count gives the rows.

## P2. MSA values (run once for each export)

```text
Attached is one Excel file. Use its first sheet. The column with the header MSA holds the area name.

Task: list every different value in the MSA column and how many data rows have it. Values that differ only in capitals or spaces are different values: list each one separately. Count empty MSA cells as (blank).

Rules:
- Read every row. Do not guess or estimate. Do not explain. Reply only in the format below.
- If there is no column headed MSA, reply with one line: CANNOT: no MSA column. Headers are: <list the headers>
- If you cannot read the file, reply with one line: CANNOT: <reason>

Reply in exactly this format, values sorted A to Z:
FILE: <file name>
<MSA value> | <number of rows>
<MSA value> | <number of rows>
(blank) | <number of rows>
TOTAL ROWS: <number of data rows>
```

Excel way: PivotTable with MSA in Rows and Count of MSA in Values.

## P3. Order Intake: statuses, types, open orders

```text
Attached is the Order Intake export. Use its first sheet.

Task 1: list every different value in the column ORDER_STATUS with its number of rows.
Task 2: list every different value in the column PRODUCT_TYPE_ABBR with its number of rows.
Task 3: for each value in the column MSA, count the rows where IS_OPEN_ORDER is TRUE and the rows where it is FALSE. If IS_OPEN_ORDER has any other value, count it under OTHER.
Task 4: give the earliest and the latest ORDER_DATE.

Rules:
- Read every row. Do not guess or estimate. Do not explain. Reply only in the format below.
- Write each value exactly as it appears in the file.
- If a column is missing or you cannot read the file, reply with one line: CANNOT: <reason>

Reply in exactly this format:
FILE: <file name>
ORDER_STATUS
<value> | <rows>
PRODUCT_TYPE_ABBR
<value> | <rows>
OPEN ORDERS BY MSA
<MSA> | TRUE <rows> | FALSE <rows> | OTHER <rows>
ORDER_DATE: <earliest> to <latest>
TOTAL ROWS: <number of data rows>
```

Excel way: PivotTables on ORDER_STATUS, PRODUCT_TYPE_ABBR, and MSA by IS_OPEN_ORDER.

## P4. Inventory by Month: months, types, dates

```text
Attached is the Inventory by Month export. Use its first sheet.

Every row belongs to the calendar month of its INVENTORY_DATE. Any day of a month counts for that month: 1 November 2026 and 30 November 2026 are both 2026-11.

Task 1: list every month that appears, oldest first, with its number of rows and the sum of the column TOTAL for that month.
Task 2: list every different value in the column PRODUCT_TYPE with its number of rows.
Task 3: say whether the INVENTORY_DATE cells are real dates, text, or a mix. Give the first three values exactly as they are shown.

Rules:
- Read every row. Do not guess or estimate. Do not explain. Reply only in the format below.
- If a column is missing or you cannot read the file, reply with one line: CANNOT: <reason>

Reply in exactly this format:
FILE: <file name>
MONTHS
<YYYY-MM> | <rows> | <sum of TOTAL>
PRODUCT_TYPE
<value> | <rows>
INVENTORY_DATE TYPE: <DATES or TEXT or MIXED>
EXAMPLES: <value 1>; <value 2>; <value 3>
TOTAL ROWS: <number of data rows>
```

Excel way: PivotTable with INVENTORY_DATE in Rows (grouped by Months and Years) and Sum of TOTAL in Values.

## P5. The hand-built report: CTL= and MX=

Attach the latest report that was built by hand (one tab per MSA).

```text
Attached is a report workbook built by hand. It has one tab per MSA, plus other tabs such as File Format.

On each MSA tab there is a heading like "Inventory By Month for <month>", cells with the text "CTL=" and "MX=", and a number in the cell to the right of each. Below or near them is a list of inventory rows with the columns PRODUCT_TYPE, MODEL_GROUP, MODEL_NUMBER, INVENTORY_DATE and TOTAL. That list may have no header row.

Task: for each MSA tab, report:
1. the heading text with the month, exactly as written;
2. the number to the right of "CTL=" and the number to the right of "MX=";
3. the sum of TOTAL in that inventory list for the rows whose PRODUCT_TYPE is Compact Track Loaders, and the sum for Mini Excavators;
4. the earliest and the latest INVENTORY_DATE in that list.

Rules:
- Do not guess. Do not explain. Reply only in the format below.
- Skip tabs that are not MSA tabs (File Format and similar).
- If a tab has no CTL= or MX= cell, write NONE for it.
- If you cannot read the file, reply with one line: CANNOT: <reason>

Reply with one line per MSA tab, in exactly this format:
<tab name> | <heading> | CTL= <number> | MX= <number> | SUM CTL <number> | SUM MX <number> | DATES <earliest> to <latest>
```

Excel way: on each tab, select the TOTAL cells of the Compact Track Loaders rows and read Sum in the status bar; then do the same for Mini Excavators.

## P6. Units per MSA for one month (fill in the month)

Use the month of the hand-built report from P5 (the month of its INVENTORY_DATE values), so that P5, P6 and P9 can be compared. Replace `<YYYY-MM>` in two places before sending.

```text
Attached is the Inventory by Month export. Use its first sheet.

Month to use: <YYYY-MM>
A row is in that month if its INVENTORY_DATE is any day of that month.

Task: for that month only, add up the column TOTAL for each MSA, separately for each of these PRODUCT_TYPE values: Compact Track Loaders, Mini Excavators, Skid-Steer Loaders. Ignore capitals and extra spaces when matching PRODUCT_TYPE. Also count the rows of that month for each MSA.

Rules:
- Read every row. Do not guess or estimate. Do not explain. Reply only in the format below.
- Write 0 where an MSA has no rows of a type.
- If a column is missing or you cannot read the file, reply with one line: CANNOT: <reason>

Reply in exactly this format, MSAs sorted A to Z:
MONTH: <YYYY-MM>
MSA | CTL | MX | SSL | ROWS
<MSA> | <sum for Compact Track Loaders> | <sum for Mini Excavators> | <sum for Skid-Steer Loaders> | <rows in the month>
TOTAL | <sum> | <sum> | <sum> | <rows>
```

Excel way: PivotTable with MSA in Rows, PRODUCT_TYPE in Columns, Sum of TOTAL in Values, and INVENTORY_DATE in Filters with only that month's dates ticked.

## P7. Model groups in Order Intake

```text
Attached is the Order Intake export. Use its first sheet.

Task: take the rows whose PRODUCT_TYPE_ABBR is CTL, MX or SSL (ignore capitals and extra spaces). List every different pair of PRODUCT_TYPE_ABBR and MODEL_GROUP among them, with its number of rows.

Rules:
- Read every row. Do not guess. Do not explain. Reply only in the format below.
- Write MODEL_GROUP exactly as it appears in the file.
- If a column is missing or you cannot read the file, reply with one line: CANNOT: <reason>

Reply in exactly this format, sorted by PRODUCT_TYPE_ABBR and then MODEL_GROUP from A to Z:
<PRODUCT_TYPE_ABBR> | <MODEL_GROUP> | <rows>
TOTAL ROWS LISTED: <number>
```

Excel way: PivotTable with PRODUCT_TYPE_ABBR and MODEL_GROUP in Rows and Count of MODEL_GROUP in Values.

## P8. Open Order (open question 3)

```text
Attached is the Open Order export. Use its first sheet.

Task 1: for each value in the column MSA, count the rows where IS_OPEN_ORDER is TRUE and the rows where it is FALSE. If IS_OPEN_ORDER has any other value, count it under OTHER.
Task 2: list every different value in the column ORDER_STATUS with its number of rows.
Task 3: give the earliest and the latest ORDER_DATE.
Task 4: list the headers of the first sheet, left to right, exactly as written.

Rules:
- Read every row. Do not guess or estimate. Do not explain. Reply only in the format below.
- If a column is missing or you cannot read the file, reply with one line: CANNOT: <reason>

Reply in exactly this format:
FILE: <file name>
BY MSA
<MSA> | TRUE <rows> | FALSE <rows> | OTHER <rows>
ORDER_STATUS
<value> | <rows>
ORDER_DATE: <earliest> to <latest>
HEADERS: <header 1>; <header 2>; ...
TOTAL ROWS: <number of data rows>
```

Excel way: PivotTable with MSA in Rows, IS_OPEN_ORDER in Columns and Count in Values.

## P9. After the run: the macro's "Units in" table

Before attaching the built report, make its month match P6: either set `InventoryMonth` on Config to a day in that month before building, or choose the month in cell B3 of the Inventory by Month tab and save. Replace `<YYYY-MM>` before sending.

```text
Attached is a report built by a macro. Open the sheet named "Inventory by Month".

Cell A3 says "Units in" and cell B3 holds a month. Row 4 has the headers MSA, CTL, MX, SSL, Total. The rows below it, down to and including the row whose first cell is "Total", hold one MSA each.

Task: copy that table exactly, using the values shown (not the formulas). Also write the month in B3 as shown.

Rules:
- Do not calculate anything yourself. Do not explain. Reply only in the format below.
- If B3 is not the month <YYYY-MM>, still copy the table, and start your reply with: WRONG MONTH
- If you cannot read the sheet, reply with one line: CANNOT: <reason>

Reply in exactly this format:
MONTH IN B3: <month as shown>
MSA | CTL | MX | SSL | Total
<MSA> | <CTL> | <MX> | <SSL> | <Total>
Total | <CTL> | <MX> | <SSL> | <Total>
```

Excel way: select A3:E<last row> and copy and paste it as text.

## P10. After the run: the Run Log

```text
Attached is a report built by a macro. Open the sheet named "Run Log".

Task: copy every non-empty cell of column A, from the top down, exactly as written. Keep the order and keep every word, number and symbol. Do not shorten, merge or summarise lines.

Rules:
- Do not explain. Do not add anything before or after the lines.
- If you cannot read the sheet, reply with one line: CANNOT: <reason>

Reply in exactly this format, one cell per line:
<cell A1>
<cell A2>
(and so on, down to the last non-empty cell)
```

Excel way: select column A of the Run Log, copy it and paste it as text. This is quicker than Copilot and exact.

## P12. Describe last year's deck for Claude (PowerPoint step)

**Why it matters.** Claude already knows the report's layout, because the macro in this repo builds it. What it does not know is the deck: which slides come from the report, what each chart shows and how it is grouped. P12 has Copilot write that down, slide by slide, **with no figures, dealer names or customer numbers**, so the description can go to Claude, which then writes the macro step that fills the deck on your PC (road map step 5, an *Update PowerPoint* button). Real figures never leave the company.

Before you start: settle question 18 of the questions PDF (may an Excel macro update PowerPoint, and what may be sent to an outside AI). The P12 reply holds no figures, but send it only if the answer allows it.

**Run both steps in one Copilot chat.** Attach last year's deck (the `.pptx`) in Step 1.

### Step 1. List the slides (attach last year's deck)

```text
Attached is a PowerPoint deck. It is a yearly review of Bobcat dealer orders and inventory by MSA (metropolitan area).

Some slides may be built from an Excel report with these parts:
- Order Intake: one row per order line, with Size Class, MSA, ORDER_DATE, REQUESTED_DATE, ORDER_STATUS, IS_OPEN_ORDER, ordered quantity, PRODUCT_TYPE_ABBR (CTL track loader, SSL skid-steer loader, MX mini excavator), MODEL_GROUP, MODEL_NUMBER, ORGANIZATION_NAME.
- Inventory by Month: a table of units in inventory for one month by MSA and product type (CTL, MX, SSL, Total), and below it one row per model and month with TOTAL units.
- Size classes, in three families: track loaders by rated operating capacity (0<2000 up to 3500&Over), mini excavators by nominal tonnes (>00-01 up to >06-10), skid-steer loaders by frame size (EX Small, Small Frame, Med Frame, Large Frame).

Task: list every slide in the deck, in order.

For each slide:
- Title: the slide title exactly as written (or NO TITLE).
- Content: what is on it, as a list of the elements (chart, table, text box, picture), for example "bar chart + text box".
- From report: YES if its figures could come from the report above, MAYBE if they could in part or you cannot tell, NO if not (title slide, agenda, other data).
- Why: one short reason for YES or MAYBE.

Rules:
- Do not copy any figure, dealer name, customer name or customer number. MSA names, model names and size class names are fine.
- Do not explain beyond the format. Reply only in the format below.
- If you cannot read the deck, reply with one line: CANNOT: <reason>

Reply in exactly this format:
DECK: <file name>
SLIDES: <number of slides>
<slide number> | <title> | <content> | <YES, MAYBE or NO> | <why>
```

Check before going on: the number of slides should match the deck. If a slide you know comes from the report is marked NO, tell Copilot in the chat (for example `Slide 6 is from the report`) and ask it to redo the list.

### Step 2. Describe each report slide (same chat)

```text
Keep the slide list. Now describe in detail every slide you marked YES or MAYBE.

For each such slide, fill in every line of the block below. Write "none" when a line does not apply and "cannot tell" when the slide does not show it.

Rules:
- Do not copy any figure, dealer name, customer name or customer number. Write N where a figure would go (for example "N units, up N% on last year"). MSA names, model names, size class names, month names and years are fine.
- Describe what is on the slide. Where you guess how a figure was worked out, start that line with GUESS:.
- Do not explain beyond the format. Reply only in the format below, one block per slide.

Reply in exactly this format:
SLIDE <number>: <title>
LAYOUT: <how the elements sit on the slide, e.g. chart left two thirds, text box right>
ELEMENT <n>: <chart / table / text box / picture>
  TYPE: <for a chart: clustered bar, stacked column, line, pie, ...; for a table: rows x columns>
  SHOWS: <the measure, e.g. units ordered, units in inventory, number of open orders>
  GROUPED BY: <what the categories or rows are, e.g. MSA, size class, product type, month>
  SERIES OR COLUMNS: <each series or column name exactly as written>
  FILTER: <what is left out or picked, e.g. CTL only, open orders only, cancelled left out>
  PERIOD: <the month, months or year shown, and any comparison, e.g. vs last year>
  LABELS: <axis titles, data labels, legend, units, sort order>
  TEXT: <for a text box: what the text says with figures replaced by N>
  LIKELY SOURCE: <Order Intake, Inventory by Month units table, Inventory by Month detail, size classes, or not in the report>
STYLE: <template or master name if shown, fonts, the colours used for CTL / MX / SSL if consistent>
NOTES: <anything else needed to rebuild this slide, e.g. one slide per MSA, a footnote, a logo>
```

Check before sending on: every YES and MAYBE slide has a block, and no block holds a real figure or a dealer or customer name. Delete any that slipped in.

### Step 3. Send it to Claude (you, not Copilot)

Copy both replies and paste them to Claude under this text:

```text
This is a description of last year's PowerPoint deck for the Open Orders - Order Intake - Inventory Analysis By MSA report, written by Copilot from the deck itself, with figures replaced by N. The report is built by the Excel macro modOpenOrdersReport (v2.4) in the gustjf/ai repo (source/open-orders-intake-inventory/build-open-orders-report.bas).

Task: plan the Update PowerPoint step (road map step 5 in open-orders-macro-manual.tex): which slides the macro can fill from the report and how (native PowerPoint charts updated in place on last year's deck), which slides need data the report does not have, and the questions I still need to ask. Answers already given on the deck are in open-orders-report-questions.tex (questions 14 to 18).

Slide list (P12 Step 1):
<paste>

Slide details (P12 Step 2):
<paste>
```

---

## What the Stronger AI review does with the answers

| Answer | Compared with | A mismatch means |
|---|---|---|
| P1 | The column names the macro looks for (macro-instructions.md, Section 5) | A renamed column: the macro would skip the file and say so in the Run Log |
| P2 | P2 for the other exports | The same MSA spelt two ways, so its figures would split across two rows |
| P3, P4 | The status order, the product types and the date handling | A new status or type to place, or text dates to check |
| P5 | The CTL= / MX= rule (sum of TOTAL for the month, by MSA) | The hand-built figures follow another rule. The anonymised sample suggests a check is worth it: its Austin tab shows CTL= 21, MX= 2, while its own inventory list sums to 59 and 45 |
| P6 | P9, the macro's own table | A bug to fix before the report is used |
| P11 | The Run Log of the first build (its UNMAPPED lines) | A model still missing from the size class file, or a slide name that does not match MODEL_GROUP |
| P7 | `size-class-master.xlsx` | Models to add to the size class list, which would otherwise be UNMAPPED (P11 Step 2 already does this when P11 is used) |
| P8 | P3, open orders per MSA | If the counts match, Open Order is just the open part of Order Intake and is not needed (question 3) |
| P10 | The expected Run Log lines | Files skipped, rows dropped, UNMAPPED models and warnings to deal with |
| P12 | Not compared: it goes to Claude, with the answers to questions 14 to 18 | Slides marked NO that do come from the report: correct them in the chat before sending |
