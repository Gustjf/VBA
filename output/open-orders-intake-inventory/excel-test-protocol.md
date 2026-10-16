# Open Orders Report Macro v2.8: Excel Test Protocol

This protocol checks macro version 2.8 on a Windows PC with Excel. It takes about 30 minutes and needs a web browser for GitHub. You set up a macro workbook, build three reports (the test files with two option sets, and the Austin sample), compare each with the expected values below, test the Config control panel, and send back what you saw. Follow the steps in order. If CHECK 1 or CHECK 2 fails, stop and send back what you saw. For any other difference, write it down and carry on.

Version 2.8 has run in LibreOffice (with the substitutions listed in `run_macro_libreoffice.py`) and builds exactly what the Python reference builds, but it has never run in Excel. The steps marked **(Excel only)** cover what LibreOffice could not show: the buttons, the message boxes, the date and hint checks on Config, gridlines, frozen panes, the value-list filters, the month drop-down, and parts of the new look (double and medium rules, the banding and the UNMAPPED peach, bold caption labels, the page setup for printing) and, new in 2.5, the Open Orders by MSA chart, which LibreOffice did not draw at all (CHECK 6). New in 2.6 is the visible total at the top of the Order Intake tab: LibreOffice Calc recalculated it correctly under real filters, and CHECK 7 confirms Excel does the same when you filter by MSA and size class. New in 2.7: the size class comes only from the size class workbook (17 classes in three families on File Format Table 1, the skid-steer frame classes among them); a model it does not list is UNMAPPED, so the tables on File Format sit three rows lower and the UNMAPPED rows below are new. New in 2.8: the allowed classes are read from the size class workbook's Set-up tab (one table each for CTL, MX and SSL), or are the 17 built-in ones when it has none; every Run Log gets a "Size classes:" line, and Run H tests the Set-up tab.

## Setup

1. Open the GitHub repository `Gustjf/AI`. Choose Code > Download ZIP and extract it into `C:\Temp`, which creates `C:\Temp\AI-main`. Below, "the source folder" means `C:\Temp\AI-main\source\open-orders-intake-inventory`. It holds the macro `build-open-orders-report.bas`, the folder `test-fixtures` (synthetic exports for all MSAs, a test size-class file `Size Class.xlsx`, and decoys the macro must ignore) and the folder `sample-data` (anonymised Austin exports and the hand-built target workbook). The size class master list is `C:\Temp\AI-main\output\open-orders-intake-inventory\size-class-master.xlsx`; for Run C and for your own runs keep a local copy next to the builder (step 2).
2. Create an empty folder `C:\Temp\oo-test`. Every run writes its report there. Copy `size-class-master.xlsx` from `C:\Temp\AI-main\output\open-orders-intake-inventory` into it: `C:\Temp\oo-test\size-class-master.xlsx` is the local master list you point `SizeClassFile` at.
3. In Excel, open a new blank workbook, choose File > Save As, set Save as type to "Excel Macro-Enabled Workbook (*.xlsm)", name it `Open Orders Report Builder.xlsm` and save it in `C:\Temp\oo-test`.
4. Press Alt+F11 (VBA editor). In Project Explorer select `VBAProject (Open Orders Report Builder.xlsm)`, choose File > Import File and pick `build-open-orders-report.bas` from the source folder. A module `modOpenOrdersReport` appears under Modules.
5. **CHECK 1 (compile).** Choose Debug > Compile VBAProject. Expected: nothing happens and no message appears. If a "Compile error" appears, write down the message and the highlighted line, stop, and send both back.
6. Close the VBA editor and save. Press Alt+F8, select `ShowOpenOrdersReportConfig`, click Run. **CHECK 2 (Config sheet).** Expected:
   - A message box: "The Config sheet is ready...".
   - Sheet `Config` holds nine settings in A2 to A10: `InputFolder`, `OutputFolder`, `OutputFileName`, `SizeClassFile`, `MSAColumn`, `InventoryMonth`, `ShowCancelled`, `ShowSSL`, `OpenAfterBuild`. The Value cells B2 to B10 are pale yellow. Row 12 has **five** buttons: Build report, Check setup, Choose input folder..., Choose output folder..., Choose size class file.... Below them are the "Last run" box (rows 14 to 19) and the "If something goes wrong" box (rows 21 to 26, five rows including "An MSA is missing or wrong").
   - There is **no** `SizeClassMap` sheet.
   - Press Alt+F8 again: only `ShowOpenOrdersReportConfig` is listed.

## Run A: test fixtures, default options

7. Set the Config values for Run A (type them in column B, on the row of each name):

| Setting | Run A value |
|---|---|
| InputFolder | `C:\Temp\AI-main\source\open-orders-intake-inventory\test-fixtures` |
| OutputFolder | `C:\Temp\oo-test` |
| OutputFileName | `fixtures-default.xlsx` |
| SizeClassFile | `C:\Temp\AI-main\source\open-orders-intake-inventory\test-fixtures\Size Class.xlsx` |
| MSAColumn | blank |
| InventoryMonth | blank |
| ShowCancelled | FALSE |
| ShowSSL | FALSE |
| OpenAfterBuild | TRUE |

8. Click **Build report**. Expected: no "not built" message; the status bar shows "Open Orders report: reading ..." then "writing ..."; a message "Report built in ... s" that names "MSAs: 4: (no MSA), Amarillo, Brownsville, Dallas-Fort Worth [North] Metroplex" and "Inventory month: September 2026". The report opens with five tabs, in this order: **File Format, Order Intake, Inventory by Month, Open Orders by MSA, Run Log**. On the builder's Config sheet the Last run box shows an amber "Report built - 12 items to look at", a link to `fixtures-default.xlsx`, the four MSAs, and 12 detail lines (5 WARNING: three size class file rows, two Order Intake files, the blank MSA; 7 RUN LOG: the two file lines with "dropped", the three (no MSA) lines, Amarillo UNMAPPED E25, T999, the TBD date).

9. **CHECK 3 (style, Excel only).** On all five tabs: no gridlines. The type is Cambria (the Run Log: Consolas). On File Format, Order Intake, Inventory by Month and Open Orders by MSA: A1 is a 14-point bold title, D1 a one-line italic note (on Order Intake the note is in H1, because F1 and G1 hold the visible total, CHECK 7), and a thicker rule runs under row 1. Tables have horizontal rules only, **no vertical lines**. Each header row is bold white text on green, with a thicker rule above and a double rule under it (Order Intake row 2; Inventory rows 4 and 11). The body rows are banded: every second visible row light green, the first row under each header white. A thicker rule sits above the first row of each size class (Order Intake) and each MSA (inventory detail), and a double rule above the Total row (Inventory row 9, not banded). Text sits one step in from the cell edge; dates and numbers are right-aligned, IS_OPEN_ORDER and PRODUCT_TYPE_ABBR centred. The UNMAPPED Size Class cells (Order Intake A30 and A31) are bold on peach; the month cell (Inventory B3) is yellow with a thin box. No other colours. On File Format the captions read "Table 1. Size classes, in report order." and so on, with only "Table 1." in bold; Table 1 has 17 classes in three families (Track loaders (CTL), Mini excavators (MX), Skid-steer loaders (SSL)), a thicker rule above the first MX and the first SSL row, and UNMAPPED last (A22, bold on peach, not banded). Order Intake is frozen below row 2 and Inventory by Month below row 11: scrolling down keeps the headers in view. **Banding and filters:** on Order Intake, clear the filter on ORDER_STATUS: the cancelled rows appear and the bands still alternate green and white with no two greens together; undo with Ctrl+Z (or close without saving). Then **print preview** (Ctrl+P) on Order Intake: landscape, all 11 columns across one page, green header and bands visible, the page header shows "Open Orders - Order Intake - Inventory Analysis By MSA" on the left and "Built <today>" on the right, the footer "Order Intake" on the left and "Page 1 of 1" in the centre; File Format previews on a single page. Close the preview without printing. Optional exact test: click A3 on Order Intake, then Home > Conditional Formatting > Manage Rules: two rules, first `=$A3="UNMAPPED"` (peach, Stop If True ticked, applies to `=$A$3:$A$31`), then `=MOD(SUBTOTAL(103,$A$3:$A3),2)=0` (";" instead of "," on a PC that uses ";"; light green, applies to `=$A$3:$K$31`). Close the dialog with Cancel.

10. **CHECK 4 (hidden rows, Excel only).** On Order Intake, the filter buttons on E2 (ORDER_STATUS) and H2 (PRODUCT_TYPE_ABBR) show the filter icon, and the row numbers jump: rows 4, 13, 18, 22, 25 and 29 are hidden (the two SSL rows S76 and S590, and the four cancelled CTL/MX rows, "Cancelled" in mixed case included). 23 of 29 rows are visible (status bar: "23 of 29 records found" or similar). G1 shows 24. Clear the filter on ORDER_STATUS (button > Clear Filter From "ORDER_STATUS"): rows 4, 18, 22, 25 appear and G1 changes to 28. Clear the filter on PRODUCT_TYPE_ABBR: rows 13 and 29 appear and G1 shows 30. Close the report without saving, then reopen `C:\Temp\oo-test\fixtures-default.xlsx`: the six rows are hidden again.

11. Check the report against Table A.

**Table A: Run A expected results**

| Item | Expected |
|---|---|
| Order Intake rows | A3:K31 (29 rows); AutoFilter on A2:K31; B2 reads MSA |
| Visible total | F1 reads `Total sum of ordered Quantity of visible selection:` (bold, right-aligned, running left over D1 and E1, nothing cut off) and G1 shows **24** right after it (bold, a little larger, left in its cell); click G1: `=SUBTOTAL(109,$G$3:$G$31)`. The italic note is in H1 |
| Order Intake order | A3 0<2000 T450 (Amarillo); A5:A10 class 2300<2600: T66-2 9/20/2026, T650 9/15/2026 (no MSA), T66-2 9/15/2026, T66-2 9/15/2026, T650 9/10/2026 (Brownsville), T64-2 7/1/2026; A14:A18 class 3000<3500: T740 8/12 ORDERED, T740 TBD ORDERED (C15 shows the text TBD), T740 8/5 Not yet processed, T770 8/20 INVOICED, T740 8/11 Cancelled; A29 Med Frame S590 (hidden, an SSL row); A30 UNMAPPED T999 (ORDERED), A31 UNMAPPED E25 (Not yet processed; not in the test size class file, so no class whatever its name) |
| PRODUCT_TYPE_ABBR | H23 shows MX (the export had " mx "); every H cell is CTL, MX or SSL |
| Inventory month cell | B3 shows September 2026, yellow and boxed; clicking it shows a drop-down with July 2026, August 2026, September 2026 |
| Units in table | A4 reads MSA; A5:E8 = (no MSA) 0 0 0 0; Amarillo 6 8 4 18; Brownsville 0 0 0 0; Dallas-Fort Worth [North] Metroplex 0 0 0 0. Row 9 Total 6 8 4 18 (bold). Click B6: `=SUMIFS($G$12:$G$27,$A$12:$A$27,$A6,$B$12:$B$27,"Compact Track Loaders",$F$12:$F$27,$B$3)` |
| Inventory detail | Header A11:I11 (MSA ... TOTAL, then ORGANIZATION_NAME in H and ORGANIZATION_CUSTOMER_NUMBER in I), rows A12:I27 (16 rows, Amarillo then Dallas; CTL, MX, SSL; newest date first). H12 = Bobcat of Amarillo North, I12 = 900102. Visible: rows 12, 13, 14, 16, 17, 18 (Amarillo CTL T66-2, T770, T450 and MX E35, E10, E20 of September 2026); the rest hidden by the filters on Month and PRODUCT_TYPE |
| File Format | Captions "Table 1. Size classes, in report order." (A3; A5:A21 the 17 classes, A22 UNMAPPED), "Table 2. Order status rank inside a size class." (A24), "Table 3. Default filters." (A33; five rows, the fifth A39:D39 `Open Orders by MSA`, `PRODUCT_TYPE_ABBR`, `CTL, MX on the chart`, `ShowSSL = FALSE`; the italic note in A40), "Table 4. This build." (A42; the item after MSAs is `Open orders`: `ORDERED and Not yet processed: 20 rows, 21 units (CTL 14, MX 6, SSL 1), on the Open Orders by MSA tab`); "Table 5. Months in this build." in F3 with July, August, September 2026 below |
| Open Orders by MSA | A1 title, A3 "Table 1. Open orders by MSA: ...", header A4:E4 MSA CTL MX SSL Total; A5:E8 = (no MSA) 1 0 0 1; Amarillo 8 5 1 14; Brownsville 1 0 0 1; Dallas-Fort Worth [North] Metroplex 4 1 0 5. Row 9 Total 14 6 1 21 (bold, a double rule above it). The chart starts at A11 (CHECK 6) |

12. **CHECK 5 (the month drop-down, Excel only).** On Inventory by Month, pick August 2026 in B3. Expected: B3 still shows the text "August 2026" (left-aligned; not "Aug-26", which would mean it became a date), and the table changes to Amarillo 6 2 1 9 and Dallas-Fort Worth [North] Metroplex 5 2 1 8, Total 11 4 2 17. (The detail rows stay filtered to September: the note in D1 says so.) Pick September 2026 again. If the numbers do not change, write down what B3 holds and what B6 shows.

12a. **CHECK 6 (Open Orders by MSA, Excel only).** Click the Open Orders by MSA tab.
   - The table is as in Table A. Click B6: the formula bar shows exactly `=SUMIFS('Order Intake'!$G$3:$G$31,'Order Intake'!$B$3:$B$31,$A6,'Order Intake'!$H$3:$H$31,B$4,'Order Intake'!$E$3:$E$31,"ORDERED")+SUMIFS('Order Intake'!$G$3:$G$31,'Order Intake'!$B$3:$B$31,$A6,'Order Intake'!$H$3:$H$31,B$4,'Order Intake'!$E$3:$E$31,"Not yet processed")` (";" instead of "," on a PC that uses ";"), and B6 shows 8.
   - D1 is the italic note ending "Chart: CTL and MX."
   - The chart starts at A11, two rows under the Total row, and does not cover the table. Its title reads "Open orders by product type and MSA (ordered quantity)". It has **two** bar groups, CTL and MX (no SSL), with **four** bars in each, one per MSA, and the legend at the bottom lists the four MSAs in table order: (no MSA) green, Amarillo grey, Brownsville blue, Dallas-Fort Worth [North] Metroplex rust. Every bar has its value on it (CTL 1, 8, 1, 4; MX 0, 5, 0, 1). The vertical axis is titled "Ordered quantity" and has horizontal gridlines only. The text is Cambria, only the title bold; the chart has no fill and a thin light-grey border.
   - Click the chart > Chart Design > Select Data: the chart range is `='Open Orders by MSA'!$A$4:$C$8` (or the four series listed separately, each on one row 5 to 8 with values from columns B and C), and "Switch Row/Column" shows the series are the rows. Click Cancel.
   - Print preview (Ctrl+P): one landscape page with the table and the whole chart, the same page header and footer as the other tabs (footer "Open Orders by MSA"). Close the preview.
   - If anything about the chart looks wrong (a series per product type instead of per MSA, a Total bar, overlapping labels, the wrong font), take a screenshot and send it back.

12b. **CHECK 7 (the visible total, Excel only).** In `fixtures-default.xlsx`, on Order Intake (filters as built, G1 shows 24):
   - Filter MSA (the button on B2) to Amarillo only. Expected: G1 shows **16** (15 rows; one T66-2 row has quantity 2).
   - Then filter Size Class (the button on A2) to 2300<2600 only. Expected: G1 shows **5** (four rows, three T66-2 and one T64-2: 1 + 2 + 1 + 1).
   - Select the G cells of the visible rows: the status bar's Sum is 5 too, the same figure G1 shows without selecting anything.
   - Scroll down: rows 1 and 2 (the total and the headers) stay in view.
   - Clear the MSA and Size Class filters (button > Clear Filter From ...): G1 shows 24 again. Close the report without saving.
   - If G1 does not change when a filter changes, press F9 and note whether it changes then (that would mean calculation is set to Manual on the PC: File > Options > Formulas).

13. Click the Run Log tab. A1 reads `Built <date time> from <folder> in <n> s`, row 3 is empty, and A2 and rows 4 to 27 match this text exactly:

```
Options: InventoryMonth=latest in file, ShowCancelled=FALSE, ShowSSL=FALSE, MSAColumn=auto, SizeClassFile=Size Class.xlsx
File Order Intake 2026.xlsx: 30 rows -> 28 written, 1 dropped (ATT 1), 1 blank ignored; MSA column MSA, 3 MSAs
File Order Intake Brownsville.xlsx: 2 rows -> 1 written, 1 blank ignored; MSA from file name (Brownsville)
File Inventory by Month 2026.xlsx: 17 rows -> 16 written, 1 dropped (Attachments 1); MSA column CITY, 2 MSAs
Size class file Size Class.xlsx: 20 usable rows, 3 warnings
Size classes: the 17 built into the macro (Size Class.xlsx has no Set-up tab)
Size class file row 21 T630: "2000-2300" is not a size class - row skipped
Size class file row 22 T650: listed again (first in row 20) - row 20 is used
Size class file row 24 E100: "Over10" is not a size class - row skipped
(no MSA): Order Intake 1 row (CTL 1, MX 0, SSL 0; 0 CANCELLED), 0 hidden by the filters
(no MSA): no inventory rows
(no MSA): Open orders 1 row (ORDERED 1, Not yet processed 0): CTL 1, MX 0, SSL 0 units
Amarillo: Order Intake 20 rows (CTL 11, MX 7, SSL 2; 4 CANCELLED), 5 hidden by the filters; UNMAPPED model groups: E25, T999
Amarillo: Inventory 10 rows; September 2026: CTL 6, MX 8, SSL 4 units
Amarillo: Open orders 13 rows (ORDERED 12, Not yet processed 1): CTL 8, MX 5, SSL 1 units
Brownsville: Order Intake 1 row (CTL 1, MX 0, SSL 0; 0 CANCELLED), 0 hidden by the filters
Brownsville: no inventory rows
Brownsville: Open orders 1 row (ORDERED 1, Not yet processed 0): CTL 1, MX 0, SSL 0 units
Dallas-Fort Worth [North] Metroplex: Order Intake 7 rows (CTL 4, MX 3, SSL 0; 1 CANCELLED), 1 hidden by the filters
Dallas-Fort Worth [North] Metroplex: Inventory 6 rows; September 2026: CTL 0, MX 0, SSL 0 units
Dallas-Fort Worth [North] Metroplex: Open orders 5 rows (ORDERED 4, Not yet processed 1): CTL 4, MX 1, SSL 0 units
File Open Orders 2026.xlsx: 4 rows, MSA column MSA, 2 MSAs - not used by this report
Order Intake 2026.xlsx ORDER_DATE: 6 text value(s) converted to dates, 1 not a date and left as is (row 25)
2 Order Intake files found, all are read: Order Intake 2026.xlsx, Order Intake Brownsville.xlsx
Order Intake 2026.xlsx: 1 row with a blank MSA - shown as (no MSA)
```

If the PC is not set to month-first dates, one more line may follow ("WARNING: text dates not in yyyy-mm-dd form ..."); note it and carry on. A line "WARNING: Excel's filters on the ... tab hid other rows than the rules ..." means the macro had to set hidden rows itself because Excel's filter did something else: the report is still right, but note it and send it back (it is the one thing this run is designed to catch).

## Run B: test fixtures, other options

14. Keep the Run A values but change four: `OutputFileName` to `fixtures-variant.xlsx`, `InventoryMonth` to `2026-08-15` (type it as year-month-day: Excel reads that form the same way whatever the regional settings; any day in August works), `ShowCancelled` to TRUE, `ShowSSL` to TRUE.
15. Click Build report. Same five tabs. Check Table B.

**Table B: Run B expected results**

| Item | Expected |
|---|---|
| Order Intake | The same 29 rows in the same order as Run A; **no** hidden rows and no filter icons (only the AutoFilter buttons); H1 ends "No rows filtered out."; G1 (the visible total) shows 30 |
| Month cell | B3 August 2026 |
| Units in table | (no MSA) 0 0 0 0; Amarillo 6 2 1 9; Brownsville 0 0 0 0; Dallas-Fort Worth [North] Metroplex 5 2 1 8; Total 11 4 2 17 |
| Inventory detail | The same 16 rows; visible: rows 15, 19, 21, 22, 23, 25, 27 (the August rows of Amarillo T66-2, E88, S76 and Dallas T740, T740, E55, S650); the July and September rows hidden by the Month filter; no filter on PRODUCT_TYPE |
| Open Orders by MSA | The table exactly as in Run A (14 6 1 21; the table does not depend on the options). D1 ends "Chart: CTL, MX and SSL." The chart has **three** bar groups, CTL, MX and SSL (Amarillo 1 in SSL, the others 0), four bars in each. File Format A39 reads `CTL, MX, SSL on the chart` and `ShowSSL = TRUE` |

16. The Run Log matches Run A's text except these lines (A2 and rows 15, 16, 21, 22); the open orders lines are the same as in Run A:

```
Options: InventoryMonth=August 2026, ShowCancelled=TRUE, ShowSSL=TRUE, MSAColumn=auto, SizeClassFile=Size Class.xlsx
Amarillo: Order Intake 20 rows (CTL 11, MX 7, SSL 2; 4 CANCELLED), 0 hidden by the filters; UNMAPPED model groups: E25, T999
Amarillo: Inventory 10 rows; August 2026: CTL 6, MX 2, SSL 1 units
Dallas-Fort Worth [North] Metroplex: Order Intake 7 rows (CTL 4, MX 3, SSL 0; 1 CANCELLED), 0 hidden by the filters
Dallas-Fort Worth [North] Metroplex: Inventory 6 rows; August 2026: CTL 5, MX 2, SSL 1 units
```

## Run C: Austin sample

17. Change these values: `InputFolder` to `C:\Temp\AI-main\source\open-orders-intake-inventory\sample-data`, `OutputFileName` to `sample.xlsx`, `SizeClassFile` to `C:\Temp\oo-test\size-class-master.xlsx` (the local copy from step 2; use Choose size class file...), `InventoryMonth` to `2026-11-08`, `ShowCancelled` FALSE, `ShowSSL` FALSE.
18. Click Build report. Expected: the Last run box shows an amber "Report built - 1 item to look at": 1 RUN LOG line, the Austin UNMAPPED model groups (as in step 19). The master list loads with no warnings. That line is expected with the 2026-10-15 master list until the eight model groups it names are added to it. Check Table C.

**Table C: Run C expected results**

| Item | Expected |
|---|---|
| Order Intake | A3:K280 (278 rows, all MSA = Austin); 108 hidden (68 SSL and 40 cancelled CTL/MX), 170 visible; G1 (the visible total) shows 170, `=SUBTOTAL(109,$G$3:$G$280)` (every quantity in the sample is 1). Filter Size Class to 2600<3000: G1 shows 6; to >03-04 instead: 16; to UNMAPPED instead: 15 (the E55, E60 and E88 rows, last on the tab). Clear the filter again |
| Month cell | B3 November 2026; the drop-down lists April 2026 to December 2026 |
| Units in table | Austin 57 43 7 107; Total 57 43 7 107 |
| Inventory detail | Header A8:I8, rows A9:I307 (299 rows, with ORGANIZATION_NAME and the customer number in H and I); visible: the 29 CTL and MX rows of November 2026 (18 CTL, then 11 MX) |
| Open Orders by MSA | Austin 86 34 43 163; Total 86 34 43 163 (163 open rows: 104 ORDERED, 59 Not yet processed, all quantity 1). The chart starts at A8 with **one** series, Austin (green), in two bar groups CTL 86 and MX 34 |

19. The Run Log: A2 and rows 4 to 12 match exactly:

```
Options: InventoryMonth=November 2026, ShowCancelled=FALSE, ShowSSL=FALSE, MSAColumn=auto, SizeClassFile=size-class-master.xlsx
File Order Intake Austin_ANONYMIZED.xlsx: 278 rows -> 278 written; MSA from file name (Austin)
File Inventory by Month Austin_ANONYMIZED.xlsx: 299 rows -> 299 written; MSA from file name (Austin)
Size class file size-class-master.xlsx: 29 usable rows, 0 warnings
Size classes: 17 from the Set-up tab (CTL 6, MX 7, SSL 4)
Austin: Order Intake 278 rows (CTL 160, MX 50, SSL 68; 44 CANCELLED), 108 hidden by the filters; UNMAPPED model groups: E60, E55, E88, T86, S510, T64, T66, S66
Austin: Inventory 299 rows; November 2026: CTL 57, MX 43, SSL 7 units
Austin: Open orders 163 rows (ORDERED 104, Not yet processed 59): CTL 86, MX 34, SSL 43 units
File Open Order Austin_ANONYMIZED.xlsx: 167 rows, MSA from file name (Austin) - not used by this report
Open orders: 2 rows where IS_OPEN_ORDER disagrees with the status (TRUE with another status, or FALSE with ORDERED / Not yet processed)
```

With the 2026-10-15 master list the SSL rows carry their frame classes (EX Small, Small Frame, Med Frame, Large Frame) and are still hidden by the SSL filter, and the eight model groups the list does not have are UNMAPPED: the T64, T66 and T86 rows are all cancelled and the S510 and S66 rows are SSL, so all of those are hidden, while 15 visible rows (E55, E60, E88) are UNMAPPED and sit at the end of the tab. A class never changes a count, so the 170 visible rows and every figure in Table C are the same as before. The last line is information only (it does not turn the Last run box amber): in the sample one Completely processed row has IS_OPEN_ORDER TRUE and one Not yet processed row has it FALSE. The tab follows ORDER_STATUS.

20. **Side-by-side check.** Open `sample-data\001. Open Orders - Order Intake - Inventory Analysis By MSA_ANONYMIZED_TRIMMED.xlsx` (do not edit or save it) and use View > View Side by Side with `sample.xlsx`. Expected:
    - The 170 visible Order Intake rows are the hand-built tab's 170 rows. Their size classes are in the same order except for the 15 E55, E60 and E88 rows: the hand-built tab classes them >05-06 (E55, E60) and >06-10 (E88); the report shows them UNMAPPED at the end, because the master list does not have them yet. Once they are added with those classes, the order matches (checked in the cloud with a copy of the list extended by those three rows). Inside a class the rows are now grouped by order status (ORDERED, Not yet processed, INVOICED, Completely processed) before the date, so many sit in a different position: that is the agreed new order, not an error.
    - The visible November inventory rows are the hand-built tab's 29 rows, in the same order. Three TOTAL values differ by 1 or 2 (anonymisation noise in the sample).
    - The hand-built CTL= 21 and MX= 2 are not the 57 and 43 of the new table: the anonymised sample cannot reproduce them (its own inventory rows add up to 59 and 45). CTL= and MX= are the sum of TOTAL (answered 2026-10-12); prompt P5 in `copilot-prompts.md` checks this on the real hand-built report.

## Run D: rebuild with the report open

21. Leave `sample.xlsx` open, switch to the builder workbook (Run C values) and click Build report. Expected: a question ending "already exists. Overwrite it?". Choose Yes. The report builds again without an error.

## Run E (optional): export already open

22. Open `test-fixtures\Order Intake 2026.xlsx` in Excel (click Enable Editing if it opens in Protected View). Set the Run A values from step 7 with `OutputFileName` = `fixtures-default-2.xlsx` and click Build report. Expected: the build matches Run A, and the export is still open afterwards.

## Run F: the Config control panel (Excel only)

23. Set the Run A values from step 7 again. Click **Check setup**. Expected: a message starting "Ready to build." that lists "MSAs: 4: (no MSA), Amarillo, Brownsville, Dallas-Fort Worth [North] Metroplex", "Inventory months: July 2026, August 2026, September 2026 (the report shows September 2026)", the report path, and 6 warnings (three size class file rows, two Order Intake files, the blank MSA, and "Models in Order Intake that are not in the size class file (UNMAPPED until added): E25, T999"). The Last run box: "Setup check: ready to build (6 warnings)" in amber. No file is written.
24. Click B9 (`ShowSSL`): a hint appears and a drop-down offers TRUE and FALSE (if the drop-down shows one item "TRUE,FALSE", the PC uses ";" as list separator: note it and send it back). Type `maybe` and press Enter: Excel refuses it ("Pick TRUE or FALSE from the list."); click Cancel. Click B7 (`InventoryMonth`) and type `next week`: refused ("Type a date such as 8/1/2026..."); click Cancel.
25. Click C2 and try to type: Excel says the cell is on a protected sheet.
26. Set `SizeClassFile` (B5) to `C:\Temp\no-such-file.xlsx` and click **Build report**. Expected: "The report was not built. Fix this on the Config sheet..." with the line "SizeClassFile not found: C:\Temp\no-such-file.xlsx. Click Choose size class file."; the Last run box turns red, "Not built: 1 problem to fix"; no new workbook opens.
27. Click **Choose size class file...**. Expected: a file picker showing Excel workbooks. Pick `test-fixtures\Size Class.xlsx` and click OK: B5 now holds its full path.
28. Click **Choose input folder...**, pick the `sample-data` folder, OK: B2 holds that path. Set B2 back to the test-fixtures path.
29. Set `MSAColumn` (B6) to `MSA` and click Check setup. Expected: 2 problems: "MSAColumn "MSA" is not a header of Inventory by Month 2026.xlsx (headers: ...)" and the same for Order Intake Brownsville.xlsx (those two test files have no MSA column on purpose; your real exports all have one). Clear B6.
30. Click the link in the Last run box's "Last report" row: the last report opens.

## Run G: upgrade from a version 1.3 Config

31. In the builder workbook: Review > Unprotect Sheet on Config, then right-click the Config tab > Delete. Insert a new sheet (the + after the last tab), rename it `Config`, and type this version 1.3 table into A1:B10 (header in row 1):

| A | B |
|---|---|
| Setting | Value |
| InputFolder | `C:\Temp\AI-main\source\open-orders-intake-inventory\test-fixtures` |
| OutputFolder | `C:\Temp\oo-test` |
| OutputFileName | `upgraded.xlsx` |
| InventorySnapshotDate | `2026-08-08` |
| DropCancelled | FALSE |
| IncludeSSL | TRUE |
| HeaderCountMode | InventoryUnits |
| InventoryHeaderRow | 7 |
| OpenAfterBuild | TRUE |

32. Press Alt+F8 and run `ShowOpenOrdersReportConfig`. Expected: the message ends "Upgraded to v2.8: HeaderCountMode and InventoryHeaderRow are no longer used; set SizeClassFile."; rows 2 to 10 now read InputFolder, OutputFolder, OutputFileName (values kept), SizeClassFile (blank), MSAColumn (blank), InventoryMonth 8/8/2026, ShowCancelled **TRUE** (DropCancelled FALSE turned around), ShowSSL TRUE, OpenAfterBuild TRUE; the panel, five buttons and both boxes are back; the Last run box shows the upgrade note.
33. Click Check setup. Expected: 1 problem, "SizeClassFile is blank. Click Choose size class file...". Set B5 with the button to `test-fixtures\Size Class.xlsx` and click Build report: the report builds (`upgraded.xlsx`, month August 2026, nothing hidden in Order Intake).

## Run H: the Set-up tab (allowed size classes, v2.8)

34. Set the Run A values from step 7 again, with `OutputFileName` = `fixtures-setup.xlsx` and `SizeClassFile` = `C:\Temp\AI-main\source\open-orders-intake-inventory\test-fixtures\Size Class Set-up.xlsx` (use Choose size class file...). That test file has a Set-up tab whose three tables differ from the usual 17 classes on purpose: `>00-01` is missing, `>10-15` and `XL Frame` are added, one table has a blank cell, one class has spaces around it, one class is listed twice and one cell says "unmapped". Click **Check setup**. Expected: "Ready to build." with 9 warnings, among them `Set-up tab C9: "2000<2300" is listed again (first in A6) - ignored`, `Set-up tab C10: "unmapped" is not a size class name (UNMAPPED marks models not in the file) - ignored`, `Size class file row 11 E10: ">00-01" is not a size class - row skipped`, and "Models in Order Intake that are not in the size class file (UNMAPPED until added): E25, T999, E10".
35. Click **Build report**. Expected: the Last run box shows "Report built - 15 items to look at". In the report:
    - File Format Table 1 lists 18 classes: the six CTL classes, then `>01-02` to `>06-10` and `>10-15` (no `>00-01`), then EX Small, Small Frame, Med Frame, Large Frame and XL Frame; a thicker rule above `>01-02` and above EX Small; UNMAPPED in A23, bold on peach. Because Table 1 is one row longer, Table 2 starts in A25, Table 3 in A34 and Table 4 in A43.
    - Order Intake: A27 `>10-15` E88; A28 Med Frame S590 (hidden); A29:A31 UNMAPPED E10, T999, E25. Rows 4, 13, 18, 21, 24 and 28 are hidden; G1 shows 24.
    - The Run Log, A2 and rows 4 to 30, matches exactly:

```
Options: InventoryMonth=latest in file, ShowCancelled=FALSE, ShowSSL=FALSE, MSAColumn=auto, SizeClassFile=Size Class Set-up.xlsx
File Order Intake 2026.xlsx: 30 rows -> 28 written, 1 dropped (ATT 1), 1 blank ignored; MSA column MSA, 3 MSAs
File Order Intake Brownsville.xlsx: 2 rows -> 1 written, 1 blank ignored; MSA from file name (Brownsville)
File Inventory by Month 2026.xlsx: 17 rows -> 16 written, 1 dropped (Attachments 1); MSA column CITY, 2 MSAs
Size class file Size Class Set-up.xlsx: 19 usable rows, 6 warnings
Size classes: 18 from the Set-up tab (CTL 6, MX 7, SSL 5)
Set-up tab C9: "2000<2300" is listed again (first in A6) - ignored
Set-up tab C10: "unmapped" is not a size class name (UNMAPPED marks models not in the file) - ignored
Size class file row 11 E10: ">00-01" is not a size class - row skipped
Size class file row 21 T630: "2000-2300" is not a size class - row skipped
Size class file row 22 T650: listed again (first in row 20) - row 20 is used
Size class file row 24 E100: "Over10" is not a size class - row skipped
(no MSA): Order Intake 1 row (CTL 1, MX 0, SSL 0; 0 CANCELLED), 0 hidden by the filters
(no MSA): no inventory rows
(no MSA): Open orders 1 row (ORDERED 1, Not yet processed 0): CTL 1, MX 0, SSL 0 units
Amarillo: Order Intake 20 rows (CTL 11, MX 7, SSL 2; 4 CANCELLED), 5 hidden by the filters; UNMAPPED model groups: E25, T999, E10
Amarillo: Inventory 10 rows; September 2026: CTL 6, MX 8, SSL 4 units
Amarillo: Open orders 13 rows (ORDERED 12, Not yet processed 1): CTL 8, MX 5, SSL 1 units
Brownsville: Order Intake 1 row (CTL 1, MX 0, SSL 0; 0 CANCELLED), 0 hidden by the filters
Brownsville: no inventory rows
Brownsville: Open orders 1 row (ORDERED 1, Not yet processed 0): CTL 1, MX 0, SSL 0 units
Dallas-Fort Worth [North] Metroplex: Order Intake 7 rows (CTL 4, MX 3, SSL 0; 1 CANCELLED), 1 hidden by the filters
Dallas-Fort Worth [North] Metroplex: Inventory 6 rows; September 2026: CTL 0, MX 0, SSL 0 units
Dallas-Fort Worth [North] Metroplex: Open orders 5 rows (ORDERED 4, Not yet processed 1): CTL 4, MX 1, SSL 0 units
File Open Orders 2026.xlsx: 4 rows, MSA column MSA, 2 MSAs - not used by this report
Order Intake 2026.xlsx ORDER_DATE: 6 text value(s) converted to dates, 1 not a date and left as is (row 25)
2 Order Intake files found, all are read: Order Intake 2026.xlsx, Order Intake Brownsville.xlsx
Order Intake 2026.xlsx: 1 row with a blank MSA - shown as (no MSA)
```

36. (Excel only, the drop-down.) Open `C:\Temp\oo-test\size-class-master.xlsx` (your local master list). On its **Set-up** tab, type `XL Frame` in the first empty yellow cell under "SSL size class" (C9). Table 4 now ends with XL Frame. On the **Size Class** tab, click D2 and open the drop-down: XL Frame is the last entry. Type `XXL Frame` in D2 and press Enter: Excel refuses it ("Not a size class"); click Cancel. Close the workbook **without saving**.

## Pass criteria

| Check | Pass when |
|---|---|
| CHECK 1, compile | No message after Debug > Compile VBAProject. |
| CHECK 2, Config | Nine settings in A2:A10 as listed, five buttons, both boxes, no SizeClassMap sheet, only ShowOpenOrdersReportConfig in Alt+F8. |
| CHECK 3, style | No gridlines, Cambria; horizontal rules only; green headers with white text and a double rule under each; light-green bands on visible rows, also after clearing a filter; UNMAPPED A30 and A31 bold on peach; month cell yellow; frozen headers; print preview landscape, one page wide, with the page header and footer. |
| CHECK 4, hidden rows | Rows 4, 13, 18, 22, 25, 29 hidden behind the two filters; clearing them shows the rows; still hidden after reopening. |
| CHECK 5, month drop-down | Picking August 2026 changes the table to the Run B figures. |
| CHECK 6, Open Orders by MSA | The table and B6 formula as given; one chart under the table, one series per MSA, two groups (three in Run B), labels and title as described; one printed page. |
| CHECK 7, visible total | G1 shows 24, then 16 (Amarillo), 5 (Amarillo, 2300<2600), the status bar Sum agrees, row 1 stays in view, 24 again after clearing; 28 and 30 in CHECK 4. |
| Run A, B, C | Tables A, B, C match and the Run Log text matches exactly; Run C's Last run box shows the 1 item of step 18. |
| Run C, side by side | Only the differences described in step 20 (the 15 UNMAPPED E55, E60, E88 rows included). |
| Run D, E | The overwrite question appears and Yes builds; with an export open, the build matches Run A. |
| Run F | Steps 23 to 30 behave as described. |
| Run G | The upgrade converts the values as in step 32 and the build then runs. |
| Run H, Set-up tab | Check setup and the build as in steps 34 and 35 (Table 1 with 18 classes, Tables 2 to 4 one row lower, E10 UNMAPPED, the Run Log exactly); the drop-down follows the Set-up tab and refuses other entries (step 36). |

## What to send back

1. Your Excel version (File > Account > About Excel), the Windows regional format (Settings > Time & language > Language & region), and the results of CHECK 1 and CHECK 2. For a failure, the exact message and the step number.
2. After each of Runs A, B, C and H, click the Run Log tab, select A1:A35, copy, and paste the text into your reply. List every table row or check that did not match, with what you saw, and a screenshot of the Open Orders by MSA tab of each run and of the top of the Order Intake tab after CHECK 7's two filters. Say whether Runs D to G passed; give the step number of anything that differed.
3. Any error message, word for word, with the step number.
4. If you can, upload the four reports (`fixtures-default.xlsx`, `fixtures-variant.xlsx`, `sample.xlsx`, `fixtures-setup.xlsx`) to `source/open-orders-intake-inventory/excel-run/` in the `Gustjf/AI` repository (Add file > Upload files, commit directly to main; create `excel-run/.gitkeep` first if the folder does not exist). The maintainer then runs, from `source/open-orders-intake-inventory`:

```
python3 check_report.py excel-run/fixtures-default.xlsx --expected test-fixtures/expected.json
python3 check_report.py excel-run/fixtures-variant.xlsx --expected test-fixtures/expected.json
python3 check_report.py excel-run/sample.xlsx
python3 check_report.py excel-run/fixtures-setup.xlsx --expected test-fixtures/expected.json
python3 compare_with_handbuilt.py excel-run/sample.xlsx
```

## Troubleshooting

- **Macros are disabled.** Click Enable Content on the yellow bar. A red "Security Risk" bar with no button means the file is marked as downloaded: close it, right-click it > Properties > tick Unblock > OK, reopen. If macros stay blocked, ask IT to add `C:\Temp\oo-test` as a Trusted Location.
- **"Expected: end of statement", or the whole module on one line after import.** The `.bas` lost its Windows line endings. Remove the module (right-click > Remove, No to export), download the file again with GitHub's "Download raw file" button, and repeat step 4.
- **A file is open in another program, or OneDrive is still syncing.** Close it everywhere, wait for the sync, click Build report again.
- **A folder picker opens when you click Build report.** InputFolder is empty: cancel, type the path, build again.
- **"SizeClassFile not found" although the file is there.** The path in B5 has a typo or points to a OneDrive file that is "online only": right-click it in File Explorer > Always keep on this device, or use Choose size class file....
