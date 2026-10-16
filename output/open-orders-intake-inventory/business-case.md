# Open Orders Report Automation: Business Case

The open orders report for the Bobcat dealer group is built by hand, MSA by MSA, and the MSA tabs have not always been treated the same way. An Excel macro can build the full report from the same exports in seconds. The decision for dealer operations is whether the team adopts it. We recommend a first run in the team's own Excel, then adoption once that run passes and the few open questions below are settled. The report is expected to be built about once a year, so the main gain is a correct, repeatable report rather than hours saved.

## The current and proposed process

The report `001. Open Orders - Order Intake - Inventory Analysis By MSA` is built by hand from the Order Intake, Inventory by Month and Open Order exports. For each MSA (metropolitan statistical area), the analyst filters to compact track loaders (CTL) and mini excavators (MX), handles cancelled orders, assigns a size class to each order line from its model group, sorts and shades each size-class block, copies one inventory month across and fills in the CTL and MX counts.

Under the proposed process (macro version 2.4, after the manager's review of 2026-10-12 and the follow-up answers), the analyst exports the three files for all MSAs into one folder and clicks "Build report" in an Excel workbook. One build covers every MSA: one Order Intake tab with an MSA column, grouped by size class, then order status, then date; one Inventory by Month tab with every inventory month and, for the month picked in its drop-down, units in stock (the sum of TOTAL) by MSA; a File Format tab that states the rules; and a Run Log. The workbook is set like a printed engineering report (green headers, banded rows, no vertical lines, page numbers) and prints one page wide. Size classes come from one master list, a size class workbook the team owns. Cancelled orders and skid-steer loaders stay in the report, hidden by default filters, so the data is there when it is needed. A Python version builds the same report where macros are blocked.

## Time per run

| Step | Today (minutes) | Automated (minutes) |
|---|---|---|
| Export from Snowflake | ____ | ____ |
| Build the report | ____ | under 1 |
| Review and fix | ____ | ____ |
| Total | ____ | ____ |

MSAs in the report: ____ Runs per year: about 1.

Fill in your own figures; hours saved per year = `(today total - automated total) x runs per year / 60`.

## Risks removed

| Error today | How the automated report prevents it |
|-----------------|-------------------------|
| Cancelled orders treated differently between MSAs | One rule for every MSA; cancelled orders are kept and hidden by one filter |
| Filters that differ between MSAs | One build applies the same filters to every MSA, stated on the File Format tab |
| Rows wrongly classed | Size class comes from one master list; models missing from it are flagged and listed |
| Inventory copied from the wrong month | Every month is kept; one setting or the tab's drop-down picks the month counted, shown on the tab and in the Run Log |
| New models silently left out | Models missing from the master list are flagged in the report, the Run Log and the setup check |
| A stray or old export used | Every file read is listed with its row count; a second file of a kind is flagged |

## Evidence and effort

The logic was checked against the hand-built Austin tab: every order line and every inventory line was reproduced, in the same size-class order, apart from small differences caused by the anonymisation of the sample. It was also tested on synthetic data with known answers, where all automated checks pass, and the macro itself, run in LibreOffice, produced reports identical to the Python version. The first run in the team's own Excel is still to do, and a short test protocol covers it.

The build was a proof of concept, a hardening pass, a second version after the manager's review and a small update for the answers that followed, each done in one working session with an AI assistant (Stronger AI review) under the analyst's direction. It needs no new software or licences, only a macro-enabled Excel workbook; Python is optional. Analyst time spent on the build: ____ hours.

## Open questions and next steps

Most decisions were settled on the 2026-10-12 call and the answers that followed: the exports' location column is the MSA; CTL and MX are the sum of TOTAL for the month, counted by MSA, with a date anywhere in a month counted in that month; the order status order stands; the organisation name and customer number stay on the inventory tab for now; and the size class workbook stays a local file that the builder is pointed at. One question is still open: whether the Open Order export is needed at all. Until it is answered, the macro lists the file and does not use it. The next steps raise a few more, listed in the questions PDF (Rev. 4): how boundary models are classed on the size class list, real exports for the Excel test, and what the PowerPoint deck shows and what IT allows.

Next steps: the Excel test run, with short AI prompts to check the latest exports; the team's own size class list replacing the starter list; feeding the PowerPoint deck from the report; and handing the builder workbook and the one-page manual to the person who runs it.
