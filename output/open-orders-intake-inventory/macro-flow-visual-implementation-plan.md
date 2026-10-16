# Open Orders Macro: Build-Flow Visual, Implementation Plan for the Stronger AI Review

This file is an **architect's brief**. It was written by one AI session (2026-10-12) for a later **Stronger AI review** session (the "implementer") that builds one deliverable: a one-page, landscape PDF drawn in TikZ that shows how the Open Orders report macro (v2.4) runs one build, with its traceability and error-proofing features on the page. The owner uses the page to talk their manager through the macro in about five minutes, just before running the report for the first time. The owner runs the implementer with the prompt at the very end (Section 11).

Nothing else in the repo changes. The macro, the manual, the business case and the scripts are read, not edited.

---

## 1. Mission and deliverables

Make a page that a manager or director can read in a minute and explain to someone else in two: what goes in, what the macro refuses to build, what it writes down, and what it was tested against. The page states facts and shows the mechanism; it draws no conclusions for the reader. It must not look like the AI-made visuals they see every day (Section 6 says what that means).

| Deliverable | Path | Notes |
|---|---|---|
| The drawing (source) | `source/open-orders-intake-inventory/open-orders-macro-flow.tex` | One TikZ picture on one landscape US Letter page; compiles with `./build.sh`. |
| The drawing (PDF) | `output/open-orders-intake-inventory/open-orders-macro-flow.pdf` | Exactly one page, 792 x 612 pt. Mirrors the source path (repo rule). |
| Map rows | `MAP.md` | One row for the `.tex`, one for the `.pdf`, same commit. Mark this plan's row "Executed" in the same commit. |

Not deliverables: no changes to the `.bas`, the manual, the business case, the instructions, the scripts or the fixtures. No second page, no slides, no Markdown copy of the page.

---

## 2. Standing rules

1. Read the repo rules file and `MAP.md` first and obey them: push directly to `main`, no pull request, update `MAP.md` in the same commit, source under `source/open-orders-intake-inventory/`, output under `output/open-orders-intake-inventory/`.
2. **Do everything yourself.** This is one file. The only delegation allowed is one optional **second AI reviewer** at the end (Section 9, step 6), who sees the rendered PNG and the checklist in Section 6, nothing else.
3. **Every statement on the page comes from Section 5 of this brief**, and every fact in Section 5 is traced to a source file in Section 3. If a wording in Section 5 is wrong against the source, the source wins and the page follows the source; say so in your final report. Do not add facts of your own, however true.
4. **No conclusions on the page.** The test for each sentence: it names something the macro does, something that happens in a given case, or a test result. "Saves time", "removes errors", "reliable", "consistent" and the like fail the test. Section 6 has the banned list.
5. Facts beat decoration. When the page is too full, cut by the order in Section 5.6, never by shrinking type below 7 pt.
6. Report faithfully: what compiled, what the checks said, what you cut, what you could not verify.

---

## 3. Read first, then stop reading

| Read | Why | How much |
|---|---|---|
| `CLAUDE.md`, `MAP.md` | Repo rules; where things are. | All (short). |
| `output/open-orders-intake-inventory/macro-instructions.md` | The source of nearly every fact on the page: Section 3 (each run), 5.7 (File Format tab and Run Log), 6 (how it was validated), 7 (decisions), 9 (troubleshooting). | Sections 3, 5.7, 6, 7, 9. Skip the rest. |
| `output/open-orders-intake-inventory/business-case.md` | The by-hand process in one paragraph; the "Risks removed" table (facts, not its verdicts). | All (one page). |
| `source/open-orders-intake-inventory/build-open-orders-report.bas` | The header comment (lines 1 to 60) states the rules; `BuildOpenOrdersReport` (from line 215) shows the check-then-build order; `CheckSetup` (from line 419) lists what stops a build and what only warns; `WriteFileFormat` (from line 2055) lists what Table 4 "This build" records. | Those four places only. Do not read the 2,800 lines. |
| `source/open-orders-intake-inventory/test-fixtures/expected.json` | Real Run Log lines (`runs[0].log_lines`), quoted verbatim on the page. | The `log_lines` list of run 0 only. |
| `source/open-orders-intake-inventory/open-orders-macro-manual.tex` | The one-page manual the manager already has: same colours, same author, same series name ("Strategy"); the page must not repeat it. | All (short). |
| `Specs/exam-style-latex-spec.md` | The house conventions this page borrows from (header, Rev. and date, author, hyperref metadata) and departs from (it is a drawing, not an exam sheet). | Skim. |
| `build.sh` | How the PDF is built into the mirrored tree. | All (short). |

Do not open the sample workbooks, the fixtures' `.xlsx` files, the POC report or the Python scripts: nothing on the page comes from them, and the numbers they would give are already in `macro-instructions.md` Section 6.

---

## 4. Who reads it, and how the five minutes go

The reader is the owner's manager (and, second-hand, a director). They know the report, the exports and the by-hand process; they have seen the one-page manual. They have not seen the macro run. They are used to AI-made slides and will discount anything that looks like one.

The owner talks while the manager looks. The page is read in this order, and the layout (Section 7) is built around it, so the five minutes need no pointing:

1. Title and the two-line strip under it: what the by-hand process was per MSA, and what one build does instead. (30 s)
2. The flow, stations 1 to 5, left to right: exports and the master list in; the Config sheet; the setup check and its stop; the build rules; the workbook. (2 min)
3. The "Stopped or flagged" panel: the cases the macro refuses or flags, and what the user sees. (1 min)
4. The "Written down" panel: what every report records about itself, with a real Run Log line. (1 min)
5. The "Tested" strip and the title block: what was checked, what is still to run, the one open question. (30 s)

The manager draws the conclusions. The page gives them the facts in the order that makes those conclusions easy to draw and easy to repeat upstairs.

---

## 5. The content, fixed (every word on the page)

Word budget for the whole page: **at most 300 words** (`pdftotext file.pdf - | wc -w`), not counting the title block. Wording below is final unless a source contradicts it (Section 2, rule 3). Keep sentence case and the terms the manual uses: MSA, size class, master list, Config, Build report, Check setup, Last run box, Run Log, UNMAPPED, CTL, MX, SSL.

### 5.1 Title and strip

- Title: **Open Orders report macro** with the subtitle **One build, from the exports to the workbook**. Small, top left, like a drawing's heading, not a centred slide title.
- Strip (two lines, plain text, directly under the title; the only "before and after" on the page):
  - *By hand, for each MSA:* filter to CTL and MX, handle cancelled orders, assign a size class per model, sort and shade each class, copy one inventory month across, fill in the CTL and MX counts. (Source: business-case.md, "The current and proposed process".)
  - *One build, all MSAs:* export the three files to one folder, click Build report, read the Run Log. About one run a year. (Source: macro-instructions.md Section 3; business-case.md.)

### 5.2 The flow: five stations and one gate

Boxes left to right, numbered 1 to 5 in the box header, arrows between them. Text inside each box is the list below, no more.

1. **Exports** — "Three Snowflake exports for all MSAs, in one folder": `Order Intake*.xlsx` (required), `Inventory by Month*.xlsx`, `Open Order*.xlsx` (listed, not used). A second box beneath, joined to station 4 by its own arrow: **Size class workbook** — "The team's master list: model group to size class. A new model is a row here, not a code change." (Sources: .bas header; macro-instructions.md 3.1 and 5.4.)
2. **Config sheet** — "Nine settings, five buttons. Build report is the only way to build; Alt+F8 lists one macro." (Source: .bas header, "HOW TO RUN"; manual Table 1.)
3. **Setup check** — "Runs before every build. Reads every file; checks the settings, both folders, the master list, each export's columns; lists the MSAs and months found." Below it a gate (a diamond or a split arrow) labelled **Problems?** with the branch down to a flat box: **Not built.** "One list of what to fix, in a message and in the Last run box. Nothing is created." Side note on the gate, smaller: "Warnings do not stop a build: two files of one kind, rows with a blank MSA, models not on the master list." (Sources: `CheckSetup`; macro-instructions.md 3.3 and 3.4.)
4. **Build** — header "Same rules for every MSA": "Every file of a kind merged; MSA from the MSA column, else the file name. Keep CTL, MX, SSL. Size class from the master list; not on it: UNMAPPED, bold on peach, and listed. Sort: size class, order status, date, MSA. Cancelled and SSL rows kept, hidden by filters." (Source: .bas header, "BUSINESS RULES".) Typeset the word UNMAPPED on the peach fill, bold, exactly as the report shows it.
5. **Workbook** — "Four tabs": File Format (the rules, and this build), Order Intake (all MSAs), Inventory by Month (units by MSA for the month in the drop-down; every month's detail), Run Log. Beneath: "Last run box on Config: Report built, or n items to look at." A return arrow from here back to the size class workbook box, labelled "UNMAPPED: add the model to the list, build again." (Sources: .bas header; macro-instructions.md 3.5.)

### 5.3 Panel "Stopped or flagged" (error-proofing)

A two-column table in a boxed panel, header row "Case" / "What happens". Rows, in this order; keep at least the first five:

| Case | What happens |
|---|---|
| A folder, file, column or the master list is missing | Not built. Listed before anything is created. |
| A model is not on the master list | UNMAPPED, bold on peach; named in the Run Log and the Last run box. An E-model is placed by the tonnes rule instead, and logged. |
| Two exports of one kind are in the folder | Both read; a warning names them. |
| An earlier report or an Excel lock file is in the folder | Ignored. |
| Excel's filter hides other rows than the rules | Every row is re-checked; WARNING in the Run Log. |
| A date is stored as text | Converted if it reads as a date; else left as is and listed by row. |
| The report is open with unsaved changes | The build stops. |
| A failure part-way | The Last run box names the step; the part-built file is left open, unsaved. |

(Sources: `CheckSetup`; `BuildOpenOrdersReport` and its `Fail:` block; macro-instructions.md 3.1, 3.4 and Section 9.)

### 5.4 Panel "Written down" (traceability)

A boxed panel with two parts.

Part one, a short list headed **In every report, File Format tab, Table 4 "This build"**: built when; builder version; input folder; each file read, with rows read, written, dropped, blank, and the MSA column used; the size class file and its usable rows; the inventory month; the MSAs; the options. And one line: "Tables 1 to 3 on the same tab print the size-class rule, the status order and the default filters." (Source: `WriteFileFormat`; macro-instructions.md 5.7.)

Part two, headed **Run Log, one line per file and per MSA** in monospace (Consolas is the report's log font; use `\ttfamily` here), quoting verbatim from `expected.json`, `runs[0].log_lines`:

```
File Order Intake 2026.xlsx: 30 rows -> 28 written, 1 dropped (ATT 1), 1 blank ignored; MSA column MSA, 3 MSAs
Amarillo: Order Intake 20 rows (CTL 11, MX 7, SSL 2; 4 CANCELLED), 5 hidden by the filters; UNMAPPED model groups: T999
```

Under it, two plain lines: "Rows read = written + dropped + blank, per file." and "Hidden rows stay in the file: clear the filter to see them. Nothing is added by hand after a build." (Sources: macro-instructions.md 3.6, 3.7, 5.7.)

### 5.5 Strip "Tested" and the title block

A one-line or two-line strip across the bottom, headed **Tested (v2.4, 2026-10-12)**, facts in this order, separated by rules or bullets:

- Synthetic exports with known answers: two runs, 22 checks each, pass.
- Austin sample against the hand-built tab: 170 of 170 visible rows, same size-class order; 29 of 29 inventory rows in order, 3 totals off by 1 or 2 (anonymised sample).
- The macro run in LibreOffice against the Python version of the rules: identical builds, 19 comparison checks.
- Not yet run in the team's Excel: a 25-minute test protocol is ready.
- One question open: whether the Open Order export is needed (listed, not used).

(Source: macro-instructions.md Section 6, the v2.0, v2.1 and v2.4 entries, and Section 7 "Still open".)

Title block, bottom right, drawn as a small ruled table in the drawing-sheet manner, fields: **Title** Open Orders report macro, build flow; **Series** Strategy; **Macro** v2.4; **Rev.** 1; **Date** 2026-10-12; **Drawn** Gust F.; **Sheet** 1 of 1. No logo, no company name.

### 5.6 Cut order when the page is full

Cut in this order, one at a time, recompiling after each: (1) the last three rows of "Stopped or flagged"; (2) the second Run Log line; (3) the "Tables 1 to 3" line in "Written down"; (4) the side note on warnings at the gate; (5) the "About one run a year" sentence. Never cut: the five stations, the gate and its "Not built" box, the first five "Stopped or flagged" rows, the first Run Log line, the "Tested" strip, the title block.

---

## 6. The look: a drawing, not a slide

The brief's hardest requirement is that the page does not read as AI-made. Those visuals share a style: rounded boxes with pastel fills and soft shadows, an icon per box, three or four equal "benefit cards" with check marks, a gradient or bold colour bar, chevron arrows, a centred headline with a tagline, headings like "Key benefits" or "Why it matters", and vocabulary like seamless, robust, streamlined, empowered. They are symmetrical, generic and number-free. This page is the opposite on every count. It is drawn the way an engineer documents a process on a drawing sheet, which also matches the report itself (v2.4 is "set like a printed engineering report, in green").

### 6.1 Rules

- **Sheet.** A thin frame (0.8 pt) around the whole drawing area, square corners. A title block in the bottom-right corner (Section 5.5). No header or footer outside the frame; `\pagestyle{empty}`.
- **Boxes.** Square corners, 0.5 pt black outline, white fill. Station headers (the number and name) are set bold white on the report's header green, exactly as the report's own header rows are. Body text inside boxes is left-aligned, ragged right, sentence case.
- **Arrows.** Thin (0.5 pt) black lines with small filled heads (`-{Stealth[length=4pt, width=3pt]}`). Orthogonal routing (`-|`, `|-`) where a straight line does not do. No chevrons, no curved swooshes, no double-headed decoration.
- **Colour.** Four colours and no others, each with one meaning: black for text and lines; **header green `548235`** for station headers and the two panel headers; **light green `E2EFDA`** as the fill of the "Written down" panel body only, if a fill is wanted at all; **peach `FCE4D6`** for the word UNMAPPED and nothing else. White elsewhere. No gradients, no shadows, no transparency.
- **Type.** Latin Modern (`lmodern`, T1): sans (`\sffamily`) for everything in the drawing, bold for headers, `\ttfamily` for the Run Log lines and file names. Sizes: title 14 pt bold, station headers 9 pt bold, box and table text 8 pt, side notes 7.5 pt, nothing under 7 pt. Hierarchy comes from weight and size, never from colour.
- **Words.** Specific nouns: file names, tab names, settings, counts, a real log line. Banned anywhere on the page (grep the `.tex` for them): benefit, benefits, key, takeaway, seamless, robust, streamline, streamlined, powerful, effortless, ensure, ensures, leverage, empower, transform, journey, solution, optimise, optimize, efficient, efficiency, reliable, consistent, best practice, game changer, unlock, elevate; also exclamation marks, check-mark symbols, emoji and icons of any kind.
- **Shape of the page.** Asymmetric by design: the flow takes about two thirds of the width across the upper half; the two panels below are of different widths; the "Tested" strip runs under them; the title block sits in the corner. Nothing is centred except text inside a station header.
- **Numbering.** Stations 1 to 5 in their headers. No other numbered callouts are needed; if a note must point at a station, write "(station 3)".

### 6.2 What to do when it starts to look like a slide

If, on the rendered PNG, you see a row of equal boxes, a big headline, any rounded corner, more than four colours, or a heading that announces a feeling, fix that before anything else. Make the gate branch visibly drop out of the flow line (the stop is the point). Let the panels be as tall as their text and no taller.

---

## 7. Layout (US Letter landscape, 11 x 8.5 in, 0.5 in margins, drawing area 10 x 7.5 in)

Draw in one `tikzpicture` with `x=1in, y=1in` and absolute coordinates on the 10 x 7.5 in sheet (origin bottom left). Fix the picture's size with `\useasboundingbox (0,0) rectangle (10,7.5);` so the page cannot grow. Rough placement (adjust by inches, keep the regions):

```
 y=7.5 +--------------------------------------------------------------------------------+
       | Open Orders report macro   (title, 14 pt)                                      |
       | One build, from the exports to the workbook                                    |
       | By hand, for each MSA: ...                     One build, all MSAs: ...        |
 y=6.4 |                                                                                |
       |  [1 Exports]--->[2 Config]--->[3 Setup check]--->[4 Build]--->[5 Workbook]     |
       |  [Size class  ]                 |  Problems?         ^   same rules      |      |
       |  [workbook    ]-----------------)------------------>+                    |      |
 y=4.6 |        ^                        v                                        |      |
       |        |                 [Not built. One list ...]                       |      |
       |        +--- UNMAPPED: add the model to the list, build again ------------+      |
 y=4.1 |                                                                                |
       |  +-- Stopped or flagged (error-proofing) --+   +-- Written down -------------+ |
       |  | Case            | What happens          |   | File Format, Table 4: ...   | |
       |  | ...             | ...                   |   | Run Log: (two tt lines)     | |
       |  | (5 to 8 rows)   |                       |   | Rows read = ...             | |
 y=1.4 |  +------------------------------------------+   +-----------------------------+ |
       |  Tested (v2.4, 2026-10-12): fixtures 2 x 22 | Austin 170/170 ... | ...        |
 y=0.9 |                                                            +-- title block ---+ |
       |                                                            | Title | Rev. | ... | |
 y=0   +------------------------------------------------------------+-------------------+
```

Widths: stations about 1.6 in each with 0.3 in gaps; the "Stopped or flagged" panel about 5.4 in wide, "Written down" about 4.2 in; the title block about 3.2 x 0.8 in. The flow's arrows run along one horizontal line at the stations' vertical centre; the "Not built" box hangs below station 3; the return arrow from station 5 goes down, left along the bottom of the flow region, and up into the size class workbook box.

---

## 8. Technical skeleton and the build loop

### 8.1 Environment

A cloud session has no TeX by default. Install once (verified 2026-10-12 on this environment, Ubuntu noble, TeX Live 2023; about four minutes):

```
sudo apt-get update && sudo DEBIAN_FRONTEND=noninteractive apt-get install -y --no-install-recommends \
  latexmk texlive-latex-base texlive-latex-recommended texlive-latex-extra texlive-pictures texlive-fonts-recommended lmodern
```

`lmodern` is a separate Debian package that `--no-install-recommends` leaves out; without it `lmodern.sty` is not found. `pdfinfo`, `pdftoppm` and `pdftotext` (poppler) are already present.

### 8.2 Preamble (start from this; it is the house preamble reduced to what a drawing needs)

```latex
\documentclass[10pt]{article}
\usepackage[letterpaper, landscape, margin=0.5in]{geometry}
\usepackage[T1]{fontenc}
\usepackage{lmodern}
\usepackage{microtype}
\usepackage{xcolor}
\usepackage{tikz}
\usetikzlibrary{positioning, arrows.meta, calc, fit, shapes.geometric}
\usepackage[hidelinks, pdftitle={Open Orders Report Macro: Build Flow}, pdfauthor={Gust F.}]{hyperref}
\pagestyle{empty}
\setlength{\parindent}{0pt}
\definecolor{headgreen}{HTML}{548235}
\definecolor{bandgreen}{HTML}{E2EFDA}
\definecolor{flagpeach}{HTML}{FCE4D6}
\tikzset{
  every node/.style={font=\sffamily\footnotesize, inner sep=4pt},
  box/.style={draw, line width=0.5pt, fill=white, align=left, text width=1.45in},
  station/.style={draw, line width=0.5pt, fill=headgreen, text=white, font=\sffamily\bfseries\small,
                  align=center, text width=1.45in, inner ysep=3pt},
  flow/.style={-{Stealth[length=4pt, width=3pt]}, line width=0.5pt},
  gate/.style={draw, line width=0.5pt, diamond, aspect=2.2, inner sep=1pt, font=\sffamily\scriptsize},
  panelhead/.style={fill=headgreen, text=white, font=\sffamily\bfseries\small, align=left, inner ysep=3pt},
  flag/.style={fill=flagpeach, font=\sffamily\bfseries\footnotesize, inner sep=1pt}
}
```

Put the station body text in a `box` node placed `below=0pt of` its `station` node so header and body read as one box with a green header row (the report's own look). Build tables inside panels with `tabular` in a node (`align=left`, `text width` set) or with `\matrix` in TikZ; `tabular` is simpler and wraps nothing, so set a `p{..}` width per column. Typeset the Run Log lines with `\ttfamily\scriptsize` (7 pt). Measured: the first line is 413 pt (5.7 in) wide, the second 450 pt (6.2 in), so neither fits a panel unwrapped. Let each wrap inside the panel's `text width` with a 1 em hanging indent, the way a log viewer wraps a long line: `\hangindent=1em\hangafter=1 <line>\par`. In a 3.9 in panel each takes two lines, which reads well. (The shorter inventory line, `Amarillo: Inventory 10 rows; September 2026: CTL 6, MX 8, SSL 4 units`, is 264 pt and fits unwrapped, but the two lines above are the ones that show reconciliation, hidden rows and UNMAPPED.)

This preamble, the `tikzset` styles, a station with its body box, the inline UNMAPPED flag, the gate, Stealth arrows, a `tabular` inside a node and the wrapped log lines were compiled together in a smoke test on 2026-10-12: one page, 792 x 612 pt, no Overfull boxes.

### 8.3 Build, check, look, repeat

From the repo root:

```
./build.sh source/open-orders-intake-inventory/open-orders-macro-flow.tex
pdfinfo output/open-orders-intake-inventory/open-orders-macro-flow.pdf | grep -E "Pages|Page size"
grep -E "Overfull|Underfull|Missing character|undefined" .build/open-orders-intake-inventory/open-orders-macro-flow.log
pdftotext output/open-orders-intake-inventory/open-orders-macro-flow.pdf - | wc -w
pdftoppm -r 110 -png output/open-orders-intake-inventory/open-orders-macro-flow.pdf /tmp/flow   # then view /tmp/flow-1.png
```

Expect `Pages: 1` and `Page size: 792 x 612 pts (letter)`. View the PNG after every change that moves anything; TikZ text overflowing a node does not show in the log. Iterate on the layout until Section 9 passes. Keep the `.tex` tidy: one block per region, a comment line naming it, coordinates as plain numbers, no macros defined beyond the `tikzset` above.

---

## 9. Acceptance checklist (fill in honestly in the final report)

1. [ ] `pdfinfo`: Pages 1; Page size 792 x 612 pts.
2. [ ] Log: no Overfull box over 2 pt, no missing characters, no undefined references. The PNG shows no text clipped by or overflowing a box, no overlapping labels, no arrow crossing text.
3. [ ] Every statement on the page is in Section 5, and each fact there is traceable to its named source; the final report lists any wording changed because the source said otherwise.
4. [ ] Word count at most 300 (title block excluded); cuts, if any, followed Section 5.6.
5. [ ] Colours: only black, white, `548235`, `E2EFDA`, `FCE4D6`; peach only on UNMAPPED. No gradient, shadow, transparency, rounded corner, icon, emoji or check mark.
6. [ ] `grep -niE` for the banned words in Section 6.1 finds nothing in the `.tex`; no exclamation mark on the page.
7. [ ] The sentence test of Section 2, rule 4: no sentence tells the reader what to conclude.
8. [ ] Smallest type 7 pt or larger; readable in the PNG at 110 dpi without zooming.
9. [ ] The reading order of Section 4 is the visual order: title, strip, stations 1 to 5 with the gate dropping out, panels, tested strip, title block.
10. [ ] Title block complete (Title, Series, Macro, Rev. 1, Date, Drawn, Sheet 1 of 1). `hyperref` metadata set.
11. [ ] Optional second AI reviewer, given only the PNG and Section 6: "Does this read as AI-made, and why?" Findings addressed or answered in the final report.
12. [ ] `MAP.md`: rows for the `.tex` and the `.pdf` added; this plan's row says it was executed. One commit on `main` with the `.tex`, the `.pdf` and `MAP.md`; pushed.

---

## 10. Things that bite on a page like this

- TikZ nodes with `text width` wrap but never grow sideways; a long file name or a Run Log line in `\ttfamily` does not wrap and silently overflows. Check the PNG, not the log.
- `\useasboundingbox` keeps the page at one sheet; without it a node placed past the margin pushes the picture onto a second page or off the paper.
- `geometry` with `landscape` and `letterpaper` gives 792 x 612; if `pdfinfo` shows 612 x 792 the option was lost (check the `\documentclass` line has no `a4paper` and `geometry` loads first).
- `latexmk` through `build.sh` writes aux files to `.build/`, which is git-ignored; only the PDF lands in `output/`. Do not commit `.build/`.
- `microtype` warns once about `\ttfamily` and letter spacing in some TeX Live builds; harmless.
- `-{Stealth[...]}` needs `arrows.meta`; `diamond` needs `shapes.geometric`; a missing library fails with "Unknown arrow tip" or "unknown shape".
- Latin Modern's `\sffamily` in `\bfseries` at 8 pt is fine; do not switch to another sans to "look less like TeX": the font is not the tell, the layout and the words are.
- The manual's TikZ warning sign and `tcolorbox` are not needed here; do not load `tcolorbox`.
- `positioning` keys set the anchor: `below=0pt of a.south west` makes the new node's anchor `north`, so an `anchor=north west` written *before* it is overridden and the node is centred on that corner (the smoke test's panel table landed half outside its panel this way). Write `anchor=` after the `below=`/`right=` key, or place with `at (a.south west)` plus `anchor=north west`.
- `latexmk` remembers a failed run: after fixing the cause it can answer "Nothing to do" and fail again with the old error. Delete `.build/open-orders-intake-inventory/open-orders-macro-flow.*` and rebuild.
- `\settowidth` inside a `tikzpicture` (between path commands) measures in `nullfont` and returns 0 pt with a stream of "Missing character" warnings. Measure outside the picture, or trust the figures in Section 8.2.
- A wrapped `\ttfamily` log line produces one "Underfull \hbox" warning per wrapped paragraph; that is expected and allowed. Only Overfull boxes fail the checklist.

---

## 11. The prompt to paste into the Stronger AI review

> Read `output/open-orders-intake-inventory/macro-flow-visual-implementation-plan.md` in full and execute it as the implementer. Follow its standing rules exactly: do the work yourself in one `.tex` file, keep every statement on the page to Section 5 and its sources, apply the look rules of Section 6, lay it out as Section 7 describes, build with `./build.sh` and check it as Section 8 says, and stop only when the Section 9 checklist passes. Install TeX first with the line in Section 8.1. Change nothing else in the repository. Finish by updating `MAP.md` and committing the `.tex`, the `.pdf` and `MAP.md` to `main` in one commit, then report the Section 9 checklist filled in honestly, with the word count, the `pdfinfo` line, anything you cut by Section 5.6, and anything you could not verify.
