# Gust SWI Template v6: Reusable Prompt

Version 6, 2026-10-23. A style guide with two tools in the same house style:

- **Tool 1, the SWI:** a Standard Work Instruction on any topic.
- **Tool 2, the Boardroom Deck:** a 4 to 6 sheet deck that presents a finished project to a manager. It has an agenda with a timeline drawn to scale, then design process, architecture, user guide and value sheets.

Paste everything below into another AI. At the end, name the tool you want and add your content.

---

# Tool 1: Standard Work Instruction (SWI)

Write a Standard Work Instruction (SWI) as one self-contained pdfLaTeX file in TikZ. It shows how something works (a file, tool, machine or process), then puts its single most important task on its own emphasized sheet. Follow this specification exactly.

## Purpose and split
- The reader is someone who uses the thing and does not build it. They should get it from the pictures; the text only explains them.
- **About 80%: how it works.** The overview sheet, then one sheet per screen, tab or part, then a sheet on how to check the result and fix common problems.
- **About 20%: the critical task**, on the last sheet, emphasized (see "The critical-task sheet"). The overview sheet points to it.
- Where several cases run the same way (e.g. several product types), write one set of steps with a small lookup table of what differs. Do not repeat the steps for each case.
- Typical length: 4 to 6 sheets.

## Page and grid
- `\documentclass[10pt]{article}`, paper 17 x 11 in (11 x 17 landscape), `margin=0.5in`, `\pagestyle{empty}`.
- Packages: `fontenc` (T1), `lmodern`, `microtype`, `enumitem`, `xcolor` with the `table` option, `graphicx`, `tikz` (libraries `positioning, arrows.meta, calc, fit`), `hyperref` (`hidelinks`, `pdftitle`, `pdfauthor`).
- Every sheet is one `tikzpicture` with `x=1in, y=1in`, drawn on a 12 x 7.5 in grid and wrapped in `\scalebox{1.333333}{...}` to fill the page. Use `\useasboundingbox (0,0) rectangle (12,7.5)`, a 0.8pt frame and `\newpage` between sheets.
- Font: sans serif (`\sffamily\small` for every node). Turn hyphenation off (`\hyphenpenalty=10000`, `\exhyphenpenalty=10000`).
- Keep 0.25 in inside the frame. Nothing may overlap the title block (x 5.7 to 12, y 0 to 0.42).

## Colors (the only ones allowed)
- `headgreen` `#548235`: station headers and markers.
- `bandgreen` `#E2EFDA`: light sections and table bands.
- `flagpeach` `#FCE4D6`: warnings and changed items.
- `inputyellow` `#FFFFCC`: fields the reader fills in.
- `alertorange` `#C55A11`: only on the critical-task sheet (frame, banner, the overview's pointer box).
- Step colors, in this order, one per pictured step (see "Colored step bullets"): `stepa` `#C55A11` orange, `stepb` `#2E75B6` blue, `stepc` `#7030A0` purple, `stepd` `#BF8F00` dark gold, `stepe` `#C00000` red.

## Building blocks (use the macros in the template)
- **Heading:** `\heading{Title}{One-sentence subtitle}` at the top left. The title is Title Case, under about 40 characters.
- **Station:** a green header bar (`station`), number then title, with a white `box` under it. Station titles start with a verb or name a thing ("Control panel", "Check the result"). Join stations with `flow` arrows, left to right or top to bottom.
- **Box contents:** an italic one-line tag via `\by{...}`, then bullets (`itemize`) or numbered steps (`enumerate`). Bullets are short fragments; steps are imperative sentences.
- **Picture:** `\node[shot] at (x,y) {\includegraphics[width=W in]{file.png}};`, anchored at its top left. Its height is H = W x (pixel height / pixel width). Write W, H and the pixel size in a comment above the sheet.
- **Markers on a picture:** open a scope over the picture, `\begin{scope}[shift={(x,y)}, x=W in, y=-H in]`. Inside it, (0,0) is the top left of the picture and (1,1) its bottom right. Place each `\node[marker] at (fx,fy) {1};` from the pixel position: fx = px / width, fy = py / height. Place a marker in empty space beside the cell or header it points to, never over its text. Repeat the same number or letter in the text with `\num{1}`. A step's markers take that step's color (see "Colored step bullets").
- **Buttons and settings:** `\btn{Button name}` for buttons, `\texttt{}` for files and settings, `\val{}` for values typed exactly. `\yellow{}` and `\peach{}` highlight inline words in those colors.
- **Reference table:** a `tabular` inside a node, `\arraystretch` about 1.4. Group the rows under full-width `\cellcolor` band rows by severity (e.g. "Stops the work", "Work can continue"). Put each symptom or message, quoted exactly, in the first column (`\ttfamily`, or bold for its key words) and the fix in the second, as an imperative sentence.
- **Colored step bullets:** station bars are always green; never color the bar itself. A step whose action or result is shown in a picture gets its number bullet in its own color, unique to that step on the sheet: `\stepnum{stepa}{3}`, then `stepb` for the next pictured step, and so on. Its markers (`\node[stepa] at ...`), frames and pointer arrows on the pictures take the same color, so the eye goes from the step straight to its place in the picture. Steps not shown in a picture (open, find, the final check) keep the plain green `\num` bullet. Never use a step color for anything else on that sheet.
- **Pointer:** the `point` arrow links a marker to more than one place; give it the step's color (`\draw[point, stepb]`). Frame the places it points to in the same color (`\draw[line width=1.4pt, stepb] ... rectangle ...`).
- **Title block:** `\titleblock{Short title}{n of N}` on every sheet. It shows Title, Project, Version, Rev., Date, Drawn and Sheet. Set these once in the `\titleblock` macro.

## Pictures
- Use pictures of the real thing, never mock-ups. Keep any reference grid visible (labels, row and column headings, part numbers) so references in the text can be checked against the picture. Crop each picture to its content.
- Make the pictures in a way that can be repeated when the subject changes. To show an example change, change a copy and shade what changed `flagpeach`, in bold. Say so on the sheet.
- One picture per idea. A full view goes on its own sheet; a close-up goes beside the steps that use it.

## Sheet types
1. **Overview ("at a glance"):** 3 or 4 stations in a row with flow arrows. Each holds a thumbnail picture or 3 bullets. Under them: a small table of the key types or terms, an "In this guide" list of sheets (the critical sheet in bold), and a `flagpeach` box with a 2pt `alertorange` border. That box names the critical task, gives it in 2 sentences, and points to its sheet.
2. **Annotated picture:** the picture large on the left (about 7 in wide), markers on it, and the meaning of each numbered marker in a column on the right. Group rows with brackets and rotated labels on the left if they fall into families. Below the picture go one or two stations: "Add a ..." (steps) and "Good to know" (bullets).
3. **Wide picture:** the picture full width (about 11.5 in), lettered markers (A, B, C), then a row of 3 or 4 stations under it with the same letters in their headers.
4. **How to check it:** a top-to-bottom chain of stations on the left (what happens at each step, and where to see the result), and the reference table of problems and fixes on the right.
5. **The critical task (last sheet):** see below.

## The critical-task sheet
- Draw a 3pt `alertorange` frame on the sheet edge in place of the black one. Put a white-on-orange banner at the top right: **THE MOST IMPORTANT JOB**.
- The title names the task and the cases it covers ("<Task>: <case A>, <case B>, <case C>"). The subtitle says the steps are the same for every case and names the example values.
- Left column (about 4.45 in): numbered stations, one step each, 0.12 in apart, 5 to 7 steps. Put the lookup table of what differs by case in the step that needs it. The last step is the check that proves it worked, with the exact text to expect.
- Right column: an "after" picture for each step that changes something visible. Put each step's markers on it with the step's number, and rectangles around the places a step changes, all in that step's color ("Colored step bullets").
- Under the steps: a `flagpeach` box beginning **Do not:** that lists the mistakes that break things, in one or two lines.

## Writing
- Present tense, plain words, the reader's own names for things, the same name for a thing everywhere.
- Say what the reader does and what they will see. Skip how it works inside, unless the reader acts on it.
- Every fact comes from the real file or code. Any example value is called an example.
- Nothing may be cut off or overlap. Compile, look at every sheet, and move or shrink until it all fits.

## Template

```latex
% <Subject> SWI: <what it covers>. Gust SWI Template v3.
% Sheets: 1 at a glance, 2 <part>, 3 <part>, 4 how to check it, 5 the critical task (orange frame).
% Pictures in <folder>/. Grid 12 x 7.5 in, scaled 4/3.
\documentclass[10pt]{article}
\usepackage[paperwidth=17in, paperheight=11in, margin=0.5in]{geometry}
\usepackage[T1]{fontenc}
\usepackage{lmodern}
\usepackage{microtype}
\usepackage{enumitem}
\usepackage[table]{xcolor}
\usepackage{graphicx}
\usepackage{tikz}
\usetikzlibrary{positioning, arrows.meta, calc, fit}
\usepackage[hidelinks, pdftitle={<Title>}, pdfauthor={Gust F.}]{hyperref}
\graphicspath{{<folder>/}}
\pagestyle{empty}
\setlength{\parindent}{0pt}
\hyphenpenalty=10000
\exhyphenpenalty=10000
\setlist[itemize]{label=\textbullet, leftmargin=0.85em, labelsep=0.35em, itemsep=1.5pt, parsep=0pt,
                  topsep=0pt, partopsep=0pt}
\definecolor{headgreen}{HTML}{548235}
\definecolor{bandgreen}{HTML}{E2EFDA}
\definecolor{flagpeach}{HTML}{FCE4D6}
\definecolor{inputyellow}{HTML}{FFFFCC}
\definecolor{alertorange}{HTML}{C55A11}
\definecolor{stepa}{HTML}{C55A11}
\definecolor{stepb}{HTML}{2E75B6}
\definecolor{stepc}{HTML}{7030A0}
\definecolor{stepd}{HTML}{BF8F00}
\definecolor{stepe}{HTML}{C00000}
\tikzset{
  every node/.style={font=\sffamily\small, inner sep=5pt},
  box/.style={draw, line width=0.5pt, fill=white, align=left},
  station/.style={draw, line width=0.5pt, fill=headgreen, text=white, font=\sffamily\bfseries\normalsize,
                  align=left, inner ysep=4pt, text height=1.6ex, text depth=0.45ex},
  flow/.style={-{Stealth[length=9pt, width=6.5pt]}, line width=1.2pt, line join=miter},
  point/.style={-{Stealth[length=6pt, width=4.5pt]}, line width=0.9pt},
  marker/.style={circle, fill=headgreen, text=white, draw=white, line width=1pt, font=\sffamily\bfseries\small,
                 inner sep=0pt, minimum size=0.25in},
  % one colour per pictured step: the bullet, its markers, frames and arrows
  stepa/.style={marker, fill=stepa}, stepb/.style={marker, fill=stepb}, stepc/.style={marker, fill=stepc},
  stepd/.style={marker, fill=stepd}, stepe/.style={marker, fill=stepe},
  shot/.style={draw, line width=0.5pt, inner sep=0pt, anchor=north west},
}
\newcommand{\by}[1]{{\footnotesize\itshape #1}\\[3pt]}
\newcommand{\peach}[1]{{\setlength{\fboxsep}{1pt}\colorbox{flagpeach}{\textbf{#1}}}}
\newcommand{\yellow}[1]{{\setlength{\fboxsep}{1pt}\colorbox{inputyellow}{#1}}}
\newcommand{\btn}[1]{{\setlength{\fboxsep}{1.5pt}\fbox{\footnotesize\bfseries #1}}}
\newcommand{\num}[1]{\tikz[baseline=(n.base)]\node(n)[marker, minimum size=0.2in, font=\sffamily\bfseries\footnotesize]{#1};}
\newcommand{\stepnum}[2]{\tikz[baseline=(n.base)]\node(n)[#1, minimum size=0.2in, font=\sffamily\bfseries\footnotesize]{#2};}
\newcommand{\val}[1]{{\ttfamily #1}}
\newcommand{\titleblock}[2]{%
  \draw[line width=0.5pt] (5.7,0) -- (5.7,0.42) -- (12,0.42);
  \foreach \x in {7.3, 8.95, 9.5, 9.95, 10.75, 11.45} \draw[line width=0.5pt] (\x,0) -- (\x,0.42);
  \foreach \x/\label/\value in {5.7/Title/#1, 7.3/Project/<Project>, 8.95/Version/<v>, 9.5/Rev./1,
                                9.95/Date/<YYYY-MM-DD>, 10.75/Drawn/Gust F., 11.45/Sheet/#2} {
    \node[anchor=north west, inner sep=0pt, font=\sffamily\fontsize{7}{8}\selectfont] at (\x+0.07,0.36) {\label};
    \node[anchor=south west, inner sep=0pt, font=\sffamily\bfseries\footnotesize] at (\x+0.07,0.07) {\value};
  }}
\newcommand{\heading}[2]{%
  \node[anchor=north west, inner sep=0pt, font=\sffamily\bfseries\fontsize{14}{16}\selectfont] at (0.25,7.25) {#1};
  \node[anchor=north west, inner sep=0pt, font=\sffamily\footnotesize] at (0.25,6.88) {#2};}

\begin{document}
% ================================================================ sheet 1: at a glance
\noindent\scalebox{1.333333}{%
\begin{tikzpicture}[x=1in, y=1in]
\useasboundingbox (0,0) rectangle (12,7.5);
\draw[line width=0.8pt] (0,0) rectangle (12,7.5);
\heading{<Subject> at a Glance}{<What it is and what it is for, in one sentence.>}
\node[station, anchor=north west, text width=2.55in] (h1) at (0.25,6.4) {1\quad <Part one>};
\node[box, anchor=north, text width=2.55in] (b1) at (h1.south) {\by{<What it decides>}
  \begin{itemize} \item <Fact> \item <Fact> \item <Fact> \end{itemize}};
\node[station, anchor=north west, text width=2.55in] (h2) at (3.2,6.4) {2\quad <Part two>};
\node[box, anchor=north, text width=2.55in] (b2) at (h2.south) {\by{<What it decides>}
  \begin{itemize} \item <Fact> \item <Fact> \end{itemize}};
\draw[flow] (h1.east) -- (h2.west);
\node[anchor=north west, draw=alertorange, line width=2pt, fill=flagpeach, text width=5.15in, align=left,
      inner sep=8pt] at (6.15,2.05)
  {{\bfseries\large Most important job: <task>}\hfill{\bfseries Sheet <N>}\\[5pt] <The task in two sentences.>};
\titleblock{At a Glance}{1 of <N>}
\end{tikzpicture}}

\newpage
% ================================================================ sheet 2: annotated picture
% picture: <w> x <h> px, 7.0 in wide, so <7.0*h/w> in high, top left at (0.75,6.4)
\noindent\scalebox{1.333333}{%
\begin{tikzpicture}[x=1in, y=1in]
\useasboundingbox (0,0) rectangle (12,7.5);
\draw[line width=0.8pt] (0,0) rectangle (12,7.5);
\heading{<Part one>}{<What this part holds, and what the reader does here.>}
\node[shot] at (0.75,6.4) {\includegraphics[width=7.0in]{<picture>.png}};
\begin{scope}[shift={(0.75,6.4)}, x=7.0in, y=-<H>in]
  \node[marker] at (<fx>,<fy>) {1};
\end{scope}
\node[anchor=north west, align=left, text width=3.6in, inner sep=0pt] at (8.05,6.4) {%
  \textbf{\large <The parts>}\\[6pt]
  \num{1}\enspace\textbf{<Name>}\\[1pt] {\footnotesize <What it is; whether it matters to the reader.>}};
\titleblock{<Part one>}{2 of <N>}
\end{tikzpicture}}

\newpage
% ================================================================ sheet N: the critical task
\noindent\scalebox{1.333333}{%
\begin{tikzpicture}[x=1in, y=1in]
\useasboundingbox (0,0) rectangle (12,7.5);
\draw[line width=3pt, alertorange] (0,0) rectangle (12,7.5);
\node[anchor=north west, inner sep=0pt, font=\sffamily\bfseries\fontsize{14}{16}\selectfont]
  at (0.25,7.25) {<Task>: <case A>, <case B>, <case C>};
\node[anchor=north east, fill=alertorange, text=white, font=\sffamily\bfseries\normalsize, inner sep=5pt]
  at (11.75,7.28) {THE MOST IMPORTANT JOB};
\node[anchor=north west, inner sep=0pt, font=\sffamily\footnotesize] at (0.25,6.88)
  {The same steps for every case. Example values are shaded \peach{peach} in the pictures.};
\node[station, anchor=north west, text width=4.45in] (s1) at (0.25,6.45) {\num{1}\enspace <Step one (not pictured)>};
\node[station, anchor=north west, text width=4.45in] (s2) at ($(s1.south west)+(0,-0.12)$) {\stepnum{stepa}{2}\enspace <Step two (pictured)>};
\node[box, anchor=north, text width=4.45in, font=\sffamily\footnotesize] (t2) at (s2.south) {%
  \begin{tabular}{@{}p{0.5in}p{1.5in}p{1.5in}@{}}
  \textbf{<A>} & <where> & e.g.\ \val{<value>} \\
  \textbf{<B>} & <where> & e.g.\ \val{<value>} \\
  \end{tabular}};
\node[station, anchor=north west, text width=4.45in] (s3) at ($(t2.south west)+(0,-0.12)$) {\num{3}\enspace <Check it worked>};
\node[box, anchor=north, text width=4.45in] (t3) at (s3.south) {<What to expect, word for word.>};
\node[anchor=north west, draw, line width=0.5pt, fill=flagpeach, text width=4.45in, align=left,
      font=\sffamily\footnotesize] at ($(t3.south west)+(0,-0.2)$) {\textbf{Do not:} <the mistakes that break it>.};
\node[shot] at (5.15,6.0) {\includegraphics[width=6.6in]{<after-picture>.png}};
\begin{scope}[shift={(5.15,6.0)}, x=6.6in, y=-<H>in]
  \node[stepa] at (<fx>,<fy>) {2};
\end{scope}
\titleblock{<Task>}{<N> of <N>}
\end{tikzpicture}}
\end{document}
```

# Tool 2: The Boardroom Deck

Write a Boardroom Deck as one self-contained pdfLaTeX file in TikZ: 4 to 6 sheets that present a finished tool or project to a manager and carry the meeting about it. It uses the SWI's page, grid, colors, title block, heading, stations, boxes, markers and `\btn`/`\val`. It has no critical-task sheet and no `alertorange`. Follow this specification exactly. The reference build is the QMS Editor deck (`source/qms-editor/qms-editor-overview-v11-2026-10-23.tex` in the `gustjf/vba` repo).

## Purpose
- The reader is the owner's manager. In a few minutes per sheet they should see that the work was planned, built in stages and proven. They should also see what the system is, how people use it, and what it is worth to the company.
- Concise and professional, in the owner's plain words, with no flourish. Each sheet's subtitle is the first sentence the presenter says.
- **No fact on two sheets.** Where two sheets touch the same topic, one gives the reason or criterion and the other the result.
- Every fact comes from the owner's notes. Name people only as the owner names them. If you draft next steps or timings yourself, say so when you hand over.

## Sheet order
1. **Agenda**: only when there is a meeting.
2. **Design Process**
3. **Architecture**
4. **User Guide**
5. **Value**

On every sheet, `\nav{n}` at the top right lists the sheets with the current one in bold green. The title block shows "Sheet n of N". Raise Rev. on every change.

## Sheet types
1. **Agenda** (pattern: a timeline drawn to scale).
   - The meeting runs from x = 0.5 to x = 11.5, so minute m sits at x = 0.5 + 11 m / M (M is the length in minutes).
   - **Part bars** at y 6.0 to 6.32: green, "Part n  Name" on the left and "nn min" on the right, with a 0.08 in gap between bars.
   - **Timeline row** at y 5.3 to 5.86: one segment per item, white with a bold label. The **centerpiece** item is green with white text; the closing item (value and next steps, or open discussion) is `bandgreen`.
   - **Ticks:** a tick and a bold time under each boundary (tick 5.3 to 5.18, label at 5.13). Use short unlabelled ticks inside a run of equal items.
   - **Dividers:** a dashed vertical line at each part boundary, from 4.95 down to the outcome band.
   - **Under each part,** show what is presented:
     - Miniatures of the deck's own sheets (`\thumb...` in a scope scaled 1/12, 1.0 x 0.625 in), with "Sheet n / name / minutes" beside each.
     - The centerpiece in a box with a 1.6pt green border, `\badge{Centerpiece}`, a short line on why it matters, and its live-demo steps as a `\badge`-numbered list.
     - A row of `\docicon` pages for a run of similar items, one under each segment, then one "For each ..." list. Don't repeat their IDs, which the timeline already shows.
   - **Outcome:** a full-width `bandgreen` box at the bottom with **We leave with:** and the decisions, in one line.
2. **Design Process**
   - Five stations in a row with flow arrows (text width 1.92 in, x = 0.25, 2.61, 4.97, 7.33, 9.69), all the same height with `\fixh`.
   - Titles are past tense and start with an action verb: Defined the requirements, Studied the inputs, Chose the platform, Built and verified, Refined with feedback. Bullets are past tense and start with a verb too.
   - Underneath: a **How it was proven** table (Test | What it checks | Result, with real counts such as "32 of 32 identical"). Beside it, a **Proof of concept** box with a large green check mark (`\ding{51}`), an upbeat one-line caption, and three or four bullets.
3. **Architecture:** three columns, then a lower row.
   - **Left, "Where the files live":** lettered cards (`\num{A}`, B, C), one per place: the controlled library, each user's PC, the owner's build and backup folder.
   - **Middle, the app as a container:**
     - A station bar named for the outer layer, with its technology right-aligned in the bar (`\mdseries\footnotesize\itshape`) and its role in italic just under the bar.
     - Inside, the other layers as `\layerbody` boxes.
     - A pair that works as one goes in a thin, rounded green group box. Its small-caps label sits on the lower-left border, so it never crowds the arrow at the top.
   - **Flows:** inside the container, a one-way flow runs straight down the middle (or straight up). Keep one-way flows in the same direction, never against each other.
   - **Exchanges:** where two parts trade files both ways, draw two short arrows side by side, one each way, each labelled with what it carries (for example `.tex` down, PDF up).
   - **Input and output arrows** between the files column and the container:
     - Straight, with no bends. The out arrow leaves from the layer that writes the files and lands inside the target card.
     - Labels are centred in the gap, 0.07 in clear of the arrow. A long label goes under its arrow in short, centred lines.
   - **Right:** a tree of what is inside the package.
   - **Lower row:** the anatomy of one file, and a pipeline of `chip` boxes with the outputs as `outchip`.
4. **User Guide**
   - Five stations: set up, load, edit, finish, hand off.
   - A **Loading and saving** table of every button (`\btn`), grouped under `bandgreen` band rows (Load, Save).
   - A box on **what the tool takes care of**: a bold lead word and one sentence each (layout, error checking, fast feedback, simple entry). Treat the inside as a black box: show benefits, not internals.
5. **Value**
   - Three stations: Cost, Consistency and control, Scale and continuity. Each opens with `\stat{<big figure>}{<caption>}`, for example `$0`, `PDF only`, `30`.
   - A full-width **Proposed next steps** numbered list. Show an open point only while it is still open.

## Finish rules
- A box starts at its title bar. Anchor the box or rectangle at the bar's `.south` or `.south west`, never at a typed y.
- A table under a title bar sits in a box with `inner ysep=0pt` and `\par\vspace{5pt}` after the tabular, so its green band touches the bar with no white strip.
- Boxes in a row have the same height, and boxes side by side end on the same line. Set `\fixh` after the first compile.
- No arrow label or text touches a line or a box edge.
- Compile, look at every sheet at full size, and fix anything cut off, crowded or misaligned before you hand it over.

## Template additions

Use the SWI template's preamble, then add this:

```latex
\usepackage{array}
\usepackage{pifont}
\setlist[enumerate]{leftmargin=1.2em, labelsep=0.35em, itemsep=2.5pt, parsep=0pt, topsep=0pt, partopsep=0pt}
\tikzset{
  layer/.style={draw=headgreen, line width=0.5pt, fill=bandgreen, align=left, anchor=north west},
  chip/.style={draw, line width=0.5pt, fill=white, align=center, font=\sffamily\footnotesize,
               minimum height=0.42in, inner sep=4pt},
  outchip/.style={chip, fill=bandgreen, font=\sffamily\bfseries\footnotesize},
}
\renewcommand{\by}[1]{{\footnotesize\itshape #1\par}\vspace{5pt}}
% fixed-height box body, so the boxes in a row line up
\newcommand{\fixh}[2]{\begin{minipage}[t][#1][t]{\linewidth}#2\end{minipage}}
% big figure at the top of a box
\newcommand{\stat}[2]{{\color{headgreen}\bfseries\fontsize{26}{28}\selectfont #1}\\[2pt]{\footnotesize\itshape #2\par}\vspace{8pt}}
\newcommand{\badge}[1]{{\setlength{\fboxsep}{2pt}\colorbox{headgreen}{\textcolor{white}{\bfseries #1}}}}
% a layer of the app: bold name, technology on the right, role underneath
\newcommand{\layerbody}[3]{\textbf{#1}\hfill{\footnotesize\itshape #2}\\[2pt]{\footnotesize #3}}
% the sheets at the top right, the current one in green
\newcommand{\navitem}[3]{\ifnum#1=#2{\color{headgreen}\bfseries #3}\else #3\fi}
\newcommand{\nav}[1]{%
  \node[anchor=north east, inner sep=0pt, font=\sffamily\footnotesize] at (11.75,7.22)
    {\navitem{#1}{1}{1\enspace Agenda}\qquad\navitem{#1}{2}{2\enspace Design Process}\qquad
     \navitem{#1}{3}{3\enspace Architecture}\qquad\navitem{#1}{4}{4\enspace User Guide}\qquad
     \navitem{#1}{5}{5\enspace Value}};}
% miniatures for the agenda, drawn in each sheet's own 12 x 7.5 coordinates (use in a scope scaled 1/12)
\newcommand{\thumbframe}{%
  \draw[line width=0.5pt, fill=white] (0,0) rectangle (12,7.5);
  \fill (0.25,7.0) rectangle (3.4,7.25);
  \draw[line width=0.3pt] (5.7,0) -- (5.7,0.42) -- (12,0.42);}
\newcommand{\thumbbox}[4]{% x1, y top, x2, y bottom: a box with its green bar
  \draw[line width=0.3pt] (#1,#2) rectangle (#3,#4); \fill[headgreen] (#1,#2-0.3) rectangle (#3,#2);}
% one \thumb<sheet> per sheet: \thumbframe, then one \thumbbox per box on that sheet, e.g.
\newcommand{\thumbdesign}{%
  \thumbframe
  \foreach \a in {0.25, 2.61, 4.97, 7.33, 9.69} \thumbbox{\a}{6.4}{\a+2.06}{3.84};
  \thumbbox{0.25}{3.6}{7.39}{1.3} \thumbbox{7.75}{3.6}{11.75}{1.3}}
% a document page with a folded corner, 0.56 x 0.72 in, anchored at its top left
\newcommand{\docicon}[2]{% x, y
  \draw[line width=0.5pt, fill=white] (#1,#2) -- (#1+0.42,#2) -- (#1+0.56,#2-0.14) -- (#1+0.56,#2-0.72) -- (#1,#2-0.72) -- cycle;
  \draw[line width=0.5pt] (#1+0.42,#2) -- (#1+0.42,#2-0.14) -- (#1+0.56,#2-0.14);
  \fill[headgreen] (#1+0.07,#2-0.08) rectangle (#1+0.36,#2-0.15);
  \draw[line width=0.5pt] (#1+0.28,#2-0.36) circle (0.08);
  \foreach \y in {0.54, 0.6} \draw[line width=0.4pt] (#1+0.07,#2-\y) -- (#1+0.49,#2-\y);}
```

Sheet pieces (each goes inside the SWI sheet frame: `\scalebox`, `tikzpicture`, bounding box, frame, `\heading`, `\nav{n}`, `\titleblock`):

```latex
% ---- agenda: M = 90 min from x = 0.5 to 11.5; a minute is 11/90 in
\foreach \a/\b/\t/\m in {0.5/2.9044/{Part 1\quad <Name>}/20 min, 2.9844/9.0156/{Part 2\quad <Name>}/50 min,
                         9.0956/11.5/{Part 3\quad <Name>}/20 min} {
  \draw[line width=0.5pt, fill=headgreen] (\a,6.0) rectangle (\b,6.32);
  \node[anchor=west, text=white, font=\sffamily\bfseries\normalsize, text height=1.6ex, text depth=0.45ex] at (\a,6.16) {\t};
  \node[anchor=east, text=white, font=\sffamily\bfseries\normalsize, text height=1.6ex, text depth=0.45ex] at (\b,6.16) {\m};
}
\foreach \a/\b/\t in {0.5/1.1111/{<Item>}, 1.1111/2.3333/{<Item>}} {
  \draw[line width=0.5pt, fill=white] (\a,5.3) rectangle (\b,5.86);
  \node[align=center, font=\sffamily\bfseries\footnotesize, inner sep=0pt] at ({(\a+\b)/2},5.58) {\t};
}
\draw[line width=0.5pt, fill=headgreen] (2.9444,5.3) rectangle (5.3889,5.86);        % the centerpiece
\node[text=white, font=\sffamily\bfseries\normalsize, inner sep=0pt] at (4.1667,5.58) {<Centerpiece>};
\foreach \x/\t in {0.5/0:00, 1.1111/0:05, 2.3333/0:15, 11.5/1:30} {
  \draw[line width=0.5pt] (\x,5.3) -- (\x,5.18);
  \node[anchor=north, font=\sffamily\bfseries\footnotesize, inner sep=0pt] at (\x,5.13) {\t};
}
\foreach \x in {2.9444, 9.0556} \draw[line width=0.5pt, dash pattern=on 4pt off 3pt] (\x,4.95) -- (\x,1.72);
\begin{scope}[shift={(0.5,{4.75-0.625})}, scale=0.083333, x=1in, y=1in] \thumbdesign \end{scope}
\node[anchor=north west, align=left, inner sep=0pt, font=\sffamily\footnotesize] at (1.62,4.72)
  {\textbf{Sheet 2}\\[2pt] Design Process\\[2pt] {\itshape 5 min}};
\node[anchor=north west, draw=headgreen, line width=1.6pt, fill=white, text width=2.05in, align=left, inner sep=7pt]
  at (3.12,4.8) {\badge{Centerpiece}\enspace{\footnotesize\itshape <why>}\\[6pt]
  {\bfseries\fontsize{20}{22}\selectfont <ID>}\\[8pt] \textbf{Live demo}\hfill{\footnotesize\itshape 12 min}\\[3pt]
  {\footnotesize\begin{enumerate}[label=\badge{\arabic*}, leftmargin=1.45em, labelsep=0.35em, itemsep=3pt]
  \item <Step> \item <Step> \end{enumerate}}};
\node[anchor=north west, draw, line width=0.5pt, fill=bandgreen, text width=10.82in, align=left] at (0.5,1.5)
  {\textbf{We leave with:} <the decisions>.};

% ---- architecture: the app as a container (outer layer's name on the bar, its role under it)
\node[station, anchor=north west, text width=3.86in] (ha) at (4.15,6.4)
  {<Outer layer>\hfill{\mdseries\footnotesize\itshape <Technology>}};
\node[anchor=north west, inner sep=0pt, font=\sffamily\footnotesize\itshape] (cap) at ($(ha.south west)+(0.12,-0.1)$) {<Role>};
\node[layer, text width=3.62in] (S) at ($(cap.south west)+(0,-0.1)$) {\layerbody{<Layer>}{<Tech>}{<Role>}};
\node[layer, text width=3.62in] (CE) at ($(S.south west)+(0,-0.1)$) {\layerbody{<Layer that writes the files>}{<Tech>}{<Role>}};
\node[layer, text width=3.38in] (G1) at ($(CE.south west)+(0.12,-0.44)$) {\layerbody{<Layer>}{<Tech>}{<Role>}};
\node[layer, text width=3.38in] (G2) at ($(G1.south west)+(0,-0.24)$) {\layerbody{<Layer>}{<Tech>}{<Role>}};
\draw[line width=0.5pt, draw=headgreen, rounded corners=3pt] ($(CE.south west)+(0,-0.32)$) rectangle ($(G2.south east)+(0.12,-0.14)$);
\node[anchor=west, fill=white, inner xsep=3pt, inner ysep=0pt, text=headgreen, font=\sffamily\bfseries\scriptsize]
  at ($(G2.south west)+(0,-0.14)$) {<GROUP NAME>};
\draw[line width=0.5pt] (ha.south west) rectangle ($(G2.south east)+(0.24,-0.28)$);
\draw[flow] (G2.north) -- (G1.south);                          % one-way flow, straight up the middle
\draw[flow] ($(CE.south)+(-0.2,0)$) -- ($(CE.south)+(-0.2,-0.32)$);   % an exchange: two arrows, one each way
\draw[flow] ($(CE.south)+(0.2,-0.32)$) -- ($(CE.south)+(0.2,0)$);
\node[anchor=east, inner sep=0pt, font=\sffamily\scriptsize] at ($(CE.south)+(-0.29,-0.16)$) {<sent>};
\node[anchor=west, inner sep=0pt, font=\sffamily\scriptsize] at ($(CE.south)+(0.29,-0.16)$) {<returned>};
\draw[flow] (3.48,0 |- S.west) -- (S.west);                    % in, from the files card (right edge 3.45)
\node[anchor=south, font=\sffamily\scriptsize, inner sep=0pt] at ($(3.8,0 |- S.west)+(0,0.07)$) {Input};
\coordinate (out) at ($(CE.north west)+(0,-0.15)$);            % out, straight into the files card
\draw[flow] (out) -- (3.48,0 |- out);
\node[anchor=north, font=\sffamily\scriptsize, inner sep=0pt, align=center] at ($(3.8,0 |- out)+(0,-0.07)$)
  {<Short>\\<stacked>\\<lines>};

% ---- a table under its title bar, with no white strip
\node[station, anchor=north west, text width=6.55in] (hi) at (0.25,3.95) {<Title>};
\node[box, anchor=north, inner ysep=0pt, text width=6.55in, font=\sffamily\footnotesize] at (hi.south) {%
  \setlength{\tabcolsep}{5pt}\renewcommand{\arraystretch}{1.7}%
  \begin{tabular}{@{}p{2.35in}p{4.05in}@{}}
  \rowcolor{bandgreen}\multicolumn{2}{@{}l@{}}{\bfseries\enspace <Group>} \\
  \btn{<Button>} & <What it does> \\
  \end{tabular}\par\vspace{5pt}};

% ---- value: a big figure opens each box
\node[box, anchor=north, text width=3.46in] at (h1.south) {\fixh{2.15in}{%
  \stat{\$0}{<caption>}
  \begin{itemize} \item <Fact> \end{itemize}}};
```


Add your content below this line. Name the tool first.
- **SWI:** the subject, the critical task and its cases, the facts, and each picture with its pixel size.
- **Boardroom Deck:** the project, the audience, the meeting length (if there is a meeting) and its centerpiece, the items to show, and the owner's notes (requirements, design steps, tests and results, components, costs, next steps).
