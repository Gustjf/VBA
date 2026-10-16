# Gust SWI Template v2: Reusable Prompt

Version 2, 2026-10-16. A style guide for Standard Work Instructions on any topic.

Paste everything below into another AI. Add your content at the end: the subject, the critical task, and the pictures you have (file names and pixel sizes).

---

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
- `alertorange` `#C55A11`: only on the critical-task sheet (frame, banner, markers, pointers).

## Building blocks (use the macros in the template)
- **Heading:** `\heading{Title}{One-sentence subtitle}` at the top left. The title is Title Case, under about 40 characters.
- **Station:** a green header bar (`station`), number then title, with a white `box` under it. Station titles start with a verb or name a thing ("Control panel", "Check the result"). Join stations with `flow` arrows, left to right or top to bottom.
- **Box contents:** an italic one-line tag via `\by{...}`, then bullets (`itemize`) or numbered steps (`enumerate`). Bullets are short fragments; steps are imperative sentences.
- **Picture:** `\node[shot] at (x,y) {\includegraphics[width=W in]{file.png}};`, anchored at its top left. Its height is H = W x (pixel height / pixel width). Write W, H and the pixel size in a comment above the sheet.
- **Markers on a picture:** open a scope over the picture, `\begin{scope}[shift={(x,y)}, x=W in, y=-H in]`. Inside it, (0,0) is the top left of the picture and (1,1) its bottom right. Place each `\node[marker] at (fx,fy) {1};` from the pixel position: fx = px / width, fy = py / height. Place a marker in empty space beside the cell or header it points to, never over its text. Repeat the same number or letter in the text with `\num{1}` (or `\hotnum{1}` on the critical sheet).
- **Buttons and settings:** `\btn{Button name}` for buttons, `\texttt{}` for files and settings, `\val{}` for values typed exactly. `\yellow{}` and `\peach{}` highlight inline words in those colors.
- **Reference table:** a `tabular` inside a node, `\arraystretch` about 1.4. Group the rows under full-width `\cellcolor` band rows by severity (e.g. "Stops the work", "Work can continue"). Put each symptom or message, quoted exactly, in the first column (`\ttfamily`, or bold for its key words) and the fix in the second, as an imperative sentence.
- **Pointer:** the `point` arrow (orange) links a marker to more than one place.
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
- Right column: an "after" picture for each step that changes something visible. Put orange `hot` markers on it with the step's number, and orange rectangles around the places a step changes.
- Under the steps: a `flagpeach` box beginning **Do not:** that lists the mistakes that break things, in one or two lines.

## Writing
- Present tense, plain words, the reader's own names for things, the same name for a thing everywhere.
- Say what the reader does and what they will see. Skip how it works inside, unless the reader acts on it.
- Every fact comes from the real file or code. Any example value is called an example.
- Nothing may be cut off or overlap. Compile, look at every sheet, and move or shrink until it all fits.

## Template

```latex
% <Subject> SWI: <what it covers>. Gust SWI Template v2.
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
\tikzset{
  every node/.style={font=\sffamily\small, inner sep=5pt},
  box/.style={draw, line width=0.5pt, fill=white, align=left},
  station/.style={draw, line width=0.5pt, fill=headgreen, text=white, font=\sffamily\bfseries\normalsize,
                  align=left, inner ysep=4pt, text height=1.6ex, text depth=0.45ex},
  flow/.style={-{Stealth[length=9pt, width=6.5pt]}, line width=1.2pt, line join=miter},
  point/.style={-{Stealth[length=6pt, width=4.5pt]}, line width=0.9pt, alertorange},
  marker/.style={circle, fill=headgreen, text=white, draw=white, line width=1pt, font=\sffamily\bfseries\small,
                 inner sep=0pt, minimum size=0.25in},
  hot/.style={marker, fill=alertorange},
  shot/.style={draw, line width=0.5pt, inner sep=0pt, anchor=north west},
}
\newcommand{\by}[1]{{\footnotesize\itshape #1}\\[3pt]}
\newcommand{\peach}[1]{{\setlength{\fboxsep}{1pt}\colorbox{flagpeach}{\textbf{#1}}}}
\newcommand{\yellow}[1]{{\setlength{\fboxsep}{1pt}\colorbox{inputyellow}{#1}}}
\newcommand{\btn}[1]{{\setlength{\fboxsep}{1.5pt}\fbox{\footnotesize\bfseries #1}}}
\newcommand{\num}[1]{\tikz[baseline=(n.base)]\node(n)[marker, minimum size=0.2in, font=\sffamily\bfseries\footnotesize]{#1};}
\newcommand{\hotnum}[1]{\tikz[baseline=(n.base)]\node(n)[hot, minimum size=0.2in, font=\sffamily\bfseries\footnotesize]{#1};}
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
\node[station, anchor=north west, text width=4.45in] (s1) at (0.25,6.45) {\hotnum{1}\enspace <Step one>};
\node[station, anchor=north west, text width=4.45in] (s2) at ($(s1.south west)+(0,-0.12)$) {\hotnum{2}\enspace <Step two>};
\node[box, anchor=north, text width=4.45in, font=\sffamily\footnotesize] (t2) at (s2.south) {%
  \begin{tabular}{@{}p{0.5in}p{1.5in}p{1.5in}@{}}
  \textbf{<A>} & <where> & e.g.\ \val{<value>} \\
  \textbf{<B>} & <where> & e.g.\ \val{<value>} \\
  \end{tabular}};
\node[station, anchor=north west, text width=4.45in] (s3) at ($(t2.south west)+(0,-0.12)$) {\hotnum{3}\enspace <Check it worked>};
\node[box, anchor=north, text width=4.45in] (t3) at (s3.south) {<What to expect, word for word.>};
\node[anchor=north west, draw, line width=0.5pt, fill=flagpeach, text width=4.45in, align=left,
      font=\sffamily\footnotesize] at ($(t3.south west)+(0,-0.2)$) {\textbf{Do not:} <the mistakes that break it>.};
\node[shot] at (5.15,6.0) {\includegraphics[width=6.6in]{<after-picture>.png}};
\begin{scope}[shift={(5.15,6.0)}, x=6.6in, y=-<H>in]
  \node[hot] at (<fx>,<fy>) {2};
\end{scope}
\titleblock{<Task>}{<N> of <N>}
\end{tikzpicture}}
\end{document}
```

Add your content below this line: the subject, the critical task and its cases, the facts, and each picture with its pixel size.
