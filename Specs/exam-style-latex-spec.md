# Exam-Style LaTeX Format — Reusable Prompt

Paste everything below into another AI, then add your content at the end.

---

Write a one-page LaTeX document in the style of an engineering or university exam, adapted for business use: plain, black and white, and easy to read. Follow this specification exactly.

## Page and fonts
- `\documentclass[11pt]{article}`, US Letter.
- Margins: `top=0.8in, bottom=0.8in, left=1in, right=1in, headheight=14pt` (geometry package).
- Font: Latin Modern via `\usepackage[T1]{fontenc}` and `\usepackage{lmodern}`. It looks like Computer Modern, but every glyph (including `\S`, `\textdegree`, `>`) is an outline font; leaving these out can embed bitmap fonts.
- Load `microtype` for even word spacing in justified text.
- Line spacing: `\linespread{1.05}`.
- Black and white only. No colors, no icons, no shaded boxes.
- Must fit on exactly one page; tighten wording before cutting content. Check the compiled PDF is 1 page with no overfull boxes.
- `\parindent` 0pt, `\parskip` 0pt.

## Header and footer (fancyhdr)
- Left header: `Strategy`, `\small`. This is the department/series name; do not repeat the document title here.
- Right header: `Rev.~N \quad YYYY-MM-DD`, `\small`. N is the revision number (start at 1, increase by 1 for each issued change); the date is the date of that revision.
- Left footer: author name (e.g. `Gust F.`), `\small`. The author goes here only, not in the title block.
- Center footer: `Page \thepage\ of \pageref{LastPage}` (lastpage package), `\small`.
- Header rule 0.4pt.

## Title block
- Centered: title in `\Large\bfseries`, subtitle in `\normalsize`. No author or date line (both are in the header/footer).
- Followed by a full-width 0.8pt rule.

## Opening block
- A paragraph starting with a bold run-in label: **Scope.** for procedures and reference documents, **Instructions.** for forms, tests, and checklists. State what the document covers and where to start. Keep it to 3–4 lines.
- Followed by a full-width 0.4pt rule.

## Sections
- Numbered sections: `\normalsize\bfseries`, label as `1.`, `2.`, … (titlesec).
- Spacing: `\titlespacing*{\section}{0pt}{0.9em}{0.3em}`.
- Procedures (ordered steps): enumerate lists labeled `(a)`, `(b)`, `(c)`: `leftmargin=2.2em, itemsep=0.1em, topsep=0.15em, ref=\alph*`. Every step is an imperative sentence ending in a period ("Apply…", "Wash…", "Allow to dry.").
- Non-sequential content (features, safeguards, requirements): itemize lists with en-dash bullets, same spacing. Descriptive sentences are fine here.
- Keep items short enough to fit on one line where possible; split or trim rather than wrap.
- Use consistent terminology and full product/item names everywhere.
- Cross-reference with `\label`/`\ref`, never typed numbers: `see Section~\ref{sec:install}`, `as in Section~\ref{sec:install}(\ref{step:import})` (renders "Section 1(c)"), `Table~\ref{tab:items}`.
- Menu paths and button names are set in italics with `>` separators using the `\menu{}` and `\gt` macros (e.g. `\menu{View\gt Macros\gt Record Macro}`). File names, extensions, and code/macro names are set with `\texttt{}`. Keyboard shortcuts are plain text (`Alt+F11`).

## Tables
- Centered, `\small`, full grid: `|` vertical dividers between every column and `\hline` between body rows. Put a double rule under the header row with `\hhline{|=|=|…|}` (one `=|` per column).
- Header row in bold. Text cells are left-aligned (`L` column type) and vertically centered; center only short values such as numbers or Y/N. An optional first column of code or item names is left-aligned monospace (`\footnotesize\ttfamily`, fixed width via `m{1.75in}`; size it to the longest name).
- Caption above via `\captionof{table}{Sentence-case description.}\label{tab:…}`, which renders as **Table 1.** Sentence-case description. Keep the table in a `center` environment (not a float) so it stays where it is placed.
- Notes below are optional; omit the note unless an abbreviation genuinely needs defining. Form: *Note.* Abbreviation definitions.
- Use `tabularx` at `\linewidth`; `\arraystretch` 1.25.
- Keep one idea per column (e.g. separate "Output sheet" and "Exception area" columns rather than one combined cell).

## PDF metadata
- Load `hyperref` last, with `hidelinks` (no colored link boxes), `pdftitle` set to the document title, and `pdfauthor` set to the author.

## Packages
`geometry, fontenc (T1), lmodern, microtype, enumitem, titlesec, tabularx, array, caption, hhline, fancyhdr, lastpage, hyperref`

Compile with `pdflatex` twice so the page count and cross-references resolve.

## Template

```latex
\documentclass[11pt]{article}
\usepackage[letterpaper, top=0.8in, bottom=0.8in, left=1in, right=1in, headheight=14pt]{geometry}
\usepackage[T1]{fontenc}
\usepackage{lmodern}
\usepackage{microtype}
\usepackage{enumitem}
\usepackage{titlesec}
\usepackage{tabularx}
\usepackage{array}
\usepackage{caption}
\usepackage{hhline}
\usepackage{fancyhdr}
\usepackage{lastpage}
\usepackage[hidelinks, pdftitle={DOCUMENT TITLE}, pdfauthor={Author Name}]{hyperref}

\linespread{1.05}

\pagestyle{fancy}
\fancyhf{}
\fancyhead[L]{\small Strategy}
\fancyhead[R]{\small Rev.~1 \quad YYYY-MM-DD}
\fancyfoot[L]{\small Author Name}
\fancyfoot[C]{\small Page \thepage\ of \pageref{LastPage}}
\renewcommand{\headrulewidth}{0.4pt}

\titleformat{\section}{\normalsize\bfseries}{\thesection.}{0.6em}{}
\titlespacing*{\section}{0pt}{0.9em}{0.3em}
\setlist[enumerate]{leftmargin=2.2em, itemsep=0.1em, topsep=0.15em, label=(\alph*), ref=\alph*}
\setlist[itemize]{leftmargin=2.2em, itemsep=0.1em, topsep=0.15em, label=--}
\captionsetup[table]{font=small, labelfont=bf, labelsep=period, skip=0.4em, position=top, hypcap=false}
\setlength{\parindent}{0pt}
\setlength{\parskip}{0pt}

\newcommand{\menu}[1]{\textit{#1}}
\newcommand{\gt}{~>\ }
\renewcommand{\tabularxcolumn}[1]{m{#1}}
\newcolumntype{L}{>{\raggedright\arraybackslash}X}

\begin{document}

\begin{center}
  {\Large\bfseries DOCUMENT TITLE}\\[0.35em]
  {\normalsize Subtitle}
\end{center}

\vspace{-0.4em}
\noindent\rule{\linewidth}{0.8pt}

\vspace{0.2em}
\textbf{Scope.} What this document covers and where to start (see Section~\ref{sec:first} and Table~\ref{tab:items}).

\vspace{-0.2em}
\noindent\rule{\linewidth}{0.4pt}

\section{First Section}\label{sec:first}
\begin{enumerate}
  \item Imperative step one.\label{step:one}
  \item Imperative step two.
\end{enumerate}

\begin{center}
\small
\captionof{table}{Description of table.}\label{tab:items}
\renewcommand{\arraystretch}{1.25}
\begin{tabularx}{\linewidth}{|>{\footnotesize\ttfamily\raggedright\arraybackslash}m{1.75in}|*{3}{L|}}
  \hline
  \multicolumn{1}{|l|}{\bfseries Item} & \bfseries A & \bfseries B & \bfseries C \\
  \hhline{|=|=|=|=|}
  ItemName & Text & Text & Text \\
  \hline
\end{tabularx}
\end{center}

\section{Reference Section}
\begin{itemize}
  \item Descriptive, non-sequential point; refer back to steps as in Section~\ref{sec:first}(\ref{step:one}).
\end{itemize}

\end{document}
```

## My content
[Paste your content here.]
