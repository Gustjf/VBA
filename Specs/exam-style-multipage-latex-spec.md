# Exam-Style LaTeX Format, Multi-Page — Reusable Prompt

Paste everything below into another AI, then add your content at the end. For a one-page document, use `exam-style-latex-spec.md` instead.

---

Write a multi-page LaTeX document in the style of an engineering or university exam, adapted for business use: plain, black and white, and easy to read and navigate across many pages. Follow this specification exactly.

## Page and fonts
- `\documentclass[11pt]{article}`, US Letter.
- Margins: `top=0.9in, bottom=0.9in, left=1in, right=1in, headheight=14pt` (geometry package).
- Font: Latin Modern via `\usepackage[T1]{fontenc}` and `\usepackage{lmodern}`. Every glyph (including `\S`, `\textdegree`, `>`) is then an outline font.
- Load `microtype` for even word spacing in justified text.
- Line spacing: `\linespread{1.1}`.
- Paragraphs: `\parindent` 0pt, `\parskip` 0.5em, so paragraphs are separated by space rather than indented.
- No widows or orphans: `\widowpenalty=10000`, `\clubpenalty=10000`. Use `\raggedbottom` so pages end naturally instead of stretching.
- Black and white only. No colors, no icons, no shaded boxes.
- Do not cut content to save pages. Readability comes first. Check that the compiled PDF has no overfull boxes and no `??` references.

## Header and footer (fancyhdr)
- Left header: `Strategy`, `\small`. This is the department/series name; do not repeat the document title here.
- Right header: `Rev.~N \quad YYYY-MM-DD`, `\small`. N is the revision number (start at 1, increase by 1 for each issued change); the date is the date of that revision.
- Left footer: author name (e.g. `Gust F.`), `\small`. The author goes here only, not in the title block.
- Center footer: the current section number and title (`\nouppercase{\leftmark}`), `\small`, so a reader always knows where they are. Keep section titles under about 45 characters.
- Right footer: `Page \thepage\ of \pageref{LastPage}` (lastpage package), `\small`, on the outer edge where it is easy to find when flipping pages.
- Header rule 0.4pt.

## First page
- Title block, centered: title in `\Large\bfseries`, subtitle in `\normalsize`. No author or date line (both are in the header/footer). Followed by a full-width 0.8pt rule.
- Opening block: a paragraph starting with a bold run-in label: **Scope.** for procedures and reference documents, **Instructions.** for plans, forms, tests, and checklists. State what the document covers and which section to start with. Keep it to 6 lines or fewer. Followed by a full-width 0.4pt rule.
- Contents: a table of contents listing sections only (`tocdepth` 1), set in `\small`, followed by `\markboth{}{}` so the first page footer shows no section name.
- Section 1 starts on page 2.

## Sections
- **Every numbered section starts on a new page** (`\newcommand{\sectionbreak}{\clearpage}` with titlesec). Never add manual `\clearpage` before a section.
- Section headings: `\large\bfseries`, labeled `1.`, `2.`, …, with a 0.4pt rule under the title.
- Subsections: `\normalsize\bfseries`, numbered `1.1`, `1.2`, …. Subsubsections: `\normalsize\bfseries\itshape`, numbered `1.1.1`. Never type numbers into heading text.
- Keep headings with their content: every subsection and subsubsection format starts with `\Needspace{…\baselineskip}` so a heading never sits alone at the bottom of a page.
- An optional lead paragraph directly under a section heading says what the section covers in 1–2 sentences.

## Lists and steps
- Procedures (ordered steps): enumerate lists labeled `(a)`, `(b)`, `(c)`: `leftmargin=2.2em, itemsep=0.25em, topsep=0.3em, parsep=0pt, ref=\alph*`. All lists set `before={\@endparpenalty\@M}` so a page never breaks directly after a list; whatever follows keeps at least the list's last item with it. Every step is an imperative sentence ending in a period ("Apply…", "Wash…", "Allow to dry.").
- Non-sequential content (features, safeguards, requirements): itemize lists with en-dash bullets, same spacing. Descriptive sentences are fine here.
- A named group of steps (a session, a phase, "How to hem") gets a stand-alone bold heading line via `\block{…}`, which reserves space and stays with the list below it. Do not end the heading with a colon.
- A recurring closing check or note uses `\closing{Pass check.}{Both legs within 1/8 inch.}`: a bold run-in label ending in a period, kept on the same page as the list above it so it never starts a page alone.
- Use consistent terminology and full product/item names everywhere.
- Cross-reference with `\label`/`\ref`, never typed numbers or section names: `see Section~\ref{sec:install}`, `as in Section~\ref{sec:install}(\ref{step:import})` (renders "Section 1(c)"), `Table~\ref{tab:items}`.
- Menu paths and button names are set in italics with `>` separators using the `\menu{}` and `\gt` macros (e.g. `\menu{View\gt Macros\gt Record Macro}`). File names, extensions, and code/macro names are set with `\texttt{}`. Keyboard shortcuts are plain text (`Alt+F11`).

## Tables
- Use `xltabular` at `\linewidth` (tabularx columns that can break across pages); `\small`; `\arraystretch` 1.3.
- Full grid: `|` vertical dividers between every column and `\hline` after every body row. Put a double rule under the header row with `\hhline{|=|=|…|}` (one `=|` per column).
- Header row in bold, repeated at the top of every continuation page under a small italic line *Table N, continued* (`\endfirsthead` / `\endhead`).
- End the last row with `\\* \hline` so the closing rule never moves to the next page alone.
- Do not split a table that fits on one page: put `\Needspace{N\baselineskip}` before it, with N its estimated height in lines (caption, header and wrapped rows). A table that fills most of a page starts on a new page together with its heading. Only a table longer than a page breaks, and never with a single row on either side.
- Column widths: `Y{f}` (left-aligned) and `C{f}` (centered) column types, where the `f` factors sum to the number of columns. Text cells are left-aligned and vertically centered; center only short values such as numbers, prices, times or Y/N.
- Caption above as the table's first row: `\caption{Sentence-case description.}\label{tab:…}\\`, which renders as **Table 1.** Sentence-case description.
- Notes below are optional; omit the note unless an abbreviation genuinely needs defining. Form: *Note.* Abbreviation definitions.
- Keep one idea per column.
- Fill-in tables (blank rows for writing) use `\arraystretch` 1.75–2.1. A blank form that fills a page starts on a new page.

## PDF metadata and navigation
- Load `hyperref` last, with `hidelinks` (no colored link boxes), `bookmarksnumbered`, `bookmarksopen`, `pdftitle` set to the document title, and `pdfauthor` set to the author. The PDF gets a clickable bookmark outline and the contents and cross-references are clickable.

## Packages
`geometry, fontenc (T1), lmodern, microtype, enumitem, titlesec, tabularx, array, longtable, xltabular, caption, hhline, needspace, fancyhdr, lastpage, hyperref`

Compile with `pdflatex` three times so the page count, contents and cross-references resolve.

## Template

```latex
\documentclass[11pt]{article}
\usepackage[letterpaper, top=0.9in, bottom=0.9in, left=1in, right=1in, headheight=14pt]{geometry}
\usepackage[T1]{fontenc}
\usepackage{lmodern}
\usepackage{microtype}
\usepackage{enumitem}
\usepackage{titlesec}
\usepackage{tabularx}
\usepackage{array}
\usepackage{longtable}
\usepackage{xltabular}
\usepackage{caption}
\usepackage{hhline}
\usepackage{needspace}
\usepackage{fancyhdr}
\usepackage{lastpage}
\usepackage[hidelinks, bookmarksnumbered, bookmarksopen, pdftitle={DOCUMENT TITLE}, pdfauthor={Author Name}]{hyperref}

\linespread{1.1}
\setlength{\parindent}{0pt}
\setlength{\parskip}{0.5em}
\widowpenalty=10000
\clubpenalty=10000
\raggedbottom

\pagestyle{fancy}
\fancyhf{}
\fancyhead[L]{\small Strategy}
\fancyhead[R]{\small Rev.~1 \quad YYYY-MM-DD}
\fancyfoot[L]{\small Author Name}
\fancyfoot[C]{\small\nouppercase{\leftmark}}
\fancyfoot[R]{\small Page \thepage\ of \pageref{LastPage}}
\renewcommand{\headrulewidth}{0.4pt}
\renewcommand{\sectionmark}[1]{\markboth{\thesection.\ #1}{}}

\setcounter{secnumdepth}{3}
\setcounter{tocdepth}{1}
\newcommand{\sectionbreak}{\clearpage}
\titleformat{\section}{\large\bfseries}{\thesection.}{0.6em}{}[\vspace{0.2em}{\titlerule[0.4pt]}]
\titlespacing*{\section}{0pt}{0pt}{0.8em}
\titleformat{\subsection}{\Needspace{8\baselineskip}\normalsize\bfseries}{\thesubsection}{0.6em}{}
\titlespacing*{\subsection}{0pt}{1.2em}{0.3em}
\titleformat{\subsubsection}{\Needspace{6\baselineskip}\normalsize\bfseries\itshape}{\thesubsubsection}{0.6em}{}
\titlespacing*{\subsubsection}{0pt}{0.8em}{0.2em}

% no page break directly after a list, so a closing check or note stays with the list's last item
\makeatletter
\setlist{parsep=0pt, before={\@endparpenalty\@M}}
\makeatother
\setlist[enumerate]{leftmargin=2.2em, itemsep=0.25em, topsep=0.3em, label=(\alph*), ref=\alph*}
\setlist[itemize]{leftmargin=2.2em, itemsep=0.25em, topsep=0.3em, label=--}

\captionsetup[table]{font=small, labelfont=bf, labelsep=period, skip=0.4em, position=top, hypcap=false}
\renewcommand{\tabularxcolumn}[1]{m{#1}}
\newcolumntype{Y}[1]{>{\raggedright\arraybackslash\hsize=#1\hsize}X}
\newcolumntype{C}[1]{>{\centering\arraybackslash\hsize=#1\hsize}X}
\renewcommand{\arraystretch}{1.3}
\setlength{\LTpre}{0.8em}
\setlength{\LTpost}{0.8em}

\newcommand{\block}[1]{\par\Needspace{5\baselineskip}\vspace{0.3em}\textbf{#1}\par\nopagebreak}
\newcommand{\closing}[2]{\par\textbf{#1} #2\par}
\newcommand{\menu}[1]{\textit{#1}}
\newcommand{\gt}{~>\ }

\begin{document}

\begin{center}
  {\Large\bfseries DOCUMENT TITLE}\\[0.35em]
  {\normalsize Subtitle}
\end{center}

\vspace{-0.4em}
\noindent\rule{\linewidth}{0.8pt}

\textbf{Scope.} What this document covers and where to start (see Section~\ref{sec:first}).

\noindent\rule{\linewidth}{0.4pt}

{\small\let\sectionbreak\relax\tableofcontents}
\markboth{}{}

\section{First Section}\label{sec:first}
One or two sentences on what this section covers.

\subsection{A Procedure}
\block{How to do the task}
\begin{enumerate}
  \item Imperative step one.\label{step:one}
  \item Imperative step two.
\end{enumerate}
\closing{Check.}{What a correct result looks like.}

\Needspace{8\baselineskip}
\begingroup\small
\begin{xltabular}{\linewidth}{|Y{0.8}|Y{1.6}|C{0.6}|}
\caption{Description of table.}\label{tab:items}\\
\hline
\textbf{Item} & \textbf{Description} & \textbf{Qty} \\
\hhline{|=|=|=|}
\endfirsthead
\multicolumn{3}{l}{\small\itshape Table~\ref{tab:items}, continued}\\
\hline
\textbf{Item} & \textbf{Description} & \textbf{Qty} \\
\hhline{|=|=|=|}
\endhead
Item one & Text & 1 \\ \hline
Item two & Text & 2 \\* \hline
\end{xltabular}
\endgroup

\section{Reference Section}
\begin{itemize}
  \item Descriptive, non-sequential point; refer back to steps as in Section~\ref{sec:first}(\ref{step:one}).
\end{itemize}

\end{document}
```

## My content
[Paste your content here.]
