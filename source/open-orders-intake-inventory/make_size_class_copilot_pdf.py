#!/usr/bin/env python3
"""
Write size-class-copilot-prompt.tex, the one-page PDF of Copilot prompt P11 (build the size class file from the
PowerPoint slide), with the four prompt boxes copied word for word from
output/open-orders-intake-inventory/copilot-prompts.md (a tab is shown as <TAB>).
Run it again whenever P11 changes, then: ./build.sh source/open-orders-intake-inventory/size-class-copilot-prompt.tex

Usage: python3 make_size_class_copilot_pdf.py [out.tex]
"""
import os, re, sys
HERE = os.path.dirname(os.path.abspath(__file__))
md = open(os.path.join(HERE, '..', '..', 'output', 'open-orders-intake-inventory', 'copilot-prompts.md'), encoding='utf-8').read()
p11=md[md.index('## P11.'):md.index('## P1. Headers')]
blocks=[b.replace('\t',' <TAB> ') for b in re.findall(r'```text\n(.*?)```',p11,re.S)]
assert len(blocks)==4
def box(t): return '\\begin{prompt}\n'+t.rstrip('\n')+'\n\\end{prompt}'
T=r'''\documentclass[11pt]{article}
\usepackage[letterpaper, top=0.5in, bottom=0.5in, left=0.5in, right=0.5in, headheight=14pt, includeheadfoot]{geometry}
\usepackage[T1]{fontenc}
\usepackage{lmodern}
\usepackage{microtype}
\usepackage{enumitem}
\usepackage{titlesec}
\usepackage{multicol}
\usepackage{fvextra}
\usepackage{fancyhdr}
\usepackage{lastpage}
\usepackage[hidelinks, pdftitle={Open Orders Report: Build the Size Class File in Copilot (P11)}, pdfauthor={Gust F.}]{hyperref}

\pagestyle{fancy}
\fancyhf{}
\fancyhead[L]{\small Strategy}
\fancyhead[R]{\small Rev.~2 \quad 2026-10-15}
\fancyfoot[L]{\small Gust F.}
\fancyfoot[C]{\small Page \thepage\ of \pageref{LastPage}}
\renewcommand{\headrulewidth}{0.4pt}

\titleformat{\section}{\small\bfseries}{Step \thesection.}{0.5em}{}
\titlespacing*{\section}{0pt}{0.35em}{0.1em}
\setlist[enumerate]{leftmargin=1.8em, itemsep=0em, topsep=0.1em, label=(\alph*), ref=\alph*}
\setlength{\parindent}{0pt}
\setlength{\parskip}{0pt}
\setlength{\columnsep}{0.25in}

\newcommand{\menu}[1]{\textit{#1}}
\newcommand{\gt}{~>\ }
% Prompt text, copied word for word from copilot-prompts.md P11 (tabs shown as <TAB>): framed, wrapped, ready to copy and paste.
\DefineVerbatimEnvironment{prompt}{Verbatim}{breaklines, breakanywhere=false, frame=single, framesep=0.3em,
  fontsize=\fontsize{6.1}{6.8}\selectfont, fontfamily=lmss, breaksymbolleft={}, breakindent=0.8em}
\newcommand{\checknote}[1]{\par\vspace{0.15em}{\footnotesize\textbf{Check.} #1}\par}

\begin{document}

\begin{center}
  {\large\bfseries Build the Size Class File in Copilot (P11)}\\[0.25em]
  {\small From the PowerPoint slide to the workbook Config \texttt{SizeClassFile} points at}
\end{center}

\vspace{-0.7em}
\noindent\rule{\linewidth}{0.8pt}

{\footnotesize\textbf{Instructions.} Do this first: it is the most important Copilot prompt. The macro builds nothing without a size class file, and a model missing from it lands in UNMAPPED. The slide has the latest classes, but not in the macro's layout or MODEL\_GROUP spelling. Run Steps~\ref{sec:s1} to \ref{sec:s3} \textbf{in one Copilot chat}, in order, attaching the file each one names; Step~\ref{sec:s4} is optional. Paste each boxed prompt as it is; where it shows <TAB>, a tab is meant. If Copilot loses the earlier table, start a new chat and paste it under the prompt.\par}

\vspace{-0.1em}
\noindent\rule{\linewidth}{0.4pt}

\begin{multicols}{2}\footnotesize
\section{Read the slide (attach the PowerPoint file)}\label{sec:s1}
@B0@
\checknote{Count the models on the slide and compare with MODELS ON SLIDE. Each CLASS LABELS line must point to the class you expect; if not, correct Copilot in the chat (e.g.\ \texttt{"Under 2,000" is 0<2000}) and ask it to redo the table.}

\section{Check against Order Intake (attach it)}\label{sec:s2}
@B1@
\checknote{Every blank Size Class needs one before Step~\ref{sec:s3}: find the capacity on the spec sheet and tell Copilot, e.g.\ \texttt{T870 is 3500\&Over}. A model left blank stays UNMAPPED.}

\section{The finished table (same chat)}\label{sec:s3}
@B2@
\vspace{0.1em}
{\footnotesize If your Copilot can create files, ask instead: \textit{Create an Excel workbook with this table on the first sheet, named Size Class, and Size Class formatted as text}; then go to step~(\ref{step:save}).}
\begin{enumerate}\footnotesize
  \item Open a new workbook; set column D (Size Class) to \menu{Home\gt Number Format\gt Text} before pasting.
  \item Paste into cell A1 of the \textbf{first sheet} (the macro reads only that one).
  \item Check that ROWS matches the rows pasted and no Size Class cell is empty.
  \item Save it as a local \texttt{.xlsx} (e.g.\ in \texttt{C:\textbackslash Temp\textbackslash oo-test}).\label{step:save}
  \item On Config, click \menu{Choose size class file\ldots}, pick it, then click \menu{Check setup}; fix anything it lists in the same chat and paste again.
\end{enumerate}

\section{Optional: what changed}\label{sec:s4}
{\footnotesize Attach \texttt{size-class-master.xlsx} to show what the slide changed.\par}
@B3@
\end{multicols}

\end{document}
'''
for i,b in enumerate(blocks): T=T.replace(f'@B{i}@',box(b))
out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, 'size-class-copilot-prompt.tex')
open(out, 'w', encoding='utf-8').write('% Generated by make_size_class_copilot_pdf.py from copilot-prompts.md P11; edit those, not this file.\n' + T)
print(out)
