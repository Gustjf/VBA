#!/usr/bin/env python3
"""
Write the LaTeX source of a printable copy of the size-class master list:
    output/open-orders-intake-inventory/size-class-master.xlsx
 -> source/open-orders-intake-inventory/size-class-master.tex
Then build the PDF with ./build.sh source/open-orders-intake-inventory/size-class-master.tex
 -> output/open-orders-intake-inventory/size-class-master.pdf

The list is read the way the macro reads it (first sheet, columns by header name), every usable row is
printed in file order, split into track loaders (T...), skid-steer loaders (S...) and mini excavators
(E...), with the class rules above each table. The note every excavator row repeats ("Rated Operating
Capacity column holds the operating weight; ...") is printed once above its table instead.
Run it again whenever the master list changes.

Usage: python3 make_size_class_pdf.py [size-class-master.xlsx] [out.tex]
"""
import datetime, os, sys
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_IN = os.path.join(HERE, '..', '..', 'output', 'open-orders-intake-inventory', 'size-class-master.xlsx')
DEFAULT_OUT = os.path.join(HERE, 'size-class-master.tex')
MX_NOTE = '; Rated Operating Capacity column holds the operating weight (lb)'
CHECK = 'REVIEW:'


def norm(h):
    return str(h or '').replace(' ', '').replace('_', '').upper()


def tex(s):
    s = '' if s is None else str(s)
    for a, b in (('\\', r'\textbackslash{}'), ('&', r'\&'), ('%', r'\%'), ('#', r'\#'), ('_', r'\_'),
                 ('$', r'\$'), ('~', r'\textasciitilde{}'), ('^', r'\textasciicircum{}')):
        s = s.replace(a, b)
    return s


def num(v):
    if v is None or v == '':
        return ''
    if isinstance(v, (int, float)):
        return f'{v:,.0f}' if float(v).is_integer() else f'{v:,.1f}'
    return tex(v)


def read_rows(path):
    ws = openpyxl.load_workbook(path, data_only=True).worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    hdr = [norm(h) for h in rows[0]]
    col = lambda *names: next((i for i, h in enumerate(hdr) if h in names), None)
    c_name = col('NAMEOFMACHINE', 'MACHINE', 'MODEL', 'MODELGROUP')
    c_cls = col('SIZECLASS', 'CLASS')
    c_roc = col('RATEDOPERATINGCAPACITY', 'RATEDOPERATORCAPACITY', 'ROC', 'OPERATINGCAPACITY')
    c_hp = col('HP')
    c_notes = col('NOTES')
    get = lambda r, c: r[c] if c is not None and c < len(r) else None
    out = []
    for r in rows[1:]:
        name = str(get(r, c_name) or '').strip()
        cls = str(get(r, c_cls) or '').strip()
        if name and cls:
            out.append((name, cls, get(r, c_roc), get(r, c_hp), str(get(r, c_notes) or '').strip()))
    return ws.title, out


def note_cell(note, family):
    if family == 'E' and note.endswith(MX_NOTE):
        note = note[:-len(MX_NOTE)]
    s = tex(note)
    return s.replace(tex(CHECK), r'\textbf{' + tex(CHECK) + '}')


def table(rows, family, figure_head, caption):
    lines = [r'\begin{longtable}{@{}p{0.75in}p{0.8in}>{\raggedleft\arraybackslash}p{0.75in}'
             r'>{\raggedleft\arraybackslash}p{0.4in}@{\hspace{1.2em}}p{3.1in}@{}}',
             r'\caption{' + caption + r'}\\',
             r'\toprule',
             r'\bfseries Machine & \bfseries Size Class & \bfseries ' + figure_head + r' & \bfseries HP & '
             r'\bfseries Notes \\',
             r'\midrule\endfirsthead',
             r'\toprule',
             r'\bfseries Machine & \bfseries Size Class & \bfseries ' + figure_head + r' & \bfseries HP & '
             r'\bfseries Notes \\',
             r'\midrule\endhead',
             r'\bottomrule\endlastfoot']
    prev = None
    for name, cls, roc, hp, note in rows:
        if prev is not None and cls != prev:
            lines.append(r'\midrule')
        prev = cls
        lines.append(f'{tex(name)} & {tex(cls)} & {num(roc)} & {num(hp)} & {note_cell(note, family)} \\\\')
    lines.append(r'\end{longtable}')
    return '\n'.join(lines)


def main():
    src = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_IN
    dst = sys.argv[2] if len(sys.argv) > 2 else DEFAULT_OUT
    sheet, rows = read_rows(src)
    fam = {k: [r for r in rows if r[0].upper().startswith(k)] for k in 'TSE'}
    other = [r for r in rows if r[0][:1].upper() not in 'TSE']
    n_check = sum(CHECK in r[4] for r in rows)
    today = datetime.date.today().isoformat()
    body = [
        r'\section{Compact track loaders (CTL)}',
        r'Classes by rated operating capacity (ROC) in lb: 0<2000, 2000<2300, 2300<2600, 2600<3000, 3000<3500, '
        r'3500\&Over. The macro reads each model\textquoteright s class from this list; the ROC column is the '
        r'owner\textquoteright s reference and is not read by the macro.',
        table(fam['T'], 'T', 'ROC (lb)', f'Track loaders, {len(fam["T"])} models.'),
        r'\section{Skid-steer loaders (SSL)}',
        r'The frame classes EX Small, Small Frame, Med Frame and Large Frame are the four SSL classes the macro '
        r'reads from this list.',
        table(fam['S'], 'S', 'ROC (lb)', f'Skid-steer loaders, {len(fam["S"])} models.'),
        r'\section{Mini excavators (MX)}',
        r'Classes by nominal tonnes: >00-01, >01-02, >02-03, >03-04, >04-05, >05-06, >06-10. The macro reads '
        r'each model\textquoteright s class from this list; a model missing from it is UNMAPPED until it is '
        r'added. The figure column holds the operating weight in lb (the file\textquoteright s Rated Operating '
        r'Capacity column); it is information only.',
        table(fam['E'], 'E', 'Weight (lb)', f'Mini excavators, {len(fam["E"])} models.'),
    ]
    if other:
        body += [r'\section{Other models}', table(other, 'O', 'Figure', f'Other, {len(other)} models.')]
    doc = PREAMBLE.replace('@DATE@', today) + '\n'.join([
        r'\begin{document}',
        r'\begin{center}',
        r'  {\Large\bfseries Open Orders Report: Size Class Master List}\\[0.35em]',
        r'  {\normalsize Printed copy of \texttt{' + tex(os.path.basename(src)) + r'}, sheet ``' + tex(sheet) + "''}",
        r'\end{center}',
        r'\vspace{0.4em}',
        r'{\small ' + f'{len(rows)} models. The macro reads this workbook as Config \\texttt{{SizeClassFile}}: '
        r'MODEL\_GROUP is matched to Machine (exact name, then the first word of the name); a model not on '
        r'the list is UNMAPPED. ' + f'{n_check} rows are marked \\textbf{{{tex(CHECK)}}}: '
        r'the owner confirms them. Rows, figures and classes are transcribed from the owner\textquoteright s '
        r'Bobcat Equipment Size Class sheet (photo, 2026-10-15); only the class spellings <2000 and '
        r'3500 \& Over are written the macro\textquoteright s way. HP is information only.}',
        *body,
        r'\end{document}', ''])
    with open(dst, 'w', encoding='utf-8', newline='\n') as f:
        f.write(doc)
    print(f'{dst}: {len(rows)} rows ({len(fam["T"])} CTL, {len(fam["S"])} SSL, {len(fam["E"])} MX, '
          f'{len(other)} other), {n_check} CHECK')


PREAMBLE = r"""% Generated by make_size_class_pdf.py from size-class-master.xlsx; edit the workbook, not this file.
\documentclass[10pt]{article}
\usepackage[letterpaper, top=0.7in, bottom=0.7in, left=0.75in, right=0.75in, headheight=14pt]{geometry}
\usepackage[T1]{fontenc}
\usepackage{lmodern}
\usepackage{microtype}
\usepackage{titlesec}
\usepackage{array}
\usepackage{longtable}
\usepackage{booktabs}
\usepackage{caption}
\usepackage{etoolbox}
\usepackage{fancyhdr}
\usepackage{lastpage}
\usepackage[hidelinks, pdftitle={Open Orders Report: Size Class Master List}, pdfauthor={Gust F.}]{hyperref}

\pagestyle{fancy}
\fancyhf{}
\fancyhead[L]{\small Strategy}
\fancyhead[R]{\small Size class master list \quad @DATE@}
\fancyfoot[L]{\small Gust F.}
\fancyfoot[C]{\small Page \thepage\ of \pageref{LastPage}}
\renewcommand{\headrulewidth}{0.4pt}

\titleformat{\section}{\normalsize\bfseries}{\thesection.}{0.6em}{}
\titlespacing*{\section}{0pt}{1.0em}{0.3em}
\captionsetup[table]{font=small, labelfont=bf, labelsep=period, skip=0.4em}
\setlength{\parindent}{0pt}
\setlength{\LTpre}{0.5em}
\setlength{\LTpost}{0.3em}
\renewcommand{\arraystretch}{1.1}
\AtBeginEnvironment{longtable}{\small}

"""

if __name__ == '__main__':
    main()
