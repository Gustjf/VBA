#!/usr/bin/env python3
"""Verify an Open Orders report (version 2.8 layout) against its data contract.

Usage:
    python3 check_report.py <report.xlsx> [--expected expected.json] [--run auto|default|variant|setup]
                            [--compare other.xlsx]

Output: one line per check, then a RESULT line:
    PASS <rule id> <sheet>
    FAIL <rule id> <sheet>: <reason, first offending cells>
    RESULT: PASS (<n> checks)   or   RESULT: FAIL (<k> of <n> checks failed)
Exit code: 0 when every check passes, 1 when any check fails, 2 on a usage error (missing or unreadable
file, bad option).

Rule ids (plan Section 7.2; R5a, R5b, E7 and C5 from open-orders-by-msa-tab-implementation-plan.md Section 7.2; R3g
and E8 from v2.6, the visible total; v2.7: the 17 size classes in three families, the size class file the only source
of a class, size-class-file-only-implementation-plan.md Section 7.3; v2.8: the classes of a build are the ones on its
File Format Table 1, from the size class file's Set-up tab or the macro's 17 built-in ones, and E9):
    R1    sheet names and order: File Format, Order Intake, Inventory by Month, Open Orders by MSA, Run Log
    R2    style, per sheet (v2.4; v2.6 moves the Order Intake note to H1): gridlines off; no theme colour; fills only where set: 548235 on every
          header cell (white bold text), FFFFCC on the month cell B3, FCE4D6 on the File Format UNMAPPED cell;
          every other border and font colour black; every written cell in Cambria (the Run Log in Consolas);
          title A1 14 bold, note D1 (Order Intake: H1) italic, a medium rule under row 1; on the data tabs every table ruled
          horizontally only (no vertical lines), a medium rule above and a double rule under the header, its
          body banded (a conditional format MOD(SUBTOTAL(103,...),2)=0 filling every second visible row
          E2EFDA); on Order Intake a first rule $A3="UNMAPPED" filling FCE4D6 ahead of the banding; columns
          aligned and indented as set, freeze panes, column widths (tolerance 0.3), number formats, UNMAPPED
          Size Class cells bold, the month cell boxed; print setup: landscape, one page wide, page header and
          footer, header rows repeated; on Open Orders by MSA the header A4:E4, widths 36 and 4 x 12, #,##0 on
          B5:E(5+nc), banding on the MSA rows only, one page tall, no title rows and no freeze panes
    R2F   File Format: the captions (Table 1. to Table 5.), the size class table (v2.8: n classes, 1 <= n <= 98, then
          UNMAPPED; each class row's Family and Rule text, the families in the order CTL, MX, SSL, no class twice, a
          medium rule above each family after the first, the UNMAPPED row bold and outside the banding; Table 2
          two rows under it, Tables 3 and 4 nine rows apart), the status order, default filter (five rows,
          the fifth the Open Orders by MSA chart) and build tables with their grid and banding, the months list;
          the build table's "Open orders" item equals this checker's own tally of the Order Intake rows
    R3a   Order Intake headers exact; every size class one of File Format Table 1's or UNMAPPED
    R3b   size class rank (the order of File Format Table 1, UNMAPPED last) never decreases
    R3c   inside a class, status rank never decreases; inside class and status, ORDER_DATE never increases and
          no date follows a non-date; inside class, status and day, the city never goes back
    R3d   PRODUCT_TYPE_ABBR is CTL, MX or SSL, trimmed and upper-cased
    R3e   hidden rows are exactly the CANCELLED rows (ShowCancelled FALSE) and the SSL rows (ShowSSL FALSE);
          AutoFilter ref A2:K<n+2> and its filter columns match the options
    R3f   date cells in C and D have a date format; thin rules between data rows; a medium rule above the
          first row of each size class (the double header rule above the first)
    R3g   the visible total in the frozen title row: F1 "Total sum of ordered Quantity of visible selection:" bold and
          right-aligned with B1:E1 empty (so it shows in full), G1 =SUBTOTAL(109,$G$3:$G$<2 + max(n, 1)>) bold,
          #,##0, left-aligned and indented (it reads on from the label); evaluated (the numbers of the rows not hidden), it equals this checker's own sum of
          column G over the visible rows; the note in H1
    R4a   inventory summary: month cell in the months list with its list validation; one row per city,
          alphabetical; every cell a SUMIFS over the full detail ranges with the right criteria; Total
          column and row formulas; a double rule above the Total row
    R4b   inventory detail: headers exact (MSA ... TOTAL, ORGANIZATION_NAME, ORGANIZATION_CUSTOMER_NUMBER); sorted
          by MSA, type order, date newest first; Month is the month of INVENTORY_DATE; hidden rows exactly the other months and the hidden types; AutoFilter and filter
          columns match; a medium rule above each MSA's first row
    R5a   Open Orders by MSA table: A1 title, D1 note, A3 caption, A4:E4 headers; one row per MSA (every MSA of the
          report, alphabetical, as the inventory summary); every B..D cell the two-SUMIFS formula over the full
          Order Intake ranges (rows 3 to 2 + max(n, 1)) with criteria $A r, <col>$4 and "ORDERED" / "Not yet
          processed"; E =SUM(B r:D r); the Total row formulas, bold, a double rule above it; evaluated, the values
          equal this checker's own tally of the open rows (ORDER_STATUS ORDERED or Not yet processed compared as
          SUMIFS compares: whole text, ignoring case, not ignoring spaces); "No MSAs: no chart." when nc = 0
    R5b   Open Orders by MSA chart, read from the .xlsx zip (not through openpyxl's chart reader): one drawing and one
          clustered column chart when nc > 0 (none when nc = 0) with its title, one series per MSA in table order
          (tx $A$r, cat $B$4:$C$4 or to $D$4 with ShowSSL, val $B$r:$C$r or to $D$r), anchored below the table
    R6    Run Log: A1 starts 'Built ', A2 options line parses, every MSA is named in a log line; every MSA has its
          open orders line (or "no open orders") with this checker's own figures, and the zeros, stray-space
          WARNING and IS_OPEN_ORDER lines are there exactly when the Order Intake rows call for them
    E0    --expected: a run is selected (by the options line, or --run)
    E1    --expected: sheet order
    E2    --expected: Order Intake rows [class, city, group, status, order date, hidden]
    E3    --expected: summary (city, CTL, MX, SSL, formulas evaluated) and the month cell
    E4    --expected: detail rows [city, type, group, month, total, hidden]
    E5    --expected: File Format build table, the file rows (files and city columns) and the Open orders item
    E6    --expected: Run Log rows equal log_lines (unified diff on mismatch)
    E7    --expected: Open Orders by MSA table values (formulas evaluated) and the chart's series count and categories
    E8    --expected: the Order Intake visible total G1 (evaluated over the rows the filters show)
    E9    --expected: the size classes on File Format Table 1, with their families, in order (v2.8)
    C1    --compare: same sheet names in the same order
    C2    --compare: Order Intake A..K and hidden flags equal row by row; F1 and G1 (formula text and value) equal
    C3    --compare: month cell, summary values (formulas evaluated with the data), detail A..G and hidden flags
    C4    --compare: Run Log A2 and rows 4.. equal line for line
    C5    --compare: Open Orders by MSA A4:E(5+nc) equal cell for cell (formula text and evaluated value); same
          chart series count, categories and anchor row
"""

import argparse
import datetime
import difflib
import json
import posixpath
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

from openpyxl import load_workbook
from openpyxl.styles.numbers import is_date_format
from openpyxl.utils import column_index_from_string, get_column_letter, range_boundaries


# ----------------------------------------------------------------- data contract

SHEETS = ['File Format', 'Order Intake', 'Inventory by Month', 'Open Orders by MSA', 'Run Log']
OO_SHEET = 'Open Orders by MSA'
OPEN_STATUSES = ('ORDERED', 'NOT YET PROCESSED')       # as written (SUMIFS ignores case, not spaces)
ABBRS = ['CTL', 'MX', 'SSL']
OO_WIDTHS = [36, 12, 12, 12, 12]
OO_CAPTION = ('Table 1. Open orders by MSA: sum of SUM(F.ORDERED_QUANTITY) over the Order Intake rows whose '
              'ORDER_STATUS is ORDERED or Not yet processed.')
OO_NOTE_RE = re.compile(r'Source: the Order Intake tab, built \d{4}-\d{2}-\d{2}\. Ordered quantity of orders with '
                        r'ORDER_STATUS ORDERED or Not yet processed, by MSA and product type\. Chart: (CTL and MX|CTL, MX and SSL)\.')
OO_CHART_TITLE = 'Open orders by product type and MSA (ordered quantity)'
OO_NO_CHART = 'No MSAs: no chart.'
OI_TOTAL_LABEL = 'Total sum of ordered Quantity of visible selection:'
OI_NOTE_RE = re.compile(r'Source: Snowflake Order Intake export, \d+ files?, built \d{4}-\d{2}-\d{2}\. .+')
# v2.8: a build's size classes are the rows of its File Format Table 1 (the size class file's Set-up tab, or the macro's
# 17 built-in classes), in rank order 1, 2, ...; UNMAPPED and anything else rank 99, so at most 98 classes
UNMAPPED = 'UNMAPPED'
UNKNOWN_RANK, MAX_CLASSES = 99, 98
# Table 1's families in report order: (abbreviation, Family text, Rule text)
FAMILIES = [('CTL', 'Track loaders (CTL)', 'by rated operating capacity in lb; the class is read from the size class file'),
            ('MX', 'Mini excavators (MX)', 'by nominal tonnes; the class is read from the size class file'),
            ('SSL', 'Skid-steer loaders (SSL)', 'by frame size; the class is read from the size class file')]
FAMILY_BY_TEXT = {f[1]: f for f in FAMILIES}
UNMAPPED_ROW = ['UNMAPPED', 'Any', 'model not in the size class file; add it there and build again']
STATUS_ORDER = ['ORDERED', 'Not yet processed', 'INVOICED', 'Completely processed', '(any other status)', 'CANCELLED']
STATUS_RANK = {'ORDERED': 1, 'NOT YET PROCESSED': 2, 'INVOICED': 3, 'COMPLETELY PROCESSED': 4, 'CANCELLED': 6}
TYPE_NAMES = ['Compact Track Loaders', 'Mini Excavators', 'Skid-Steer Loaders']
OI_HEADERS = ['Size Class', 'MSA', 'ORDER_DATE', 'REQUESTED_DATE', 'ORDER_STATUS', 'IS_OPEN_ORDER',
              'SUM(F.ORDERED_QUANTITY)', 'PRODUCT_TYPE_ABBR', 'MODEL_GROUP', 'MODEL_NUMBER', 'ORGANIZATION_NAME']
SUM_HEADERS = ['MSA', 'CTL', 'MX', 'SSL', 'Total']
INV_HEADERS = ['MSA', 'PRODUCT_TYPE', 'MODEL_GROUP', 'MODEL_NUMBER', 'INVENTORY_DATE', 'Month', 'TOTAL',
               'ORGANIZATION_NAME', 'ORGANIZATION_CUSTOMER_NUMBER']
MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September',
          'October', 'November', 'December']
FONT_NAME, LOG_FONT_NAME = 'Cambria', 'Consolas'
HEAD_FILL, BAND_FILL, FLAG_FILL, INPUT_FILL = '548235', 'E2EFDA', 'FCE4D6', 'FFFFCC'
REPORT_TITLE = 'Open Orders - Order Intake - Inventory Analysis By MSA'
OI_ALIGN, SUM_ALIGN, INV_ALIGN = 'LLRRLCRCLLL', 'LRRRR', 'LLLLRLRLR'
# widths as Excel's UI shows them; the file stores about 0.71 more (openpyxl and Excel alike)
OI_WIDTHS = [20.29, 18, 18.86, 22.57, 19.29, 21, 35.57, 22.29, 17.86, 24, 56.86]
INV_WIDTHS = [18, 25.14, 14, 16, 17, 16, 11.29, 56.86, 36]
WIDTH_PAD, WIDTH_TOL = 0.71, 0.3
CAP_CLASSES, CAP_STATUS, CAP_FILTERS, CAP_BUILD, CAP_MONTHS = (
    'Table 1. Size classes, in report order.', 'Table 2. Order status rank inside a size class.',
    'Table 3. Default filters.', 'Table 4. This build.', 'Table 5. Months in this build.')
FF_CAPTIONS = [CAP_CLASSES, CAP_STATUS, CAP_FILTERS, CAP_BUILD]
FILE_ITEMS = ['Order Intake file', 'Inventory by Month file', 'Open Order file']
OPTIONS_RE = re.compile(r'Options: InventoryMonth=(latest in file|[A-Z][a-z]+ \d{4}), ShowCancelled=(TRUE|FALSE), '
                        r'ShowSSL=(TRUE|FALSE), MSAColumn=(.+), SizeClassFile=(.+)')
MONTH_RE = re.compile(r'(%s) (\d{4})' % '|'.join(MONTHS))


class Options:
    def __init__(self, month, show_cancelled, show_ssl, city_column, size_class_file):
        self.month, self.show_cancelled, self.show_ssl = month, show_cancelled, show_ssl
        self.city_column, self.size_class_file = city_column, size_class_file


DEFAULT_OPTIONS = Options(None, False, False, 'auto', '')


# ------------------------------------------------------------------- value helpers

def text(value):
    return '' if value is None else str(value).strip()


def is_empty(value):
    return text(value) == ''


def is_date(value):
    return isinstance(value, (datetime.datetime, datetime.date))


def day_of(value):
    return value.toordinal() if is_date(value) else None


def is_number(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def shown(value):
    if value is None:
        return 'empty'
    if isinstance(value, datetime.datetime):
        return value.strftime('%Y-%m-%d')
    return repr(value) if isinstance(value, str) else str(value)


def same_value(a, b):
    if is_number(a) and is_number(b):
        return abs(a - b) < 1e-9
    return norm(a) == norm(b)


def norm(value):
    return None if value == '' else value


def month_text(value):
    return f'{MONTHS[value.month - 1]} {value.year}' if is_date(value) else ''


def month_key(textv):
    m = MONTH_RE.fullmatch(text(textv))
    return int(m.group(2)) * 12 + MONTHS.index(m.group(1)) if m else None


def class_table(sh):
    """File Format Table 1: (caption row, [(class, family abbreviation)] in rank order, problems). The class rows run
    from two rows under the caption to the UNMAPPED row."""
    ws = sh.ws
    r = next((r for r in range(1, ws.max_row + 1) if text(sh.raw(r, 1)) == CAP_CLASSES), None)
    if r is None:
        return None, [], [f'no "{CAP_CLASSES}" table']
    out = []
    for row in range(r + 2, r + 3 + MAX_CLASSES):
        name, fam = text(sh.raw(row, 1)), FAMILY_BY_TEXT.get(text(sh.raw(row, 2)))
        if name == UNMAPPED:
            return r, out, []
        if not name or fam is None:
            return r, out, [f'A{row}:B{row} {shown(sh.raw(row, 1))} / {shown(sh.raw(row, 2))} is not a class row, and no '
                            f'UNMAPPED row comes before it']
        out.append((name, fam[0]))
    return r, out, [f'no UNMAPPED row within {MAX_CLASSES} classes']


def class_ranks(ff):
    """{class: rank} from File Format Table 1 (empty when the sheet or the table is missing)."""
    _, classes, _ = class_table(ff) if ff else (None, [], [])
    return {name: i for i, (name, _) in enumerate(classes, start=1)}


def status_rank(value):
    return STATUS_RANK.get(text(value).upper(), 5)


def num_text(value):
    return str(int(value)) if is_number(value) and float(value).is_integer() else str(value)


def parse_options(value):
    m = OPTIONS_RE.fullmatch(text(value))
    if not m:
        return None
    month, canc, ssl, city, scf = m.groups()
    if month != 'latest in file' and month_key(month) is None:
        return None
    return Options(None if month == 'latest in file' else month, canc == 'TRUE', ssl == 'TRUE', city, scf)


# ------------------------------------------------------------------- style helpers

STRENGTH = {None: 0, 'hair': 1, 'thin': 2, 'medium': 3, 'double': 4, 'thick': 5}


def _style(cell, side):
    s = getattr(cell.border, side) if cell.border is not None else None
    return s.style if s is not None and s.style else None


def hline(ws, row, col):
    """Style of the rule above row in column col (top of row, bottom of row-1; the stronger when both are set)."""
    a = _style(ws.cell(row=row, column=col), 'top')
    b = _style(ws.cell(row=row - 1, column=col), 'bottom') if row > 1 else None
    return a if STRENGTH.get(a, 0) >= STRENGTH.get(b, 0) else b


def vline(ws, row, col):
    """Style of the rule left of column col in row (left of col, right of col-1)."""
    a = _style(ws.cell(row=row, column=col), 'left')
    b = _style(ws.cell(row=row, column=col - 1), 'right') if col > 1 else None
    return a if STRENGTH.get(a, 0) >= STRENGTH.get(b, 0) else b


def grid_problems(ws, head, last, ncols, lines, col=1):
    """Horizontal rules over rows head..last: lines maps a row to the rule wanted above it (thin otherwise, and
    below the last row); no vertical rules, the table's outer edges included."""
    probs = []
    for r in range(head, last + 2):
        want = lines.get(r, 'thin')
        for c in range(col, col + ncols):
            got = hline(ws, r, c)
            if got != want:
                probs.append(f'{get_column_letter(c)}{r} rule above {got}, want {want}')
                return probs
    for r in range(head, last + 1):
        for c in range(col, col + ncols + 1):
            want = None
            got = vline(ws, r, c)
            if got != want:
                probs.append(f'{get_column_letter(c)}{r} rule left {got}, want {want}')
                return probs
    return probs


def align_problems(ws, row, spec, col=1):
    for i, a in enumerate(spec):
        al = ws.cell(row=row, column=col + i).alignment
        want = {'L': 'left', 'R': 'right', 'C': 'center'}[a]
        if al.horizontal != want or (a != 'C' and not al.indent) or al.vertical != 'center':
            return [f'{get_column_letter(col + i)}{row} aligned {al.horizontal}/{al.vertical} indent {al.indent}, want {want}/center'
                    + ('' if a == 'C' else ' indent 1')]
    return []


def fill_hex(fill):
    """'RRGGBB' of a solid fill (fgColor, else bgColor as conditional formats store it), None for no fill."""
    if fill is None or fill.fill_type in (None, 'none'):
        return None
    for colour in (fill.fgColor, fill.bgColor):
        if colour is not None and colour.type == 'rgb' and str(colour.rgb)[-6:] not in ('000000',):
            return str(colour.rgb)[-6:].upper()
    return '000000'


def colour_problems(ws, font_name, heads, fills):
    """No theme colour (the default font colour aside); fills exactly where fills says ({(row, col): colour}),
    the header cells heads in it with white text; black borders, black text elsewhere; the sheet's font."""
    bad = []
    for (r, c), want in sorted(fills.items()):
        if fill_hex(ws.cell(row=r, column=c).fill) != want:
            bad.append(f'{get_column_letter(c)}{r} not filled {want}')
            break
    for r, c in sorted(heads):
        fc = ws.cell(row=r, column=c).font.color
        if not (fc is not None and fc.type == 'rgb' and str(fc.rgb)[-6:].upper() == 'FFFFFF'):
            bad.append(f'{get_column_letter(c)}{r} header text not white')
            break
    for row in ws.iter_rows():
        for c in row:
            f = fill_hex(c.fill)
            if f is not None and fills.get((c.row, c.column)) != f:
                bad.append(f'{c.coordinate} has a fill {f}')
            fc = c.font.color if c.font is not None else None
            if (c.row, c.column) not in heads and fc is not None and not (fc.type == 'theme' and fc.theme == 1 and not fc.tint) and \
                    not (fc.type == 'rgb' and str(fc.rgb)[-6:] == '000000') and fc.type != 'auto':
                bad.append(f'{c.coordinate} font colour {fc.rgb if fc.type == "rgb" else fc.type}')
            if c.value is not None and (c.font is None or c.font.name != font_name):
                bad.append(f'{c.coordinate} font {c.font.name if c.font else None}, want {font_name}')
            b = c.border
            if b is not None:
                for side in (b.left, b.right, b.top, b.bottom):
                    if side is not None and side.style and side.color is not None and \
                            not (side.color.type == 'rgb' and str(side.color.rgb)[-6:] == '000000') and side.color.type != 'auto' \
                            and not (side.color.type == 'indexed' and side.color.indexed in (8, 64)):
                        bad.append(f'{c.coordinate} border colour')
                        break
            if len(bad) >= 5:
                return bad
    return bad


def cf_rule(ws, ref, formula, colour):
    """The rule over exactly ref with that formula and fill, or None."""
    for cf in ws.conditional_formatting:
        if str(cf.sqref) == ref:
            for rule in cf.rules:
                if rule.type == 'expression' and [x.lstrip('=') for x in rule.formula] == [formula] and \
                        rule.dxf is not None and fill_hex(rule.dxf.fill) == colour:
                    return rule
    return None


def band_problems(ws, first, last, col, ncols):
    """The banding rule on rows first..last: one conditional format over exactly that range, formula
    MOD(SUBTOTAL(103,$A$first:$Afirst),2)=0 with A the first column, filling EDEDED."""
    if last < first:
        return []
    a, z = get_column_letter(col), get_column_letter(col + ncols - 1)
    ref, formula = f'{a}{first}:{z}{last}', f'MOD(SUBTOTAL(103,${a}${first}:${a}{first}),2)=0'
    for cf in ws.conditional_formatting:
        if str(cf.sqref) != ref:
            continue
        for rule in cf.rules:
            if rule.type == 'expression' and [x.lstrip('=') for x in rule.formula] == [formula] and \
                    rule.dxf is not None and fill_hex(rule.dxf.fill) == BAND_FILL:
                return []
        return [f'{ref} banding rule {[(r.type, r.formula) for r in cf.rules]}, want {formula} filling {BAND_FILL}']
    return [f'{ref} has no banding rule ({formula})']


def fill_cells(ws, kind, inv, heads):
    """{(row, col): colour} of every static fill the sheet should have."""
    out = {rc: HEAD_FILL for rc in heads}
    if kind == 'inv' and inv and inv.h:
        out[(3, 2)] = INPUT_FILL
    if kind == 'ff':
        r = next((r for r in range(1, ws.max_row + 1) if str(ws.cell(row=r, column=1).value or '') == CAP_CLASSES), None)
        u = None if r is None else next((x for x in range(r + 2, r + 3 + MAX_CLASSES)
                                         if text(ws.cell(row=x, column=1).value) == UNMAPPED), None)
        if u is not None:
            out[(u, 1)] = FLAG_FILL
    return out


def header_cells(ws, kind, inv):
    if kind == 'oi':
        return {(2, c) for c in range(1, 12)}
    if kind == 'inv':
        return {(4, c) for c in range(1, 6)} | ({(inv.h, c) for c in range(1, len(INV_HEADERS) + 1)} if inv and inv.h else set())
    if kind == 'oo':
        return {(4, c) for c in range(1, 6)}
    if kind == 'ff':
        widths_ = {CAP_CLASSES: 3, CAP_STATUS: 2, CAP_FILTERS: 4, CAP_BUILD: 4}
        return {(r + 1, c) for r in range(1, ws.max_row + 1) if str(ws.cell(row=r, column=1).value or '') in widths_
                for c in range(1, widths_[str(ws.cell(row=r, column=1).value)] + 1)}
    return set()


def print_problems(ws, title_rows, pages_tall):
    probs = []
    ps = ws.page_setup
    if ps.orientation != 'landscape':
        probs.append(f'orientation {ps.orientation}')
    fit = ws.sheet_properties.pageSetUpPr
    if not (fit is not None and fit.fitToPage) or int(ps.fitToWidth if ps.fitToWidth is not None else 1) != 1 or \
            int(ps.fitToHeight if ps.fitToHeight is not None else 1) != pages_tall:
        probs.append(f'not fitted 1 page wide x {pages_tall or "any"} tall (fitToWidth {ps.fitToWidth}, fitToHeight {ps.fitToHeight})')
    want = {'header left': REPORT_TITLE, 'header right': None, 'footer left': '&A', 'footer center': 'Page &P of &N'}
    parts = {'header left': ws.oddHeader.left, 'header right': ws.oddHeader.right, 'footer left': ws.oddFooter.left,
             'footer center': ws.oddFooter.center}
    for k, part in parts.items():
        t = part.text or ''
        if (want[k] is None and not t.startswith('Built ')) or (want[k] is not None and t != want[k]):
            probs.append(f'page {k} {t!r}')
        elif (part.font or '').split(',')[0].strip('"') != FONT_NAME:
            probs.append(f'page {k} font {part.font!r}')
    got = (ws.print_title_rows or '').replace('$', '') or None
    if got != title_rows:
        probs.append(f'print title rows {got}, want {title_rows}')
    return probs


def widths(ws):
    out = {}
    for dim in ws.column_dimensions.values():
        if dim.width is None:
            continue
        low = dim.min or column_index_from_string(dim.index)
        for col in range(low, (dim.max or low) + 1):
            out[col] = dim.width
    return out


def filter_columns(ws):
    """{colId: ('values', [..]) or ('custom', [(operator, val), ...])} of the sheet's AutoFilter."""
    out = {}
    for fc in ws.auto_filter.filterColumn:
        if fc.filters is not None and fc.filters.filter:
            out[fc.colId] = ('values', sorted(text(v) for v in fc.filters.filter))
        elif fc.customFilters is not None:
            out[fc.colId] = ('custom', [(cf.operator or 'equal', text(cf.val)) for cf in fc.customFilters.customFilter])
        else:
            out[fc.colId] = ('other', [])
    return out


# ------------------------------------------------------------------- workbook access

class Sheet:
    def __init__(self, ws, cached):
        self.ws, self.cached_ws, self.name = ws, cached, ws.title

    def raw(self, row, col):
        return self.ws.cell(row=row, column=col).value

    def get(self, row, col):
        return self.raw(row, col)

    def hidden(self, row):
        dim = self.ws.row_dimensions.get(row)
        return bool(dim is not None and dim.hidden)

    def last_row(self, col=1, start=1):
        r = self.ws.max_row
        while r >= start and is_empty(self.raw(r, col)):
            r -= 1
        return r


class Book:
    def __init__(self, path):
        self.path = path
        self.formulas = load_workbook(path)
        self.values = load_workbook(path, data_only=True)
        self.names = list(self.formulas.sheetnames)
        self._sheets = {}

    def sheet(self, name):
        if name not in self.names:
            return None
        if name not in self._sheets:
            self._sheets[name] = Sheet(self.formulas[name], self.values[name])
            self._sheets[name].book = self
        return self._sheets[name]


class Tally:
    def __init__(self):
        self.checks = self.failed = 0

    def add(self, rule, sheet, problems, limit=5):
        self.checks += 1
        if not problems:
            print(f'PASS {rule} {sheet}')
            return True
        self.failed += 1
        extra = len(problems) - limit
        print(f'FAIL {rule} {sheet}: ' + '; '.join(problems[:limit]) + (f' (+{extra} more)' if extra > 0 else ''))
        return False


# ------------------------------------------------------------------- formula evaluation (for E3, C3, R3g and E8)

REF_RE = re.compile(r'\$?([A-Z]{1,3})\$?(\d+)')


def split_args(s, sep=','):
    """s split at sep outside quotes, brackets and quoted sheet names."""
    out, cur, depth, quoted, sheet_q = [], '', 0, False, False
    for ch in s:
        if ch == '"' and not sheet_q:
            quoted = not quoted
        elif ch == "'" and not quoted:
            sheet_q = not sheet_q
        elif not quoted and not sheet_q and ch == '(':
            depth += 1
        elif not quoted and not sheet_q and ch == ')':
            depth -= 1
        if ch == sep and depth == 0 and not quoted and not sheet_q:
            out.append(cur)
            cur = ''
        else:
            cur += ch
    return out + [cur]


SHEET_REF_RE = re.compile(r"(?:'((?:[^']|'')+)'|([A-Za-z0-9_.]+))!(.+)")


def cells_of(sh, ref):
    """(sheet, [(row, col), ...]) of a reference, on sh or on the sheet it names ('Order Intake'!$G$3:$G$31)."""
    ref = ref.strip()
    m = SHEET_REF_RE.fullmatch(ref)
    if m:
        name = (m.group(1) or m.group(2)).replace("''", "'")
        target = sh.book.sheet(name) if getattr(sh, 'book', None) else None
        if target is None:
            raise ValueError(f'reference to a sheet that is not there: {ref}')
        sh, ref = target, m.group(3)
    a, b = ref.split(':') if ':' in ref else (ref, ref)
    (c1, r1), (c2, r2) = [(column_index_from_string(m.group(1)), int(m.group(2))) for m in map(REF_RE.fullmatch, (a, b))]
    return sh, [(r, c) for r in range(min(r1, r2), max(r1, r2) + 1) for c in range(min(c1, c2), max(c1, c2) + 1)]


def criterion(sh, arg):
    arg = arg.strip()
    if arg.startswith('"') and arg.endswith('"'):
        return arg[1:-1].replace('""', '"')
    target, cells = cells_of(sh, arg)
    return evaluate(target, *cells[0])


def fold(value):
    """Text as SUMIFS compares it: the whole text, ignoring case, not ignoring stray spaces."""
    return ('' if value is None else str(value)).casefold()


def evaluate(sh, row, col, depth=0):
    """Value of a cell, evaluating the formula forms the report writes (SUMIFS with text or cell criteria, SUM,
    SUBTOTAL(109, ...) over the rows not hidden, and a + of such terms), across sheets."""
    v = sh.raw(row, col)
    if not (isinstance(v, str) and v.startswith('=')) or depth > 5:
        return v
    total = 0
    for term in split_args(v.strip()[1:], '+'):
        m = re.fullmatch(r'(SUMIFS|SUM|SUBTOTAL)\((.*)\)', term.strip(), re.I)
        if not m:
            return None
        args = split_args(m.group(2))
        if m.group(1).upper() == 'SUBTOTAL':           # 109: the sum of what the filters show (hidden rows skipped)
            if len(args) != 2 or args[0].strip() != '109':
                return None
            target, cells = cells_of(sh, args[1])
            total += sum(x for x in (evaluate(target, r, c, depth + 1) for r, c in cells if not target.hidden(r))
                         if is_number(x))
            continue
        if m.group(1).upper() == 'SUM':
            for a in args:
                target, cells = cells_of(sh, a)
                total += sum(x for x in (evaluate(target, r, c, depth + 1) for r, c in cells) if is_number(x))
            continue
        sum_sheet, sum_cells = cells_of(sh, args[0])
        pairs = [(cells_of(sh, args[i]), fold(criterion(sh, args[i + 1]))) for i in range(1, len(args) - 1, 2)]
        for k, (r, c) in enumerate(sum_cells):
            if all(fold(rs.raw(*cells[k])) == crit for (rs, cells), crit in pairs):
                x = sum_sheet.raw(r, c)
                if is_number(x):
                    total += x
    return total


# ------------------------------------------------------------------- the report's parts

class OI:
    """Order Intake rows 3.. (A..K) with hidden flags."""
    def __init__(self, sh):
        self.sh = sh
        self.last = max(sh.last_row(1, 3), 2)
        self.rows = [[sh.raw(r, c) for c in range(1, 12)] for r in range(3, self.last + 1)]
        self.hidden = [sh.hidden(r) for r in range(3, self.last + 1)]


class INV:
    """Inventory by Month: month cell, summary rows, detail rows."""
    def __init__(self, sh):
        self.sh = sh
        self.month = text(sh.raw(3, 2))
        self.h = next((r for r in range(5, sh.ws.max_row + 1)
                       if sh.raw(r, 1) == 'MSA' and sh.raw(r, 2) == 'PRODUCT_TYPE'), None)
        if self.h is None:
            self.cities, self.detail, self.hidden, self.last = [], [], [], 0
            return
        self.n = self.h - 7
        self.cities = [sh.raw(r, 1) for r in range(5, 5 + self.n)]
        self.last = max(sh.last_row(1, self.h + 1), self.h)
        self.detail = [[sh.raw(r, c) for c in range(1, 1 + len(INV_HEADERS))] for r in range(self.h + 1, self.last + 1)]
        self.hidden = [sh.hidden(r) for r in range(self.h + 1, self.last + 1)]


def all_cities(oi, inv):
    names = {}
    for r in oi.rows:
        names.setdefault(text(r[1]).upper(), text(r[1]))
    for r in inv.detail:
        names.setdefault(text(r[0]).upper(), text(r[0]))
    return sorted(names.values(), key=str.upper)


def raw_text(value):
    return '' if value is None else str(value)


class OpenTally:
    """This checker's own count of the open orders on the Order Intake tab: rows whose ORDER_STATUS is ORDERED or
    Not yet processed as SUMIFS compares it (whole text, ignoring case, not ignoring stray spaces), by MSA."""
    def __init__(self, oi, cities):
        self.cities = cities
        self.rows = {c.upper(): 0 for c in cities}
        self.ordered = {c.upper(): 0 for c in cities}
        self.units = {c.upper(): {a: 0 for a in ABBRS} for c in cities}
        self.spaced, self.spaced_texts, self.disagree = 0, [], 0
        for row in oi.rows:
            status, key, abbr = raw_text(row[4]), text(row[1]).upper(), text(row[7])
            open_ = status.upper() in OPEN_STATUSES
            if open_ and key in self.rows:
                self.rows[key] += 1
                self.ordered[key] += status.upper() == 'ORDERED'
                if abbr in ABBRS and is_number(row[6]):
                    self.units[key][abbr] += row[6]
            elif not open_ and status.strip(' ').upper() in OPEN_STATUSES:
                self.spaced += 1
                if status not in self.spaced_texts:
                    self.spaced_texts.append(status)
            if isinstance(row[5], bool) and row[5] != open_:
                self.disagree += 1

    def table(self):
        out = [[c] + [self.units[c.upper()][a] for a in ABBRS] + [sum(self.units[c.upper()].values())] for c in self.cities]
        return out + [['Total'] + [sum(r[k] for r in out) for k in range(1, 5)]]

    def total_rows(self):
        return sum(self.rows.values())

    def build_text(self):
        t = self.table()[-1]
        n = self.total_rows()
        return (f'ORDERED and Not yet processed: {n} row{"" if n == 1 else "s"}, {num_text(t[4])} units (CTL {num_text(t[1])}, '
                f'MX {num_text(t[2])}, SSL {num_text(t[3])}), on the Open Orders by MSA tab')

    def city_line(self, c):
        k = c.upper()
        n = self.rows[k]
        if not n:
            return f'{c}: no open orders (ORDERED or Not yet processed)'
        u = self.units[k]
        return (f'{c}: Open orders {n} row{"" if n == 1 else "s"} (ORDERED {self.ordered[k]}, Not yet processed '
                f'{n - self.ordered[k]}): CTL {num_text(u["CTL"])}, MX {num_text(u["MX"])}, SSL {num_text(u["SSL"])} units')

    def tail_lines(self):
        """The open orders lines after the inventory "shows zeros" block, in their order."""
        out = []
        if self.total_rows() == 0:
            out.append('Open orders: no rows with ORDER_STATUS ORDERED or Not yet processed - the Open Orders by MSA tab '
                       'shows zeros')
        if self.spaced:
            listed = ', '.join(f'"{t}"' for t in self.spaced_texts[:10]) + (', ...' if len(self.spaced_texts) > 10 else '')
            out.append(f'WARNING: {self.spaced} row{"" if self.spaced == 1 else "s"} have ORDER_STATUS ORDERED or Not yet '
                       f'processed with extra spaces and are not counted on the Open Orders by MSA tab: {listed}')
        if self.disagree:
            out.append(f'Open orders: {self.disagree} row{"" if self.disagree == 1 else "s"} where IS_OPEN_ORDER disagrees '
                       f'with the status (TRUE with another status, or FALSE with ORDERED / Not yet processed)')
        return out


class OO:
    """Open Orders by MSA: the MSA rows 5.. down to the Total row."""
    def __init__(self, sh):
        self.sh = sh
        self.tr = next((r for r in range(5, sh.ws.max_row + 1) if text(sh.raw(r, 1)) == 'Total'), None)
        self.n = (self.tr - 5) if self.tr is not None else 0
        self.cities = [sh.raw(r, 1) for r in range(5, 5 + self.n)]

    def values(self):
        """[[MSA, CTL, MX, SSL, Total], ..., ['Total', ...]] with the formulas evaluated."""
        if self.tr is None:
            return []
        return [[text(self.sh.raw(r, 1))] + [evaluate(self.sh, r, c) for c in range(2, 6)] for r in range(5, self.tr + 1)]


# ------------------------------------------------------------------- charts, read from the .xlsx zip

NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'pr': 'http://schemas.openxmlformats.org/package/2006/relationships',
      'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'}


def _rels(z, part):
    """{Id: (type, part path)} of the relationships of a part in the zip."""
    folder, name = posixpath.split(part)
    rels = posixpath.join(folder, '_rels', name + '.rels')
    if rels not in z.namelist():
        return {}
    out = {}
    for rel in ET.fromstring(z.read(rels)).findall('pr:Relationship', NS):
        if rel.get('TargetMode') == 'External':
            continue
        t = rel.get('Target')
        out[rel.get('Id')] = (rel.get('Type') or '', t.lstrip('/') if t.startswith('/') else posixpath.normpath(posixpath.join(folder, t)))
    return out


def _chart(xml):
    root = ET.fromstring(xml)
    chart = root.find('c:chart', NS)
    plot = chart.find('c:plotArea', NS) if chart is not None else None
    out = {'kinds': [], 'barDir': None, 'grouping': None, 'series': [], 'title': None}
    if plot is None:
        return out
    out['kinds'] = [e.tag.split('}')[1] for e in plot if e.tag.endswith('Chart')]
    title = chart.find('c:title', NS)
    if title is not None:
        out['title'] = ''.join(t.text or '' for t in title.iter('{%s}t' % NS['a']))
    bar = plot.find('c:barChart', NS)
    if bar is None:
        return out
    val = lambda e: e.get('val') if e is not None else None
    out['barDir'], out['grouping'] = val(bar.find('c:barDir', NS)), val(bar.find('c:grouping', NS))
    for ser in bar.findall('c:ser', NS):
        def ref(*paths):
            for p in paths:
                e = ser.find(p, NS)
                if e is not None:
                    return e.text
            return None
        out['series'].append((ref('c:tx/c:strRef/c:f'), ref('c:cat/c:strRef/c:f', 'c:cat/c:numRef/c:f', 'c:cat/c:multiLvlStrRef/c:f'),
                              ref('c:val/c:numRef/c:f')))
    return out


def chart_facts(path, sheet_name):
    """(drawings, charts) of a sheet read from the zip: the number of drawing parts and, per chart, its type,
    series refs, title text and the 0-based row of its top-left anchor. None when the sheet is not in the zip."""
    with zipfile.ZipFile(path) as z:
        wb = ET.fromstring(z.read('xl/workbook.xml'))
        rid = next((s.get('{%s}id' % NS['r']) for s in wb.find('m:sheets', NS) if s.get('name') == sheet_name), None)
        rels = _rels(z, 'xl/workbook.xml')
        if rid is None or rid not in rels:
            return None
        drawings = [p for t, p in _rels(z, rels[rid][1]).values() if t.endswith('/drawing')]
        charts = []
        for dp in drawings:
            drels = _rels(z, dp)
            for anchor in ET.fromstring(z.read(dp)):
                frm = anchor.find('xdr:from', NS)
                row = int(frm.find('xdr:row', NS).text) if frm is not None else None
                for el in anchor.iter('{%s}chart' % NS['c']):
                    rel = drels.get(el.get('{%s}id' % NS['r']))
                    if rel is not None:
                        charts.append(dict(_chart(z.read(rel[1])), anchor_row=row))
        return len(drawings), charts


def plain_ref(ref):
    """A chart reference without $ signs and sheet quotes, so Excel's and openpyxl's spellings compare equal."""
    return re.sub(r"[$']", '', ref or '')


def ref_texts(sh, ref):
    """The texts of the cells a chart reference names (on sh)."""
    m = re.fullmatch(r'.*!([A-Z]+\d+(?::[A-Z]+\d+)?)', plain_ref(ref))
    if not m:
        return None
    _, cells = cells_of(sh, m.group(1))
    return [text(sh.raw(r, c)) for r, c in cells]


def chart_summary(facts, sh):
    """(series count, categories, anchor row) of the one chart, or None."""
    if not facts or len(facts[1]) != 1:
        return None
    ch = facts[1][0]
    cats = ref_texts(sh, ch['series'][0][1]) if ch['series'] else []
    return len(ch['series']), cats, ch['anchor_row']


# ------------------------------------------------------------------- R2 style

def check_style(sh, kind, oi=None, inv=None, oo=None):
    ws, probs = sh.ws, []
    if ws.sheet_view.showGridLines is not False:
        probs.append('gridlines are on')
    heads = header_cells(ws, kind, inv)
    probs += colour_problems(ws, LOG_FONT_NAME if kind == 'log' else FONT_NAME, heads, fill_cells(ws, kind, inv, heads))
    ncols = {'oi': 11, 'inv': len(INV_HEADERS), 'ff': 6, 'oo': 5, 'log': 0}[kind]
    if kind in ('oi', 'inv', 'ff', 'oo'):
        a1 = ws['A1']
        if not (a1.font.b and abs((a1.font.sz or 0) - 14) < 0.01):
            probs.append('A1 is not the title style (14 bold)')
        note = ws['H1' if kind == 'oi' else 'D1']
        if not note.font.i:
            probs.append(f'{note.coordinate} is not the note style (italic)')
        for c in range(1, ncols + 1):
            if hline(ws, 2, c) != 'medium':
                probs.append(f'no medium rule under {get_column_letter(c)}1')
                break
    heads = {'oi': [(2, OI_ALIGN)], 'inv': [(4, SUM_ALIGN)] + ([(inv.h, INV_ALIGN)] if inv and inv.h else []), 'ff': [],
             'oo': [(4, SUM_ALIGN)], 'log': []}[kind]
    for r, spec in heads:
        for c in range(1, len(spec) + 1):
            if not ws.cell(row=r, column=c).font.b:
                probs.append(f'{get_column_letter(c)}{r} header not bold')
                break
        probs += align_problems(ws, r, spec)
        if r + 1 <= ws.max_row:
            probs += align_problems(ws, r + 1, spec)
    want_freeze = {'oi': 'A3', 'inv': f'A{inv.h + 1}' if inv and inv.h else None, 'ff': None, 'oo': None, 'log': None}[kind]
    if (ws.freeze_panes or None) != want_freeze:
        probs.append(f'freeze panes {ws.freeze_panes}, want {want_freeze}')
    if kind in ('oi', 'inv', 'oo'):
        got = widths(ws)
        for col, w in enumerate({'oi': OI_WIDTHS, 'inv': INV_WIDTHS, 'oo': OO_WIDTHS}[kind], start=1):
            if abs(got.get(col, 0) - (w + WIDTH_PAD)) > WIDTH_TOL:
                probs.append(f'column {get_column_letter(col)} width {got.get(col)}, want {w + WIDTH_PAD:.2f}')
    if kind == 'oi':
        for i, row in enumerate(oi.rows):
            r = 3 + i
            if text(row[0]) == UNMAPPED and not ws.cell(row=r, column=1).font.b:
                probs.append(f'A{r} UNMAPPED not bold')
            if ws.cell(row=r, column=7).number_format != '0' and is_number(row[6]):
                probs.append(f'G{r} number format {ws.cell(row=r, column=7).number_format}, want 0')
        probs += band_problems(ws, 3, 2 + len(oi.rows), 1, 11)
        n = len(oi.rows)
        if n:
            flag = cf_rule(ws, f'A3:A{2 + n}', '$A3="UNMAPPED"', FLAG_FILL)
            bandr = cf_rule(ws, f'A3:K{2 + n}', 'MOD(SUBTOTAL(103,$A$3:$A3),2)=0', BAND_FILL)
            if flag is None or not flag.stopIfTrue or (bandr is not None and not flag.priority < bandr.priority):
                probs.append(f'A3:A{2 + n} has no first, stopping rule $A3="UNMAPPED" filling {FLAG_FILL}')
    if kind == 'inv' and inv.h:
        if any(vline(ws, 3, c) != 'thin' for c in (2, 3)) or hline(ws, 3, 2) != 'thin' or hline(ws, 4, 2) not in ('thin', 'medium'):
            probs.append('B3 (month cell) is not boxed by thin rules')
        probs += grid_problems(ws, 4, 5 + inv.n, 5, {4: 'medium', 5: 'double', 5 + inv.n: 'double'})
        probs += band_problems(ws, 5, 4 + inv.n, 1, 5)
        probs += band_problems(ws, inv.h + 1, inv.h + len(inv.detail), 1, len(INV_HEADERS))
        for r in range(5, inv.h - 1):
            for c in range(2, 6):
                if ws.cell(row=r, column=c).number_format != '#,##0':
                    probs.append(f'{get_column_letter(c)}{r} number format {ws.cell(row=r, column=c).number_format}, want #,##0')
        for r in range(inv.h + 1, inv.last + 1):
            if ws.cell(row=r, column=7).number_format != '#,##0' and is_number(sh.raw(r, 7)):
                probs.append(f'G{r} number format, want #,##0')
                break
    if kind == 'oo' and oo is not None and oo.tr is not None:
        probs += grid_problems(ws, 4, oo.tr, 5, {4: 'medium', 5: 'double', oo.tr: 'double'})
        probs += band_problems(ws, 5, 4 + oo.n, 1, 5)
        for r in range(5, oo.tr + 1):
            for c in range(2, 6):
                if ws.cell(row=r, column=c).number_format != '#,##0':
                    probs.append(f'{get_column_letter(c)}{r} number format {ws.cell(row=r, column=c).number_format}, want #,##0')
    title_rows = {'oi': '1:2', 'inv': f'{inv.h}:{inv.h}' if inv and inv.h else None, 'ff': None, 'oo': None, 'log': None}[kind]
    probs += print_problems(ws, title_rows, 1 if kind in ('ff', 'oo') else 0)
    return probs


# ------------------------------------------------------------------- R2F File Format

def check_file_format(sh, opts, tally=None):
    ws, probs = sh.ws, []
    if text(sh.raw(1, 1)) != 'File Format':
        probs.append(f'A1={shown(sh.raw(1, 1))}')
    rows = {text(sh.raw(r, 1)): r for r in range(1, ws.max_row + 1) if text(sh.raw(r, 1)) in FF_CAPTIONS}
    for cap in FF_CAPTIONS:
        if cap not in rows:
            probs.append(f'no "{cap}" table')
    if probs:
        return probs
    r, classes, table_probs = class_table(sh)
    n = len(classes)
    probs += table_probs
    if [text(sh.raw(r + 1, c)) for c in (1, 2, 3)] != ['Class', 'Family', 'Rule']:
        probs.append('size class table headers')
    if not table_probs:
        u = r + 2 + n                                        # the UNMAPPED row
        if n == 0:
            probs.append('size class table has no classes')
        order = [ab for ab, _, _ in FAMILIES]
        lines = {r + 1: 'medium', r + 2: 'double'}
        for i, (name, fam) in enumerate(classes):
            row, rule = r + 2 + i, next(f[2] for f in FAMILIES if f[0] == fam)
            if text(sh.raw(row, 3)) != rule:
                probs.append(f'C{row} {shown(sh.raw(row, 3))}, want {rule!r}')
            if i and order.index(fam) < order.index(classes[i - 1][1]):
                probs.append(f'B{row} {fam} after {classes[i - 1][1]}: families out of the order CTL, MX, SSL')
            if i and fam != classes[i - 1][1]:
                lines[row] = 'medium'                        # a medium rule above each family after the first
            if name.upper() in [c.upper() for c, _ in classes[:i]]:
                probs.append(f'A{row} {name!r} listed twice')
        if [text(sh.raw(u, c)) for c in (1, 2, 3)] != UNMAPPED_ROW:
            probs.append(f'A{u}:C{u} {[text(sh.raw(u, c)) for c in (1, 2, 3)]}, want {UNMAPPED_ROW}')
        if not ws.cell(row=u, column=1).font.b:
            probs.append(f'A{u} UNMAPPED not bold')
        probs += grid_problems(ws, r + 1, u, 3, lines) + band_problems(ws, r + 2, u - 1, 1, 3)
        # v2.8: the tables under Table 1 move with it: Table 2 two rows under its UNMAPPED row, then 9 rows apart
        want_rows = {CAP_STATUS: u + 2, CAP_FILTERS: u + 11, CAP_BUILD: u + 20}
        for cap, want in want_rows.items():
            if rows[cap] != want:
                probs.append(f'"{cap}" in A{rows[cap]}, want A{want} (Table 1 has {n} classes)')
    r = rows[CAP_STATUS]
    if [text(sh.raw(r + 2 + i, 2)) for i in range(6)] != STATUS_ORDER or \
            [sh.raw(r + 2 + i, 1) for i in range(6)] != [1, 2, 3, 4, 5, 6]:
        probs.append('status order table differs from ' + ', '.join(STATUS_ORDER))
    probs += grid_problems(ws, r + 1, r + 7, 2, {r + 1: 'medium', r + 2: 'double'}) + band_problems(ws, r + 2, r + 7, 1, 2)
    r = rows[CAP_FILTERS]
    filt = [[text(sh.raw(r + 2 + i, c)) for c in range(1, 5)] for i in range(5)]
    ssl = f'ShowSSL = {"TRUE" if opts.show_ssl else "FALSE"}'
    want = [['Order Intake', 'ORDER_STATUS', 'all' if opts.show_cancelled else 'all except CANCELLED'],
            ['Order Intake', 'PRODUCT_TYPE_ABBR', 'all (CTL, MX, SSL)' if opts.show_ssl else 'CTL, MX'],
            ['Inventory by Month', 'Month'], ['Inventory by Month', 'PRODUCT_TYPE'],
            [OO_SHEET, 'PRODUCT_TYPE_ABBR', 'CTL, MX, SSL on the chart' if opts.show_ssl else 'CTL, MX on the chart', ssl]]
    for got, w in zip(filt, want):
        if got[:len(w)] != w:
            probs.append(f'filter row {got}, want {w}')
    probs += grid_problems(ws, r + 1, r + 6, 4, {r + 1: 'medium', r + 2: 'double'}) + band_problems(ws, r + 2, r + 6, 1, 4)
    if text(sh.raw(r + 7, 1)) != 'Note. Clear a filter to see the rows it hides.' or not ws.cell(row=r + 7, column=1).font.i:
        probs.append('the italic "Note. Clear a filter" line is missing')
    r = rows[CAP_BUILD]
    items = [text(sh.raw(i, 1)) for i in range(r + 2, ws.max_row + 1) if not is_empty(sh.raw(i, 1))]
    probs += grid_problems(ws, r + 1, r + 1 + len(items), 4, {r + 1: 'medium', r + 2: 'double'})
    probs += band_problems(ws, r + 2, r + 1 + len(items), 1, 4)
    for item in ['Built', 'Builder', 'Input folder', 'Order Intake file', 'Size class file', 'Inventory month', 'MSAs',
                 'Open orders']:
        if item not in items:
            probs.append(f'build table has no {item} row')
    if 'MSAs' in items and 'Open orders' in items and items.index('Open orders') != items.index('MSAs') + 1:
        probs.append('build table: Open orders is not the item after MSAs')
    got = build_item(sh, 'Open orders')
    if tally is not None and got is not None and got != tally.build_text():
        probs.append(f'build table Open orders {got!r}, own tally {tally.build_text()!r}')
    if text(sh.raw(3, 6)) != CAP_MONTHS:
        probs.append(f'F3 is not "{CAP_MONTHS}"')
    months, r = [], 4
    while not is_empty(sh.raw(r, 6)):
        months.append(text(sh.raw(r, 6)))
        r += 1
    keys = [month_key(m) for m in months]
    if None in keys or keys != sorted(set(keys)):
        probs.append(f'months list not chronological month texts: {months}')
    if months:
        probs += grid_problems(ws, 4, 3 + len(months), 1, {}, col=6) + band_problems(ws, 4, 3 + len(months), 6, 1)
    return probs


def months_list(sh):
    out, r = [], 4
    while sh is not None and not is_empty(sh.raw(r, 6)):
        out.append(text(sh.raw(r, 6)))
        r += 1
    return out


def build_files(sh):
    if sh is None:
        return []
    r0 = next((r for r in range(1, sh.ws.max_row + 1) if text(sh.raw(r, 1)) == CAP_BUILD), None)
    if r0 is None:
        return []
    return [[text(sh.raw(r, 1)), text(sh.raw(r, 2))] for r in range(r0 + 2, sh.ws.max_row + 1)
            if text(sh.raw(r, 1)) in FILE_ITEMS]


def build_item(sh, item):
    """The value of one item of the File Format build table, or None."""
    if sh is None:
        return None
    r0 = next((r for r in range(1, sh.ws.max_row + 1) if text(sh.raw(r, 1)) == CAP_BUILD), None)
    if r0 is None:
        return None
    return next((text(sh.raw(r, 2)) for r in range(r0 + 2, sh.ws.max_row + 1) if text(sh.raw(r, 1)) == item), None)


# ------------------------------------------------------------------- R3 Order Intake

def check_order_intake(sh, oi, opts, cities, ranks):
    found = {k: [] for k in ('R3a', 'R3b', 'R3c', 'R3d', 'R3e', 'R3f')}
    heads = [text(sh.raw(2, c)) for c in range(1, 12)]
    if heads != OI_HEADERS:
        found['R3a'].append(f'headers {heads}')
    crank = {c.upper(): i for i, c in enumerate(cities)}
    prev = None
    for i, row in enumerate(oi.rows):
        r = 3 + i
        cls = text(row[0])
        if cls not in ranks and cls != UNMAPPED:
            found['R3a'].append(f'A{r}={shown(row[0])}')
        key = (ranks.get(cls, UNKNOWN_RANK), status_rank(row[4]), day_of(row[2]), crank.get(text(row[1]).upper(), -1))
        if prev is not None:
            if key[0] < prev[0]:
                found['R3b'].append(f'A{r} {cls} after a later class')
            elif key[0] == prev[0]:
                if key[1] < prev[1]:
                    found['R3c'].append(f'E{r} status {shown(row[4])} after a later status')
                elif key[1] == prev[1]:
                    if key[2] is not None and prev[2] is None:
                        found['R3c'].append(f'C{r} date after a non-date')
                    elif key[2] is not None and prev[2] is not None and key[2] > prev[2]:
                        found['R3c'].append(f'C{r} newer than the row above')
                    elif key[2] == prev[2] and key[3] < prev[3]:
                        found['R3c'].append(f'B{r} city out of order')
        prev = key
        if text(row[7]) not in ('CTL', 'MX', 'SSL') or row[7] != text(row[7]):
            found['R3d'].append(f'H{r}={shown(row[7])}')
        # Excel's filter "<>CANCELLED" ignores case but not stray spaces, so the status is compared untrimmed
        want_hidden = ((not opts.show_cancelled and ('' if row[4] is None else str(row[4])).upper() == 'CANCELLED')
                       or (not opts.show_ssl and text(row[7]) == 'SSL'))
        if oi.hidden[i] != want_hidden:
            found['R3e'].append(f'row {r} hidden={oi.hidden[i]}, want {want_hidden}')
        for c in (3, 4):
            if is_date(row[c - 1]) and not is_date_format(sh.ws.cell(row=r, column=c).number_format):
                found['R3f'].append(f'{get_column_letter(c)}{r} not a date format')
    n = len(oi.rows)
    lines = {2: 'medium', 3: 'double'}
    lines.update({3 + i: 'medium' for i in range(1, n) if text(oi.rows[i - 1][0]) != text(oi.rows[i][0])})
    found['R3f'] += grid_problems(sh.ws, 2, 2 + n, 11, lines)
    if sh.ws.auto_filter.ref != f'A2:K{2 + n}':
        found['R3e'].append(f'AutoFilter {sh.ws.auto_filter.ref}, want A2:K{2 + n}')
    want_fc = {}
    if n and not opts.show_cancelled:
        want_fc[4] = ('custom', [('notEqual', 'CANCELLED')])
    if n and not opts.show_ssl:
        want_fc[7] = ('values', ['CTL', 'MX'])
    got_fc = filter_columns(sh.ws)
    if got_fc != want_fc:
        found['R3e'].append(f'filter columns {got_fc}, want {want_fc}')
    return found


def visible_total(oi):
    """This checker's own sum of SUM(F.ORDERED_QUANTITY) over the Order Intake rows not hidden."""
    return sum(row[6] for row, hid in zip(oi.rows, oi.hidden) if not hid and is_number(row[6]))


def check_visible_total(sh, oi):
    """R3g: the visible total in the title row (F1 label, G1 SUBTOTAL 109 over the whole column G), its value
    against the checker's own sum of the visible rows, and the note moved to H1."""
    ws, probs = sh.ws, []
    if text(sh.raw(1, 6)) != OI_TOTAL_LABEL:
        probs.append(f'F1={shown(sh.raw(1, 6))}, want {OI_TOTAL_LABEL!r}')
    f1 = ws['F1']
    if not f1.font.b or f1.alignment.horizontal != 'right':
        probs.append(f'F1 bold {f1.font.b}, aligned {f1.alignment.horizontal}; want bold, right')
    filled = [f'{get_column_letter(c)}1' for c in range(2, 6) if not is_empty(sh.raw(1, c))]
    if filled:
        probs.append(f'{", ".join(filled)} not empty (the F1 label runs over them)')
    want = f'=SUBTOTAL(109,$G$3:$G${2 + max(len(oi.rows), 1)})'
    got = sh.raw(1, 7)
    if not isinstance(got, str) or got.replace(' ', '').upper() != want.upper():
        probs.append(f'G1={shown(got)}, want {want}')
    else:
        value, own = evaluate(sh, 1, 7), visible_total(oi)
        if not same_value(value, own):
            probs.append(f'G1 evaluates to {shown(value)}, own sum of the visible rows {num_text(own)}')
    g1 = ws['G1']
    if not g1.font.b or g1.number_format != '#,##0' or g1.alignment.horizontal != 'left' or not g1.alignment.indent:
        probs.append(f'G1 bold {g1.font.b}, format {g1.number_format}, aligned {g1.alignment.horizontal} indent '
                     f'{g1.alignment.indent}; want bold, #,##0, left indent 1')
    if not OI_NOTE_RE.fullmatch(text(sh.raw(1, 8))):
        probs.append(f'H1={shown(sh.raw(1, 8))}, want the source note')
    return probs


# ------------------------------------------------------------------- R4 Inventory by Month

def check_inventory(sh, inv, opts, cities, months):
    a, b = [], []
    if inv.h is None:
        return ['no detail header row (MSA, PRODUCT_TYPE, ...)'], ['no detail block']
    ws = sh.ws
    if text(sh.raw(3, 1)) != 'Units in':
        a.append(f'A3={shown(sh.raw(3, 1))}')
    if months:
        if inv.month not in months:
            a.append(f'month cell {inv.month!r} not in the months list')
        want_dv = f"'File Format'!$F$4:$F${3 + len(months)}"
        dvs = [dv for dv in ws.data_validations.dataValidation if 'B3' in str(dv.sqref).split()]
        if not dvs or dvs[0].type != 'list' or (dvs[0].formula1 or '').lstrip('=') != want_dv:
            a.append(f'B3 has no list validation {want_dv} ({[(d.type, d.formula1) for d in dvs]})')
    elif inv.month:
        a.append(f'month cell {inv.month!r} but no months list')
    if [text(sh.raw(4, c)) for c in range(1, 6)] != SUM_HEADERS:
        a.append('summary headers')
    if [text(c) for c in inv.cities] != cities:
        a.append(f'summary MSAs {inv.cities}, want {cities}')
    first, last = inv.h + 1, inv.h + max(len(inv.detail), 1)
    rng = lambda col: f'${col}${first}:${col}${last}'
    for i in range(inv.n):
        r = 5 + i
        for j, tname in enumerate(TYPE_NAMES):
            want = f'=SUMIFS({rng("G")},{rng("A")},$A{r},{rng("B")},"{tname}",{rng("F")},$B$3)'
            if sh.raw(r, 2 + j) != want:
                a.append(f'{get_column_letter(2 + j)}{r}={shown(sh.raw(r, 2 + j))}, want {want}')
        if sh.raw(r, 5) != f'=SUM(B{r}:D{r})':
            a.append(f'E{r}={shown(sh.raw(r, 5))}')
    tr = 5 + inv.n
    if text(sh.raw(tr, 1)) != 'Total':
        a.append(f'A{tr} is not Total')
    for col in range(2, 6):
        L = get_column_letter(col)
        want = f'=SUM({L}5:{L}{4 + inv.n})' if inv.n else 0
        if sh.raw(tr, col) != want:
            a.append(f'{L}{tr}={shown(sh.raw(tr, col))}, want {want}')
    if hline(ws, tr, 2) != 'double':
        a.append(f'Total row has no double rule above it')
    if [text(sh.raw(inv.h, c)) for c in range(1, 1 + len(INV_HEADERS))] != INV_HEADERS:
        b.append('detail headers')
    crank = {c.upper(): i for i, c in enumerate(cities)}
    prev = None
    for i, row in enumerate(inv.detail):
        r = inv.h + 1 + i
        t = TYPE_NAMES.index(row[1]) if row[1] in TYPE_NAMES else None
        if t is None:
            b.append(f'B{r}={shown(row[1])} is not a product type name')
            continue
        key = (crank.get(text(row[0]).upper(), -1), t, day_of(row[4]))
        if prev is not None and (key[:2] < prev[:2] or (key[:2] == prev[:2] and (
                (key[2] is not None and prev[2] is None) or (None not in (key[2], prev[2]) and key[2] > prev[2])))):
            b.append(f'row {r} out of order')
        prev = key
        if text(row[5]) != month_text(row[4]):
            b.append(f'F{r}={shown(row[5])}, want {month_text(row[4])!r}')
        want_hidden = (bool(inv.month) and text(row[5]) != inv.month) or (not opts.show_ssl and t == 2)
        if inv.hidden[i] != want_hidden:
            b.append(f'row {r} hidden={inv.hidden[i]}, want {want_hidden}')
        if is_date(row[4]) and not is_date_format(ws.cell(row=r, column=5).number_format):
            b.append(f'E{r} not a date format')
    n = len(inv.detail)
    lines = {inv.h: 'medium', inv.h + 1: 'double'}
    lines.update({inv.h + 1 + i: 'medium' for i in range(1, n) if text(inv.detail[i - 1][0]) != text(inv.detail[i][0])})
    b += grid_problems(ws, inv.h, inv.h + n, len(INV_HEADERS), lines)
    if ws.auto_filter.ref != f'A{inv.h}:I{inv.h + n}':
        b.append(f'AutoFilter {ws.auto_filter.ref}, want A{inv.h}:I{inv.h + n}')
    want_fc = {}
    if n and inv.month:
        want_fc[5] = ('values', [inv.month])
    if n and not opts.show_ssl:
        want_fc[1] = ('values', sorted(TYPE_NAMES[:2]))
    if filter_columns(ws) != want_fc:
        b.append(f'filter columns {filter_columns(ws)}, want {want_fc}')
    return a, b


def summary_values(sh, inv):
    return [[text(sh.raw(5 + i, 1))] + [evaluate(sh, 5 + i, c) for c in (2, 3, 4)] for i in range(inv.n)] if inv.h else []


# ------------------------------------------------------------------- R5 Open Orders by MSA

def check_open_orders(sh, oo, oi, opts, cities, tally):
    """R5a: the table and its formulas, evaluated against the checker's own tally."""
    probs, ws = [], sh.ws
    if text(sh.raw(1, 1)) != OO_SHEET:
        probs.append(f'A1={shown(sh.raw(1, 1))}')
    m = OO_NOTE_RE.fullmatch(text(sh.raw(1, 4)))
    if not m or m.group(1) != ('CTL, MX and SSL' if opts.show_ssl else 'CTL and MX'):
        probs.append(f'D1={shown(sh.raw(1, 4))}')
    if text(sh.raw(3, 1)) != OO_CAPTION:
        probs.append(f'A3={shown(sh.raw(3, 1))}, want {OO_CAPTION!r}')
    if [text(sh.raw(4, c)) for c in range(1, 6)] != SUM_HEADERS:
        probs.append(f'headers {[text(sh.raw(4, c)) for c in range(1, 6)]}, want {SUM_HEADERS}')
    if oo.tr is None:
        return probs + ['no Total row']
    if [text(c) for c in oo.cities] != cities:
        probs.append(f'MSAs {oo.cities}, want {cities}')
    last = 2 + max(len(oi.rows), 1)
    rng = lambda col: f"'Order Intake'!${col}$3:${col}${last}"
    for i in range(oo.n):
        r = 5 + i
        for c in range(2, 5):
            L = get_column_letter(c)
            want = '=' + '+'.join(f'SUMIFS({rng("G")},{rng("B")},$A{r},{rng("H")},{L}$4,{rng("E")},"{st}")'
                                  for st in ('ORDERED', 'Not yet processed'))
            if sh.raw(r, c) != want:
                probs.append(f'{L}{r}={shown(sh.raw(r, c))}, want {want}')
        if sh.raw(r, 5) != f'=SUM(B{r}:D{r})':
            probs.append(f'E{r}={shown(sh.raw(r, 5))}')
    for c in range(2, 6):
        L = get_column_letter(c)
        want = f'=SUM({L}5:{L}{4 + oo.n})' if oo.n else 0
        if sh.raw(oo.tr, c) != want:
            probs.append(f'{L}{oo.tr}={shown(sh.raw(oo.tr, c))}, want {want}')
    if hline(ws, oo.tr, 2) != 'double':
        probs.append('Total row has no double rule above it')
    if not all(ws.cell(row=oo.tr, column=c).font.b for c in range(1, 6)):
        probs.append('Total row not bold')
    got, want = oo.values(), tally.table()
    probs += diff_rows('table row (evaluated) vs own tally', got, want)
    if oo.n == 0 and text(sh.raw(oo.tr + 2, 1)) != OO_NO_CHART:
        probs.append(f'A{oo.tr + 2} is not "{OO_NO_CHART}"')
    return probs


def check_open_orders_chart(sh, oo, opts, path):
    """R5b: the chart, from the zip."""
    facts = chart_facts(path, OO_SHEET)
    if facts is None:
        return ['sheet not found in the zip']
    drawings, charts = facts
    if oo.tr is None:
        return ['no table, so no chart to check']
    if oo.n == 0:
        return [f'{len(charts)} chart(s), want none (no MSAs)'] if charts else []
    probs = []
    if drawings != 1 or len(charts) != 1:
        return [f'{drawings} drawing(s) and {len(charts)} chart(s), want one of each']
    ch = charts[0]
    if ch['kinds'] != ['barChart'] or ch['barDir'] != 'col' or ch['grouping'] != 'clustered':
        probs.append(f'chart {ch["kinds"]} barDir {ch["barDir"]} grouping {ch["grouping"]}, want one clustered column barChart')
    if ch['title'] != OO_CHART_TITLE:
        probs.append(f'chart title {ch["title"]!r}, want {OO_CHART_TITLE!r}')
    if len(ch['series']) != oo.n:
        probs.append(f'{len(ch["series"])} series, want {oo.n} (one per MSA)')
    last = 'D' if opts.show_ssl else 'C'
    for i, (tx, cat, val) in enumerate(ch['series']):
        r = 5 + i
        want = (f'{OO_SHEET}!A{r}', f'{OO_SHEET}!B4:{last}4', f'{OO_SHEET}!B{r}:{last}{r}')
        got = tuple(plain_ref(x) for x in (tx, cat, val))
        if got != want:
            probs.append(f'series {i + 1} refs {got}, want {want}')
    if ch['anchor_row'] is None or ch['anchor_row'] < 6 + oo.n:
        probs.append(f'chart anchored at row {None if ch["anchor_row"] is None else ch["anchor_row"] + 1}, want A{7 + oo.n} or below')
    return probs


# ------------------------------------------------------------------- R6 Run Log

def log_lines(sh):
    return [sh.raw(r, 1) for r in range(4, sh.last_row(1, 4) + 1)]


def check_run_log(sh, cities, tally=None):
    probs = []
    if not text(sh.raw(1, 1)).startswith('Built '):
        probs.append(f'A1={shown(sh.raw(1, 1))}')
    opts = parse_options(sh.raw(2, 1))
    if opts is None:
        probs.append(f'A2 does not parse: {shown(sh.raw(2, 1))}')
    lines = [text(x) for x in log_lines(sh)]
    for c in cities:
        if not any(line.startswith(c + ': ') for line in lines):
            probs.append(f'no Run Log line for MSA {c}')
    if tally is not None:
        for c in cities:                  # the open orders line follows the MSA's inventory line
            want = tally.city_line(c)
            inv_at = next((i for i, l in enumerate(lines) if l.startswith(c + ': Inventory ') or l == c + ': no inventory rows'), None)
            if want not in lines:
                probs.append(f'no Run Log line {want!r}')
            elif inv_at is None or lines.index(want) != inv_at + 1:
                probs.append(f'{want!r} is not the line after the inventory line of {c}')
        got = [l for l in lines if l.startswith('Open orders: ') or 'are not counted on the Open Orders by MSA tab' in l]
        want = tally.tail_lines()
        if len(got) != len(want) or any(g != w and not same_spaced_warning(g, w, tally) for g, w in zip(got, want)):
            probs.append(f'open orders lines {got}, own tally {want}')
    return probs, opts or DEFAULT_OPTIONS


SPACED_RE = re.compile(r'(WARNING: \d+ rows? have ORDER_STATUS ORDERED or Not yet processed with extra spaces and are not '
                       r'counted on the Open Orders by MSA tab: )(.*?)(, \.\.\.)?')


def same_spaced_warning(got, want, tally):
    """The stray-space WARNING lists its texts in export order, which the report's sorted rows cannot show: compare
    the row count, and that the listed texts are distinct stray-space texts of the report, all of them when there
    are ten or fewer, ten and "..." when there are more."""
    g, w = SPACED_RE.fullmatch(got), SPACED_RE.fullmatch(want)
    if not (g and w) or g.group(1) != w.group(1):
        return False
    listed = re.findall(r'"((?:[^"])*)"', g.group(2))
    distinct = set(tally.spaced_texts)
    return (len(listed) == len(set(listed)) == min(10, len(distinct)) and set(listed) <= distinct
            and bool(g.group(3)) == (len(distinct) > 10))


# ------------------------------------------------------------------- E rules

def oi_facts(oi):
    out = []
    for row, hid in zip(oi.rows, oi.hidden):
        d = row[2]
        out.append([text(row[0]), text(row[1]), row[8], row[4], d.strftime('%Y-%m-%d') if is_date(d) else d, hid])
    return out


def detail_facts(inv):
    return [[row[0], row[1], row[2], row[5], row[6], hid] for row, hid in zip(inv.detail, inv.hidden)]


def diff_rows(label, got, want, limit=5):
    probs = []
    if len(got) != len(want):
        probs.append(f'{len(got)} {label}, want {len(want)}')
    for i, (g, w) in enumerate(zip(got, want)):
        if any(not same_value(x, y) for x, y in zip(g, w)) or len(g) != len(w):
            probs.append(f'{label} {i + 1}: {g}, want {w}')
    return probs


def run_expected(tally, expected, which, book, oi, inv, run_log, oo=None):
    a2 = text(run_log.raw(2, 1)) if run_log else ''
    runs = expected.get('runs', [])
    run = next((r for r in runs if (r['name'] == which if which != 'auto' else r['options_line'] == a2)), None)
    if run is None:
        tally.add('E0', '(expected)', [f'no run matching {which if which != "auto" else repr(a2)}'])
        return
    tally.add('E0', '(expected)', [])
    tally.add('E1', '(workbook)', [] if book.names == run['sheet_order'] else [f'sheets {book.names}'])
    tally.add('E2', 'Order Intake', diff_rows('row', oi_facts(oi), run['oi_rows']) if oi else ['no Order Intake'])
    if inv:
        probs = [] if inv.month == run['month_cell'] else [f'month cell {inv.month!r}, want {run["month_cell"]!r}']
        tally.add('E3', 'Inventory by Month', probs + diff_rows('summary row', summary_values(inv.sh, inv), run['summary']))
        tally.add('E4', 'Inventory by Month', diff_rows('detail row', detail_facts(inv), run['detail_rows']))
    else:
        tally.add('E3', 'Inventory by Month', ['no sheet'])
        tally.add('E4', 'Inventory by Month', ['no sheet'])
    probs = diff_rows('file row', build_files(book.sheet('File Format')), run['build_files'])
    item = build_item(book.sheet('File Format'), 'Open orders')
    if 'build_open_orders' in run and item != run['build_open_orders']:
        probs.append(f'Open orders item {item!r}, want {run["build_open_orders"]!r}')
    tally.add('E5', 'File Format', probs)
    got = [text(x) for x in log_lines(run_log)] if run_log else []
    probs = [] if got == run['log_lines'] else [f'Run Log rows differ ({len(got)} here, {len(run["log_lines"])} expected)']
    tally.add('E6', 'Run Log', probs)
    if probs:
        for line in difflib.unified_diff(run['log_lines'], got, 'expected', 'report', lineterm=''):
            print(line)
    if 'open_orders' in run:
        tally.add('E7', OO_SHEET, expected_open_orders(run, book, oo))
    if 'oi_visible_total' in run:
        value = evaluate(oi.sh, 1, 7) if oi else None
        tally.add('E8', 'Order Intake', [] if same_value(value, run['oi_visible_total'])
                  else [f'G1 evaluates to {shown(value)}, want {run["oi_visible_total"]}'])
    if 'classes' in run:                                     # v2.8: the size classes of this build, File Format Table 1
        ff = book.sheet('File Format')
        _, classes, probs = class_table(ff) if ff else (None, [], ['no File Format sheet'])
        got = [list(c) for c in classes]
        tally.add('E9', 'File Format', probs + ([] if got == run['classes'] else [f'Table 1 classes {got}, want {run["classes"]}']))


def expected_open_orders(run, book, oo):
    if oo is None:
        return ['no sheet']
    probs = diff_rows('table row', oo.values(), run['open_orders'])
    summary = chart_summary(chart_facts(book.path, OO_SHEET), oo.sh)
    want = run['open_orders_chart']
    if want['series'] == 0:
        if summary is not None:
            probs.append('a chart, want none')
    elif summary is None or summary[:2] != (want['series'], want['categories']):
        probs.append(f'chart (series, categories) {summary[:2] if summary else None}, want {(want["series"], want["categories"])}')
    return probs


# ------------------------------------------------------------------- C rules

def compare_rows(label, a, b, ha, hb, limit=10):
    probs = []
    if len(a) != len(b):
        probs.append(f'{len(a)} {label} here, {len(b)} in other')
    for i, (x, y) in enumerate(zip(a, b)):
        bad = [get_column_letter(c + 1) for c, (u, v) in enumerate(zip(x, y)) if not same_value(u, v)]
        if bad:
            probs.append(f'{label} {i + 1} differs in {",".join(bad)}: {[shown(v) for v in x]} vs {[shown(v) for v in y]}')
        if ha[i] != hb[i]:
            probs.append(f'{label} {i + 1} hidden {ha[i]} vs {hb[i]}')
    return probs


def run_compare(tally, book, other):
    tally.add('C1', '(workbook)', [] if book.names == other.names else [f'sheets {book.names} vs {other.names}'])
    s1, s2 = book.sheet('Order Intake'), other.sheet('Order Intake')
    if s1 and s2:
        o1, o2 = OI(s1), OI(s2)
        probs = compare_rows('row', o1.rows, o2.rows, o1.hidden, o2.hidden)
        for c in (6, 7):                                 # the visible total: label, formula text and value
            f1, f2, v1, v2 = s1.raw(1, c), s2.raw(1, c), evaluate(s1, 1, c), evaluate(s2, 1, c)
            if not same_value(f1, f2) or not same_value(v1, v2):
                probs.append(f'{get_column_letter(c)}1: {shown(f1)} = {shown(v1)} vs {shown(f2)} = {shown(v2)}')
        tally.add('C2', 'Order Intake', probs)
    else:
        tally.add('C2', 'Order Intake', ['sheet missing'])
    s1, s2 = book.sheet('Inventory by Month'), other.sheet('Inventory by Month')
    if s1 and s2:
        i1, i2 = INV(s1), INV(s2)
        probs = [] if i1.month == i2.month else [f'month cell {i1.month!r} vs {i2.month!r}']
        sv1, sv2 = summary_values(s1, i1), summary_values(s2, i2)
        probs += compare_rows('summary row', sv1, sv2, [False] * len(sv1), [False] * len(sv2))
        probs += compare_rows('detail row', i1.detail, i2.detail, i1.hidden, i2.hidden)
        tally.add('C3', 'Inventory by Month', probs)
    else:
        tally.add('C3', 'Inventory by Month', ['sheet missing'])
    l1, l2 = book.sheet('Run Log'), other.sheet('Run Log')
    if not (l1 and l2):
        tally.add('C4', 'Run Log', ['sheet missing'])
        return
    a, b = [text(l1.raw(2, 1))] + [text(x) for x in log_lines(l1)], [text(l2.raw(2, 1))] + [text(x) for x in log_lines(l2)]
    tally.add('C4', 'Run Log', [] if a == b else [f'Run Log differs ({len(a)} vs {len(b)} lines incl. A2)'])
    if a != b:
        for line in difflib.unified_diff(a, b, 'this', 'other', lineterm=''):
            print(line)
    s1, s2 = book.sheet(OO_SHEET), other.sheet(OO_SHEET)
    if not (s1 and s2):
        tally.add('C5', OO_SHEET, ['sheet missing'])
        return
    o1, o2 = OO(s1), OO(s2)
    probs = []
    if o1.tr != o2.tr:
        probs.append(f'Total row {o1.tr} vs {o2.tr}')
    for r in range(4, max(o1.tr or 4, o2.tr or 4) + 1):
        for c in range(1, 6):
            f1, f2 = s1.raw(r, c), s2.raw(r, c)
            v1, v2 = evaluate(s1, r, c), evaluate(s2, r, c)
            if not same_value(f1, f2) or not same_value(v1, v2):
                probs.append(f'{get_column_letter(c)}{r}: {shown(f1)} = {shown(v1)} vs {shown(f2)} = {shown(v2)}')
    c1, c2 = (chart_summary(chart_facts(bk.path, OO_SHEET), sh) for bk, sh in ((book, s1), (other, s2)))
    if c1 != c2:
        probs.append(f'chart (series, categories, anchor row) {c1} vs {c2}')
    tally.add('C5', OO_SHEET, probs)


# ------------------------------------------------------------------- driver

def open_book(path, label):
    try:
        return Book(path)
    except Exception as exc:
        print(f'error: cannot open {label} {path}: {exc}', file=sys.stderr)
        sys.exit(2)


def main(argv=None):
    ap = argparse.ArgumentParser(description='Verify an Open Orders report (v2) against the data contract.')
    ap.add_argument('report')
    ap.add_argument('--expected', help='expected.json with the run facts (E rules)')
    ap.add_argument('--run', default='auto', choices=['auto', 'default', 'variant', 'setup'])
    ap.add_argument('--compare', help='a second .xlsx to compare with (C rules)')
    args = ap.parse_args(argv)
    book = open_book(args.report, 'report')
    expected = None
    if args.expected:
        try:
            expected = json.load(open(args.expected, encoding='utf-8'))
        except (OSError, ValueError) as exc:
            print(f'error: cannot read expected file {args.expected}: {exc}', file=sys.stderr)
            sys.exit(2)
    other = open_book(args.compare, 'compare file') if args.compare else None

    tally = Tally()
    tally.add('R1', '(workbook)', [] if book.names == SHEETS else [f'sheets {book.names}, want {SHEETS}'])
    ff, ois, invs, oos, log = (book.sheet(n) for n in SHEETS)
    oi = OI(ois) if ois else None
    inv = INV(invs) if invs else None
    oo = OO(oos) if oos else None
    cities = all_cities(oi, inv) if oi and inv else []
    open_tally = OpenTally(oi, cities) if oi and inv else None
    if log:
        probs, opts = check_run_log(log, cities, open_tally)
        tally.add('R6', 'Run Log', probs)
    else:
        tally.add('R6', 'Run Log', ['no Run Log sheet; defaults assumed'])
        opts = DEFAULT_OPTIONS
    for sh, kind in ((ff, 'ff'), (ois, 'oi'), (invs, 'inv'), (oos, 'oo'), (log, 'log')):
        if sh is not None:
            tally.add('R2', sh.name, check_style(sh, kind, oi, inv, oo))
    if ff:
        tally.add('R2F', 'File Format', check_file_format(ff, opts, open_tally))
    if oi:
        found = check_order_intake(ois, oi, opts, cities, class_ranks(ff))
        for rule in ('R3a', 'R3b', 'R3c', 'R3d', 'R3e', 'R3f'):
            tally.add(rule, 'Order Intake', found[rule])
        tally.add('R3g', 'Order Intake', check_visible_total(ois, oi))
    if inv:
        a, b = check_inventory(invs, inv, opts, cities, months_list(ff))
        tally.add('R4a', 'Inventory by Month', a)
        tally.add('R4b', 'Inventory by Month', b)
    if oo and oi and open_tally:
        tally.add('R5a', OO_SHEET, check_open_orders(oos, oo, oi, opts, cities, open_tally))
        tally.add('R5b', OO_SHEET, check_open_orders_chart(oos, oo, opts, book.path))
    else:
        tally.add('R5a', OO_SHEET, ['no Open Orders by MSA sheet' if not oo else 'no Order Intake or Inventory by Month sheet'])
        tally.add('R5b', OO_SHEET, ['no Open Orders by MSA sheet' if not oo else 'no Order Intake or Inventory by Month sheet'])
    if expected is not None:
        run_expected(tally, expected, args.run, book, oi, inv, log, oo)
    if other is not None:
        run_compare(tally, book, other)
    print(f'RESULT: FAIL ({tally.failed} of {tally.checks} checks failed)' if tally.failed
          else f'RESULT: PASS ({tally.checks} checks)')
    return 1 if tally.failed else 0


if __name__ == '__main__':
    sys.exit(main())
