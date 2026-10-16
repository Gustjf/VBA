#!/usr/bin/env python3
"""
Write the size-class master list for the Open Orders report (Config SizeClassFile):
output/open-orders-intake-inventory/size-class-master.xlsx

One sheet "Size Class", header row 1: HP, Rated Operating Capacity, Name of machine, Size Class, Notes.
Rows: the owner's "Bobcat Equipment Size Class" sheet (photo of the screen, 2026-10-15), transcribed as shown,
nothing carried over from the earlier starter list (83 models from the v1 map, replaced):
    Name of machine   the sheet's model column (SSL, CTL, MX tables)
    Rated Operating   loaders: rated operating capacity in lb; excavators: the sheet's operating weight in lb
      Capacity
    HP                the sheet's horsepower; blank where the sheet's cell is blank
    Size Class        the sheet's size class. Only the spelling is changed where the macro needs it:
                      "<2000" -> "0<2000", "3500 & Over" -> "3500&Over". The SSL frame classes (EX Small,
                      Small Frame, Med Frame, Large Frame) are kept as on the sheet: since macro v2.7 they are
                      the four SSL classes the macro reads, and this list is the only source of a class
    Notes             "REVIEW: ..." on every row with something to confirm (a blank cell on the sheet); "as on
                      the sheet" otherwise
Every class must be one of the macro's 17 (build_report_reference.SIZE_CLASSES): the script stops with an error
on any other, so a mistyped class never reaches the workbook.

A second sheet "Set-up" holds the allowed values of the Size Class column in three tables the owner edits (CTL, SSL,
MX; seeded with the macro's 17 classes), and a fourth column that stacks them by formula (CTL, MX, SSL, top to
bottom, blank cells skipped). The Size Class column (D2:D500, room for rows the owner adds) has a drop-down of that
list (the workbook name AllowedClasses), and Excel refuses any other entry. The macro (v2.8) and the reference read
the same three tables as the allowed classes (by their headers: CTL size class, MX size class, SSL size class) and
the rows from the first sheet, so "Size Class" stays first; make_size_class_pdf.py reads only the first sheet.
Run it again only to regenerate the list; the owner's own edits to the workbook replace it.

Usage: python3 make_size_class_master.py [output.xlsx]
"""
import os, sys
import openpyxl
from openpyxl.styles import Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import build_report_reference as ref   # the style sheet (fonts, fills, rules) of the report

DEFAULT_OUT = os.path.join(HERE, '..', '..', 'output', 'open-orders-intake-inventory', 'size-class-master.xlsx')
HEADERS = ['HP', 'Rated Operating Capacity', 'Name of machine', 'Size Class', 'Notes']
REVIEW = 'REVIEW:'
OK = 'as on the sheet'
VALID_ROWS = 500          # the Size Class drop-down covers D2:D500, room for the rows the owner adds
SETUP_SHEET = 'Set-up'
SETUP_SLOTS = 15          # cells per family table on Set-up (the current classes, then empty cells to add to)
SETUP_FIRST = 5           # first value row of the Set-up tables (caption row 3, header row 4)
LIST_NAME = 'AllowedClasses'
# Set-up tables, left to right: (key, column, family, header); the drop-down stacks them in report order CTL, MX, SSL
SETUP_TABLES = [('CTL', 1, 'Track loaders (CTL)', 'CTL size class'), ('SSL', 3, 'Skid-steer loaders (SSL)', 'SSL size class'),
                ('MX', 5, 'Mini excavators (MX)', 'MX size class')]
SETUP_LIST_COL = 7

# Class spellings changed so the macro reads them (sheet spelling -> macro spelling)
SPELLING = {'0<2000': '<2000', '3500&Over': '3500 & Over'}

# (name, size class as on the sheet, HP, ROC or operating weight lb, review notes)
SKIDS = [
    ('S70', 'EX Small', 23.5, 760, []),
    ('S450', 'Small Frame', 49, 1370, []),
    ('S590', 'Med Frame', 68, 2000, []),
    ('S64-2', 'Med Frame', 68, 2300, []),
    ('S650', 'Med Frame', 74, 2690, []),
    ('S66-2', 'Med Frame', 74, 2400, []),
    ('S76-2', 'Med Frame', 74, 2900, []),
    ('S770', 'Large Frame', 92, 3350, []),
    ('S86-2', 'Large Frame', 105, 3500, []),
]
LOADERS = [
    ('T450', '0<2000', 55, 1490, []),
    ('T550', '0<2000', None, None, ['REVIEW: HP and ROC blank on the sheet']),
    ('T595', '2000<2300', 70, 2200, []),
    ('T62', '2000<2300', None, None, ['REVIEW: HP and ROC blank on the sheet']),
    ('T64-2', '2300<2600', 68, 2300, []),
    ('T66-2', '2300<2600', 74, 2450, []),
    ('T650', '2300<2600', 74, 2570, []),
    ('T76-2', '2600<3000', 74, 2900, []),
    ('T740', '3000<3500', 74, None, ['REVIEW: ROC blank on the sheet']),
    ('T770', '3000<3500', 92, 3475, []),
    ('T86-2', '3500&Over', 105, 3800, []),
]
EXCAVATORS = [
    ('E10', '>00-01', 10.2, 2593, []),
    ('E17', '>01-02', 15, 3774, []),
    ('E20', '>01-02', 15, 4167, []),
    ('E26', '>02-03', 24.8, 6489, []),
    ('E32', '>03-04', 24.8, 7403, []),
    ('E35', '>03-04', 24.8, 7659, []),
    ('E38', '>03-04', 33.2, 7893, []),
    ('E40', '>04-05', 40.4, 9180, []),
    ('E48', '>04-05', 42.6, 9830, []),
]
EXCAVATOR_NOTE = 'Rated Operating Capacity column holds the operating weight (lb)'

def rows():
    bad = [(name, cls) for name, cls, *_ in LOADERS + SKIDS + EXCAVATORS if cls not in ref.SIZE_CLASSES]
    if bad:
        raise SystemExit('Not one of the macro\'s size classes (spell it as on the File Format tab): '
                         + ', '.join(f'{n} "{c}"' for n, c in bad))
    out = []
    for name, cls, hp, roc, review in LOADERS + SKIDS:
        notes = list(review)
        if cls in SPELLING:
            notes.insert(0, f'class written {cls} (sheet: {SPELLING[cls]})')
        out.append((hp, roc, name, cls, '; '.join(notes) or OK))
    for name, cls, hp, wt, review in EXCAVATORS:
        out.append((hp, wt, name, cls, '; '.join(review + [EXCAVATOR_NOTE]) if review else f'{OK}; {EXCAVATOR_NOTE}'))
    return out

def write(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Size Class'
    spec = 'RRLLL'
    ref.header_row(ws, 1, 1, HEADERS, spec)
    data = rows()
    groups = {2, 2 + len(LOADERS), 2 + len(LOADERS) + len(SKIDS)}   # first row of each family
    for i, row in enumerate(data, start=2):
        for col, v in enumerate(row, start=1):
            c = ws.cell(row=i, column=col, value=v)
            if col == 2:
                c.number_format = ref.UNITS_FMT
    ref.ruled_table(ws, 1, 1 + len(data), spec, groups)
    ref.finish_fonts(ws)
    for i in range(2, 2 + len(data)):
        ws.cell(row=i, column=5).alignment = Alignment(wrap_text=False)
    # Size Class: a drop-down of the classes on the Set-up tab (the name AllowedClasses); any other entry is refused.
    # A blank cell is allowed (the macro skips that row with a warning).
    dv = DataValidation(type='list', formula1=LIST_NAME, allow_blank=True,
                        showDropDown=False, showInputMessage=True, showErrorMessage=True, errorStyle='stop',
                        promptTitle='Size Class', prompt='Pick the size class from the list (the classes on the Set-up tab).',
                        errorTitle='Not a size class',
                        error='Pick one of the size classes in the list. The list comes from the Set-up tab; the macro '
                              'reads only the classes on the File Format tab of its report.')
    dv.add(f'D2:D{VALID_ROWS}')
    ws.add_data_validation(dv)
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = f'A1:E{1 + len(data)}'
    ws.sheet_view.showGridLines = False
    ref.widths(ws, [8, 24, 18, 14, 110])
    write_setup(wb)
    wb.active = 0                          # opens on Size Class, the sheet the macro reads (the first)
    wb.save(path)
    return len(data)

def setup_families():
    """The macro's 17 classes split into the Set-up tables (CTL, SSL, MX), checked against the reference."""
    c = ref.SIZE_CLASSES
    ctl, mx, ssl = c[:6], c[6:13], c[13:]
    assert (len(ctl), len(mx), len(ssl)) == (6, 7, 4) and ctl[0] == '0<2000' and mx[0] == '>00-01' and ssl[0] == 'EX Small', \
        'SIZE_CLASSES no longer splits into 6 CTL, 7 MX, 4 SSL: update SETUP_TABLES'
    return [ctl, ssl, mx]

def write_setup(wb):
    """The Set-up sheet: three tables of allowed classes (yellow cells: type or clear), the stacked drop-down list
    built from them by formula, and the name AllowedClasses the Size Class drop-down points at."""
    ws = wb.create_sheet(SETUP_SHEET)
    first, last = SETUP_FIRST, SETUP_FIRST + SETUP_SLOTS - 1
    ref.title(ws, 'Set-up', 'Allowed values of the Size Class drop-down on the Size Class tab.', SETUP_LIST_COL, note_col=3)
    rng = {}
    for k, ((key, col, family, head), classes) in enumerate(zip(SETUP_TABLES, setup_families()), start=1):
        ref.caption(ws, first - 2, col, f'Table {k}.', f'{family}.')
        ref.header_row(ws, first - 1, col, [head], 'L')
        for i in range(SETUP_SLOTS):
            c = ws.cell(row=first + i, column=col, value=classes[i] if i < len(classes) else None)
            c.fill = ref.INPUT_FILL
        ref.ruled_table(ws, first - 1, last, 'L', col=col, unbanded=SETUP_SLOTS)   # yellow input cells, no banding
        L = openpyxl.utils.get_column_letter(col)
        rng[key] = (f'${L}${first}:${L}${last}', f'${L}${first}')
    # Table 4: the drop-down list, CTL then MX then SSL (the report's order), blank cells skipped; one array formula
    # per cell (Excel and LibreOffice alike, any Excel version)
    G = openpyxl.utils.get_column_letter(SETUP_LIST_COL)
    nlist = 3 * SETUP_SLOTS
    ref.caption(ws, first - 2, SETUP_LIST_COL, 'Table 4.', 'Drop-down list (formulas).')
    ref.header_row(ws, first - 1, SETUP_LIST_COL, ['Drop-down shows (do not type)'], 'L')
    order = [rng['CTL'], rng['MX'], rng['SSL']]
    counts = [f'COUNTA({r})' for r, _ in order]
    for i in range(nlist):
        row = first + i
        k = f'ROWS(${G}${first}:{G}{row})'
        f = '""'
        for j in reversed(range(3)):
            r, top = order[j]
            before = '-'.join([k] + counts[:j])
            f = (f'IF({k}<=' + '+'.join(counts[:j + 1]) + f',INDEX({r},SMALL(IF({r}<>"",ROW({r})-ROW({top})+1),{before})),{f})')
        cell = f'{G}{row}'
        ws[cell] = ArrayFormula(cell, '=' + f)
    ref.ruled_table(ws, first - 1, first + nlist - 1, 'L', col=SETUP_LIST_COL, unbanded=nlist)
    # AllowedClasses: Table 4 from its first cell down to the last class (INDEX, not the volatile OFFSET)
    sheet = f"'{SETUP_SHEET}'!"
    total = '+'.join(f'COUNTA({sheet}{r})' for r, _ in order)
    wb.defined_names[LIST_NAME] = DefinedName(
        LIST_NAME, attr_text=f'{sheet}${G}${first}:INDEX({sheet}${G}${first}:${G}${first + nlist - 1},MAX(1,{total}))')
    # How to use, to the right of the tables (nothing beyond it, so each line shows in full)
    notes = [
        ('How to use', True),
        ('Add a class: type it in the first empty yellow cell of its table.', False),
        ('Remove a class: select its cell and press Delete. Do not insert or delete whole rows.', False),
        ('The Size Class tab\'s drop-down offers every class in Tables 1 to 3, CTL, then MX, then SSL, top to bottom;', False),
        ('Table 4 shows it as it is now. Each table holds up to %d classes.' % SETUP_SLOTS, False),
        ('', False),
        ('What the macro reads (v2.8)', True),
        ('Its allowed size classes are Tables 1 to 3, found by their headers; keep the header texts as they are.', False),
        ('It ranks them CTL, then MX, then SSL, top to bottom: the order of the report and its File Format tab.', False),
        ('A class added here is read by the next build; a class removed here makes rows that still carry it UNMAPPED', False),
        ('(Check setup names them). The rows come from the first tab: keep Size Class first.', False),
    ]
    note_col = SETUP_LIST_COL + 2
    for i, (text_, bold) in enumerate(notes):
        if text_:
            ref.put(ws, first - 2 + i, note_col, text_).font = ref.BOLD_FONT if bold else ref.BODY_FONT
    ref.finish_fonts(ws)
    ws.sheet_view.showGridLines = False
    ref.widths(ws, [24, 3, 24, 3, 24, 3, 30, 3, 90])

if __name__ == '__main__':
    out = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else DEFAULT_OUT)
    n = write(out)
    print(f'Wrote {n} models to {out}')
