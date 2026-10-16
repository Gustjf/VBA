"""Build synthetic "Snowflake export" Excel fixtures for the v2 report (all cities in one set of files).

Makes, in out_folder:
  - Order Intake 2026.xlsx: Amarillo and Dallas-Fort Worth [North] Metroplex rows with an MSA column headed MSA,
    an ATT row (dropped), a blank row, a row with a blank MSA, ORDER_DATE as text for Dallas (incl. TBD), MSA
    spellings that differ only in case and spaces;
  - Order Intake Brownsville.xlsx: no MSA column (MSA from the file name), shuffled headers, empty row 2;
  - Inventory by Month 2026.xlsx: MSA column headed CITY (the auto-detect fallback), Amarillo and Dallas rows over three months, an
    Attachments row (dropped), no Brownsville rows;
  - Open Orders 2026.xlsx: a few rows with an MSA column, only logged;
  - Size Class.xlsx: the size class file (owner's header spelling; an invalid class, T650 twice, an SSL frame
    class, a lower-case "ex small", an "Over10" that is no longer a class, a ROC column the reader ignores,
    "T740 Compact Track Loader"; T999 and E25 absent, so both are UNMAPPED); one sheet, no Set-up tab, so the
    macro uses its 17 built-in classes;
  - Size Class Set-up.xlsx (v2.8): the same rows (E88 classed >10-15) plus a Set-up tab whose three tables are the
    allowed classes: a blank gap, padding spaces, a class listed in two tables, the reserved word "unmapped",
    >00-01 removed (so E10 is UNMAPPED), >10-15 and XL Frame added (18 classes, so the File Format rows move), and a
    stale copy of the drop-down list in the fourth column, which must not be read; used by the third run "setup";
  - decoys a correct build must ignore (an old report and an Excel lock file);
  - expected.json: the exact content a correct report must contain for two runs ("default", "variant").

expected.json is computed from the data tables below, not by re-reading the workbooks and not with the
reference implementation. The script then asserts hand-computed results against it, prints a summary, and
exits non-zero if an assertion fails.

Usage:
  python3 make_test_fixtures.py [out_folder]
  Default out_folder: test-fixtures next to this script. Created if missing; files written there are
  overwritten; nothing is deleted (remove old fixture files with git rm).
"""
import collections
import datetime
import json
import os
import sys

import openpyxl

DEFAULT_OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'test-fixtures')
AM, DF, BR, NOC = 'Amarillo', 'Dallas-Fort Worth [North] Metroplex', 'Brownsville', '(no MSA)'
N, S = 'Bobcat of Amarillo North', 'Bobcat of Amarillo South'
TYPE_NAME = {'CTL': 'Compact Track Loaders', 'MX': 'Mini Excavators', 'SSL': 'Skid-Steer Loaders'}
# The macro's 17 built-in classes (v2.7, used when the size class file has no Set-up tab), with their family, in
# report order; a run's classes rank 1, 2, ... in its own list, UNMAPPED 99
BUILTIN_CLASSES = ([(c, 'CTL') for c in ('0<2000', '2000<2300', '2300<2600', '2600<3000', '3000<3500', '3500&Over')]
                   + [(c, 'MX') for c in ('>00-01', '>01-02', '>02-03', '>03-04', '>04-05', '>05-06', '>06-10')]
                   + [(c, 'SSL') for c in ('EX Small', 'Small Frame', 'Med Frame', 'Large Frame')])
UNMAPPED_RANK = 99
STATUS_RANK = {'ORDERED': 1, 'NOT YET PROCESSED': 2, 'INVOICED': 3, 'COMPLETELY PROCESSED': 4, 'CANCELLED': 6}
MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October',
          'November', 'December']
SHEET_ORDER = ['File Format', 'Order Intake', 'Inventory by Month', 'Open Orders by MSA', 'Run Log']
# Open Orders by MSA (v2.5): ORDER_STATUS compared as Excel's SUMIFS compares text (whole text, ignoring case, not
# ignoring stray spaces)
OPEN_STATUSES = ('ORDERED', 'NOT YET PROCESSED')
ABBRS = ('CTL', 'MX', 'SSL')
SC_SETUP_FILE = 'Size Class Set-up.xlsx'
OI_FILE, BR_FILE, INV_FILE, OO_FILE, SC_FILE = ('Order Intake 2026.xlsx', 'Order Intake Brownsville.xlsx',
                                                 'Inventory by Month 2026.xlsx', 'Open Orders 2026.xlsx', 'Size Class.xlsx')
OI_HEADER = ['MSA', 'ORDER_DATE', 'REQUESTED_DATE', 'ORDER_STATUS', 'IS_OPEN_ORDER', 'SUM(F.ORDERED_QUANTITY)',
             'PRODUCT_TYPE_ABBR', 'MODEL_GROUP', 'MODEL_NUMBER', 'ORGANIZATION_NAME']
OI_HEADER_BROWNSVILLE = ['MODEL_GROUP', 'ORGANIZATION_NAME', 'ORDER_STATUS', 'PRODUCT_TYPE_ABBR', 'ORDER_DATE',
                         'MODEL_NUMBER', 'SUM(F.ORDERED_QUANTITY)', 'REQUESTED_DATE', 'IS_OPEN_ORDER']
INV_HEADER = ['ORGANIZATION_NAME', 'ORGANIZATION_CUSTOMER_NUMBER', 'PRODUCT_TYPE', 'MODEL_GROUP', 'MODEL_NUMBER',
              'INVENTORY_DATE', 'TOTAL', 'CITY']
OO_HEADER = ['SALES_ORDER_NUMBER', 'ORDER_DATE', 'REQUESTED_DATE', 'ORDER_STATUS', 'IS_OPEN_ORDER',
             'SUM(F.ORDERED_QUANTITY)', 'ORACLE_ORDER', 'PRODUCT_TYPE_ABBR', 'MODEL_GROUP', 'MODEL_NUMBER',
             'ORGANIZATION_NAME', 'ORGANIZATION_CUSTOMER_NUMBER', 'MSA']
SC_HEADER = ['HP', 'Rated Operator Capacity', 'Name of machine', 'Size Class']


def D(text):
    """'2026-09-08' -> datetime(2026, 9, 8); '2026-09-08 14:30' -> datetime(2026, 9, 8, 14, 30)."""
    return datetime.datetime.strptime(text, '%Y-%m-%d %H:%M' if ' ' in text else '%Y-%m-%d')


# A str ORDER_DATE is written as text; D(...) values are written as real dates. city None = blank cell.
OIRow = collections.namedtuple('OIRow', 'id city order_date requested status is_open qty abbr group model org size')
InvRow = collections.namedtuple('InvRow', 'id city org cust ptype group model date total')

# ------------------------------------------------------------------ data tables
# Order Intake 2026.xlsx in file order. None = the blank row. Last field: expected size class.
OI_2026 = [
    OIRow('A01', AM, D('2026-09-15'), D('2026-10-15'), 'ORDERED', True, 1, 'CTL', 'T66-2', 'T66-2 T4V1', N, '2300<2600'),
    OIRow('A02', AM, D('2026-08-20'), D('2026-09-30'), 'INVOICED', False, 1, 'CTL', 'T770', 'T770 T4 V3', S, '3000<3500'),
    OIRow('A03', AM, D('2026-09-02'), D('2026-11-01'), 'CANCELLED', False, 1, 'CTL', 'T595', 'T595 T4', N, '2000<2300'),
    OIRow('A04', AM, D('2026-07-11'), D('2026-08-11'), 'ORDERED', True, 1, 'MX', 'E35', 'E35R2 25HP', N, '>03-04'),
    OIRow('A05', AM, D('2026-09-18'), D('2026-12-01'), 'Not yet processed', True, 1, 'MX', 'E25', 'E25 R2', S, 'UNMAPPED'),
    OIRow('A06', AM, D('2026-08-05'), D('2026-09-05'), 'ORDERED', True, 1, 'SSL', 'S590', 'S590 T4', N, 'Med Frame'),
    OIRow('A07', AM, D('2026-06-30'), D('2026-07-30'), 'ORDERED', True, 1, 'CTL', 'T999', 'T999 PROTO', S, 'UNMAPPED'),
    OIRow('A08', AM, D('2026-09-01'), D('2026-10-01'), 'CANCELLED', False, 1, 'MX', 'E48', 'E48 T4 V1', S, '>04-05'),
    OIRow('A09', AM, D('2026-08-14'), D('2026-09-14'), 'CANCELLED', False, 1, 'SSL', 'S76', 'S76 T4', N, '2600<3000'),
    OIRow('A10', AM, D('2026-09-15'), D('2026-10-20'), 'ORDERED', True, 2, 'CTL', 'T66-2', 'T66-2 T4V2', S, '2300<2600'),
    OIRow('A11', ' amarillo ', D('2026-05-19'), D('2026-06-19'), 'Completely processed', False, 1, 'CTL', 'T450',
          'T450 T4', N, '0<2000'),
    OIRow('A12', AM, D('2026-09-25'), D('2026-10-25'), 'ORDERED', True, 1, 'CTL', 'T86-2', 'T86-2 T4', S, '3500&Over'),
    OIRow('A13', AM, D('2026-08-28'), D('2026-09-28'), 'ORDERED', True, 1, 'MX', 'E10', 'E10 T4', N, '>00-01'),
    OIRow('A14', AM, D('2026-07-22'), D('2026-08-22'), 'INVOICED', False, 1, 'MX', 'E88', 'E88 T4', S, '>06-10'),
    OIRow('A15', AM, D('2026-09-09'), D('2026-09-30'), 'ORDERED', True, 1, 'CTL', 'T76-2', 'T76-2 T4', N, '2600<3000'),
    OIRow('A16', AM, D('2026-08-03'), D('2026-09-03'), 'ORDERED', True, 1, 'MX', 'E20', 'E20 T4', S, '>01-02'),
    OIRow('A17', AM, D('2026-07-01'), D('2026-09-17'), 'ORDERED', True, 1, 'CTL', 'T64-2', 'T64-2 T4V1', N, '2300<2600'),
    OIRow('A18', AM, D('2026-09-20'), D('2026-10-30'), 'ORDERED', True, 1, 'CTL', 'T66-2', 'T66-2 T4V1', N, '2300<2600'),
    OIRow('A19', AM, D('2026-08-11'), D('2026-09-11'), 'Cancelled', False, 1, 'CTL', 'T740', 'T740 T4', S, '3000<3500'),
    OIRow('A20', AM, D('2026-09-04'), D('2026-10-04'), 'ORDERED', True, 1, ' mx ', 'e35', 'e35 R2', N, '>03-04'),
    None,
    OIRow('C01', DF, '2026-08-12', D('2026-09-12'), 'ORDERED', True, 1, 'CTL', 'T740', 'T740 T4', 'Bobcat of Dallas', '3000<3500'),
    OIRow('C02', DF, '2026-08-05 00:00:00.000', D('2026-09-05'), 'Not yet processed', True, 1, 'CTL', 'T740', 'T740 T4 V2',
          'Bobcat of Fort Worth', '3000<3500'),
    OIRow('C03', DF, 'TBD', D('2026-10-01'), 'ORDERED', True, 1, 'CTL', 'T740', 'T740 T4', 'Bobcat of Dallas', '3000<3500'),
    OIRow('C04', DF, '2026-07-30', D('2026-08-30'), 'INVOICED', False, 1, 'MX', 'E55', 'E55 T4', 'Bobcat of Fort Worth', '>05-06'),
    OIRow('C05', 'Dallas-Fort Worth  [North]  Metroplex', '2026-08-19', D('2026-09-19'), 'CANCELLED', False, 1, 'MX', 'E26',
          'E26 T4', 'Bobcat of Dallas', '>02-03'),
    OIRow('C06', DF, '2026-08-21T00:00:00Z', D('2026-09-21'), 'ORDERED', True, 1, 'MX', 'E60', 'E60 T4', 'Bobcat of Dallas', '>05-06'),
    OIRow('C07', DF, '2026-08-02', D('2026-08-15'), 'ORDERED', True, 1, 'CTL', 'T76-2', 'T76-2 T4', 'Bobcat of Fort Worth', '2600<3000'),
    OIRow('X01', AM, D('2026-09-03'), D('2026-10-03'), 'ORDERED', True, 4, 'ATT', 'AT-100', 'AT-100 BUCKET', N, None),
    OIRow('X02', None, D('2026-09-15'), D('2026-10-10'), 'ORDERED', True, 1, 'CTL', 'T650', 'T650 T4', 'Bobcat of Nowhere', '2300<2600'),
]
OI_2026_FIRST_ROW = 2
# Order Intake Brownsville.xlsx: no MSA column; shuffled header; empty row 2; one data row in row 3.
OI_BROWNSVILLE = [
    OIRow('B01', BR, D('2026-09-10'), D('2026-10-10'), 'ORDERED', True, 1, 'CTL', 'T650', 'T650 T4 V2',
          'Bobcat of Brownsville', '2300<2600'),
]
OI_BROWNSVILLE_FIRST_ROW = 3
# Inventory by Month 2026.xlsx in file order. Fields: id, CITY, ORGANIZATION_NAME, customer no., type abbr (or
# another type name), MODEL_GROUP, MODEL_NUMBER, INVENTORY_DATE, TOTAL.
INV_2026 = [
    InvRow('I01', AM, N, 900101, 'MX', 'E35', 'E35R2 25HP', D('2026-09-08'), 2),
    InvRow('I02', AM, N, 900102, 'CTL', 'T66-2', 'T66-2 T4V1', D('2026-09-08'), 3),
    InvRow('I03', AM, S, 900103, 'SSL', 'S590', 'S590 T4', D('2026-09-08'), 4),
    InvRow('I04', AM, S, 900104, 'CTL', 'T770', 'T770 T4 V3', D('2026-09-08'), 1),
    InvRow('I05', 'AMARILLO', N, 900105, 'MX', 'E10', 'E10 T4', D('2026-09-08'), 5),
    InvRow('I06', AM, N, 900106, 'CTL', 'T66-2', 'T66-2 T4V1', D('2026-08-08'), 6),
    InvRow('I07', AM, S, 900107, 'MX', 'E88', 'E88 T4', D('2026-08-08'), 2),
    InvRow('Z01', AM, N, 900199, 'Attachments', 'AT-100', 'AT-100 BUCKET', D('2026-09-08'), 7),
    InvRow('I08', AM, S, 900108, 'CTL', 'T450', 'T450 T4', D('2026-09-08'), 2),
    InvRow('I09', AM, N, 900109, 'SSL', 'S76', 'S76 T4', D('2026-08-08'), 1),
    InvRow('I10', AM, S, 900110, 'MX', 'E20', 'E20 T4', D('2026-09-08 14:30'), 1),
    InvRow('D01', DF, 'Bobcat of Dallas', 910001, 'CTL', 'T740', 'T740 T4', D('2026-08-08'), 3),
    InvRow('D02', DF, 'Bobcat of Dallas', 910002, 'MX', 'E55', 'E55 T4', D('2026-08-08'), 2),
    InvRow('D03', DF, 'Bobcat of Fort Worth', 910003, 'CTL', 'T76-2', 'T76-2 T4', D('2026-07-08'), 4),
    InvRow('D04', DF, 'Bobcat of Fort Worth', 910004, 'SSL', 'S650', 'S650 T4', D('2026-08-08'), 1),
    InvRow('D05', DF, 'Bobcat of Dallas', 910005, 'MX', 'E26', 'E26 T4', D('2026-07-08'), 1),
    InvRow('D06', DF, 'Bobcat of Fort Worth', 910006, 'CTL', 'T740', 'T740 T4 V2', D('2026-08-08'), 2),
]
# Open Orders 2026.xlsx: open rows of A01, A04, C01, C06 (only logged).
OO_IDS = ['A01', 'A04', 'C01', 'C06']
# Size Class.xlsx rows from row 2: (HP, ROC, name, class). Expected: 20 usable rows and the three warnings below.
# v2.7: the ROC column is ignored (T595's 2350 no longer warns); rows 23 and 24 are appended so no row number moves.
SIZE_CLASS_ROWS = [
    (74, 2450, 'T66-2', '2300<2600'), (92, 3475, 'T770', '3000<3500'), (74, 2350, 'T595', '2000<2300'),
    (24.8, None, 'E35', '>03-04'), (66, 2100, 'S590', 'Med Frame'), (41.7, None, 'E48', '>04-05'),
    (74, 2900, 'S76', '2600<3000'), (49, 1400, 'T450', '0<2000'), (105, 3800, 'T86-2', '3500&Over'),
    (10.2, None, 'E10', '>00-01'), (72.4, None, 'E88', '>06-10'), (74, 2900, 'T76-2', '2600<3000'),
    (17.6, None, 'E20', '>01-02'), (66, 2300, 'T64-2', '2300<2600'), (92, 3200, 'T740 Compact Track Loader', '3000<3500'),
    (48, None, 'E55', '>05-06'), (24.8, None, 'E26', '>02-03'), (48, None, 'E60', '>05-06'),
    (74, 2570, 'T650', '2300<2600'), (68, 2230, 'T630', '2000-2300'), (74, 2675, 'T650', '2600<3000'),
    (23.5, 760, 'S70', 'ex small'), (None, None, 'E100', 'Over10'),
]
SIZE_CLASS_WARNINGS = [
    'Size class file row 21 T630: "2000-2300" is not a size class - row skipped',
    'Size class file row 22 T650: listed again (first in row 20) - row 20 is used',
    'Size class file row 24 E100: "Over10" is not a size class - row skipped',
]
# Size Class Set-up.xlsx: the same rows, but E88 (row 12) classed >10-15, a class only its Set-up tab allows
SIZE_CLASS_ROWS_SETUP = [r if r[2] != 'E88' else (r[0], r[1], 'E88', '>10-15') for r in SIZE_CLASS_ROWS]
# Its Set-up tab: title A1, captions row 3, headers row 4, values from row 5 (None = an empty cell), the layout of
# size-class-master.xlsx; columns A (CTL), C (SSL), E (MX); G holds a stale drop-down list the macro must not read
SETUP_FIRST_ROW = 5
SETUP_TABLES = [
    ('A', 'Table 1. Track loaders (CTL).', 'CTL size class',
     ['0<2000', '2000<2300', None, '2300<2600', '2600<3000', ' 3000<3500 ', '3500&Over']),
    ('C', 'Table 2. Skid-steer loaders (SSL).', 'SSL size class',
     ['EX Small', 'Small Frame', 'Med Frame', 'Large Frame', '2000<2300', 'unmapped', 'XL Frame']),
    ('E', 'Table 3. Mini excavators (MX).', 'MX size class',
     ['>01-02', '>02-03', '>03-04', '>04-05', '>05-06', '>06-10', '>10-15']),
    ('G', 'Table 4. Drop-down list (formulas).', 'Drop-down shows (do not type)',
     ['STALE 1', 'STALE 2', '>00-01']),
]
# Hand-worked from SETUP_TABLES: read CTL, then MX, then SSL, top to bottom; blank skipped, trimmed, a repeat
# (ignoring case) or the word UNMAPPED warned and ignored
SETUP_CLASSES = ([(c, 'CTL') for c in ('0<2000', '2000<2300', '2300<2600', '2600<3000', '3000<3500', '3500&Over')]
                 + [(c, 'MX') for c in ('>01-02', '>02-03', '>03-04', '>04-05', '>05-06', '>06-10', '>10-15')]
                 + [(c, 'SSL') for c in ('EX Small', 'Small Frame', 'Med Frame', 'Large Frame', 'XL Frame')])
SETUP_WARNINGS = [
    'Set-up tab C9: "2000<2300" is listed again (first in A6) - ignored',
    'Set-up tab C10: "unmapped" is not a size class name (UNMAPPED marks models not in the file) - ignored',
    'Size class file row 11 E10: ">00-01" is not a size class - row skipped',
] + SIZE_CLASS_WARNINGS
# What each size class file gives a run: the file line's counts, the classes line, the warnings, the class list and
# the Order Intake rows whose class differs from the OIRow size field
SC_SPECS = {
    SC_FILE: {'usable': 20, 'warnings': SIZE_CLASS_WARNINGS, 'classes': BUILTIN_CLASSES, 'overrides': {},
              'classes_line': 'Size classes: the 17 built into the macro (%s has no Set-up tab)' % SC_FILE},
    SC_SETUP_FILE: {'usable': 19, 'warnings': SETUP_WARNINGS, 'classes': SETUP_CLASSES,
                    'overrides': {'A13': 'UNMAPPED', 'A14': '>10-15'},
                    'classes_line': 'Size classes: 18 from the Set-up tab (CTL 6, MX 7, SSL 5)'},
}
# The three runs. The 'options' keys and their order are the Run Log options line.
RUNS = [
    {'name': 'default', 'month': None,
     'options': {'InventoryMonth': 'latest in file', 'ShowCancelled': 'FALSE', 'ShowSSL': 'FALSE',
                 'MSAColumn': 'auto', 'SizeClassFile': SC_FILE}},
    {'name': 'variant', 'month': '2026-08-15',
     'options': {'InventoryMonth': 'August 2026', 'ShowCancelled': 'TRUE', 'ShowSSL': 'TRUE',
                 'MSAColumn': 'auto', 'SizeClassFile': SC_FILE}},
    {'name': 'setup', 'month': None,
     'options': {'InventoryMonth': 'latest in file', 'ShowCancelled': 'FALSE', 'ShowSSL': 'FALSE',
                 'MSAColumn': 'auto', 'SizeClassFile': SC_SETUP_FILE}},
]


# ------------------------------------------------------------------ expected-value helpers
def parse_text_date(text):
    """Text ORDER_DATE -> datetime, or None when it does not parse (e.g. 'TBD')."""
    s = text.strip()
    if s.endswith('Z'):
        s = s[:-1]
    if len(s) >= 11 and s[10] == 'T':
        s = s[:10] + ' ' + s[11:]
    if s.endswith('.000'):
        s = s[:-4]
    for fmt in ('%Y-%m-%d', '%Y-%m-%d %H:%M:%S'):
        try:
            return datetime.datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None


def city_text(raw):
    """Trimmed, runs of spaces collapsed; blank = (no city)."""
    return ' '.join((raw or '').split()) or NOC


def first_spellings():
    """Upper-case city -> first spelling seen, in reading order: Order Intake files (alphabetical, kept rows),
    then the Inventory file (kept rows)."""
    seen = {}
    for r in OI_2026:
        if r is not None and r.abbr.strip().upper() in ('CTL', 'MX', 'SSL'):
            seen.setdefault(city_text(r.city).upper(), city_text(r.city))
    seen.setdefault(BR.upper(), BR)
    for r in INV_2026:
        if r.ptype in TYPE_NAME:
            seen.setdefault(city_text(r.city).upper(), city_text(r.city))
    return seen


def canon_city(raw):
    return first_spellings()[city_text(raw).upper()]


def month_of(d):
    return '%s %d' % (MONTHS[d.month - 1], d.year)


def rows_word(n):
    return '%d row%s' % (n, '' if n == 1 else 's')


def oi_kept():
    """Kept Order Intake rows with their expected city, in export order (files alphabetical, rows in order)."""
    kept = []
    for r in [r for r in OI_2026 if r is not None]:
        if r.abbr.strip().upper() in ('CTL', 'MX', 'SSL'):
            kept.append((r, canon_city(r.city)))
    for r in OI_BROWNSVILLE:
        kept.append((r, BR))
    return kept


def inv_kept():
    return [(r, canon_city(r.city)) for r in INV_2026 if r.ptype in TYPE_NAME]


def cities_sorted():
    names = {c for _, c in oi_kept()} | {c for _, c in inv_kept()}
    return sorted(names, key=str.upper)


def is_open(status):
    """An open order row: ORDER_STATUS ORDERED or Not yet processed, as SUMIFS compares (case ignored, spaces not)."""
    return (status or '').upper() in OPEN_STATUSES


def open_orders_facts(cities):
    """Per city: (rows, ORDERED rows, units by type) over the open rows (all three types); then the stray-space
    statuses (distinct texts, export order, with their row count) and the IS_OPEN_ORDER disagreements."""
    per = {c: [0, 0, {t: 0 for t in ABBRS}] for c in cities}
    spaced_rows, spaced_texts, disagree = 0, [], 0
    for r, c in oi_kept():
        open_ = is_open(r.status)
        if open_:
            per[c][0] += 1
            per[c][1] += r.status.upper() == 'ORDERED'
            if isinstance(r.qty, (int, float)) and not isinstance(r.qty, bool):
                per[c][2][r.abbr.strip().upper()] += r.qty
        elif r.status.strip(' ').upper() in OPEN_STATUSES:
            spaced_rows += 1
            if r.status not in spaced_texts:
                spaced_texts.append(r.status)
        if isinstance(r.is_open, bool) and r.is_open != open_:
            disagree += 1
    return per, spaced_rows, spaced_texts, disagree


def compute_run(run):
    opts = run['options']
    show_c, show_s = opts['ShowCancelled'] == 'TRUE', opts['ShowSSL'] == 'TRUE'
    cities = cities_sorted()
    rank = {c: i for i, c in enumerate(cities)}
    inv = inv_kept()
    months = sorted({(r.date.year, r.date.month) for r, _ in inv})
    if run['month']:
        chosen = datetime.datetime.strptime(run['month'], '%Y-%m-%d')
        chosen = (chosen.year, chosen.month)
    else:
        chosen = max(months)
    month_text = '%s %d' % (MONTHS[chosen[1] - 1], chosen[0])

    # Order Intake: every kept row, sorted by class, status, newest date (undated last), city, export order
    sc = SC_SPECS[opts['SizeClassFile']]
    class_rank = {c: i for i, (c, _) in enumerate(sc['classes'], start=1)}
    size_of = lambda r: sc['overrides'].get(r.id, r.size)
    rows = []
    for k, (r, city) in enumerate(oi_kept()):
        day = r.order_date if not isinstance(r.order_date, str) else parse_text_date(r.order_date)
        status_rank = STATUS_RANK.get(r.status.strip().upper(), 5)
        hidden = (not show_c and r.status.upper() == 'CANCELLED') or (not show_s and r.abbr.strip().upper() == 'SSL')
        shown_date = day.strftime('%Y-%m-%d') if day is not None else r.order_date
        rows.append(((class_rank.get(size_of(r), UNMAPPED_RANK), status_rank, day is None,
                      -(day.toordinal() if day else 0), rank[city], k),
                     [size_of(r), city, r.group, r.status, shown_date, hidden], r.id))
    rows.sort(key=lambda x: x[0])
    oi_rows, oi_ids = [x[1] for x in rows], [x[2] for x in rows]
    # v2.6: the visible total in Order Intake G1, the ordered quantity of the rows the filters show
    hidden_ids = {i for i, row in zip(oi_ids, oi_rows) if row[5]}
    visible_total = sum(r.qty for r, _ in oi_kept() if r.id not in hidden_ids)

    # Inventory: summary units for the chosen month (all three types), detail sorted city, type, newest date
    summary = []
    for c in cities:
        units = [sum(r.total for r, rc in inv if rc == c and r.ptype == t and (r.date.year, r.date.month) == chosen)
                 for t in ('CTL', 'MX', 'SSL')]
        summary.append([c] + units)
    det = sorted(enumerate(inv), key=lambda x: (rank[x[1][1]], ['CTL', 'MX', 'SSL'].index(x[1][0].ptype),
                                                -x[1][0].date.toordinal(), x[0]))
    detail = [[c, TYPE_NAME[r.ptype], r.group, month_of(r.date), r.total,
               (r.date.year, r.date.month) != chosen or (not show_s and r.ptype == 'SSL')] for _, (r, c) in det]
    detail_ids = [r.id for _, (r, c) in det]

    # Run Log, in the order of plan Section 5.2.4
    log = [
        'File %s: 30 rows -> 28 written, 1 dropped (ATT 1), 1 blank ignored; MSA column MSA, 3 MSAs' % OI_FILE,
        'File %s: 2 rows -> 1 written, 1 blank ignored; MSA from file name (Brownsville)' % BR_FILE,
        'File %s: 17 rows -> 16 written, 1 dropped (Attachments 1); MSA column CITY, 2 MSAs' % INV_FILE,
        'Size class file %s: %d usable rows, %d warnings' % (opts['SizeClassFile'], sc['usable'], len(sc['warnings'])),
        sc['classes_line'],
    ] + sc['warnings']
    unmapped = collections.OrderedDict()
    for r, c in oi_kept():
        if size_of(r) == 'UNMAPPED':
            unmapped.setdefault(c, [])
            if r.group.strip() not in unmapped[c]:
                unmapped[c].append(r.group.strip())
    open_per, spaced_rows, spaced_texts, disagree = open_orders_facts(cities)
    for c in cities:
        mine = [(r, row) for (r, cc), row in zip(oi_kept(), [None] * 99) if cc == c]
        if mine:
            n = {t: sum(r.abbr.strip().upper() == t for r, _ in mine) for t in ('CTL', 'MX', 'SSL')}
            canc = sum(r.status.strip().upper() == 'CANCELLED' for r, _ in mine)
            hid = sum(1 for row in oi_rows if row[1] == c and row[5])
            line = '%s: Order Intake %s (CTL %d, MX %d, SSL %d; %d CANCELLED), %d hidden by the filters' % (
                c, rows_word(len(mine)), n['CTL'], n['MX'], n['SSL'], canc, hid)
            if c in unmapped:
                line += '; UNMAPPED model groups: ' + ', '.join(unmapped[c])
            log.append(line)
        else:
            log.append('%s: no Order Intake rows' % c)
        n_inv = sum(1 for _, rc in inv if rc == c)
        if n_inv:
            units = next(s for s in summary if s[0] == c)[1:]
            log.append('%s: Inventory %s; %s: CTL %d, MX %d, SSL %d units' % (c, rows_word(n_inv), month_text, *units))
        else:
            log.append('%s: no inventory rows' % c)
        n_open, n_ordered, units = open_per[c]
        if n_open:
            log.append('%s: Open orders %s (ORDERED %d, Not yet processed %d): CTL %d, MX %d, SSL %d units' % (
                c, rows_word(n_open), n_ordered, n_open - n_ordered, units['CTL'], units['MX'], units['SSL']))
        else:
            log.append('%s: no open orders (ORDERED or Not yet processed)' % c)
    log += [
        'File %s: 4 rows, MSA column MSA, 2 MSAs - not used by this report' % OO_FILE,
        '%s ORDER_DATE: 6 text value(s) converted to dates, 1 not a date and left as is (row 25)' % OI_FILE,
        '2 Order Intake files found, all are read: %s, %s' % (OI_FILE, BR_FILE),
        '%s: 1 row with a blank MSA - shown as (no MSA)' % OI_FILE,
    ]
    # Open Orders by MSA lines after the inventory "shows zeros" block (none of which the fixtures produce)
    total_open = sum(p[0] for p in open_per.values())
    if total_open == 0:
        log.append('Open orders: no rows with ORDER_STATUS ORDERED or Not yet processed - the Open Orders by MSA tab shows zeros')
    if spaced_rows:
        log.append('WARNING: %s have ORDER_STATUS ORDERED or Not yet processed with extra spaces and are not counted on '
                   'the Open Orders by MSA tab: %s' % (rows_word(spaced_rows), ', '.join('"%s"' % t for t in spaced_texts[:10])
                                                       + (', ...' if len(spaced_texts) > 10 else '')))
    if disagree:
        log.append('Open orders: %s where IS_OPEN_ORDER disagrees with the status (TRUE with another status, or FALSE '
                   'with ORDERED / Not yet processed)' % rows_word(disagree))
    open_orders = [[c] + [open_per[c][2][t] for t in ABBRS] + [sum(open_per[c][2].values())] for c in cities]
    open_orders.append(['Total'] + [sum(row[k] for row in open_orders) for k in range(1, 5)])
    tot = open_orders[-1]
    build_open = ('ORDERED and Not yet processed: %s, %d units (CTL %d, MX %d, SSL %d), on the Open Orders by MSA tab'
                  % (rows_word(total_open), tot[4], tot[1], tot[2], tot[3]))
    files = [['Order Intake file', log[0][len('File '):]], ['Order Intake file', log[1][len('File '):]],
             ['Inventory by Month file', log[2][len('File '):]],
             ['Open Order file', 'Open Orders 2026.xlsx: 4 rows, MSA column MSA, 2 MSAs - not used by this report']]
    return {'name': run['name'], 'options': opts,
            'options_line': 'Options: ' + ', '.join('%s=%s' % kv for kv in opts.items()),
            'sheet_order': SHEET_ORDER, 'cities': cities, 'months': [month_of(datetime.datetime(y, m, 1)) for y, m in months],
            'month_cell': month_text, 'oi_rows': oi_rows, 'oi_ids': oi_ids, 'oi_visible_total': visible_total,
            'summary': summary,
            'detail_rows': detail, 'detail_ids': detail_ids, 'build_files': files,
            'open_orders': open_orders,
            'open_orders_split': [[c, open_per[c][1], open_per[c][0] - open_per[c][1]] for c in cities],
            'open_orders_chart': {'series': len(cities), 'categories': ['CTL', 'MX', 'SSL'] if show_s else ['CTL', 'MX']},
            'build_open_orders': build_open, 'log_lines': log,
            'classes': [list(c) for c in sc['classes']]}


def build_expected():
    return {'generated_by': 'make_test_fixtures.py', 'version': '2.8',
            'columns': {'oi_rows': ['Size Class', 'MSA', 'MODEL_GROUP', 'ORDER_STATUS', 'ORDER_DATE', 'hidden'],
                        'summary': ['MSA', 'CTL', 'MX', 'SSL'],
                        'detail_rows': ['MSA', 'PRODUCT_TYPE', 'MODEL_GROUP', 'Month', 'TOTAL', 'hidden'],
                        'open_orders': ['MSA', 'CTL', 'MX', 'SSL', 'Total'],
                        'open_orders_split': ['MSA', 'ORDERED rows', 'Not yet processed rows'],
                        'classes': ['Size class', 'Family (CTL, MX or SSL)']},
            'runs': [compute_run(run) for run in RUNS]}


# ------------------------------------------------------------------ writers
def write_book(path, title, header, records, first_row=2, blank_rows=()):
    """One sheet: header in row 1, then records (dicts keyed by header name) from first_row, skipping the
    worksheet rows in blank_rows. None stays empty."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = title
    for col, name in enumerate(header, start=1):
        ws.cell(row=1, column=col, value=name)
    row = first_row
    for rec in records:
        while row in blank_rows:
            row += 1
        for col, name in enumerate(header, start=1):
            value = rec.get(name)
            if value is not None:
                cl = ws.cell(row=row, column=col, value=value)
                if isinstance(value, datetime.datetime):
                    cl.number_format = 'mm-dd-yy'
        row += 1
    wb.save(path)


def oi_record(r, with_city=True):
    rec = {'ORDER_DATE': r.order_date, 'REQUESTED_DATE': r.requested, 'ORDER_STATUS': r.status,
           'IS_OPEN_ORDER': r.is_open, 'SUM(F.ORDERED_QUANTITY)': r.qty, 'PRODUCT_TYPE_ABBR': r.abbr,
           'MODEL_GROUP': r.group, 'MODEL_NUMBER': r.model, 'ORGANIZATION_NAME': r.org}
    if with_city:
        rec['MSA'] = r.city
    return rec


def write_setup_file(path):
    """Size Class Set-up.xlsx: sheet Size Class (the rows) first, then the Set-up tab of SETUP_TABLES."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Size Class'
    for col, name in enumerate(SC_HEADER, start=1):
        ws.cell(row=1, column=col, value=name)
    for i, row in enumerate(SIZE_CLASS_ROWS_SETUP, start=2):
        for col, value in enumerate(row, start=1):
            if value is not None:
                ws.cell(row=i, column=col, value=value)
    st = wb.create_sheet('Set-up')
    st['A1'] = 'Set-up'
    st['C1'] = 'Allowed values of the Size Class drop-down on the Size Class tab.'
    for col, caption, header, values in SETUP_TABLES:
        st[f'{col}{SETUP_FIRST_ROW - 2}'] = caption
        st[f'{col}{SETUP_FIRST_ROW - 1}'] = header
        for i, value in enumerate(values):
            if value is not None:
                st[f'{col}{SETUP_FIRST_ROW + i}'] = value
    st['I3'] = 'How to use'
    st['I4'] = 'Add a class: type it in the first empty yellow cell of its table.'
    wb.save(path)


def write_fixtures(out):
    os.makedirs(out, exist_ok=True)
    blank_at = OI_2026_FIRST_ROW + OI_2026.index(None)
    write_book(os.path.join(out, OI_FILE), 'Order Intake', OI_HEADER,
               [oi_record(r) for r in OI_2026 if r is not None], OI_2026_FIRST_ROW, {blank_at})
    write_book(os.path.join(out, BR_FILE), 'Order Intake', OI_HEADER_BROWNSVILLE,
               [oi_record(r, False) for r in OI_BROWNSVILLE], OI_BROWNSVILLE_FIRST_ROW)
    write_book(os.path.join(out, INV_FILE), 'Inventory', INV_HEADER,
               [{'ORGANIZATION_NAME': r.org, 'ORGANIZATION_CUSTOMER_NUMBER': r.cust,
                 'PRODUCT_TYPE': TYPE_NAME.get(r.ptype, r.ptype), 'MODEL_GROUP': r.group, 'MODEL_NUMBER': r.model,
                 'INVENTORY_DATE': r.date, 'TOTAL': r.total, 'CITY': r.city} for r in INV_2026])
    by_id = {r.id: r for r in OI_2026 if r is not None}
    write_book(os.path.join(out, OO_FILE), 'Open Order', OO_HEADER,
               [dict(oi_record(by_id[i]), SALES_ORDER_NUMBER=700001 + k, ORGANIZATION_CUSTOMER_NUMBER=900201 + k)
                for k, i in enumerate(OO_IDS)])
    write_book(os.path.join(out, SC_FILE), 'Size Class', SC_HEADER,
               [dict(zip(SC_HEADER, row)) for row in SIZE_CLASS_ROWS])
    write_setup_file(os.path.join(out, SC_SETUP_FILE))
    report = openpyxl.Workbook()
    report.active.title = 'File Format'
    report.active['A1'] = 'old report - the builder must ignore this file'
    report.save(os.path.join(out, '001. Open Orders - Order Intake - Inventory Analysis By MSA_2025.xlsx'))
    open(os.path.join(out, '~$Order Intake 2026.xlsx'), 'wb').close()


def write_expected(out, expected):
    with open(os.path.join(out, 'expected.json'), 'w', encoding='utf-8') as fh:
        json.dump(expected, fh, indent=2)
        fh.write('\n')


# ------------------------------------------------------------------ summary and self-check
def print_summary(expected):
    for run in expected['runs']:
        print(run['options_line'])
        print('  Order Intake: %d rows, %d hidden, visible total %d; month %s; months %s' % (
            len(run['oi_rows']), sum(r[5] for r in run['oi_rows']), run['oi_visible_total'], run['month_cell'],
            run['months']))
        for s in run['summary']:
            print('    %-40s CTL %d  MX %d  SSL %d' % tuple(s))
        print('  Inventory detail: %d rows, %d hidden; %d Run Log lines' % (
            len(run['detail_rows']), sum(r[5] for r in run['detail_rows']), len(run['log_lines'])))
        print('  Open Orders by MSA (chart: %d series, %s):' % (run['open_orders_chart']['series'],
                                                                ', '.join(run['open_orders_chart']['categories'])))
        for s in run['open_orders']:
            print('    %-40s CTL %d  MX %d  SSL %d  Total %d' % tuple(s))


def check(label, actual, expected):
    assert actual == expected, '%s: computed %r, hand-computed %r' % (label, actual, expected)


# Worked out by hand from the tables (class, then status, then newest date with undated last, MSA, export order);
# v2.7 (size-class-file-only plan Section 7.1): S590 is Med Frame (after the MX classes), E25 is UNMAPPED and
# sorts last after T999 (ORDERED before Not yet processed).
HAND_OI_ORDER = ['A11', 'A03', 'A18', 'X02', 'A01', 'A10', 'B01', 'A17', 'A15', 'C07', 'A09', 'C01', 'C03',
                 'C02', 'A02', 'A19', 'A12', 'A13', 'A16', 'C05', 'A20', 'A04', 'A08', 'C06', 'C04', 'A14',
                 'A06', 'A07', 'A05']
# v2.8, run "setup" (Size Class Set-up.xlsx): >00-01 is not allowed, so E10 (A13) is UNMAPPED and sorts first of the
# UNMAPPED rows (ORDERED, the newest); E88 (A14) is >10-15, the last MX class; the rest as HAND_OI_ORDER.
HAND_OI_ORDER_SETUP = ['A11', 'A03', 'A18', 'X02', 'A01', 'A10', 'B01', 'A17', 'A15', 'C07', 'A09', 'C01', 'C03',
                       'C02', 'A02', 'A19', 'A12', 'A16', 'C05', 'A20', 'A04', 'A08', 'C06', 'C04', 'A14',
                       'A06', 'A13', 'A07', 'A05']
HAND_DETAIL_ORDER = ['I02', 'I04', 'I08', 'I06', 'I01', 'I05', 'I10', 'I07', 'I03', 'I09',
                     'D01', 'D06', 'D03', 'D02', 'D05', 'D04']


def self_check(expected):
    runs = {r['name']: r for r in expected['runs']}
    d, v, u = runs['default'], runs['variant'], runs['setup']
    check('cities', d['cities'], [NOC, AM, BR, DF])
    check('months', d['months'], ['July 2026', 'August 2026', 'September 2026'])
    check('order intake order', d['oi_ids'], HAND_OI_ORDER)
    check('order intake order (variant)', v['oi_ids'], HAND_OI_ORDER)
    # v2.7: the same six rows as v2.6; A06 (S590, now Med Frame) sorts after the MX classes, so it is listed last
    check('default hidden rows', [i for i, r in zip(d['oi_ids'], d['oi_rows']) if r[5]], ['A03', 'A09', 'A19', 'C05', 'A08', 'A06'])
    check('default hidden rows as a set (unchanged)', sorted(i for i, r in zip(d['oi_ids'], d['oi_rows']) if r[5]),
          sorted(['A06', 'A03', 'A09', 'A19', 'C05', 'A08']))
    check('variant hidden rows', sum(r[5] for r in v['oi_rows']), 0)
    check('classes via the size class file', [r[0] for r in d['oi_rows']][:3], ['0<2000', '2000<2300', '2300<2600'])
    check('T740 by first word, E25 and T999 unmapped (not in the file), S590 a frame class',
          {r[2]: r[0] for r in d['oi_rows'] if r[2] in ('T740', 'E25', 'T999', 'S590')},
          {'T740': '3000<3500', 'E25': 'UNMAPPED', 'T999': 'UNMAPPED', 'S590': 'Med Frame'})
    check('the file lines (v2.8: no Set-up tab, the 17 built-in classes; v2.7: Over10 not a class)', d['log_lines'][3:8],
          ['Size class file %s: 20 usable rows, 3 warnings' % SC_FILE,
           'Size classes: the 17 built into the macro (Size Class.xlsx has no Set-up tab)'] + SIZE_CLASS_WARNINGS)
    check('built-in classes', [c for c, _ in BUILTIN_CLASSES], [
        '0<2000', '2000<2300', '2300<2600', '2600<3000', '3000<3500', '3500&Over', '>00-01', '>01-02', '>02-03', '>03-04',
        '>04-05', '>05-06', '>06-10', 'EX Small', 'Small Frame', 'Med Frame', 'Large Frame'])
    check('no tonnes line', [l for l in d['log_lines'] + v['log_lines'] if 'tonnes' in l], [])
    check('order by class (Section 7.1)', [(i, r[0]) for i, r in zip(d['oi_ids'], d['oi_rows'])], [
        ('A11', '0<2000'), ('A03', '2000<2300'), ('A18', '2300<2600'), ('X02', '2300<2600'), ('A01', '2300<2600'),
        ('A10', '2300<2600'), ('B01', '2300<2600'), ('A17', '2300<2600'), ('A15', '2600<3000'), ('C07', '2600<3000'),
        ('A09', '2600<3000'), ('C01', '3000<3500'), ('C03', '3000<3500'), ('C02', '3000<3500'), ('A02', '3000<3500'),
        ('A19', '3000<3500'), ('A12', '3500&Over'), ('A13', '>00-01'), ('A16', '>01-02'), ('C05', '>02-03'),
        ('A20', '>03-04'), ('A04', '>03-04'), ('A08', '>04-05'), ('C06', '>05-06'), ('C04', '>05-06'),
        ('A14', '>06-10'), ('A06', 'Med Frame'), ('A07', 'UNMAPPED'), ('A05', 'UNMAPPED')])
    check('city spellings', sorted({r[1] for r in d['oi_rows']}), sorted([NOC, AM, BR, DF]))
    check('TBD row', [r[4] for r in d['oi_rows'] if r[2] == 'T740' and r[1] == DF], ['2026-08-12', 'TBD', '2026-08-05'])
    check('default month', d['month_cell'], 'September 2026')
    check('default summary', d['summary'], [[NOC, 0, 0, 0], [AM, 6, 8, 4], [BR, 0, 0, 0], [DF, 0, 0, 0]])
    check('variant month', v['month_cell'], 'August 2026')
    check('variant summary', v['summary'], [[NOC, 0, 0, 0], [AM, 6, 2, 1], [BR, 0, 0, 0], [DF, 5, 2, 1]])
    check('detail order', d['detail_ids'], HAND_DETAIL_ORDER)
    check('default visible detail', [i for i, r in zip(d['detail_ids'], d['detail_rows']) if not r[5]],
          ['I02', 'I04', 'I08', 'I01', 'I05', 'I10'])
    check('variant visible detail', [i for i, r in zip(v['detail_ids'], v['detail_rows']) if not r[5]],
          ['I06', 'I07', 'I09', 'D01', 'D06', 'D02', 'D04'])
    # v2.5: each MSA has three lines (Order Intake, inventory, open orders), so the MSA lines start at index 8 (v2.8:
    # the size classes line follows the size class file line)
    check('default Amarillo line', d['log_lines'][11],
          'Amarillo: Order Intake 20 rows (CTL 11, MX 7, SSL 2; 4 CANCELLED), 5 hidden by the filters; '
          'UNMAPPED model groups: E25, T999')
    check('default Amarillo inventory line', d['log_lines'][12], 'Amarillo: Inventory 10 rows; September 2026: CTL 6, MX 8, SSL 4 units')
    check('default Dallas line', d['log_lines'][17],
          'Dallas-Fort Worth [North] Metroplex: Order Intake 7 rows (CTL 4, MX 3, SSL 0; 1 CANCELLED), 1 hidden by the filters')
    check('variant Dallas inventory line', v['log_lines'][18],
          'Dallas-Fort Worth [North] Metroplex: Inventory 6 rows; August 2026: CTL 5, MX 2, SSL 1 units')
    check('variant Amarillo line', v['log_lines'][11],
          'Amarillo: Order Intake 20 rows (CTL 11, MX 7, SSL 2; 4 CANCELLED), 0 hidden by the filters; '
          'UNMAPPED model groups: E25, T999')
    check('number of log lines', (len(d['log_lines']), len(v['log_lines']), len(u['log_lines'])), (24, 24, 27))
    check('TBD worksheet row', OI_2026_FIRST_ROW + [r for r in OI_2026].index(
        next(r for r in OI_2026 if r is not None and r.id == 'C03')), 25)
    # Open Orders by MSA (plan open-orders-by-msa-tab Section 7.1, hand-computed from the Order Intake table)
    hand_table = [[NOC, 1, 0, 0, 1], [AM, 8, 5, 1, 14], [BR, 1, 0, 0, 1], [DF, 4, 1, 0, 5], ['Total', 14, 6, 1, 21]]
    check('open orders table', d['open_orders'], hand_table)
    check('open orders table (variant, same)', v['open_orders'], hand_table)
    check('open rows per MSA (ORDERED, Not yet processed)', d['open_orders_split'],
          [[NOC, 1, 0], [AM, 12, 1], [BR, 1, 0], [DF, 4, 1]])
    check('open rows: 20 = 18 ORDERED + 2 Not yet processed',
          (sum(s[1] + s[2] for s in d['open_orders_split']), sum(s[1] for s in d['open_orders_split']),
           sum(s[2] for s in d['open_orders_split'])), (20, 18, 2))
    check('chart default', d['open_orders_chart'], {'series': 4, 'categories': ['CTL', 'MX']})
    check('chart variant', v['open_orders_chart'], {'series': 4, 'categories': ['CTL', 'MX', 'SSL']})
    check('the " mx " row counts as MX', [r.qty for r, c in oi_kept() if r.id == 'A20' and is_open(r.status)], [1])
    check('default (no MSA) open line', d['log_lines'][10],
          '(no MSA): Open orders 1 row (ORDERED 1, Not yet processed 0): CTL 1, MX 0, SSL 0 units')
    check('default Amarillo open line', d['log_lines'][13],
          'Amarillo: Open orders 13 rows (ORDERED 12, Not yet processed 1): CTL 8, MX 5, SSL 1 units')
    check('variant Dallas open line', v['log_lines'][19],
          'Dallas-Fort Worth [North] Metroplex: Open orders 5 rows (ORDERED 4, Not yet processed 1): CTL 4, MX 1, SSL 0 units')
    check('no stray-space, disagreement or zeros line', [l for l in d['log_lines'] if 'Open orders: ' in l or
                                                         'extra spaces' in l], [])
    check('build table item', d['build_open_orders'],
          'ORDERED and Not yet processed: 20 rows, 21 units (CTL 14, MX 6, SSL 1), on the Open Orders by MSA tab')
    # v2.6 visible total: 29 kept rows, every quantity 1 but A10's 2, so 30 units; the default run hides six rows
    # of quantity 1 (A06, A03, A09, A19, C05, A08), the variant none
    check('visible total (default)', d['oi_visible_total'], 24)
    check('visible total (variant)', v['oi_visible_total'], 30)
    # v2.8, run "setup": the classes come from the Set-up tab of Size Class Set-up.xlsx (hand-worked in SETUP_CLASSES)
    check('setup: classes', u['classes'], [list(c) for c in SETUP_CLASSES])
    check('setup: 18 classes, CTL 6, MX 7, SSL 5', [sum(f == k for _, f in SETUP_CLASSES) for k in ('CTL', 'MX', 'SSL')], [6, 7, 5])
    check('setup: order intake order', u['oi_ids'], HAND_OI_ORDER_SETUP)
    check('setup: E10 UNMAPPED (>00-01 not allowed), E88 >10-15 (added)',
          {r[2]: r[0] for r in u['oi_rows'] if r[2] in ('E10', 'E88', 'S590', 'T740')},
          {'E10': 'UNMAPPED', 'E88': '>10-15', 'S590': 'Med Frame', 'T740': '3000<3500'})
    check('setup: hidden rows as the default run', [i for i, r in zip(u['oi_ids'], u['oi_rows']) if r[5]],
          ['A03', 'A09', 'A19', 'C05', 'A08', 'A06'])
    check('setup: file lines', u['log_lines'][3:11], [
        'Size class file Size Class Set-up.xlsx: 19 usable rows, 6 warnings',
        'Size classes: 18 from the Set-up tab (CTL 6, MX 7, SSL 5)',
        'Set-up tab C9: "2000<2300" is listed again (first in A6) - ignored',
        'Set-up tab C10: "unmapped" is not a size class name (UNMAPPED marks models not in the file) - ignored',
        'Size class file row 11 E10: ">00-01" is not a size class - row skipped',
        'Size class file row 21 T630: "2000-2300" is not a size class - row skipped',
        'Size class file row 22 T650: listed again (first in row 20) - row 20 is used',
        'Size class file row 24 E100: "Over10" is not a size class - row skipped'])
    check('setup: Amarillo line (UNMAPPED in export order)', u['log_lines'][14],
          'Amarillo: Order Intake 20 rows (CTL 11, MX 7, SSL 2; 4 CANCELLED), 5 hidden by the filters; '
          'UNMAPPED model groups: E25, T999, E10')
    check('setup: the rest of the Run Log as the default run', u['log_lines'][11:], d['log_lines'][8:11] + [u['log_lines'][14]]
          + d['log_lines'][12:])
    check('setup: figures as the default run', (u['summary'], u['open_orders'], u['oi_visible_total'], u['detail_ids']),
          (d['summary'], d['open_orders'], d['oi_visible_total'], d['detail_ids']))


def main(argv):
    out = os.path.abspath(argv[1]) if len(argv) > 1 else DEFAULT_OUT
    expected = build_expected()
    write_fixtures(out)
    write_expected(out, expected)
    print('Wrote fixtures and expected.json to %s' % out)
    print_summary(expected)
    self_check(expected)
    print('Self-check passed.')


if __name__ == '__main__':
    main(sys.argv)
