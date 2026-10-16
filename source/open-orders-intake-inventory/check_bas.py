#!/usr/bin/env python3
"""Static checker for the open-orders VBA module (no Excel needed).

Usage: python3 check_bas.py [path/to/module.bas]   (default: build-open-orders-report.bas
beside this script). Prints PASS or FAIL per check (a FAIL lists at most 10 line numbers),
then RESULT: PASS or RESULT: FAIL (<k> checks failed). Exit 0 if all pass, 1 if any fails,
2 if the file cannot be read.

Checks: bytes (CRLF on every line, file ends in CRLF, ASCII only); header (Attribute
VB_Name on line 1, Option Explicit before the first Sub/Function); line-length (no line
over 1,000 characters); continuations (at most 20 per logical line, at most 50 Array(
arguments, none dangling at EOF); blocks (Sub, Function, Property, Type, With, For, Do,
While, Select Case, If, #If nest and close in order; a one-line If opens nothing);
reserved-names (Dim, Static, Const, module-level and parameter names, case-insensitive);
entry-points-and-forbidden (ShowOpenOrdersReportConfig declared once as Public Sub with no
arguments and no other such Sub, so it is the only macro in Excel's Macros list; the button
macros declared once as Public Sub with only Optional arguments; no duplicate procedure names, though Property Get/Let/Set may pair; no forbidden API or
option text outside comments).
"""
import os
import re
import sys

RESERVED = {w.lower() for w in """
    Abs CBool CByte CCur CDate CDbl CDec CInt CLng CLngLng CLngPtr CSng CStr CVar CVErr
    Date Debug DoEvents Fix Int Len LenB Me PSet Scale Sgn String
    Call Case Close Const Dim Do Each Else End Erase Error Exit For Function Get GoTo If
    Input Let Lock Loop Name Next On Open Option Print Put ReDim Resume Return Seek Select
    Set Static Step Stop Sub Then To Type Unlock Wend While With Write""".split()}
ENTRY_POINTS = ['ShowOpenOrdersReportConfig']           # the only macros in the Macros list (Alt+F8)
BUTTON_MACROS = ['BuildOpenOrdersReport', 'CheckOpenOrdersReportSetup',
                 'ChooseOpenOrdersInputFolder', 'ChooseOpenOrdersOutputFolder',
                 'ChooseOpenOrdersSizeClassFile']   # Config buttons only
FORBIDDEN = [(r'Scripting\.Dictionary', 'Scripting.Dictionary'), (r'\bCreateObject\s*\(', 'CreateObject('),
             (r'\bGetObject\s*\(', 'GetObject('), (r'\bDeclare\s', 'Declare '), (r'\bPtrSafe\b', 'PtrSafe'),
             (r'\.VBProject\b', '.VBProject'), (r'\bOption\s+Base\b', 'Option Base'),
             (r'\bOption\s+Compare\s+Text\b', 'Option Compare Text')]
# Block rules: (regex, action, block name); a None name means group 1 of the regex, titled.
RULES = [
    (r'(?:(?:Public|Private|Friend|Static)\s+)*(Sub|Function|Property)\b', 'open', None),
    (r'End\s+(Sub|Function|Property|Type|With|Select|If)\b', 'close', None),
    (r'(?:(?:Public|Private)\s+)*Type\b', 'open', 'Type'), (r'With\b', 'open', 'With'),
    (r'For\b', 'open', 'For'), (r'Next\b', 'close', 'For'), (r'Do\b', 'open', 'Do'), (r'Loop\b', 'close', 'Do'),
    (r'While\b', 'open', 'While'), (r'Wend\b', 'close', 'While'), (r'Select\s+Case\b', 'open', 'Select'),
    (r'If\b', 'open', 'If'), (r'Else(?:If)?\b', 'mid', 'If'),
    (r'#If\b', 'open', '#If'), (r'#End\s+If\b', 'close', '#If'), (r'#Else(?:If)?\b', 'mid', '#If'),
]
STRING_OR_COMMENT = re.compile(r'"((?:[^"]|"")*)"|\'.*$')
PROC = re.compile(r'(?:(?:Public|Private|Friend|Static)\s+)*(Sub|Function|Property\s+(Get|Let|Set))\s+(\w+)', re.I)
DECL = re.compile(r'((?:(?:Public|Private|Static|Global|Dim|Const)\s+)+)'
                  r'(?!(?:Sub|Function|Property|Type|Enum|Declare|Event)\b)(.*)', re.I)

def code_part(line, keep=False):
    """The line without its comment. String contents are blanked (quotes kept) unless keep."""
    def repl(m):
        if m.group(0).startswith("'"):
            return ''
        return m.group(0) if keep else '"' + ' ' * len(m.group(1)) + '"'
    return STRING_OR_COMMENT.sub(repl, line)

def logical(lines):
    """(first line, code, continuation count) per logical line; comments removed, continuations joined."""
    out, cur, start = [], [], 0
    for i, raw in enumerate(lines + [''], 1):  # the extra '' closes a dangling continuation
        code = code_part(raw).rstrip()
        cont = code.endswith(' _')
        start = start if cur else i
        cur.append((code[:-2] if cont else code).strip())
        if not cont:
            out.append((start, ' '.join(cur), len(cur) - 1))
            cur = []
    return out

def split_top(text):
    """Items of a comma list, split at commas outside parentheses; stops at an unmatched ')'."""
    items, cur, depth = [], '', 0
    for c in text:
        depth += (c == '(') - (c == ')')
        if depth < 0:
            break
        if c == ',' and depth == 0:
            items.append(cur)
            cur = ''
        else:
            cur += c
    return items + [cur]

def item_name(item):
    """Declared name in a Dim, Const or parameter item, after modifiers such as ByRef or Optional."""
    m = re.match(r'\s*(?:(?:ByVal|ByRef|Optional|ParamArray|WithEvents)\s+)*(\w+)', item, re.I)
    return m.group(1) if m else None

def one_line_if(s):
    """True for 'If c Then <code>': an If with code after Then on the same line."""
    parts = re.split(r'\bThen\b', s, maxsplit=1, flags=re.I)
    return bool(re.match(r'If\b', s, re.I)) and len(parts) > 1 and parts[1].strip(' \t:') != ''

def statements(code):
    """Split a logical line on ':'. A one-line If takes the rest of the line as one statement."""
    parts = [p.strip() for p in code.split(':')]
    for k in range(len(parts)):
        if one_line_if(':'.join(parts[k:])):
            return parts[:k] + [':'.join(parts[k:])]
    return parts

def classify(s):
    """(action, block name) for one statement: 'open', 'close', 'mid' (Else) or None."""
    for rx, action, name in RULES:
        m = re.match(rx, s, re.I)
        if m:
            return action, name or m.group(1).title()
    return None, None

def array_counts(code):
    """Top-level argument count of each Array( call in a logical line."""
    counts = []
    for m in re.finditer(r'\bArray\s*\(', code, re.I):
        items = split_top(code[m.end():])
        counts.append(len(items) if any(x.strip() for x in items) else 0)
    return counts

def parse(lns):
    """Procedures [(line, kind, name)] and declared names [(line, name)] for variables,
    constants and parameters."""
    procs, names = [], []
    for ln, code, _ in lns:
        for s in statements(code):
            m, d = PROC.match(s), DECL.match(s)
            if m:
                procs.append((ln, (m.group(2) or m.group(1)).title(), m.group(3)))
                rest = s[m.end():].lstrip()
                if rest.startswith('('):  # parameter list
                    names += [(ln, n) for n in map(item_name, split_top(rest[1:])) if n]
            elif d:
                names += [(ln, n) for n in map(item_name, split_top(d.group(2))) if n]
    return procs, names

def check_bytes(data, raw):
    probs = [] if data.endswith(b'\r\n') else ['file does not end with CRLF']
    for i, r in enumerate(raw, 1):
        if i < len(raw) and not r.endswith('\r'):
            probs.append(f'line {i}: bare LF')
        if '\r' in (r[:-1] if r.endswith('\r') else r):
            probs.append(f'line {i}: bare CR')
        if re.search(r'[^\x00-\x7f]', r):
            probs.append(f'line {i}: non-ASCII byte')
    return probs

def check_header(lines, lns):
    probs = []
    if not re.match(r'Attribute VB_Name = "[^"]+"\s*$', lines[0] if lines else ''):
        probs.append('line 1 is not Attribute VB_Name = "<name>"')
    subs = [ln for ln, code, _ in lns if re.match(r'(?:(?:Public|Private|Friend|Static)\s+)*(Sub|Function)\b', code, re.I)]
    opts = [ln for ln, code, _ in lns if re.match(r'Option\s+Explicit$', code, re.I)]
    if not opts:
        probs.append('no Option Explicit line')
    elif subs and opts[0] > subs[0]:
        probs.append(f'Option Explicit (line {opts[0]}) is after the first Sub/Function (line {subs[0]})')
    return probs

def check_continuations(lines, lns):
    probs = []
    if lines and code_part(lines[-1]).rstrip().endswith(' _'):
        probs.append(f'line {len(lines)}: continuation at end of file')
    for ln, code, conts in lns:
        if conts > 20:
            probs.append(f'line {ln}: {conts} continuations')
        probs += [f'line {ln}: Array( with {n} arguments' for n in array_counts(code) if n > 50]
    return probs

def check_blocks(lns):
    stack = []  # open blocks as (name, line)
    for ln, code, _ in lns:
        for s in statements(code):
            if not s or one_line_if(s):
                continue
            action, name = classify(s)
            if action == 'open':
                stack.append((name, ln))
            elif action == 'mid' and (not stack or stack[-1][0] != name):
                return [f'line {ln}: "{s[:30]}" is outside an {name} block']
            elif action == 'close':
                if not stack:
                    return [f'line {ln}: "{s[:30]}" has no open block']
                top, at = stack.pop()
                if top != name:
                    return [f'line {ln}: "{s[:30]}" does not match open {top} (line {at})']
    return [f'end of file: {top} (line {at}) is never closed' for top, at in stack]

def check_entries(lines, lns, procs):
    probs = []
    for ep in ENTRY_POINTS:
        hits = [ln for ln, code, _ in lns if re.match(r'Public\s+Sub\s+' + ep + r'\s*\(\s*\)$', code, re.I)]
        if len(hits) != 1:
            probs.append(f'Public Sub {ep}() declared {len(hits)} times, expected 1 (lines {hits})')
    listed = [(ln, m.group(1)) for ln, code, _ in lns
              for m in [re.match(r'(?:Public\s+)?Sub\s+(\w+)\s*\(\s*\)$', code, re.I)] if m]
    extra = [f'{name} (line {ln})' for ln, name in listed if name not in ENTRY_POINTS]
    if extra:
        probs.append(f'only {ENTRY_POINTS} may be a Public Sub without arguments (shown by Alt+F8): {extra}')
    for bm in BUTTON_MACROS:
        hits = [ln for ln, code, _ in lns
                if re.match(r'Public\s+Sub\s+' + bm + r'\s*\(\s*Optional\b[^)]*\)$', code, re.I)]
        if len(hits) != 1:
            probs.append(f'Public Sub {bm}(Optional ...) declared {len(hits)} times, expected 1 (lines {hits})')
    groups = {}
    for ln, kind, name in procs:
        groups.setdefault(name.lower(), []).append((ln, kind, name))
    for items in groups.values():
        kinds = [kind for _, kind, _ in items]
        if len(items) > 1 and ('Sub' in kinds or 'Function' in kinds or len(set(kinds)) < len(kinds)):
            probs.append(f'{items[0][2]} declared more than once (lines {[ln for ln, _, _ in items]})')
    for i, raw in enumerate(lines, 1):
        code = code_part(raw, keep=True)
        probs += [f'line {i}: forbidden {text}' for pat, text in FORBIDDEN if re.search(pat, code, re.I)]
    return probs

def main(argv):
    default = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'build-open-orders-report.bas')
    path = argv[1] if len(argv) > 1 else default
    try:
        with open(path, 'rb') as f:
            data = f.read()
    except OSError as err:
        print(f'cannot read {path}: {err}', file=sys.stderr)
        return 2
    raw = data.decode('latin-1').split('\n')  # line number = index + 1
    lines = [r[:-1] if r.endswith('\r') else r for r in raw]
    if lines[-1] == '':
        lines.pop()
    lns = logical(lines)
    procs, names = parse(lns)
    checks = [
        ('bytes', check_bytes(data, raw)),
        ('header', check_header(lines, lns)),
        ('line-length', [f'line {i}: {len(s)} characters' for i, s in enumerate(lines, 1) if len(s) > 1000]),
        ('continuations', check_continuations(lines, lns)),
        ('blocks', check_blocks(lns)),
        ('reserved-names', [f"line {ln}: '{n}' is a reserved word" for ln, n in names if n.lower() in RESERVED]),
        ('entry-points-and-forbidden', check_entries(lines, lns, procs)),
    ]
    failed = 0
    for name, probs in checks:
        if probs:
            failed += 1
            more = f' (+{len(probs) - 10} more)' if len(probs) > 10 else ''
            print(f'FAIL {name}: ' + '; '.join(probs[:10]) + more)
        else:
            print(f'PASS {name}')
    print(f'RESULT: FAIL ({failed} checks failed)' if failed else 'RESULT: PASS')
    return 1 if failed else 0

if __name__ == '__main__':
    sys.exit(main(sys.argv))
