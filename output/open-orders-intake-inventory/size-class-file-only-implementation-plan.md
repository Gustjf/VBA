# Open Orders Macro: Size Class from the Master List Only, Implementation Plan for Claude Opus (high effort)

This file is an **architect's brief**. It was written by one AI session (2026-10-15) for a later session, Claude Opus run at high reasoning effort (the "implementer"), that changes how the Open Orders report macro assigns a size class to an order line. The owner runs the implementer with the prompt at the very end (Section 13).

The request, in the owner's words (2026-10-15, after a Check setup on the Austin sample with the new 29-model master list): *I am getting these errors. I expected most of these errors, except for the "tonnes rule" - I don't think this is intended behavior. The size class master sheet must be the only input now for assigning a size class. The previous behavior for assigning size classes was a workaround because we did not have the size class sheet created. There are three different categories of equipment at Bobcat we are looking at: skid steer loader (SSL), compact track loader (CTL) and excavator (MX). The SSLs use the size classes EX Small, Small Frame, Med Frame, Large Frame; the CTL uses <2000, 2000<2300, etc.; the MX uses >00-01, >01-02. This is just how the business organically grew and where we are at today.*

What the owner saw (the Last run box, "Setup check: ready to build (10 warnings)"): nine `WARNING: Size class file row N S...: "Med Frame" is not a size class - row skipped` lines, one per skid-steer row of the master list, and one `WARNING: Models in Order Intake that are not in the size class file: S86-2 (UNMAPPED until added), ..., E60 (tonnes rule), E55 (tonnes rule), E88 (tonnes rule)`. The nine are the macro refusing the SSL frame classes; the three "tonnes rule" items are the v1 fallback placing E-models by their model number. Both go away with this change: the frame classes become classes, and a model not in the file is UNMAPPED whatever its name.

Everything below turns that into a contract the macro, the Python reference, the checker, the fixtures, the LibreOffice run, the master-list scripts and the documents all follow.

---

## 1. Mission and deliverables

Macro **v2.7** reads a model's size class from the size class workbook and from nowhere else. The accepted classes are the seventeen the business uses (six CTL, seven MX, four SSL; Section 5.1). The tonnes rule, the `Over10` class and the rated-operating-capacity cross-check are removed. The File Format tab's Table 1 shows the three families. Nothing else about the report changes.

| Deliverable | Path | Notes |
|---|---|---|
| Macro v2.7 | `source/open-orders-intake-inventory/build-open-orders-report.bas` | Section 5. CRLF, ASCII; `check_bas.py` passes. |
| Reference v2.7 | `source/open-orders-intake-inventory/build_report_reference.py` | Same rules, same Table 1, same log lines, `FF_*` rows moved. |
| Checker v2.7 | `source/open-orders-intake-inventory/check_report.py` | The 17 classes, Table 1 with 18 rows, no `Over10` (Section 7.3). |
| Fixtures | `source/open-orders-intake-inventory/make_test_fixtures.py`, `test-fixtures/Size Class.xlsx`, `test-fixtures/expected.json` | Section 7.1; the hand-computed order asserted in `self_check`. |
| Master-list scripts and outputs | `make_size_class_master.py`, `make_size_class_pdf.py`, `size-class-master.tex`; `output/.../size-class-master.xlsx`, `size-class-master.pdf` | Section 8.1: the SSL rows and E40 lose their REVIEW notes; the PDF's rule paragraphs rewritten. |
| POC report | `output/open-orders-intake-inventory/001. Open Orders - Order Intake - Inventory Analysis By MSA_POC.xlsx` | Rebuilt by the reference from the Austin sample with the regenerated master list. |
| Documents | `output/.../macro-instructions.md`, `excel-test-protocol.md`, `copilot-prompts.md` (+ `make_size_class_copilot_pdf.py`, `size-class-copilot-prompt.tex`, `.pdf`), `source/.../open-orders-macro-manual.tex` + `output/.../open-orders-macro-manual.pdf` (Rev. 11), `open-orders-macro-flow.tex` and `-alt.tex` + their PDFs (Rev. 15), `open-orders-team-review.pdf` (sheet 3 is the flow drawing) | Section 8. |
| Map rows | `MAP.md` | Every edited file's row says v2.7 where it names the version and no longer mentions the tonnes rule, Over10, the ROC check, frame-class warnings, "11 items" or "13 REVIEW rows"; this plan's row marked "Executed". Same commit. |

Not deliverables: no new Config setting, no change to the Order Intake columns, filters, visible total, the Inventory by Month or Open Orders by MSA tabs, the status order or any existing Run Log line other than the ones Section 5 names. No alias spellings for classes (Section 11, question 1). The historic documents (`size-class-reference.md`, the `stronger-ai-*` plans, `open-orders-report-questions.*`, `macro-flow-visual-implementation-plan.md`, the v2.5 tab plan) describe earlier revisions and stay as they are; say so in the final report.

---

## 2. Standing rules

1. Read the repo rules file (`CLAUDE.md`) and `MAP.md` first and obey them: push directly to `main` (a cloud session also pushes the same commit to its `claude/<name>` branch), no pull request, update `MAP.md` in the same commit, source under `source/open-orders-intake-inventory/`, output under `output/open-orders-intake-inventory/`, the one-page manual rebuilt with `./build.sh` in the same commit as any user-facing macro change.
2. **One rule, three places.** Every rule in Section 5 is implemented in the macro, in `build_report_reference.py` and in `check_report.py`; the fixtures assert it. A rule you cannot implement in one of the three is not implemented: stop and say so in the report.
3. **Delete, do not disable.** The tonnes rule, `Over10` and the ROC cross-check are removed from the code, not commented out or left behind a flag. Dead helpers (`TonnesClass`, `LoaderClass`, `RocValue`, `tonnes_class`, `loader_class`, `roc_value`, the `LOADER_RULE`/`TONNES_RULE` tables and their text constants) go with them, unless another live caller remains (check before deleting; `make_size_class_master.py` imports `ref.loader_class` today and must stop).
4. **Delegation (Section 6.1).** Smaller AIs do bounded, verifiable work from self-contained briefs: Sonnet for an independent recomputation, the VBA diff review and the document edits; Haiku for sweeps and counts. A worker never changes the contract, and you verify every worker's output yourself before it goes in a commit.
5. **Every change is validated before it is committed**: `check_bas.py` on the module, the reference on the fixtures and the Austin sample, `check_report.py --expected` on every fixture build, the macro in LibreOffice compared with the reference (`--compare`), `compare_with_handbuilt.py` on the Austin build (Section 7.2 says what it must show). What you could not run (real Excel), you say.
6. Report faithfully at the end: what passed, what was not run, what you decided that this brief left open, and what the owner still owes (Section 11).

---

## 3. Read first, then stop reading

| Read | Why | How much |
|---|---|---|
| `CLAUDE.md`, `MAP.md` | Repo rules; where things are. | All (short). |
| This brief | The contract. | All. |
| `build-open-orders-report.bas` | Header comment (1 to 70: the rule bullet at 43 to 44). `CheckOpenOrdersSetup` 556 to 567 (the "Models in Order Intake that are not in the size class file" warning and its tonnes suffix). `RunLogLineToCheck` 594 to 601. `ConfigHelp` 805 to 806. **The size class file section 1107 to 1314**: `LoadSizeClassFile` (ROC column 1130, ROC cross-check 1172 to 1178), `RocValue`, `LookupClass`, `LoaderClass`, `TonnesClass`, `SizeClasses`, `CanonClass`, `ClassRank`. `PrepareData` 1602 to 1688 (the assignment at 1666 to 1679, `mTonnes` at 192, 1610, 1672). `BuildLogLines` 2078 (the tonnes line). `WriteFileFormat` 2125 to 2204 (Table 1 at 2130 to 2144; `FF_*` constants at 93 to 96). `StyleTable` and `Band` 2627 to 2655 (the `unbanded` argument). | Those places. |
| `build_report_reference.py` | Docstring 30; constants 68 to 76 and 101 (`FF_*`); `canon_class`, `loader_class`, `tonnes_class` 193 to 202; `load_size_class_file` 321 to 369; the assignment 504 to 516 and `self.tonnes` 467 to 468; the log line 835 to 836; Table 1 at 864 to 871. | Those. |
| `check_report.py` | `CLASSES`, `OVER10`, `CLASS_RANK`, `UNKNOWN_RANK` 110 to 115; the captions 134 to 137; `class_rank` 202; `fill_cells` for `ff` 355 to 364; `check_file_format` Table 1 at 875 to 889; R3a/R3b at 975 to 980. | Those. |
| `make_test_fixtures.py` | `CLASS_RANK` 36 to 39; the Order Intake table 73 to 107; `SIZE_CLASS_ROWS` and `SIZE_CLASS_WARNINGS` 138 to 152; the log 300 to 345; `HAND_OI_ORDER` and `self_check` 470 to 530. | Those. |
| `make_size_class_master.py`, `make_size_class_pdf.py` | All (short); the rule paragraphs of the PDF script at 105 to 120. | All. |
| `run_macro_libreoffice.py` | The docstring table of substitutions only; no substitution changes. `--check-only` prints the Last run box, which Section 7.1 and 7.2 need. | That. |
| `compare_with_handbuilt.py` | Lines 51 to 57 (the two checks). | Those. |
| `macro-instructions.md`, `excel-test-protocol.md`, `copilot-prompts.md` | Only the lines Section 8 lists. | Those lines. |
| `open-orders-macro-manual.tex`, `open-orders-macro-flow.tex` (line 136 and the title block), `open-orders-macro-flow-alt.tex` (line 129), `build.sh` | Section 8.3 and 8.4. | Those places. |

Do not open the sample or fixture workbooks cell by cell; use scripts. Do not read the other plans in `output/` beyond Section 11 of `stronger-ai-implementation-plan.md` (the traps list, carried into Section 12 here).

---

## 4. The request in the report's terms

Today (v2.6) a model group gets its class in three steps: the size class file (exact name, then first word), else the tonnes rule for an MX model named `E<digits>` (number / 10 = nominal tonnes, upper bound inclusive, `Over10` above 10 t), else `UNMAPPED`. The file reader accepts thirteen class spellings plus `Over10`, and for a model whose name starts with T or S it recomputes the class from the rated operating capacity column and warns when the file disagrees. The owner's real master list (29 models, built from the Bobcat Equipment Size Class sheet on 2026-10-15) classes its nine skid-steer loaders as `EX Small`, `Small Frame`, `Med Frame` and `Large Frame`, which the reader rejects, so those models read UNMAPPED and their rows produce nine warnings; and three Austin mini excavators not on the sheet (E55, E60, E88) were placed by the tonnes rule instead of being flagged.

After this change: **the size class file is the only source of a class.** The reader accepts seventeen spellings (Section 5.1). A model not in the file is UNMAPPED, whatever its product type or name. The file's class is never second-guessed. Table 1 of the File Format tab lists the seventeen classes in three families and says the class comes from the file.

A fact worth knowing for Section 7.2: with the 29-model master list, the Austin sample's Order Intake has eight model groups not on the list: `T64`, `T66`, `T86` (CTL; the list has `T64-2`, `T66-2`, `T86-2`, which the first-word match does not cover because the whole first word differs), `E55`, `E60`, `E88` (MX), `S510`, `S66` (SSL). All eight read UNMAPPED after the change; adding them to the workbook is the owner's job (Section 11, question 4), not the macro's.

---

## 5. The contract, fixed

### 5.1 The size classes

Seventeen classes in this order, which is the sort rank (1 to 17), the order of Table 1 and the order of the `SizeClasses()` / `SIZE_CLASSES` / `CLASSES` / `CLASS_RANK` tables in all four Python and VBA places:

| Rank | Class | Family |
|---|---|---|
| 1 to 6 | `0<2000`, `2000<2300`, `2300<2600`, `2600<3000`, `3000<3500`, `3500&Over` | Track loaders (CTL) |
| 7 to 13 | `>00-01`, `>01-02`, `>02-03`, `>03-04`, `>04-05`, `>05-06`, `>06-10` | Mini excavators (MX) |
| 14 to 17 | `EX Small`, `Small Frame`, `Med Frame`, `Large Frame` | Skid-steer loaders (SSL) |
| 99 | `UNMAPPED` | Any |

The CTL and MX blocks keep their v2.6 order, so every existing CTL and MX row keeps its place; the SSL block comes after MX (the Open Orders by MSA and Units tables already list CTL, MX, SSL in that order; SSL rows are hidden by default, so the default view of Order Intake is unchanged apart from the rows that leave the tonnes rule). `Over10` is no longer a class: a file row carrying it is skipped with the "is not a size class" warning, and nothing in the code produces it. `ClassRank` of an unknown class stays 99.

The canonical spelling is the one in the table (`EX Small` with capital E and X, `Med Frame`, not `Medium`). The reader matches ignoring case after `Trim`, as today (`ex small` is accepted and written `EX Small`); it does not accept other spellings (Section 11, question 1).

### 5.2 Reading the size class file

Columns found by name as today: `Name of machine` (aliases unchanged) and `Size Class` (aliases unchanged). The rated-operating-capacity column is **no longer looked for or read**; a file may still have it (the master list does), and it is ignored.

Rows: blank rows skipped silently; then, per row, in this order and with these exact texts (unchanged): `<where>: no machine name - row skipped`; `<where>: no size class - row skipped`; `<where>: "<raw>" is not a size class - row skipped`; `<where>: listed again (first in row N) - row N is used`. The ROC cross-check warning (`... ROC n is class X by the rule, file says Y`) is removed. `<where>` = `Size class file row <n>` plus a space and the name when there is one, as today. The "no usable rows" problem text is unchanged. The Run Log line `Size class file <name>: <n> usable rows, <m> warnings` is unchanged.

Lookup is unchanged: the model group trimmed and upper-cased, equal to a name, else equal to a name's first word; first row of a name wins.

### 5.3 Assigning a class to an Order Intake row

`cls = LookupClass(uMg)`; if empty, `cls = "UNMAPPED"` and the model group (or `(blank)`) is added to the MSA's list and the all-MSAs list, exactly as the UNMAPPED branch does today, **for every product type**. The `E#*` branch, `TonnesClass`, `mTonnes`, the Run Log line `Order Intake: placed by the tonnes rule (not in the size class file): ...` and the `(tonnes rule)` suffix of the setup-check warning are removed. The `RunLogLineToCheck` patterns `placed by the tonnes rule` and `not in the size class file` are removed (nothing produces them on the Run Log any more; the UNMAPPED line is caught by `UNMAPPED`).

No family check: the macro does not compare a class's family with `PRODUCT_TYPE_ABBR` (an SSL model given a CTL class by the file is written as the file says). Section 11, question 3.

### 5.4 The setup-check warning

`CheckOpenOrdersSetup` (Check setup only, as today) warns once when kept Order Intake rows have model groups that are not in the file, with the new wording:

`Models in Order Intake that are not in the size class file (UNMAPPED until added): <g1>, <g2>, ...`

in the order first seen across the files (the existing `mUnmappedAll` order). The per-item `(UNMAPPED until added)` suffix goes; the parenthesis moves into the label. Same in the reference's `--check-only` equivalent if it has one (it does not today: the harness prints the macro's own box).

### 5.5 File Format, Table 1

Caption unchanged (`Table 1.` `Size classes, in report order.`); header `Class`, `Family`, `Rule` unchanged. Eighteen body rows: the seventeen classes of Section 5.1 in order, then `UNMAPPED`. Family and Rule texts:

| Rows | Family | Rule |
|---|---|---|
| CTL (6) | `Track loaders (CTL)` | `by rated operating capacity in lb; the class is read from the size class file` |
| MX (7) | `Mini excavators (MX)` | `by nominal tonnes; the class is read from the size class file` |
| SSL (4) | `Skid-steer loaders (SSL)` | `by frame size; the class is read from the size class file` |
| UNMAPPED | `Any` | `model not in the size class file; add it there and build again` (unchanged) |

With `r = FF_CLASS_ROW = 3`: header `r+1`; CTL `r+2` to `r+7`; MX `r+8` to `r+14`; SSL `r+15` to `r+18`; UNMAPPED `r+19`. A medium rule above `r+8` and above `r+15` (the two family boundaries). The UNMAPPED row bold and peach (`CLR_FLAG` interior) as today. **It is now the 18th body row, which the banding formula would fill green and hide the peach**, so the table is styled with its last row unbanded (`StyleTable ... unbanded:=1`, the way the Total row of the Units table is; the reference's `ruled_table(..., unbanded=1)` at line 606 is the same thing, and its `groups` argument draws the two family rules) and the banding range ends at `r+18`. The constants `LOADER_RULE_TEXT` and `TONNES_RULE_TEXT` are replaced by the three texts above (one constant each, or an array; your choice, same text in the reference).

Table 1 grows by three rows, so the fixed rows move in **both** the macro and the reference: `FF_STATUS_ROW` 21 to **24**, `FF_FILTER_ROW` 30 to **33**, `FF_BUILD_ROW` 39 to **42**. `FF_CLASS_ROW` and `FF_MONTH_COL` stay 3 and 6; the Months list in column F is unaffected. The checker finds tables by caption and needs no row constant, but its Table 1 offsets change (Section 7.3).

### 5.6 Config, help, upgrade

No new setting; no change to the Config layout or the upgrade path (a v2.6 Config sheet is a v2.7 Config sheet; `mUpgradeNote` says v2.7 and keeps its content). The help row `Size Class is UNMAPPED` may say `one of the 17 class names on the File Format tab`; nothing else in `ConfigHelp` changes.

### 5.7 Header comment and version

The header's rule bullet (lines 43 to 44) becomes: `Size Class from MODEL_GROUP via the size class file (exact name, then the name's first word), the only source of a class: 17 classes in three families (CTL by rated operating capacity, MX by nominal tonnes, SSL by frame size). A model not in the file is UNMAPPED (bold, peach).` `2.6` becomes `2.7` everywhere the implementer edits where it names the macro version: the module header and `BUILDER_TEXT`, `VERSION` in the reference, the checker's and runner's docstrings, `'version'` in `make_test_fixtures.py` (and therefore `expected.json`), the instructions, the test protocol, the manual (Rev. 11, dated the build day), the MAP rows of those files. Files not edited keep `2.6`.

---

## 6. Work plan

### 6.1 Where the smaller AIs earn their keep

Use them where the work is bounded, the brief can be self-contained, and the result can be checked against something fixed. Skip them where you would have to explain the whole contract first.

| When | Who | Task | How you verify |
|---|---|---|---|
| Phase 1, before editing `self_check` | Sonnet | Recompute `HAND_OI_ORDER`, the Amarillo UNMAPPED list and the default hidden rows from the Order Intake table under Section 5.1 and 7.1, without seeing Section 7.1's answer (this brief's author had one Sonnet do exactly that: it agreed on every id; a second run is cheap insurance if you change the fixture rows beyond Section 7.1). | Compare with Section 7.1; any difference is investigated before it is asserted. |
| Phase 3, after `check_bas.py` passes | Sonnet | Review the VBA diff only, with Section 5 of this brief: compile risks (reserved names, `Array` sizes, continuations, a `Dim` left over from a deleted helper, an `Option Explicit` miss), rule mismatches against Section 5, any surviving mention of tonnes or Over10. | Fix what is real; ignore style. |
| Phase 4 | Sonnet | The three Markdown edits of Section 8.2 from a brief that carries the exact replacement text, the new Run Log lines pasted from `expected.json`, and the Austin lines pasted from the reference build's Run Log sheet. | Diff the result; grep for the stale words (next row). |
| End of Phase 4 | Haiku | A sweep: grep every edited file and `MAP.md` for `tonnes`, `Over10`, `14 class`, `14 names`, `13 classes`, `13 rows`, `9 warnings`, `11 items`, `ROC rule`, `by the rule`, `not a macro size class`, `frame class`, `LOADER_RULE`, `TONNES_RULE`, `v2.6` / `2.6` (where it names the macro), and report every hit with file and line. Also count the pages of the rebuilt manual PDF. | Every hit is either fixed or is in a historic document Section 1 lists as untouched. |
| Phase 5 | Haiku | Read the four class tables (`SizeClasses()` in the .bas, `SIZE_CLASSES` in the reference, `CLASSES` in the checker, `CLASS_RANK` in the fixtures) and confirm all four are the Section 5.1 list in the same order with no `Over10`. | A one-line yes or the differing entries. |

Give each worker the file paths, the line numbers from Section 3 and the exact text it needs; it has no other context. Never have a worker run `git`.

### 6.2 Phase 0: Orient

Read Section 3. Run the baseline so you know it is green before you touch anything:

```
cd source/open-orders-intake-inventory
python3 check_bas.py
python3 make_test_fixtures.py test-fixtures
python3 build_report_reference.py test-fixtures /tmp/ref-default.xlsx --size-class-file "test-fixtures/Size Class.xlsx"
python3 check_report.py /tmp/ref-default.xlsx --expected test-fixtures/expected.json
python3 run_macro_libreoffice.py test-fixtures /tmp/lo-default.xlsx --size-class-file "test-fixtures/Size Class.xlsx"
python3 check_report.py /tmp/lo-default.xlsx --expected test-fixtures/expected.json --compare /tmp/ref-default.xlsx
```

(The variant run adds `--inventory-month 2026-08-15 --show-cancelled --show-ssl`, or `--set` for the harness; see each script's `--help`.) If the baseline is not green, stop and report; do not build on a red baseline.

### 6.3 Phase 1: The executable spec first (fixtures and checker)

1. `make_test_fixtures.py`: Section 7.1 (the class table, the size class rows, the expected classes, the log lines, `HAND_OI_ORDER`, `self_check`, `version` 2.7). Regenerate `test-fixtures/` and `expected.json`; `self_check` passes.
2. `check_report.py`: Section 7.3. Run it against the v2.6 reference output from Phase 0: R2F (Table 1), E6 (the tonnes log line) and E2 (the E25 and S590 classes, the order) must now **fail** for the right reasons. That failing run is evidence the checks bite; keep its output for the report.

### 6.4 Phase 2: Reference v2.7

`build_report_reference.py`: Section 5 in Python (the class list and rank, the reader without the ROC check, the assignment without the tonnes branch, Table 1 with three families and the unbanded last row, the `FF_*` rows, the docstring). Then `make_size_class_master.py` (Section 8.1), regenerate the master list, build the fixtures (default and variant) and the Austin sample; `check_report.py --expected` passes on both fixture runs; `compare_with_handbuilt.py` on the Austin build shows what Section 7.2 says.

### 6.5 Phase 3: Macro v2.7

`build-open-orders-report.bas`: Section 5 in VBA. `check_bas.py` after every edit. Then the LibreOffice run on both fixture runs (with `--expected` and `--compare` against the reference builds), `--check-only` on the fixtures (the Last run box text of Section 7.1) and the Austin sample (`--compare` against the reference's Austin build; the Last run box of Section 7.2). The Sonnet diff review (Section 6.1).

### 6.6 Phase 4: Documents and generated outputs

Section 8. Rebuild the manual, the master-list PDF, the P11 PDF and the flow drawings with `./build.sh`; the manual must stay one page. The Haiku sweep.

### 6.7 Phase 5: Review and ship

The Haiku class-table check. Fill in Section 10 honestly. One commit (or a few logical ones), `MAP.md` in the same commit, push to `main` (and to the session branch).

---

## 7. Fixtures and checks

### 7.1 Fixtures (hand-computed for this brief from the tables in `make_test_fixtures.py`, confirmed by an independent Sonnet recomputation)

**`Size Class.xlsx`** (`SIZE_CLASS_ROWS`, header unchanged: HP, Rated Operator Capacity, Name of machine, Size Class; the ROC column stays in the file to prove it is ignored):

- Row 6 `S590`: class `2000<2300` becomes **`Med Frame`**.
- Row 8 `S76` keeps `2600<3000` (an SSL model with a loader class is accepted: no family check).
- Row 4 `T595` keeps ROC 2350 and class `2000<2300`; its ROC warning **disappears**.
- Appended after the existing 21 rows, so no existing row number moves: row 23 `(23.5, 760, 'S70', 'ex small')` (lower case, accepted and written `EX Small`); row 24 `(None, None, 'E100', 'Over10')` (skipped: no longer a class).
- `E25` and `T999` stay absent.

Expected: **20 usable rows, 3 warnings**, in this order:

```
Size class file row 21 T630: "2000-2300" is not a size class - row skipped
Size class file row 22 T650: listed again (first in row 20) - row 20 is used
Size class file row 24 E100: "Over10" is not a size class - row skipped
```

and the file line `Size class file Size Class.xlsx: 20 usable rows, 3 warnings`.

**Order Intake expected classes** (the `size` field): `A05` E25 `>02-03` becomes **`UNMAPPED`**; `A06` S590 `2000<2300` becomes **`Med Frame`**; every other row unchanged (`A20` `e35` still `>03-04` by the case-insensitive lookup; `T740` rows by first word; `T650` by the first listing). `CLASS_RANK` in the fixtures script becomes the Section 5.1 table.

**`HAND_OI_ORDER`** (class rank, status rank, newest date first with undated last, MSA with `(no MSA)` first, export order):

```
A11, A03, A18, X02, A01, A10, B01, A17, A15, C07, A09, C01, C03, C02, A02, A19,
A12, A13, A16, C05, A20, A04, A08, C06, C04, A14, A06, A07, A05
```

By class: `0<2000` A11; `2000<2300` A03; `2300<2600` A18, X02, A01, A10, B01, A17 (the 09-15 tie: X02 is `(no MSA)`, then A01 before A10 by export order); `2600<3000` A15, C07, A09; `3000<3500` C01, C03 (TBD, undated last among ORDERED), C02, A02, A19; `3500&Over` A12; `>00-01` A13; `>01-02` A16; `>02-03` C05; `>03-04` A20, A04; `>04-05` A08; `>05-06` C06, C04; `>06-10` A14; `Med Frame` A06; `UNMAPPED` A07 (ORDERED) then A05 (Not yet processed). Same order in both runs.

**Unchanged facts** (assert that they are): default hidden rows `A06, A03, A09, A19, C05, A08` (A06 now hidden as SSL under a frame class); variant hides none; the visible totals 24 and 30; the Units summaries; the detail order; the Open Orders by MSA table, split and charts; the TBD row.

**Run Log changes**: the line `Order Intake: placed by the tonnes rule (not in the size class file): E25` is gone, so each run has **23** lines (was 24). The Amarillo line ends `; UNMAPPED model groups: E25, T999` (export order: A05 before A07). The `SIZE_CLASS_WARNINGS` are the three above. `self_check`: `classes via the size class file` first three are `0<2000`, `2000<2300`, `2300<2600`; the `T740 by first word, E25 by tonnes, T999 unmapped` check becomes `{'T740': '3000<3500', 'E25': 'UNMAPPED', 'T999': 'UNMAPPED', 'S590': 'Med Frame'}`; `number of log lines` (23, 23); the line indexes of the per-MSA checks are unchanged (the warnings count is still 3 and nothing before the MSA lines moved).

**Check setup on the fixtures** (`run_macro_libreoffice.py --check-only`): the Last run box warning reads `Models in Order Intake that are not in the size class file (UNMAPPED until added): E25, T999`; take the warning count the box shows into the test protocol's step 23 (it was 6).

### 7.2 Austin sample (the shipped master list, 29 models, regenerated by Section 8.1)

The size class file loads with **0 warnings** (the nine frame-class warnings are gone, and the file's ROC column is ignored). Eight model groups are UNMAPPED: `T64`, `T66`, `T86`, `E55`, `E60`, `E88`, `S510`, `S66`, in the export's order of first appearance (copy the exact `Austin: Order Intake ...; UNMAPPED model groups: ...` line from the reference build's Run Log into the test protocol; do not type it from memory). No tonnes line. The figures do not change, because a class never changes a count: 278 rows, 108 hidden, visible total 170, November units CTL 57 / MX 43 / SSL 7, open orders 86 / 34 / 43 = 163. The Last run box after a build is amber with the lines `RunLogLineToCheck` still catches (the Austin UNMAPPED line among them; it was 11 items with the nine warnings and the tonnes line); take the exact count and wording from the harness's printed box into the protocol's step 18 and Table C.

`compare_with_handbuilt.py`: the `visible rows as a set` check still passes (170 of 170; it ignores column A). The `size-class order` check **fails** with the shipped list, and that is expected: the hand-built Austin tab classes E55 and E60 as `>05-06` and E88 as `>06-10` (15 visible rows), which the macro now writes as UNMAPPED and sorts last. Prove the macro agrees once the list is complete: build Austin a second time with a **scratch** copy of the master list (not committed) that adds `E55 >05-06`, `E60 >05-06`, `E88 >06-10`; on that build both checks pass. Report both runs. (T64, T66, T86, S510 and S66 do not appear among the hand-built tab's visible rows: their Austin rows are cancelled or SSL, so they do not affect the comparison.)

### 7.3 Checker v2.7 (`check_report.py`)

- `CLASSES` = the seventeen of Section 5.1 in order; `CLASS_RANK` from it; `OVER10` and `CLASS_RANK[OVER10]` removed; `UNKNOWN_RANK` 99 stays. R3a (every class known) and R3b (rank never decreases) then follow Section 5.1 with no other change.
- `check_file_format`, Table 1: eighteen body rows (`range(18)`), expected `CLASSES + [UNMAPPED]`; the header check unchanged; UNMAPPED bold at `r + 19`; the grid `r + 1` to `r + 19` with `medium` at `r + 1`, `double` at `r + 2`, `medium` at `r + 8` and `r + 15`; banding `r + 2` to `r + 18` (the UNMAPPED row outside it); new: the Family and Rule cells of every class row equal Section 5.5's texts (the checker today reads only column A; add the two columns, it is three lines).
- `fill_cells` for `ff`: the peach cell at `(r + 19, 1)`.
- E2 / E6 unchanged in mechanism; they compare the new `expected.json`.
- Docstring: v2.7, the class count.

### 7.4 LibreOffice run and static check

`run_macro_libreoffice.py`: no substitution changes expected; `AUTOFILTER_CALLS` and the `SetUpPage` count unchanged. If the harness stops on a changed count, you changed something outside the contract. `check_bas.py`: no change expected; it will object if a deleted helper leaves a dangling `Dim` or an `Array` grows past its limit (the new `SizeClasses()` has 17 arguments).

---

## 8. Documents and generated outputs

### 8.1 The master list and its PDF

`make_size_class_master.py`: the docstring (lines 13 to 18: the frame classes are macro classes now); `NOT_A_CLASS` and the `ref.canon_class` note go, replaced by a hard assertion that every class in the script is one of `ref.SIZE_CLASSES` (the script must fail loudly, not annotate, if someone types a class the macro will not read); `ref.loader_class` is no longer imported; the E40 REVIEW note (line 74) becomes `[]` (there is no rule to disagree with; Section 11, question 2). Result: 29 models, **3** REVIEW rows (T550, T62, T740: blank HP or ROC on the sheet). Regenerate `output/.../size-class-master.xlsx`.

`make_size_class_pdf.py`: the paragraph at 111 to 113 (frame classes skipped) becomes a one-sentence note that the frame classes are the SSL classes the macro reads; the paragraph at 116 to 118 (nominal tonnes = model number / 10, Over10, "an E-model missing from the list is still placed by this rule") becomes: the MX classes are read from the list, a model missing from it is UNMAPPED; the CTL rule paragraph likewise says the class is read from the list (keep the ROC figure column; it is the owner's reference). The intro's REVIEW count is computed, so it follows. Regenerate `size-class-master.tex` and `./build.sh` it.

### 8.2 `macro-instructions.md`, `excel-test-protocol.md`, `copilot-prompts.md` (v2.7)

`macro-instructions.md`: line 82 (troubleshooting: the "placed by the tonnes rule" paragraph goes; UNMAPPED is the only state); 121 (sort order: CTL, MX, SSL classes in the File Format order, then UNMAPPED last); 127 (a row is used when it has a name and one of the 17 classes; the ROC sentence goes); 129 (the three families and that the class comes from the file); 130 (a model not in the list is UNMAPPED, no exception); 131 (the master list: 29 models, 3 REVIEW rows, loads with 0 warnings); 156 (the Run Log order: the tonnes rule item goes); 160 (the Last run patterns: `tonnes rule` and `not in the size class file` go); 207 (the behaviour list); 262 (the setup-check message, new wording of Section 5.4, "(or by the tonnes rule for E-models)" goes). Section 7 (decisions): add the Section 11 decisions as "decided during the build (2026-10-15)". Section 8: the class list is in one place per file (name the four).

`excel-test-protocol.md`: line 36 (Run A's expected RUN LOG items: `E25 UNMAPPED` instead of `E25 by the tonnes rule`); 84 (the T595 line goes), 85 (stays), the new E100 line added after the T650 line; 103 (the tonnes line goes); the Run A and Run B Run Log blocks replaced by the lines of the new `expected.json` (paste, never type); 136 (step 18: the Austin Last run box, Section 7.2); 154 to 168 (the Austin file line `29 usable rows, 0 warnings`, the nine warnings gone, the tonnes line gone, the UNMAPPED line pasted from the reference build); 172 (the SSL rows now carry frame classes but are still hidden; the 170 visible rows unchanged); 189 (step 23: the new warning wording and count); 229 (Table C's item count). Grep the file for `tonnes`, `11 items`, `9 warnings`, `frame class` when done.

`copilot-prompts.md` P11: lines 53 to 54 (the MX classes are the seven `>00-01` to `>06-10`, no Over10, "as the business defines them", no formula), 85 (an E-model not on the slide gets no class by rule: Copilot leaves Size Class blank and the owner fills it from the spec sheet, as the loaders already do), 111 (`one of the 17 names`), 410 (the context block: three families), 491 (the Run Log's UNMAPPED lines). Add the four SSL frame classes to Step 1's list so Copilot keeps the slide's SSL labels as they are. Then `python3 make_size_class_copilot_pdf.py` and `./build.sh` the P11 `.tex` (one page; its Rev. 2, dated the build day).

### 8.3 Manual, Rev. 11 (`open-orders-macro-manual.tex`, one page)

Header `Rev.~11 \quad <build date>`; Scope `(v2.7)`; Running the report step (c): `Read the Run Log, Last run box and chart; a model the size class workbook does not list is UNMAPPED: add it there and rebuild.`; Road map row 2 `Done <date> (v2.7).`; row 4 (Adopt the list) "Done when": `The workbook is the only source of a class; nothing UNMAPPED.` Rebuild with `./build.sh source/open-orders-intake-inventory/open-orders-macro-manual.tex`; confirm one page (`pdfinfo`); commit `.tex` and `.pdf` together.

### 8.4 The flow drawing (Rev. 15) and the team review deck

`open-orders-macro-flow.tex` line 136 and `open-orders-macro-flow-alt.tex` line 129: the bullet `Size class from the master list; MX E-models not on it by tonnes; else UNMAPPED` becomes `Size class from the master list only; a model not on it is UNMAPPED`. The header comments (line 8 / 9) likewise. Title block: Rev. 15, the build date, macro v2.7. Keep the two files in step (the alt is a copy without the Open Order block). Rebuild both PDFs and `open-orders-team-review.pdf` (its sheet 3 inputs the flow source; bump the deck's Rev. to 11 with the date in its own title block, nothing else on sheets 1 and 2). Nothing else on the drawing changes.

### 8.5 `MAP.md`

Rows 73 to 75, 79 to 82, 85, 87, 91, 100, 101, 115 (the P11 sentence), 117, 118, 119, 120, 121, 122, 125 to 130 as their files change (v2.7 where the version is named; no tonnes rule, Over10, ROC check, frame-class warnings, "11 items" or "13 REVIEW rows"); a row for `size-class-copilot-prompt.pdf` says Rev. 2; this plan's row ends with "Executed <date>."

---

## 9. Not deliverables, restated for the implementer

Do not: add class aliases; add a family check; add a Config setting; touch the Order Intake columns or the visible total; reword existing Run Log lines other than the ones named in Section 5; change the Open Orders by MSA or Inventory tabs; edit the historic documents of Section 1; change `compare_with_handbuilt.py` to make the Austin class-order check pass (the fix is the owner's list, Section 7.2).

---

## 10. Acceptance checklist (fill in, honestly, in the final report)

- [ ] `check_bas.py`: PASS.
- [ ] `make_test_fixtures.py`: `self_check` passes with the Section 7.1 order, classes, warnings and 23 log lines asserted.
- [ ] A v2.6 reference build fails the v2.7 checker for the expected reasons (R2F Table 1, E2 classes and order, E6 the tonnes line); output kept.
- [ ] Reference, fixtures default and variant: `check_report.py --expected` PASS.
- [ ] Reference, Austin sample with the regenerated master list: `check_report.py` PASS; 0 file warnings; the eight UNMAPPED groups; figures unchanged (Section 7.2).
- [ ] `compare_with_handbuilt.py`: set check PASS and class-order check FAIL with the shipped list (reported as expected), both PASS with the scratch list extended by E55, E60, E88.
- [ ] Macro in LibreOffice, fixtures default and variant and Austin: `--expected` (fixtures) and `--compare` against the reference builds PASS; `--check-only` Last run box text as Section 7.1; the Austin Last run box count and wording taken into the protocol.
- [ ] Sonnet VBA diff review done; findings listed with what was fixed or why not.
- [ ] Haiku sweep: no stale word left in an edited file; the four class tables identical.
- [ ] Master list regenerated: 29 models, 3 REVIEW rows; `size-class-master.pdf` rebuilt; POC report rebuilt.
- [ ] Manual Rev. 11 one page; flow drawings Rev. 15 and the deck rebuilt; P11 PDF Rev. 2.
- [ ] Instructions, protocol and P11 updated; `MAP.md` rows updated in the same commit; this plan marked Executed.
- [ ] Not run: real Excel (the owner runs the test protocol).
- [ ] Section 11 decisions and anything decided beyond them listed for the owner.

---

## 11. Decisions taken in this brief, and open questions (defaults apply until answered)

| # | Question | Default taken here |
|---|---|---|
| 1 | Should the reader also accept the sheet's own spellings (`<2000`, `3500 & Over`) as aliases of `0<2000` and `3500&Over`, so the owner's sheet needs no respelling? | No. One spelling per class, as printed on the File Format tab; `make_size_class_master.py` and P11 respell. Easy to add later (one alias table in three files) if the owner asks. |
| 2 | The E40 REVIEW note (sheet says `>04-05`, the old tonnes rule said `>03-04`). | Dropped. The sheet is the rule now; there is nothing to compare with. |
| 3 | Should the macro warn when a model's class family does not match its `PRODUCT_TYPE_ABBR` (an S-model with a CTL class, an E-model with a frame class)? | Not built. The sheet is trusted; a Run Log count line could be added later in all three places if the owner wants a guard. |
| 4 | The eight Austin model groups not on the sheet (`T64`, `T66`, `T86`, `E55`, `E60`, `E88`, `S510`, `S66`). Are `T64`/`T64-2`, `T66`/`T66-2`, `T86`/`T86-2`, `S66`/`S66-2` the same machines, and which class do `E55`, `E60`, `E88`, `S510` get? | The owner adds rows to the workbook (the hand-built tab used `>05-06`, `>05-06`, `>06-10` for the three excavators). The macro does not guess a suffix rule. |
| 5 | Order of the SSL block: after MX (chosen) or between CTL and MX? | After MX: existing CTL and MX rows keep their places, and the product-type columns elsewhere read CTL, MX, SSL. |
| 6 | `Over10` dropped as a class? | Yes; the sheet has no such class and nothing produces it. A future >10 t class is a one-line addition to the four tables. |

---

## 12. Things that went wrong once already, and new traps

Carried over from the v2 and v2.5 plans:
- Reading the sample workbooks cell by cell wasted context. Use scripts.
- No theme colours; RGB only. The `.bas` stays CRLF and ASCII, no line over 1,000 characters, no statement over 20 continuations, no `Array(...)` over 50 arguments; `check_bas.py` after every edit.
- Reserved words as variable names break compilation; `Name`, `Type`, `Date`, `Len` are in the list.
- `FF_BUILD_ROW` and its siblings are defined in the macro **and** the reference: change both or the tables land on different rows and the C-rules fail.
- openpyxl drops charts on load; nothing here reloads a build, keep it that way.

New for this change:
- **The banding hides the peach.** Table 1's UNMAPPED row moves from an odd body row to an even one; without the unbanded last row the green band wins and R2 fails on the fill. Section 5.5.
- **The ROC column stays in the files.** The master list and the fixture both still have it; the reader must not warn about it, and `AliasColumn` must simply not be asked for it.
- **`mUnmapped` is per MSA and `mUnmappedAll` across MSAs**; the E25 row joins both lists through the existing `AddToList` calls. Do not invent a third list.
- **The `E#*` test** (`uMg Like "E#*"`) is the only place the product type entered classification; after the change `abbr` is not consulted at all in the assignment, and the Python `re.match(r'^E(\d+)', ...)` goes with it.
- **`canon_class` in the reference is imported by `make_size_class_master.py`**; keep its name and signature, drop `Over10` from it.
- **`self_check` line indexes** of the per-MSA log checks (10, 11, 16, 17, 9) do not move, because the three file warnings still precede the MSA lines; the tonnes line that disappears was after them. If an index does move, you added or removed a line the contract did not name.
- **`compare_with_handbuilt.py` will go red** on the class-order check with the shipped list. That is the data, not the macro (Section 7.2); do not "fix" the script.
- **The flow drawing and the deck share a source**: `open-orders-team-review.tex` inputs `open-orders-macro-flow.tex` with docmute; rebuild the deck after the flow, and keep `-alt.tex` in step by hand.

---

## 13. The prompt to paste into Claude Opus (high reasoning effort)

> Read `output/open-orders-intake-inventory/size-class-file-only-implementation-plan.md` in full and execute it as the implementer, at high reasoning effort. Follow its standing rules exactly: push to `main` (and to this session's branch), keep `MAP.md` current, build to the Section 5 contract, implement every rule in the macro, the Python reference and the checker, assert the Section 7.1 order and figures in the fixtures, run the macro in LibreOffice and compare it with the reference, run `compare_with_handbuilt.py` both ways as Section 7.2 says, regenerate the master list and its PDF, rebuild the POC report, and keep the one-page manual (Rev. 11), the flow drawings (Rev. 15), the instructions, the test protocol and Copilot prompt P11 current. Use Sonnet and Haiku workers only where Section 6.1 says and with self-contained briefs; verify their output yourself and never let them change the contract. Work through Phases 0 to 5 in order and finish with the Section 10 checklist filled in honestly, including what was not run (real Excel) and the Section 11 decisions I should confirm.
