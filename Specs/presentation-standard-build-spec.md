# Presentation Standard: Build Spec and Run Prompts

This file sets up a later set of Claude sessions to build a **company presentation standard**. The standard is a small library of LaTeX presentation templates, plus the tools and training that let colleagues with weaker work AIs (GPT, Gemini, Sonnet 4.6 in their work tools) produce decks as good as the Open Orders by MSA team review. It is released to the owner's SharePoint document control site. The owner's manager has approved defining a new standard if it works.

It runs as four stages, each started by pasting one prompt from Section 9:

| Stage | Who | Effort | Job |
|---|---|---|---|
| 1. Plan | Fable | high | Turn this spec into a build plan with every design decision made. |
| 2. Build | Opus orchestrating Sonnet subagents | high | Execute the plan: templates, tools, prompts, docs, release package. |
| 3. Review | Fable | high | Audit the build against the plan and the rubric (Section 7); write a numbered fix plan. |
| 4. Fix | Opus orchestrating Sonnet and Haiku subagents | high | Apply the fix plan, re-run every check, release. |

Stages 3 and 4 may run a second time (Section 8). Fill in Section 1 before Stage 1.

---

## 1. Owner inputs (fill in before Stage 1)

Defaults are in brackets; delete what does not apply.

- **Presentation types for release 1** (3 to 5; ask the manager and two heavy presenters which they give most):
  [1. Team review / meeting deck: agenda timeline, review sheets, discussion questions (the MSA deck pattern)
   2. Process or system flow one-pager (the build-flow drawing pattern)
   3. Monthly KPI / project status deck: figures and charts read from a CSV exported from Excel
   4. Proposal / decision one-pager (A3 style: problem, current state, options, recommendation, ask)
   5. How-to / training one-pager (the exam-style spec pattern)]
- **Formats**: [11 x 17 landscape for print and wall review, plus 16:9 for screens. The plan may drop one per type with a reason.]
- **Colleague toolchain the owner will install**: [MiKTeX + TeXworks (ships with MiKTeX) on Windows, local only. Overleaf only if the data use policy covers it.]
- **Company look**: [Neutral house style from the MSA deck (green `548235` headers, white boxes, title block bottom right). Replace with brand colours and logo rules if the brand team has them.]
- **Document control**: document number [`TBD`], owner [Gust F.], approver [manager], SharePoint library path [`TBD`].
- **Pilot users**: [2 colleagues, named in the training plan only by role].

---

## 2. Mission and the one principle

Make it possible for a colleague who has never used LaTeX to go from their notes and an Excel export to a correct, good-looking deck in under an hour, with their work AI doing only the writing.

**The principle: AI only for what needs AI.** The owner will not use AI, theirs or the colleagues', for work a person can be taught or a script can do. Decide every step by this split:

| AI does (judgment and language) | Templates, scripts or people do (deterministic) |
|---|---|
| Turn rough notes into tight headlines, bullets and speaker notes within the field limits | Layout, fonts, colours, spacing, page size: fixed in the template |
| Choose which template and which optional sections fit the audience | Compiling, and checking overflow, page count and missing fields |
| Draft discussion questions, an agenda, a recommendation's wording | Every number, table and chart: read straight from the colleague's CSV by LaTeX (`pgfplots`, `pgfplotstable`), never retyped by an AI |
| Review the wording against the audience and purpose (optional prompt) | Dates, revision numbers, title block fields: typed once by the person |

Anything in the right-hand column that ends up as an AI step in the plan is a defect.

---

## 3. Fixed constraints

1. **Repo rules.** Read `CLAUDE.md`, then `MAP.md`. Keep `MAP.md` current in the same commit, use lowercase-hyphen names, and mirror sources and outputs: `source/presentation-standard/` and `output/presentation-standard/`. Push to `main`, and also to the session's `claude/*` branch if one is given. Never force-push.
2. **Synthetic data only in this repo.** Every example deck and CSV is made up; real company decks live only on SharePoint.
3. **Colleague machines are Windows with no admin rights after install, and no Python, Perl or Node.** Everything a colleague runs is a double-click `.bat` and built-in PowerShell, with no other runtime. Note: `latexmk` on MiKTeX needs Perl, so the build script calls the engine directly, the needed number of times.
4. **The standard must work offline** after the one-time install (MiKTeX packages preinstalled or listed for IT; no on-the-fly downloads at a colleague's desk).
5. **Weak-model-proof by construction.** A colleague's AI edits exactly one file per deck, the content file. It never touches the template, preamble or TikZ. If the AI writes rubbish into the content file, the result is a clear error message or a visible on-page flag, never a silently broken layout.
6. **Teachable.** A person can learn the whole workflow in a 30-minute session using a one-page quick-start.
7. **Golden reference.** The house style and quality bar are `source/open-orders-intake-inventory/open-orders-team-review.tex` (306 lines) and `source/open-orders-intake-inventory/open-orders-macro-flow.tex` (191 lines), with their PDFs in `output/open-orders-intake-inventory/`. Read lines 1 to 120 of the team review for the preamble, styles, title block and sheet grid. Read further only when a template copies a specific sheet.
8. **Repo build.** `./build.sh` (latexmk) builds every `.tex` under `source/` in the cloud. The colleague build (`build.bat`) must produce the same PDF from the same files.

---

## 4. What Stage 1 must decide (the planner's questions)

Fable answers each question in the plan with a decision and a one-line reason. Stage 2 must not redesign anything.

1. **Content file format.** Default: `content.tex` containing only `\set{field}{value}` lines, comment headers and `\begin{items}...\end{items}` lists, with the fields documented per template. Decide how special characters (`& % $ # _ { } ~ ^ \`), smart quotes, en and em dashes, and pasted Unicode from Word, Outlook and Copilot are handled. Options are pdfLaTeX with `inputenc` and a check, or LuaLaTeX with `fontspec`. Weigh this against MiKTeX's slow first LuaLaTeX run and font cache.
2. **Field limits and how they are enforced in LaTeX itself.** For example, measure each text box with `\sbox` and `\ifdim` against its allotted height. On overflow, still produce the PDF but print a red **OVERFLOW: field name** stamp on the page and write a log line, so a non-technical user sees the problem where it is. Also set character and bullet-count limits, stated in the prompt.
3. **Missing and unknown fields.** A required field left empty prints a visible placeholder (**MISSING: field name**). An unknown `\set` key (an AI typo) stops with a plain-English error naming the line.
4. **Data path.** How a colleague's Excel data becomes tables and charts with no AI and no retyping: the CSV export settings (delimiter, decimal, date format, encoding), the column-name contract per template, and what happens when a column is missing.
5. **Template inventory.** For each Section 1 type: the sheets or pages, which are optional (switches in the content file), the formats, and the field list with limits and one example value each.
6. **Shared style package.** One `.sty` holding the house style (colours, TikZ styles, title block, fonts, the overflow and missing machinery) that every template loads, so the look is defined once and changed in one place with a version number.
7. **Colleague tools.** `build.bat` (compile, open the PDF, show the errors and stamps in plain English) and `check.ps1` (scan the content file for unescaped characters, empty required fields and over-limit text before compiling; scan the log after). Also decide the "new deck" workflow: copy a template folder, rename it, edit the content file.
8. **Colleague prompts.** For each template, one pasteable prompt for a weak model. It takes the colleague's notes and the template's field table, and returns **only** the content file in one code block. It includes one complete filled example and the limits. Add one optional "wording review" prompt. Prompts must fit a small chat window: target under 1,500 words each including the example.
9. **The standard document.** A controlled document (multi-page exam-style, `Specs/exam-style-multipage-latex-spec.md`) that states the rules of the standard: when to use which template, the AI-versus-person split from Section 2, naming and revision rules, review and approval of decks, how to request a new template, and a change log. Header fields follow Section 1's document control inputs.
10. **SharePoint release package.** The folder layout (for example `01 Standard/`, `02 Templates/<name>/`, `03 Prompts/`, `04 Tools/`, `05 Examples/`, `06 Training/`), what is read-only, the versioned zip, and how a colleague gets a fresh copy.
11. **Training.** A one-page quick-start (exam-style spec), a 30-minute session plan with a practice exercise using synthetic data, and an install checklist the owner works through on each PC.
12. **Acceptance test.** A protocol the owner runs on a real colleague PC with the colleagues' real work AIs. It covers install, building every example, deliberately breaking each check, and one prompt run per template through GPT, Gemini and Sonnet 4.6, with pass criteria. It follows the pattern of `output/open-orders-intake-inventory/excel-test-protocol.md`.
13. **Build waves and agent briefs for Stage 2** (Section 6), including which outputs are verified by script versus by rendered-page review.

---

## 5. Deliverables (Stage 2 builds, Stage 4 finalizes)

| Deliverable | Path |
|---|---|
| Build plan (Stage 1 output) | `source/presentation-standard/build-plan.md` |
| Shared style package | `source/presentation-standard/tex/<name>.sty` |
| One folder per template: `template.tex`, `content.tex` (filled synthetic example), `fields.md` (field table with limits), sample `data.csv` where used | `source/presentation-standard/templates/<template-name>/` |
| Colleague tools | `source/presentation-standard/tools/build.bat`, `check.ps1` |
| Colleague prompts, one per template plus the wording review | `source/presentation-standard/prompts/` (Markdown, pasteable) |
| Standard document, quick-start, training plan | `source/presentation-standard/docs/*.tex` to `output/presentation-standard/*.pdf` |
| Example PDFs, one per template and format | `output/presentation-standard/examples/` |
| Acceptance test protocol, install checklist | `output/presentation-standard/*.md` |
| Release zip for SharePoint | `output/presentation-standard/presentation-standard-v<N>.zip` |
| Review report and fix plan (Stage 3 output) | `source/presentation-standard/review-<round>.md` |
| Map rows | `MAP.md` |

---

## 6. Agent rules (Stages 2 and 4)

- The orchestrator (Opus) owns the shared `.sty`, assembly, compiling, checks, `MAP.md` and the commit. It writes the `.sty` first, before fanning out, because every template depends on it.
- Sonnet subagents each own one template folder, one prompt, one document or one tool, in parallel waves. Give each a self-contained brief: the decisions it needs from the plan, the field table, the `.sty` interface, and the golden-reference line ranges it may read. Forbid repo-wide reading. Keep each subagent's context under 80,000 tokens.
- Haiku subagents (Stage 4, and Stage 2's prompt test) stand in for the colleagues' weak models. For each template, a Haiku agent gets **only** the colleague prompt and a page of messy synthetic notes, and returns a content file. That file must compile and pass `check.ps1`'s rules with no human repair. Run it 3 times per template; at least 3 of 3 must pass. This is a proxy only: the owner's acceptance test with the real tools is the actual gate.
- Nothing is accepted from a subagent without the orchestrator compiling it and rendering every page to PNG (`pdftoppm -r 60`) for a look. Check page count, overflow and missing stamps, overlaps, text running off boxes, and parity with the golden reference.
- `check.ps1` cannot run in the Linux cloud container unless PowerShell is installed (`pwsh`). Install `pwsh` if the network allows; otherwise write a line-for-line equivalent test in bash, say so in the report, and add the real script run to the owner's acceptance test.

---

## 7. Review rubric (Stage 3 grades against this; Stages 1 and 2 build to it)

Each item is Pass, Fail (with file:line and a concrete fix) or N/A.

1. **Principle.** No AI step does deterministic work (Section 2 table). No number in any example reached the PDF through an AI.
2. **Weak-model proof.** Every bad-input case produces a visible stamp or a plain-English error, never a silent layout break. Test cases: empty field, over-limit field, unknown key, each special character, smart quotes, an emoji, a missing CSV column, an extra CSV column, wrong decimal separator.
3. **House style parity.** Rendered pages match the golden reference's style: colours, title block, spacing, type sizes, no overlaps.
4. **One source of style.** All look-and-feel lives in the `.sty`; templates hold layout only; content files hold content only.
5. **Colleague path.** Copy folder, edit `content.tex`, double-click `build.bat`, PDF opens: no step needs the command line, admin rights, internet or the owner.
6. **Prompts.** Each colleague prompt is under the word target, self-contained, includes limits and a full example, asks for the content file only, and passed the Haiku proxy test 3 of 3.
7. **Docs.** The standard, quick-start and training plan agree with the templates and tools as built; the quick-start fits one page; every document has Rev. and date.
8. **Release.** The zip unpacks into the planned SharePoint layout and builds offline on a clean MiKTeX install, according to the install checklist.
9. **Repo.** `./build.sh` builds every `.tex` cleanly with no overfull boxes; `MAP.md` rows exist for every file; examples use synthetic data only.

---

## 8. Stopping rule

At most two review-and-fix rounds. After round 2, any open Fail goes into `source/presentation-standard/open-items.md` with the reason and a proposed fix for the owner, and release proceeds if rubric items 1, 2 and 5 pass. Reviewers judge against the rubric and the plan, not personal taste; a new idea goes to open items, not into the fix plan.

---

## 9. Run prompts (paste one per session)

### Stage 1: Plan (Fable, effort high)

```
Read CLAUDE.md, MAP.md, then Specs/presentation-standard-build-spec.md in full. You are the architect for Stage 1. Section 1 holds my inputs. Answer every question in Section 4 with a decision and a one-line reason, and write the result as source/presentation-standard/build-plan.md. Write it as an architect's brief that an Opus orchestrator can execute without redesigning anything. Include the field table for every template (field, required or optional, limit, example), the .sty interface, the exact content-file grammar, the bad-input test cases, the build waves with self-contained Sonnet subagent briefs, and the checks that close each wave. Hold to the Section 2 principle: flag and remove any step where AI does work a script or a taught person could do. Read the golden reference only at the line ranges in Section 3.7. Do not build templates in this stage. List any decision you could not make without me at the top under "Owner questions" with your recommended default, then proceed on the defaults. Update MAP.md, commit and push per CLAUDE.md.
```

### Stage 2: Build (Opus orchestrator, effort high, Sonnet subagents)

```
Read CLAUDE.md, MAP.md, Specs/presentation-standard-build-spec.md (Sections 2, 3, 5, 6, 7) and source/presentation-standard/build-plan.md in full. You are the Stage 2 orchestrator. Execute the build plan exactly; do not redesign. Write the shared .sty yourself first, then run the plan's waves with Sonnet subagents (model "sonnet") in parallel, each on a self-contained brief with no repo-wide reading. Run the Haiku proxy prompt test from Section 6. Compile everything, render every page to PNG and look at it, run every bad-input case, and fix what fails before moving to the next wave. Grade yourself against the Section 7 rubric and put the table at the end of a short build report in source/presentation-standard/build-report.md, with any rubric Fail you could not fix. Build the release zip. Update MAP.md, commit and push per CLAUDE.md.
```

### Stage 3: Review (Fable, effort high)

```
Read CLAUDE.md, MAP.md, Specs/presentation-standard-build-spec.md, source/presentation-standard/build-plan.md and source/presentation-standard/build-report.md. You are the Stage 3 reviewer; be hard and specific. Grade the build against every Section 7 rubric item. Verify it yourself, not from the build report: compile, render pages to PNG and look at them, run the bad-input cases, read each colleague prompt as a weak model would, and check the docs against the files as built. Write source/presentation-standard/review-<round>.md: the graded rubric table, then a numbered fix plan for an Opus orchestrator with Sonnet and Haiku subagents. Each fix gets file:line, what is wrong, the exact change, how to verify it, and which agent tier should do it (Haiku for mechanical edits, Sonnet for writing or layout, the orchestrator for the .sty and tools). Order by severity. Obey Section 8: no new features, new ideas go to open items. Update MAP.md, commit and push per CLAUDE.md.
```

### Stage 4: Fix (Opus orchestrator, effort high, Sonnet and Haiku subagents)

```
Read CLAUDE.md, MAP.md, Specs/presentation-standard-build-spec.md (Sections 6 to 8) and the latest source/presentation-standard/review-<round>.md. You are the Stage 4 orchestrator. Apply the fix plan in order, giving each item to the agent tier it names (model "haiku" or "sonnet") with a self-contained brief; do the .sty and tools yourself. After all fixes: rebuild everything, re-render and look at every page, re-run all bad-input cases and the Haiku proxy test, and re-grade the Section 7 rubric. Append the result to the review file as "Fix round <round> result". If this was round 2, follow Section 8: write open-items.md and release. Bump revisions on changed documents, rebuild the release zip, update MAP.md, commit and push per CLAUDE.md.
```

### After Stage 4 (owner, by hand)

Run the acceptance test protocol on one real colleague PC with the real work AIs. Then upload the release zip to the SharePoint library, route the standard document for approval, and hold the 30-minute training with the pilot users. Bring every acceptance-test failure back as a new Stage 3 input.
