# CLAUDE.md

## Purpose
This repo holds the owner's **work** material: VBA macros and the prompts, specs, reports, test tooling and documents around them. Personal projects do not belong here: they live in the `gustjf/ai` repo.

## Layout
- Repo root: the standalone macros (Visio, Fleet Fuel) plus `README.md`, `MAP.md`, `CLAUDE.md`, `build.sh`, `LICENSE` and config.
- `Specs/`: reusable prompts and specs to paste into an AI. The owner opens these.
- `output/<Topic>/`: finished deliverables (PDFs, manuals, plans, workbooks). The owner opens these.
- `source/<Topic>/`: everything behind them (`.bas` and `.tex` sources, scripts, test fixtures, sample data). The owner rarely opens these.

## Standing rules
- **Always push to `main`. Do not ask for confirmation.** Commit and push directly to `main` (no feature branches, no pull requests unless explicitly requested).
- **Cloud sessions**: a session given a `claude/<name>` branch pushes the same commit to that branch and to `main` (a fast-forward), so no branch is ever ahead of `main`. Never create a history unrelated to `main`, never force-push.
- **Read `MAP.md` first** to locate files. Do not scan the repo; open only the files the map says are relevant.
- **Keep `MAP.md` current**: whenever you add, rename, move, or delete a file, update its row (path, type, one-line description) in the same commit, under the right section ("Open these" for `Specs/`, `output/` and the root macros, "Behind the scenes" for `source/` and plumbing) and topic heading.
- Keep files self-contained and reusable (prompts should be pasteable as-is).
- Use concise, descriptive filenames in lowercase-hyphen style.
- **Mirror outputs and sources**: generated deliverables (e.g. PDFs) go in `output/<Topic>/`; their source files (e.g. `.tex`) go in `source/<Topic>/` using the same subfolder name, so the two trees mirror each other. `./build.sh` compiles `source/<Topic>/x.tex` to `output/<Topic>/x.pdf`.
- **Keep the Open Orders macro manual current**: whenever `source/open-orders-intake-inventory/build-open-orders-report.bas` changes anything a user sees or does (setup, buttons, Config settings, how to run or build), update `source/open-orders-intake-inventory/open-orders-macro-manual.tex` in the same commit: bump the Rev. and date in its header, keep it to one page, rebuild `output/open-orders-intake-inventory/open-orders-macro-manual.pdf` with `./build.sh`, and commit both.
- **Specs live in `Specs/`**: reusable prompts/specs (e.g. `*-spec.md`) go in the `Specs/` folder. `Specs/exam-style-latex-spec.md` and `Specs/exam-style-multipage-latex-spec.md` are copies of the same files in `gustjf/ai`; when one changes, change the other.
