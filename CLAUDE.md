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
- **Version and change date in file names (from 2026-10-16)**: every document Claude creates or changes (specs and prompts in `Specs/`, deliverables in `output/`, and their `.tex` sources in `source/`) is named `<name>-v<N>-<YYYY-MM-DD>.<ext>`, e.g. `Specs/gust-swi-template-v6-2026-10-23.md`. A new file starts at `v1` with the date it was created. Each later change raises N by 1 and sets the date to the day of the change. Rename the file with `git mv` (no old copies kept) and, in the same commit, update its `MAP.md` row and every path that points to it (other files, scripts, this file). A mirrored `.tex`/`.pdf` pair always has the same versioned name. An existing document gets its version name the first time it changes (start at `v1`, or one above the Rev. printed in it). Exempt: names that tools, code or another repo depend on (`README.md`, `MAP.md`, `CLAUDE.md`, `LICENSE`, `build.sh`, config files, code and scripts (`.bas`, `.py`, `.sh`), data that macros or scripts open by name, test fixtures, the picture folders, and the two exam-style specs mirrored from `gustjf/ai`).
- **Mirror outputs and sources**: generated deliverables (e.g. PDFs) go in `output/<Topic>/`; their source files (e.g. `.tex`) go in `source/<Topic>/` using the same subfolder name, so the two trees mirror each other. `./build.sh` compiles `source/<Topic>/x.tex` to `output/<Topic>/x.pdf`.
- **Keep the Open Orders macro manual current**: whenever `source/open-orders-intake-inventory/build-open-orders-report.bas` changes anything a user sees or does (setup, buttons, Config settings, how to run or build), update `source/open-orders-intake-inventory/open-orders-macro-manual.tex` in the same commit: bump the Rev. and date in its header, keep it to one page, rebuild `output/open-orders-intake-inventory/open-orders-macro-manual.pdf` with `./build.sh`, and commit both.
- **Specs live in `Specs/`**: reusable prompts/specs (e.g. `*-spec.md`) go in the `Specs/` folder. `Specs/exam-style-latex-spec.md` and `Specs/exam-style-multipage-latex-spec.md` are copies of the same files in `gustjf/ai`; when one changes, change the other.
