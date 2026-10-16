# VBA: work macros and documents

Work VBA macros and everything around them: specs, manuals, plans, test tooling. Personal projects live in the `gustjf/ai` repo.

## Open these

**Standalone macros** (repo root): [FleetFuelReport.txt](FleetFuelReport.txt), [ConvertVisioToMacroVisio](ConvertVisioToMacroVisio), [ImportDiagramsToMasterVisio](ImportDiagramsToMasterVisio), [VisioToggleLayers](VisioToggleLayers).

**Prompts and specs to paste** go in [`Specs/`](Specs/): the presentation standard build spec, the flowchart pilot prompt, and the exam-style LaTeX specs.

**Finished documents** go in [`output/`](output/), one folder per project:

| Project | Start with |
|---|---|
| Open Orders / Order Intake / Inventory report | [open-orders-macro-manual.pdf](output/open-orders-intake-inventory/open-orders-macro-manual.pdf), [macro-instructions.md](output/open-orders-intake-inventory/macro-instructions.md), then the rest of [the folder](output/open-orders-intake-inventory/) |
| ESG Reporting macros | [Excel_Macros_Doosan_Bobcat_ESG_Reporting.pdf](output/ESG%20Reporting/Excel_Macros_Doosan_Bobcat_ESG_Reporting.pdf) |
| Flowchart pilot | [flowchart-style-guide.pdf](output/flowchart-pilot/flowchart-style-guide.pdf), [flowchart-prompting-sheet.pdf](output/flowchart-pilot/flowchart-prompting-sheet.pdf) |

The Open Orders macro itself is [`source/open-orders-intake-inventory/build-open-orders-report.bas`](source/open-orders-intake-inventory/build-open-orders-report.bas) (import it into Excel).

## Behind the scenes

[`source/<Topic>/`](source/) mirrors `output/<Topic>/` and holds the `.bas` and `.tex` sources, the Python checkers, test fixtures and sample data. `./build.sh` turns `source/<Topic>/x.tex` into `output/<Topic>/x.pdf`.

[`MAP.md`](MAP.md) lists every file with a one-line description, in the same two sections.
