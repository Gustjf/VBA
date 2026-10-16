# Compiling the ESG Reporting .tex

Source file: `source/ESG Reporting/Excel_Macros_Doosan_Bobcat_ESG_Reporting.tex`

## 1. Check whether LaTeX is already installed
Open Command Prompt and run:
```
pdflatex --version
```
If that prints a version number, go to step 3.

## 2. Install MiKTeX (Windows)
- Download it from **miktex.org**. Choose **"Install for me only"**, which usually doesn't need admin rights.
- In MiKTeX Console, set **"Install missing packages on-the-fly"** to **Yes**. The first build will then fetch the packages the file uses (`lmodern`, `titlesec`, `lastpage`, `hhline`, and so on) by itself.
- If your PC is locked down, ask IT to install **MiKTeX** or **TeX Live**. Both are free and widely used.
- **Overleaf** (a web-based LaTeX editor) needs no install: upload the `.tex` and click Recompile. Check first whether company content is allowed on outside sites; this document names Doosan Bobcat and describes internal reporting.

## 3. Build the PDF
Put the `.tex` in its own folder and rename it without parentheses, e.g. `esg-macros.tex`, so it's easier to type. Then run in that folder:
```
pdflatex esg-macros.tex
pdflatex esg-macros.tex
```
Run it **twice**. The first pass works out the page count and the section and table numbers; the second pass fills them in. With only one pass you'll see "Page 1 of ??" or "Section ??".

## 4. Check the result
Open `esg-macros.pdf` and confirm it's one page with no "??". You can delete the extra files the build leaves behind (`.aux`, `.log`, `.out`).

## If something goes wrong
- **"File `xyz.sty` not found":** a package is missing. Turn on the MiKTeX on-the-fly install setting, or install it from MiKTeX Console.
- **The build stops at a `?` prompt:** type `x` and press Enter to quit. The error is in the `.log` file, at the line that starts with `!`.
- **To edit and build with one click:** install **TeXworks** (included with MiKTeX) or VS Code with the LaTeX Workshop extension.
