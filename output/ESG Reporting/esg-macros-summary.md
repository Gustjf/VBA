# ESG Reporting Macros — Summary

1. **What it is:** a one-page procedure for installing and running the three ESG reporting macros: Energy, Fleet Fuel, and Greenhouse.
2. **Why LaTeX rather than Word:** the layout is set by a written style guide, so every document comes out looking the same, and the plain-text source can be kept under version control. Changing wording doesn't break the formatting.
3. **What changed:**
   - The page is easier to read: slightly more space between lines and left-aligned table cells.
   - It now looks like a controlled document: a "Strategy" header, a revision number with date, and "Page X of Y".
   - Section, step, and table references update themselves when steps change.
4. **Document control:** the header shows "Rev. 1 2026-10-09". Each change that gets issued raises the revision number, and the author is in the footer.
5. **Reuse:** the same style guide (`Specs/exam-style-latex-spec.md`) can produce other one-page procedures in the same format.
6. **Open questions:**
   - Is "Strategy" the right header label?
   - Should the right side of the footer carry a document number?
   - Who approves new revisions?
