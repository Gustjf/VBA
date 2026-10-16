# Flowchart pilot: test report (2026-10-14)

What was tested: `source/flowchart-pilot/flowchart-style-guide.tex` (Rev. 1) and `flowchart-prompt.txt`, used exactly as the prompting sheet says, by stand-in models with no tools and nothing but the pasted prompt and the attached style guide. All process notes were synthetic; no test chart is shipped.

## Stand-in model runs (proxy for Sonnet 4.6 and a GPT equivalent)

| Run | Model | Notes given | Engine copied verbatim | Compiled first time | Page quality |
|---|---|---|---|---|---|
| A | Sonnet | Expense claim: 4 steps, decision, policy file, loop, before-you-start | yes | yes | good; chose a key box and a loop, explained one judgement call in one line |
| A | Haiku | same | yes | yes | good; 5 bullets in one station, still fits |
| B | Sonnet | Monthly report build: three input files, macro, check, review, publish, mapping table | yes | yes | good; document stations at both ends, side document, gate with warning |
| C | Sonnet | Supplier onboarding: long messy notes, decision at the end | yes | yes | good; six stations, two gates, side document, remark; trimmed the notes to the limits |
| C | Haiku | same | yes | yes | good; role lead lines, two gates, loop; one wrapped outcome heading (engine now wraps it) |
| D | Haiku | Visitor sign-in: three lines of notes | yes | yes | good; sparse by nature (3 stations); added one inferred bullet |

Every run returned one complete file from `\documentclass` to `\end{document}`, no invented commands, no questions back. Haiku (weaker than Sonnet 4.6) is the floor; GPT and Gemini were not run here, so the owner's acceptance test with the real work tools is still the gate.

## Engine checks (drawn by the owner's session, throwaway)

- 2, 3, 4, 5 and 6 stations; stations with document symbols; empty tags; empty second summary column.
- Gate under a middle station with a side document feeding a later station (the reference drawing's shape); gate with no side document; two gates.
- Side document feeding the same, an earlier and a later station; two side documents; one and two loops; loop from a station that already has an arrow entering it (offset).
- Overlong content: the page shrinks by up to 45 % and otherwise stamps TOO TALL in red.
- Rule violations (gate under the last station, gates under neighbouring stations, side document under a gated station, loop from a gated station, station numbers that do not exist): compile still succeeds, each violation is a red line at the bottom left naming the fix.
- Characters `& _ # $ ^` typed as is, `\%`, smart quotes, en and em dashes, accents, `€`, arrows, ticks and crosses: all render. A bare `%` or an emoji still breaks the build (named in the sheet's problem table).

## Open points for the owner

- Run the prompting sheet once on the manager's PC with their real work AI (Copilot or Gemini) and the MiKTeX install; bring any failure back as a change to Part 1 of the style guide or to the sheet.
- The style guide is one file by design (pasteable); when the full presentation standard is built, Part 2 becomes the shared `.sty` and the content commands stay the same.
