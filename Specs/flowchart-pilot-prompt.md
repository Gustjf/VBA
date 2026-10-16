# Flowchart Pilot: Prompt for Fable

A pilot of the presentation standard (`Specs/presentation-standard-build-spec.md`), restricted to flowcharts drawn like the system architecture drawing. Fable produces two things: one general-purpose `.tex` style guide (no example subjects) and a prompting instructions sheet. The manager follows the sheet and gives both files to Sonnet 4.6 or a GPT equivalent to make their own flowcharts.

Paste everything in the block below into a Fable session at effort high.

---

```
Read CLAUDE.md and MAP.md first and follow them.

My boss was impressed with the presentations I make at work, especially the MSA deck, and wants to make them themselves. At work they only have access to weaker models: Sonnet 4.6, or a GPT or Gemini equivalent. My boss said I am free to define a new standard if it works. It will be released on our SharePoint document control site, which meets our data use policy. I can install MiKTeX, Overleaf or similar on their PC, whatever is easiest for them.

This pilot is restricted to one thing: flowcharts, drawn as nicely as the system architecture drawing, source/open-orders-intake-inventory/open-orders-macro-flow.tex (the PDF is in output/open-orders-intake-inventory/).

The outputs are two files:
1. A .tex style guide for flowcharts: one good, general-purpose .tex file my boss can use to generate whatever flowchart they want, not a set of example subjects. It holds the house style and the building blocks (boxes, decision gates, loops, notes, title block and so on) with clear placeholders and instructions written for the model. My boss attaches or pastes it into Sonnet 4.6 (or the GPT equivalent) together with a description of their process, and the model writes a new flowchart .tex from it that compiles first time and looks as good as the system architecture drawing.
2. A prompting instructions sheet: what my boss does, step by step, from opening the chat to printing the PDF. It includes the exact prompt to paste with the style guide, what to give the model about their process, how to ask for changes, and how to fix the common problems.

Rules I care about:
- Optimize for Sonnet 4.6, and make sure a GPT equivalent can use it too. Test this as well as you can by having a weaker model follow the sheet with only the two files.
- Use AI only for what needs AI (wording, structure, judgment). Do not offload to AI anything the template or a person I can easily teach could do. Be resourceful.
- It must be safe: placeholders only in the file, no company data, nothing that could embarrass me or the company if it were forwarded. You may draw throwaway test flowcharts to prove it works, but do not ship them as content.
- It must wow my boss: whatever they generate from it should look excellent with no hand-tuning.
- Specs/presentation-standard-build-spec.md is the larger plan this pilot belongs to. Use it as background and keep this pilot compatible with it, but do not build what it lists beyond these two files.

You do the thinking: plan it, build it with subagents where it helps, compile everything and look at the rendered pages, fix what is wrong, and tell me at the end what I need to do by hand. Put source files under source/flowchart-pilot/ and generated files under output/flowchart-pilot/. Update MAP.md, commit and push per CLAUDE.md.
```
