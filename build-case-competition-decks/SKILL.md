---
name: build-case-competition-decks
description: Build or redesign MBA and business case competition presentations as editable PowerPoint (.pptx) files from a supplied playbook, case solution, research pack, or slide draft. Use for competition decks, shortlisting submissions, finalist presentations, and case solution slides requiring precise readable content, clearly divided layouts, content-specific diagrams, evidence, implementation, and economics. Also use to emulate the structural strengths of supplied reference decks without copying their density. This skill packages an existing argument into a presentation; use a strategy skill separately if the user needs the case solved first.
---

# Build Case Competition Decks

Produce a persuasive, self-contained, editable PowerPoint. Treat the playbook as the content authority, the competition brief as the constraint authority, and reference decks as design evidence. Preserve the mechanism and proof while removing repetition. Deliver the actual `.pptx`, not only an outline, script, PDF, or slide images.

## Read the right resources

- Read [content-and-story.md](references/content-and-story.md) before extracting and compressing a playbook.
- Read [visual-system.md](references/visual-system.md) and [slide-archetypes.md](references/slide-archetypes.md) before storyboarding. Choose a layout because it fits the argument.
- Read [editable-powerpoint.md](references/editable-powerpoint.md) before building. Use the optional native-object helpers in `scripts/pptx_primitives.py` or another available PowerPoint library.
- Read [quality-gates.md](references/quality-gates.md) before final review. Run `scripts/audit_pptx.py` and visually inspect the rendered deck.
- Consult [reference-deck-analysis.md](references/reference-deck-analysis.md) when interpreting visual preferences or checking why a rule exists. It records a critical examination of 39 supplied PDF pages and two organizer-linked winning decks; these are references, not proof of universal winning rules.

## 1. Establish the contract and read every input

Identify the playbook, case brief, rubric, slide cap, time limit, audience, mandatory sections, supplied branding, and any template. Follow an explicit slide cap, including whether cover and appendix count. Keep mandatory content visible; speaker notes cannot substitute for required submitted content.

If unspecified, use a 16:9, balanced competition submission with roughly 8–12 core slides, adjusted to the argument. State the assumption briefly and proceed. Use a more spacious live-presentation mode when requested. Do not delay for cosmetic choices.

Extract text and tables from all relevant files. For PDF or PPT references, also render and inspect the slides: text extraction does not reveal hierarchy, diagrams, density, or clipping. Inspect existing PPT objects when editable originals are supplied; a PDF cannot establish original object editability or animations. Do not claim otherwise.

## 2. Build a source ledger and a decision story

Create a working ledger with `claim ID | exact meaning/value | unit/base/period | source locator | status | destination`. Distinguish **case fact**, **external evidence**, **team inference**, **assumption**, **proposal/target**, and **unknown**. Reuse claim IDs across chart data, slides, and appendix.

Extract the actual decision, root causes, selected recommendations, mechanisms, owners, rollout, economics, risks, and success gates. Map rubric requirements to slides. Identify contradictions and material gaps before dressing them up. Recalculate derived values; never quietly choose the most attractive version. Preserve named programmes and approved positions unless editing them is authorized.

Write the answer in one sentence, then the slide headlines in order. A judge reading only those headlines should understand what to do, why it fits this case, how it works, and when it merits scaling. Use claim-led titles when supported; use an honest question or proposed action when evidence is incomplete.

## 3. Write slide copy before laying it out

For each slide specify:

`purpose / judge question / takeaway title / 2–4 named zones / visible copy / visual relationship / source IDs / notes / likely objection`.

Give each zone a distinct job: evidence, mechanism, implication, decision, or control. Use a short meaningful heading and one or two substantive sentences or two or three short bullets. Keep actor, action, scope, condition, and measurable result where supplied. Do not reduce content to vague labels such as “AI-driven synergy” or remove essential qualifiers to fit.

Use **130–220 visible body words as a working range for a substantial read-through slide**, not a quota. Prefer **60–120 for live presentation**. Exclude navigation and source footer from that estimate; include table and diagram text. Visual or financial slides may need fewer words. If a slide exceeds the range, examine its hierarchy; do not automatically shrink text. Aim for 20–45 words per explanatory zone and 4–10 words per diagram node. Split, simplify, or move supporting detail to an allowed appendix before reducing body font.

Never invent interviews, survey responses, benchmarks, quotes, causal effects, or financial returns. Label targets and illustrative personas visibly. External research may validate or supplement a playbook when authorized; it must not silently replace its case logic. Do not import historical reference-deck statistics into the new case.

## 4. Design a coherent visual system

Choose one case-relevant palette, one or two available font families, and a stable grid. Keep a restrained section indicator, clear title, organized content area, and a short source line. Use 2–4 major zones and generous separation; make the most important evidence or mechanism visually dominant.

Vary layouts across the deck: evidence chart, comparison, process, persona, operating model, roadmap, economics, and decision. Do not use the same three-card layout repeatedly. For six or more content slides, aim for at least four suitable layout families; do not add meaningless variety to satisfy a count.

Use diagrams that expose an actual relationship: sequence, handoff, boundary, cause, trade-off, funnel, allocation, feedback, or timing. Create at least one distinctive case-specific schematic when warranted. Label connectors and decision gates where meaning is not obvious. Use a cycle only for a real feedback loop, a funnel only for genuine narrowing, and equal cards only for comparable items.

Use a consistent icon or sticker family as a navigation aid or explanatory cue. Prefer editable shape-based icons; allow images for photos, logos, and illustrations while keeping business text and diagram labels editable. Avoid ornamental gears, trophy stickers, excessive logos, neon highlights, and decorative flows with no semantic role. Do not reproduce unrelated reference-team branding or portraits.

## 5. Build with native PowerPoint objects

Use native text boxes, shapes, connectors, tables, and editable charts with embedded data. Name important objects logically; group components when practical without flattening them. Keep typography and colour tokens consistent. Use explicit font sizes, manual layout, and sufficient text-box height. Never rasterize a whole slide, table, chart, or business diagram to hide layout problems.

Put full source locators, assumptions, calculations, and presenter explanations in speaker notes; keep decision-critical qualifications on the slide. Provide an appendix when allowed for detailed evidence and calculations. Label projections and proposals in charts as well as prose.

Build two or three representative slides first—one evidence slide, one mechanism diagram, and a dense comparison or economics slide. Render them and fix the system before completing the deck. Continue without requesting approval for routine reversible design choices.

## 6. Verify, revise, and deliver

Reconcile repeated numbers across main slides, native chart data, notes, and appendix. Check totals, denominators, percentage points versus percent, cost periods, and ROI versus benefit–cost ratio. Check that every recommendation retains its case-specific mechanism and that the last substantive slide asks for a concrete decision.

Run the structural audit, render all slides, inspect a contact sheet for rhythm and each slide at readable size for clipping, overlap, alignment, contrast, awkward wrapping, spelling, and missing glyphs. A successful save or clean lint report is not visual validation. Review tiny text and crowded tables manually. Fix and re-render affected slides.

Confirm editability by reopening the PPTX, inspecting object types, and changing sample text, a diagram shape, a table cell, and chart data in a copy where present. Confirm the copy still opens/renders. Do not claim Microsoft PowerPoint testing unless actually performed. If rendering is unavailable, say so explicitly and provide the editable deck with a limited verification statement.

Save the final deck using the environment's durable artifact workflow, then provide a clickable `.pptx` link and a concise account of the result and any unresolved content issue. Include a PDF preview or source files only when requested or useful. Keep temporary renders and audits out of final deliverables unless requested.
