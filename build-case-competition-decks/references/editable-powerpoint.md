# Build and verify editable PowerPoint

## Native-object contract

| Element | Required implementation |
|---|---|
| Titles, body, labels, captions, sources | Native PowerPoint text |
| Business diagrams, frameworks, process flows | Native shapes and connectors, with native text |
| Comparison/financial/roadmap tables | Native tables or individually editable cells and labels |
| Quantitative charts | Native chart with embedded data workbook when supported |
| Photos, logos, complex illustrations | Images allowed; keep captions editable |
| Complete slide | Never a single screenshot, PDF page, or background picture containing the content |

An SVG is not equivalent to individually editable PowerPoint shapes. Use SVG/PNG only for permitted visual assets, not as a workaround for the business diagram requirement. A PDF preview accompanies, but does not replace, the PPTX.

## Choose the runtime

Inspect installed libraries and available presentation tools first. Use python-pptx, PptxGenJS, or a suitable native presentation API. Do not assume a package or renderer exists. Use runtime-owned dependency environment variables where provided; never overwrite system variables.

The bundled `scripts/pptx_primitives.py` provides optional `Deck` and `Theme` classes using python-pptx. It supplies a base, text, panels, connectors, metric callouts, a short flow, a native table, native bars, simple shape icons, and notes. It is a starting vocabulary, not an automatic design system or a guarantee of fit. Write case-specific layouts and adapt sizing; do not force every slide through `panel()`.

Example from a task-local script (resolve the installed skill root first):

```python
from pathlib import Path
import sys
skill_root = Path('/absolute/path/to/build-case-competition-decks')
sys.path.insert(0, str(skill_root / 'scripts'))
from pptx_primitives import Deck, Theme

d = Deck(Theme(primary='1D4F91', accent='007C83'))
s = d.base('Target the largest verified service gap first',
           section='Diagnosis', source='Source: supplied playbook, table 2.')
d.bar_chart(s, ['Region A', 'Region B'], {'Current rate': [.22, .12]},
            .50, 1.80, 7.50, 4.40, number_format='0%')
d.panel(s, 'Implication',
        'Use the verified regional gap to choose the first pilot. Validate its cause before scaling the same intervention elsewhere.',
        8.35, 1.85, 4.40, 3.60)
d.notes(s, 'Illustrative example only. Replace values with sourced playbook data.\n'
           'Claim IDs: C01, C02. Preserve base, period, and uncertainty.')
d.save('example.pptx')
```

Never use the example's numbers as case facts. Keep actual chart arrays, repeated totals, and appendix calculations tied to the same source ledger or task-local data object.

## Build discipline

1. Define slide size, theme tokens, font roles, and safe content rectangle.
2. Write all visible copy in the storyboard before positioning objects.
3. Build representative chart, diagram, and table slides; render them.
4. Complete the deck with native objects, named components, and meaningful reading order.
5. Put long references and calculation workings in notes; keep critical source qualifiers visible.
6. Avoid auto-shrink. If a box cannot fit at a readable size, change its content or layout.

Name objects with useful roles such as `title:`, `heading:`, `body:`, `label:`, `diagram:`, `table:`, `chart:`, `footer:`, and `nav:`. The audit uses these prefixes to apply different font heuristics. Do not relabel body text as footer to evade warnings. With another builder, equivalent native structures and honest manual validation are acceptable.

Use actual connectors where practical and attach them to relevant shapes. Keep connection lines clear of labels. Add shape objects in a logical reading order; inspect ordering/accessibility in PowerPoint when available. Verify local font availability and substituted glyphs, including ₹, %, arrows, and en dashes. Do not claim fonts are embedded unless actually embedded and licensed.

Check fit within each containing panel, not only within the slide. Even a 100-word slide can overflow if narrow nodes or shallow boxes force additional lines. Leave roughly 10–15% vertical slack in explanatory text boxes, measure with the chosen font and weight, and inspect long words inside diamonds. Keep connector labels separate from body text. Remove inherited theme shadows/effects unless deliberately used; native objects can inherit styling that was not specified in code.

## Render and inspect

Use a provided presentation renderer if available. Otherwise an installed LibreOffice/soffice renderer can produce a PDF:

```bash
soffice -env:UserInstallation=file:///tmp/case-deck-render-profile --headless --convert-to pdf --outdir /absolute/preview /absolute/deck.pptx
pdftoppm -scale-to 1600 -png /absolute/preview/deck.pdf /absolute/preview/slide
```

Use a unique temporary profile/output directory for concurrent builds. Check command exit status, actual PDF existence, page count, and modification time; a stale PDF is not a successful render. A PDF library can also rasterize the rendered file and make a contact sheet. Do not send every intermediate image to the user.

Inspect every slide at readable resolution and the contact sheet as a whole. Check nested diagrams and small text separately. Renderer differences can exist; name the renderer in the verification statement if relevant, and do not claim testing in Microsoft PowerPoint unless it occurred.

## Structural and editability audit

```bash
python /absolute/skill/scripts/audit_pptx.py /absolute/deck.pptx --mode read --json /absolute/audit.json
```

The script checks ZIP/XML integrity, internal relationship targets, top-level bounds, approximate word density, font roles, notes presence, picture dominance, native object counts, and potential text overflow. It inventories charts and workbooks. Exit 1 means structural errors; exit 0 can still include warnings. Review warnings individually. Its text-fit estimate is heuristic; it cannot establish semantic accuracy, overlap, contrast, or visual quality. Tables, icons, outlines and intentional containment can make generic overlap tests misleading; inspect the rendered result.

In a temporary copy, change one title/body text, one native diagram shape, one table cell, and a chart's data where those elements exist. Reopen and render the copy. Inspect the PPTX package for native chart XML and its workbook relationship. The test copy is not a deliverable. Do not claim that a chart is editable merely because a decorative chart image sits beside a text label.

## Technical sources

- [python-pptx chart documentation](https://python-pptx.readthedocs.io/en/latest/user/charts.html): native chart creation and chart data, formatting, axes and labels.
- [python-pptx notes documentation](https://python-pptx.readthedocs.io/en/latest/user/notes.html): access speaker notes through the notes text frame.

Documentation checked 6 October 2026; helpers exercised against installed python-pptx 1.0.2. Verify capabilities in the current environment instead of assuming every chart type or version supports the same operations.
