# ai-doc-converter

Convert PowerPoint dashboard slides into clean, responsive HTML that preserves
the original colors, fonts, sizes, and spacing — without absolute positioning.

## How it works

1. **Extract** — parses shapes with `python-pptx`: geometry, solid fills,
   borders, text runs, images, tables, and connectors. Theme colors resolve
   through the master color map (`clrMap`), and text styles missing on a run
   are inherited the OOXML way: layout placeholder → master placeholder →
   master `txStyles`. The slide background resolves slide → layout → master.
2. **Infer layout** — builds a containment tree from shape geometry, then
   groups siblings into rows by vertical overlap and columns by horizontal
   order.
3. **Emit** — semantic HTML using flexbox/grid plus a stylesheet with CSS
   variables. Side-by-side cards become a responsive grid that stacks on
   mobile; column widths follow the original shape proportions.

Connector shapes are handled by intent: lines **with arrowheads** become SVG
flow arrows (with nearby text labels attached, rotating vertical on mobile);
lines **without arrowheads** become `<hr>` dividers. Slide-number, footer,
and date placeholders are skipped.

## Usage

```bash
pip install -r requirements.txt
python -m ai_doc_converter dashboard.pptx -o html_out [--slide N]
```

Output:

```
html_out/
  index.html    # semantic, responsive markup
  styles.css    # CSS variables for colors/fonts + generated classes
  assets/       # images extracted losslessly from ppt/media
```

`examples_output.zip` contains three converted reference decks
(operating_model, cit_mit_flow, toc_dark).

## Notes

- Non-web-safe PowerPoint fonts (e.g. Gadugi) are mapped to fallback stacks
  in `FONT_FALLBACKS`; extend that dict for your corporate fonts, or add
  `@font-face` rules to the generated CSS.
- Layout inference assumes dashboard-style slides: filled rectangles acting
  as cards/containers with text, icons, tables, and flow arrows. Freeform or
  heavily overlapping art slides will not reflow well.
- Gradient fills, images-as-background, and per-level (lvl2+) list styles are
  not yet resolved; vertical divider lines are dropped (they don't survive
  responsive reflow).
