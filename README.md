# ai-doc-converter

Convert PowerPoint dashboard slides into clean, responsive HTML that preserves
the original colors, fonts, sizes, and spacing — without absolute positioning.

## How it works

1. **Extract** — parses shapes with `python-pptx`: geometry, solid fills,
   borders, text runs (font/size/weight/color), images, and tables. Theme
   colors and fonts come from `theme1.xml`.
2. **Infer layout** — builds a containment tree from shape geometry (which
   shapes sit inside which), then groups siblings into rows by vertical
   overlap and columns by horizontal order.
3. **Emit** — semantic HTML using flexbox/grid plus a stylesheet with CSS
   variables for every color and font. Side-by-side cards become a responsive
   grid (`auto-fit, minmax(320px, 1fr)`) that stacks on mobile.

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

## Notes

- Non-web-safe PowerPoint fonts (e.g. Gadugi) are mapped to fallback stacks
  in `FONT_FALLBACKS`; extend that dict for your corporate fonts, or add
  `@font-face` rules to the generated CSS.
- Layout inference assumes dashboard-style slides: filled rectangles acting
  as cards/containers with text, icons, and tables placed inside them.
  Freeform or heavily overlapping art slides will not reflow well.
- Sizes are emitted in `rem` (16px root), so the whole page scales by
  changing the root font size; a media query reduces it on small screens.
