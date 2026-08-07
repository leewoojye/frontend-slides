# Frontend Slides PPTX Fidelity Importer

This development plugin imports the self-contained `*.figma.html` produced by
`scripts/pptx-to-figma-html.py` directly into **Figma Slides**.

## Why this exists

Figma Slides can import `.pptx` natively, but PowerPoint imports can substitute
fonts and omit unsupported objects. The fidelity converter takes a different
route:

1. render non-text PowerPoint artwork to SVG at the source slide geometry;
2. keep PowerPoint text boxes/placeholders as explicit positioned metadata;
3. embed both into one standalone HTML file;
4. let this plugin create one native Figma `SLIDE` per PowerPoint slide;
5. import SVG artwork as vectors and text/placeholders as editable Figma text.

The plugin does **not** delete or rewrite existing slides.

## Install as a local development plugin

Figma Desktop is required for development plugins.

1. Open or create a Figma Slides deck.
2. Open **Plugins → Development → Import plugin from manifest…**.
3. Select this folder's `manifest.json`.
4. Run **Frontend Slides PPTX Fidelity Importer** from the Development plugins list.

If Figma reports that the manifest ID is invalid in your environment, create a
blank local Slides plugin once and copy the generated `id` into this
`manifest.json`; the source files need no other changes.

## Import

Generate the HTML first:

```bash
python scripts/pptx-to-figma-html.py presentation.pptx \
  --output presentation.figma.html
```

Then run the plugin and choose `presentation.figma.html`.

The same HTML can be opened directly in a browser. It contains no external image
paths; slide artwork is embedded as inline SVG.

## Fidelity notes

- Static artwork, images, borders, fills, and most PowerPoint geometry are kept
  in an SVG background imported as editable vectors.
- Text boxes and empty PowerPoint placeholders are imported as editable text.
- When the original PowerPoint font is unavailable to Figma, the importer tries
  the requested family/style first and falls back to Inter.
- Complex grouped text, chart-internal labels, SmartArt, video, animation, and
  some Office-only effects can still differ. The SVG background remains the
  visual source of truth for those non-text objects.
