# PowerPoint → standalone HTML → Figma Slides

This fork now has a fidelity-first path for PowerPoint templates whose original
layout should survive the trip to Figma Slides as closely as possible.

## Why the original conversion loses layout

`scripts/extract-pptx.py` intentionally extracts slide content, not rendering
geometry. It keeps text strings, image files, and notes, but it does not retain
all PowerPoint coordinates, text-box margins, placeholder geometry, master
layout visuals, typography, and drawing styles. That is useful for redesigning a
presentation, but not for faithfully transporting a template.

The generated `KU_DRI_LAB_PPT_Template_16_9.html` also points at a sibling
`KU_DRI_LAB_PPT_Template_16_9-assets/` directory. Moving or opening only the HTML
therefore breaks the images.

## Fidelity-first pipeline

The new path is:

```text
PPTX
  ↓
scripts/pptx-to-figma-html.py
  ├─ PowerPoint artwork → text-stripped PDF → per-slide SVG
  ├─ text + placeholders → positioned HTML + manifest metadata
  └─ everything embedded into one *.figma.html
  ↓
Frontend Slides PPTX Fidelity Importer (local Figma Slides plugin)
  ├─ SVG → Figma vectors
  └─ text/placeholders → editable Figma text
  ↓
Figma Slides
```

This is deliberately hybrid. Rendering non-text artwork keeps visual fidelity,
while reconstructing text separately keeps the important template fields
editable.

## Install converter dependencies

### macOS

```bash
brew install --cask libreoffice
brew install poppler
python3 -m pip install python-pptx
```

### Ubuntu / Debian

```bash
sudo apt-get install libreoffice poppler-utils fonts-noto-cjk
python3 -m pip install python-pptx
```

## Convert the included KU DRI Lab template

```bash
python3 scripts/pptx-to-figma-html.py \
  KU_DRI_LAB_PPT_Template_16_9.pptx \
  --output KU_DRI_LAB_PPT_Template_16_9.figma.html
```

The output HTML is **standalone**. Unlike the older hand-authored HTML, it does
not require `KU_DRI_LAB_PPT_Template_16_9-assets/` next to it.

Open it directly in a browser to inspect the result. Add `?figma=1` to show all
slides vertically for comparison:

```text
file:///.../KU_DRI_LAB_PPT_Template_16_9.figma.html?figma=1
```

The HTML also contains an embedded `<script id="figma-deck-manifest">` JSON
block used by the Figma plugin.

## Import into Figma Slides

Use Figma Desktop:

1. Open or create a **Figma Slides** file.
2. Open **Plugins → Development → Import plugin from manifest…**.
3. Select `figma-plugin/pptx-fidelity-importer/manifest.json`.
4. Run **Frontend Slides PPTX Fidelity Importer**.
5. Choose the generated `*.figma.html`.
6. Click **Import slides**.

Each source PowerPoint slide becomes a native Figma `SLIDE`. Static artwork is
imported as SVG vectors, and text/template placeholders are recreated as
editable text layers. The importer adds slides and does not delete existing
slides.

## Fidelity tradeoffs

This path targets layout fidelity rather than perfect semantic PowerPoint
round-tripping.

Expected to preserve well:

- slide aspect and absolute placement;
- slide/master/layout artwork that LibreOffice can render;
- images and logos without external HTML asset paths;
- fills, borders, basic diagrams and vector geometry;
- normal text boxes and common placeholders;
- editable text in Figma Slides.

May still require adjustment:

- a font that is not installed/available in Figma;
- highly customized grouped text transforms;
- SmartArt/chart-internal labels;
- PowerPoint animations/transitions;
- audio/video;
- Office-only effects whose rendering differs in LibreOffice.

For a template like `KU_DRI_LAB_PPT_Template_16_9.pptx`, this route is much more
appropriate than asking an agent to infer a new HTML layout from extracted
content, because the source layout itself is treated as authoritative.

## Validation

`.github/workflows/test-pptx-figma-fidelity.yml` converts the checked-in KU DRI
Lab template on GitHub Actions and verifies that:

- a standalone HTML file is produced;
- all slides contain embedded SVG artwork;
- editable text metadata exists;
- the result contains no references to the old external assets directory.

The workflow uploads the converted `KU_DRI_LAB_PPT_Template_16_9.figma.html` as
an Actions artifact for visual inspection.
