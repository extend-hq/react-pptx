# OpenDocument support

**OpenDocument support is experimental.** Rendering fidelity varies, especially for chart
layouts and custom drawing geometry. These formats are view only.

OpenDocument presentations and drawings are parsed by the same Rust/Wasm pipeline as PowerPoint
files and rendered from the same `PresentationDocument` model. The runtime does not invoke
LibreOffice, upload the file to a conversion service, or first rewrite it as `.pptx`.

## File types

| Extension | Kind | `document.format` |
| --- | --- | --- |
| `.odp` | Presentation | `odp` |
| `.otp` | Presentation template | `odp` |
| `.fodp` | Presentation as one flat XML file | `odp` |
| `.odg` | Drawing | `odg` |
| `.otg` | Drawing template | `odg` |
| `.fodg` | Drawing as one flat XML file | `odg` |

The format is detected from the bytes: the `mimetype` entry of the package, or the
`office:mimetype` attribute of a flat XML file. File names are not consulted. OpenDocument text
and spreadsheet files are rejected with an `unsupported-format` error, and encrypted files with
`encrypted-document`.

A drawing's pages are shown as slides.

## What is drawn

Page size, slide and master backgrounds, text boxes and placeholders with their formatting,
bullets and numbering, basic shapes, custom shapes that correspond to a preset, paths and
polygons, lines and connectors, rotation, groups, images, tables with merged cells, solid,
gradient and bitmap fills, outlines and speaker notes. Master-page decorations appear on every
slide, and the master's footer, slide-number and fixed-date fields are filled in on slides that
display them.

Supported custom shapes evaluate enhanced-geometry formulas and paths. Embedded charts use
a saved preview when available; otherwise, supported charts render to SVG from their data.
Chart styles, scatter/bubble domains and cached values are retained for downstream extraction.

## What is not

- A custom shape whose geometry cannot be evaluated falls back to an approximation, with a
  `shape-geometry-approximated` warning.
- Chart layouts and custom geometry do not yet fully match the originating application.
  Unsupported charts may use a saved preview; chart data can remain available even when
  the chart cannot be drawn.
- Audio and video, animations, transitions and slide comments are not read.
- Image cropping and hatch fills are ignored.

## Editing

OpenDocument files are view only. The playground's editor saves changes into the original PPTX
package, which these files do not have.

The parser side is documented in Rust Office's
[OpenDocument input](https://github.com/extend-hq/rust-office/blob/main/docs/opendocument.md).
