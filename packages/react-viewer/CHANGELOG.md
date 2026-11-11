# @extend-ai/react-pptx

## 0.5.0

### Minor Changes

- Bundle the published rust-pptx and rust-pptx-wasm 0.3.0 crates, with rust-office-art 0.1.8 and rust-office-charts 0.1.10, including corrected gradients, connector geometry, text styles, script fonts, and bullet inheritance.

  Expand EMF/WMF rendering for compound paths, curves, transforms, pen styles, brush fills, and RGB bitmap transfers. Improve font fallback and bullet layout, and expose slide-scoped parser warnings.

  Update compatible presentations in place and preserve scroll position when rebuilding, keeping the viewer stable as edited content is applied.

## 0.4.1

- Fix parsing with default options when Web Workers are unavailable, including server-side parsing. Pass the same explicit boolean defaults used by the worker path.

## 0.4.0

- Bundle rust-pptx 0.2.0 and expose document-scoped object, paragraph, table-cell and UTF-16 word references.
- Add `extractPresentationBlocks`, `flattenPresentationBlocks`, source/block/word geometry resolvers, `PptxBlockHighlights`, `PptxWordHighlights`, and `revealPptxSource`.
- Resolve citations on mounted slide content, with explicit mismatch, invisibility and unsupported-range states. Exclude thumbnails and preserve references across zoom and nested transforms.
- Add a virtualized playground Blocks panel with text, word and JSON views, type and text filters, hover highlighting, click-to-reveal, and a show-all toggle.

## 0.3.0

- Update the bundled Rust parser to `rust-pptx` 0.1.7 (`rust-office-art` 0.1.6, `rust-office-charts` 0.1.8).
- Expose the parser's presentation metadata in the model: slide, layout and master `transition` and `timing` (animation timelines and build lists), Office Math runs (`TextRun.math` with MathML), text fields (`TextRun.field`), merged table cells (`TableCell.merge`), click and hover actions, editing locks and connector endpoints, ink drawings, media playback trims, bookmarks and poster frames, embedded OLE objects and controls, modern comments, and accessibility titles, descriptions and decorative flags.
- Render Office Math runs as sanitized MathML, skip merged table cells, turn presentation jump and slide-link actions into in-viewer navigation, show media poster frames with a play control and honour trim points, offer embedded objects as downloads, and paint `fallbackPicture` artwork when the primary picture is unavailable.
- Add the hand-written `model-features.pptx` fixture and parser contract tests for the new fields.
- Remove the TypeScript chart fallback renderer. Charts are rendered to SVG by the bundled Rust parser (`rust-office-charts`), so `d3-*`, `regl`, `topojson-client`, and the TopoJSON atlas chunk are gone from the package.
- Bundle TanStack Virtual; the package now has no runtime dependencies beyond its React peers.
- Accept picture fills, `<desc>` notes, markers, and filter primitives in parser-rendered chart SVG instead of dropping to a placeholder.
- Chart nodes without parser SVG fall back to the built-in series placeholder, whose bars and lines now carry `data-rpv-chart-series` / `data-rpv-chart-point` attributes.
- Excel map charts render from the geometry PowerPoint saves in the chart part (`cx:geoCache`); a map chart saved without it shows the placeholder instead of the former TopoJSON approximation.

## 0.2.5

- Refresh the shared Rust parser and OfficeArt renderer for cached SmartArt shapes, pictures, effects, and charts.
- Render saved SmartArt text frames independently, preserve authoritative cached labels, and resolve cached drawings across relationship scopes.
- Add the opt-in `include3dModels` parser option for original 3D model assets and scene metadata while retaining cached preview rendering.

## 0.2.3

- Use the Rust Office crates by Extend and refresh the bundled WebAssembly.

## 0.2.2

### Patch Changes

- f4cb3b0: Fix paragraph styling in placeholders and render PowerPoint's default table style.

  - Slide paragraphs no longer take their `lvl`/`marL` from the layout placeholder's prompt paragraphs ("Click to edit Master text styles", "Second level", …). They inherit only through `a:lstStyle`, which stops plain bullet lists from stair-stepping.
  - A placeholder without a `type` (e.g. `<p:ph idx="1"/>`) defaults to `obj` per ECMA-376 and now uses the master's `bodyStyle` instead of `otherStyle`, restoring its bullets, indents and size.
  - When `ppt/tableStyles.xml` doesn't define the requested style, the built-in "Medium Style 2 - Accent 1" (`{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}`, PowerPoint's default) is used, so its fills, banding and borders render.

## 0.2.1

### Patch Changes

- Add first-class controlled and uncontrolled responsive zoom modes, resize-aware resolved zoom,
  and controller APIs for consumer-owned zoom controls while preserving `fitMode` compatibility.

## 0.1.1

### Patch Changes

- Improve PowerPoint import fidelity for word spacing, hyperlink styling, diagonal stripe shapes, and straight connectors, including zero-width vertical timeline stems.

## 0.3.0

- Add a cold-start `createPptxThumbnailRenderer` API with package-owned PNG, canvas, ImageBitmap,
  and SVG outputs, explicit thumbnail font policy, cancellation, bounded concurrency, shared
  presentation assets, and phase-level performance measurements.
- Make slide navigation observable before off-window rendering completes, isolate resource
  readiness per slide mount, update zoom and fit geometry without rebuilding mounted slide DOM,
  and reduce thumbnail work outside the visible rail.
- Expose consumer-owned slide thumbnail rails through `usePptxViewerThumbnails`, including stable
  container refs, resolution bounds, render status, imperative rendering, and cleanup. Add
  DOCX-style visible/prefetch render windows and virtualize the built-in thumbnail scrollport.

## 0.1.0

Initial public release of the React PowerPoint viewer, including:

- Browser-native PPTX and legacy PPT parsing through Rust and WebAssembly.
- A normalized presentation model and virtualized React viewer.
- Embedded and host-provided font support with missing-font diagnostics.
- Native legacy PowerPoint slide, drawing, text, and image fallback support.
- Fidelity improvements for placeholders, backgrounds, text layout, gradients, and image fills.
