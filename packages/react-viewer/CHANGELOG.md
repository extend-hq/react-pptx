# @extend-ai/react-pptx

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

## Unreleased

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
