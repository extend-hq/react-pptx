# @extend-ai/react-pptx

## Unreleased

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
