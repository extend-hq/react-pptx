---
'@extend-ai/react-pptx': minor
---

Open OpenDocument presentations and drawings (`.odp`, `.otp`, `.fodp`, `.odg`, `.otg`, `.fodg`). They are detected from their bytes, parsed by the bundled Rust parser into the same presentation model as PowerPoint files, and reported as the new `odp` and `odg` values of `PresentationFormat`. They are view only. See `docs/opendocument.md` for what is and is not drawn.

`PresentationFormat` gains `'odp'` and `'odg'`; code that switches on it exhaustively needs the new cases. `isPresentationFormat` and `PRESENTATION_FORMATS` are exported for checking a format value.

Bundle the published rust-pptx and rust-pptx-wasm 0.4.0 crates, which read these formats through the new rust-odf 0.1.0 crate.
