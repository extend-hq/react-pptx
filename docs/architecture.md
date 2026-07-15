# Architecture

The package deliberately separates untrusted file parsing, stable public data, and rendering.

1. `pptx-core` checks magic bytes and resource ceilings before opening ZIP or OLE containers.
2. PPTX parsing resolves the presentation relationship graph, slide order and dimensions,
   core shapes/text/images, binary assets, and metadata into `PresentationDocument`.
3. Legacy PPT parsing opens OLE compound storage and resolves the active binary presentation:
   - `Current User` identifies the newest `UserEditAtom`.
   - The bounded edit chain merges full and incremental persist-object directories, with newer
     entries replacing older ones.
   - The live `DocumentContainer` and `SlideListWithText` resolve source slide order and prevent
     stale slides or master placeholder text from becoming user-visible slides.
   - Referenced master drawing content is normalized before slide-local content. Supported
     OfficeArt client anchors, child anchors, and group coordinate spaces become EMU transforms
     in the shared `PresentationDocument` model.
   - The OfficeArt blip store and OLE `Pictures` stream provide image assets. PNG/JPEG data can
     render directly; decoded EMF/WMF data is converted to a browser image by the React renderer.
4. `pptx-wasm` exposes the Rust model to browsers. The npm runtime loads it lazily and permits
   a custom URL, response, bytes, or compiled module.
5. Both PPTX and legacy PPT parsers produce the same public normalized model. There is no second,
   opaque rendering representation or format-specific React surface.
6. The repository-owned React renderer resolves assets asynchronously before completing a slide render.
   It renders PNG/JPEG assets directly and uses a bounded browser-side conversion for EMF/WMF
   static fallbacks. No LibreOffice process, server conversion, or native desktop dependency is
   part of the runtime path.
7. React owns document lifecycle and slide-level state. The package-owned renderer creates the
   HTML/SVG slide tree, and each node maps directly back to the public model so rendering and
   diagnostics remain inspectable.

Continuous mode mounts a window around the viewport. Single-slide mode replaces the surface.
Both modes dispose observers, blob URLs, asynchronous assets, and highlights on replacement.

Both parsers converge on the same public model and viewer controller, but fidelity is format-
specific. The legacy path currently targets static viewing of common PowerPoint 97–2003 decks.
Advanced OfficeArt effects, editable binary charts and tables, animations, transitions, active
OLE objects, and media timelines may be omitted or represented by a static picture. Every known
degradation is surfaced through structured warnings.
