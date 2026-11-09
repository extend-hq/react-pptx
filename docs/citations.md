# Source-grounded citations

React PPTX can resolve blocks produced by the matching Rust parser into bounding boxes on mounted slides. Extraction is opt-in and does not require a second parse of the presentation bytes.

```tsx
import {
  ReactPptxViewer,
  PptxBlockHighlights,
  parsePresentation,
  extractPresentationBlocks,
  flattenPresentationBlocks,
} from '@extend-ai/react-pptx';

const parsed = await parsePresentation(file);
const structured = await extractPresentationBlocks(parsed.document, { includeWords: true });
const blocks = flattenPresentationBlocks(structured.blocks);
const selected = blocks.find(block => block.type === 'text');

<div ref={rootRef}>
  <ReactPptxViewer source={parsed} onReady={setController} />
  <PptxBlockHighlights rootRef={rootRef} blocks={selected ? [selected] : []} />
</div>
```

`file`, `rootRef` and `setController` are supplied by the host. Run the asynchronous extraction outside render, keep the parsed model and structured blocks in state, and cancel obsolete loads when the source changes. `rootRef` should enclose one viewer; thumbnails are excluded from citation resolution. `PptxBlockHighlights` accepts a block or word as `block`, or multiple targets through `blocks`. It refreshes after mounting, scrolling, resizing, font loading and zoom. It does not mount every slide in a virtualized deck.

Use `await revealPptxSource(controller, root, sourceRef)` to mount and navigate to a citation's slide, then resolve or highlight it. The returned geometry reports whether the target resolved. For persistent multiple-slide overlays, use continuous mode and handle `not_mounted` until the relevant slides mount.

For custom overlays, `resolvePptxSourceGeometry(root, refs)`, `resolvePptxBlockGeometries(root, blocks)` and `resolvePptxWordGeometries(root, words)` expose geometry without rendering a highlight. The latter two return maps keyed by block/word ID. `normalizedPptxBoundingBox(fragment)` converts a fragment to fractional slide coordinates.

Each geometry result has `scope: 'mounted_slides'`, a status (`rendered`, `partially_rendered`, `not_mounted`, `not_visible`, `unresolved`), a fragment array and explicit issues. Fragments contain zero-based `slideIndex`, `slideWidth`, `slideHeight`, `boundingBox` and `coordinateSpace: 'slide-css-px'`. Boxes are axis-aligned and relative to the unscaled slide surface, so they remain stable when the viewer zooms. A wrapped phrase can produce several rectangles. Rotated text produces enclosing rectangles, not a claimed exact glyph polygon.

A source ref includes document identity, slide identity and occurrence index, node ID and node path, optional source-part identity, and optional row, column and paragraph indices. Raw shape IDs are insufficient because they repeat between slides. Ranges use inclusive start and exclusive end UTF-16 offsets within the paragraph. Preserve these refs when passing extracted content to an external service and return them with its citations; do not locate repeated words by searching visible text.

The source identity hashes the exact input bytes. Repacking or editing creates a new identity; models and source refs must come from compatible parser versions. Treat parsed models as immutable. The resolver rejects document mismatches, missing paths and changed range text. It reports unsupported mathematical subranges rather than highlighting an unrelated object. Notes are extractable with `includeNotes: true`, but resolve as `not_on_slide`. Figures use the whole object frame; individual chart points, chart-label ranges, and embedded-document text do not yet have sub-object refs.

## Availability and development

React PPTX 0.4.0 bundles rust-pptx and rust-pptx-wasm 0.2.0. The normal package build uses these registry-pinned crates. A development checkout can instead select local upstream sources with `RUST_OFFICE_PATH=/absolute/path/to/extend-office pnpm build:packages`.

`PptxBlockHighlights` follows the DOCX highlight component's props: `block`, `blocks`, `color`, `colorForBlock` and `onGeometry` for the active block. `onGeometries` additionally reports a map for all supplied targets. `PptxWordHighlights` accepts `word` and `words` with the same active-geometry callback.

The playground's **Blocks** toggle opens a virtualized inspector with text, word and JSON views, type/text filters and JSON download. Hover or focus a row to highlight it; click to mount and navigate to its slide. **Show all highlights** displays all matching targets on mounted slides. Notes remain extractable and explicitly report that they are not drawn on the slide.
