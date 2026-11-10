# react-pptx

[![npm version](https://img.shields.io/npm/v/@extend-ai/react-pptx.svg)](https://www.npmjs.com/package/@extend-ai/react-pptx)
[![Monthly npm downloads](https://img.shields.io/npm/dm/@extend-ai/react-pptx.svg)](https://www.npmjs.com/package/@extend-ai/react-pptx)
[![CI](https://github.com/extend-hq/react-pptx/actions/workflows/ci.yml/badge.svg)](https://github.com/extend-hq/react-pptx/actions/workflows/ci.yml)

Browser-native React viewer for `.pptx` and legacy `.ppt`, using a Rust/WASM parser, a normalized presentation model and virtualized HTML/SVG rendering.

## Quick start

```bash
npm install @extend-ai/react-pptx
```

React and React DOM are peer dependencies.

```tsx
import { ReactPptxViewer } from "@extend-ai/react-pptx";
import "@extend-ai/react-pptx/styles.css";

export function Deck({ file }: { file: File }) {
  return <ReactPptxViewer source={file} defaultZoom="fit-width" showThumbnails showToolbar />;
}
```

`source` also accepts URLs, buffers and parsed presentations. `parsePresentation`, `PptxViewerController` and `createPptxThumbnailRenderer` support custom integrations. WASM loads lazily; set a custom source before parsing with `setWasmSource`. See the [package API](packages/react-viewer/README.md).

## Development

Use Node.js from [`.nvmrc`](.nvmrc), pnpm 9.15.0, Rust, `wasm-pack` and the `wasm32-unknown-unknown` target. Supported Node ranges are in [package.json](package.json).

```bash
pnpm install --frozen-lockfile
pnpm dev
pnpm build
pnpm typecheck
pnpm test
pnpm exec playwright install chromium
pnpm test:visual --project=chromium
```

`pnpm dev` rebuilds WASM, the model and the viewer before starting the playground at [localhost:4173](http://localhost:4173). Restart it after changing the Rust dependency.

The [playground](https://react-pptx-alpha.vercel.app/) opens PPTX files in edit mode. Click slide text or table cells to edit directly, drag an object’s border to move it, or use its eight handles to resize it. Double-click selects a word; Escape finishes text editing and keeps the object selected. Arrow keys move selected objects, and Tab cycles through them. Changes save when you leave a text box without reloading the slide. Undo/redo and Download preserve edits in the original PPTX package. The **Read only** switch disables editing; legacy PPT files and inherited layout/master objects remain read only. This editor is implemented in the playground; the published viewer API remains a viewer. The default presentation is a licensed excerpt with [source and adaptation credits](apps/playground/public/samples/ATTRIBUTION.txt).

| Work | Start here |
| --- | --- |
| Public viewer, controller and hooks | [packages/react-viewer](packages/react-viewer) |
| Presentation model | [packages/presentation-model](packages/presentation-model) |
| Rust/WASM adapter and browser loader | [packages/wasm](packages/wasm) |
| Playground and validation | [apps/playground](apps/playground), [tests](tests), [fidelity harness](docs/fidelity-harness.md) |
| Architecture and releases | [Architecture](docs/architecture.md), [release guide](docs/releasing.md) |

Parser and shared rendering fixes belong in [Rust Office](https://github.com/extend-hq/rust-office); this repository owns the browser adapter and React rendering. The private native-reference corpus uses `pnpm powerpoint-oracle:validate` and `pnpm powerpoint-oracle:compare`.

## Rendering boundaries

- Legacy `.ppt` is parsed directly in the browser. Unsupported content emits `degraded-rendering` warnings through `onWarning` / `showDiagnostics`; see [legacy format support](docs/legacy-ppt.md).
- Viewing is static. Slide transitions and animation timelines are exposed on the model (`transition`, `timing`) but not played; macros, active OLE controls, password decryption and authoring/export are outside the supported scope. Embedded audio and video play through native controls, and embedded objects are offered as downloads.
- SmartArt uses saved drawings; 3D models display saved Office previews. `include3dModels` optionally exposes original assets and scene metadata. See [OfficeArt visuals](packages/react-viewer/README.md#officeart-visuals).

## Live demo

[![PowerPoint viewer with a presentation slide and navigation controls](./docs/readme-assets/extend-ui-pptx-viewer.png)](https://www.extend.ai/ui/docs/components/pptx-viewer)

Click the screenshot for the viewer demo and integration docs. [MIT license](LICENSE) · [Third-party notices](THIRD_PARTY_NOTICES.md).

### Block citations

Extract paragraph, table-cell and word blocks with `extractPresentationBlocks`, resolve their source references with `resolvePptxSourceGeometry`, and render overlays with `PptxBlockHighlights`. See the [citation API and source-build instructions](docs/citations.md).
