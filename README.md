# react-pptx

[![npm version](https://img.shields.io/npm/v/@extend-ai/react-pptx.svg)](https://www.npmjs.com/package/@extend-ai/react-pptx)
[![CI](https://github.com/extend-hq/react-pptx/actions/workflows/ci.yml/badge.svg)](https://github.com/extend-hq/react-pptx/actions/workflows/ci.yml)

`@extend-ai/react-pptx` is a browser-native React viewer for PowerPoint `.pptx` and
legacy `.ppt` files. It combines a Rust/WebAssembly parser, a format-neutral model,
and a virtualized HTML/SVG rendering engine.

## Workspace

- `crates/pptx-core` — safe format detection, PresentationML parsing, legacy OLE/PPT records, and normalized models.
- `crates/pptx-wasm` / `packages/wasm` — lazy browser Wasm runtime.
- `packages/presentation-model` — presentation types bundled into the public viewer.
- `packages/react-viewer` — the published React viewer and controller API.
- `apps/playground` — interactive upload, search, navigation, zoom, diagnostics, and fidelity test surface.

## Run the playground

```bash
pnpm install
pnpm dev
```

Open [http://localhost:4173](http://localhost:4173). The dev task generates a deterministic
two-slide fixture automatically, or you can drop any `.ppt` or `.pptx` file onto the viewer.

## Use the package

```bash
npm install @extend-ai/react-pptx
```

React and React DOM are peer dependencies already present in a React application.

```tsx
import { ReactPptxViewer } from '@extend-ai/react-pptx';
import '@extend-ai/react-pptx/styles.css';

export function Deck({ file }: { file: File }) {
  return <ReactPptxViewer source={file} mode="continuous" showThumbnails showToolbar />;
}
```

The package also exports `parsePresentation`, `initWasm`, `setWasmSource`, normalized
model types, search results, warnings, and an imperative `PptxViewerController`.

## Validation

```bash
pnpm build
pnpm typecheck
pnpm test
pnpm exec playwright install chromium
pnpm test:visual --project=chromium
```

The private PowerPoint-oracle corpus uses `pnpm powerpoint-oracle:validate` and
`pnpm powerpoint-oracle:compare`. See [the fidelity contract](docs/fidelity-harness.md).

## Fidelity boundary

PPTX parsing and rendering are owned by this repository. The Rust/Wasm parser resolves
PresentationML packages into the public normalized model, and the React package renders that
model with its own HTML/SVG slide surface. Shapes, styled text, images, groups, tables, charts,
backgrounds, transforms, and theme colors all travel through the same inspectable data model.

Legacy `.ppt` files use a separate, browser-native Rust pipeline; they are not converted
to `.pptx` and do not require LibreOffice or a rendering server. The parser follows the
`Current User` and `UserEditAtom` chain, merges the live persist-object directory, and
uses `SlideListWithText` to resolve source slide order. It then maps supported OfficeArt
shape/group anchors, text, pictures, and referenced master drawing content into the same
normalized model used by the public React API.

PNG pictures render directly. Compressed EMF and WMF records are decoded by Rust and
converted to browser-safe image data before a slide is displayed, which provides a
static fallback for content such as legacy chart previews. Unsupported effects and
editable chart/table semantics emit structured `degraded-rendering` warnings instead of
silently disappearing. This is PowerPoint-close static viewing, not complete binary
PowerPoint emulation: animation timelines, macros, active OLE controls, password
decryption, media playback, and authoring/export remain outside v1.

The regression suite includes the original three-slide PowerPoint 97–2003 OLE fixture
at `tests/fixtures/legacy/file-example-250kb.ppt`. Rust and browser tests open that file
directly and verify live slide order, text, normalized drawing nodes, thumbnails, and
diagnostics without a conversion step.

## License

MIT. See [third-party notices](THIRD_PARTY_NOTICES.md) for the Apache-2.0 EMF/WMF converter.

Maintainers can find the versioning, npm Trusted Publishing, and automated GitHub
release process in [the release guide](docs/releasing.md).
