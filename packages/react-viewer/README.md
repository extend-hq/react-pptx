# @extend-ai/react-pptx

[![npm version](https://img.shields.io/npm/v/@extend-ai/react-pptx.svg)](https://www.npmjs.com/package/@extend-ai/react-pptx)

Browser-native PowerPoint viewing for React.

```bash
npm install @extend-ai/react-pptx
```

React and React DOM are peer dependencies already present in a React application.

## Main API

- `ReactPptxViewer` renders URLs, `Blob`, `ArrayBuffer`, `Uint8Array`, parsed presentations, or normalized documents.
- `parsePresentation` performs native parsing and returns the same normalized model the viewer renders.
- `initWasm` and `setWasmSource` support eager initialization and custom Wasm/CDN hosting.
- `PptxViewerController` provides navigation, zoom, fit, text search, highlighting, model access,
  and detached thumbnail rendering.

Import `@extend-ai/react-pptx/styles.css` alongside the component; the stylesheet is emitted as
a separate package asset. Parsing and rendering are client-side; importing the JavaScript package
during SSR is safe.

Unknown or degraded content is surfaced through `onWarning` and `showDiagnostics` rather
than silently removed.

## Supported presentation formats

`.pptx` and legacy PowerPoint 97–2003 `.ppt` files are both accepted by the same `source`
prop and `parsePresentation()` API. Legacy files are opened natively from OLE compound
storage by the Rust/Wasm parser; the browser does not need LibreOffice, a server-side
conversion endpoint, or a temporary `.pptx` copy.

For `.ppt`, the parser follows the `Current User`/`UserEditAtom` edit chain and live
persist-object directory to recover the actual slide list. Supported OfficeArt group and
shape anchors, text, pictures, and referenced master drawing content are normalized into
the same `PresentationDocument` consumed by the React viewer. PNG/JPEG pictures render
directly, while decoded EMF/WMF records use a bounded browser-side image conversion for
static previews.

Legacy rendering is static, not full PowerPoint emulation. Advanced effects, editable
binary charts and tables, animations, transitions, active OLE objects, and timed media may
degrade to a static picture or remain unsupported. Inspect `onWarning` or enable
`showDiagnostics` to receive explicit `degraded-rendering` warnings:

```tsx
<ReactPptxViewer
  source={file}
  showDiagnostics
  onWarning={(warning) => {
    if (warning.code === 'degraded-rendering') {
      console.warn(warning.message, warning.slideIndex);
    }
  }}
/>
```

The repository regression suite opens a real three-slide PowerPoint 97–2003 file at
`tests/fixtures/legacy/file-example-250kb.ppt` directly in Rust and the playground. It
checks slide order, text, normalized nodes, thumbnail navigation, and diagnostics without
a conversion step.

## WebAssembly asset

Native PowerPoint parsing runs through a WebAssembly module that loads lazily. Most apps can
use the default loader. If your bundler or deployment needs to host the binary explicitly,
configure it before the first presentation is parsed:

```ts
import { setWasmSource } from '@extend-ai/react-pptx';

setWasmSource('https://cdn.example.com/pptx_wasm_bg.wasm');
// A URL, Request, Response, ArrayBuffer/TypedArray, or WebAssembly.Module also works.
```

The binary is also available as a package subpath for bundlers that support asset URL imports:

```ts
import wasmUrl from '@extend-ai/react-pptx/pptx_wasm_bg.wasm?url';
import { setWasmSource } from '@extend-ai/react-pptx';

setWasmSource(wasmUrl);
```

Configured strings, URLs, request URLs, bytes, and `WebAssembly.Module` values are forwarded to
the parser worker. A `Response` is supported on the main thread because it cannot be forwarded
reliably to a worker. `initWasm()` optionally accepts the same source when eager initialization
is useful.

## Fonts

Fonts embedded in a `.pptx` are decoded by the Rust parser and registered before the first
slide is rendered. Applications can also provide licensed/self-hosted fonts. The viewer
preserves the requested PowerPoint family, then appends Office-compatible and script-complete
fallbacks for Latin, CJK, Arabic, Hebrew, emoji, and math glyphs.

```tsx
<ReactPptxViewer
  source={file}
  fonts={{
    sources: [
      { family: 'Brand Sans', source: '/fonts/brand-sans.woff2' },
      {
        family: 'Brand Sans',
        source: '/fonts/brand-sans-bold.woff2',
        descriptors: { weight: '700' },
      },
    ],
    fallbacks: { 'Brand Sans': ['Inter', 'Arial'] },
    fallbackFamilies: ['Noto Sans', 'Noto Sans CJK SC', 'Noto Sans Arabic'],
  }}
  onWarning={(warning) => {
    if (warning.code === 'missing-font') console.warn(warning.message);
  }}
/>
```

The package does not redistribute proprietary Microsoft fonts. For identical metrics across
machines, self-host the fonts your license permits or install metric-compatible families such
as Carlito and Caladea. `resolvePptxFontFamily()` and `OFFICE_FONT_FALLBACKS` are exported for
applications that also render presentation text outside the viewer.

## Custom scroll areas and hooks

Pass the actual scrolling viewport through `virtualization.scrollElement`. This works with
Radix, Base UI, coss, or any host-owned scroll area and makes that element the observer root.

```tsx
const [scrollElement, setScrollElement] = useState<HTMLDivElement | null>(null);
const viewer = usePptxViewer();
const parsed = usePptxPresentation(file);

return (
  <ScrollArea.Root>
    <ScrollArea.Viewport ref={setScrollElement}>
      {parsed.presentation ? (
        <ReactPptxViewer
          ref={viewer.ref}
          source={parsed.presentation}
          virtualization={{ enabled: true, overscanViewport: 2, scrollElement }}
        />
      ) : null}
    </ScrollArea.Viewport>
  </ScrollArea.Root>
);
```

`usePptxPresentation` (`usePptxModel` alias) separates parsing from rendering, while
`usePptxViewer` exposes the reactive navigation/search/zoom controller. `onViewportReady`
provides the internal slide surface for advanced integrations and test harnesses.

## Thumbnail hook

`usePptxViewerThumbnails` exposes detached slide previews for a consumer-owned filmstrip. PPTX
thumbnails use the same live DOM/SVG renderer and loaded fonts as the main viewer, so attach each
stable `containerRef` to a regular element rather than a canvas.

```tsx
import { ReactPptxViewer, usePptxViewer, usePptxViewerThumbnails } from '@extend-ai/react-pptx';

export function PresentationWithThumbnails({ file }: { file: ArrayBuffer }) {
  const viewer = usePptxViewer();
  const { thumbnails } = usePptxViewerThumbnails(viewer.controller, {
    resolution: { maxWidth: 180, maxHeight: 120 },
    renderWindow: {
      visibleSlideIndexes: [0, 1, 2],
      prefetchSlideIndexes: [3, 4, 5],
    },
  });

  return (
    <div style={{ display: 'grid', gridTemplateColumns: '200px 1fr', gap: 16 }}>
      <aside style={{ display: 'grid', gap: 12, alignContent: 'start' }}>
        {thumbnails.map((thumbnail) => (
          <button
            key={thumbnail.slideIndex}
            type="button"
            onClick={() => void viewer.controller?.goToSlide(thumbnail.slideIndex)}
          >
            <div
              ref={thumbnail.containerRef}
              style={{
                width: thumbnail.width,
                height: thumbnail.height,
                overflow: 'hidden',
              }}
            />
          </button>
        ))}
      </aside>
      <ReactPptxViewer ref={viewer.ref} source={file} mode="slide" />
    </div>
  );
}
```

The result also provides `renderThumbnail(slideIndex, element)` for imperative integrations and
`rerenderAttachedThumbnails()` for refreshing every currently attached preview. `disabled` keeps
the metadata stable while pausing and disposing thumbnail rendering. For large decks, virtualize
the rail against its actual scroll element: mount `containerRef` only for virtual rows, pass those
rows through `renderWindow.visibleSlideIndexes`, and use `prefetchSlideIndexes` to warm nearby
slides. The built-in `showThumbnails` filmstrip follows the same windowed behavior.
