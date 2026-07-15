# Legacy `.ppt` support

Legacy PowerPoint files use OLE compound storage and the MS-PPT record format rather than
PresentationML. They are parsed directly by the Rust/Wasm pipeline and rendered client-side;
the runtime does not invoke LibreOffice, upload the file to a conversion service, or first
rewrite it as `.pptx`.

## Native parsing path

For a normal PowerPoint 97–2003 file, the parser:

1. Opens the OLE `PowerPoint Document`, `Current User`, and optional `Pictures` streams under
   resource limits.
2. Uses `Current User` to find the newest `UserEditAtom`, follows the bounded edit chain, and
   combines its full and incremental persist-object directories into the live object map.
3. Resolves the live `DocumentContainer`, document dimensions, master references, and the
   `SlideListWithText` source order. This avoids rendering deleted historical objects or treating
   master placeholder text as an extra slide.
4. Reads Unicode and single-byte text atoms and associates outline text with the OfficeArt shape
   that owns it.
5. Maps OfficeArt client anchors, child anchors, and nested group coordinate systems into the
   normalized EMU transform used by `PresentationDocument`.
6. Resolves a slide's referenced master and places its supported non-placeholder drawing content
   before slide-local nodes. This preserves common inherited backgrounds and decorations without
   requiring a separate React renderer for `.ppt`.
7. Resolves supported OfficeArt picture references through the blip store and `Pictures` stream.
   PNG/JPEG assets render directly. Compressed EMF and WMF records are decompressed by Rust, then
   converted to a browser-safe image before the slide render completes.

Files without a `Current User` stream use a bounded sequential text-recovery fallback. That path
is intentionally lower fidelity and emits a `degraded-rendering` warning.

## Fidelity and diagnostics

The output is the same normalized `PresentationDocument` consumed by navigation, search,
thumbnails, custom scroll areas, and the public viewer hooks. Callers do not need a separate API
or component for `.ppt`.

Legacy support is static and deliberately honest about unsupported binary semantics. Common text,
shapes, groups, inherited drawing content, and pictures are recoverable, while advanced OfficeArt
effects, editable charts/tables, SmartArt-like constructs, animations, transitions, active OLE
objects, and timed media can be omitted or represented by a static preview. Known loss is reported
with a structured `degraded-rendering` warning and can be inspected through `onWarning` or
`showDiagnostics`; the viewer does not claim pixel identity with desktop PowerPoint. Encrypted
streams fail closed with `encrypted-document`, and password decryption is not attempted.

## Real-file regression

The checked-in regression suite opens a real PowerPoint 97–2003 OLE file directly in Rust and in
the browser playground: `tests/fixtures/legacy/file-example-250kb.ppt`. The three-slide deck covers
body text, master/background artwork, a chart preview, and a table. Tests assert live source slide
order, meaningful extracted text, absence of master placeholder text as a slide, renderable
normalized nodes, thumbnail navigation, and visible diagnostics.

The fixture is never converted during the test or at runtime. Reference exports may still be used
offline by the fidelity harness as an oracle, but they are not an application dependency.
