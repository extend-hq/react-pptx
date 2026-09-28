---
"@extend-ai/react-pptx": patch
---

Fix paragraph styling in placeholders and render PowerPoint's default table style.

- Slide paragraphs no longer take their `lvl`/`marL` from the layout placeholder's prompt paragraphs ("Click to edit Master text styles", "Second level", …). They inherit only through `a:lstStyle`, which stops plain bullet lists from stair-stepping.
- A placeholder without a `type` (e.g. `<p:ph idx="1"/>`) defaults to `obj` per ECMA-376 and now uses the master's `bodyStyle` instead of `otherStyle`, restoring its bullets, indents and size.
- When `ppt/tableStyles.xml` doesn't define the requested style, the built-in "Medium Style 2 - Accent 1" (`{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}`, PowerPoint's default) is used, so its fills, banding and borders render.
