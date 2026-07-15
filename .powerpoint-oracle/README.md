# Private PowerPoint oracle corpus

The corpus itself is intentionally not committed. Create `manifest.json`, source decks,
PowerPoint-rendered reference PNGs, and captured viewer PNGs in this directory. Run
`pnpm powerpoint-oracle:validate` before `pnpm powerpoint-oracle:compare`.

The manifest must pin browser, viewport, DPR, and a font fingerprint. Comparisons fail
closed when that renderer environment changes.

On macOS with Microsoft PowerPoint installed, build and compare the private corpus with:

```sh
pnpm powerpoint-oracle:export "/absolute/path/to/pptx-corpus"
pnpm powerpoint-oracle:capture
pnpm powerpoint-oracle:compare
```

`export` skips the `invalid-parser-fixtures` directory, automates PowerPoint PDF export,
renders 96-DPI reference PNGs, and records the fonts PowerPoint actually used. `capture`
loads matching fonts from the local Office font cache into Chromium without copying them
into the repository, captures every slide at its intrinsic dimensions, and enforces the
pinned browser/font environment.
