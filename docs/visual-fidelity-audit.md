# PowerPoint visual fidelity audit — November 2026

The audit compared 362 slides from 98 presentations against native PowerPoint
references: 349 slides from 97 corpus presentations and 13 freshly exported
playground slides. One additional presentation was excluded because hidden slides
made its source-to-reference page mapping ambiguous. This is measured coverage,
not a claim of complete or pixel-exact PowerPoint compatibility.

## Changes

- Expanded EMF/WMF rendering for compound paths, cubic curves, transformed
  geometry, pen styles, drawing-state restoration, brush fills, and bounded
  uncompressed RGB bitmap transfers. Unsupported object handles retain their
  positions so later objects are not selected incorrectly.
- Corrected OfficeArt gradient interpolation, centered radial tiles, connector
  arrowhead placement, and positive subpixel paint bounds upstream.
- Kept Latin, East Asian, and complex-script font inheritance separate; corrected
  master/default paragraph style precedence and direct text-fill color parsing.
- Preserved inherited bullet font and size across local overrides, corrected
  hanging-indent placement, and used surrounding text color for marker fallback.
- Added an Office-compatible fallback for unavailable font families and surfaced
  slide-scoped parser warnings through the public parse result.

The source build pins `rust-pptx` and `rust-pptx-wasm` to upstream revision
`7555683a8e690baf37fe3442d9376b43750276a1`. No new crate or npm release is implied.

## Results

All 362 baseline/final pairs passed capture-identity and dimension validation.

| Measurement | Baseline | Updated |
| --- | ---: | ---: |
| Mean differing pixels per slide | 6.466% | 5.999% |
| Median differing pixels per slide | 4.873% | 4.658% |
| 95th-percentile differing pixels | 16.629% | 15.379% |
| Playground mean differing pixels | 3.778% | 2.270% |
| Playground opening slide differing pixels | 9.487% | 0.462% |
| Corpus slides within their configured limit | 174 / 349 | 177 / 349 |

114 slides improved, 83 regressed, and 165 were unchanged under the fixed
comparison algorithm. Of the regressions, 66 increased by at most 0.1 percentage
points and 17 by more; one increased by more than one percentage point. The report
retains every result, including regressions. 172 corpus slides still exceed their
configured limits, so this pass does not satisfy a universal fidelity gate.
Strict byte-level pixel differences remain approximately 40.8% in both passes,
underscoring the JPEG, antialiasing, and color-profile limits of this comparison.

## Method and evidence limits

The baseline is repository revision `d7e0676`. Both browser passes use identical
source bytes, slide dimensions, and per-slide font inputs. Captures wait for
fonts, images, and layout to settle. SHA-256 receipts identify source files,
native references, renderer bundles, WASM binaries, font inputs, and captures.
Pixel comparisons use pixelmatch 7.2.0, threshold 0.1, with antialiasing excluded;
there is no image registration or threshold relaxation.

Browser screenshots are JPEG, so compression contributes to measured differences.
Historical references were exported through PowerPoint PDF and rasterized at
96 dpi; their original application version and export-time hashes were not
recorded. Current reference bytes were sealed for this audit. A fresh native
screen check also identified an ICC-profile shift in a retained PDF gradient:
the corrected browser matches the current PowerPoint screen more closely than
that PDF reference. Original reference pixels and comparison scores remain intact.

REA native observations, PDF geometry/font analysis, source inspection, and
targeted native screen checks informed the fixes. They do not constitute an
exhaustive reverse-engineering of PowerPoint. Proprietary fonts and corpus assets
are not redistributed with the viewer.

The local evidence bundle is under `output/visual-audit-20261110/`, including the
interactive `comparison/report.html`, capture receipts, feature inventory,
native observations, and separate snapshots of intermediate iterations. These
artifacts are intentionally excluded from Git.

## Remaining limitations

- SmartArt without usable cached drawing data still uses an approximate layout.
- Multiply fill overlays and other advanced OfficeArt compositing remain partial.
- Compressed EOT/MTX embedded fonts are not decoded; fallback metrics can differ
  from native PowerPoint even when object geometry is correct.
- EMF+, complex clipping, masked/indexed bitmaps, and exact metafile text spacing
  remain incomplete. A nonempty metafile render does not imply a complete one.
- Explicit bullet colors are not represented separately from surrounding text
  color in the public model, so some list markers still use an approximate color.
- A uniform line-height multiplier does not reproduce font-dependent native
  line metrics. This causes accumulated paragraph shifts even when the exact
  font is loaded; the audit retains a measured Corbel spacing regression.
- The corpus does not exercise active WordArt warps or establish exhaustive shape
  and parameter coverage. Static comparisons do not validate animation, media
  playback, or native 3D rendering.

Warning-free parsing is not proof of visual equivalence. The corpus includes
known failures, and aggregate scores must be read alongside individual slides
and feature coverage.

## Validation

The locked production build, TypeScript checks, 107 existing JavaScript unit
tests, three existing fidelity-contract tests, and the existing Rust fixture
check pass. The upstream inheritance changes also pass 44 existing PPTX tests.
No new unit tests were added for this audit.
