# Third-party notices

## emf-converter

The viewer bundles a locally patched copy of
[`emf-converter` 1.5.0](https://github.com/ChristopherVR/emf-converter) for browser-native
rendering of EMF/WMF fallback previews embedded in legacy PowerPoint files. The compatibility
patch corrects mapping-mode scaling and `LOGFONTW` decoding. It is licensed under the Apache
License 2.0; its copyright and license remain with its authors.

The published package includes the Apache License 2.0 text at
`THIRD_PARTY_LICENSES/Apache-2.0.txt`.

## Chart rendering stack (vendored from @extend-ai/react-xlsx)

The chart renderer and chart model builder in `src/charts/` are vendored from
[`@extend-ai/react-xlsx`](https://github.com/extend-hq/react-xlsx) (MIT) so
PowerPoint charts render with the same styles and colors as Excel charts.

The chart libraries `d3-scale`, `d3-shape`, `d3-hierarchy`, `d3-geo` (ISC,
Copyright Mike Bostock), `topojson-client` (ISC), and
[`regl`](https://github.com/regl-project/regl) (MIT) are declared as regular
package dependencies.

The published package bundles the
[`us-atlas`](https://github.com/topojson/us-atlas) and
[`world-atlas`](https://github.com/topojson/world-atlas) TopoJSON boundary data
(ISC License) inside a dynamically imported chunk that only loads when a
presentation contains an Excel map chart.

The published package includes the ISC and MIT license texts at
`THIRD_PARTY_LICENSES/ISC.txt` and `THIRD_PARTY_LICENSES/MIT.txt`.

## Virtualized scrolling

The continuous-scroll mode is windowed with
[`@tanstack/virtual-core`](https://github.com/TanStack/virtual) (MIT License,
Copyright (c) 2021-present Tanner Linsley), declared as a regular package
dependency.
