# Third-party notices

## PowerPoint parsing and chart rendering

The browser runtime bundles the PowerPoint parser and chart renderer from
[extend-office](https://github.com/extend-hq/extend-office). Its MIT license
is included in the published package at `THIRD_PARTY_LICENSES/extend-office.txt`.

The React chart renderer and chart model builder used for fallback rendering
are vendored from [`@extend-ai/react-xlsx`](https://github.com/extend-hq/react-xlsx) (MIT).

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

The continuous-scroll mode and React thumbnail rails are windowed with
[`@tanstack/virtual-core`](https://github.com/TanStack/virtual) and
[`@tanstack/react-virtual`](https://github.com/TanStack/virtual) (MIT License,
Copyright (c) 2021-present Tanner Linsley), declared as regular package
dependencies.
