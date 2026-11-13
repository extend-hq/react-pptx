# Third-party notices

## Office parsing and chart rendering

The browser runtime bundles the PowerPoint parser, OpenDocument parser, shared
OfficeArt renderer, and chart renderer from [rust-office](https://github.com/extend-hq/rust-office).
Their MIT license is included in the published package at
`THIRD_PARTY_LICENSES/extend-office.txt`.

The shared OfficeArt renderer includes preset geometry definitions from
[Apache POI](https://poi.apache.org/) under the Apache License 2.0.
The published package includes the upstream license, notice, and source revision
information at `THIRD_PARTY_LICENSES/apache-poi-LICENSE.txt`,
`THIRD_PARTY_LICENSES/apache-poi-NOTICE.txt`, and
`THIRD_PARTY_LICENSES/apache-poi-SOURCE.md`.

The chart renderer embeds its own map boundary data for Excel map charts; the
published package no longer ships TopoJSON atlases or JavaScript chart libraries.

## Virtualized scrolling

The continuous-scroll mode and React thumbnail rails are windowed with
[`@tanstack/virtual-core`](https://github.com/TanStack/virtual) and
[`@tanstack/react-virtual`](https://github.com/TanStack/virtual) (MIT License,
Copyright (c) 2021-present Tanner Linsley), bundled into the published
JavaScript so the package has no runtime dependencies beyond React. The MIT
license text is included at `THIRD_PARTY_LICENSES/MIT.txt`.
