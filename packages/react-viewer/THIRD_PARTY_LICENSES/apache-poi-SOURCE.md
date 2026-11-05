# Preset geometry data

Source: [Apache POI](https://github.com/apache/poi/blob/338882ac8898df5c13a7d15f533204c5dd8607d6/poi/src/main/resources/org/apache/poi/sl/draw/geom/presetShapeDefinitions.xml), revision `338882ac8898df5c13a7d15f533204c5dd8607d6`.

Original path: `poi/src/main/resources/org/apache/poi/sl/draw/geom/presetShapeDefinitions.xml`.

The XML is retained unchanged. Apache POI license and notice are included alongside it.

SHA-256: `4a762444d8d85876881c02a5b1dedf6f73006fcd8acb7b4e393435615b37c780`.

The data contains 187 preset definitions. The evaluator accepts the standard guide operators, adjustment overrides, multiple paths, per-path fill and stroke modes, polar elliptical arcs, and quadratic/cubic segments. Geometry coordinates use EMUs unless a path declares its own coordinate dimensions. Bounds include Bézier control points and therefore can conservatively exceed the painted curve.

Evaluation limits are 1 MiB of supplied XML, 1,024 guide evaluations, 256 paths, and 8,192 path commands. Arc sweeps are limited to 16 revolutions; values must remain finite and rendered coordinates must not exceed 100,000,000 pixels. Invalid geometry returns no result. Adjustment handles, connection-site metadata, and inscribed text rectangles are retained in the preset data but are not part of the current path-returning API.
