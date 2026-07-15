# PowerPoint fidelity harness

PowerPoint-rendered reference PNGs are the visual oracle. The private corpus is stored under
`.powerpoint-oracle/` and is excluded from Git; `manifest.example.json` documents its contract.

The manifest pins browser name/version, platform, architecture, viewport, DPR, locale, timezone,
font fingerprint, source-deck hashes, per-slide references, and difference thresholds. Validation
must fail before pixel comparison when the renderer environment does not match.

```bash
pnpm powerpoint-oracle:validate .powerpoint-oracle/manifest.json
pnpm powerpoint-oracle:compare \
  .powerpoint-oracle/manifest.json \
  .powerpoint-oracle/captures \
  .powerpoint-oracle/results
```

Pixel checks complement, rather than replace, semantic assertions for slide order, bounds, z-order,
text, relationships, chart series, notes, links, and asset references.
