# Performance contract

The production benchmark corpus should include a 50-slide, approximately 20 MiB enterprise deck.
On the pinned benchmark machine the targets are:

- First visible slide at or below 1,000 ms.
- Cached slide navigation at or below 100 ms.
- Windowed continuous mode without full-deck DOM mounting.
- No retained slide DOM, observers, chart instances, or object URLs after replacement.
- No more than a 10% regression without an explicit benchmark-baseline change.

The parser returns one normalized presentation model consumed directly by the React renderer.
Continuous mode mounts and evicts slide DOM around the viewport. Binary media stays in typed
arrays; browser object URLs are created on first render, cached per asset, and revoked when the
viewer is replaced. Callers do not need to base64-encode media.
