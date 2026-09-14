import { describe, expect, it } from 'vitest';
import { normalizeViewerZoomLevel, resolveViewerZoom } from './zoom';

describe('viewer zoom resolution', () => {
  it.each([
    ['automatic', 1_920, 1_080, 100],
    ['fit-width', 1_920, 1_080, 200],
    ['fit-page', 1_920, 1_080, 150],
    ['automatic', 480, 360, 50],
    ['fit-width', 480, 360, 50],
    ['fit-page', 480, 360, 50],
  ] as const)('resolves %s in a %sx%s viewport', (level, width, height, expected) => {
    expect(resolveViewerZoom(level, width, height, 960, 720)).toBe(expected);
  });

  it('clamps numeric and responsive levels to the supported range', () => {
    expect(normalizeViewerZoomLevel(1)).toBe(10);
    expect(normalizeViewerZoomLevel(900)).toBe(400);
    expect(normalizeViewerZoomLevel(Number.NEGATIVE_INFINITY)).toBe(10);
    expect(normalizeViewerZoomLevel(Number.POSITIVE_INFINITY)).toBe(400);
    expect(resolveViewerZoom('fit-width', 10_000, 1_000, 960, 720)).toBe(400);
    expect(resolveViewerZoom('fit-page', 1, 1, 960, 720)).toBe(10);
  });
});
