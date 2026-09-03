import type { PresentationDocument, ShapeNode } from '@extend-ai/react-pptx-model';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createPptxThumbnailRenderer } from './thumbnail-renderer';
import type { ParsedPresentation } from './types';

function parsedPresentation(slideCount = 3, withImages = false): ParsedPresentation {
  const textNode = (index: number): ShapeNode => ({
    id: `shape-${index}`,
    type: 'shape',
    transform: { x: 0, y: 0, width: 2_000_000, height: 1_000_000 },
    geometry: { preset: 'rect' },
    paragraphs: [{ runs: [{ text: `Requested slide ${index}` }] }],
  });
  const document: PresentationDocument = {
    format: 'pptx',
    size: { widthEmu: 9_144_000, heightEmu: 5_143_500 },
    slides: Array.from({ length: slideCount }, (_, index) => ({
      id: `slide-${index}`,
      index,
      nodes: withImages
        ? [
            {
              id: `image-${index}`,
              type: 'image' as const,
              transform: { x: 0, y: 0, width: 2_000_000, height: 1_000_000 },
              assetId: 'shared',
            },
          ]
        : [textNode(index)],
    })),
    masters: [],
    layouts: [],
    themes: [],
    assets: withImages
      ? {
          shared: {
            id: 'shared',
            contentType: 'image/png',
            byteLength: 3,
            data: new Uint8Array([1, 2, 3]),
          },
        }
      : {},
    warnings: [],
  };
  return { kind: 'parsed-presentation', document, warnings: [] };
}

describe('thumbnail renderer', () => {
  afterEach(() => {
    document.body.replaceChildren();
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it('renders only the requested slide without a viewer surface', async () => {
    const phases: string[] = [];
    const renderer = createPptxThumbnailRenderer(parsedPresentation(), {
      onPerformanceMeasurement: (measurement) => phases.push(measurement.phase),
    });

    const result = await renderer.renderSlide(1, {
      maxWidth: 240,
      maxHeight: 135,
      output: 'svg',
    });

    expect(result.output).toBe('svg');
    expect(result.width).toBe(240);
    expect(result.height).toBe(135);
    expect(result.data).toContain('Requested slide 1');
    expect(result.data).not.toContain('Requested slide 0');
    expect(result.data).not.toContain('Requested slide 2');
    expect(result.measurements.map((measurement) => measurement.phase)).toEqual([
      'slide-dom-svg-rendering',
      'image-decoding',
    ]);
    expect(phases).toContain('embedded-font-preparation');
    renderer.destroy();
  });

  it('does not wait for the global font set under the default thumbnail policy', async () => {
    Object.defineProperty(document, 'fonts', {
      configurable: true,
      value: { ready: new Promise(() => {}) },
    });
    const renderer = createPptxThumbnailRenderer(parsedPresentation(1));

    await expect(renderer.ready).resolves.toBeUndefined();

    renderer.destroy();
    Reflect.deleteProperty(document, 'fonts');
  });

  it('bounds concurrent work and cancels a queued render', async () => {
    vi.stubGlobal('URL', {
      createObjectURL: vi.fn(() => 'blob:thumbnail-asset'),
      revokeObjectURL: vi.fn(),
    });
    const pendingDecodes: Array<() => void> = [];
    let active = 0;
    let maximumActive = 0;
    Object.defineProperty(HTMLImageElement.prototype, 'decode', {
      configurable: true,
      value: vi.fn(
        () =>
          new Promise<void>((resolve) => {
            active += 1;
            maximumActive = Math.max(maximumActive, active);
            pendingDecodes.push(() => {
              active -= 1;
              resolve();
            });
          }),
      ),
    });
    const renderer = createPptxThumbnailRenderer(parsedPresentation(4, true), {
      concurrency: 2,
    });
    const first = renderer.renderSlide(0, { output: 'svg' });
    const second = renderer.renderSlide(1, { output: 'svg' });
    const abort = new AbortController();
    const cancelled = renderer
      .renderSlide(2, { output: 'svg', signal: abort.signal })
      .catch((error: unknown) => error);
    const fourth = renderer.renderSlide(3, { output: 'svg' });

    await vi.waitFor(() => expect(pendingDecodes).toHaveLength(2));
    expect(maximumActive).toBe(2);
    abort.abort();
    await expect(cancelled).resolves.toMatchObject({ code: 'aborted' });

    pendingDecodes.shift()?.();
    await vi.waitFor(() => expect(pendingDecodes).toHaveLength(2));
    pendingDecodes.splice(0).forEach((resolve) => resolve());
    await expect(Promise.all([first, second, fourth])).resolves.toHaveLength(3);
    expect(maximumActive).toBe(2);
    renderer.destroy();
  });
});
