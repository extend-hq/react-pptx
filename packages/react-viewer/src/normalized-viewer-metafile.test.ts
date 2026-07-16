import type { PresentationDocument } from '@extend-ai/react-pptx-model';
import { describe, expect, it, vi } from 'vitest';

const { renderEmfToDataUrl, renderWmfToDataUrl } = vi.hoisted(() => ({
  renderEmfToDataUrl: vi.fn(async () => 'data:image/png;base64,converted'),
  renderWmfToDataUrl: vi.fn(async () => 'data:image/png;base64,wmf'),
}));

vi.mock('./metafile-renderer', () => ({ renderEmfToDataUrl, renderWmfToDataUrl }));

import { NormalizedPresentationViewer } from './normalized-viewer';

describe('normalized viewer metafiles', () => {
  it('converts a legacy EMF asset before a slide render completes', async () => {
    const presentation: PresentationDocument = {
      format: 'ppt',
      size: { widthEmu: 9_144_000, heightEmu: 6_858_000 },
      slides: [
        {
          id: 'slide-1',
          index: 0,
          nodes: [
            {
              id: 'chart-preview',
              type: 'image',
              transform: { x: 0, y: 0, width: 4_572_000, height: 3_429_000 },
              assetId: 'legacy-emf-1',
            },
          ],
        },
      ],
      masters: [],
      layouts: [],
      themes: [],
      assets: {
        'legacy-emf-1': {
          id: 'legacy-emf-1',
          contentType: 'image/x-emf',
          byteLength: 4,
          data: new Uint8Array([1, 2, 3, 4]),
        },
      },
      warnings: [],
    };
    const container = document.createElement('div');
    const viewer = new NormalizedPresentationViewer(container, presentation);

    await viewer.renderSlide();

    expect(renderEmfToDataUrl).toHaveBeenCalledOnce();
    expect(container.querySelector('img')?.src).toBe('data:image/png;base64,converted');

    await viewer.renderSlide();

    expect(renderEmfToDataUrl).toHaveBeenCalledOnce();
    viewer.destroy();
  });
});
