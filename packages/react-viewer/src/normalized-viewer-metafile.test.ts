import type { PresentationDocument } from '@extend-ai/react-pptx-model';
import { beforeEach, describe, expect, it, vi } from 'vitest';

const { renderEmfToDataUrl, renderWmfToDataUrl } = vi.hoisted(() => ({
  renderEmfToDataUrl: vi.fn(),
  renderWmfToDataUrl: vi.fn(),
}));

vi.mock('./metafile-renderer', () => ({ renderEmfToDataUrl, renderWmfToDataUrl }));

import { NormalizedPresentationViewer } from './normalized-viewer';

describe('normalized viewer metafiles', () => {
  beforeEach(() => {
    renderEmfToDataUrl.mockReset().mockResolvedValue('data:image/png;base64,converted');
    renderWmfToDataUrl.mockReset().mockResolvedValue('data:image/png;base64,wmf');
  });

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

  it('scopes resource readiness to the mount that requested it', async () => {
    let resolveConversion: ((url: string) => void) | undefined;
    renderEmfToDataUrl.mockImplementationOnce(
      () =>
        new Promise<string>((resolve) => {
          resolveConversion = resolve;
        }),
    );
    const presentation: PresentationDocument = {
      format: 'pptx',
      size: { widthEmu: 9_144_000, heightEmu: 6_858_000 },
      slides: [
        {
          id: 'slow-thumbnail',
          index: 0,
          nodes: [
            {
              id: 'slow-image',
              type: 'image',
              transform: { x: 0, y: 0, width: 4_572_000, height: 3_429_000 },
              assetId: 'slow-emf',
            },
          ],
        },
        { id: 'fast-slide', index: 1, nodes: [] },
      ],
      masters: [],
      layouts: [],
      themes: [],
      assets: {
        'slow-emf': {
          id: 'slow-emf',
          contentType: 'image/x-emf',
          byteLength: 4,
          data: new Uint8Array([1, 2, 3, 4]),
        },
      },
      warnings: [],
    };
    const container = document.createElement('div');
    const thumbnailTarget = document.createElement('div');
    const viewer = new NormalizedPresentationViewer(container, presentation);

    const thumbnail = viewer.renderThumbnailToContainer(0, thumbnailTarget);
    let thumbnailReady = false;
    void thumbnail.ready.then(() => {
      thumbnailReady = true;
    });
    let slideReady = false;
    const slideRender = viewer.renderSlide(1).then(() => {
      slideReady = true;
    });
    await new Promise((resolve) => setTimeout(resolve, 0));

    expect(slideReady).toBe(true);
    expect(thumbnailReady).toBe(false);
    resolveConversion?.('data:image/png;base64,slow');
    await Promise.all([thumbnail.ready, slideRender]);
    expect(thumbnailReady).toBe(true);
    viewer.destroy();
  });

  it('scrolls and publishes an off-window target before its resources are ready', async () => {
    let resolveConversion: ((url: string) => void) | undefined;
    renderEmfToDataUrl.mockImplementationOnce(
      () =>
        new Promise<string>((resolve) => {
          resolveConversion = resolve;
        }),
    );
    const presentation: PresentationDocument = {
      format: 'pptx',
      size: { widthEmu: 9_144_000, heightEmu: 6_858_000 },
      slides: Array.from({ length: 5 }, (_, index) => ({
        id: `slide-${index}`,
        index,
        nodes:
          index === 4
            ? [
                {
                  id: 'target-image',
                  type: 'image' as const,
                  transform: { x: 0, y: 0, width: 4_572_000, height: 3_429_000 },
                  assetId: 'target-emf',
                },
              ]
            : [],
      })),
      masters: [],
      layouts: [],
      themes: [],
      assets: {
        'target-emf': {
          id: 'target-emf',
          contentType: 'image/x-emf',
          byteLength: 4,
          data: new Uint8Array([1, 2, 3, 4]),
        },
      },
      warnings: [],
    };
    const container = document.createElement('div');
    const scrollTo = vi.fn();
    container.scrollTo = scrollTo as unknown as typeof container.scrollTo;
    const changes: number[] = [];
    const viewer = new NormalizedPresentationViewer(container, presentation, {
      onSlideChange: (index) => changes.push(index),
    });
    await viewer.renderList({ enabled: true, overscanViewport: 0 });
    scrollTo.mockClear();

    const firstNavigation = viewer.goToSlide(4);
    let navigationReady = false;
    const navigation = viewer.goToSlide(4).then(() => {
      navigationReady = true;
    });

    expect(scrollTo).toHaveBeenCalled();
    expect(viewer.currentSlideIndex).toBe(4);
    expect(changes).toEqual([0, 4]);
    expect(renderEmfToDataUrl).not.toHaveBeenCalled();
    for (let attempt = 0; attempt < 10 && !renderEmfToDataUrl.mock.calls.length; attempt += 1) {
      await new Promise((resolve) => setTimeout(resolve, 10));
    }
    expect(renderEmfToDataUrl).toHaveBeenCalledOnce();
    expect(navigationReady).toBe(false);

    resolveConversion?.('data:image/png;base64,target');
    await Promise.all([firstNavigation, navigation]);
    expect(navigationReady).toBe(true);
    expect(
      container.querySelector('[data-rpv-list-item="4"] [data-rpv-slide-wrapper]'),
    ).not.toBeNull();
    viewer.destroy();
  });

  it('remounts a pending slide that leaves and re-enters the virtual window', async () => {
    let resolveConversion: ((url: string) => void) | undefined;
    renderEmfToDataUrl.mockImplementationOnce(
      () =>
        new Promise<string>((resolve) => {
          resolveConversion = resolve;
        }),
    );
    const presentation: PresentationDocument = {
      format: 'pptx',
      size: { widthEmu: 9_144_000, heightEmu: 6_858_000 },
      slides: Array.from({ length: 5 }, (_, index) => ({
        id: `slide-${index}`,
        index,
        nodes:
          index === 0
            ? [
                {
                  id: 'slow-first-image',
                  type: 'image' as const,
                  transform: { x: 0, y: 0, width: 4_572_000, height: 3_429_000 },
                  assetId: 'slow-first-emf',
                },
              ]
            : [],
      })),
      masters: [],
      layouts: [],
      themes: [],
      assets: {
        'slow-first-emf': {
          id: 'slow-first-emf',
          contentType: 'image/x-emf',
          byteLength: 4,
          data: new Uint8Array([1, 2, 3, 4]),
        },
      },
      warnings: [],
    };
    const container = document.createElement('div');
    const rendered: number[] = [];
    const viewer = new NormalizedPresentationViewer(container, presentation, {
      onSlideRendered: (index) => rendered.push(index),
    });

    const initialRender = viewer.renderList({ enabled: true, overscanViewport: 0 });
    const firstItem = container.querySelector<HTMLElement>('[data-rpv-list-item="0"]')!;
    expect(firstItem.querySelector('[data-rpv-slide-wrapper]')).not.toBeNull();

    container.scrollTop = 4 * 744;
    container.dispatchEvent(new Event('scroll'));
    expect(firstItem.querySelector('[data-rpv-slide-wrapper]')).toBeNull();

    container.scrollTop = 0;
    container.dispatchEvent(new Event('scroll'));
    expect(firstItem.querySelector('[data-rpv-slide-wrapper]')).not.toBeNull();
    expect(rendered.filter((index) => index === 0)).toHaveLength(2);

    resolveConversion?.('data:image/png;base64,ready');
    await initialRender;
    viewer.destroy();
  });
});
