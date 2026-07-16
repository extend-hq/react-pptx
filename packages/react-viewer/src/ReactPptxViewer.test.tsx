import type { PresentationDocument, ShapeNode } from '@extend-ai/react-pptx-model';
import { act, StrictMode } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { PptxFontManager } from './fonts';
import { usePptxViewer } from './hooks';
import { ReactPptxViewer } from './ReactPptxViewer';
import { usePptxViewerThumbnails } from './thumbnails';
import type { PptxSlideThumbnailRenderWindow, PptxSlideThumbnailResolution } from './types';

const textShape: ShapeNode = {
  id: 'needle-shape',
  type: 'shape',
  transform: { x: 0, y: 0, width: 1_000_000, height: 500_000 },
  geometry: { preset: 'rect' },
  paragraphs: [{ runs: [{ text: 'Needle in this slide' }] }],
};

function documentModel(title: string, slideCount = 1): PresentationDocument {
  return {
    format: 'pptx',
    size: { widthEmu: 9_144_000, heightEmu: 6_858_000 },
    slides: Array.from({ length: slideCount }, (_, index) => ({
      id: `${title}-slide-${index}`,
      index,
      nodes: index === 0 ? [{ ...textShape }] : [],
    })),
    masters: [],
    layouts: [],
    themes: [],
    assets: {},
    warnings: [],
    metadata: { title },
  };
}

async function flushEffects(): Promise<void> {
  await act(async () => {
    await new Promise((resolve) => setTimeout(resolve, 0));
  });
}

async function waitFor(assertion: () => void): Promise<void> {
  let lastError: unknown;
  for (let attempt = 0; attempt < 30; attempt += 1) {
    try {
      assertion();
      return;
    } catch (error) {
      lastError = error;
      await flushEffects();
    }
  }
  throw lastError;
}

function installObserverStub(): void {
  class ObserverStub {
    constructor(_callback: IntersectionObserverCallback) {}
    observe() {}
    disconnect() {}
    unobserve() {}
    takeRecords(): IntersectionObserverEntry[] {
      return [];
    }
    readonly root = null;
    readonly rootMargin = '';
    readonly thresholds = [];
  }
  vi.stubGlobal('IntersectionObserver', ObserverStub);
}

function ThumbnailHarness({
  disabled = false,
  mountedSlideIndexes,
  onSlideRendered,
  renderWindow,
  resolution,
  source,
}: {
  disabled?: boolean;
  mountedSlideIndexes?: readonly number[];
  onSlideRendered?: (index: number, element: HTMLElement) => void;
  renderWindow?: PptxSlideThumbnailRenderWindow;
  resolution?: PptxSlideThumbnailResolution;
  source: PresentationDocument;
}) {
  const viewer = usePptxViewer();
  const { thumbnails } = usePptxViewerThumbnails(viewer.controller, {
    disabled,
    ...(renderWindow === undefined ? {} : { renderWindow }),
    ...(resolution === undefined ? {} : { resolution }),
  });
  const mounted = new Set(
    mountedSlideIndexes ?? thumbnails.map((thumbnail) => thumbnail.slideIndex),
  );
  return (
    <>
      <ReactPptxViewer
        ref={viewer.ref}
        source={source}
        mode="slide"
        {...(onSlideRendered ? { onSlideRendered } : {})}
      />
      <div data-testid="thumbnail-statuses">
        {thumbnails.map((thumbnail) => (
          <span
            key={thumbnail.slideIndex}
            data-slide-index={thumbnail.slideIndex}
            data-status={thumbnail.status}
          />
        ))}
      </div>
      <div data-testid="thumbnail-rail">
        {thumbnails
          .filter((thumbnail) => mounted.has(thumbnail.slideIndex))
          .map((thumbnail) => (
            <div
              key={thumbnail.slideIndex}
              ref={thumbnail.containerRef}
              data-slide-index={thumbnail.slideIndex}
              data-status={thumbnail.status}
              style={{ width: thumbnail.width, height: thumbnail.height }}
            />
          ))}
      </div>
    </>
  );
}

describe('ReactPptxViewer adapter lifecycle', () => {
  let host: HTMLDivElement;
  let root: Root;

  beforeEach(() => {
    (globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
    vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(null);
    host = document.createElement('div');
    document.body.append(host);
    root = createRoot(host);
  });

  afterEach(async () => {
    await act(async () => root.unmount());
    host.remove();
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it('publishes readiness before initial search, highlight, and thumbnails render', async () => {
    const onReady = vi.fn();
    const onSearchResults = vi.fn();
    const onThumbnailRendered = vi.fn();

    await act(async () => {
      root.render(
        <ReactPptxViewer
          source={documentModel('ready')}
          width={640}
          showThumbnails
          searchQuery="Needle"
          activeSearchResult={0}
          onReady={onReady}
          onSearchResults={onSearchResults}
          onThumbnailRendered={onThumbnailRendered}
        />,
      );
    });

    await waitFor(() => {
      expect(onReady).toHaveBeenCalledTimes(1);
      expect(onSearchResults).toHaveBeenLastCalledWith([
        expect.objectContaining({ slideIndex: 0, nodeId: 'needle-shape' }),
      ]);
      expect(onThumbnailRendered).toHaveBeenCalledWith(0, expect.any(HTMLElement));
      expect(host.querySelector('.rpv-search-highlight')).not.toBeNull();
    });
    expect(host.querySelector<HTMLElement>('.rpv-root')?.style.width).toBe('640px');
    expect(host.querySelector('[data-rpv-slide-wrapper]')).not.toBeNull();
  });

  it('mounts only the built-in thumbnail rows inside the filmstrip window', async () => {
    const source = documentModel('virtual-filmstrip', 100);
    await act(async () =>
      root.render(<ReactPptxViewer source={source} mode="slide" showThumbnails />),
    );
    await waitFor(() => {
      const mountedRows = host.querySelectorAll('.rpv-filmstrip__item');
      expect(mountedRows.length).toBeGreaterThan(0);
      expect(mountedRows.length).toBeLessThan(source.slides.length);
    });

    const filmstrip = host.querySelector<HTMLElement>('.rpv-filmstrip')!;
    await act(async () => {
      filmstrip.scrollTop = 12_000;
      filmstrip.dispatchEvent(new Event('scroll'));
    });
    await waitFor(() => {
      expect(host.querySelector('[aria-label="Go to slide 100"]')).not.toBeNull();
      expect(host.querySelectorAll('.rpv-filmstrip__item').length).toBeLessThan(
        source.slides.length,
      );
    });
  });

  it('exposes detached thumbnails for a consumer-owned rail', async () => {
    const source = documentModel('thumbnail-hook', 2);
    const render = (resolution: PptxSlideThumbnailResolution, disabled = false) => (
      <StrictMode>
        <ThumbnailHarness source={source} resolution={resolution} disabled={disabled} />
      </StrictMode>
    );

    await act(async () => root.render(render(160)));
    await waitFor(() => {
      const rail = host.querySelector<HTMLElement>('[data-testid="thumbnail-rail"]')!;
      expect(rail.children).toHaveLength(2);
      expect(rail.querySelectorAll('[data-status="ready"]')).toHaveLength(2);
      expect(rail.querySelectorAll('[data-rpv-slide-wrapper]')).toHaveLength(2);
    });

    const firstRail = host.querySelector<HTMLElement>('[data-testid="thumbnail-rail"]')!;
    const firstThumbnail = firstRail.querySelector<HTMLElement>('[data-slide-index="0"]')!;
    const firstWrapper = firstThumbnail.querySelector<HTMLElement>('[data-rpv-slide-wrapper]')!;
    expect(firstThumbnail.style.width).toBe('160px');
    expect(firstThumbnail.style.height).toBe('120px');
    expect(firstWrapper.style.width).toBe('160px');
    expect(firstWrapper.style.height).toBe('120px');

    await act(async () => root.render(render({ maxHeight: 90 })));
    await waitFor(() => {
      const rail = host.querySelector<HTMLElement>('[data-testid="thumbnail-rail"]')!;
      expect(rail.querySelectorAll('[data-status="ready"]')).toHaveLength(2);
      expect(rail.querySelector<HTMLElement>('[data-slide-index="0"]')?.style.height).toBe('90px');
    });
    const resized = host.querySelector<HTMLElement>(
      '[data-testid="thumbnail-rail"] [data-slide-index="0"] [data-rpv-slide-wrapper]',
    )!;
    expect(resized.style.width).toBe('120px');
    expect(resized.style.height).toBe('90px');

    await act(async () => root.render(render({ maxHeight: 90 }, true)));
    await waitFor(() => {
      const rail = host.querySelector<HTMLElement>('[data-testid="thumbnail-rail"]')!;
      expect(rail.querySelectorAll('[data-status="idle"]')).toHaveLength(2);
      expect(rail.querySelectorAll('[data-rpv-slide-wrapper]')).toHaveLength(0);
    });
  });

  it('prefetches a render window and adopts the cached slide when its row mounts', async () => {
    const source = documentModel('virtual-thumbnail-hook', 5);
    const onSlideRendered = vi.fn();
    const render = (
      mountedSlideIndexes: readonly number[],
      renderWindow: PptxSlideThumbnailRenderWindow,
    ) => (
      <ThumbnailHarness
        source={source}
        mountedSlideIndexes={mountedSlideIndexes}
        renderWindow={renderWindow}
        onSlideRendered={onSlideRendered}
      />
    );

    await act(async () =>
      root.render(render([0], { visibleSlideIndexes: [0], prefetchSlideIndexes: [1] })),
    );
    await waitFor(() => {
      expect(
        host
          .querySelector('[data-testid="thumbnail-statuses"] [data-slide-index="1"]')
          ?.getAttribute('data-status'),
      ).toBe('ready');
      expect(
        host.querySelectorAll('[data-testid="thumbnail-rail"] [data-rpv-slide-wrapper]'),
      ).toHaveLength(1);
    });
    expect(onSlideRendered.mock.calls.filter(([index]) => index === 1)).toHaveLength(1);

    await act(async () =>
      root.render(render([1], { visibleSlideIndexes: [1], prefetchSlideIndexes: [2] })),
    );
    await waitFor(() => {
      const mountedThumbnail = host.querySelector<HTMLElement>(
        '[data-testid="thumbnail-rail"] [data-slide-index="1"]',
      );
      expect(mountedThumbnail?.querySelector('[data-rpv-slide-wrapper]')).not.toBeNull();
      expect(mountedThumbnail?.getAttribute('data-status')).toBe('ready');
    });
    expect(onSlideRendered.mock.calls.filter(([index]) => index === 1)).toHaveLength(1);
  });

  it('does not publish a stale adapter after its font preparation resolves', async () => {
    let resolveFirstPrepare: (() => void) | undefined;
    const prepare = vi
      .spyOn(PptxFontManager.prototype, 'prepare')
      .mockImplementationOnce(
        () =>
          new Promise<void>((resolve) => {
            resolveFirstPrepare = resolve;
          }),
      )
      .mockResolvedValue(undefined);
    const readyTitles: string[] = [];
    const onReady = vi.fn((controller) => {
      readyTitles.push(controller.getDocument()?.metadata?.title ?? 'missing');
    });
    const first = documentModel('first');
    const second = documentModel('second');

    await act(async () => {
      root.render(<ReactPptxViewer source={first} onReady={onReady} />);
    });
    await waitFor(() => expect(prepare).toHaveBeenCalledTimes(1));
    await act(async () => {
      root.render(<ReactPptxViewer source={second} onReady={onReady} />);
    });
    await waitFor(() => expect(readyTitles).toEqual(['second']));

    await act(async () => {
      resolveFirstPrepare?.();
      await Promise.resolve();
    });

    expect(onReady).toHaveBeenCalledTimes(1);
    expect(readyTitles).toEqual(['second']);
  });

  it('reacts to parse option values without looping on equivalent inline fetchInit objects', async () => {
    const source = documentModel('parse-options');
    const onLoad = vi.fn();
    const render = (header: string, maxInputBytes = 1000, formatHint: 'pptx' | 'ppt' = 'pptx') => (
      <ReactPptxViewer
        source={source}
        parseOptions={{
          formatHint,
          maxInputBytes,
          fetchInit: { headers: { 'x-test': header } },
        }}
        onLoad={onLoad}
      />
    );

    await act(async () => root.render(render('same')));
    await waitFor(() => expect(onLoad).toHaveBeenCalledTimes(1));
    await act(async () => root.render(render('same')));
    await flushEffects();
    expect(onLoad).toHaveBeenCalledTimes(1);

    await act(async () => root.render(render('changed')));
    await waitFor(() => expect(onLoad).toHaveBeenCalledTimes(2));
    await act(async () => root.render(render('changed', 2000)));
    await waitFor(() => expect(onLoad).toHaveBeenCalledTimes(3));
    await act(async () => root.render(render('changed', 2000, 'ppt')));
    await waitFor(() => expect(onLoad).toHaveBeenCalledTimes(4));
  });

  it('applies fitMode changes and rebuilds the window when virtualization props change', async () => {
    installObserverStub();
    const source = documentModel('dynamic-props', 3);
    const onViewportReady = (viewport: HTMLDivElement) => {
      Object.defineProperty(viewport, 'clientWidth', { configurable: true, value: 480 });
    };

    await act(async () => {
      root.render(
        <ReactPptxViewer
          source={source}
          mode="slide"
          fitMode="contain"
          virtualization={{ enabled: true, initialSlides: 1 }}
          onViewportReady={onViewportReady}
        />,
      );
    });
    await waitFor(() => {
      expect(host.querySelector<HTMLElement>('[data-rpv-slide-wrapper]')?.style.width).toBe(
        '480px',
      );
    });

    await act(async () => {
      root.render(
        <ReactPptxViewer
          source={source}
          mode="slide"
          fitMode="none"
          virtualization={{ enabled: true, initialSlides: 1 }}
          onViewportReady={onViewportReady}
        />,
      );
    });
    await waitFor(() => {
      expect(host.querySelector<HTMLElement>('[data-rpv-slide-wrapper]')?.style.width).toBe(
        '960px',
      );
    });

    await act(async () => {
      root.render(
        <ReactPptxViewer
          source={source}
          mode="continuous"
          fitMode="none"
          virtualization={{ enabled: true, overscanViewport: 0 }}
          onViewportReady={onViewportReady}
        />,
      );
    });
    // Windowed mode only mounts the visible slide plus one overscan slide.
    await waitFor(() => expect(host.querySelectorAll('[data-rpv-slide-wrapper]')).toHaveLength(2));

    await act(async () => {
      root.render(
        <ReactPptxViewer
          source={source}
          mode="continuous"
          fitMode="none"
          virtualization={false}
          onViewportReady={onViewportReady}
        />,
      );
    });
    await waitFor(() => {
      expect(host.querySelectorAll('[data-rpv-slide-wrapper]')).toHaveLength(3);
      expect(host.querySelector<HTMLElement>('[data-rpv-slide-wrapper]')?.style.width).toBe(
        '960px',
      );
    });
  });

  it('shows the empty state without surfacing a render error', async () => {
    const onError = vi.fn();
    await act(async () => {
      root.render(
        <ReactPptxViewer
          source={documentModel('empty', 0)}
          emptyState="Nothing here"
          onError={onError}
        />,
      );
    });

    await waitFor(() => expect(host.textContent).toContain('Nothing here'));
    expect(onError).not.toHaveBeenCalled();
    expect(host.querySelector('[data-testid="pptx-viewport"]')?.childElementCount).toBe(0);
  });

  it('reports each controlled navigation once', async () => {
    const source = documentModel('controlled', 2);
    const onSlideChange = vi.fn();

    await act(async () => {
      root.render(
        <ReactPptxViewer
          source={source}
          mode="slide"
          slideIndex={1}
          onSlideChange={onSlideChange}
        />,
      );
    });
    await waitFor(() => expect(onSlideChange).toHaveBeenCalledTimes(1));
    expect(onSlideChange).toHaveBeenLastCalledWith(1);

    await act(async () => {
      root.render(
        <ReactPptxViewer
          source={source}
          mode="slide"
          slideIndex={0}
          onSlideChange={onSlideChange}
        />,
      );
    });
    await waitFor(() => expect(onSlideChange).toHaveBeenCalledTimes(2));
    expect(onSlideChange.mock.calls).toEqual([[1], [0]]);
  });
});
