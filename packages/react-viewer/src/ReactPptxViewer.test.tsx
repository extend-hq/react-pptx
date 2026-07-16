import type { PresentationDocument, ShapeNode } from '@extend-ai/react-pptx-model';
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { PptxFontManager } from './fonts';
import { ReactPptxViewer } from './ReactPptxViewer';

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

describe('ReactPptxViewer adapter lifecycle', () => {
  let host: HTMLDivElement;
  let root: Root;

  beforeEach(() => {
    (globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
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
