import {
  forwardRef,
  useCallback,
  useEffect,
  useImperativeHandle,
  useMemo,
  useRef,
  useState,
  type CSSProperties,
} from 'react';
import { useVirtualizer, type VirtualizerOptions } from '@tanstack/react-virtual';
import type { PresentationSearchResult } from '@extend-ai/react-pptx-model';
import { PptxViewerError, toPptxViewerError } from './errors';
import { PptxFontManager } from './fonts';
import { NormalizedPresentationViewer } from './normalized-viewer';
import { parsePresentation } from './parse';
import type {
  FitMode,
  ParsedPresentation,
  PptxViewerController,
  ReactPptxViewerProps,
  SearchHighlightOptions,
  ViewerMode,
  ViewerSearchOptions,
} from './types';

interface ViewerAdapter {
  render(mode: ViewerMode, slideIndex: number): Promise<void>;
  goToSlide(index: number, options?: ScrollIntoViewOptions): Promise<void>;
  setZoom(percent: number): Promise<void>;
  setFitMode(mode: FitMode): Promise<void>;
  search(query: string | RegExp, options?: ViewerSearchOptions): PresentationSearchResult[];
  highlight(result: PresentationSearchResult, options?: SearchHighlightOptions): Promise<void>;
  clearHighlights(): void;
  renderThumbnail(index: number, target: HTMLElement, width: number): Promise<() => void>;
  destroy(): void;
}

interface AdapterState {
  adapter: ViewerAdapter;
  generation: number;
}

const THUMBNAIL_ROW_ESTIMATE = 123;
const THUMBNAIL_OVERSCAN = 1;
const THUMBNAIL_INITIAL_RECT = { height: 780, width: 184 };
const observeThumbnailFilmstripRect: NonNullable<
  VirtualizerOptions<HTMLElement, HTMLElement>['observeElementRect']
> = (instance, callback) => {
  const element = instance.scrollElement;
  if (!element) return;
  const publish = () => {
    const rect = element.getBoundingClientRect();
    callback({
      height: rect.height || element.clientHeight || THUMBNAIL_INITIAL_RECT.height,
      width: rect.width || element.clientWidth || THUMBNAIL_INITIAL_RECT.width,
    });
  };
  publish();
  if (typeof ResizeObserver === 'undefined') return;
  const observer = new ResizeObserver(publish);
  observer.observe(element);
  return () => observer.disconnect();
};

const clampSlide = (index: number, count: number): number =>
  Math.max(0, Math.min(Math.max(0, count - 1), Math.floor(index)));

function viewerShellStyle(height: number | string | undefined): CSSProperties {
  return {
    '--rpv-height': typeof height === 'number' ? `${height}px` : (height ?? 'min(76vh, 780px)'),
  } as CSSProperties;
}

function stableRequestValue(value: unknown, seen = new WeakSet<object>()): unknown {
  if (value === null || typeof value !== 'object') return value;
  if (typeof Headers !== 'undefined' && value instanceof Headers) {
    return [...value.entries()].sort(([first], [second]) => first.localeCompare(second));
  }
  if (typeof URLSearchParams !== 'undefined' && value instanceof URLSearchParams) {
    return [...value.entries()];
  }
  if (value instanceof Date) return value.toISOString();
  if (value instanceof ArrayBuffer) return [...new Uint8Array(value)];
  if (ArrayBuffer.isView(value)) {
    return [...new Uint8Array(value.buffer, value.byteOffset, value.byteLength)];
  }
  if (typeof Blob !== 'undefined' && value instanceof Blob) {
    return { type: value.type, size: value.size };
  }
  if (seen.has(value)) return '[circular]';
  seen.add(value);
  if (Array.isArray(value)) return value.map((item) => stableRequestValue(item, seen));
  return Object.fromEntries(
    Object.entries(value)
      .sort(([first], [second]) => first.localeCompare(second))
      .map(([key, entry]) => [key, stableRequestValue(entry, seen)]),
  );
}

function requestInitKey(value: Omit<RequestInit, 'signal'> | undefined): string {
  if (!value) return '';
  try {
    return JSON.stringify(stableRequestValue(value)) ?? '';
  } catch {
    return String(value);
  }
}

async function createAdapter(
  target: HTMLElement,
  parsed: ParsedPresentation,
  props: ReactPptxViewerProps,
  getCurrentProps: () => ReactPptxViewerProps,
  onCurrentSlideChange: (index: number) => void,
  reportWarning: (warning: import('@extend-ai/react-pptx-model').PresentationWarning) => void,
  isCurrent: () => boolean,
): Promise<ViewerAdapter> {
  const fontManager = new PptxFontManager(props.fonts, (warning) => {
    if (isCurrent()) reportWarning(warning);
  });
  await fontManager.prepare(parsed.document);
  if (!isCurrent()) {
    fontManager.destroy();
    throw new Error('Viewer adapter creation was superseded.');
  }
  let destroyed = false;
  const active = () => !destroyed && isCurrent();
  const callbacks = {
    onSlideChange: (index: number) => {
      if (active()) onCurrentSlideChange(index);
    },
    onSlideRendered: (index: number, element: HTMLElement) => {
      if (!active()) return;
      fontManager.applyTo(element, index);
      getCurrentProps().onSlideRendered?.(index, element);
    },
    onSlideUnmounted: (index: number) => {
      if (active()) getCurrentProps().onSlideUnmounted?.(index);
    },
    onNodeError: (nodeId: string, error: unknown) => {
      if (!active()) return;
      reportWarning({
        code: 'degraded-rendering',
        severity: 'warning',
        nodeId,
        feature: 'normalized-renderer',
        message: `A slide element could not be rendered: ${
          error instanceof Error ? error.message : String(error)
        }`,
      });
    },
  };

  const normalized = new NormalizedPresentationViewer(target, parsed.document, callbacks);
  return {
    async render(mode, slideIndex) {
      if (mode === 'continuous') {
        const virtualization = typeof props.virtualization === 'object' ? props.virtualization : {};
        await normalized.renderList({
          enabled: props.virtualization !== false,
          initialSlides: virtualization.initialSlides ?? 3,
          overscanViewport: virtualization.overscanViewport ?? 1.5,
          batchSize: virtualization.batchSize ?? 3,
          ...(virtualization.scrollElement ? { scrollElement: virtualization.scrollElement } : {}),
          initialSlideIndex: slideIndex,
        });
      } else {
        await normalized.renderSlide(slideIndex);
      }
    },
    goToSlide: (index, options) => normalized.goToSlide(index, options),
    setZoom: (percent) => normalized.setZoom(percent),
    setFitMode: (mode) => normalized.setFitMode(mode),
    search: (query, options) => normalized.searchText(query, options),
    highlight: (result, options) => normalized.highlightSearchResult(result, options),
    clearHighlights: () => normalized.clearSearchHighlights(),
    async renderThumbnail(index, thumbnailTarget, width) {
      if (!active()) return () => {};
      const handle = normalized.renderThumbnailToContainer(index, thumbnailTarget, { width });
      await handle.ready;
      if (!active()) {
        handle.dispose();
        return () => {};
      }
      fontManager.applyTo(thumbnailTarget, index);
      return () => handle.dispose();
    },
    destroy: () => {
      if (destroyed) return;
      destroyed = true;
      normalized.destroy();
      fontManager.destroy();
    },
  };
}

interface ThumbnailProps {
  adapter: ViewerAdapter;
  index: number;
  active: boolean;
  width: number;
  onSelect: () => void;
  onRendered?: (index: number, element: HTMLElement) => void;
}

function Thumbnail({ adapter, index, active, width, onSelect, onRendered }: ThumbnailProps) {
  const targetRef = useRef<HTMLDivElement>(null);
  useEffect(() => {
    const target = targetRef.current;
    if (!target) return;
    let disposed = false;
    let cleanup: (() => void) | undefined;
    void adapter
      .renderThumbnail(index, target, width)
      .then((nextCleanup) => {
        if (disposed) nextCleanup();
        else {
          cleanup = nextCleanup;
          onRendered?.(index, target);
        }
      })
      .catch(() => undefined);
    return () => {
      disposed = true;
      cleanup?.();
    };
  }, [adapter, index, onRendered, width]);
  return (
    <button
      type="button"
      className="rpv-thumbnail"
      data-active={active || undefined}
      aria-label={`Go to slide ${index + 1}`}
      aria-current={active ? 'page' : undefined}
      onClick={onSelect}
    >
      <span className="rpv-thumbnail__number">{String(index + 1).padStart(2, '0')}</span>
      <span className="rpv-thumbnail__canvas" ref={targetRef} />
    </button>
  );
}

export const ReactPptxViewer = forwardRef<PptxViewerController, ReactPptxViewerProps>(
  function ReactPptxViewer(props, forwardedRef) {
    const {
      source,
      mode = 'continuous',
      slideIndex: controlledSlide,
      initialSlide = 0,
      zoom: controlledZoom,
      defaultZoom = 100,
      fitMode = 'contain',
      showToolbar = false,
      showThumbnails = false,
      showNotes = false,
      showDiagnostics = false,
      showSlideLabels = false,
      searchQuery,
      searchOptions,
      activeSearchResult,
      searchHighlightOptions,
      className,
      style,
      width,
      height,
      virtualization,
      fonts,
      parseOptions,
      renderLoading,
      renderError,
      renderThumbnail,
      emptyState,
      toolbarClassName,
      viewportClassName,
      toolbarStyle,
      viewportStyle,
      onReady,
      onLoad,
      onError,
      onWarning,
      onSlideChange: _onSlideChange,
      onSlideRendered: _onSlideRendered,
      onSlideUnmounted: _onSlideUnmounted,
      onSearchResults,
      onThumbnailRendered,
      onViewportReady,
      ...divProps
    } = props;
    const viewportRef = useRef<HTMLDivElement>(null);
    const filmstripRef = useRef<HTMLElement>(null);
    const adapterRef = useRef<ViewerAdapter | undefined>(undefined);
    const adapterGenerationRef = useRef(0);
    const latestPropsRef = useRef(props);
    latestPropsRef.current = props;
    const [adapterState, setAdapterState] = useState<AdapterState | null>(null);
    const [parsed, setParsed] = useState<ParsedPresentation | null>(null);
    const [error, setError] = useState<PptxViewerError | null>(null);
    const [loading, setLoading] = useState(true);
    const [internalSlide, setInternalSlide] = useState(initialSlide);
    const [internalZoom, setInternalZoom] = useState(defaultZoom);
    const [results, setResults] = useState<PresentationSearchResult[]>([]);
    const [runtimeWarnings, setRuntimeWarnings] = useState<
      import('@extend-ai/react-pptx-model').PresentationWarning[]
    >([]);
    const virtualizationOptions = typeof virtualization === 'object' ? virtualization : {};
    const virtualizationEnabled =
      virtualization !== false && virtualizationOptions.enabled !== false;
    const initialSlides = virtualizationOptions.initialSlides ?? 3;
    const overscanViewport = virtualizationOptions.overscanViewport ?? 1.5;
    const batchSize = virtualizationOptions.batchSize ?? 3;
    const scrollElement = virtualizationOptions.scrollElement ?? null;
    const parseFormatHint = parseOptions?.formatHint;
    const parseMaxInputBytes = parseOptions?.maxInputBytes;
    const parseFetchInit = parseOptions?.fetchInit;
    const parseFetchInitKey = requestInitKey(parseFetchInit);
    const slide = clampSlide(controlledSlide ?? internalSlide, parsed?.document.slides.length ?? 0);
    const zoom = controlledZoom ?? internalZoom;
    const latestSlideRef = useRef(slide);
    const latestZoomRef = useRef(zoom);
    const latestFitModeRef = useRef(fitMode);
    /** Last slide index the viewer itself reported through onSlideChange. */
    const lastViewerSlideRef = useRef<number | undefined>(undefined);
    latestSlideRef.current = slide;
    latestZoomRef.current = zoom;
    latestFitModeRef.current = fitMode;
    const thumbnailCount = showThumbnails ? (parsed?.document.slides.length ?? 0) : 0;
    const getThumbnailKey = useCallback(
      (index: number) => parsed?.document.slides[index]?.id ?? index,
      [parsed],
    );
    const thumbnailVirtualizer = useVirtualizer({
      count: thumbnailCount,
      enabled: thumbnailCount > 0,
      estimateSize: () => THUMBNAIL_ROW_ESTIMATE,
      getItemKey: getThumbnailKey,
      getScrollElement: () => filmstripRef.current,
      initialRect: THUMBNAIL_INITIAL_RECT,
      observeElementRect: observeThumbnailFilmstripRect,
      overscan: THUMBNAIL_OVERSCAN,
      useFlushSync: false,
    });
    const virtualThumbnails = thumbnailVirtualizer.getVirtualItems();

    useEffect(() => {
      if (!showThumbnails || thumbnailCount === 0) return;
      thumbnailVirtualizer.scrollToIndex(slide, { align: 'auto' });
    }, [showThumbnails, slide, thumbnailCount, thumbnailVirtualizer]);

    const controller = useMemo<PptxViewerController>(
      () => ({
        async goToSlide(index, options) {
          const next = clampSlide(index, parsed?.document.slides.length ?? 0);
          setInternalSlide(next);
          await adapterRef.current?.goToSlide(next, options);
        },
        async next() {
          await this.goToSlide(latestSlideRef.current + 1, {
            behavior: 'smooth',
            block: 'center',
          });
        },
        async previous() {
          await this.goToSlide(latestSlideRef.current - 1, {
            behavior: 'smooth',
            block: 'center',
          });
        },
        async setZoom(percent) {
          const next = Math.max(10, Math.min(400, percent));
          setInternalZoom(next);
          await adapterRef.current?.setZoom(next);
        },
        async setFitMode(nextMode) {
          await adapterRef.current?.setFitMode(nextMode);
        },
        search(query, options) {
          return adapterRef.current?.search(query, options) ?? [];
        },
        async highlightSearchResult(result, options) {
          setInternalSlide(result.slideIndex);
          await adapterRef.current?.highlight(result, options);
        },
        clearSearchHighlights() {
          adapterRef.current?.clearHighlights();
        },
        isReady: () => Boolean(adapterRef.current),
        async renderThumbnail(index, target, options) {
          const adapter = adapterRef.current;
          if (!adapter) throw new Error('The PowerPoint viewer is not ready to render thumbnails.');
          const count = parsed?.document.slides.length ?? 0;
          if (!Number.isInteger(index) || index < 0 || index >= count) {
            throw new RangeError(`Slide ${index} does not exist.`);
          }
          return adapter.renderThumbnail(index, target, options?.width ?? 144);
        },
        getDocument: () => parsed?.document ?? null,
        getSlideIndex: () => latestSlideRef.current,
        getZoom: () => latestZoomRef.current,
      }),
      [adapterState?.generation, parsed],
    );
    useImperativeHandle(forwardedRef, () => controller, [controller]);

    useEffect(() => {
      const abort = new AbortController();
      adapterGenerationRef.current += 1;
      const previousAdapter = adapterRef.current;
      adapterRef.current = undefined;
      previousAdapter?.destroy();
      setAdapterState(null);
      setLoading(true);
      setError(null);
      setParsed(null);
      setRuntimeWarnings([]);
      void parsePresentation(source, {
        signal: abort.signal,
        ...(parseFormatHint ? { formatHint: parseFormatHint } : {}),
        ...(parseMaxInputBytes !== undefined ? { maxInputBytes: parseMaxInputBytes } : {}),
        ...(parseFetchInit ? { fetchInit: parseFetchInit } : {}),
      })
        .then((next) => {
          if (abort.signal.aborted) return;
          setParsed(next);
          latestPropsRef.current.onLoad?.(next);
          for (const warning of next.warnings) latestPropsRef.current.onWarning?.(warning);
        })
        .catch((reason: unknown) => {
          if (abort.signal.aborted) return;
          const nextError = toPptxViewerError(
            reason,
            'parse-failed',
            'Could not open the presentation.',
          );
          setError(nextError);
          latestPropsRef.current.onError?.(nextError);
        })
        .finally(() => {
          if (!abort.signal.aborted) setLoading(false);
        });
      return () => abort.abort();
    }, [source, parseFormatHint, parseMaxInputBytes, parseFetchInitKey]);

    useEffect(() => {
      const viewport = viewportRef.current;
      if (viewport) onViewportReady?.(viewport);
    }, [onViewportReady]);

    useEffect(() => {
      const target = viewportRef.current;
      if (!parsed || !target || error) return;
      const generation = ++adapterGenerationRef.current;
      let disposed = false;
      let createdAdapter: ViewerAdapter | undefined;
      const isCurrent = () => !disposed && adapterGenerationRef.current === generation && !error;
      const adapterProps: ReactPptxViewerProps = {
        ...props,
        virtualization: !virtualizationEnabled
          ? false
          : {
              enabled: virtualizationEnabled,
              initialSlides,
              overscanViewport,
              batchSize,
              ...(scrollElement ? { scrollElement } : {}),
            },
      };
      setAdapterState(null);
      void (async () => {
        try {
          const adapter = await createAdapter(
            target,
            parsed,
            adapterProps,
            () => latestPropsRef.current,
            (index) => {
              if (!isCurrent()) return;
              lastViewerSlideRef.current = index;
              setInternalSlide(index);
              latestPropsRef.current.onSlideChange?.(index);
            },
            (warning) => {
              if (!isCurrent()) return;
              setRuntimeWarnings((current) =>
                current.some(
                  (item) => item.code === warning.code && item.feature === warning.feature,
                )
                  ? current
                  : [...current, warning],
              );
              latestPropsRef.current.onWarning?.(warning);
            },
            isCurrent,
          );
          createdAdapter = adapter;
          if (!isCurrent()) {
            adapter.destroy();
            return;
          }
          await adapter.setZoom(latestZoomRef.current);
          if (!isCurrent()) {
            adapter.destroy();
            return;
          }
          await adapter.setFitMode(latestFitModeRef.current);
          if (!isCurrent()) {
            adapter.destroy();
            return;
          }
          await adapter.render(mode, latestSlideRef.current);
          if (!isCurrent()) {
            adapter.destroy();
            return;
          }
          adapterRef.current = adapter;
          setAdapterState({ adapter, generation });
          if (!isCurrent()) return;
          latestPropsRef.current.onReady?.(controller);
        } catch (reason: unknown) {
          if (!isCurrent()) return;
          const nextError = toPptxViewerError(
            reason,
            'render-failed',
            'Could not render the presentation.',
          );
          setError(nextError);
          latestPropsRef.current.onError?.(nextError);
        }
      })();
      return () => {
        disposed = true;
        if (adapterGenerationRef.current === generation) adapterGenerationRef.current += 1;
        if (adapterRef.current === createdAdapter) adapterRef.current = undefined;
        createdAdapter?.destroy();
        setAdapterState((current) => (current?.generation === generation ? null : current));
      };
    }, [
      parsed,
      error,
      mode,
      fonts,
      virtualizationEnabled,
      initialSlides,
      overscanViewport,
      batchSize,
      scrollElement,
      width,
    ]);

    useEffect(() => {
      // Skip echoes: controlled hosts feed onSlideChange back into slideIndex
      // while the user scrolls, and navigating to the already-visible slide
      // would snap the scroll position.
      if (adapterState && controlledSlide !== undefined && slide !== lastViewerSlideRef.current) {
        void adapterState.adapter.goToSlide(slide, { behavior: 'instant' });
      }
    }, [adapterState, controlledSlide, slide]);
    useEffect(() => {
      if (adapterState && controlledZoom !== undefined) {
        void adapterState.adapter.setZoom(zoom);
      }
    }, [adapterState, controlledZoom, zoom]);
    useEffect(() => {
      if (adapterState) void adapterState.adapter.setFitMode(fitMode);
    }, [adapterState, fitMode]);
    useEffect(() => {
      const adapter = adapterState?.adapter;
      if (!adapter || searchQuery === undefined || searchQuery === '') {
        setResults([]);
        onSearchResults?.([]);
        adapter?.clearHighlights();
        return;
      }
      const next = adapter.search(searchQuery, searchOptions);
      setResults(next);
      onSearchResults?.(next);
    }, [adapterState, searchQuery, searchOptions, onSearchResults]);
    useEffect(() => {
      const adapter = adapterState?.adapter;
      if (!adapter) return;
      if (activeSearchResult === null || activeSearchResult === undefined) {
        adapter.clearHighlights();
        return;
      }
      const result =
        typeof activeSearchResult === 'number' ? results[activeSearchResult] : activeSearchResult;
      if (result) void controller.highlightSearchResult(result, searchHighlightOptions);
    }, [adapterState, activeSearchResult, results, searchHighlightOptions, controller]);

    const go = useCallback(
      (index: number) => {
        setInternalSlide(index);
        void controller.goToSlide(index, { behavior: 'smooth', block: 'center' });
      },
      [controller],
    );
    const notes = parsed?.document.slides[slide]?.notes ?? [];
    return (
      <div
        {...divProps}
        className={['rpv-root', className].filter(Boolean).join(' ')}
        style={{ ...viewerShellStyle(height), ...(width !== undefined ? { width } : {}), ...style }}
        data-rpv-scroll-owner={scrollElement ? 'external' : 'internal'}
      >
        {showToolbar && parsed ? (
          <div
            className={['rpv-toolbar', toolbarClassName].filter(Boolean).join(' ')}
            style={toolbarStyle}
          >
            <div className="rpv-toolbar__cluster">
              <button
                type="button"
                onClick={() => void controller.previous()}
                disabled={slide === 0}
                aria-label="Previous slide"
              >
                ←
              </button>
              <span className="rpv-toolbar__count">
                {slide + 1} / {parsed.document.slides.length}
              </span>
              <button
                type="button"
                onClick={() => void controller.next()}
                disabled={slide >= parsed.document.slides.length - 1}
                aria-label="Next slide"
              >
                →
              </button>
            </div>
            <div className="rpv-toolbar__cluster">
              <button
                type="button"
                onClick={() => void controller.setZoom(zoom - 10)}
                aria-label="Zoom out"
              >
                −
              </button>
              <span className="rpv-toolbar__count">{zoom}%</span>
              <button
                type="button"
                onClick={() => void controller.setZoom(zoom + 10)}
                aria-label="Zoom in"
              >
                +
              </button>
            </div>
          </div>
        ) : null}
        <div className="rpv-workspace">
          {showThumbnails && parsed ? (
            <nav ref={filmstripRef} className="rpv-filmstrip" aria-label="Slide thumbnails">
              <div
                className="rpv-filmstrip__sizer"
                style={{ height: thumbnailVirtualizer.getTotalSize() }}
              >
                {virtualThumbnails.map((virtualThumbnail) => {
                  const item = parsed.document.slides[virtualThumbnail.index];
                  if (!item) return null;
                  const index = virtualThumbnail.index;
                  return (
                    <div
                      key={virtualThumbnail.key}
                      ref={thumbnailVirtualizer.measureElement}
                      className="rpv-filmstrip__item"
                      data-index={index}
                      style={{ transform: `translateY(${virtualThumbnail.start}px)` }}
                    >
                      {renderThumbnail ? (
                        renderThumbnail({
                          slideIndex: index,
                          slideCount: parsed.document.slides.length,
                          isCurrent: index === slide,
                          goToSlide: () => go(index),
                        })
                      ) : adapterState ? (
                        <Thumbnail
                          adapter={adapterState.adapter}
                          index={index}
                          active={index === slide}
                          width={144}
                          onSelect={() => go(index)}
                          {...(onThumbnailRendered ? { onRendered: onThumbnailRendered } : {})}
                        />
                      ) : null}
                    </div>
                  );
                })}
              </div>
            </nav>
          ) : null}
          <main className="rpv-stage">
            {loading ? (
              <div className="rpv-status">
                {renderLoading?.() ?? (
                  <>
                    <span className="rpv-spinner" />
                    Opening presentation…
                  </>
                )}
              </div>
            ) : null}
            {error ? (
              <div className="rpv-status rpv-status--error">
                {renderError?.(error) ?? (
                  <>
                    <strong>Couldn’t open this deck.</strong>
                    <span>{error.message}</span>
                  </>
                )}
              </div>
            ) : null}
            {!loading && !error && parsed?.document.slides.length === 0 ? (
              <div className="rpv-status">{emptyState ?? 'This presentation has no slides.'}</div>
            ) : null}
            <div
              ref={viewportRef}
              className={['rpv-viewport', viewportClassName].filter(Boolean).join(' ')}
              style={viewportStyle}
              data-testid="pptx-viewport"
              data-rpv-scroll-owner={scrollElement ? 'external' : 'internal'}
              aria-live="polite"
            />
            {showSlideLabels && parsed ? (
              <div className="rpv-slide-label">Slide {slide + 1}</div>
            ) : null}
          </main>
          {showDiagnostics && parsed ? (
            <aside className="rpv-diagnostics" aria-label="Rendering diagnostics">
              <h2>Diagnostics</h2>
              <dl>
                <div>
                  <dt>Format</dt>
                  <dd>{parsed.document.format.toUpperCase()}</dd>
                </div>
                <div>
                  <dt>Slides</dt>
                  <dd>{parsed.document.slides.length}</dd>
                </div>
                <div>
                  <dt>Warnings</dt>
                  <dd>{parsed.warnings.length + runtimeWarnings.length}</dd>
                </div>
              </dl>
              {parsed.warnings.map((warning, index) => (
                <p key={`${warning.code}-${index}`} data-severity={warning.severity}>
                  <strong>{warning.code}</strong>
                  {warning.message}
                </p>
              ))}
              {runtimeWarnings.map((warning, index) => (
                <p
                  key={`runtime-${warning.code}-${warning.feature ?? index}`}
                  data-severity={warning.severity}
                >
                  <strong>{warning.code}</strong>
                  {warning.message}
                </p>
              ))}
            </aside>
          ) : null}
        </div>
        {showNotes && notes.length ? (
          <section className="rpv-notes">
            <h2>Speaker notes</h2>
            {notes.map((note, index) => (
              <p key={index}>{note.text}</p>
            ))}
          </section>
        ) : null}
      </div>
    );
  },
);
