import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import wasmUrl from '@extend-ai/react-pptx/pptx_wasm_bg.wasm?url';
import { useVirtualizer } from '@tanstack/react-virtual';
import {
  ReactPptxViewer,
  setWasmSource,
  type ParsedPresentation,
  type PresentationSearchResult,
  type PresentationSource,
  type ViewerMode,
  usePptxViewer,
  usePptxViewerThumbnails,
} from '@extend-ai/react-pptx';
import { Icon } from './icons';

setWasmSource(wasmUrl);

const SAMPLE_SOURCE = '/viewer-smoke.pptx';
const ZOOM_OPTIONS = [50, 75, 90, 100, 110, 125, 150, 175, 200] as const;
const MIN_ZOOM = ZOOM_OPTIONS[0];
const MAX_ZOOM = ZOOM_OPTIONS[ZOOM_OPTIONS.length - 1] ?? 200;
const THUMBNAIL_ROW_ESTIMATE = 134;
const THUMBNAIL_OVERSCAN = 3;
const THUMBNAIL_PREFETCH = 3;

function buildThumbnailRenderWindow(visibleIndexes: readonly number[], slideCount: number) {
  const visibleSlideIndexes = [...new Set(visibleIndexes)]
    .filter((index) => index >= 0 && index < slideCount)
    .sort((left, right) => left - right);
  if (!visibleSlideIndexes.length) {
    return { visibleSlideIndexes, prefetchSlideIndexes: [] };
  }

  const visible = new Set(visibleSlideIndexes);
  const prefetchSlideIndexes: number[] = [];
  const first = visibleSlideIndexes[0] ?? 0;
  const last = visibleSlideIndexes.at(-1) ?? first;
  for (let distance = 1; distance <= THUMBNAIL_PREFETCH; distance += 1) {
    for (const index of [first - distance, last + distance]) {
      if (index >= 0 && index < slideCount && !visible.has(index)) {
        prefetchSlideIndexes.push(index);
      }
    }
  }
  return { visibleSlideIndexes, prefetchSlideIndexes };
}

export function App() {
  const viewer = usePptxViewer();
  const inputRef = useRef<HTMLInputElement>(null);
  const [source, setSource] = useState<PresentationSource>(SAMPLE_SOURCE);
  const [fileName, setFileName] = useState('viewer-smoke.pptx');
  const [mode, setMode] = useState<ViewerMode>('continuous');
  const [zoom, setZoom] = useState(100);
  const [slide, setSlide] = useState(0);
  const [slideCount, setSlideCount] = useState(0);
  const [search, setSearch] = useState('');
  const [results, setResults] = useState<readonly PresentationSearchResult[]>([]);
  const [activeResult, setActiveResult] = useState<number | null>(null);
  const [diagnostics, setDiagnostics] = useState(false);
  const [thumbnailsOpen, setThumbnailsOpen] = useState(true);
  const [dragging, setDragging] = useState(false);
  const [parsed, setParsed] = useState<ParsedPresentation | null>(null);
  const [dark, setDark] = useState(false);
  const thumbnailScrollRef = useRef<HTMLDivElement>(null);
  const thumbnailVirtualizer = useVirtualizer({
    count: slideCount,
    enabled: thumbnailsOpen && slideCount > 0,
    estimateSize: () => THUMBNAIL_ROW_ESTIMATE,
    getScrollElement: () => thumbnailScrollRef.current,
    initialRect: { height: 640, width: 204 },
    overscan: THUMBNAIL_OVERSCAN,
    useFlushSync: false,
  });
  const virtualThumbnails = thumbnailVirtualizer.getVirtualItems();
  const virtualThumbnailIndexesKey = virtualThumbnails
    .map((thumbnail) => thumbnail.index)
    .join(',');
  const thumbnailRenderWindow = useMemo(
    () =>
      buildThumbnailRenderWindow(
        virtualThumbnailIndexesKey
          ? virtualThumbnailIndexesKey.split(',').map((index) => Number(index))
          : [],
        slideCount,
      ),
    [slideCount, virtualThumbnailIndexesKey],
  );
  const { thumbnails } = usePptxViewerThumbnails(viewer.controller, {
    disabled: !thumbnailsOpen,
    renderWindow: thumbnailRenderWindow,
    resolution: { maxWidth: 152, maxHeight: 114 },
  });

  useEffect(() => {
    if (!thumbnailsOpen || slideCount === 0) return;
    thumbnailVirtualizer.scrollToIndex(slide, { align: 'auto' });
  }, [slide, slideCount, thumbnailVirtualizer, thumbnailsOpen]);

  const openFile = useCallback((file?: File) => {
    if (!file) return;
    const extension = file.name.split('.').pop()?.toLowerCase();
    if (extension !== 'ppt' && extension !== 'pptx') return;
    setSource(file);
    setFileName(file.name);
    setSlide(0);
    setSearch('');
    setActiveResult(null);
  }, []);

  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.target instanceof HTMLInputElement || event.target instanceof HTMLTextAreaElement)
        return;
      if (event.key === 'ArrowRight' || event.key === 'PageDown') void viewer.controller?.next();
      if (event.key === 'ArrowLeft' || event.key === 'PageUp') void viewer.controller?.previous();
    };
    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, [viewer.controller]);

  const stepResult = (direction: 1 | -1) => {
    if (!results.length) return;
    const next =
      ((activeResult ?? (direction > 0 ? -1 : 0)) + direction + results.length) % results.length;
    setActiveResult(next);
  };

  const resetSample = () => {
    setSource(`${SAMPLE_SOURCE}?t=${Date.now()}`);
    setFileName('viewer-smoke.pptx');
    setSlide(0);
    setSearch('');
    setActiveResult(null);
  };

  return (
    <div className={dark ? 'playground-shell dark' : 'playground-shell'}>
      <div className="playground-layout">
        <section className="toolbar-card" aria-label="Presentation toolbar">
          <div className="status-row">
            <div className="file-summary">
              <span className="file-icon">
                <Icon name="file" />
              </span>
              <div>
                <strong>{fileName}</strong>
                <span>
                  {parsed
                    ? `${parsed.document.format.toUpperCase()} · ${slideCount} slides`
                    : 'Opening presentation…'}
                </span>
              </div>
            </div>
            <div className="status-actions">
              <span className="status-text">{parsed ? 'Ready' : 'Loading…'}</span>
              <button className="button outline small" type="button" onClick={() => setDark(!dark)}>
                {dark ? 'Light' : 'Dark'} theme
              </button>
              <button
                className="button outline small"
                type="button"
                data-active={diagnostics || undefined}
                onClick={() => setDiagnostics(!diagnostics)}
              >
                <Icon name="diagnostics" />
                Diagnostics
              </button>
            </div>
          </div>

          <div className="toolbar-row">
            <div className="button-group" role="group" aria-label="View mode">
              <button
                className="button outline"
                type="button"
                data-active={mode === 'continuous' || undefined}
                onClick={() => setMode('continuous')}
              >
                Continuous
              </button>
              <button
                className="button outline"
                type="button"
                data-active={mode === 'slide' || undefined}
                onClick={() => setMode('slide')}
              >
                Single slide
              </button>
            </div>

            <div className="button-group" role="group" aria-label="Slide navigation">
              <button
                className="button outline icon-only"
                type="button"
                aria-label="Previous slide"
                disabled={slide === 0}
                onClick={() => void viewer.controller?.previous()}
              >
                <Icon name="arrow-left" />
              </button>
              <span className="page-count">
                {slideCount ? slide + 1 : 0} / {slideCount}
              </span>
              <button
                className="button outline icon-only"
                type="button"
                aria-label="Next slide"
                disabled={slide >= slideCount - 1}
                onClick={() => void viewer.controller?.next()}
              >
                <Icon name="arrow-right" />
              </button>
            </div>

            <label className="search-control">
              <Icon name="search" />
              <input
                value={search}
                onChange={(event) => {
                  setSearch(event.target.value);
                  setActiveResult(null);
                }}
                placeholder="Search slide text"
                aria-label="Search slide text"
              />
              {search ? <span>{results.length} hits</span> : null}
            </label>

            <div className="button-group search-buttons" role="group" aria-label="Search results">
              <button
                className="button outline icon-only"
                type="button"
                aria-label="Previous search result"
                disabled={!results.length}
                onClick={() => stepResult(-1)}
              >
                <Icon name="arrow-left" />
              </button>
              <button
                className="button outline icon-only"
                type="button"
                aria-label="Next search result"
                disabled={!results.length}
                onClick={() => stepResult(1)}
              >
                <Icon name="arrow-right" />
              </button>
            </div>

            <div className="button-group" role="group" aria-label="Zoom">
              <button
                className="button outline icon-only"
                type="button"
                aria-label="Zoom out"
                disabled={zoom <= MIN_ZOOM}
                onClick={() => setZoom(Math.max(MIN_ZOOM, zoom - 10))}
              >
                <Icon name="minus" />
              </button>
              <label className="select-control">
                <span className="sr-only">Zoom</span>
                <select value={zoom} onChange={(event) => setZoom(Number(event.target.value))}>
                  {ZOOM_OPTIONS.map((value) => (
                    <option key={value} value={value}>
                      {value}%
                    </option>
                  ))}
                </select>
              </label>
              <button
                className="button outline icon-only"
                type="button"
                aria-label="Zoom in"
                disabled={zoom >= MAX_ZOOM}
                onClick={() => setZoom(Math.min(MAX_ZOOM, zoom + 10))}
              >
                <Icon name="plus" />
              </button>
            </div>

            <div className="button-group">
              <button
                className="button outline"
                type="button"
                data-active={thumbnailsOpen || undefined}
                aria-expanded={thumbnailsOpen}
                aria-controls="playground-slide-thumbnails"
                onClick={() => setThumbnailsOpen(!thumbnailsOpen)}
              >
                <Icon name="slides" />
                Slides
              </button>
              <button className="button outline" type="button" onClick={resetSample}>
                <Icon name="replace" />
                Reset
              </button>
            </div>

            <div className="button-group toolbar-end">
              <div className="readonly-control">
                <span>Read only</span>
                <span className="switch" role="switch" aria-checked="true">
                  <span />
                </span>
              </div>
              <button
                className="button primary"
                type="button"
                onClick={() => inputRef.current?.click()}
              >
                <Icon name="upload" />
                Import
              </button>
              <input
                ref={inputRef}
                hidden
                type="file"
                accept=".ppt,.pptx,application/vnd.ms-powerpoint,application/vnd.openxmlformats-officedocument.presentationml.presentation"
                onChange={(event) => openFile(event.target.files?.[0])}
              />
            </div>
          </div>
        </section>

        <section
          id="viewer-workspace"
          className="viewer-panel"
          data-dragging={dragging || undefined}
          onDragEnter={(event) => {
            event.preventDefault();
            setDragging(true);
          }}
          onDragOver={(event) => event.preventDefault()}
          onDragLeave={(event) => {
            if (!event.currentTarget.contains(event.relatedTarget as Node | null))
              setDragging(false);
          }}
          onDrop={(event) => {
            event.preventDefault();
            setDragging(false);
            openFile(event.dataTransfer.files[0]);
          }}
        >
          <div className="viewer-content">
            {thumbnailsOpen ? (
              <nav
                id="playground-slide-thumbnails"
                className="slide-rail"
                aria-label="Slide thumbnails"
              >
                <div className="slide-rail__header">
                  <strong>Slides</strong>
                  <span>{slideCount}</span>
                </div>
                <div ref={thumbnailScrollRef} className="slide-rail__list">
                  {thumbnails.length ? (
                    <div
                      className="slide-rail__sizer"
                      style={{ height: thumbnailVirtualizer.getTotalSize() }}
                    >
                      {virtualThumbnails.map((virtualThumbnail) => {
                        const thumbnail = thumbnails[virtualThumbnail.index];
                        if (!thumbnail) return null;
                        return (
                          <div
                            key={virtualThumbnail.key}
                            ref={thumbnailVirtualizer.measureElement}
                            className="slide-rail__item"
                            data-index={virtualThumbnail.index}
                            style={{ transform: `translateY(${virtualThumbnail.start}px)` }}
                          >
                            <button
                              className="slide-thumbnail"
                              type="button"
                              data-active={thumbnail.slideIndex === slide || undefined}
                              data-status={thumbnail.status}
                              aria-current={thumbnail.slideIndex === slide ? 'page' : undefined}
                              aria-label={`Go to slide ${thumbnail.slideNumber}`}
                              aria-posinset={thumbnail.slideNumber}
                              aria-setsize={slideCount}
                              onClick={() =>
                                void viewer.controller?.goToSlide(thumbnail.slideIndex)
                              }
                            >
                              <span className="slide-thumbnail__number">
                                {thumbnail.slideNumber}
                              </span>
                              <span
                                className="slide-thumbnail__preview"
                                ref={thumbnail.containerRef}
                                style={{ width: thumbnail.width, height: thumbnail.height }}
                              />
                            </button>
                          </div>
                        );
                      })}
                    </div>
                  ) : (
                    <div className="slide-rail__empty">Preparing previews…</div>
                  )}
                </div>
              </nav>
            ) : null}
            <div className="viewer-host">
              <ReactPptxViewer
                ref={viewer.ref}
                source={source}
                mode={mode}
                zoom={zoom}
                slideIndex={slide}
                showDiagnostics={diagnostics}
                virtualization={{
                  enabled: true,
                  initialSlides: 3,
                  batchSize: 3,
                  overscanViewport: 1.5,
                }}
                searchQuery={search}
                activeSearchResult={activeResult}
                onLoad={(next) => {
                  setParsed(next);
                  setSlideCount(next.document.slides.length);
                }}
                onSlideChange={setSlide}
                onSearchResults={setResults}
                onError={(error) => console.error(error)}
                height="100%"
              />
            </div>
          </div>
          <div className="drop-overlay">
            <div>
              <Icon name="upload" />
              <span>
                <strong>Drop a PowerPoint file to import</strong>
                <small>PPT and PPTX files are parsed locally in your browser.</small>
              </span>
            </div>
          </div>
        </section>
      </div>
    </div>
  );
}
