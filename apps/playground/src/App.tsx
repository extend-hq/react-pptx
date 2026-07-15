import { useCallback, useEffect, useRef, useState } from 'react';
import wasmUrl from '@extend-ai/react-pptx/pptx_wasm_bg.wasm?url';
import {
  ReactPptxViewer,
  setWasmSource,
  type ParsedPresentation,
  type PptxViewerController,
  type PresentationSearchResult,
  type PresentationSource,
  type ViewerMode,
} from '@extend-ai/react-pptx';
import { Icon } from './icons';

setWasmSource(wasmUrl);

const SAMPLE_SOURCE = '/viewer-smoke.pptx';
const ZOOM_OPTIONS = [50, 75, 90, 100, 110, 125, 150, 175, 200] as const;
const MIN_ZOOM = ZOOM_OPTIONS[0];
const MAX_ZOOM = ZOOM_OPTIONS[ZOOM_OPTIONS.length - 1] ?? 200;

export function App() {
  const viewerRef = useRef<PptxViewerController>(null);
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
  const [thumbnailsOpen, setThumbnailsOpen] = useState(false);
  const [dragging, setDragging] = useState(false);
  const [parsed, setParsed] = useState<ParsedPresentation | null>(null);
  const [dark, setDark] = useState(false);

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
      if (event.key === 'ArrowRight' || event.key === 'PageDown') void viewerRef.current?.next();
      if (event.key === 'ArrowLeft' || event.key === 'PageUp') void viewerRef.current?.previous();
    };
    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, []);

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
                onClick={() => void viewerRef.current?.previous()}
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
                onClick={() => void viewerRef.current?.next()}
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
          <ReactPptxViewer
            ref={viewerRef}
            source={source}
            mode={mode}
            zoom={zoom}
            slideIndex={slide}
            showThumbnails={thumbnailsOpen}
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
