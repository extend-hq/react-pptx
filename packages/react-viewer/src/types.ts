import type {
  PresentationDocument,
  PresentationFormat,
  PresentationSearchResult,
  PresentationWarning,
} from '@extend-ai/react-pptx-model';
import type { CSSProperties, HTMLAttributes, ReactNode, Ref } from 'react';

export type BinaryPresentationSource = ArrayBuffer | Uint8Array | Blob | string | URL;

export interface ParsedPresentation {
  readonly kind: 'parsed-presentation';
  readonly document: PresentationDocument;
  readonly warnings: readonly PresentationWarning[];
}

export type PptxPerformancePhase =
  | 'download'
  | 'wasm-parsing'
  | 'embedded-font-preparation'
  | 'slide-dom-svg-rendering'
  | 'image-decoding'
  | 'raster-encoding';

export interface PptxPerformanceMeasurement {
  /** Wall-clock duration for this phase. */
  durationMs: number;
  phase: PptxPerformancePhase;
  /** Present for work attributable to one slide. */
  slideIndex?: number;
}

export type PresentationSource =
  BinaryPresentationSource | PresentationDocument | ParsedPresentation;

export type ViewerMode = 'slide' | 'continuous';
/** @deprecated Superseded by `ViewerZoomLevel` and the `zoom` prop. */
export type FitMode = 'contain' | 'none';

export type ViewerZoomMode = 'automatic' | 'fit-page' | 'fit-width';
export type ViewerZoomLevel = number | ViewerZoomMode;

export type ViewerZoomState = {
  level: ViewerZoomLevel;
  /** Resolved percentage; `100` means actual size. */
  resolvedZoom: number;
};

export interface ParsePresentationOptions {
  signal?: AbortSignal;
  formatHint?: PresentationFormat;
  maxInputBytes?: number;
  fetchInit?: Omit<RequestInit, 'signal'>;
  /** Receives cold-start download and native parsing timings. */
  onPerformanceMeasurement?: (measurement: PptxPerformanceMeasurement) => void;
}

export interface VirtualizationOptions {
  /**
   * Mount only slides near the viewport in continuous mode, windowed with
   * TanStack Virtual at fixed offsets so scrolling never shifts layout.
   * Default `true`.
   */
  enabled?: boolean;
  /**
   * @deprecated The virtualized window is viewport-driven; this option is
   * ignored and kept only for API compatibility.
   */
  initialSlides?: number;
  /** Overscan measured in viewport heights. Default `1.5`. */
  overscanViewport?: number;
  /** Slides rendered per cooperative batch. Default `3`. */
  batchSize?: number;
  /**
   * Scroll viewport owned by the host application. Pass the viewport element
   * from Radix, Base UI, coss, or another custom scroll-area implementation.
   */
  scrollElement?: HTMLElement | null;
}

export interface PptxFontSource {
  /** Font family used by PowerPoint text runs. */
  family: string;
  /** URL/CSS font source or font bytes. */
  source: string | ArrayBuffer | Uint8Array | Blob;
  descriptors?: FontFaceDescriptors;
}

export interface PptxFontOptions {
  /** Host-provided web fonts. Embedded PPTX fonts are loaded automatically. */
  sources?: readonly PptxFontSource[];
  /** Per-family fallback overrides, keyed case-sensitively or in lowercase. */
  fallbacks?: Readonly<Record<string, string | readonly string[]>>;
  /** Extra script-complete families appended to every stack. */
  fallbackFamilies?: readonly string[];
  /** Apply built-in Office/CJK/RTL fallback stacks. Default `true`. */
  useOfficeFallbacks?: boolean;
  /** Load supported fonts embedded in the PPTX. Default `true`. */
  loadEmbeddedFonts?: boolean;
  /** Wait for registered fonts before rendering slides. Default `true`. */
  waitForFonts?: boolean;
  /** Report unavailable requested families through `onWarning`. Default `true`. */
  reportMissingFonts?: boolean;
  /** Maximum wait for each font and `document.fonts.ready`. Default `5000`. */
  loadTimeoutMs?: number;
  onFontLoaded?: (family: string, face: FontFace) => void;
  onFontsReady?: () => void;
}

export interface ViewerSearchOptions {
  matchCase?: boolean;
  wholeWord?: boolean;
  useRegex?: boolean;
  snippetRadius?: number;
  includeShapes?: boolean;
  includeTables?: boolean;
  includeGroups?: boolean;
}

export interface SearchHighlightOptions {
  className?: string;
  borderColor?: string;
  backgroundColor?: string;
  boxShadow?: string;
  borderRadius?: number | string;
  borderWidth?: number | string;
  padding?: number;
  zIndex?: number;
  style?: Record<string, string | number | undefined>;
  scrollIntoView?: boolean | ScrollIntoViewOptions;
}

export interface ThumbnailRenderContext {
  slideIndex: number;
  slideCount: number;
  isCurrent: boolean;
  goToSlide: () => void;
}

export type PptxSlideThumbnailResolution =
  | number
  | {
      maxHeight?: number;
      maxWidth?: number;
    };

export interface PptxSlideThumbnailRenderOptions {
  /** Rendered thumbnail width in CSS pixels. */
  width?: number;
}

export type PptxThumbnailOutput = 'blob' | 'canvas' | 'imageBitmap' | 'svg';

export interface PptxThumbnailRendererOptions {
  /**
   * Thumbnail-specific font policy. Defaults to skipping embedded fonts and
   * not waiting for the global font set.
   */
  fonts?: PptxFontOptions;
  /** Maximum number of slide renders executing at once. Default `2`. */
  concurrency?: number;
  onPerformanceMeasurement?: (measurement: PptxPerformanceMeasurement) => void;
  onWarning?: (warning: PresentationWarning) => void;
}

export interface PptxThumbnailRenderOptions {
  maxHeight?: number;
  maxWidth?: number;
  /** Output pixel density. Default `1`. */
  pixelRatio?: number;
  output?: PptxThumbnailOutput;
  /** PNG encoding quality where supported. */
  quality?: number;
  signal?: AbortSignal;
}

export interface PptxThumbnailRenderResult<T> {
  data: T;
  height: number;
  measurements: readonly PptxPerformanceMeasurement[];
  output: PptxThumbnailOutput;
  slideIndex: number;
  width: number;
}

export interface PptxThumbnailRenderer {
  readonly presentation: ParsedPresentation;
  readonly ready: Promise<void>;
  renderSlide(
    slideIndex: number,
    options: PptxThumbnailRenderOptions & { output: 'canvas' },
  ): Promise<PptxThumbnailRenderResult<HTMLCanvasElement>>;
  renderSlide(
    slideIndex: number,
    options: PptxThumbnailRenderOptions & { output: 'imageBitmap' },
  ): Promise<PptxThumbnailRenderResult<ImageBitmap>>;
  renderSlide(
    slideIndex: number,
    options: PptxThumbnailRenderOptions & { output: 'svg' },
  ): Promise<PptxThumbnailRenderResult<string>>;
  renderSlide(
    slideIndex: number,
    options?: PptxThumbnailRenderOptions & { output?: 'blob' },
  ): Promise<PptxThumbnailRenderResult<Blob>>;
  /** Renders exactly the supplied indexes, preserving their order. */
  renderSlides(
    slideIndexes: readonly number[],
    options?: PptxThumbnailRenderOptions,
  ): Promise<PptxThumbnailRenderResult<Blob | HTMLCanvasElement | ImageBitmap | string>[]>;
  destroy(): void;
}

export interface PptxSlideThumbnailRenderWindow {
  /**
   * Slide indexes whose attached thumbnail containers should render first.
   *
   * Use this for the thumbnails currently mounted in a virtualized rail.
   */
  visibleSlideIndexes?: readonly number[];
  /**
   * Slide indexes to render into a detached cache after visible thumbnails.
   * Prefetched slides are adopted when their container mounts.
   */
  prefetchSlideIndexes?: readonly number[];
}

export interface UsePptxViewerThumbnailsOptions {
  /** Desired thumbnail bounds. A number constrains both dimensions. Default `160`. */
  resolution?: PptxSlideThumbnailResolution;
  /** Prioritizes and prefetches slides for consumer-owned virtualized thumbnail rails. */
  renderWindow?: PptxSlideThumbnailRenderWindow;
  /** Prevent rendering while preserving stable thumbnail metadata. Default `false`. */
  disabled?: boolean;
}

export type PptxSlideThumbnailStatus = 'idle' | 'rendering' | 'ready' | 'error';

export interface PptxSlideThumbnailItem {
  /** Source slide aspect ratio. */
  aspectRatio: number;
  /** Natural slide height in CSS pixels. */
  contentHeight: number;
  /** Natural slide width in CSS pixels. */
  contentWidth: number;
  /** Stable ref callback that renders into an attached consumer-owned element. */
  containerRef: (element: HTMLElement | null) => void;
  /** Last thumbnail rendering error, if any. */
  error?: Error;
  /** Rendered thumbnail height in CSS pixels. */
  height: number;
  /** Normalized source slide represented by this thumbnail. */
  slide: PresentationDocument['slides'][number];
  /** Zero-based slide index. */
  slideIndex: number;
  /** One-based slide number for display. */
  slideNumber: number;
  /** Current thumbnail render status. */
  status: PptxSlideThumbnailStatus;
  /** Renders this thumbnail into a consumer-owned element. */
  renderToContainer: (element: HTMLElement) => Promise<void>;
  /** Rendered thumbnail width in CSS pixels. */
  width: number;
}

export interface PptxViewerThumbnails {
  /** Renders the requested slide into a consumer-owned element. */
  renderThumbnail: (slideIndex: number, element: HTMLElement) => Promise<void>;
  /** Re-renders every thumbnail currently attached through `containerRef`. */
  rerenderAttachedThumbnails: () => Promise<void>;
  /** Thumbnail metadata and render helpers for each slide. */
  thumbnails: PptxSlideThumbnailItem[];
}

export interface PptxViewerController {
  goToSlide(index: number, scrollOptions?: ScrollIntoViewOptions): Promise<void>;
  next(): Promise<void>;
  previous(): Promise<void>;
  setZoom(level: ViewerZoomLevel): Promise<void>;
  /** @deprecated Use `setZoom()` with a responsive zoom mode. */
  setFitMode(mode: FitMode): Promise<void>;
  search(query: string | RegExp, options?: ViewerSearchOptions): PresentationSearchResult[];
  highlightSearchResult(
    result: PresentationSearchResult,
    options?: SearchHighlightOptions,
  ): Promise<void>;
  clearSearchHighlights(): void;
  /** True after the slide renderer is ready to service thumbnail requests. */
  isReady(): boolean;
  /** Renders a detached slide thumbnail and returns its cleanup function. */
  renderThumbnail(
    index: number,
    target: HTMLElement,
    options?: PptxSlideThumbnailRenderOptions,
  ): Promise<() => void>;
  getDocument(): PresentationDocument | null;
  getSlideIndex(): number;
  /** Current numeric percentage or active responsive mode. */
  getZoom(): ViewerZoomLevel;
  /** Current clamped percentage after resolving the viewport-dependent mode. */
  getResolvedZoom(): number;
}

export interface ReactPptxViewerProps extends Omit<
  HTMLAttributes<HTMLDivElement>,
  'children' | 'onError' | 'onLoad'
> {
  source: PresentationSource;
  mode?: ViewerMode;
  /** Controlled zero-based slide index. */
  slideIndex?: number;
  /** Initial zero-based slide index for uncontrolled usage. */
  initialSlide?: number;
  /** Controlled zoom percentage or responsive zoom mode. */
  zoom?: ViewerZoomLevel;
  /** Initial zoom percentage or responsive zoom mode for uncontrolled usage. */
  defaultZoom?: ViewerZoomLevel;
  /** @deprecated Superseded by `zoom`. `contain` maps to `automatic`. */
  fitMode?: FitMode;
  width?: number;
  height?: number | string;
  showToolbar?: boolean;
  showThumbnails?: boolean;
  showNotes?: boolean;
  showDiagnostics?: boolean;
  showSlideLabels?: boolean;
  virtualization?: boolean | VirtualizationOptions;
  /** Font loading, substitution, and missing-font diagnostics. */
  fonts?: PptxFontOptions;
  parseOptions?: ParsePresentationOptions;
  searchQuery?: string | RegExp;
  searchOptions?: ViewerSearchOptions;
  activeSearchResult?: number | PresentationSearchResult | null;
  searchHighlightOptions?: SearchHighlightOptions;
  renderThumbnail?: (context: ThumbnailRenderContext) => ReactNode;
  renderLoading?: () => ReactNode;
  renderError?: (error: import('./errors').PptxViewerError) => ReactNode;
  emptyState?: ReactNode;
  toolbarClassName?: string;
  viewportClassName?: string;
  toolbarStyle?: CSSProperties;
  viewportStyle?: CSSProperties;
  onReady?: (controller: PptxViewerController) => void;
  onLoad?: (presentation: ParsedPresentation) => void;
  onError?: (error: import('./errors').PptxViewerError) => void;
  onWarning?: (warning: PresentationWarning) => void;
  onSlideChange?: (index: number) => void;
  /** Reports semantic zoom changes and responsive percentage recalculation. */
  onZoomChange?: (state: ViewerZoomState) => void;
  onSlideRendered?: (index: number, element: HTMLElement) => void;
  onSlideUnmounted?: (index: number) => void;
  onSearchResults?: (results: readonly PresentationSearchResult[]) => void;
  onThumbnailRendered?: (index: number, element: HTMLElement) => void;
  /** Receives parsing, font preparation, and rendering timings. */
  onPerformanceMeasurement?: (measurement: PptxPerformanceMeasurement) => void;
  /** Supplies the internal slide surface to host integrations and test harnesses. */
  onViewportReady?: (element: HTMLDivElement) => void;
  ref?: Ref<PptxViewerController>;
}

export type {
  PresentationAsset,
  PresentationDocument,
  PresentationEmbeddedFont,
  PresentationFormat,
  PresentationMetadata,
  PresentationSearchResult,
  PresentationSlide,
  PresentationTheme,
  PresentationWarning,
  Slide,
  SlideComment,
  SlideLayout,
  SlideMaster,
  SlideNode,
  SlideNote,
} from '@extend-ai/react-pptx-model';
