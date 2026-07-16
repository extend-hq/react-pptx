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

export type PresentationSource =
  BinaryPresentationSource | PresentationDocument | ParsedPresentation;

export type ViewerMode = 'slide' | 'continuous';
export type FitMode = 'contain' | 'none';

export interface ParsePresentationOptions {
  signal?: AbortSignal;
  formatHint?: PresentationFormat;
  maxInputBytes?: number;
  fetchInit?: Omit<RequestInit, 'signal'>;
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

export interface PptxViewerController {
  goToSlide(index: number, scrollOptions?: ScrollIntoViewOptions): Promise<void>;
  next(): Promise<void>;
  previous(): Promise<void>;
  setZoom(percent: number): Promise<void>;
  setFitMode(mode: FitMode): Promise<void>;
  search(query: string | RegExp, options?: ViewerSearchOptions): PresentationSearchResult[];
  highlightSearchResult(
    result: PresentationSearchResult,
    options?: SearchHighlightOptions,
  ): Promise<void>;
  clearSearchHighlights(): void;
  getDocument(): PresentationDocument | null;
  getSlideIndex(): number;
  getZoom(): number;
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
  /** Controlled zoom percentage. */
  zoom?: number;
  defaultZoom?: number;
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
  onSlideRendered?: (index: number, element: HTMLElement) => void;
  onSlideUnmounted?: (index: number) => void;
  onSearchResults?: (results: readonly PresentationSearchResult[]) => void;
  onThumbnailRendered?: (index: number, element: HTMLElement) => void;
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
