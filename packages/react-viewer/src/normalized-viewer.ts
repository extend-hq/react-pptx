import type {
  FillStyle,
  LineStyle,
  PresentationAsset,
  PresentationDocument,
  PresentationSearchResult,
  SlideNode,
  TextParagraph,
} from '@extend-ai/react-pptx-model';
import type {
  FitMode,
  SearchHighlightOptions,
  ViewerSearchOptions,
  VirtualizationOptions,
} from './types';
import { convertEmfToDataUrl, convertWmfToDataUrl } from 'emf-converter';
import { OFFICE_FONT_FALLBACKS } from './fonts';

const EMU_PER_CSS_PIXEL = 9_525;

interface NormalizedViewerCallbacks {
  onSlideChange?: (index: number) => void;
  onSlideRendered?: (index: number, element: HTMLElement) => void;
  onSlideUnmounted?: (index: number) => void;
  onNodeError?: (nodeId: string, error: unknown) => void;
}

interface DisposableHandle {
  element: HTMLElement;
  target: HTMLElement;
  ready: Promise<void>;
  dispose(): void;
}

type RenderListOptions = VirtualizationOptions & {
  showSlideLabels?: boolean;
  initialSlideIndex?: number;
};

interface ListPlaceholder {
  item: HTMLElement;
  mount(): Promise<void>;
  unmount(): void;
  isMounted(): boolean;
}

interface ActiveListState {
  generation: number;
  options: RenderListOptions;
  placeholders: ListPlaceholder[];
}

const METAFILE_FONT_MAP = Object.fromEntries(
  Object.entries(OFFICE_FONT_FALLBACKS).map(([family, fallbacks]) => [
    family,
    fallbacks.find(
      (candidate) => candidate.toLowerCase() !== family && !candidate.endsWith('serif'),
    ) ??
      fallbacks[0] ??
      family,
  ]),
);

function safeCssColorToken(raw: string): string | undefined {
  const value = raw.trim();
  if (!value || /[\u0000-\u001f\u007f;{}'"\\]/.test(value)) return undefined;
  if (/^(?:url|var|image|element|cross-fade)\s*\(/i.test(value)) return undefined;
  if (/^#(?:[\da-f]{3,4}|[\da-f]{6}|[\da-f]{8})$/i.test(value)) return value;
  if (/^(?:[\da-f]{3,4}|[\da-f]{6}|[\da-f]{8})$/i.test(value)) return `#${value}`;
  if (/^(?:rgb|rgba|hsl|hsla)\([\d\s.,%+/-]+\)$/i.test(value)) return value;
  if (/^[a-z]+$/i.test(value)) {
    if (/^(?:inherit|initial|unset|revert|currentcolor)$/i.test(value)) return undefined;
    const probe = document.createElement('span');
    probe.style.color = value;
    return probe.style.color ? value : undefined;
  }
  return undefined;
}

function color(value?: { value: string; alpha?: number }): string | undefined {
  if (!value) return undefined;
  const safe = safeCssColorToken(value.value);
  if (!safe) return undefined;
  if (value.alpha === undefined || value.alpha >= 1) return safe;
  const alpha = Math.max(0, Math.min(1, value.alpha));
  return `color-mix(in srgb, ${safe} ${Math.round(alpha * 100)}%, transparent)`;
}

function fill(fillStyle?: FillStyle): string | undefined {
  if (!fillStyle || fillStyle.type === 'none') return undefined;
  if (fillStyle.type === 'solid') return color(fillStyle.color);
  if (fillStyle.type === 'gradient') {
    const stops = fillStyle.stops.map((stop) => {
      const stopColor = color(stop.color);
      return stopColor ? `${stopColor} ${stop.position * 100}%` : undefined;
    });
    if (stops.some((stop) => !stop)) return undefined;
    return `linear-gradient(${fillStyle.angle ?? 0}deg, ${stops.join(', ')})`;
  }
  if (fillStyle.type === 'pattern') {
    const foreground = color(fillStyle.foreground) ?? 'currentColor';
    const background = color(fillStyle.background) ?? 'transparent';
    const angle = /(?:vert|vertical)/i.test(fillStyle.preset)
      ? 90
      : /(?:diag|cross)/i.test(fillStyle.preset)
        ? 45
        : 0;
    return `repeating-linear-gradient(${angle}deg, ${foreground} 0 1px, ${background} 1px 5px)`;
  }
  return undefined;
}

function safeHyperlink(raw: string | undefined): string | undefined {
  if (!raw) return undefined;
  const value = raw.trim();
  if (!value || /[\u0000-\u001f\u007f\\]/.test(value)) return undefined;
  if (value.startsWith('//')) return undefined;
  const scheme = /^([a-z][a-z\d+.-]*):/i.exec(value)?.[1]?.toLowerCase();
  if (scheme && !['http', 'https', 'mailto', 'tel'].includes(scheme)) return undefined;
  return value;
}

function applyText(container: HTMLElement, paragraphs: TextParagraph[]): void {
  container.style.whiteSpace = 'pre-wrap';
  container.style.overflowWrap = 'break-word';
  for (const paragraph of paragraphs) {
    const line = document.createElement('div');
    line.style.textAlign = paragraph.alignment ?? 'left';
    line.style.direction = paragraph.rtl ? 'rtl' : 'ltr';
    line.style.marginLeft = paragraph.level ? `${paragraph.level * 1.25}em` : '0';
    if (paragraph.spaceBefore !== undefined) line.style.marginTop = `${paragraph.spaceBefore}pt`;
    if (paragraph.spaceAfter !== undefined) line.style.marginBottom = `${paragraph.spaceAfter}pt`;
    if (paragraph.lineSpacing !== undefined) {
      line.style.lineHeight =
        paragraph.lineSpacing > 10 ? `${paragraph.lineSpacing}%` : String(paragraph.lineSpacing);
    }
    if (paragraph.bullet) {
      const marker = document.createElement('span');
      marker.textContent = `${paragraph.bullet.value ?? '•'} `;
      marker.setAttribute('aria-hidden', 'true');
      line.append(marker);
    }
    for (const run of paragraph.runs) {
      const href = safeHyperlink(run.hyperlink);
      const span = document.createElement(href ? 'a' : 'span');
      span.textContent = run.text;
      if (run.fontFamily) span.style.fontFamily = run.fontFamily;
      if (run.fontSizePt) span.style.fontSize = `${run.fontSizePt}pt`;
      if (run.bold) span.style.fontWeight = '700';
      if (run.italic) span.style.fontStyle = 'italic';
      const decorations = [
        run.underline ? 'underline' : '',
        run.strike ? 'line-through' : '',
      ].filter(Boolean);
      if (decorations.length) span.style.textDecoration = decorations.join(' ');
      if (run.color) span.style.color = color(run.color) ?? '';
      if (run.baseline) {
        span.style.verticalAlign = run.baseline > 0 ? 'super' : 'sub';
        span.style.fontSize = span.style.fontSize || '0.75em';
      }
      if (run.language) span.lang = run.language;
      if (href && span instanceof HTMLAnchorElement) {
        span.href = href;
        span.target = '_blank';
        span.rel = 'noreferrer noopener';
      }
      line.append(span);
    }
    container.append(line);
  }
}

function applyLine(element: HTMLElement, line?: LineStyle): void {
  if (!line) return;
  const width = Math.max(0, line.width ?? 1);
  const lineColor = color(line.color) ?? 'currentColor';
  const style = /dash|dot/i.test(line.dash ?? '') ? 'dashed' : 'solid';
  element.style.border = `${width}px ${style} ${lineColor}`;
}

function applyGeometry(element: HTMLElement, preset?: string): void {
  switch (preset) {
    case 'ellipse':
      element.style.borderRadius = '50%';
      break;
    case 'roundRect':
      element.style.borderRadius = '10%';
      break;
    case 'triangle':
      element.style.clipPath = 'polygon(50% 0, 100% 100%, 0 100%)';
      break;
    case 'rtTriangle':
      element.style.clipPath = 'polygon(0 0, 100% 100%, 0 100%)';
      break;
    case 'diamond':
      element.style.clipPath = 'polygon(50% 0, 100% 50%, 50% 100%, 0 50%)';
      break;
    case 'parallelogram':
      element.style.clipPath = 'polygon(20% 0, 100% 0, 80% 100%, 0 100%)';
      break;
    case 'hexagon':
      element.style.clipPath = 'polygon(25% 0, 75% 0, 100% 50%, 75% 100%, 25% 100%, 0 50%)';
      break;
    case 'chevron':
      element.style.clipPath = 'polygon(0 0, 70% 0, 100% 50%, 70% 100%, 0 100%, 30% 50%)';
      break;
    case 'pentagon':
      element.style.clipPath = 'polygon(50% 0, 100% 38%, 82% 100%, 18% 100%, 0 38%)';
      break;
    case 'star5':
      element.style.clipPath =
        'polygon(50% 0, 61% 35%, 98% 35%, 68% 57%, 79% 94%, 50% 72%, 21% 94%, 32% 57%, 2% 35%, 39% 35%)';
      break;
  }
}

function isWordCharacter(value: string | undefined): boolean {
  return Boolean(value && /[\p{L}\p{N}_]/u.test(value));
}

function searchDocument(
  presentation: PresentationDocument,
  query: string | RegExp,
  options: ViewerSearchOptions = {},
): PresentationSearchResult[] {
  if (typeof query === 'string' && !query) return [];
  const results: PresentationSearchResult[] = [];
  const expression =
    query instanceof RegExp
      ? new RegExp(query.source, query.flags.includes('g') ? query.flags : `${query.flags}g`)
      : new RegExp(
          options.useRegex ? query : query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'),
          options.matchCase ? 'g' : 'gi',
        );
  const visit = (slideIndex: number, nodes: SlideNode[]): void => {
    for (const node of nodes) {
      const texts: string[] = [];
      if (node.type === 'shape' && options.includeShapes !== false) {
        texts.push(...(node.paragraphs ?? []).map((p) => p.runs.map((run) => run.text).join('')));
      } else if (node.type === 'table' && options.includeTables !== false) {
        for (const row of node.rows)
          for (const cell of row) {
            texts.push(...cell.paragraphs.map((p) => p.runs.map((run) => run.text).join('')));
          }
      } else if (node.type === 'chart') {
        if (node.title) texts.push(node.title);
        for (const series of node.series) if (series.name) texts.push(series.name);
      } else if (node.type === 'group' && options.includeGroups !== false) {
        visit(slideIndex, node.children);
      }
      const text = texts.join('\n');
      expression.lastIndex = 0;
      let match: RegExpExecArray | null;
      while ((match = expression.exec(text))) {
        const matchedText = match[0];
        if (
          !options.wholeWord ||
          (!isWordCharacter(text[match.index - 1]) &&
            !isWordCharacter(text[match.index + matchedText.length]))
        ) {
          const radius = options.snippetRadius ?? 32;
          results.push({
            slideIndex,
            nodeId: node.id,
            nodeType: node.type,
            text,
            matchStart: match.index,
            matchEnd: match.index + matchedText.length,
            snippet: text.slice(
              Math.max(0, match.index - radius),
              match.index + matchedText.length + radius,
            ),
            bounds: node.transform,
          });
        }
        if (!matchedText) expression.lastIndex += 1;
      }
    }
  };
  for (const slide of presentation.slides) visit(slide.index, slide.nodes);
  return results;
}

export class NormalizedPresentationViewer {
  private current = 0;
  private zoom = 100;
  private fit: FitMode = 'contain';
  private objectUrls = new Set<string>();
  private assetUrls = new Map<string, string>();
  private metafileUrls = new Map<string, Promise<string | undefined>>();
  private visibleSlides = new Map<number, { ratio: number; top: number }>();
  private observer: IntersectionObserver | undefined;
  private visibilityObserver: IntersectionObserver | undefined;
  private highlight: HTMLElement | undefined;
  private pendingResources = new Set<Promise<void>>();
  private mountedHandles = new Set<DisposableHandle>();
  private handlesByTarget = new Map<HTMLElement, DisposableHandle>();
  private destroyed = false;
  private renderGeneration = 0;
  private navigationGeneration = 0;
  private lastNotifiedSlide: number | undefined;
  private listState: ActiveListState | undefined;
  private renderState:
    { mode: 'single' } | { mode: 'continuous'; options: RenderListOptions } | undefined;

  constructor(
    private readonly container: HTMLElement,
    readonly presentation: PresentationDocument,
    private readonly callbacks: NormalizedViewerCallbacks = {},
  ) {}

  get slideCount(): number {
    return this.presentation.slides.length;
  }
  get currentSlideIndex(): number {
    return this.current;
  }
  get zoomPercent(): number {
    return this.zoom;
  }
  get fitMode(): FitMode {
    return this.fit;
  }

  private get naturalSlideWidth(): number {
    const width = this.presentation.size.widthEmu / EMU_PER_CSS_PIXEL;
    return Number.isFinite(width) && width > 0 ? width : 1;
  }

  private get naturalSlideHeight(): number {
    const height = this.presentation.size.heightEmu / EMU_PER_CSS_PIXEL;
    return Number.isFinite(height) && height > 0 ? height : 1;
  }

  private isRenderActive(generation: number): boolean {
    return !this.destroyed && generation === this.renderGeneration;
  }

  private beginRender(): number | undefined {
    if (this.destroyed) return undefined;
    const generation = ++this.renderGeneration;
    this.navigationGeneration += 1;
    this.observer?.disconnect();
    this.visibilityObserver?.disconnect();
    this.observer = undefined;
    this.visibilityObserver = undefined;
    this.visibleSlides.clear();
    this.listState = undefined;
    this.clearSearchHighlights();
    return generation;
  }

  private notifySlideChange(index: number): void {
    if (this.destroyed || this.lastNotifiedSlide === index) return;
    this.lastNotifiedSlide = index;
    this.callbacks.onSlideChange?.(index);
  }

  private metafileUrl(
    assetId: string,
    contentType: string,
    data: Uint8Array,
  ): Promise<string | undefined> {
    const cached = this.metafileUrls.get(assetId);
    if (cached) return cached;
    const convert = contentType.includes('wmf') ? convertWmfToDataUrl : convertEmfToDataUrl;
    const ownedBytes = data.slice();
    const pending = convert(ownedBytes.buffer as ArrayBuffer, 2_048, 2_048, {
      dpiScale: 2,
      maxCanvasDimension: 4_096,
      maxRecords: 200_000,
      fontFamilyMap: METAFILE_FONT_MAP,
    }).then((url) => url || undefined);
    this.metafileUrls.set(assetId, pending);
    return pending;
  }

  private assetUrl(asset: PresentationAsset): Promise<string | undefined> {
    if (asset.url) return Promise.resolve(asset.url);
    if (!asset.data) return Promise.resolve(undefined);
    const contentType = asset.contentType.toLowerCase();
    if (contentType.includes('emf') || contentType.includes('wmf')) {
      return this.metafileUrl(asset.id, contentType, asset.data);
    }
    const cached = this.assetUrls.get(asset.id);
    if (cached) return Promise.resolve(cached);
    const ownedBytes = asset.data.slice();
    const url = URL.createObjectURL(
      new Blob([ownedBytes.buffer as ArrayBuffer], { type: asset.contentType }),
    );
    this.objectUrls.add(url);
    this.assetUrls.set(asset.id, url);
    return Promise.resolve(url);
  }

  private applyAsset(assetId: string, apply: (url: string) => void, nodeId: string): void {
    const asset = this.presentation.assets[assetId];
    if (!asset) return;
    const pending = this.assetUrl(asset)
      .then((url) => {
        if (url && !this.destroyed) apply(url);
      })
      .catch((error: unknown) => {
        if (!this.destroyed) this.callbacks.onNodeError?.(nodeId, error);
      })
      .finally(() => this.pendingResources.delete(pending));
    this.pendingResources.add(pending);
  }

  private applyFill(element: HTMLElement, fillStyle: FillStyle | undefined, nodeId: string): void {
    if (fillStyle?.type === 'image') {
      element.style.backgroundRepeat = fillStyle.mode === 'tile' ? 'repeat' : 'no-repeat';
      element.style.backgroundSize = fillStyle.mode === 'tile' ? 'auto' : '100% 100%';
      this.applyAsset(
        fillStyle.assetId,
        (url) => {
          element.style.backgroundImage = `url(${JSON.stringify(url)})`;
        },
        nodeId,
      );
      return;
    }
    element.style.background = fill(fillStyle) ?? 'transparent';
  }

  private createChart(node: Extract<SlideNode, { type: 'chart' }>): HTMLElement {
    const chart = document.createElement('div');
    chart.style.width = '100%';
    chart.style.height = '100%';
    chart.style.display = 'flex';
    chart.style.flexDirection = 'column';
    chart.style.boxSizing = 'border-box';
    chart.style.padding = '4%';
    chart.setAttribute('role', 'img');
    chart.setAttribute('aria-label', node.title ?? `${node.chartType} chart`);
    if (node.title) {
      const title = document.createElement('div');
      title.textContent = node.title;
      title.style.textAlign = 'center';
      title.style.fontWeight = '600';
      title.style.flex = '0 0 auto';
      chart.append(title);
    }
    const values = node.series
      .flatMap((series) => series.values)
      .filter((value): value is number => value !== null);
    if (!values.length) {
      const empty = document.createElement('div');
      empty.textContent = node.title ?? 'Chart';
      empty.style.margin = 'auto';
      chart.append(empty);
      return chart;
    }
    const palette = ['#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47'];
    if (/pie|doughnut/i.test(node.chartType)) {
      const total = values.reduce((sum, value) => sum + Math.max(0, value), 0) || 1;
      let cursor = 0;
      const stops = values.map((value, index) => {
        const start = (cursor / total) * 360;
        cursor += Math.max(0, value);
        const end = (cursor / total) * 360;
        return `${palette[index % palette.length]} ${start}deg ${end}deg`;
      });
      const pie = document.createElement('div');
      pie.style.margin = 'auto';
      pie.style.width = 'min(78%, 220px)';
      pie.style.aspectRatio = '1';
      pie.style.borderRadius = '50%';
      pie.style.background = `conic-gradient(${stops.join(', ')})`;
      if (/doughnut/i.test(node.chartType)) {
        pie.style.mask = 'radial-gradient(circle, transparent 0 42%, #000 43%)';
      }
      chart.append(pie);
      return chart;
    }
    const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    svg.setAttribute('viewBox', '0 0 1000 600');
    svg.setAttribute('preserveAspectRatio', 'none');
    svg.style.width = '100%';
    svg.style.flex = '1 1 auto';
    const maximum = Math.max(1, ...values.map((value) => Math.abs(value)));
    const pointCount = Math.max(1, ...node.series.map((series) => series.values.length));
    if (/line|scatter|area/i.test(node.chartType)) {
      node.series.forEach((series, seriesIndex) => {
        const points = series.values
          .map((value, index) =>
            value === null
              ? null
              : `${80 + (index / Math.max(1, pointCount - 1)) * 840},${540 - (value / maximum) * 460}`,
          )
          .filter((point): point is string => point !== null)
          .join(' ');
        const line = document.createElementNS('http://www.w3.org/2000/svg', 'polyline');
        line.setAttribute('points', points);
        line.setAttribute('fill', 'none');
        line.setAttribute('stroke', color(series.color) ?? palette[seriesIndex % palette.length]!);
        line.setAttribute('stroke-width', '8');
        svg.append(line);
      });
    } else {
      const seriesCount = Math.max(1, node.series.length);
      const slot = 840 / pointCount;
      const barWidth = Math.max(4, (slot * 0.72) / seriesCount);
      node.series.forEach((series, seriesIndex) => {
        series.values.forEach((value, index) => {
          if (value === null) return;
          const height = (Math.abs(value) / maximum) * 460;
          const rect = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
          rect.setAttribute('x', String(80 + index * slot + seriesIndex * barWidth));
          rect.setAttribute('y', String(540 - height));
          rect.setAttribute('width', String(barWidth * 0.9));
          rect.setAttribute('height', String(height));
          rect.setAttribute('fill', color(series.color) ?? palette[seriesIndex % palette.length]!);
          svg.append(rect);
        });
      });
    }
    chart.append(svg);
    return chart;
  }

  private createNode(
    node: SlideNode,
    coordinate: { x: number; y: number; width: number; height: number },
  ): HTMLElement {
    const href = safeHyperlink(node.hyperlink);
    const element = document.createElement(href ? 'a' : 'div');
    element.dataset.rpvNodeId = node.id;
    element.style.position = 'absolute';
    element.style.left = `${((node.transform.x - coordinate.x) / coordinate.width) * 100}%`;
    element.style.top = `${((node.transform.y - coordinate.y) / coordinate.height) * 100}%`;
    element.style.width = `${(node.transform.width / coordinate.width) * 100}%`;
    element.style.height = `${(node.transform.height / coordinate.height) * 100}%`;
    element.style.boxSizing = 'border-box';
    const transforms = [`rotate(${node.transform.rotation ?? 0}deg)`];
    if (node.transform.flipHorizontal) transforms.push('scaleX(-1)');
    if (node.transform.flipVertical) transforms.push('scaleY(-1)');
    element.style.transform = transforms.join(' ');
    element.style.transformOrigin = 'center';
    if (node.hidden) element.hidden = true;
    if (node.opacity !== undefined) element.style.opacity = String(node.opacity);
    if (node.altText) element.setAttribute('aria-label', node.altText);
    if (href && element instanceof HTMLAnchorElement) {
      element.href = href;
      element.target = '_blank';
      element.rel = 'noreferrer noopener';
    }

    if (node.type === 'shape') {
      this.applyFill(element, node.fill, node.id);
      applyGeometry(element, node.geometry.preset);
      applyLine(element, node.line);
      element.style.overflow = 'hidden';
      element.style.display = 'flex';
      element.style.flexDirection = 'column';
      element.style.whiteSpace = 'pre-wrap';
      element.style.justifyContent =
        node.verticalAlignment === 'middle'
          ? 'center'
          : node.verticalAlignment === 'bottom'
            ? 'flex-end'
            : 'flex-start';
      if (node.textInsets) {
        const { top, right, bottom, left } = node.textInsets;
        element.style.padding = `${top / EMU_PER_CSS_PIXEL}px ${right / EMU_PER_CSS_PIXEL}px ${bottom / EMU_PER_CSS_PIXEL}px ${left / EMU_PER_CSS_PIXEL}px`;
      }
      if (node.paragraphs) applyText(element, node.paragraphs);
    } else if (node.type === 'image') {
      const image = document.createElement('img');
      this.applyAsset(node.assetId, (url) => (image.src = url), node.id);
      image.alt = node.altText ?? node.name ?? '';
      image.style.width = '100%';
      image.style.height = '100%';
      image.style.objectFit = node.preserveAspectRatio === false ? 'fill' : 'contain';
      if (node.crop) {
        image.style.clipPath = `inset(${node.crop.top * 100}% ${node.crop.right * 100}% ${node.crop.bottom * 100}% ${node.crop.left * 100}%)`;
      }
      element.append(image);
    } else if (node.type === 'group') {
      const childCoordinate = node.childTransform ?? {
        x: 0,
        y: 0,
        width: node.transform.width,
        height: node.transform.height,
      };
      for (const child of node.children) element.append(this.createNode(child, childCoordinate));
    } else if (node.type === 'table') {
      const table = document.createElement('table');
      table.style.width = '100%';
      table.style.height = '100%';
      table.style.borderCollapse = 'collapse';
      table.style.tableLayout = 'fixed';
      if (node.columnWidths?.length) {
        const total = node.columnWidths.reduce((sum, width) => sum + width, 0) || 1;
        const columns = document.createElement('colgroup');
        for (const width of node.columnWidths) {
          const column = document.createElement('col');
          column.style.width = `${(width / total) * 100}%`;
          columns.append(column);
        }
        table.append(columns);
      }
      const totalRowHeight = node.rowHeights?.reduce((sum, height) => sum + height, 0) || 0;
      node.rows.forEach((row, rowIndex) => {
        const tr = table.insertRow();
        if (totalRowHeight && node.rowHeights?.[rowIndex] !== undefined) {
          tr.style.height = `${(node.rowHeights[rowIndex]! / totalRowHeight) * 100}%`;
        }
        for (const cell of row) {
          const td = tr.insertCell();
          td.colSpan = cell.colSpan ?? 1;
          td.rowSpan = cell.rowSpan ?? 1;
          td.style.padding = '0.25em';
          td.style.overflow = 'hidden';
          td.style.verticalAlign = 'middle';
          td.style.border = '1px solid currentColor';
          this.applyFill(td, cell.fill, node.id);
          for (const [side, border] of Object.entries(cell.borders ?? {})) {
            if (!border) continue;
            const width = Math.max(0, border.width ?? 1);
            const style = /dash|dot/i.test(border.dash ?? '') ? 'dashed' : 'solid';
            td.style.setProperty(
              `border-${side}`,
              `${width}px ${style} ${color(border.color) ?? 'currentColor'}`,
            );
          }
          applyText(td, cell.paragraphs);
        }
      });
      element.append(table);
    } else if (node.type === 'chart') {
      element.style.border = '1px solid color-mix(in srgb, currentColor 25%, transparent)';
      element.append(this.createChart(node));
    } else if (node.type === 'media') {
      const media = document.createElement(node.mediaType === 'audio' ? 'audio' : 'video');
      media.controls = true;
      media.style.width = '100%';
      media.style.height = '100%';
      if (node.assetId) this.applyAsset(node.assetId, (url) => (media.src = url), node.id);
      if (node.posterAssetId && media instanceof HTMLVideoElement) {
        this.applyAsset(node.posterAssetId, (url) => (media.poster = url), node.id);
      }
      element.append(media);
    } else if (node.type === 'unknown') {
      element.dataset.rpvUnsupportedFeature = node.feature;
      if (node.fallbackAssetId) {
        const image = document.createElement('img');
        image.alt = node.altText ?? node.name ?? '';
        image.style.width = '100%';
        image.style.height = '100%';
        image.style.objectFit = 'contain';
        this.applyAsset(node.fallbackAssetId, (url) => (image.src = url), node.id);
        element.append(image);
      }
    }
    return element;
  }

  private createSlide(index: number): HTMLElement {
    const slide = this.presentation.slides[index];
    if (!slide) throw new RangeError(`Slide ${index} does not exist.`);
    const element = document.createElement('section');
    element.className = 'rpv-normalized-slide';
    element.dataset.rpvSlideIndex = String(index);
    element.setAttribute('aria-label', `Slide ${index + 1} of ${this.slideCount}`);
    element.style.position = 'relative';
    element.style.width = `${this.naturalSlideWidth}px`;
    element.style.height = `${this.naturalSlideHeight}px`;
    element.style.overflow = 'hidden';
    element.style.background = '#fff';
    if (slide.background) this.applyFill(element, slide.background, slide.id);
    for (const node of slide.nodes) {
      try {
        element.append(
          this.createNode(node, {
            x: 0,
            y: 0,
            width: this.presentation.size.widthEmu,
            height: this.presentation.size.heightEmu,
          }),
        );
      } catch (error) {
        this.callbacks.onNodeError?.(node.id, error);
      }
    }
    return element;
  }

  private emptyHandle(target: HTMLElement): DisposableHandle {
    const element = document.createElement('section');
    return {
      element,
      target,
      ready: Promise.resolve(),
      dispose: () => element.remove(),
    };
  }

  private mount(index: number, target: HTMLElement, scale?: number): DisposableHandle {
    if (this.destroyed) return this.emptyHandle(target);
    this.handlesByTarget.get(target)?.dispose();
    const slide = this.createSlide(index);
    const fitScale =
      this.fit === 'contain'
        ? Math.min(1, (target.clientWidth || this.naturalSlideWidth) / this.naturalSlideWidth)
        : 1;
    const effectiveScale = scale ?? fitScale * (this.zoom / 100);
    const wrapper = document.createElement('div');
    wrapper.dataset.rpvSlideWrapper = String(index);
    wrapper.style.position = 'relative';
    wrapper.style.width = `${this.naturalSlideWidth * effectiveScale}px`;
    wrapper.style.height = `${this.naturalSlideHeight * effectiveScale}px`;
    wrapper.style.margin = '0 auto';
    slide.style.transformOrigin = 'top left';
    slide.style.transform = `scale(${effectiveScale})`;
    wrapper.append(slide);
    target.replaceChildren(wrapper);
    this.callbacks.onSlideRendered?.(index, slide);
    let disposed = false;
    const handle: DisposableHandle = {
      element: slide,
      target,
      ready: Promise.all([...this.pendingResources]).then(() => undefined),
      dispose: () => {
        if (disposed) return;
        disposed = true;
        wrapper.remove();
        this.mountedHandles.delete(handle);
        if (this.handlesByTarget.get(target) === handle) this.handlesByTarget.delete(target);
        this.callbacks.onSlideUnmounted?.(index);
      },
    };
    this.mountedHandles.add(handle);
    this.handlesByTarget.set(target, handle);
    return handle;
  }

  private disposeContainerSlides(): void {
    for (const handle of [...this.mountedHandles]) {
      if (handle.target === this.container || this.container.contains(handle.target)) {
        handle.dispose();
      }
    }
  }

  private async yieldRenderBatch(): Promise<void> {
    if (typeof requestAnimationFrame !== 'undefined') {
      await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
    } else {
      await Promise.resolve();
    }
  }

  private async mountInBatches(
    placeholders: ListPlaceholder[],
    indices: number[],
    batchSize: number,
    generation: number,
  ): Promise<void> {
    const uniqueIndices = [...new Set(indices)].filter(
      (index) => index >= 0 && index < placeholders.length,
    );
    const size =
      Number.isFinite(batchSize) && batchSize > 0 ? Math.max(1, Math.floor(batchSize)) : 3;
    for (let offset = 0; offset < uniqueIndices.length; offset += size) {
      if (!this.isRenderActive(generation)) return;
      await Promise.all(
        uniqueIndices.slice(offset, offset + size).map((index) => placeholders[index]!.mount()),
      );
      if (!this.isRenderActive(generation)) return;
      if (offset + size < uniqueIndices.length) await this.yieldRenderBatch();
    }
  }

  async renderSlide(index = this.current): Promise<void> {
    this.renderState = { mode: 'single' };
    const generation = this.beginRender();
    if (generation === undefined) return;
    this.current = Math.max(0, Math.min(this.slideCount - 1, index));
    this.disposeContainerSlides();
    this.container.replaceChildren();
    if (this.slideCount === 0) return;
    const handle = this.mount(this.current, this.container);
    await handle.ready;
    if (!this.isRenderActive(generation) || this.handlesByTarget.get(this.container) !== handle) {
      return;
    }
    this.notifySlideChange(this.current);
  }

  async renderList(options: RenderListOptions = {}): Promise<void> {
    const windowingEnabled = options.enabled !== false;
    const normalizedOptions = { ...options, enabled: windowingEnabled };
    this.renderState = { mode: 'continuous', options: normalizedOptions };
    const generation = this.beginRender();
    if (generation === undefined) return;
    this.disposeContainerSlides();
    this.container.replaceChildren();
    if (!options.scrollElement) {
      this.container.scrollTop = 0;
      this.container.scrollLeft = 0;
    }
    if (this.slideCount === 0) {
      this.current = 0;
      this.listState = { generation, options: normalizedOptions, placeholders: [] };
      return;
    }

    const initialIndex = Math.max(
      0,
      Math.min(this.slideCount - 1, options.initialSlideIndex ?? this.current),
    );
    const requestedInitial = options.initialSlides ?? 2;
    const initial =
      Number.isFinite(requestedInitial) && requestedInitial >= 0 ? Math.floor(requestedInitial) : 2;
    const requestedOverscan = options.overscanViewport ?? 1.5;
    const overscan =
      Number.isFinite(requestedOverscan) && requestedOverscan >= 0 ? requestedOverscan : 1.5;
    const viewportWidth = (options.scrollElement ?? this.container).clientWidth;
    const estimatedFitScale =
      this.fit === 'contain'
        ? Math.min(1, (viewportWidth || this.naturalSlideWidth) / this.naturalSlideWidth)
        : 1;
    const placeholders = this.presentation.slides.map((slide, index) => {
      const item = document.createElement('div');
      item.dataset.rpvListItem = String(index);
      item.style.minHeight = `${this.naturalSlideHeight * estimatedFitScale * (this.zoom / 100)}px`;
      item.style.margin = '0 auto 24px';
      let handle: DisposableHandle | undefined;
      let mountPromise: Promise<void> | undefined;
      const isActive = () =>
        this.isRenderActive(generation) && this.listState?.generation === generation;
      const mount = (): Promise<void> => {
        if (!isActive()) return Promise.resolve();
        if (handle) return handle.ready;
        if (mountPromise) return mountPromise;
        const nextHandle = this.mount(index, item);
        handle = nextHandle;
        const pending = nextHandle.ready
          .then(() => {
            if (!isActive() || handle !== nextHandle) {
              if (handle === nextHandle) handle = undefined;
              nextHandle.dispose();
            }
          })
          .finally(() => {
            if (mountPromise === pending) mountPromise = undefined;
          });
        mountPromise = pending;
        return pending;
      };
      const unmount = (): void => {
        const mounted = handle;
        handle = undefined;
        if (mounted) {
          mounted.dispose();
        }
      };
      this.container.append(item);
      return { item, mount, unmount, isMounted: () => Boolean(handle) };
    });

    const state: ActiveListState = {
      generation,
      options: normalizedOptions,
      placeholders,
    };
    this.listState = state;
    const hasObserver = typeof IntersectionObserver !== 'undefined';
    const initialIndices =
      windowingEnabled && hasObserver
        ? [...Array(Math.min(initial, this.slideCount)).keys(), initialIndex]
        : this.presentation.slides.map((_, index) => index);
    await this.mountInBatches(placeholders, initialIndices, options.batchSize ?? 3, generation);
    if (!this.isRenderActive(generation) || this.listState !== state) return;

    if (!options.scrollElement) {
      this.container.scrollTop = 0;
      this.container.scrollLeft = 0;
    }
    const initialItem = placeholders[initialIndex]?.item;
    if (initialItem && initialIndex !== 0) {
      initialItem.scrollIntoView?.({ behavior: 'instant', block: 'start' });
    }
    this.current = initialIndex;
    this.notifySlideChange(initialIndex);
    if (windowingEnabled && hasObserver) {
      const observerRoot = options.scrollElement ?? this.container;
      this.observer = new IntersectionObserver(
        (entries) => {
          if (!this.isRenderActive(generation) || this.listState !== state) return;
          for (const entry of entries) {
            const index = Number((entry.target as HTMLElement).dataset.rpvListItem);
            if (entry.isIntersecting) void placeholders[index]?.mount();
            else placeholders[index]?.unmount();
          }
        },
        {
          root: observerRoot,
          rootMargin: `${overscan * 100}% 0px`,
        },
      );
      this.visibilityObserver = new IntersectionObserver(
        (entries) => {
          if (!this.isRenderActive(generation) || this.listState !== state) return;
          for (const entry of entries) {
            const index = Number((entry.target as HTMLElement).dataset.rpvListItem);
            if (entry.isIntersecting) {
              void placeholders[index]?.mount();
              this.visibleSlides.set(index, {
                ratio: entry.intersectionRatio,
                top: entry.boundingClientRect.top,
              });
            } else this.visibleSlides.delete(index);
          }
          const visible = [...this.visibleSlides.entries()].sort(
            ([firstIndex, first], [secondIndex, second]) =>
              second.ratio - first.ratio ||
              Math.abs(first.top) - Math.abs(second.top) ||
              firstIndex - secondIndex,
          )[0];
          if (visible && visible[0] !== this.current) {
            this.current = visible[0];
            this.notifySlideChange(this.current);
          }
        },
        {
          root: observerRoot,
          threshold: [0, 0.01, 0.25, 0.5, 0.75, 1],
        },
      );
      for (const { item } of placeholders) {
        this.observer.observe(item);
        this.visibilityObserver.observe(item);
      }
    }
  }

  async goToSlide(index: number, scrollOptions?: ScrollIntoViewOptions): Promise<void> {
    if (this.destroyed || this.slideCount === 0) return;
    const next = Math.max(0, Math.min(this.slideCount - 1, index));
    const changed = next !== this.current;
    const state = this.listState;
    if (state && this.isRenderActive(state.generation)) {
      const navigation = ++this.navigationGeneration;
      const placeholder = state.placeholders[next];
      await placeholder?.mount();
      if (
        this.destroyed ||
        navigation !== this.navigationGeneration ||
        !this.isRenderActive(state.generation) ||
        this.listState !== state
      ) {
        return;
      }
      placeholder?.item.scrollIntoView?.(scrollOptions);
      this.current = next;
      if (changed) this.notifySlideChange(next);
      return;
    }
    if (!changed && this.container.querySelector(`[data-rpv-slide-index="${next}"]`)) return;
    await this.renderSlide(next);
  }

  async setZoom(percent: number): Promise<void> {
    const next = Math.max(10, Math.min(400, percent));
    if (this.zoom === next) return;
    this.zoom = next;
    if (this.renderState?.mode === 'continuous') {
      await this.renderList({ ...this.renderState.options, initialSlideIndex: this.current });
    } else if (this.renderState) await this.renderSlide(this.current);
  }
  async setFitMode(mode: FitMode): Promise<void> {
    if (this.fit === mode) return;
    this.fit = mode;
    if (this.renderState?.mode === 'continuous') {
      await this.renderList({ ...this.renderState.options, initialSlideIndex: this.current });
    } else if (this.renderState) await this.renderSlide(this.current);
  }
  searchText(query: string | RegExp, options?: ViewerSearchOptions): PresentationSearchResult[] {
    return searchDocument(this.presentation, query, options);
  }
  renderThumbnailToContainer(
    index: number,
    target: HTMLElement,
    options: { width?: number } = {},
  ): DisposableHandle {
    if (this.destroyed || index < 0 || index >= this.slideCount) return this.emptyHandle(target);
    return this.mount(index, target, (options.width ?? 144) / this.naturalSlideWidth);
  }
  async highlightSearchResult(
    result: PresentationSearchResult,
    options: SearchHighlightOptions = {},
  ): Promise<void> {
    if (this.destroyed || result.slideIndex < 0 || result.slideIndex >= this.slideCount) return;
    if (options.scrollIntoView !== false) {
      await this.goToSlide(
        result.slideIndex,
        typeof options.scrollIntoView === 'object' ? options.scrollIntoView : undefined,
      );
    } else if (this.listState) {
      const state = this.listState;
      await state.placeholders[result.slideIndex]?.mount();
      if (!this.isRenderActive(state.generation) || this.listState !== state) return;
    } else if (!this.container.querySelector(`[data-rpv-slide-index="${result.slideIndex}"]`)) {
      await this.renderSlide(result.slideIndex);
    }
    if (this.destroyed) return;
    this.clearSearchHighlights();
    const slide = this.container.querySelector<HTMLElement>(
      `[data-rpv-slide-index="${result.slideIndex}"]`,
    );
    const node = [...(slide?.querySelectorAll<HTMLElement>('[data-rpv-node-id]') ?? [])].find(
      (candidate) => candidate.dataset.rpvNodeId === result.nodeId,
    );
    const parent = node?.parentElement;
    if (!node || !parent) return;
    const highlight = document.createElement('div');
    highlight.className = ['rpv-search-highlight', options.className].filter(Boolean).join(' ');
    highlight.style.position = 'absolute';
    highlight.style.left = node.style.left;
    highlight.style.top = node.style.top;
    highlight.style.width = node.style.width;
    highlight.style.height = node.style.height;
    highlight.style.boxSizing = 'border-box';
    highlight.style.transform = node.style.transform;
    highlight.style.transformOrigin = node.style.transformOrigin;
    const padding = options.padding ?? 3;
    const width =
      typeof options.borderWidth === 'number'
        ? `${options.borderWidth}px`
        : (options.borderWidth ?? '3px');
    highlight.style.outline = `${width} solid ${options.borderColor ?? '#ef8b2c'}`;
    highlight.style.outlineOffset = `${padding}px`;
    highlight.style.background = options.backgroundColor ?? 'transparent';
    highlight.style.boxShadow =
      options.boxShadow ?? '0 0 0 2px color-mix(in srgb, #ef8b2c 30%, transparent)';
    if (options.borderRadius !== undefined) {
      highlight.style.borderRadius =
        typeof options.borderRadius === 'number'
          ? `${options.borderRadius}px`
          : options.borderRadius;
    }
    highlight.style.zIndex = String(options.zIndex ?? 2_147_483_646);
    for (const [property, value] of Object.entries(options.style ?? {})) {
      if (value === undefined) continue;
      if (property.startsWith('--') || property.includes('-')) {
        highlight.style.setProperty(property, String(value));
      } else {
        (highlight.style as unknown as Record<string, string>)[property] = String(value);
      }
    }
    highlight.style.pointerEvents = 'none';
    parent.append(highlight);
    this.highlight = highlight;
  }
  clearSearchHighlights(): void {
    this.highlight?.remove();
    this.highlight = undefined;
  }
  destroy(): void {
    if (this.destroyed) return;
    this.destroyed = true;
    this.renderGeneration += 1;
    this.navigationGeneration += 1;
    this.observer?.disconnect();
    this.visibilityObserver?.disconnect();
    this.observer = undefined;
    this.visibilityObserver = undefined;
    this.visibleSlides.clear();
    this.listState = undefined;
    this.clearSearchHighlights();
    for (const url of this.objectUrls) URL.revokeObjectURL(url);
    this.objectUrls.clear();
    this.assetUrls.clear();
    this.metafileUrls.clear();
    this.pendingResources.clear();
    for (const handle of [...this.mountedHandles]) handle.dispose();
    this.handlesByTarget.clear();
    this.renderState = undefined;
    this.container.replaceChildren();
  }
}
