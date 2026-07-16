import type {
  FillStyle,
  ImageEffects,
  LineStyle,
  PresentationAsset,
  PresentationDocument,
  PresentationSearchResult,
  PresentationTheme,
  SlideNode,
  TextParagraph,
  TextRun,
} from '@extend-ai/react-pptx-model';
import { renderChartInto } from './charts/chart-host';
import type {
  FitMode,
  SearchHighlightOptions,
  ViewerSearchOptions,
  VirtualizationOptions,
} from './types';
import { Virtualizer } from '@tanstack/virtual-core';
import { OFFICE_FONT_FALLBACKS } from './fonts';
import { renderEmfToDataUrl, renderWmfToDataUrl } from './metafile-renderer';

const EMU_PER_CSS_PIXEL = 9_525;
const DEFAULT_TEXT_HORIZONTAL_INSET_EMU = 91_440;
const DEFAULT_TEXT_VERTICAL_INSET_EMU = 45_720;
const POWERPOINT_SINGLE_LINE_HEIGHT = 1.15;

interface TextRenderOptions {
  fontScale?: number;
  lineSpacingReduction?: number;
  textWrap?: string;
  spaceFirstLastParagraph?: boolean;
}

type TextSpacingValue = number | { value: number; unit: 'points' | 'percent' };

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
  /** Deterministic offset navigation provided by the TanStack virtualizer. */
  scrollToIndex?: (index: number, options?: ScrollIntoViewOptions) => void;
}

/** Vertical spacing between slides in continuous mode, in CSS pixels. */
const LIST_ITEM_GAP = 24;

/**
 * Viewport rect used when layout measurements are unavailable (jsdom, SSR,
 * display: none) so the virtualizer still mounts an initial window.
 */
const FALLBACK_VIEWPORT_RECT = { width: 800, height: 600 };

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
    const stops = [...fillStyle.stops]
      .sort((first, second) => first.position - second.position)
      .map((stop) => {
        const stopColor = color(stop.color);
        const position = Math.max(0, Math.min(1, stop.position));
        return stopColor ? `${stopColor} ${position * 100}%` : undefined;
      });
    if (stops.some((stop) => !stop)) return undefined;
    // DrawingML's 0-degree vector points left-to-right; CSS 90deg does the
    // same. Both increase clockwise, so the coordinate-system offset is +90.
    const cssAngle = ((((fillStyle.angle ?? 0) + 90) % 360) + 360) % 360;
    return `linear-gradient(${cssAngle}deg, ${stops.join(', ')})`;
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

function boundedScale(value: number | undefined): number {
  return value === undefined || !Number.isFinite(value) ? 1 : Math.max(0.01, Math.min(1, value));
}

function lineSpacingValue(value: TextSpacingValue, reduction: number): string {
  if (typeof value === 'object') {
    if (value.unit === 'points') return `${value.value}pt`;
    return String(value.value * POWERPOINT_SINGLE_LINE_HEIGHT * (1 - reduction));
  }
  // The public model currently carries both OOXML point and percentage spacing
  // as numbers. Percentage spacing is conventionally 50 or greater; smaller
  // values are fixed point sizes.
  return value >= 50 ? `${value * POWERPOINT_SINGLE_LINE_HEIGHT * (1 - reduction)}%` : `${value}pt`;
}

function paragraphSpacingValue(value: TextSpacingValue): string {
  if (typeof value === 'object') {
    return value.unit === 'points'
      ? `${value.value}pt`
      : `${value.value * POWERPOINT_SINGLE_LINE_HEIGHT}em`;
  }
  // Percentage paragraph spacing is relative to the current line height. CSS
  // vertical percentage margins are relative to width, so express it in em.
  return value >= 50 ? `${(value / 100) * POWERPOINT_SINGLE_LINE_HEIGHT}em` : `${value}pt`;
}

function bulletCharacter(paragraph: TextParagraph, counters: Map<number, number>): string {
  const bullet = paragraph.bullet;
  if (!bullet) return '';
  if (bullet.kind === 'number') {
    const level = paragraph.level ?? 0;
    const next = counters.get(level) ?? bullet.startAt ?? 1;
    counters.set(level, next + 1);
    for (const key of counters.keys()) if (key > level) counters.delete(key);
    return `${next}.`;
  }
  if (bullet.kind === 'picture') return '•';
  const value = bullet.value ?? '•';
  // Wingdings/Symbol private-use bullets have no useful glyph when their
  // original font is unavailable in a browser. Preserve ordinary Unicode
  // bullets and use a stable round bullet for private-use code points.
  return /[\uE000-\uF8FF]/u.test(value) ? '•' : value;
}

function isEastAsianLanguage(language: string): boolean {
  return /^(?:ja|ko|zh)(?:-|$)/i.test(language);
}

function isComplexScriptLanguage(language: string): boolean {
  return /^(?:ar|dv|fa|he|ku|ps|syr|ug|ur|yi)(?:-|$)/i.test(language);
}

function runFontFamily(run: TextRun): string | undefined {
  const language = run.language ?? run.alternativeLanguage ?? '';
  const symbolFirst = /^[\u2000-\u206f\u2190-\u2bff\ue000-\uf8ff\s]+$/u.test(run.text);
  const ordered = symbolFirst
    ? [run.symbolFontFamily, run.fontFamily, run.eastAsianFontFamily, run.complexScriptFontFamily]
    : isEastAsianLanguage(language)
      ? [run.eastAsianFontFamily, run.fontFamily, run.complexScriptFontFamily, run.symbolFontFamily]
      : isComplexScriptLanguage(language) || run.rightToLeft
        ? [
            run.complexScriptFontFamily,
            run.fontFamily,
            run.eastAsianFontFamily,
            run.symbolFontFamily,
          ]
        : [
            run.fontFamily,
            run.eastAsianFontFamily,
            run.complexScriptFontFamily,
            run.symbolFontFamily,
          ];
  const families = ordered.filter(
    (family, index): family is string => Boolean(family) && ordered.indexOf(family) === index,
  );
  return families.length > 0 ? families.join(', ') : undefined;
}

function createTextRunElement(
  run: TextRun,
  paragraph: TextParagraph,
  fontScale: number,
): HTMLElement {
  const href = safeHyperlink(run.hyperlink);
  const span = document.createElement(href ? 'a' : 'span');
  span.textContent = run.text;
  const fontFamily = runFontFamily(run);
  if (fontFamily) span.style.fontFamily = fontFamily;
  if (run.fontSizePt !== undefined) span.style.fontSize = `${run.fontSizePt * fontScale}pt`;
  if (run.characterSpacingPt !== undefined)
    span.style.letterSpacing = `${run.characterSpacingPt}pt`;
  if (run.kerningThresholdPt !== undefined) {
    const effectiveFontSize = run.fontSizePt === undefined ? undefined : run.fontSizePt * fontScale;
    span.style.fontKerning =
      effectiveFontSize === undefined || effectiveFontSize >= run.kerningThresholdPt
        ? 'normal'
        : 'none';
  }
  if (run.bold) span.style.fontWeight = '700';
  if (run.italic) span.style.fontStyle = 'italic';
  const decorations = [run.underline ? 'underline' : '', run.strike ? 'line-through' : ''].filter(
    Boolean,
  );
  if (decorations.length) span.style.textDecoration = decorations.join(' ');
  if (run.color) span.style.color = color(run.color) ?? '';
  if (run.baseline) {
    span.style.verticalAlign = run.baseline > 0 ? 'super' : 'sub';
    span.style.fontSize = span.style.fontSize || '0.75em';
  }
  const language = run.language ?? run.alternativeLanguage;
  if (language) span.lang = language;
  if (run.rightToLeft !== undefined) {
    span.dir = run.rightToLeft ? 'rtl' : 'ltr';
    span.style.unicodeBidi = 'embed';
  }
  if (paragraph.latinLineBreak && !isEastAsianLanguage(language ?? '')) {
    span.style.overflowWrap = 'anywhere';
  }
  if (paragraph.eastAsianLineBreak === false && isEastAsianLanguage(language ?? '')) {
    span.style.wordBreak = 'keep-all';
  }
  if (href && span instanceof HTMLAnchorElement) {
    span.href = href;
    span.target = '_blank';
    span.rel = 'noreferrer noopener';
    // Browser link defaults are not part of DrawingML. PowerPoint stores
    // hyperlink color and underline explicitly in the run properties.
    if (!run.color) span.style.color = 'inherit';
    if (!decorations.length) span.style.textDecoration = 'none';
  }
  return span;
}

function splitRunsIntoTabRows(runs: TextRun[]): TextRun[][][] {
  const rows: TextRun[][][] = [[[]]];
  for (const run of runs) {
    for (const token of run.text.split(/([\t\n])/u)) {
      if (!token) continue;
      if (token === '\t') {
        rows.at(-1)!.push([]);
      } else if (token === '\n') {
        rows.push([[]]);
      } else {
        rows
          .at(-1)!
          .at(-1)!
          .push({ ...run, text: token });
      }
    }
  }
  return rows;
}

function appendExplicitTabLayout(
  line: HTMLElement,
  paragraph: TextParagraph,
  fontScale: number,
): boolean {
  const stops = [...(paragraph.tabStops ?? [])]
    .filter((stop) => Number.isFinite(stop.positionEmu) && stop.positionEmu >= 0)
    .sort((first, second) => first.positionEmu - second.positionEmu);
  if (stops.length === 0 || !paragraph.runs.some((run) => run.text.includes('\t'))) return false;

  const widths: number[] = [];
  let previous = 0;
  for (const stop of stops) {
    widths.push(Math.max(0, (stop.positionEmu - previous) / EMU_PER_CSS_PIXEL));
    previous = stop.positionEmu;
  }
  const template = `${widths.map((width) => `${width}px`).join(' ')} minmax(0, 1fr)`;
  for (const cells of splitRunsIntoTabRows(paragraph.runs)) {
    const row = document.createElement('span');
    row.dataset.rpvTabRow = '';
    row.style.display = 'grid';
    row.style.position = 'relative';
    row.style.gridTemplateColumns = template;
    row.style.minWidth = '0';
    row.style.minHeight = '1em';
    cells.forEach((runs, cellIndex) => {
      const cell = document.createElement('span');
      cell.dataset.rpvTabCell = String(cellIndex);
      cell.style.minWidth = '0';
      if (cellIndex === 0) {
        cell.style.gridColumn = '1';
      } else {
        const stopIndex = Math.min(cellIndex - 1, stops.length - 1);
        const stop = stops[stopIndex]!;
        if (stop.alignment === 'left') {
          cell.style.gridColumn = String(Math.min(stopIndex + 2, stops.length + 1));
        } else if (stop.alignment === 'center') {
          cell.style.position = 'absolute';
          cell.style.left = `${stop.positionEmu / EMU_PER_CSS_PIXEL}px`;
          cell.style.transform = 'translateX(-50%)';
        } else {
          cell.style.gridColumn = String(Math.max(1, stopIndex + 1));
          cell.style.justifySelf = 'end';
        }
      }
      for (const run of runs) cell.append(createTextRunElement(run, paragraph, fontScale));
      row.append(cell);
    });
    line.append(row);
  }
  return true;
}

function applyText(
  container: HTMLElement,
  paragraphs: TextParagraph[],
  options: TextRenderOptions = {},
): void {
  const fontScale = boundedScale(options.fontScale);
  const lineSpacingReduction = Math.max(0, Math.min(1, options.lineSpacingReduction ?? 0));
  const counters = new Map<number, number>();
  const wraps = !/^(?:none|nowrap)$/i.test(options.textWrap ?? 'square');
  container.style.whiteSpace = wraps ? 'pre-wrap' : 'pre';
  container.style.overflowWrap = 'normal';
  for (const [paragraphIndex, paragraph] of paragraphs.entries()) {
    const line = document.createElement('div');
    line.dataset.rpvTextParagraph = '';
    line.style.minWidth = '0';
    line.style.fontWeight = '400';
    line.style.lineHeight = String(POWERPOINT_SINGLE_LINE_HEIGHT * (1 - lineSpacingReduction));
    line.style.textAlign =
      paragraph.alignment === 'distributed' ? 'justify' : (paragraph.alignment ?? 'left');
    if (paragraph.alignment === 'distributed') {
      line.style.textAlignLast = 'justify';
      line.style.setProperty('text-justify', 'inter-character');
    } else if (paragraph.alignment === 'justify') {
      line.style.setProperty('text-justify', 'inter-word');
    }
    line.style.direction = paragraph.rtl ? 'rtl' : 'ltr';
    line.style.overflowWrap = wraps && paragraph.latinLineBreak ? 'anywhere' : 'normal';
    if (paragraph.eastAsianLineBreak !== undefined) {
      line.style.wordBreak = paragraph.eastAsianLineBreak ? 'normal' : 'keep-all';
      line.style.setProperty('line-break', paragraph.eastAsianLineBreak ? 'strict' : 'auto');
    }
    if (paragraph.hangingPunctuation) {
      line.style.setProperty('hanging-punctuation', 'first allow-end last');
    }
    if (paragraph.defaultTabSizeEmu !== undefined) {
      line.style.tabSize = `${paragraph.defaultTabSizeEmu / EMU_PER_CSS_PIXEL}px`;
    }
    line.style.marginLeft =
      paragraph.marginLeftEmu !== undefined
        ? `${paragraph.marginLeftEmu / EMU_PER_CSS_PIXEL}px`
        : paragraph.level
          ? `${paragraph.level * 1.25}em`
          : '0';
    if (paragraph.indentEmu !== undefined) {
      line.style.textIndent = `${paragraph.indentEmu / EMU_PER_CSS_PIXEL}px`;
    }
    if (
      paragraph.spaceBefore !== undefined &&
      (options.spaceFirstLastParagraph !== false || paragraphIndex !== 0)
    ) {
      line.style.marginTop = paragraphSpacingValue(paragraph.spaceBefore);
    }
    if (
      paragraph.spaceAfter !== undefined &&
      (options.spaceFirstLastParagraph !== false || paragraphIndex !== paragraphs.length - 1)
    ) {
      line.style.marginBottom = paragraphSpacingValue(paragraph.spaceAfter);
    }
    if (paragraph.lineSpacing !== undefined) {
      line.style.lineHeight = lineSpacingValue(paragraph.lineSpacing, lineSpacingReduction);
    }
    if (paragraph.bullet) {
      const marker = document.createElement('span');
      marker.dataset.rpvBullet = '';
      marker.textContent = `${bulletCharacter(paragraph, counters)}\u00a0`;
      if (!paragraph.rtl && paragraph.indentEmu !== undefined && paragraph.indentEmu < 0) {
        marker.style.display = 'inline-block';
        marker.style.width = `${-paragraph.indentEmu / EMU_PER_CSS_PIXEL}px`;
      }
      if (paragraph.bullet.fontFamily) marker.style.fontFamily = paragraph.bullet.fontFamily;
      if (paragraph.bullet.fontSizePt !== undefined) {
        marker.style.fontSize = `${paragraph.bullet.fontSizePt * fontScale}pt`;
      } else if (paragraph.bullet.sizePercent !== undefined) {
        const surroundingSize = paragraph.runs.find(
          (run) => run.fontSizePt !== undefined,
        )?.fontSizePt;
        marker.style.fontSize =
          surroundingSize !== undefined
            ? `${surroundingSize * fontScale * paragraph.bullet.sizePercent}pt`
            : `${paragraph.bullet.sizePercent * 100}%`;
      }
      marker.setAttribute('aria-hidden', 'true');
      line.append(marker);
    }
    const usedExplicitTabs =
      !paragraph.bullet && appendExplicitTabLayout(line, paragraph, fontScale);
    if (!usedExplicitTabs) {
      for (const run of paragraph.runs)
        line.append(createTextRunElement(run, paragraph, fontScale));
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

function normalizedPresetGeometryPath(preset?: string): string | undefined {
  switch (preset) {
    case 'straightConnector1':
      return 'M 0 0 L 1 1';
    case 'heart':
      // ECMA-376 preset geometry, normalized from w/h coordinates to 0..1.
      return 'M 0.5 0.25 C 0.7083333333 -0.3333333333 1.5208333333 0.25 0.5 1 C -0.5208333333 0.25 0.2916666667 -0.3333333333 0.5 0.25 Z';
    default:
      return undefined;
  }
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
    case 'diagStripe':
      element.style.clipPath = 'polygon(0 50%, 50% 0, 100% 0, 0 100%)';
      break;
    case 'hexagon':
      element.style.clipPath = 'polygon(25% 0, 75% 0, 100% 50%, 75% 100%, 25% 100%, 0 50%)';
      break;
    case 'chevron':
      element.style.clipPath = 'polygon(0 0, 70% 0, 100% 50%, 70% 100%, 0 100%, 30% 50%)';
      break;
    case 'rightArrow':
      element.style.clipPath =
        'polygon(0 25%, 75% 25%, 75% 0, 100% 50%, 75% 100%, 75% 75%, 0 75%)';
      break;
    case 'leftArrow':
      element.style.clipPath =
        'polygon(25% 0, 25% 25%, 100% 25%, 100% 75%, 25% 75%, 25% 100%, 0 50%)';
      break;
    case 'upArrow':
      element.style.clipPath =
        'polygon(50% 0, 100% 25%, 75% 25%, 75% 100%, 25% 100%, 25% 25%, 0 25%)';
      break;
    case 'downArrow':
      element.style.clipPath =
        'polygon(25% 0, 75% 0, 75% 75%, 100% 75%, 50% 100%, 0 75%, 25% 75%)';
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

function applyTextInsets(
  element: HTMLElement,
  insets: { top: number; right: number; bottom: number; left: number } | undefined,
  useBodyDefaults: boolean,
): void {
  const resolved =
    insets ??
    (useBodyDefaults
      ? {
          top: DEFAULT_TEXT_VERTICAL_INSET_EMU,
          right: DEFAULT_TEXT_HORIZONTAL_INSET_EMU,
          bottom: DEFAULT_TEXT_VERTICAL_INSET_EMU,
          left: DEFAULT_TEXT_HORIZONTAL_INSET_EMU,
        }
      : undefined);
  if (!resolved) return;
  element.style.padding = `${resolved.top / EMU_PER_CSS_PIXEL}px ${resolved.right / EMU_PER_CSS_PIXEL}px ${resolved.bottom / EMU_PER_CSS_PIXEL}px ${resolved.left / EMU_PER_CSS_PIXEL}px`;
}

function applyTextOrientation(
  element: HTMLElement,
  textRotation?: number,
  verticalText?: string,
  rotationAsVerticalFlow = false,
): void {
  const vertical = verticalText?.toLowerCase();
  if (vertical && vertical !== 'horz') {
    element.style.writingMode = vertical.includes('mongolian') ? 'vertical-lr' : 'vertical-rl';
    element.style.textOrientation = vertical.includes('wordartvert') ? 'upright' : 'mixed';
  }
  if (textRotation === undefined || !Number.isFinite(textRotation) || textRotation === 0) return;
  const normalized = ((textRotation % 360) + 360) % 360;
  if (rotationAsVerticalFlow && (normalized === 90 || normalized === 270)) {
    element.style.writingMode = 'vertical-rl';
    element.style.textOrientation = 'mixed';
    if (normalized === 270) element.style.transform = 'rotate(180deg)';
    return;
  }
  element.style.transform = `rotate(${textRotation}deg)`;
  element.style.transformOrigin = 'center';
}

function allowsTextOverflow(value: string | undefined): boolean {
  return value === undefined || /overflow/i.test(value);
}

function applyCroppedImage(
  image: HTMLImageElement,
  crop: { top: number; right: number; bottom: number; left: number },
): void {
  const horizontal = Math.max(0.000_001, 1 - crop.left - crop.right);
  const vertical = Math.max(0.000_001, 1 - crop.top - crop.bottom);
  image.style.position = 'absolute';
  image.style.maxWidth = 'none';
  image.style.maxHeight = 'none';
  image.style.width = `${100 / horizontal}%`;
  image.style.height = `${100 / vertical}%`;
  image.style.left = `${(-crop.left / horizontal) * 100}%`;
  image.style.top = `${(-crop.top / vertical) * 100}%`;
  image.style.objectFit = 'fill';
}

function isWordCharacter(value: string | undefined): boolean {
  return Boolean(value && /[\p{L}\p{N}_]/u.test(value));
}

function rgbChannels(value: string | undefined): [number, number, number] | undefined {
  const hex = value?.trim().replace(/^#/, '');
  if (!hex || hex.length !== 6 || !/^[\da-f]{6}$/i.test(hex)) return undefined;
  return [
    Number.parseInt(hex.slice(0, 2), 16),
    Number.parseInt(hex.slice(2, 4), 16),
    Number.parseInt(hex.slice(4, 6), 16),
  ];
}

/**
 * Builds a hidden SVG carrying a DrawingML duotone filter: the picture is
 * reduced to luminance, then mapped from the dark color to the light color.
 */
function createDuotoneFilter(
  id: string,
  dark: [number, number, number],
  light: [number, number, number],
): SVGSVGElement {
  const namespace = 'http://www.w3.org/2000/svg';
  const svg = document.createElementNS(namespace, 'svg');
  svg.setAttribute('width', '0');
  svg.setAttribute('height', '0');
  svg.style.position = 'absolute';
  svg.setAttribute('aria-hidden', 'true');
  const filter = document.createElementNS(namespace, 'filter');
  filter.id = id;
  filter.setAttribute('color-interpolation-filters', 'sRGB');
  const luminance = document.createElementNS(namespace, 'feColorMatrix');
  luminance.setAttribute('type', 'matrix');
  luminance.setAttribute(
    'values',
    '0.2126 0.7152 0.0722 0 0 0.2126 0.7152 0.0722 0 0 0.2126 0.7152 0.0722 0 0 0 0 0 1 0',
  );
  const transfer = document.createElementNS(namespace, 'feComponentTransfer');
  const channels: Array<['feFuncR' | 'feFuncG' | 'feFuncB', 0 | 1 | 2]> = [
    ['feFuncR', 0],
    ['feFuncG', 1],
    ['feFuncB', 2],
  ];
  for (const [name, channel] of channels) {
    const func = document.createElementNS(namespace, name);
    func.setAttribute('type', 'table');
    func.setAttribute('tableValues', `${dark[channel] / 255} ${light[channel] / 255}`);
    transfer.append(func);
  }
  filter.append(luminance, transfer);
  svg.append(filter);
  return svg;
}

/**
 * Builds a deterministic sRGB threshold filter for DrawingML `a:biLevel`.
 * A discrete 256-entry transfer avoids the browser-dependent rounding of a
 * very large CSS contrast filter around the threshold boundary.
 */
function createBiLevelFilter(id: string, threshold: number): SVGSVGElement {
  const namespace = 'http://www.w3.org/2000/svg';
  const svg = document.createElementNS(namespace, 'svg');
  svg.setAttribute('width', '0');
  svg.setAttribute('height', '0');
  svg.style.position = 'absolute';
  svg.setAttribute('aria-hidden', 'true');
  const filter = document.createElementNS(namespace, 'filter');
  filter.id = id;
  filter.setAttribute('color-interpolation-filters', 'sRGB');
  const luminance = document.createElementNS(namespace, 'feColorMatrix');
  luminance.setAttribute('type', 'matrix');
  luminance.setAttribute(
    'values',
    '0.2126 0.7152 0.0722 0 0 0.2126 0.7152 0.0722 0 0 0.2126 0.7152 0.0722 0 0 0 0 0 1 0',
  );
  const transfer = document.createElementNS(namespace, 'feComponentTransfer');
  const tableValues = Array.from({ length: 256 }, (_, index) =>
    index / 255 < threshold ? '0' : '1',
  ).join(' ');
  for (const name of ['feFuncR', 'feFuncG', 'feFuncB'] as const) {
    const func = document.createElementNS(namespace, name);
    func.setAttribute('type', 'discrete');
    func.setAttribute('tableValues', tableValues);
    transfer.append(func);
  }
  filter.append(luminance, transfer);
  svg.append(filter);
  return svg;
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
  private listCleanup: (() => void) | undefined;
  private highlight: HTMLElement | undefined;
  private pendingResources = new Set<Promise<void>>();
  private mountedHandles = new Set<DisposableHandle>();
  private handlesByTarget = new Map<HTMLElement, DisposableHandle>();
  private destroyed = false;
  private renderingSlideIndex: number | undefined;
  private customGeometrySequence = 0;
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
    this.listCleanup?.();
    this.listCleanup = undefined;
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
    const convert = contentType.includes('wmf') ? renderWmfToDataUrl : renderEmfToDataUrl;
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

  /**
   * Applies DrawingML picture recolor effects (`a:biLevel`, `a:grayscl`,
   * `a:duotone`, `a:lum`) to the element that paints the picture, matching
   * PowerPoint's rendering as closely as CSS/SVG filters allow.
   */
  private applyImageEffects(
    target: HTMLElement | SVGElement,
    effects: ImageEffects | undefined,
    filterHost: HTMLElement,
  ): void {
    if (!effects) return;
    const filters: string[] = [];
    if (effects.duotone && effects.duotone.length >= 2) {
      const dark = rgbChannels(effects.duotone[0]?.value);
      const light = rgbChannels(effects.duotone[1]?.value);
      if (dark && light) {
        const id = `rpv-duotone-${++this.customGeometrySequence}`;
        filterHost.append(createDuotoneFilter(id, dark, light));
        filters.push(`url(#${id})`);
      }
    }
    if (effects.biLevelThreshold !== undefined) {
      const threshold = Math.min(0.98, Math.max(0.02, effects.biLevelThreshold));
      const id = `rpv-bilevel-${++this.customGeometrySequence}`;
      filterHost.append(createBiLevelFilter(id, threshold));
      filters.push(`url(#${id})`);
    } else if (effects.grayscale) {
      filters.push('grayscale(1)');
    }
    if (effects.brightness) {
      filters.push(`brightness(${1 + Math.max(-1, Math.min(1, effects.brightness))})`);
    }
    if (effects.contrast) {
      filters.push(`contrast(${1 + Math.max(-1, Math.min(1, effects.contrast))})`);
    }
    if (filters.length > 0) target.style.filter = filters.join(' ');
  }

  private applyFill(element: HTMLElement, fillStyle: FillStyle | undefined, nodeId: string): void {
    if (fillStyle?.type === 'image') {
      const opacity = Math.max(0, Math.min(1, fillStyle.opacity ?? 1));
      const target =
        opacity < 1 || fillStyle.effects
          ? (() => {
              const layer = document.createElement('div');
              layer.dataset.rpvImageFill = '';
              layer.style.position = 'absolute';
              layer.style.inset = '0';
              layer.style.pointerEvents = 'none';
              layer.style.opacity = String(opacity);
              layer.style.zIndex = '0';
              if (!element.style.position) element.style.position = 'relative';
              element.prepend(layer);
              return layer;
            })()
          : element;
      target.style.backgroundRepeat = fillStyle.mode === 'tile' ? 'repeat' : 'no-repeat';
      if (fillStyle.mode === 'tile') target.style.backgroundSize = 'auto';
      else if (fillStyle.crop) {
        const horizontal = Math.max(0.000_001, 1 - fillStyle.crop.left - fillStyle.crop.right);
        const vertical = Math.max(0.000_001, 1 - fillStyle.crop.top - fillStyle.crop.bottom);
        target.style.backgroundSize = `${100 / horizontal}% ${100 / vertical}%`;
        const horizontalCrop = fillStyle.crop.left + fillStyle.crop.right;
        const verticalCrop = fillStyle.crop.top + fillStyle.crop.bottom;
        target.style.backgroundPosition = `${
          Math.abs(horizontalCrop) > 0.000_001 ? (fillStyle.crop.left / horizontalCrop) * 100 : 50
        }% ${Math.abs(verticalCrop) > 0.000_001 ? (fillStyle.crop.top / verticalCrop) * 100 : 50}%`;
      } else {
        target.style.backgroundSize = '100% 100%';
      }
      this.applyImageEffects(target, fillStyle.effects, element);
      this.applyAsset(
        fillStyle.assetId,
        (url) => {
          target.style.backgroundImage = `url(${JSON.stringify(url)})`;
        },
        nodeId,
      );
      return;
    }
    element.style.background = fill(fillStyle) ?? 'transparent';
  }

  private createCustomGeometry(
    node: Extract<SlideNode, { type: 'shape' }>,
    pathData = node.geometry.path ?? '',
  ): SVGSVGElement {
    const namespace = 'http://www.w3.org/2000/svg';
    const svg = document.createElementNS(namespace, 'svg');
    svg.dataset.rpvCustomGeometry = '';
    svg.setAttribute('viewBox', '0 0 1 1');
    svg.setAttribute('preserveAspectRatio', 'none');
    svg.style.position = 'absolute';
    svg.style.inset = '0';
    svg.style.width = '100%';
    svg.style.height = '100%';
    svg.style.overflow = 'visible';
    svg.style.pointerEvents = 'none';

    let renderedPathData = pathData;
    if (node.geometry.preset === 'straightConnector1') {
      if (node.transform.width === 0 && node.transform.height !== 0) {
        svg.style.left = '-0.5px';
        svg.style.width = '1px';
        renderedPathData = 'M 0.5 0 L 0.5 1';
      } else if (node.transform.height === 0 && node.transform.width !== 0) {
        svg.style.top = '-0.5px';
        svg.style.height = '1px';
        renderedPathData = 'M 0 0.5 L 1 0.5';
      }
    }

    const path = document.createElementNS(namespace, 'path');
    path.setAttribute('d', renderedPathData);
    path.setAttribute('vector-effect', 'non-scaling-stroke');
    path.style.fill = 'none';

    const fillStyle = node.fill;
    if (fillStyle?.type === 'solid') {
      path.style.fill = color(fillStyle.color) ?? 'none';
    } else if (fillStyle?.type === 'gradient' && fillStyle.stops.length) {
      const definitions = document.createElementNS(namespace, 'defs');
      const gradient = document.createElementNS(namespace, 'linearGradient');
      const gradientId = `rpv-gradient-${++this.customGeometrySequence}`;
      gradient.id = gradientId;
      const radians = ((fillStyle.angle ?? 0) * Math.PI) / 180;
      const horizontal = Math.cos(radians) / 2;
      const vertical = Math.sin(radians) / 2;
      gradient.setAttribute('x1', String(0.5 - horizontal));
      gradient.setAttribute('y1', String(0.5 - vertical));
      gradient.setAttribute('x2', String(0.5 + horizontal));
      gradient.setAttribute('y2', String(0.5 + vertical));
      for (const stopStyle of [...fillStyle.stops].sort(
        (first, second) => first.position - second.position,
      )) {
        const stop = document.createElementNS(namespace, 'stop');
        stop.setAttribute('offset', String(Math.max(0, Math.min(1, stopStyle.position))));
        stop.style.stopColor = color(stopStyle.color) ?? 'transparent';
        gradient.append(stop);
      }
      definitions.append(gradient);
      svg.append(definitions);
      path.style.fill = `url(#${gradientId})`;
    } else if (fillStyle?.type === 'pattern') {
      path.style.fill = color(fillStyle.foreground) ?? 'none';
    } else if (fillStyle?.type === 'image') {
      const definitions = document.createElementNS(namespace, 'defs');
      const clipPath = document.createElementNS(namespace, 'clipPath');
      const clipId = `rpv-clip-${++this.customGeometrySequence}`;
      clipPath.id = clipId;
      const clipShape = path.cloneNode() as SVGPathElement;
      clipShape.removeAttribute('style');
      clipPath.append(clipShape);
      definitions.append(clipPath);
      svg.append(definitions);

      const image = document.createElementNS(namespace, 'image');
      const horizontal = Math.max(
        0.000_001,
        1 - (fillStyle.crop?.left ?? 0) - (fillStyle.crop?.right ?? 0),
      );
      const vertical = Math.max(
        0.000_001,
        1 - (fillStyle.crop?.top ?? 0) - (fillStyle.crop?.bottom ?? 0),
      );
      image.setAttribute('x', String(-(fillStyle.crop?.left ?? 0) / horizontal));
      image.setAttribute('y', String(-(fillStyle.crop?.top ?? 0) / vertical));
      image.setAttribute('width', String(1 / horizontal));
      image.setAttribute('height', String(1 / vertical));
      image.setAttribute('preserveAspectRatio', 'none');
      image.setAttribute('clip-path', `url(#${clipId})`);
      image.style.opacity = String(Math.max(0, Math.min(1, fillStyle.opacity ?? 1)));
      this.applyAsset(
        fillStyle.assetId,
        (url) => image.setAttribute('href', url),
        node.id,
      );
      svg.append(image);
    }

    if (node.line) {
      path.style.stroke = color(node.line.color) ?? 'currentColor';
      path.style.strokeWidth = `${Math.max(0, node.line.width ?? 1)}px`;
      if (/dash/i.test(node.line.dash ?? '')) path.style.strokeDasharray = '6 4';
      else if (/dot/i.test(node.line.dash ?? '')) path.style.strokeDasharray = '1 3';
    }
    svg.append(path);
    return svg;
  }

  private themeForSlide(index: number | undefined): PresentationTheme | undefined {
    const slide = index === undefined ? undefined : this.presentation.slides[index];
    const master = this.presentation.masters.find((entry) => entry.id === slide?.masterId);
    return (
      this.presentation.themes.find((entry) => entry.id === master?.themeId) ??
      this.presentation.themes[0]
    );
  }

  private createChart(node: Extract<SlideNode, { type: 'chart' }>): HTMLElement {
    const container = document.createElement('div');
    container.style.width = '100%';
    container.style.height = '100%';
    container.setAttribute('role', 'img');
    container.setAttribute('aria-label', node.title ?? `${node.chartType} chart`);
    if (renderChartInto(container, node, this.themeForSlide(this.renderingSlideIndex))) {
      return container;
    }
    return this.createChartFallback(node);
  }

  private createChartFallback(node: Extract<SlideNode, { type: 'chart' }>): HTMLElement {
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
      .filter((value): value is number => typeof value === 'number' && Number.isFinite(value));
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
    const hasNegativeValues = values.some((value) => value < 0);
    const minimum = hasNegativeValues ? Math.min(0, ...values) : 0;
    // Keep the original PowerPoint-oracle geometry for positive-only charts.
    // Mixed-sign charts need a real zero baseline, but scaling sub-unit
    // positive values to the full plot made many existing charts enormous.
    const maximum = hasNegativeValues ? Math.max(0, ...values) : Math.max(1, ...values);
    const range = maximum - minimum || 1;
    const plotTop = hasNegativeValues ? 60 : 80;
    const plotHeight = hasNegativeValues ? 480 : 460;
    const yPosition = (value: number) => plotTop + ((maximum - value) / range) * plotHeight;
    const zeroY = yPosition(0);
    const pointCount = Math.max(1, ...node.series.map((series) => series.values.length));
    if (/line|scatter|area/i.test(node.chartType)) {
      node.series.forEach((series, seriesIndex) => {
        const points = series.values
          .map((value, index) =>
            typeof value !== 'number' || !Number.isFinite(value)
              ? null
              : `${80 + (index / Math.max(1, pointCount - 1)) * 840},${yPosition(value)}`,
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
          if (typeof value !== 'number' || !Number.isFinite(value)) return;
          const valueY = yPosition(value);
          const height = Math.abs(zeroY - valueY);
          const rect = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
          rect.setAttribute('x', String(80 + index * slot + seriesIndex * barWidth));
          rect.setAttribute('y', String(Math.min(zeroY, valueY)));
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
      element.style.color = 'inherit';
      element.style.textDecoration = 'none';
    }

    if (node.type === 'shape') {
      const presetPath = normalizedPresetGeometryPath(node.geometry.preset);
      if (node.geometry.path || presetPath) {
        element.append(this.createCustomGeometry(node, node.geometry.path ?? presetPath));
      } else {
        this.applyFill(element, node.fill, node.id);
        applyGeometry(element, node.geometry.preset);
        applyLine(element, node.line);
      }
      const shapeAutofit = node.autofit?.mode === 'shape';
      const verticalText = Boolean(node.verticalText && node.verticalText.toLowerCase() !== 'horz');
      const growsToFitText = shapeAutofit && !verticalText;
      if (growsToFitText) {
        const minimumHeight = element.style.height;
        element.style.height = 'auto';
        element.style.minHeight = minimumHeight;
        element.style.display = 'flex';
        element.style.flexDirection = 'column';
      }
      element.style.overflowX =
        shapeAutofit || allowsTextOverflow(node.horizontalOverflow) ? 'visible' : 'clip';
      element.style.overflowY =
        shapeAutofit || allowsTextOverflow(node.verticalOverflow) ? 'visible' : 'clip';
      if (node.paragraphs) {
        const textBody = document.createElement('div');
        textBody.dataset.rpvTextBody = '';
        textBody.style.position = growsToFitText ? 'relative' : 'absolute';
        if (!growsToFitText) textBody.style.inset = '0';
        textBody.style.width = '100%';
        textBody.style.height = growsToFitText ? 'auto' : '100%';
        if (growsToFitText) textBody.style.flex = '1 0 auto';
        textBody.style.boxSizing = 'border-box';
        textBody.style.display = 'flex';
        textBody.style.flexDirection = 'column';
        textBody.style.justifyContent =
          node.verticalAlignment === 'middle'
            ? 'center'
            : node.verticalAlignment === 'bottom'
              ? 'flex-end'
              : 'flex-start';
        applyTextInsets(textBody, node.textInsets, true);
        applyTextOrientation(textBody, node.textRotation, node.verticalText);

        const textContent = document.createElement('div');
        textContent.dataset.rpvTextContent = '';
        textContent.style.minWidth = '0';
        textContent.style.maxHeight = '100%';
        if (node.columnCount && node.columnCount > 1) {
          textContent.style.width = '100%';
          textContent.style.height = '100%';
          textContent.style.columnCount = String(Math.floor(node.columnCount));
          textContent.style.columnFill = 'auto';
          if (node.columnSpacing !== undefined) {
            textContent.style.columnGap = `${node.columnSpacing / EMU_PER_CSS_PIXEL}px`;
          }
          if (node.rightToLeftColumns) textContent.style.direction = 'rtl';
        }
        applyText(textContent, node.paragraphs, {
          ...(node.autofit?.mode === 'normal'
            ? {
                fontScale: node.autofit.fontScale,
                lineSpacingReduction: node.autofit.lineSpacingReduction,
              }
            : {}),
          ...(node.textWrap !== undefined ? { textWrap: node.textWrap } : {}),
          ...(node.spaceFirstLastParagraph !== undefined
            ? { spaceFirstLastParagraph: node.spaceFirstLastParagraph }
            : {}),
        });
        textBody.append(textContent);
        element.append(textBody);
      }
    } else if (node.type === 'image') {
      const image = document.createElement('img');
      this.applyAsset(node.assetId, (url) => (image.src = url), node.id);
      image.alt = node.altText ?? node.name ?? '';
      image.style.display = 'block';
      image.style.width = '100%';
      image.style.height = '100%';
      image.style.objectFit = node.preserveAspectRatio === false ? 'fill' : 'contain';
      element.style.overflow = 'hidden';
      if (node.crop) applyCroppedImage(image, node.crop);
      this.applyImageEffects(image, node.effects, element);
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
      const hasFixedRowHeights = totalRowHeight > 0;
      node.rows.forEach((row, rowIndex) => {
        const tr = table.insertRow();
        if (totalRowHeight && node.rowHeights?.[rowIndex] !== undefined) {
          tr.style.height = `${(node.rowHeights[rowIndex]! / totalRowHeight) * 100}%`;
        }
        for (const cell of row) {
          const td = tr.insertCell();
          td.colSpan = cell.colSpan ?? 1;
          td.rowSpan = cell.rowSpan ?? 1;
          if (hasFixedRowHeights) {
            td.style.position = 'relative';
            td.style.padding = '0';
          } else if (cell.textInsets) applyTextInsets(td, cell.textInsets, false);
          else td.style.padding = '0.25em';
          td.style.overflow = 'hidden';
          td.style.verticalAlign = cell.verticalAlignment ?? 'middle';
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
          const textBody = document.createElement('div');
          textBody.dataset.rpvTableTextBody = '';
          textBody.style.width = '100%';
          textBody.style.boxSizing = 'border-box';
          if (hasFixedRowHeights) {
            textBody.style.position = 'absolute';
            textBody.style.inset = '0';
            textBody.style.height = '100%';
            textBody.style.display = 'flex';
            textBody.style.flexDirection = 'column';
            textBody.style.justifyContent =
              cell.verticalAlignment === 'top'
                ? 'flex-start'
                : cell.verticalAlignment === 'bottom'
                  ? 'flex-end'
                  : 'center';
            if (cell.textInsets) applyTextInsets(textBody, cell.textInsets, false);
            else textBody.style.padding = '0.25em';
          }
          applyTextOrientation(textBody, cell.textRotation, undefined, true);
          applyText(textBody, cell.paragraphs);
          td.append(textBody);
        }
      });
      element.append(table);
    } else if (node.type === 'chart') {
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
    this.renderingSlideIndex = index;
    const element = document.createElement('section');
    element.className = 'rpv-normalized-slide';
    element.dataset.rpvSlideIndex = String(index);
    element.dataset.rpvSlideHidden = String(Boolean(slide.hidden));
    element.setAttribute('aria-label', `Slide ${index + 1} of ${this.slideCount}`);
    element.style.position = 'relative';
    element.style.width = `${this.naturalSlideWidth}px`;
    element.style.height = `${this.naturalSlideHeight}px`;
    element.style.overflow = 'hidden';
    element.style.background = '#fff';
    element.style.color = '#000';
    element.style.fontFamily = 'Arial, sans-serif';
    element.style.fontWeight = '400';
    element.style.lineHeight = String(POWERPOINT_SINGLE_LINE_HEIGHT);
    if (slide.background && slide.background.type !== 'none') {
      this.applyFill(element, slide.background, slide.id);
    }
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
    this.renderingSlideIndex = undefined;
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

  private createListPlaceholder(
    index: number,
    item: HTMLElement,
    generation: number,
    scale?: number,
  ): ListPlaceholder {
    let handle: DisposableHandle | undefined;
    let mountPromise: Promise<void> | undefined;
    const isActive = () =>
      this.isRenderActive(generation) && this.listState?.generation === generation;
    const mount = (): Promise<void> => {
      if (!isActive()) return Promise.resolve();
      if (handle) return handle.ready;
      if (mountPromise) return mountPromise;
      const nextHandle = this.mount(index, item, scale);
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
    return { item, mount, unmount, isMounted: () => Boolean(handle) };
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
    const scroller = options.scrollElement ?? this.container;
    const viewportWidth = scroller.clientWidth;
    const estimatedFitScale =
      this.fit === 'contain'
        ? Math.min(1, (viewportWidth || this.naturalSlideWidth) / this.naturalSlideWidth)
        : 1;
    const effectiveScale = estimatedFitScale * (this.zoom / 100);
    const slideHeight = this.naturalSlideHeight * effectiveScale;
    const itemStride = slideHeight + LIST_ITEM_GAP;

    if (!windowingEnabled) {
      // Non-windowed continuous mode: every slide mounts in normal flow.
      const placeholders = this.presentation.slides.map((_, index) => {
        const item = document.createElement('div');
        item.dataset.rpvListItem = String(index);
        item.style.minHeight = `${slideHeight}px`;
        item.style.margin = `0 auto ${LIST_ITEM_GAP}px`;
        this.container.append(item);
        return this.createListPlaceholder(index, item, generation);
      });
      const state: ActiveListState = { generation, options: normalizedOptions, placeholders };
      this.listState = state;
      await this.mountInBatches(
        placeholders,
        this.presentation.slides.map((_, index) => index),
        options.batchSize ?? 3,
        generation,
      );
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
      return;
    }

    // Windowed continuous mode driven by TanStack Virtual: slides live at
    // fixed absolute offsets inside a full-height sizer, so mounting and
    // unmounting content never shifts layout and scrolling stays smooth.
    const sizer = document.createElement('div');
    sizer.dataset.rpvVirtualSizer = '';
    sizer.style.position = 'relative';
    sizer.style.width = '100%';
    sizer.style.height = `${this.slideCount * itemStride - LIST_ITEM_GAP}px`;
    this.container.append(sizer);

    const placeholders = this.presentation.slides.map((_, index) => {
      const item = document.createElement('div');
      item.dataset.rpvListItem = String(index);
      item.style.position = 'absolute';
      item.style.top = '0';
      item.style.left = '0';
      item.style.width = '100%';
      item.style.height = `${slideHeight}px`;
      item.style.transform = `translateY(${index * itemStride}px)`;
      sizer.append(item);
      // Mount with the exact scale the shell geometry was computed from so
      // content and window heights always agree, even if the container is
      // still settling its layout.
      return this.createListPlaceholder(index, item, generation, effectiveScale);
    });

    const state: ActiveListState = { generation, options: normalizedOptions, placeholders };
    this.listState = state;
    const isActive = () => this.isRenderActive(generation) && this.listState === state;

    const measureViewportRect = () => {
      const rect = scroller.getBoundingClientRect();
      if (rect.width > 0 && rect.height > 0) return { width: rect.width, height: rect.height };
      if (scroller.clientHeight > 0) {
        return { width: scroller.clientWidth, height: scroller.clientHeight };
      }
      return { ...FALLBACK_VIEWPORT_RECT };
    };
    const initialRect = measureViewportRect();
    const requestedOverscan = options.overscanViewport ?? 1.5;
    const overscanViewports =
      Number.isFinite(requestedOverscan) && requestedOverscan >= 0 ? requestedOverscan : 1.5;
    const overscan = Math.max(1, Math.ceil((overscanViewports * initialRect.height) / itemStride));
    // Content rendered above the slide list (inside the same scroller) offsets
    // every virtual position; re-measured whenever the layout settles.
    const measureScrollMargin = () =>
      sizer.getBoundingClientRect().top - scroller.getBoundingClientRect().top + scroller.scrollTop;
    let scrollMargin = measureScrollMargin();
    // When the container is still collapsed while the document loads, the
    // initial alignment lands wrong; repeat it once real layout arrives as
    // long as the user has not scrolled yet.
    let needsSettleAlignment = initialIndex > 0;
    let lastProgrammaticTop: number | undefined;

    const reconcile = (instance: Virtualizer<HTMLElement, HTMLElement>): void => {
      if (!isActive()) return;
      const windowed = new Set(instance.getVirtualItems().map((item) => item.index));
      placeholders.forEach((placeholder, index) => {
        if (windowed.has(index)) void placeholder.mount();
        else placeholder.unmount();
      });
      // Report the slide with the largest visible overlap, matching how
      // PowerPoint's own thumbnail rail tracks the reading position.
      const offset = (instance.scrollOffset ?? 0) - scrollMargin;
      const viewportHeight = instance.scrollRect?.height ?? initialRect.height;
      const firstCandidate = Math.max(0, Math.floor(offset / itemStride));
      const lastCandidate = Math.min(
        this.slideCount - 1,
        Math.max(firstCandidate, Math.floor((offset + viewportHeight) / itemStride)),
      );
      let best = firstCandidate;
      let bestOverlap = Number.NEGATIVE_INFINITY;
      for (let index = firstCandidate; index <= lastCandidate; index += 1) {
        const start = index * itemStride;
        const overlap =
          Math.min(offset + viewportHeight, start + slideHeight) - Math.max(offset, start);
        if (overlap > bestOverlap + 0.5) {
          bestOverlap = overlap;
          best = index;
        }
      }
      if (best !== this.current) {
        this.current = best;
        this.notifySlideChange(best);
      }
    };

    const virtualizer = new Virtualizer<HTMLElement, HTMLElement>({
      count: this.slideCount,
      getScrollElement: () => scroller,
      estimateSize: () => slideHeight,
      gap: LIST_ITEM_GAP,
      overscan,
      scrollMargin,
      initialRect,
      initialOffset: initialIndex > 0 ? scrollMargin + initialIndex * itemStride : 0,
      observeElementRect: (instance, callback) => {
        const publish = () => {
          callback(measureViewportRect());
          const nextMargin = measureScrollMargin();
          if (Math.abs(nextMargin - scrollMargin) > 1) {
            scrollMargin = nextMargin;
            instance.setOptions({ ...instance.options, scrollMargin: nextMargin });
            instance.measure();
          }
          if (needsSettleAlignment && scroller.clientHeight > 0) {
            needsSettleAlignment = false;
            const undisturbed =
              scroller.scrollTop === 0 ||
              (lastProgrammaticTop !== undefined &&
                Math.abs(scroller.scrollTop - lastProgrammaticTop) <= 2);
            if (undisturbed) state.scrollToIndex?.(initialIndex);
          }
        };
        publish();
        if (typeof ResizeObserver === 'undefined') return;
        const observer = new ResizeObserver(publish);
        observer.observe(scroller);
        return () => observer.disconnect();
      },
      observeElementOffset: (_instance, callback) => {
        const onScroll = () => callback(scroller.scrollTop, true);
        callback(scroller.scrollTop, false);
        scroller.addEventListener('scroll', onScroll, { passive: true });
        return () => scroller.removeEventListener('scroll', onScroll);
      },
      scrollToFn: (offset, { adjustments, behavior }) => {
        const top = offset + (adjustments ?? 0);
        lastProgrammaticTop = top;
        if (typeof scroller.scrollTo === 'function') {
          scroller.scrollTo(behavior ? { top, behavior } : { top });
        } else {
          scroller.scrollTop = top;
        }
      },
      onChange: (instance) => reconcile(instance),
    });
    state.scrollToIndex = (index, scrollOptions) => {
      if (!isActive()) return;
      virtualizer.scrollToIndex(index, {
        align:
          scrollOptions?.block === 'center'
            ? 'center'
            : scrollOptions?.block === 'end'
              ? 'end'
              : scrollOptions?.block === 'nearest'
                ? 'auto'
                : 'start',
        behavior: scrollOptions?.behavior === 'smooth' ? 'smooth' : 'auto',
      });
    };
    const teardown = virtualizer._didMount();
    this.listCleanup = teardown;
    virtualizer._willUpdate();

    const initialWindow = virtualizer.getVirtualItems().map((item) => item.index);
    await this.mountInBatches(
      placeholders,
      [...initialWindow, initialIndex],
      options.batchSize ?? 3,
      generation,
    );
    if (!isActive()) return;
    if (initialIndex > 0) state.scrollToIndex(initialIndex);
    else if (!options.scrollElement) {
      this.container.scrollTop = 0;
      this.container.scrollLeft = 0;
    }
    this.current = initialIndex;
    this.notifySlideChange(initialIndex);
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
      // Only move the viewport when navigation actually changes the slide;
      // controlled hosts echo the visible slide back through goToSlide while
      // the user scrolls, and re-scrolling to it would snap the viewport.
      if (changed) {
        if (state.scrollToIndex) state.scrollToIndex(next, scrollOptions);
        else placeholder?.item.scrollIntoView?.(scrollOptions);
      }
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
    this.listCleanup?.();
    this.listCleanup = undefined;
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
