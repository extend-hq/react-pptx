export type PresentationFormat = 'pptx' | 'ppt';

export interface PresentationSize {
  widthEmu: number;
  heightEmu: number;
}

export interface PresentationWarning {
  code:
    | 'unsupported-feature'
    | 'degraded-rendering'
    | 'missing-asset'
    | 'missing-font'
    | 'corrupt-content'
    | 'encrypted-document'
    | 'resource-limit'
    | (string & {});
  message: string;
  severity: 'info' | 'warning' | 'error';
  slideIndex?: number;
  nodeId?: string;
  partName?: string;
  feature?: string;
}

export interface PresentationAsset {
  id: string;
  contentType: string;
  byteLength: number;
  fileName?: string;
  data?: Uint8Array;
  url?: string;
}

export interface PresentationEmbeddedFont {
  /** Typeface name declared by PowerPoint. */
  family: string;
  /** Asset containing a browser-loadable OpenType or TrueType font. */
  assetId: string;
  style: 'normal' | 'italic';
  weight: '400' | '700';
}

export interface Rect {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface Transform extends Rect {
  rotation?: number;
  flipHorizontal?: boolean;
  flipVertical?: boolean;
}

export interface ColorValue {
  value: string;
  alpha?: number;
}

/** DrawingML picture recolor effects declared on an `a:blip`. */
export interface ImageEffects {
  /** `a:biLevel` black/white threshold, normalized to 0..1. */
  biLevelThreshold?: number;
  /** `a:grayscl` grayscale recolor. */
  grayscale?: boolean;
  /** `a:duotone` two-color ramp, dark then light. */
  duotone?: [ColorValue, ColorValue] | ColorValue[];
  /** `a:lum` brightness, normalized to -1..1. */
  brightness?: number;
  /** `a:lum` contrast, normalized to -1..1. */
  contrast?: number;
}

export type FillStyle =
  | { type: 'none' }
  | { type: 'solid'; color: ColorValue }
  | {
      type: 'gradient';
      angle?: number;
      stops: Array<{ position: number; color: ColorValue }>;
    }
  | { type: 'pattern'; preset: string; foreground: ColorValue; background: ColorValue }
  | {
      type: 'image';
      assetId: string;
      mode: 'stretch' | 'tile';
      crop?: { top: number; right: number; bottom: number; left: number };
      /** Opacity applied to the image fill only, normalized to 0..1. */
      opacity?: number;
      effects?: ImageEffects;
    };

export interface LineStyle {
  color?: ColorValue;
  width?: number;
  dash?: string;
  startArrow?: string;
  endArrow?: string;
}

export interface TextRun {
  text: string;
  fontFamily?: string;
  fontSizePt?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: ColorValue;
  baseline?: number;
  /** OOXML run character spacing, in points. */
  characterSpacingPt?: number;
  language?: string;
  hyperlink?: string;
}

export interface TextSpacing {
  value: number;
  /** Percent values are normalized factors: 1 is 100%. */
  unit: 'points' | 'percent';
}

export interface TextParagraph {
  runs: TextRun[];
  alignment?: 'left' | 'center' | 'right' | 'justify' | 'distributed';
  level?: number;
  bullet?: {
    kind: 'character' | 'number' | 'picture';
    value?: string;
    fontFamily?: string;
    fontSizePt?: number;
    sizePercent?: number;
    startAt?: number;
  };
  /** Structured values preserve OOXML units; numbers remain accepted for 0.1.x compatibility. */
  lineSpacing?: TextSpacing | number;
  spaceBefore?: TextSpacing | number;
  spaceAfter?: TextSpacing | number;
  rtl?: boolean;
  /** Exact OOXML paragraph left margin, in EMUs. */
  marginLeftEmu?: number;
  /** Exact OOXML first-line or hanging indent, in EMUs. */
  indentEmu?: number;
}

export interface TextAutofit {
  mode: 'none' | 'normal' | 'shape';
  /** PowerPoint normal-autofit font scale, normalized to 0..1. */
  fontScale?: number;
  /** PowerPoint normal-autofit line-spacing reduction, normalized to 0..1. */
  lineSpacingReduction?: number;
}

export interface BaseSlideNode {
  id: string;
  name?: string;
  transform: Transform;
  hidden?: boolean;
  opacity?: number;
  hyperlink?: string;
  altText?: string;
  sourcePart?: string;
}

export interface ShapeNode extends BaseSlideNode {
  type: 'shape';
  geometry: { preset?: string; path?: string; adjustments?: Record<string, number> };
  fill?: FillStyle;
  line?: LineStyle;
  paragraphs?: TextParagraph[];
  verticalAlignment?: 'top' | 'middle' | 'bottom';
  textInsets?: { top: number; right: number; bottom: number; left: number };
  autofit?: TextAutofit;
  textRotation?: number;
  verticalText?: string;
  horizontalOverflow?: string;
  verticalOverflow?: string;
  textWrap?: 'none' | 'square' | (string & {});
  columnCount?: number;
  columnSpacing?: number;
  rightToLeftColumns?: boolean;
  spaceFirstLastParagraph?: boolean;
}

export interface ImageNode extends BaseSlideNode {
  type: 'image';
  assetId: string;
  crop?: { top: number; right: number; bottom: number; left: number };
  preserveAspectRatio?: boolean;
  effects?: ImageEffects;
}

export interface GroupNode extends BaseSlideNode {
  type: 'group';
  children: SlideNode[];
  childTransform?: Transform;
}

export interface TableCell {
  rowSpan?: number;
  colSpan?: number;
  fill?: FillStyle;
  borders?: Partial<Record<'top' | 'right' | 'bottom' | 'left', LineStyle>>;
  paragraphs: TextParagraph[];
  textInsets?: { top: number; right: number; bottom: number; left: number };
  verticalAlignment?: 'top' | 'middle' | 'bottom';
  textRotation?: number;
}

export interface TableNode extends BaseSlideNode {
  type: 'table';
  rows: TableCell[][];
  columnWidths?: number[];
  rowHeights?: number[];
}

export interface ChartSeries {
  name?: string;
  categories?: Array<string | number>;
  values: Array<number | null>;
  color?: ColorValue;
}

export interface ChartNode extends BaseSlideNode {
  type: 'chart';
  chartType: string;
  title?: string;
  series: ChartSeries[];
  hasLegend?: boolean;
  /**
   * Raw DrawingML chart part XML (`c:chartSpace` or `cx:chartSpace`).
   * Viewers use this for full-fidelity chart rendering; the parsed summary
   * fields above remain available for search and fallbacks.
   */
  chartXml?: string;
  /** Companion Microsoft chart style part (`style*.xml`), when present. */
  chartStyleXml?: string;
  /** Companion Microsoft chart color style part (`colors*.xml`), when present. */
  chartColorsXml?: string;
}

export interface MediaNode extends BaseSlideNode {
  type: 'media';
  mediaType: 'audio' | 'video' | 'ole' | 'unknown';
  assetId?: string;
  posterAssetId?: string;
}

export interface UnknownNode extends BaseSlideNode {
  type: 'unknown';
  feature: string;
  fallbackAssetId?: string;
}

export type SlideNode =
  ShapeNode | ImageNode | GroupNode | TableNode | ChartNode | MediaNode | UnknownNode;

export interface PresentationTheme {
  id: string;
  name?: string;
  colors: Record<string, string>;
  majorFonts?: Partial<Record<'latin' | 'eastAsia' | 'complexScript', string>>;
  minorFonts?: Partial<Record<'latin' | 'eastAsia' | 'complexScript', string>>;
}

export interface SlideMaster {
  id: string;
  name?: string;
  themeId?: string;
  nodes: SlideNode[];
}

export interface SlideLayout {
  id: string;
  name?: string;
  masterId?: string;
  nodes: SlideNode[];
}

export interface SlideNote {
  text: string;
}

export interface SlideComment {
  id: string;
  author?: string;
  text: string;
  createdAt?: string;
  x?: number;
  y?: number;
}

export interface PresentationSlide {
  id: string;
  index: number;
  name?: string;
  hidden?: boolean;
  masterId?: string;
  layoutId?: string;
  background?: FillStyle;
  nodes: SlideNode[];
  notes?: SlideNote[];
  comments?: SlideComment[];
  warnings?: PresentationWarning[];
  sourcePart?: string;
}

export interface PresentationMetadata {
  title?: string;
  subject?: string;
  creator?: string;
  company?: string;
  createdAt?: string;
  modifiedAt?: string;
}

export interface PresentationDocument {
  format: PresentationFormat;
  size: PresentationSize;
  slides: PresentationSlide[];
  masters: SlideMaster[];
  layouts: SlideLayout[];
  themes: PresentationTheme[];
  assets: Record<string, PresentationAsset>;
  /** Fonts embedded in the package and decoded by the native parser. */
  embeddedFonts?: PresentationEmbeddedFont[];
  warnings: PresentationWarning[];
  metadata?: PresentationMetadata;
}

export type Slide = PresentationSlide;

export interface PresentationSearchResult {
  slideIndex: number;
  nodeId: string;
  nodeType: SlideNode['type'];
  text: string;
  matchStart: number;
  matchEnd: number;
  snippet: string;
  bounds?: Rect;
}

export function collectSlideText(slide: PresentationSlide): string {
  const chunks: string[] = [];
  const visit = (nodes: SlideNode[]): void => {
    for (const node of nodes) {
      if (node.type === 'shape') {
        for (const paragraph of node.paragraphs ?? []) {
          chunks.push(paragraph.runs.map((run) => run.text).join(''));
        }
      } else if (node.type === 'table') {
        for (const row of node.rows) {
          for (const cell of row) {
            chunks.push(...cell.paragraphs.map((p) => p.runs.map((run) => run.text).join('')));
          }
        }
      } else if (node.type === 'group') {
        visit(node.children);
      } else if (node.type === 'chart') {
        if (node.title) chunks.push(node.title);
        for (const series of node.series) if (series.name) chunks.push(series.name);
      }
    }
  };
  visit(slide.nodes);
  return chunks.join('\n');
}
