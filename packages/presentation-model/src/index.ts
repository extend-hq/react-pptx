import type { SlideTiming, SlideTransition } from './animation';
export * from './animation';

/**
 * The source format the parser detected: PowerPoint (`pptx`), legacy
 * PowerPoint (`ppt`), OpenDocument presentation (`odp`) or OpenDocument
 * drawing (`odg`).
 */
export type PresentationFormat = 'pptx' | 'ppt' | 'odp' | 'odg';

export const PRESENTATION_FORMATS: readonly PresentationFormat[] = ['pptx', 'ppt', 'odp', 'odg'];

export function isPresentationFormat(value: unknown): value is PresentationFormat {
  return typeof value === 'string' && (PRESENTATION_FORMATS as readonly string[]).includes(value);
}

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

export interface MathElement {
  name: string;
  namespace?: string;
  attributes?: { name: string; namespace?: string; value: string }[];
  children?: MathElement[];
  text?: string;
}

export interface OfficeMath {
  text: string;
  display: boolean;
  sourceXml: string;
  root?: MathElement;
  justification?: string;
  mathMl: string;
  fontSizePt?: number;
  unsupported?: string[];
}

export interface TextRunProperties {
  fontFamily?: string;
  /** Explicit East Asian typeface declared for this run. */
  eastAsianFontFamily?: string;
  /** Explicit complex-script typeface declared for this run. */
  complexScriptFontFamily?: string;
  /** Explicit symbol typeface declared for this run. */
  symbolFontFamily?: string;
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
  alternativeLanguage?: string;
  /** Run-level bidi override from DrawingML. */
  rightToLeft?: boolean;
  /** Minimum font size, in points, at which pair kerning is enabled. */
  kerningThresholdPt?: number;
  hyperlink?: string;
}

export interface TextRun extends TextRunProperties {
  math?: OfficeMath;
  text: string;
  field?: TextField;
}

export interface TextField {
  id?: string;
  type?: string;
  hasCachedText: boolean;
  paragraphProperties?: TextFieldParagraphProperties;
  sourcePart?: string;
  sourceXml: string;
}

export interface TextFieldParagraphProperties extends Omit<TextParagraph, 'runs'> {
  marginRightEmu?: number;
  fontAlignment?: string;
  bulletNone?: boolean;
  bulletColor?: ColorValue;
  bulletColorFollowText?: boolean;
  bulletFontFollowText?: boolean;
  bulletSizeFollowText?: boolean;
  bulletImageAssetId?: string;
  defaultRunProperties?: TextRunProperties;
}

export interface TextTabStop {
  positionEmu: number;
  alignment: 'left' | 'center' | 'right' | 'decimal';
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
  /** Default interval used for tabs without an explicit stop, in EMUs. */
  defaultTabSizeEmu?: number;
  /** Explicit DrawingML paragraph tab stops. */
  tabStops?: TextTabStop[];
  eastAsianLineBreak?: boolean;
  latinLineBreak?: boolean;
  hangingPunctuation?: boolean;
}

export interface TextAutofit {
  mode: 'none' | 'normal' | 'shape';
  /** PowerPoint normal-autofit font scale, normalized to 0..1. */
  fontScale?: number;
  /** PowerPoint normal-autofit line-spacing reduction, normalized to 0..1. */
  lineSpacingReduction?: number;
}

export interface ObjectAction {
  relationshipId?: string;
  target?: string;
  external?: boolean;
  action?: string;
  tooltip?: string;
  targetFrame?: string;
  invalidUrl?: string;
  history?: boolean;
  highlightClick?: boolean;
  endSound?: boolean;
  soundAssetId?: string;
  soundName?: string;
  soundBuiltin?: boolean;
}

export interface BaseSlideNode {
  id: string;
  name?: string;
  transform: Transform;
  hidden?: boolean;
  opacity?: number;
  hyperlink?: string;
  altText?: string;
  accessibilityTitle?: string;
  description?: string;
  decorative?: boolean;
  clickAction?: ObjectAction;
  hoverAction?: ObjectAction;
  sourcePart?: string;
  sourceId?: string;
  editing?: ObjectEditing;
}

export interface ObjectEditing {
  kind: 'shape' | 'connector' | 'picture' | 'group' | 'graphicFrame' | 'contentPart';
  locks?: ObjectLocks;
  connector?: ConnectorConnections;
  textBox?: boolean;
  preferRelativeResize?: boolean;
  isComment?: boolean;
  sourceXml: string;
}

export interface ObjectLocks {
  noGrouping?: boolean;
  noUngrouping?: boolean;
  noSelection?: boolean;
  noRotation?: boolean;
  noChangeAspect?: boolean;
  noMove?: boolean;
  noResize?: boolean;
  noEditPoints?: boolean;
  noAdjustHandles?: boolean;
  noChangeArrowheads?: boolean;
  noChangeShapeType?: boolean;
  noTextEdit?: boolean;
  noCrop?: boolean;
  noDrilldown?: boolean;
}

export interface ConnectorConnections {
  start?: ConnectorEndpoint;
  end?: ConnectorEndpoint;
}

export interface ConnectorEndpoint {
  shapeId: number;
  connectionSite: number;
}

/** Resolves an endpoint within its original part; ambiguous IDs remain unresolved. */
export function resolveConnectorEndpoint(
  nodes: readonly SlideNode[],
  connector: BaseSlideNode,
  endpoint: ConnectorEndpoint,
): SlideNode | undefined {
  if (!connector.sourcePart || !Number.isInteger(endpoint.shapeId)
    || endpoint.shapeId < 0 || endpoint.shapeId > 0xffffffff) return undefined;
  let match: SlideNode | undefined;
  let ambiguous = false;
  const visit = (children: readonly SlideNode[]): void => {
    for (const node of children) {
      if (node.sourcePart === connector.sourcePart && node.sourceId !== undefined
        && /^[\t\n\r ]*[+-]?[0-9]+[\t\n\r ]*$/.test(node.sourceId)) {
        const id = Number(node.sourceId);
        if (id === endpoint.shapeId) {
          if (match) ambiguous = true;
          else match = node;
        }
      }
      if (node.type === 'group') visit(node.children);
    }
  };
  visit(nodes);
  return ambiguous ? undefined : match;
}

export interface ShapeNode extends BaseSlideNode {
  type: 'shape';
  shapeSvg?: string;
  paintBounds?: PicturePaintBounds;
  sourceXml?: string;
  geometry: { preset?: string; path?: string; adjustments?: Record<string, number> };
  fill?: FillStyle;
  line?: LineStyle;
  paragraphs?: TextParagraph[];
  verticalAlignment?: 'top' | 'middle' | 'bottom';
  textInsets?: { top: number; right: number; bottom: number; left: number };
  autofit?: TextAutofit;
  textRotation?: number;
  textTransform?: Transform;
  verticalText?: string;
  horizontalOverflow?: string;
  verticalOverflow?: string;
  textWrap?: 'none' | 'square' | (string & {});
  columnCount?: number;
  columnSpacing?: number;
  rightToLeftColumns?: boolean;
  spaceFirstLastParagraph?: boolean;
}

export interface PicturePaintBounds {
  xPx: number;
  yPx: number;
  widthPx: number;
  heightPx: number;
  layoutWidthPx: number;
  layoutHeightPx: number;
}

export interface Model3DAsset {
  partName: string;
  contentType: string;
  data?: Uint8Array;
}

export interface Model3DPreview {
  modelRelationshipId?: string;
  rasterRelationshipId: string;
  renderer?: string;
  rendererVersion?: string;
}

export interface Model3D {
  modelRelationshipId?: string;
  linkedModelRelationshipId?: string;
  asset?: Model3DAsset;
  preview?: Model3DPreview;
  sourceXml: string;
}

export interface ImageNode extends BaseSlideNode {
  type: 'image';
  model3d?: Model3D;
  ink?: {
    assetId?: string;
    sourceXml?: string;
    drawing?: {
      sourceBounds: { x: number; y: number; width: number; height: number };
      strokes: Array<{
        id?: string;
        timeOffset?: string;
        points: Array<{ x: number; y: number; pressure: number; contact: boolean }>;
        brush: { width: number; height: number; color: string; opacity: number; rectangle: boolean; pencil: boolean; ignorePressure: boolean };
      }>;
    };
  };
  assetId: string;
  crop?: { top: number; right: number; bottom: number; left: number };
  fallbackAssetId?: string;
  preserveAspectRatio?: boolean;
  effects?: ImageEffects;
  pictureSvg?: string;
  fallbackPicture?: { svg: string; paintBounds?: PicturePaintBounds | null };
  paintBounds?: PicturePaintBounds;
  sourceXml?: string;
}

export interface GroupNode extends BaseSlideNode {
  type: 'group';
  children: SlideNode[];
  childTransform?: Transform;
}

export interface TableCell {
  merge?: {
    horizontal: boolean;
    vertical: boolean;
    origin?: { row: number; column: number };
  };
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
  chartPartName?: string;
  /** Chart artwork at the dimensions of the slide frame. */
  chartSvg?: string;
  /** Companion Microsoft chart style part (`style*.xml`), when present. */
  chartStyleXml?: string;
  /** Companion Microsoft chart color style part (`colors*.xml`), when present. */
  chartColorsXml?: string;
}

export interface MediaNode extends BaseSlideNode {
  type: 'media';
  mediaType: 'audio' | 'video' | 'ole' | 'control' | 'unknown';
  assetId?: string;
  posterAssetId?: string;
  sourceUrl?: string;
  poster?: ImageNode;
  playback?: {
    trimStartMs?: number;
    /** Duration removed from the end of the media. */
    trimEndMs?: number;
    fadeInMs?: number;
    fadeOutMs?: number;
    bookmarks?: { name: string; timeMs: number }[];
  };
  embeddedObject?: {
    control?: {
      classId?: string;
      persistence?: string;
      sourceXml?: string;
      binaryAssetIds?: string[];
      properties?: Array<{ name: string; value?: string }>;
    };
    programId?: string;
    name?: string;
    relationshipId?: string;
    linked: boolean;
    showAsIcon?: boolean;
    previewWidthEmu?: number;
    previewHeightEmu?: number;
    followColorScheme?: string;
    updateAutomatic?: boolean;
  };
  sourceXml?: string;
}

export interface UnknownNode extends BaseSlideNode {
  type: 'unknown';
  feature: string;
  fallbackAssetId?: string;
}

export type SlideNode =
  ShapeNode | ImageNode | GroupNode | TableNode | ChartNode | MediaNode | UnknownNode;

export interface PresentationTheme {
  formatSchemeXml?: string;
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
  transition?: SlideTransition;
  timing?: SlideTiming;
  nodes: SlideNode[];
}

export interface SlideLayout {
  id: string;
  name?: string;
  masterId?: string;
  transition?: SlideTransition;
  timing?: SlideTiming;
  nodes: SlideNode[];
}

export interface SlideNote {
  text: string;
  paragraphs?: TextParagraph[];
}

export interface SlideComment {
  id: string;
  author?: string;
  authorId?: string;
  authorInitials?: string;
  sourceIndex?: number;
  sourcePart?: string;
  text: string;
  paragraphs?: TextParagraph[];
  replies?: SlideComment[];
  modern?: ModernComment;
  createdAt?: string;
  modifiedAt?: string;
  position?: { x: number; y: number; unit: 'eighthPoint' };
  x?: number;
  y?: number;
}

export interface ModernComment {
  status: string;
  authorUserId?: string;
  authorProviderId?: string;
  authorExtensionsXml?: string;
  anchors?: CommentAnchor[];
  positionEmu?: { x: number; y: number; sourceX: string; sourceY: string };
  task?: CommentTask;
  extensionsXml?: string;
}

export interface CommentTask {
  startDate?: string;
  dueDate?: string;
  assignedTo: string[];
  title: string;
  completion: string;
  completionFraction?: number;
}

export interface CommentAnchor {
  kind: 'slide' | 'drawingElement' | 'textRange' | 'unknown';
  monikers: CommentMoniker[];
}

export interface CommentMoniker {
  namespace: string;
  name: string;
  attributes: Record<string, string>;
  children?: CommentMoniker[];
  text?: string;
}

export interface PresentationSlide {
  id: string;
  index: number;
  name?: string;
  hidden?: boolean;
  masterId?: string;
  layoutId?: string;
  background?: FillStyle;
  transition?: SlideTransition;
  timing?: SlideTiming;
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
  documentId?: string;
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
            if (cell.merge?.horizontal || cell.merge?.vertical) continue;
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

export type * from "./blocks";
export { flattenPresentationBlocks } from "./blocks";
