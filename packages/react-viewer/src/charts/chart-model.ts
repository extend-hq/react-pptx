/**
 * Builds the vendored react-xlsx chart model (`XlsxChart`) from raw PowerPoint
 * chart part XML, then applies the same Microsoft chart style pipeline
 * (style/colors companion parts, theme palette, built-in Office defaults) that
 * react-xlsx uses, so charts render with identical styles and colors.
 *
 * The parsing helpers are ported from react-xlsx `charts.ts`; the base model
 * builders replace the Duke Sheets wasm output with a direct DrawingML parse
 * because PowerPoint chart parts always carry their cached data.
 */
import type { ChartNode, PresentationTheme } from '@extend-ai/react-pptx-model';
import type {
  XlsxChart,
  XlsxChartAxis,
  XlsxChartDataLabels,
  XlsxChartPointDataLabel,
  XlsxChartPointStyle,
  XlsxChartSeries,
  XlsxChartTypeGroup,
  XlsxChartWall,
  XlsxThemePalette,
} from './chart-types';

const EMU_PER_PIXEL = 9525;

/** Default Office series palette when the theme has no accent colors. */
const SERIES_COLORS = ['#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47'];

const THEME_COLOR_INDEX_BY_NAME: Record<string, number> = {
  accent1: 4,
  accent2: 5,
  accent3: 6,
  accent4: 7,
  accent5: 8,
  accent6: 9,
  dk1: 1,
  dk2: 3,
  folHlink: 11,
  hlink: 10,
  lt1: 0,
  lt2: 2,
  tx1: 1,
  tx2: 3,
  bg1: 0,
  bg2: 2,
};

const PRIMARY_CHART_TYPE_LOCAL_NAMES = [
  'barChart',
  'lineChart',
  'line3DChart',
  'stockChart',
  'radarChart',
  'scatterChart',
  'pieChart',
  'pie3DChart',
  'doughnutChart',
  'areaChart',
  'area3DChart',
  'bar3DChart',
  'ofPieChart',
  'bubbleChart',
  'surfaceChart',
  'surface3DChart',
] as const;

type ChartStyleAppearance = {
  axisLabelColor?: string | undefined;
  axisLineColor?: string | undefined;
  chartAreaBorderColor?: string | undefined;
  chartAreaFillColor?: string | undefined;
  chartAreaNoFill?: boolean | undefined;
  paletteOffset?: number | undefined;
  textColor?: string | undefined;
  titleColor?: string | undefined;
};

/** Maps a slide's resolved PowerPoint theme onto the Excel-style palette. */
export function buildThemePaletteFromPresentationTheme(
  theme: PresentationTheme | undefined,
): XlsxThemePalette | null {
  if (!theme) return null;
  const colorsByIndex: Record<number, string> = {};
  for (const [name, value] of Object.entries(theme.colors ?? {})) {
    const index = THEME_COLOR_INDEX_BY_NAME[name];
    if (index === undefined) continue;
    const normalized = normalizeHexColor(value);
    if (normalized && colorsByIndex[index] === undefined) colorsByIndex[index] = normalized;
  }
  if (Object.keys(colorsByIndex).length === 0) return null;
  return {
    colorsByIndex,
    majorLatinFont: theme.majorFonts?.latin,
    minorLatinFont: theme.minorFonts?.latin,
  };
}

function clampUnitInterval(value: number): number {
  return Math.max(0, Math.min(1, value));
}

function isElementNode(node: Node | ChildNode | null | undefined): node is Element {
  return node != null && node.nodeType === 1;
}

function parseXml(xml: string): XMLDocument | null {
  if (typeof DOMParser === 'undefined') return null;
  try {
    const document = new DOMParser().parseFromString(xml, 'application/xml');
    return document.getElementsByTagName('parsererror').length > 0 ? null : document;
  } catch {
    return null;
  }
}

function getLocalChildren(parent: ParentNode, localName: string): Element[] {
  return Array.from(parent.childNodes).filter(
    (node): node is Element => node.nodeType === 1 && (node as Element).localName === localName,
  );
}

function getLocalDescendants(parent: ParentNode, localName: string): Element[] {
  return Array.from((parent as Element | Document).getElementsByTagName('*')).filter(
    (node) => node.localName === localName,
  );
}

function getFirstLocalChild(parent: ParentNode | null, localName: string): Element | null {
  if (!parent) return null;
  return getLocalChildren(parent, localName)[0] ?? null;
}

function getFirstLocalDescendant(parent: ParentNode | null, localName: string): Element | null {
  if (!parent) return null;
  return getLocalDescendants(parent, localName)[0] ?? null;
}

function normalizeHexColor(value: string): string | null {
  const hex = value.replace(/^#/, '');
  if (hex.length === 8) return `#${hex.slice(2).toLowerCase()}`;
  if (hex.length === 6) return `#${hex.toLowerCase()}`;
  return null;
}

function parseHexColor(color: string): [number, number, number] | null {
  const normalized = normalizeHexColor(color);
  const match = normalized ? /^#([0-9a-f]{6})$/.exec(normalized) : null;
  if (!match?.[1]) return null;
  const hex = match[1];
  return [
    Number.parseInt(hex.slice(0, 2), 16),
    Number.parseInt(hex.slice(2, 4), 16),
    Number.parseInt(hex.slice(4, 6), 16),
  ];
}

function rgbToHsl(red: number, green: number, blue: number): [number, number, number] {
  const normalizedRed = red / 255;
  const normalizedGreen = green / 255;
  const normalizedBlue = blue / 255;
  const max = Math.max(normalizedRed, normalizedGreen, normalizedBlue);
  const min = Math.min(normalizedRed, normalizedGreen, normalizedBlue);
  const lightness = (max + min) / 2;
  if (max === min) return [0, 0, lightness];
  const delta = max - min;
  const saturation = lightness > 0.5 ? delta / (2 - max - min) : delta / (max + min);
  let hue = 0;
  switch (max) {
    case normalizedRed:
      hue = (normalizedGreen - normalizedBlue) / delta + (normalizedGreen < normalizedBlue ? 6 : 0);
      break;
    case normalizedGreen:
      hue = (normalizedBlue - normalizedRed) / delta + 2;
      break;
    default:
      hue = (normalizedRed - normalizedGreen) / delta + 4;
      break;
  }
  return [hue / 6, saturation, lightness];
}

function hueToRgb(p: number, q: number, t: number): number {
  let nextT = t;
  if (nextT < 0) nextT += 1;
  if (nextT > 1) nextT -= 1;
  if (nextT < 1 / 6) return p + (q - p) * 6 * nextT;
  if (nextT < 1 / 2) return q;
  if (nextT < 2 / 3) return p + (q - p) * (2 / 3 - nextT) * 6;
  return p;
}

function hslToRgb(hue: number, saturation: number, lightness: number): [number, number, number] {
  if (saturation === 0) {
    const gray = Math.round(lightness * 255);
    return [gray, gray, gray];
  }
  const q =
    lightness < 0.5 ? lightness * (1 + saturation) : lightness + saturation - lightness * saturation;
  const p = 2 * lightness - q;
  return [
    Math.round(hueToRgb(p, q, hue + 1 / 3) * 255),
    Math.round(hueToRgb(p, q, hue) * 255),
    Math.round(hueToRgb(p, q, hue - 1 / 3) * 255),
  ];
}

function rgbToHex(red: number, green: number, blue: number): string {
  return `#${[red, green, blue]
    .map((channel) => Math.max(0, Math.min(255, Math.round(channel))).toString(16).padStart(2, '0'))
    .join('')}`;
}

function applyLightnessTransform(baseColor: string, modifier = 1, offset = 0): string | null {
  const rgb = parseHexColor(baseColor);
  if (!rgb) return normalizeHexColor(baseColor);
  const [hue, saturation, lightness] = rgbToHsl(rgb[0], rgb[1], rgb[2]);
  const nextLightness = clampUnitInterval(lightness * modifier + offset);
  const [nextRed, nextGreen, nextBlue] = hslToRgb(hue, saturation, nextLightness);
  return rgbToHex(nextRed, nextGreen, nextBlue);
}

function resolveThemeColor(name: string | null, themePalette?: XlsxThemePalette | null): string | null {
  if (!name) return null;
  const index = THEME_COLOR_INDEX_BY_NAME[name];
  return index === undefined ? null : (themePalette?.colorsByIndex[index] ?? null);
}

function resolveThemeTypeface(
  typeface: string | null,
  themePalette?: XlsxThemePalette | null,
): string | null {
  if (!typeface) return null;
  if (typeface === '+mn-lt' || typeface === '+mn-ea' || typeface === '+mn-cs') {
    return themePalette?.minorLatinFont ?? null;
  }
  if (typeface === '+mj-lt' || typeface === '+mj-ea' || typeface === '+mj-cs') {
    return themePalette?.majorLatinFont ?? null;
  }
  return typeface;
}

function readChartTextTypeface(
  textPropertiesNode: Element | null,
  themePalette?: XlsxThemePalette | null,
): string | null {
  if (!textPropertiesNode) return null;
  const defaultRunProperties =
    getFirstLocalDescendant(textPropertiesNode, 'defRPr') ??
    getFirstLocalDescendant(textPropertiesNode, 'rPr');
  if (!defaultRunProperties) return null;
  const typeface =
    getFirstLocalChild(defaultRunProperties, 'latin')?.getAttribute('typeface') ??
    getFirstLocalChild(defaultRunProperties, 'ea')?.getAttribute('typeface') ??
    getFirstLocalChild(defaultRunProperties, 'cs')?.getAttribute('typeface') ??
    null;
  const resolved = resolveThemeTypeface(typeface, themePalette)?.trim() ?? '';
  return resolved.length > 0 ? resolved : null;
}

function resolveChartColorNode(
  node: Element | null,
  themePalette?: XlsxThemePalette | null,
): string | null {
  if (!node) return null;
  let baseColor: string | null = null;
  if (node.localName === 'srgbClr') {
    baseColor = normalizeHexColor(`#${node.getAttribute('val') ?? ''}`);
  } else if (node.localName === 'schemeClr') {
    baseColor = resolveThemeColor(node.getAttribute('val'), themePalette);
  } else if (node.localName === 'sysClr') {
    baseColor = normalizeHexColor(`#${node.getAttribute('lastClr') ?? ''}`);
  }
  if (!baseColor) return null;
  let lightnessModifier = 1;
  let lightnessOffset = 0;
  for (const transformNode of Array.from(node.childNodes).filter(isElementNode)) {
    const rawValue = Number(transformNode.getAttribute('val') ?? Number.NaN);
    if (!Number.isFinite(rawValue)) continue;
    if (transformNode.localName === 'lumMod') {
      lightnessModifier *= rawValue / 100000;
    } else if (transformNode.localName === 'lumOff') {
      lightnessOffset += rawValue / 100000;
    } else if (transformNode.localName === 'tint') {
      lightnessOffset += (1 - lightnessOffset) * (rawValue / 100000);
    } else if (transformNode.localName === 'shade') {
      lightnessModifier *= rawValue / 100000;
    }
  }
  return applyLightnessTransform(baseColor, lightnessModifier, lightnessOffset);
}

function isChartColorElement(node: Element | null | undefined): node is Element {
  return Boolean(
    node && (node.localName === 'schemeClr' || node.localName === 'srgbClr' || node.localName === 'sysClr'),
  );
}

function findFirstChartColorElement(node: Element | null): Element | null {
  if (!node) return null;
  if (isChartColorElement(node)) return node;
  for (const localName of ['srgbClr', 'schemeClr', 'sysClr']) {
    for (const candidate of getLocalDescendants(node, localName)) {
      if (isChartColorElement(candidate)) return candidate;
    }
  }
  return null;
}

function resolveChartFillColor(
  shapeNode: Element | null,
  themePalette?: XlsxThemePalette | null,
): string | null {
  if (!shapeNode || getFirstLocalChild(shapeNode, 'noFill')) return null;
  const solidFill = getFirstLocalChild(shapeNode, 'solidFill');
  if (solidFill) {
    const colorNode = findFirstChartColorElement(
      Array.from(solidFill.childNodes).find(isElementNode) ?? null,
    );
    return resolveChartColorNode(colorNode, themePalette);
  }
  const gradientFill = getFirstLocalChild(shapeNode, 'gradFill');
  const gradientStops = gradientFill
    ? getLocalDescendants(gradientFill, 'gs')
        .map((stopNode) => ({
          colorNode: Array.from(stopNode.childNodes).find(isElementNode) ?? null,
          position: Number(stopNode.getAttribute('pos') ?? Number.NaN),
        }))
        .filter((stop) => Boolean(stop.colorNode))
    : [];
  if (gradientStops.length === 0) return null;
  gradientStops.sort((left, right) => {
    const leftPos = Number.isFinite(left.position) ? left.position : 0;
    const rightPos = Number.isFinite(right.position) ? right.position : 0;
    return leftPos - rightPos;
  });
  const midpointStop =
    gradientStops.find((stop) => Number.isFinite(stop.position) && stop.position >= 50000) ??
    gradientStops[Math.floor(gradientStops.length / 2)] ??
    gradientStops[0];
  return resolveChartColorNode(midpointStop?.colorNode ?? null, themePalette);
}

function resolveChartLineStyle(
  shapeNode: Element | null,
  themePalette?: XlsxThemePalette | null,
): { color: string | null; hidden: boolean; widthPx: number | undefined } {
  const lineNode =
    shapeNode?.localName === 'ln' ? shapeNode : shapeNode ? getFirstLocalChild(shapeNode, 'ln') : null;
  if (!lineNode) return { color: null, hidden: false, widthPx: undefined };
  if (getFirstLocalChild(lineNode, 'noFill')) return { color: null, hidden: true, widthPx: undefined };
  const solidFill = getFirstLocalChild(lineNode, 'solidFill');
  const colorNode = solidFill
    ? findFirstChartColorElement(Array.from(solidFill.childNodes).find(isElementNode) ?? null)
    : null;
  const widthValue = Number(lineNode.getAttribute('w') ?? Number.NaN);
  return {
    color: resolveChartColorNode(colorNode, themePalette),
    hidden: false,
    widthPx: Number.isFinite(widthValue) ? Math.max(1, widthValue / EMU_PER_PIXEL) : undefined,
  };
}

function normalizeLegendPosition(position: string | undefined): string | undefined {
  if (!position) return undefined;
  switch (position) {
    case 'bottom':
      return 'b';
    case 'left':
      return 'l';
    case 'right':
      return 'r';
    case 'top':
      return 't';
    default:
      return position;
  }
}

function readChartNumericAttribute(parent: Element | null, localName: string): number | undefined {
  const node = parent ? getFirstLocalChild(parent, localName) : null;
  const value = Number(node?.getAttribute('val') ?? Number.NaN);
  return Number.isFinite(value) ? value : undefined;
}

function readChartBooleanAttribute(parent: Element | null, localName: string): boolean | undefined {
  const node = parent ? getFirstLocalChild(parent, localName) : null;
  if (!node) return undefined;
  const rawValue = node.getAttribute('val');
  if (rawValue == null) return true;
  if (rawValue === '1' || rawValue === 'true') return true;
  if (rawValue === '0' || rawValue === 'false') return false;
  return undefined;
}

function readChartLabelFontSizePt(textPropertiesNode: Element | null): number | undefined {
  if (!textPropertiesNode) return undefined;
  const runPropertiesNode =
    getFirstLocalDescendant(textPropertiesNode, 'defRPr') ??
    getFirstLocalDescendant(textPropertiesNode, 'rPr');
  const rawSize = Number(runPropertiesNode?.getAttribute('sz') ?? Number.NaN);
  if (!Number.isFinite(rawSize) || rawSize <= 0) return undefined;
  return rawSize / 100;
}

function parseChartPointDataLabelsFromXml(labelsNode: Element): XlsxChartPointDataLabel[] {
  const fallbackFontSizePt = readChartLabelFontSizePt(getFirstLocalChild(labelsNode, 'txPr'));
  const labels: XlsxChartPointDataLabel[] = [];
  for (const pointLabelNode of getLocalChildren(labelsNode, 'dLbl')) {
    const index = readChartNumericAttribute(pointLabelNode, 'idx');
    if (typeof index !== 'number' || !Number.isFinite(index)) continue;
    const layoutNode = getFirstLocalChild(pointLabelNode, 'layout');
    const manualLayoutNode = layoutNode ? getFirstLocalChild(layoutNode, 'manualLayout') : null;
    labels.push({
      deleted: readChartBooleanAttribute(pointLabelNode, 'delete'),
      fontSizePt: readChartLabelFontSizePt(getFirstLocalChild(pointLabelNode, 'txPr')) ?? fallbackFontSizePt,
      index,
      showBubbleSize: readChartBooleanAttribute(pointLabelNode, 'showBubbleSize'),
      showCategoryName: readChartBooleanAttribute(pointLabelNode, 'showCatName'),
      showPercent: readChartBooleanAttribute(pointLabelNode, 'showPercent'),
      showSeriesName: readChartBooleanAttribute(pointLabelNode, 'showSerName'),
      showValue: readChartBooleanAttribute(pointLabelNode, 'showVal'),
      x: readChartNumericAttribute(manualLayoutNode, 'x'),
      y: readChartNumericAttribute(manualLayoutNode, 'y'),
    });
  }
  return labels;
}

function parseChartDataLabelsFromXml(labelsNode: Element | null): XlsxChartDataLabels | null {
  if (!labelsNode) return null;
  const pointLabels = parseChartPointDataLabelsFromXml(labelsNode);
  const labels: XlsxChartDataLabels = {
    pointLabels: pointLabels.length > 0 ? pointLabels : undefined,
    raw: {},
    showBubbleSize: readChartBooleanAttribute(labelsNode, 'showBubbleSize'),
    showCategoryName: readChartBooleanAttribute(labelsNode, 'showCatName'),
    showLegendKey: readChartBooleanAttribute(labelsNode, 'showLegendKey'),
    showPercent: readChartBooleanAttribute(labelsNode, 'showPercent'),
    showSeriesName: readChartBooleanAttribute(labelsNode, 'showSerName'),
    showValue: readChartBooleanAttribute(labelsNode, 'showVal'),
  };
  const hasValue =
    labels.showBubbleSize !== undefined ||
    labels.showCategoryName !== undefined ||
    labels.showLegendKey !== undefined ||
    labels.showPercent !== undefined ||
    (labels.pointLabels?.length ?? 0) > 0 ||
    labels.showSeriesName !== undefined ||
    labels.showValue !== undefined;
  return hasValue ? labels : null;
}

function cellValueToNumber(value: string): number | null {
  const parsed = Number(value.replace(/,/g, ''));
  return Number.isFinite(parsed) ? parsed : null;
}

function parseChartCacheValues(
  parentNode: Element | null,
  cacheName: 'numCache' | 'strCache',
  mode: 'category' | 'value',
): Array<number | string | null> | null {
  if (!parentNode) return null;
  const referenceNode =
    getFirstLocalChild(parentNode, 'numRef') ??
    getFirstLocalChild(parentNode, 'strRef') ??
    parentNode;
  const cacheNode =
    getFirstLocalChild(referenceNode, cacheName) ??
    // Literal (unlinked) data lives in numLit/strLit with the same pt shape.
    getFirstLocalChild(referenceNode, cacheName === 'numCache' ? 'numLit' : 'strLit');
  if (!cacheNode) return null;
  const pointCount = readChartNumericAttribute(cacheNode, 'ptCount');
  const pointNodes = getLocalChildren(cacheNode, 'pt')
    .map((pointNode) => {
      const rawIndex = Number(pointNode.getAttribute('idx') ?? Number.NaN);
      return {
        index: Number.isFinite(rawIndex) ? rawIndex : 0,
        value: getFirstLocalChild(pointNode, 'v')?.textContent ?? '',
      };
    })
    .sort((left, right) => left.index - right.index);
  if (pointNodes.length === 0) return null;
  const maxIndex = pointNodes.reduce((max, point) => Math.max(max, point.index), 0);
  const targetLength = Math.max(
    pointNodes.length,
    Number.isFinite(pointCount ?? Number.NaN) ? Number(pointCount) : 0,
    maxIndex + 1,
  );
  const values = Array.from({ length: targetLength }, () => null as number | string | null);
  for (const point of pointNodes) {
    if (point.index < 0 || point.index >= values.length) continue;
    if (mode === 'value') {
      values[point.index] = cellValueToNumber(point.value);
    } else {
      values[point.index] = point.value.length > 0 ? point.value : null;
    }
  }
  return values;
}

function parseChartMultiLevelCacheValues(
  parentNode: Element | null,
  mode: 'category' | 'value',
): Array<number | string | null> | null {
  if (!parentNode) return null;
  const referenceNode = getFirstLocalChild(parentNode, 'multiLvlStrRef') ?? parentNode;
  const cacheNode = getFirstLocalChild(referenceNode, 'multiLvlStrCache');
  if (!cacheNode) return null;
  const levelNodes = getLocalChildren(cacheNode, 'lvl');
  if (levelNodes.length === 0) return null;
  const pointCount = readChartNumericAttribute(cacheNode, 'ptCount');
  const primaryLevelNode =
    mode === 'category' ? (levelNodes[levelNodes.length - 1] ?? levelNodes[0]) : levelNodes[0];
  if (!primaryLevelNode) return null;
  const pointNodes = getLocalChildren(primaryLevelNode, 'pt')
    .map((pointNode) => {
      const rawIndex = Number(pointNode.getAttribute('idx') ?? Number.NaN);
      return {
        index: Number.isFinite(rawIndex) ? rawIndex : 0,
        value: getFirstLocalChild(pointNode, 'v')?.textContent ?? '',
      };
    })
    .sort((left, right) => left.index - right.index);
  if (pointNodes.length === 0) return null;
  const maxIndex = pointNodes.reduce((max, point) => Math.max(max, point.index), 0);
  const targetLength = Math.max(
    pointNodes.length,
    Number.isFinite(pointCount ?? Number.NaN) ? Number(pointCount) : 0,
    maxIndex + 1,
  );
  const values = Array.from({ length: targetLength }, () => null as number | string | null);
  for (const point of pointNodes) {
    if (point.index < 0 || point.index >= values.length) continue;
    if (mode === 'value') {
      values[point.index] = cellValueToNumber(point.value);
      continue;
    }
    values[point.index] = point.value.length > 0 ? point.value : null;
  }
  return values;
}

function parseChartPointStyles(
  seriesNode: Element,
  themePalette?: XlsxThemePalette | null,
): XlsxChartPointStyle[] {
  const pointStyles: XlsxChartPointStyle[] = [];
  for (const dataPointNode of getLocalChildren(seriesNode, 'dPt')) {
    const indexValue = readChartNumericAttribute(dataPointNode, 'idx');
    if (indexValue === undefined) continue;
    const shapeProperties = getFirstLocalChild(dataPointNode, 'spPr');
    const lineStyle = resolveChartLineStyle(shapeProperties, themePalette);
    pointStyles.push({
      color: resolveChartFillColor(shapeProperties, themePalette) ?? undefined,
      explosion: readChartNumericAttribute(dataPointNode, 'explosion'),
      index: indexValue,
      lineColor: lineStyle.color ?? undefined,
    });
  }
  return pointStyles;
}

function parseInvertNegativeStyle(
  seriesNode: Element,
  themePalette?: XlsxThemePalette | null,
): { color: string | undefined; lineColor: string | undefined } {
  const invertNode = getFirstLocalDescendant(seriesNode, 'invertSolidFillFmt');
  const shapeProperties = invertNode ? getFirstLocalChild(invertNode, 'spPr') : null;
  if (!shapeProperties) return { color: undefined, lineColor: undefined };
  const lineStyle = resolveChartLineStyle(shapeProperties, themePalette);
  return {
    color: resolveChartFillColor(shapeProperties, themePalette) ?? undefined,
    lineColor: lineStyle.color ?? undefined,
  };
}

function readChartAxisFromXml(axisNode: Element | null): Partial<XlsxChartAxis> | null {
  if (!axisNode) return null;
  const numFmt = getFirstLocalChild(axisNode, 'numFmt');
  const scalingNode = getFirstLocalChild(axisNode, 'scaling');
  const axis: Partial<XlsxChartAxis> = {
    crossId: readChartNumericAttribute(axisNode, 'crossAx'),
    crosses: getFirstLocalChild(axisNode, 'crosses')?.getAttribute('val') ?? undefined,
    crossBetween: getFirstLocalChild(axisNode, 'crossBetween')?.getAttribute('val') ?? undefined,
    delete:
      getFirstLocalChild(axisNode, 'delete')?.getAttribute('val') === '1'
        ? true
        : getFirstLocalChild(axisNode, 'delete')?.getAttribute('val') === '0'
          ? false
          : undefined,
    id: readChartNumericAttribute(axisNode, 'axId'),
    labelPosition: getFirstLocalChild(axisNode, 'tickLblPos')?.getAttribute('val') ?? undefined,
    logBase: readChartNumericAttribute(scalingNode, 'logBase'),
    orientation: getFirstLocalChild(scalingNode ?? axisNode, 'orientation')?.getAttribute('val') ?? undefined,
    majorGridlines: Boolean(getFirstLocalChild(axisNode, 'majorGridlines')),
    majorTickMark: getFirstLocalChild(axisNode, 'majorTickMark')?.getAttribute('val') ?? undefined,
    majorUnit: readChartNumericAttribute(axisNode, 'majorUnit'),
    max: readChartNumericAttribute(scalingNode, 'max'),
    min: readChartNumericAttribute(scalingNode, 'min'),
    minorGridlines: Boolean(getFirstLocalChild(axisNode, 'minorGridlines')),
    minorTickMark: getFirstLocalChild(axisNode, 'minorTickMark')?.getAttribute('val') ?? undefined,
    minorUnit: readChartNumericAttribute(axisNode, 'minorUnit'),
    position: getFirstLocalChild(axisNode, 'axPos')?.getAttribute('val') ?? undefined,
    tickLabelSkip: readChartNumericAttribute(axisNode, 'tickLblSkip'),
    tickMarkSkip: readChartNumericAttribute(axisNode, 'tickMarkSkip'),
  };
  if (numFmt) {
    axis.numberFormat = {
      formatCode: numFmt.getAttribute('formatCode') ?? undefined,
      sourceLinked:
        numFmt.getAttribute('sourceLinked') === '1'
          ? true
          : numFmt.getAttribute('sourceLinked') === '0'
            ? false
            : undefined,
    };
  }
  return axis;
}

function mergeChartAxis(
  target: XlsxChartAxis | null | undefined,
  patch: Partial<XlsxChartAxis> | null | undefined,
): XlsxChartAxis | null {
  if (!patch) return target ?? null;
  return { ...(target ?? {}), ...patch };
}

function readChartWallFromXml(
  wallNode: Element | null,
  themePalette?: XlsxThemePalette | null,
): XlsxChartWall | null {
  if (!wallNode) return null;
  const shapeProperties = getFirstLocalChild(wallNode, 'spPr');
  const lineStyle = resolveChartLineStyle(shapeProperties, themePalette);
  return {
    fillColor: resolveChartFillColor(shapeProperties, themePalette) ?? undefined,
    hidden: shapeProperties ? getFirstLocalChild(shapeProperties, 'noFill') != null : undefined,
    lineColor: lineStyle.color ?? undefined,
    thickness: readChartNumericAttribute(wallNode, 'thickness'),
  };
}

function readChartColorPalette(
  colorsXml: string | undefined,
  themePalette?: XlsxThemePalette | null,
): string[] {
  if (!colorsXml) return [];
  const colorStyleDocument = parseXml(colorsXml);
  if (!colorStyleDocument?.documentElement) return [];
  return Array.from(colorStyleDocument.documentElement.childNodes)
    .filter((child): child is Element => isElementNode(child) && child.localName !== 'variation')
    .map(
      (child) =>
        resolveChartColorNode(child, themePalette) ??
        resolveChartColorNode(findFirstChartColorElement(child), themePalette),
    )
    .filter((color): color is string => typeof color === 'string');
}

function readChartStyleAppearance(
  styleXml: string | undefined,
  themePalette?: XlsxThemePalette | null,
): ChartStyleAppearance {
  if (!styleXml) return {};
  const styleDocument = parseXml(styleXml);
  if (!styleDocument) return {};
  const dataPointNode = getFirstLocalDescendant(styleDocument, 'dataPoint');
  const fillRefNode = dataPointNode ? getFirstLocalChild(dataPointNode, 'fillRef') : null;
  const index = Number(fillRefNode?.getAttribute('idx') ?? Number.NaN);
  const chartAreaNode = getFirstLocalDescendant(styleDocument, 'chartArea');
  const chartAreaShapeProperties = chartAreaNode ? getFirstLocalChild(chartAreaNode, 'spPr') : null;
  const chartAreaFontRef = chartAreaNode ? getFirstLocalChild(chartAreaNode, 'fontRef') : null;
  const chartAreaFontColor = chartAreaFontRef
    ? resolveChartColorNode(Array.from(chartAreaFontRef.childNodes).find(isElementNode) ?? null, themePalette)
    : null;
  const titleNode = getFirstLocalDescendant(styleDocument, 'title');
  const titleFontRef = titleNode ? getFirstLocalChild(titleNode, 'fontRef') : null;
  const titleColor = titleFontRef
    ? resolveChartColorNode(Array.from(titleFontRef.childNodes).find(isElementNode) ?? null, themePalette)
    : null;
  const axisStyleNode =
    getFirstLocalDescendant(styleDocument, 'categoryAxis') ??
    getFirstLocalDescendant(styleDocument, 'valueAxis');
  const axisShapeProperties = axisStyleNode ? getFirstLocalChild(axisStyleNode, 'spPr') : null;
  const axisFontRef = axisStyleNode ? getFirstLocalChild(axisStyleNode, 'fontRef') : null;
  const chartAreaNoFill = chartAreaShapeProperties
    ? getFirstLocalChild(chartAreaShapeProperties, 'noFill') != null
    : false;
  return {
    axisLabelColor: axisFontRef
      ? (resolveChartColorNode(Array.from(axisFontRef.childNodes).find(isElementNode) ?? null, themePalette) ??
        undefined)
      : undefined,
    axisLineColor: resolveChartLineStyle(axisShapeProperties, themePalette).color ?? undefined,
    chartAreaBorderColor: resolveChartLineStyle(chartAreaShapeProperties, themePalette).color ?? undefined,
    chartAreaFillColor: resolveChartFillColor(chartAreaShapeProperties, themePalette) ?? undefined,
    chartAreaNoFill,
    paletteOffset: Number.isFinite(index) ? index : undefined,
    textColor: chartAreaFontColor ?? undefined,
    titleColor: titleColor ?? chartAreaFontColor ?? undefined,
  };
}

function buildThemeSeriesPalette(themePalette?: XlsxThemePalette | null): string[] {
  const themeColors = [4, 5, 6, 7, 8, 9]
    .map((index) => themePalette?.colorsByIndex[index] ?? null)
    .filter((color): color is string => Boolean(color));
  return themeColors.length > 0 ? themeColors : SERIES_COLORS;
}

function normalizeBuiltinSurfaceStyleId(styleId: number | undefined): number | null {
  if (typeof styleId !== 'number' || !Number.isFinite(styleId)) return null;
  return styleId >= 100 ? styleId - 100 : styleId;
}

function getBuiltinSurfacePalette(
  styleId: number | undefined,
  wireframe: boolean | undefined,
): string[] | null {
  const normalized = normalizeBuiltinSurfaceStyleId(styleId);
  if (normalized === 34 || (wireframe === true && normalized == null)) {
    return ['#5b9bd5', '#ed7d31', '#a5a5a5'];
  }
  if (normalized === 35 || normalized === 36 || (wireframe !== true && normalized == null)) {
    return ['#2f5597', '#4472c4', '#5b9bd5', '#8faadc', '#d9e2f3'];
  }
  return null;
}

function applyBuiltinSurfaceDefaults(chart: XlsxChart): void {
  if (chart.chartType !== 'Surface') return;
  const builtinPalette = getBuiltinSurfacePalette(chart.chartStyleId, chart.wireframe);
  if ((!chart.chartColorPalette || chart.chartColorPalette.length === 0) && builtinPalette) {
    chart.chartColorPalette = builtinPalette;
  }
  const wallFill = chart.wireframe ? '#d0d0d0' : '#d9d9df';
  const wallLine = chart.wireframe ? '#a6a6a6' : '#a8adb7';
  chart.floor = {
    ...(chart.floor ?? {}),
    fillColor: chart.floor?.fillColor ?? wallFill,
    lineColor: chart.floor?.lineColor ?? wallLine,
  };
  chart.sideWall = {
    ...(chart.sideWall ?? {}),
    fillColor: chart.sideWall?.fillColor ?? wallFill,
    lineColor: chart.sideWall?.lineColor ?? wallLine,
  };
  chart.backWall = {
    ...(chart.backWall ?? {}),
    fillColor: chart.backWall?.fillColor ?? wallFill,
    lineColor: chart.backWall?.lineColor ?? wallLine,
  };
  if (!chart.surfaceMaterial && chart.wireframe !== true) {
    chart.surfaceMaterial = 'flat';
  }
}

export function applyBuiltinChartDefaults(
  chart: XlsxChart,
  themePalette?: XlsxThemePalette | null,
): void {
  const darkBuiltInStyle =
    typeof chart.chartStyleId === 'number' && chart.chartStyleId >= 140 && chart.chartStyleId < 150;
  const textColor = themePalette?.colorsByIndex[1] ?? themePalette?.colorsByIndex[3] ?? null;
  const minorTypeface = themePalette?.minorLatinFont?.trim() || undefined;
  const derivedAxisColor = textColor ? applyLightnessTransform(textColor, 0.35, 0.55) : null;
  const derivedBorderColor = textColor
    ? applyLightnessTransform(textColor, chart.is3d ? 0.28 : 0.22, chart.is3d ? 0.6 : 0.7)
    : null;
  if (darkBuiltInStyle) {
    chart.chartAreaFillColor = chart.chartAreaFillColor ?? '#1f1f1f';
    chart.chartAreaBorderColor = chart.chartAreaBorderColor ?? '#1f1f1f';
    chart.textColor = chart.textColor ?? '#f5f5f5';
    chart.titleColor = chart.titleColor ?? '#f5f5f5';
    chart.axisLabelColor = chart.axisLabelColor ?? '#d9d9d9';
    chart.axisLineColor = chart.axisLineColor ?? '#8c8c8c';
  }
  chart.chartAreaBorderColor = chart.chartAreaBorderColor ?? derivedBorderColor ?? undefined;
  chart.textColor = chart.textColor ?? textColor ?? undefined;
  chart.titleColor = chart.titleColor ?? textColor ?? undefined;
  chart.axisLabelColor = chart.axisLabelColor ?? derivedAxisColor ?? textColor ?? undefined;
  chart.axisLineColor = chart.axisLineColor ?? derivedAxisColor ?? textColor ?? undefined;
  chart.fontFamily = chart.fontFamily ?? minorTypeface;
  chart.titleFontFamily = chart.titleFontFamily ?? chart.fontFamily ?? minorTypeface;

  const seriesPalette =
    chart.chartColorPalette && chart.chartColorPalette.length > 0
      ? chart.chartColorPalette
      : buildThemeSeriesPalette(themePalette);
  if (!chart.chartColorPalette || chart.chartColorPalette.length === 0) {
    chart.chartColorPalette = seriesPalette;
  }

  chart.series = chart.series.map((series, index) => {
    const fallbackColor = seriesPalette[index % seriesPalette.length];
    return {
      ...series,
      color: series.color ?? series.lineColor ?? fallbackColor,
      lineColor: series.lineColor ?? series.color ?? fallbackColor,
      markerColor: series.markerColor ?? series.color ?? series.lineColor ?? fallbackColor,
      markerLineColor: series.markerLineColor ?? series.lineColor ?? series.color ?? fallbackColor,
    };
  });
  chart.typeGroups = chart.typeGroups?.map((group, groupIndex) => ({
    ...group,
    series: group.series.map((series, seriesIndex) => {
      const fallbackColor = seriesPalette[(groupIndex + seriesIndex) % seriesPalette.length];
      return {
        ...series,
        color: series.color ?? series.lineColor ?? fallbackColor,
        lineColor: series.lineColor ?? series.color ?? fallbackColor,
        markerColor: series.markerColor ?? series.color ?? series.lineColor ?? fallbackColor,
        markerLineColor: series.markerLineColor ?? series.lineColor ?? series.color ?? fallbackColor,
      };
    }),
  }));
  applyBuiltinSurfaceDefaults(chart);
}

function resolveScatterChartType(scatterStyle: string | null | undefined): string {
  switch (scatterStyle) {
    case 'line':
    case 'lineMarker':
      return 'ScatterLines';
    case 'smooth':
    case 'smoothMarker':
      return 'ScatterSmooth';
    default:
      return 'Scatter';
  }
}

function findPrimaryChartTypeNode(plotAreaNode: Element | null): Element | null {
  if (!plotAreaNode) return null;
  for (const localName of PRIMARY_CHART_TYPE_LOCAL_NAMES) {
    const node = getLocalChildren(plotAreaNode, localName)[0];
    if (node) return node;
  }
  return null;
}

/** Maps a classic chart-type node to the normalized react-xlsx chart type. */
function resolveClassicGroupChartType(chartTypeNode: Element): {
  chartType: string;
  is3d: boolean | undefined;
} {
  const grouping = getFirstLocalChild(chartTypeNode, 'grouping')?.getAttribute('val');
  switch (chartTypeNode.localName) {
    case 'barChart':
    case 'bar3DChart': {
      const isHorizontalBar = getFirstLocalChild(chartTypeNode, 'barDir')?.getAttribute('val') === 'bar';
      const is3d = chartTypeNode.localName === 'bar3DChart' ? true : undefined;
      if (grouping === 'percentStacked') {
        return { chartType: isHorizontalBar ? 'BarPercentStacked' : 'ColumnPercentStacked', is3d };
      }
      if (grouping === 'stacked') {
        return { chartType: isHorizontalBar ? 'BarStacked' : 'ColumnStacked', is3d };
      }
      return { chartType: isHorizontalBar ? 'BarClustered' : 'ColumnClustered', is3d };
    }
    case 'areaChart':
    case 'area3DChart': {
      const is3d = chartTypeNode.localName === 'area3DChart' ? true : undefined;
      if (grouping === 'stacked') return { chartType: 'AreaStacked', is3d };
      if (grouping === 'percentStacked') return { chartType: 'AreaPercentStacked', is3d };
      return { chartType: 'Area', is3d };
    }
    case 'lineChart':
    case 'line3DChart': {
      const is3d = chartTypeNode.localName === 'line3DChart' ? true : undefined;
      if (grouping === 'stacked') return { chartType: 'LineStacked', is3d };
      if (grouping === 'percentStacked') return { chartType: 'LinePercentStacked', is3d };
      return { chartType: 'Line', is3d };
    }
    case 'pieChart':
      return { chartType: 'Pie', is3d: undefined };
    case 'pie3DChart':
      return { chartType: 'Pie3D', is3d: true };
    case 'doughnutChart':
      return { chartType: 'Doughnut', is3d: undefined };
    case 'ofPieChart':
      return { chartType: 'BarOfPie', is3d: undefined };
    case 'scatterChart':
      return {
        chartType: resolveScatterChartType(
          getFirstLocalChild(chartTypeNode, 'scatterStyle')?.getAttribute('val'),
        ),
        is3d: undefined,
      };
    case 'radarChart':
      return { chartType: 'Radar', is3d: undefined };
    case 'surfaceChart':
      return { chartType: 'Surface', is3d: false };
    case 'surface3DChart':
      return { chartType: 'Surface', is3d: true };
    case 'stockChart':
      return { chartType: 'Stock', is3d: undefined };
    case 'bubbleChart':
      return { chartType: 'Bubble', is3d: undefined };
    default:
      return { chartType: 'ColumnClustered', is3d: undefined };
  }
}

function isScatterLikeChartType(chartType: string): boolean {
  return (
    chartType === 'Scatter' ||
    chartType === 'ScatterLines' ||
    chartType === 'ScatterSmooth' ||
    chartType === 'Bubble'
  );
}

function readSeriesName(seriesNode: Element): string | undefined {
  const textNode = getFirstLocalChild(seriesNode, 'tx');
  if (!textNode) return undefined;
  const cached = parseChartCacheValues(textNode, 'strCache', 'category');
  const cachedName = cached?.find((value) => typeof value === 'string' && value.length > 0);
  if (typeof cachedName === 'string') return cachedName;
  const literal = getFirstLocalChild(textNode, 'v')?.textContent ?? '';
  return literal.length > 0 ? literal : undefined;
}

/** Fully parses one classic `c:ser` node into the normalized series model. */
function parseClassicSeries(
  seriesNode: Element,
  chartType: string,
  seriesId: string,
  themePalette?: XlsxThemePalette | null,
): XlsxChartSeries {
  const isScatterChart = isScatterLikeChartType(chartType);
  const categories = isScatterChart
    ? (parseChartCacheValues(getFirstLocalChild(seriesNode, 'xVal'), 'numCache', 'value') ??
      parseChartMultiLevelCacheValues(getFirstLocalChild(seriesNode, 'xVal'), 'category'))
    : (parseChartCacheValues(getFirstLocalChild(seriesNode, 'cat'), 'strCache', 'category') ??
      parseChartCacheValues(getFirstLocalChild(seriesNode, 'cat'), 'numCache', 'category') ??
      parseChartMultiLevelCacheValues(getFirstLocalChild(seriesNode, 'cat'), 'category'));
  const values = isScatterChart
    ? parseChartCacheValues(getFirstLocalChild(seriesNode, 'yVal'), 'numCache', 'value')
    : parseChartCacheValues(getFirstLocalChild(seriesNode, 'val'), 'numCache', 'value');
  const bubbleSizes =
    chartType === 'Bubble'
      ? parseChartCacheValues(getFirstLocalChild(seriesNode, 'bubbleSize'), 'numCache', 'value')
      : null;
  const shapeProperties = getFirstLocalChild(seriesNode, 'spPr');
  const markerNode = getFirstLocalChild(seriesNode, 'marker');
  const markerShapeProperties = markerNode ? getFirstLocalChild(markerNode, 'spPr') : null;
  const lineStyle = resolveChartLineStyle(shapeProperties, themePalette);
  const markerLineStyle = resolveChartLineStyle(markerShapeProperties, themePalette);
  const fillColor = resolveChartFillColor(shapeProperties, themePalette);
  const markerSize = readChartNumericAttribute(markerNode, 'size');
  const markerSymbol = markerNode
    ? (getFirstLocalChild(markerNode, 'symbol')?.getAttribute('val') ?? undefined)
    : undefined;
  const pointStyles = parseChartPointStyles(seriesNode, themePalette);
  const invertNegativeStyle = parseInvertNegativeStyle(seriesNode, themePalette);
  const seriesExplosion = readChartNumericAttribute(seriesNode, 'explosion');
  const resolvedLineColor = lineStyle.hidden
    ? undefined
    : (lineStyle.color ?? fillColor ?? undefined);
  return {
    bubbleSizeRef: null,
    bubbleSizes: bubbleSizes
      ? bubbleSizes.map((value) => (typeof value === 'number' && Number.isFinite(value) ? value : null))
      : [],
    categories: categories ?? [],
    categoriesRef: null,
    color: fillColor ?? lineStyle.color ?? undefined,
    dataPoints: [],
    dataPointStyles: pointStyles.length > 0 ? pointStyles : undefined,
    id: seriesId,
    invertIfNegative: readChartBooleanAttribute(seriesNode, 'invertIfNegative'),
    lineColor: resolvedLineColor,
    lineWidthPx: lineStyle.hidden ? undefined : lineStyle.widthPx,
    marker: undefined,
    markerColor:
      resolveChartFillColor(markerShapeProperties, themePalette) ??
      fillColor ??
      lineStyle.color ??
      undefined,
    markerLineColor: markerLineStyle.color ?? lineStyle.color ?? fillColor ?? undefined,
    markerSize,
    markerSymbol,
    name: readSeriesName(seriesNode),
    negativeColor: invertNegativeStyle.color,
    negativeLineColor: invertNegativeStyle.lineColor,
    raw: {},
    shapeProperties: {
      xmlExplosion: seriesExplosion ?? undefined,
      xmlFillColor: fillColor ?? undefined,
      xmlLineHidden: lineStyle.hidden ? true : undefined,
      xmlLineColor: lineStyle.color ?? undefined,
      xmlLineWidthPx: lineStyle.widthPx ?? undefined,
      xmlNegativeFillColor: invertNegativeStyle.color ?? undefined,
      xmlNegativeLineColor: invertNegativeStyle.lineColor ?? undefined,
    },
    smooth: readChartBooleanAttribute(seriesNode, 'smooth'),
    values: values
      ? values.map((value) => (typeof value === 'number' && Number.isFinite(value) ? value : null))
      : [],
    valuesRef: null,
  };
}

function readGroupAxisIds(chartTypeNode: Element): number[] {
  return getLocalChildren(chartTypeNode, 'axId')
    .map((node) => Number(node.getAttribute('val') ?? Number.NaN))
    .filter((value) => Number.isFinite(value));
}

function readChartTitle(chartNode: Element | null): string | undefined {
  const titleNode = chartNode ? getFirstLocalChild(chartNode, 'title') : null;
  if (!titleNode) return undefined;
  const texts = getLocalDescendants(titleNode, 't')
    .map((node) => node.textContent ?? '')
    .filter((value) => value.trim().length > 0);
  const joined = texts.join('');
  return joined.length > 0 ? joined : undefined;
}

function defaultAnchor(): XlsxChart['anchor'] {
  return {
    kind: 'two-cell',
    from: { col: 0, colOffsetEmu: 0, row: 0, rowOffsetEmu: 0 },
    to: { col: 8, colOffsetEmu: 0, row: 15, rowOffsetEmu: 0 },
  };
}

function createEmptyChart(id: string): XlsxChart {
  return {
    anchor: defaultAnchor(),
    axes: [],
    categoryAxis: null,
    chartType: 'ColumnClustered',
    dataLabels: null,
    editable: false,
    id,
    legend: null,
    seriesAxis: null,
    series: [],
    sheetIndex: 0,
    sideWall: null,
    backWall: null,
    floor: null,
    typeGroups: [],
    valueAxis: null,
    workbookSheetIndex: 0,
    zIndex: 0,
  };
}

/** Builds the base chart model from a classic `c:chartSpace` document. */
function buildClassicChartBase(id: string, chartDocument: XMLDocument, themePalette?: XlsxThemePalette | null): XlsxChart {
  const chart = createEmptyChart(id);
  const chartNode = getFirstLocalDescendant(chartDocument, 'chart');
  const plotAreaNode = chartNode ? getFirstLocalChild(chartNode, 'plotArea') : null;
  chart.title = readChartTitle(chartNode);
  chart.autoTitleDeleted = readChartBooleanAttribute(chartNode, 'autoTitleDeleted');
  chart.name = chart.title;

  const groupNodes = plotAreaNode
    ? Array.from(plotAreaNode.childNodes)
        .filter(isElementNode)
        .filter((node) =>
          (PRIMARY_CHART_TYPE_LOCAL_NAMES as readonly string[]).includes(node.localName ?? ''),
        )
    : [];
  const typeGroups: XlsxChartTypeGroup[] = [];
  let seriesCounter = 0;
  for (const [groupIndex, groupNode] of groupNodes.entries()) {
    const resolved = resolveClassicGroupChartType(groupNode);
    const groupSeries = getLocalChildren(groupNode, 'ser').map((seriesNode) => {
      const parsed = parseClassicSeries(
        seriesNode,
        resolved.chartType,
        `${id}-series-${seriesCounter}`,
        themePalette,
      );
      seriesCounter += 1;
      return parsed;
    });
    typeGroups.push({
      axisIds: readGroupAxisIds(groupNode),
      chartType: resolved.chartType,
      dataLabels: parseChartDataLabelsFromXml(getFirstLocalChild(groupNode, 'dLbls')),
      gapWidth: readChartNumericAttribute(groupNode, 'gapWidth'),
      is3d: resolved.is3d,
      overlap: readChartNumericAttribute(groupNode, 'overlap'),
      raw: { groupIndex, xmlChartType: groupNode.localName ?? undefined },
      series: groupSeries,
      varyColors: readChartBooleanAttribute(groupNode, 'varyColors'),
    });
  }
  chart.series = typeGroups.flatMap((group) => group.series);
  chart.typeGroups = typeGroups.length > 1 ? typeGroups : [];
  const primaryGroup = typeGroups[0];
  if (primaryGroup) {
    chart.chartType = primaryGroup.chartType;
    chart.is3d = primaryGroup.is3d;
    chart.gapWidth = primaryGroup.gapWidth;
    chart.overlap = primaryGroup.overlap;
    chart.varyColors = primaryGroup.varyColors;
  }

  // Populate the axis list with ids so combo groups can resolve their axes.
  if (plotAreaNode) {
    const categoryAxisNodes = [
      ...getLocalChildren(plotAreaNode, 'catAx'),
      ...getLocalChildren(plotAreaNode, 'dateAx'),
    ];
    const valueAxisNodes = getLocalChildren(plotAreaNode, 'valAx');
    const seriesAxisNodes = getLocalChildren(plotAreaNode, 'serAx');
    const axes = [...categoryAxisNodes, ...valueAxisNodes, ...seriesAxisNodes]
      .map((axisNode) => readChartAxisFromXml(axisNode))
      .filter((axis): axis is XlsxChartAxis => axis !== null);
    chart.axes = axes;
  }
  return chart;
}

// ---------------------------------------------------------------------------
// Modern chartEx (`cx:chartSpace`) support: funnel, waterfall, treemap,
// sunburst, histogram/pareto, box & whisker, and region maps.
// ---------------------------------------------------------------------------

function humanizeChartExLayoutLabel(layout: string | undefined): string | undefined {
  if (!layout) return undefined;
  return layout
    .replace(/([a-z0-9])([A-Z])/g, '$1 $2')
    .replace(/[-_]+/g, ' ')
    .trim()
    .replace(/\b\w/g, (match) => match.toUpperCase());
}

function resolveChartExLayoutChartType(layout: string | undefined): string {
  switch (layout) {
    case 'boxWhisker':
      return 'BoxWhisker';
    case 'clusteredColumn':
      return 'ColumnClustered';
    case 'funnel':
      return 'Funnel';
    case 'paretoLine':
      return 'Line';
    case 'regionMap':
      return 'RegionMap';
    case 'sunburst':
      return 'Sunburst';
    case 'treemap':
      return 'Treemap';
    case 'waterfall':
      return 'Waterfall';
    default:
      return layout ? `Unsupported(cx:${layout})` : 'ColumnClustered';
  }
}

type ChartExDimension = {
  dimType: string;
  levels: Array<Array<string | null>>;
};

/** Reads a `cx:strDim`/`cx:numDim` level list; points carry text directly. */
function readChartExDimension(dimensionNode: Element): ChartExDimension {
  const levels = getLocalChildren(dimensionNode, 'lvl').map((levelNode) => {
    const pointCount = Number(levelNode.getAttribute('ptCount') ?? Number.NaN);
    const pointNodes = getLocalChildren(levelNode, 'pt')
      .map((pointNode) => ({
        index: Number(pointNode.getAttribute('idx') ?? Number.NaN),
        value: pointNode.textContent ?? '',
      }))
      .filter((point) => Number.isFinite(point.index) && point.index >= 0);
    const maxIndex = pointNodes.reduce((max, point) => Math.max(max, point.index), -1);
    const length = Math.max(Number.isFinite(pointCount) ? pointCount : 0, maxIndex + 1);
    const values = Array.from({ length }, () => null as string | null);
    for (const point of pointNodes) {
      if (point.index < values.length) values[point.index] = point.value;
    }
    return values;
  });
  return {
    dimType: dimensionNode.getAttribute('type') ?? (dimensionNode.localName === 'numDim' ? 'val' : 'cat'),
    levels,
  };
}

type ChartExHistogramBin = {
  count: number;
  label: string;
  lower: number;
  upper: number;
};

function niceHistogramStep(value: number): number {
  if (!Number.isFinite(value) || value <= 0) return 1;
  const exponent = Math.floor(Math.log10(value));
  const scale = 10 ** exponent;
  const normalized = value / scale;
  if (normalized <= 1) return scale;
  if (normalized <= 2) return scale * 2;
  if (normalized <= 5) return scale * 5;
  return scale * 10;
}

function formatHistogramBinLabel(lower: number, upper: number, index: number, closedRight: boolean): string {
  const leftBracket = closedRight ? (index === 0 ? '[' : '(') : '[';
  const rightBracket = closedRight ? ']' : ')';
  return `${leftBracket}${Number(lower.toFixed(6))},${Number(upper.toFixed(6))}${rightBracket}`;
}

function buildChartExHistogramBins(
  values: number[],
  binning: Record<string, unknown> | null,
  sortByFrequency: boolean,
): ChartExHistogramBin[] {
  if (values.length === 0) return [];
  const minValue = Math.min(...values);
  const maxValue = Math.max(...values);
  const explicitWidth =
    typeof binning?.binWidth === 'number' && Number.isFinite(binning.binWidth) && binning.binWidth > 0
      ? binning.binWidth
      : typeof binning?.width === 'number' && Number.isFinite(binning.width) && binning.width > 0
        ? binning.width
        : undefined;
  const explicitCount =
    typeof binning?.binCount === 'number' && Number.isFinite(binning.binCount) && binning.binCount > 0
      ? binning.binCount
      : typeof binning?.count === 'number' && Number.isFinite(binning.count) && binning.count > 0
        ? binning.count
        : undefined;
  const closedRight = binning?.intervalClosed === 'r' || binning?.intervalClosed === 'right';
  const mean = values.reduce((sum, value) => sum + value, 0) / values.length;
  const variance =
    values.reduce((sum, value) => sum + (value - mean) ** 2, 0) / Math.max(1, values.length);
  const standardDeviation = Math.sqrt(Math.max(0, variance));
  const allIntegers = values.every((value) => Math.abs(value - Math.round(value)) < 1e-9);
  const scottWidth =
    standardDeviation > 0 ? (3.49 * standardDeviation) / Math.cbrt(values.length) : undefined;
  const fallbackWidth =
    explicitCount != null
      ? (maxValue - minValue) / Math.max(1, explicitCount)
      : (scottWidth ?? (maxValue - minValue) / Math.max(1, Math.ceil(Math.log2(values.length) + 1)));
  const roughWidth =
    explicitWidth ??
    (allIntegers
      ? Math.max(1, Math.ceil(Math.max(fallbackWidth, 1e-6)))
      : niceHistogramStep(Math.max(fallbackWidth, 1e-6)));
  const binWidth = Math.max(roughWidth, 1e-6);
  const start =
    explicitWidth != null || explicitCount != null ? Math.floor(minValue / binWidth) * binWidth : minValue;
  const end = Math.max(start + binWidth, start + Math.ceil((maxValue - start) / binWidth) * binWidth);
  const binCount = Math.max(1, Math.ceil((end - start) / binWidth));
  const bins = Array.from({ length: binCount }, (_, index) => {
    const lower = start + binWidth * index;
    const upper = lower + binWidth;
    return {
      count: 0,
      label: formatHistogramBinLabel(lower, upper, index, closedRight),
      lower,
      upper,
    } satisfies ChartExHistogramBin;
  });
  values.forEach((value) => {
    if (!Number.isFinite(value)) return;
    const offset = (value - start) / binWidth;
    let binIndex = Math.floor(offset);
    if (closedRight && Math.abs(offset - Math.round(offset)) < 1e-9 && value > start) {
      binIndex -= 1;
    }
    if (value >= end) binIndex = bins.length - 1;
    if (value <= start) binIndex = 0;
    const target = bins[Math.max(0, Math.min(bins.length - 1, binIndex))];
    if (target) target.count += 1;
  });
  if (sortByFrequency) {
    bins.sort((left, right) => right.count - left.count || left.lower - right.lower);
  }
  return bins;
}

function buildChartExHistogramSeries(
  series: XlsxChartSeries,
  binning: Record<string, unknown> | null,
  sortByFrequency: boolean,
): XlsxChartSeries {
  if (!binning) return series;
  const numericValues = series.values.filter(
    (value): value is number => typeof value === 'number' && Number.isFinite(value),
  );
  if (numericValues.length === 0) return series;
  const bins = buildChartExHistogramBins(numericValues, binning, sortByFrequency);
  if (bins.length === 0) return series;
  return {
    ...series,
    categories: bins.map((bin) => bin.label),
    categoriesRef: null,
    raw: {
      ...series.raw,
      chartExHistogramBins: bins,
      chartExSourceValues: numericValues,
    },
    values: bins.map((bin) => bin.count),
  };
}

function buildChartExParetoLineSeries(
  series: XlsxChartSeries,
  name: string | undefined,
): XlsxChartSeries {
  const counts = series.values.map((value) =>
    typeof value === 'number' && Number.isFinite(value) ? value : 0,
  );
  const total = counts.reduce((sum, value) => sum + value, 0);
  let running = 0;
  const cumulative = counts.map((value) => {
    running += value;
    return total > 0 ? (running / total) * 100 : 0;
  });
  return {
    ...series,
    color: undefined,
    lineColor: undefined,
    markerColor: undefined,
    markerLineColor: undefined,
    markerSize: 7,
    markerSymbol: 'circle',
    name: name ?? 'Pareto',
    raw: {
      ...(series.raw ?? {}),
      chartExLayout: 'paretoLine',
    },
    values: cumulative,
  };
}

function collapseChartExPointSeries(chartType: string, series: XlsxChartSeries[]): XlsxChartSeries[] {
  if (chartType !== 'Funnel' && chartType !== 'Waterfall') {
    if (
      (chartType === 'Sunburst' || chartType === 'Treemap') &&
      series.length > 1 &&
      series.every((entry) => {
        const raw = entry.raw && typeof entry.raw === 'object' ? (entry.raw as Record<string, unknown>) : null;
        return raw?.dimType === 'size';
      })
    ) {
      const primarySeries = series.find((entry) => entry.hidden !== true) ?? series[0] ?? null;
      if (!primarySeries) return series;
      return [{ ...primarySeries, dataPoints: [], hidden: false }];
    }
    return series;
  }
  const primarySeries = series.find((entry) => entry.hidden !== true) ?? series[0] ?? null;
  if (!primarySeries) return series;
  // Unlike the react-xlsx wasm model (one series per funnel/waterfall point),
  // the XML builder produces a single series whose cached categories are the
  // stage labels, so they must be preserved.
  return [{ ...primarySeries, dataPoints: [], hidden: false }];
}

function readChartExBinning(seriesNode: Element): Record<string, unknown> | null {
  const layoutPrNode = getFirstLocalChild(seriesNode, 'layoutPr');
  const binningNode = layoutPrNode ? getFirstLocalChild(layoutPrNode, 'binning') : null;
  if (!binningNode) return null;
  const binning: Record<string, unknown> = {};
  for (const attribute of Array.from(binningNode.attributes)) {
    const rawValue = attribute.value;
    const numeric = Number(rawValue);
    binning[attribute.localName || attribute.name] =
      Number.isFinite(numeric) && rawValue.trim() !== '' ? numeric : rawValue;
  }
  for (const child of Array.from(binningNode.childNodes).filter(isElementNode)) {
    const rawValue = child.getAttribute('val');
    if (rawValue == null) continue;
    const numeric = Number(rawValue);
    binning[child.localName ?? ''] = Number.isFinite(numeric) && rawValue.trim() !== '' ? numeric : rawValue;
  }
  return binning;
}

function readChartExAxis(axisNode: Element): XlsxChartAxis {
  const numFmt = getFirstLocalChild(axisNode, 'numFmt');
  const valScaling = getFirstLocalChild(axisNode, 'valScaling');
  const axis: XlsxChartAxis = {
    delete: axisNode.getAttribute('hidden') === '1' ? true : undefined,
    id: Number.isFinite(Number(axisNode.getAttribute('id'))) ? Number(axisNode.getAttribute('id')) : undefined,
    majorGridlines: getFirstLocalChild(axisNode, 'majorGridlines') ? true : undefined,
    max: Number.isFinite(Number(valScaling?.getAttribute('max'))) ? Number(valScaling?.getAttribute('max')) : undefined,
    min: Number.isFinite(Number(valScaling?.getAttribute('min'))) ? Number(valScaling?.getAttribute('min')) : undefined,
    minorGridlines: getFirstLocalChild(axisNode, 'minorGridlines') ? true : undefined,
    raw: {},
  };
  if (numFmt) {
    axis.numberFormat = {
      formatCode: numFmt.getAttribute('formatCode') ?? undefined,
      sourceLinked:
        numFmt.getAttribute('sourceLinked') === '1'
          ? true
          : numFmt.getAttribute('sourceLinked') === '0'
            ? false
            : undefined,
    };
  }
  return axis;
}

/** Builds the base chart model from a modern `cx:chartSpace` document. */
function buildChartExBase(id: string, chartDocument: XMLDocument, themePalette?: XlsxThemePalette | null): XlsxChart {
  const chart = createEmptyChart(id);
  const chartNode = getFirstLocalDescendant(chartDocument, 'chart');
  const chartDataNode = getFirstLocalDescendant(chartDocument, 'chartData');
  const plotAreaNode = chartNode ? getFirstLocalChild(chartNode, 'plotArea') : null;
  const plotAreaRegion = plotAreaNode ? getFirstLocalChild(plotAreaNode, 'plotAreaRegion') : null;

  const dataById = new Map<number, ChartExDimension[]>();
  if (chartDataNode) {
    for (const dataNode of getLocalChildren(chartDataNode, 'data')) {
      const dataId = Number(dataNode.getAttribute('id') ?? Number.NaN);
      if (!Number.isFinite(dataId)) continue;
      const dimensions = [
        ...getLocalChildren(dataNode, 'strDim'),
        ...getLocalChildren(dataNode, 'numDim'),
      ].map(readChartExDimension);
      dataById.set(dataId, dimensions);
    }
  }

  const seriesNodes = plotAreaRegion ? getLocalChildren(plotAreaRegion, 'series') : [];
  const seriesLayouts = seriesNodes.map((node) => node.getAttribute('layoutId') ?? undefined);
  const primaryLayout = seriesLayouts.find((value): value is string => Boolean(value && value.length > 0));
  const chartType = resolveChartExLayoutChartType(primaryLayout);
  const fallbackTitle = humanizeChartExLayoutLabel(primaryLayout);
  const titleNode = chartNode ? getFirstLocalChild(chartNode, 'title') : null;
  const titleText = titleNode
    ? getLocalDescendants(titleNode, 'v')
        .map((node) => node.textContent ?? '')
        .join('')
        .trim()
    : '';
  const chartTitle = titleText.length > 0 ? titleText : titleNode ? 'Chart Title' : fallbackTitle;

  const normalizedSeries = seriesNodes.map((seriesNode, seriesIndex): XlsxChartSeries => {
    const dataId = readChartNumericAttribute(seriesNode, 'dataId');
    const dimensions = dataId !== undefined ? (dataById.get(dataId) ?? []) : [];
    const categoryDimension =
      dimensions.find((dimension) => dimension.dimType === 'cat') ??
      dimensions.find((dimension) => dimension.dimType === 'name') ??
      null;
    const valueDimension =
      dimensions.find(
        (dimension) =>
          dimension.dimType === 'val' ||
          dimension.dimType === 'y' ||
          dimension.dimType === 'colorVal' ||
          dimension.dimType === 'size',
      ) ??
      dimensions.find((dimension) => dimension !== categoryDimension) ??
      categoryDimension;
    const colorStringDimension = dimensions.find((dimension) => dimension.dimType === 'colorStr') ?? null;
    // Levels arrive innermost-first; the leaf level provides the display label.
    const categoryLevels = categoryDimension?.levels ?? [];
    const leafCategories = categoryLevels[0] ?? [];
    const pointCount = Math.max(
      leafCategories.length,
      valueDimension?.levels[0]?.length ?? 0,
    );
    const hierarchyCategories: string[][] = Array.from({ length: pointCount }, (_, pointIndex) => {
      const path: string[] = [];
      for (let level = categoryLevels.length - 1; level >= 0; level -= 1) {
        const label = categoryLevels[level]?.[pointIndex];
        if (typeof label === 'string' && label.length > 0) path.push(label);
      }
      return path;
    });
    const values = (valueDimension?.levels[0] ?? []).map((value) => {
      if (value == null) return null;
      return cellValueToNumber(value);
    });
    const colorStrings = (colorStringDimension?.levels[0] ?? []).map((value) => {
      if (typeof value !== 'string') return null;
      const trimmed = value.trim();
      return trimmed.length > 0 ? trimmed : null;
    });
    const shapeProperties = getFirstLocalChild(seriesNode, 'spPr');
    const fillColor = resolveChartFillColor(shapeProperties, themePalette);
    const lineStyle = resolveChartLineStyle(shapeProperties, themePalette);
    const textNode = getFirstLocalChild(seriesNode, 'tx');
    const nameText = textNode
      ? getLocalDescendants(textNode, 'v')
          .map((node) => node.textContent ?? '')
          .join('')
          .trim()
      : '';
    const layoutPrNode = getFirstLocalChild(seriesNode, 'layoutPr');
    const geographyNode = layoutPrNode ? getFirstLocalChild(layoutPrNode, 'geography') : null;
    const geography = geographyNode
      ? Object.fromEntries(
          Array.from(geographyNode.attributes).map((attribute) => [
            attribute.localName || attribute.name,
            attribute.value,
          ]),
        )
      : null;
    const valueColorsNode = getFirstLocalChild(seriesNode, 'valueColors');
    const valueColors = valueColorsNode
      ? Array.from(valueColorsNode.childNodes)
          .filter(isElementNode)
          .map((node) => resolveChartColorNode(findFirstChartColorElement(node) ?? node, themePalette))
          .filter((value): value is string => typeof value === 'string' && value.length > 0)
      : [];
    return {
      bubbleSizeRef: null,
      bubbleSizes: [],
      categories: leafCategories.map((value) => (value != null && value.length > 0 ? value : null)),
      categoriesRef: null,
      color: fillColor ?? undefined,
      dataPoints: [],
      dataPointStyles: undefined,
      hidden: seriesNode.getAttribute('hidden') === '1' ? true : undefined,
      id: `${id}-series-${seriesIndex}`,
      lineColor: lineStyle.color ?? fillColor ?? undefined,
      lineWidthPx: lineStyle.widthPx,
      markerColor: fillColor ?? undefined,
      markerLineColor: lineStyle.color ?? fillColor ?? undefined,
      name: nameText.length > 0 ? nameText : undefined,
      raw: {
        chartExColorStrings: colorStrings,
        chartExHierarchyCategories: hierarchyCategories,
        dimType: valueDimension?.dimType,
        layoutId: seriesNode.getAttribute('layoutId') ?? undefined,
        layoutPr: geography ? { geography } : undefined,
        valueColors: valueColors.length > 0 ? valueColors : undefined,
      },
      values,
      valuesRef: null,
    };
  });

  const clusteredColumnIndex = seriesLayouts.findIndex((layout) => layout === 'clusteredColumn');
  const hasParetoLine = seriesLayouts.includes('paretoLine');
  const paretoLineSeriesIndex = seriesLayouts.findIndex((layout) => layout === 'paretoLine');
  const clusteredNode = clusteredColumnIndex >= 0 ? (seriesNodes[clusteredColumnIndex] ?? null) : null;
  const binning = clusteredNode ? readChartExBinning(clusteredNode) : null;
  const clusteredBase =
    clusteredColumnIndex >= 0
      ? (normalizedSeries[clusteredColumnIndex] ?? normalizedSeries[0] ?? null)
      : null;
  const primaryHistogramSeries =
    clusteredBase && binning ? buildChartExHistogramSeries(clusteredBase, binning, hasParetoLine) : null;
  const synthesizedParetoSeries =
    hasParetoLine && primaryHistogramSeries && primaryHistogramSeries.values.length > 0
      ? buildChartExParetoLineSeriesFromNodes(
          primaryHistogramSeries,
          paretoLineSeriesIndex >= 0 ? normalizedSeries[paretoLineSeriesIndex] : undefined,
        )
      : null;
  const resolvedSeries = synthesizedParetoSeries
    ? [primaryHistogramSeries as XlsxChartSeries, synthesizedParetoSeries]
    : primaryHistogramSeries
      ? [
          primaryHistogramSeries,
          ...normalizedSeries.filter((_, seriesIndex) => seriesIndex !== clusteredColumnIndex),
        ]
      : collapseChartExPointSeries(chartType, normalizedSeries);
  const resolvedChartType = primaryHistogramSeries ? 'ColumnClustered' : chartType;

  const axes = plotAreaNode ? getLocalChildren(plotAreaNode, 'axis').map(readChartExAxis) : [];
  const legendNode = chartNode ? getFirstLocalChild(chartNode, 'legend') : null;

  chart.axes = axes;
  chart.categoryAxis = axes[0] ?? null;
  chart.valueAxis = axes.find((axis) => axis.numberFormat || axis.majorGridlines) ?? axes[1] ?? null;
  chart.chartExLayout = primaryLayout;
  chart.chartType = resolvedChartType;
  if (primaryHistogramSeries) chart.gapWidth = 0;
  chart.legend = legendNode
    ? {
        overlay: legendNode.getAttribute('overlay') === '1',
        position: normalizeLegendPosition(legendNode.getAttribute('pos') ?? undefined),
        raw: {},
      }
    : null;
  chart.name = chartTitle;
  chart.series = resolvedSeries;
  chart.title = chartTitle;
  chart.typeGroups = synthesizedParetoSeries
    ? [
        {
          chartType: 'ColumnClustered',
          gapWidth: 0,
          raw: { gapWidth: 0, layout: 'clusteredColumn' },
          series: [primaryHistogramSeries as XlsxChartSeries],
        },
        {
          chartType: 'Line',
          raw: { layout: 'paretoLine' },
          series: [synthesizedParetoSeries],
        },
      ]
    : [];
  const firstSeriesDataLabels = seriesNodes[0]
    ? parseChartDataLabelsFromXml(getFirstLocalChild(seriesNodes[0], 'dataLabels'))
    : null;
  chart.dataLabels = firstSeriesDataLabels;
  return chart;
}

function buildChartExParetoLineSeriesFromNodes(
  histogramSeries: XlsxChartSeries,
  paretoSeries: XlsxChartSeries | undefined,
): XlsxChartSeries {
  return buildChartExParetoLineSeries(histogramSeries, paretoSeries?.name);
}

// ---------------------------------------------------------------------------
// Regex fallbacks (mirrors react-xlsx behaviour for malformed chart XML).
// ---------------------------------------------------------------------------

function resolveColorFromXmlFragment(
  fragment: string,
  themePalette?: XlsxThemePalette | null,
): string | undefined {
  if (!fragment) return undefined;
  const srgbMatch = fragment.match(/<a:srgbClr\b[^>]*\bval="([0-9a-fA-F]{6,8})"/i);
  if (srgbMatch?.[1]) return normalizeHexColor(srgbMatch[1]) ?? undefined;
  const schemeMatch =
    fragment.match(/<a:schemeClr\b[^>]*\bval="([^"]+)"[^>]*>([\s\S]*?)<\/a:schemeClr>/i) ??
    fragment.match(/<a:schemeClr\b[^>]*\bval="([^"]+)"[^>]*/i);
  if (!schemeMatch?.[1]) return undefined;
  const baseColor = resolveThemeColor(schemeMatch[1], themePalette);
  if (!baseColor) return undefined;
  const transforms = schemeMatch[2] ?? '';
  let lightnessModifier = 1;
  let lightnessOffset = 0;
  for (const match of transforms.matchAll(/<a:(lumMod|lumOff|tint|shade)\b[^>]*\bval="(-?\d+(?:\.\d+)?)"/gi)) {
    const transform = match[1]?.toLowerCase();
    const rawValue = Number(match[2] ?? Number.NaN);
    if (!transform || !Number.isFinite(rawValue)) continue;
    if (transform === 'lummod') {
      lightnessModifier *= rawValue / 100000;
    } else if (transform === 'lumoff') {
      lightnessOffset += rawValue / 100000;
    } else if (transform === 'tint') {
      lightnessOffset += (1 - lightnessOffset) * (rawValue / 100000);
    } else if (transform === 'shade') {
      lightnessModifier *= rawValue / 100000;
    }
  }
  return applyLightnessTransform(baseColor, lightnessModifier, lightnessOffset) ?? undefined;
}

function readHexColorFromXmlFragment(
  fragment: string,
  preferLine = false,
  themePalette?: XlsxThemePalette | null,
): string | undefined {
  const source = preferLine
    ? (fragment.match(/<a:ln\b[\s\S]*?<\/a:ln>/i)?.[0] ?? '')
    : (fragment.match(/<a:solidFill\b[\s\S]*?<\/a:solidFill>/i)?.[0] ?? '');
  return resolveColorFromXmlFragment(source, themePalette);
}

type FallbackSeriesStyle = {
  color?: string | undefined;
  lineColor?: string | undefined;
};

function parseFallbackSeriesStylesFromChartXml(
  chartXml: string,
  themePalette?: XlsxThemePalette | null,
): FallbackSeriesStyle[] {
  const seriesBlocks = chartXml.match(/<c:ser\b[\s\S]*?<\/c:ser>/gi) ?? [];
  if (seriesBlocks.length === 0) return [];
  return seriesBlocks.map((seriesBlock) => {
    const shapeBlock = seriesBlock.match(/<c:spPr\b[\s\S]*?<\/c:spPr>/i)?.[0] ?? '';
    return {
      color: readHexColorFromXmlFragment(shapeBlock, false, themePalette),
      lineColor: readHexColorFromXmlFragment(shapeBlock, true, themePalette),
    };
  });
}

// ---------------------------------------------------------------------------
// Style application (ported from react-xlsx `applyChartStyleFromXml`).
// ---------------------------------------------------------------------------

function applyChartSeriesStyleFromXml(
  chart: XlsxChart,
  chartTypeNode: Element,
  themePalette?: XlsxThemePalette | null,
): void {
  const seriesNodes = getLocalChildren(chartTypeNode, 'ser');
  chart.series = chart.series.map((series, index) => {
    const seriesNode = seriesNodes[index];
    if (!seriesNode) return series;
    const shapeProperties = getFirstLocalChild(seriesNode, 'spPr');
    const markerNode = getFirstLocalChild(seriesNode, 'marker');
    const markerShapeProperties = getFirstLocalChild(markerNode ?? chartTypeNode, 'spPr');
    const lineStyle = resolveChartLineStyle(shapeProperties, themePalette);
    const markerLineStyle = resolveChartLineStyle(markerShapeProperties, themePalette);
    const fillColor = resolveChartFillColor(shapeProperties, themePalette);
    const markerSize = readChartNumericAttribute(markerNode, 'size');
    const markerSymbolNode = markerNode ? getFirstLocalChild(markerNode, 'symbol') : null;
    const markerSymbol = markerSymbolNode?.getAttribute('val') ?? undefined;
    const pointStyles = parseChartPointStyles(seriesNode, themePalette);
    const seriesExplosion = readChartNumericAttribute(seriesNode, 'explosion');
    const invertNegativeStyle = parseInvertNegativeStyle(seriesNode, themePalette);
    const invertIfNegative = readChartBooleanAttribute(seriesNode, 'invertIfNegative');
    const isScatterChart = isScatterLikeChartType(chart.chartType);
    const cachedCategories = isScatterChart
      ? (parseChartCacheValues(getFirstLocalChild(seriesNode, 'xVal'), 'numCache', 'value') ??
        parseChartMultiLevelCacheValues(getFirstLocalChild(seriesNode, 'xVal'), 'category'))
      : (parseChartCacheValues(getFirstLocalChild(seriesNode, 'cat'), 'strCache', 'category') ??
        parseChartCacheValues(getFirstLocalChild(seriesNode, 'cat'), 'numCache', 'category') ??
        parseChartMultiLevelCacheValues(getFirstLocalChild(seriesNode, 'cat'), 'category'));
    const cachedValues = isScatterChart
      ? parseChartCacheValues(getFirstLocalChild(seriesNode, 'yVal'), 'numCache', 'value')
      : parseChartCacheValues(getFirstLocalChild(seriesNode, 'val'), 'numCache', 'value');
    const cachedBubbleSizes =
      chart.chartType === 'Bubble'
        ? parseChartCacheValues(getFirstLocalChild(seriesNode, 'bubbleSize'), 'numCache', 'value')
        : null;
    const resolvedLineColor = lineStyle.hidden
      ? undefined
      : (lineStyle.color ?? fillColor ?? series.lineColor ?? series.color);
    return {
      ...series,
      bubbleSizes: cachedBubbleSizes
        ? cachedBubbleSizes.map((value) =>
            typeof value === 'number' && Number.isFinite(value) ? value : null,
          )
        : series.bubbleSizes,
      categories: cachedCategories ?? series.categories,
      color: fillColor ?? lineStyle.color ?? series.color,
      dataPointStyles: pointStyles.length > 0 ? pointStyles : series.dataPointStyles,
      lineColor: resolvedLineColor,
      lineWidthPx: lineStyle.hidden ? undefined : (lineStyle.widthPx ?? series.lineWidthPx),
      markerColor:
        resolveChartFillColor(markerShapeProperties, themePalette) ??
        fillColor ??
        lineStyle.color ??
        undefined,
      markerLineColor: markerLineStyle.color ?? lineStyle.color ?? fillColor ?? undefined,
      markerSize: markerSize ?? series.markerSize,
      markerSymbol,
      smooth: readChartBooleanAttribute(seriesNode, 'smooth') ?? series.smooth,
      invertIfNegative: invertIfNegative ?? series.invertIfNegative,
      shapeProperties: {
        ...series.shapeProperties,
        xmlExplosion: seriesExplosion ?? undefined,
        xmlFillColor: fillColor ?? undefined,
        xmlLineHidden: lineStyle.hidden ? true : undefined,
        xmlLineColor: lineStyle.color ?? undefined,
        xmlLineWidthPx: lineStyle.widthPx ?? undefined,
        xmlNegativeFillColor: invertNegativeStyle.color ?? undefined,
        xmlNegativeLineColor: invertNegativeStyle.lineColor ?? undefined,
      },
      negativeColor: invertNegativeStyle.color ?? series.negativeColor,
      negativeLineColor: invertNegativeStyle.lineColor ?? series.negativeLineColor,
      values: cachedValues
        ? cachedValues.map((value) => (typeof value === 'number' && Number.isFinite(value) ? value : null))
        : series.values,
    };
  });
}

export function applyChartStyleFromXml(
  chart: XlsxChart,
  chartXml: string,
  styleXml: string | undefined,
  colorsXml: string | undefined,
  themePalette?: XlsxThemePalette | null,
): void {
  const fallbackSeriesStyles = parseFallbackSeriesStylesFromChartXml(chartXml, themePalette);
  const applyFallbackSeriesStyles = (): void => {
    if (fallbackSeriesStyles.length === 0) return;
    chart.series = chart.series.map((series, seriesIndex) => {
      const fallbackStyle = fallbackSeriesStyles[seriesIndex];
      if (!fallbackStyle) return series;
      const fallbackColor = fallbackStyle.color ?? fallbackStyle.lineColor;
      return {
        ...series,
        color: series.color ?? fallbackColor,
        lineColor: series.lineColor ?? fallbackStyle.lineColor ?? fallbackColor,
        markerColor: series.markerColor ?? fallbackColor ?? series.color,
        markerLineColor:
          series.markerLineColor ?? fallbackStyle.lineColor ?? fallbackColor ?? series.lineColor,
      };
    });
  };
  const applyRelationshipStyles = (): ChartStyleAppearance => {
    const palette = readChartColorPalette(colorsXml, themePalette);
    if (palette.length > 0) chart.chartColorPalette = palette;
    const styleAppearance = readChartStyleAppearance(styleXml, themePalette);
    chart.axisLabelColor = styleAppearance.axisLabelColor ?? chart.axisLabelColor;
    chart.axisLineColor = styleAppearance.axisLineColor ?? chart.axisLineColor;
    chart.chartAreaBorderColor = styleAppearance.chartAreaBorderColor ?? chart.chartAreaBorderColor;
    chart.chartAreaFillColor = styleAppearance.chartAreaFillColor ?? chart.chartAreaFillColor;
    chart.chartColorPaletteOffset = styleAppearance.paletteOffset ?? chart.chartColorPaletteOffset;
    chart.textColor = styleAppearance.textColor ?? chart.textColor;
    chart.titleColor = styleAppearance.titleColor ?? chart.titleColor;
    return styleAppearance;
  };

  const chartDocument = parseXml(chartXml);
  const chartNode = chartDocument ? getFirstLocalDescendant(chartDocument, 'chart') : null;
  const plotAreaNode = chartNode ? getFirstLocalChild(chartNode, 'plotArea') : null;
  const styleIdNode = chartDocument?.documentElement
    ? getFirstLocalDescendant(chartDocument.documentElement, 'style')
    : null;
  const chartTypeNode = findPrimaryChartTypeNode(plotAreaNode);

  if (!chartNode || !chartTypeNode || !plotAreaNode) {
    applyRelationshipStyles();
    const fallbackStyleId = Number(styleIdNode?.getAttribute('val') ?? Number.NaN);
    if (Number.isFinite(fallbackStyleId)) chart.chartStyleId = fallbackStyleId;
    applyFallbackSeriesStyles();
    applyBuiltinChartDefaults(chart, themePalette);
    return;
  }
  const plotArea = plotAreaNode;

  const resolvedType = resolveClassicGroupChartType(chartTypeNode);
  chart.chartType = resolvedType.chartType;
  if (resolvedType.is3d !== undefined) chart.is3d = resolvedType.is3d;

  const legendNode = getFirstLocalChild(chartNode, 'legend');
  const legendPosition = legendNode
    ? (getFirstLocalChild(legendNode, 'legendPos')?.getAttribute('val') ?? undefined)
    : undefined;
  const legendOverlay = legendNode
    ? getFirstLocalChild(legendNode, 'overlay')?.getAttribute('val')
    : undefined;
  chart.legend = legendNode
    ? {
        overlay: legendOverlay === '1',
        position: normalizeLegendPosition(legendPosition),
        raw: chart.legend?.raw,
      }
    : chart.legend;
  const plotVisibleOnly = readChartBooleanAttribute(chartNode, 'plotVisOnly');
  if (plotVisibleOnly !== undefined) chart.plotVisibleOnly = plotVisibleOnly;
  chart.displayBlanksAs =
    getFirstLocalChild(chartNode, 'dispBlanksAs')?.getAttribute('val') ?? chart.displayBlanksAs;
  const styleId = Number(styleIdNode?.getAttribute('val') ?? Number.NaN);
  chart.chartStyleId = Number.isFinite(styleId) ? styleId : chart.chartStyleId;
  chart.firstSliceAngle = readChartNumericAttribute(chartTypeNode, 'firstSliceAng') ?? chart.firstSliceAngle;
  chart.gapWidth = readChartNumericAttribute(chartTypeNode, 'gapWidth') ?? chart.gapWidth;
  chart.overlap = readChartNumericAttribute(chartTypeNode, 'overlap') ?? chart.overlap;
  chart.bubbleScale = readChartNumericAttribute(chartTypeNode, 'bubbleScale') ?? chart.bubbleScale;
  chart.varyColors = readChartBooleanAttribute(chartTypeNode, 'varyColors') ?? chart.varyColors;
  const bubble3dNode = getFirstLocalChild(chartTypeNode, 'bubble3D');
  chart.bubble3d = bubble3dNode ? bubble3dNode.getAttribute('val') !== '0' : chart.bubble3d;
  chart.holeSize = readChartNumericAttribute(chartTypeNode, 'holeSize') ?? chart.holeSize;
  chart.radarStyle = getFirstLocalChild(chartTypeNode, 'radarStyle')?.getAttribute('val') ?? chart.radarStyle;
  chart.scatterStyle =
    getFirstLocalChild(chartTypeNode, 'scatterStyle')?.getAttribute('val') ?? chart.scatterStyle;
  chart.shape3d = getFirstLocalChild(chartTypeNode, 'shape')?.getAttribute('val') ?? chart.shape3d;
  const wireframeNode = getFirstLocalChild(chartTypeNode, 'wireframe');
  chart.wireframe = wireframeNode ? wireframeNode.getAttribute('val') !== '0' : chart.wireframe;
  const chartTypeDataLabels = parseChartDataLabelsFromXml(getFirstLocalChild(chartTypeNode, 'dLbls'));
  const firstSeriesNode = getLocalChildren(chartTypeNode, 'ser')[0] ?? null;
  const seriesDataLabels = parseChartDataLabelsFromXml(getFirstLocalChild(firstSeriesNode, 'dLbls'));
  chart.dataLabels = chartTypeDataLabels ?? seriesDataLabels ?? chart.dataLabels;
  const seriesSp3dNode = firstSeriesNode ? getFirstLocalDescendant(firstSeriesNode, 'sp3d') : null;
  chart.surfaceMaterial = seriesSp3dNode?.getAttribute('prstMaterial') ?? chart.surfaceMaterial;
  const bandFormatsNode = getLocalChildren(chartTypeNode, 'bandFmts')[0] ?? null;
  const bandFormatNodes = bandFormatsNode ? getLocalChildren(bandFormatsNode, 'bandFmt') : [];
  const bandFormatColors = bandFormatNodes
    .map((bandFormatNode) => {
      const shapeProperties = getFirstLocalChild(bandFormatNode, 'spPr');
      return resolveChartFillColor(shapeProperties, themePalette) ?? undefined;
    })
    .filter((color): color is string => typeof color === 'string' && color.length > 0);
  const bandFormatLineColors = bandFormatNodes
    .map((bandFormatNode) => {
      const shapeProperties = getFirstLocalChild(bandFormatNode, 'spPr');
      return resolveChartLineStyle(shapeProperties, themePalette).color ?? undefined;
    })
    .filter((color): color is string => typeof color === 'string' && color.length > 0);

  chart.raw = {
    ...(chart.raw ?? {}),
    bandFormatCount: bandFormatNodes.length > 0 ? bandFormatNodes.length : undefined,
    bandFormatColors: bandFormatColors.length > 0 ? bandFormatColors : undefined,
    bandFormatLineColors: bandFormatLineColors.length > 0 ? bandFormatLineColors : undefined,
    date1904: readChartBooleanAttribute(chartDocument?.documentElement ?? null, 'date1904'),
    bubble3d: chart.bubble3d,
    grouping: getFirstLocalChild(chartTypeNode, 'grouping')?.getAttribute('val') ?? undefined,
    ofPieType: getFirstLocalChild(chartTypeNode, 'ofPieType')?.getAttribute('val') ?? undefined,
    shape: getFirstLocalChild(chartTypeNode, 'shape')?.getAttribute('val') ?? undefined,
    secondPieSize: readChartNumericAttribute(chartTypeNode, 'secondPieSize'),
    scatterStyle: chart.scatterStyle,
    splitPos: readChartNumericAttribute(chartTypeNode, 'splitPos'),
    splitType: getFirstLocalChild(chartTypeNode, 'splitType')?.getAttribute('val') ?? undefined,
    xmlChartType: chartTypeNode.localName,
  };
  const view3dNode = getFirstLocalDescendant(chartNode, 'view3D');
  if (view3dNode) {
    chart.view3d = {
      depthPercent: readChartNumericAttribute(view3dNode, 'depthPercent'),
      perspective: readChartNumericAttribute(view3dNode, 'perspective'),
      rAngAx: getFirstLocalChild(view3dNode, 'rAngAx')?.getAttribute('val') === '1',
      rotX: readChartNumericAttribute(view3dNode, 'rotX'),
      rotY: readChartNumericAttribute(view3dNode, 'rotY'),
    };
  }
  chart.floor = readChartWallFromXml(getFirstLocalChild(chartNode, 'floor'), themePalette) ?? chart.floor;
  chart.sideWall =
    readChartWallFromXml(getFirstLocalChild(chartNode, 'sideWall'), themePalette) ?? chart.sideWall;
  chart.backWall =
    readChartWallFromXml(getFirstLocalChild(chartNode, 'backWall'), themePalette) ?? chart.backWall;

  const styleAppearance = applyRelationshipStyles();
  const chartTextTypeface = readChartTextTypeface(getFirstLocalChild(chartNode, 'txPr'), themePalette);
  const titleTypeface = readChartTextTypeface(getFirstLocalDescendant(chartNode, 'title'), themePalette);
  chart.fontFamily = chartTextTypeface ?? chart.fontFamily;
  chart.titleFontFamily = titleTypeface ?? chart.titleFontFamily ?? chart.fontFamily;

  const chartAreaShapeProperties = chartDocument?.documentElement
    ? getFirstLocalChild(chartDocument.documentElement, 'spPr')
    : null;
  const plotAreaShapeProperties = getFirstLocalChild(plotArea, 'spPr');
  const chartAreaNoFill = chartAreaShapeProperties
    ? getFirstLocalChild(chartAreaShapeProperties, 'noFill') != null
    : false;
  const plotAreaNoFill = plotAreaShapeProperties
    ? getFirstLocalChild(plotAreaShapeProperties, 'noFill') != null
    : false;
  chart.raw = {
    ...(chart.raw ?? {}),
    chartAreaNoFill: styleAppearance.chartAreaNoFill === true || chartAreaNoFill,
    plotAreaNoFill,
  };
  if (chartAreaShapeProperties) {
    const chartAreaFillColor = resolveChartFillColor(chartAreaShapeProperties, themePalette);
    if (chartAreaFillColor) {
      chart.chartAreaFillColor = chartAreaFillColor;
    } else if (getFirstLocalChild(chartAreaShapeProperties, 'noFill')) {
      chart.chartAreaFillColor = 'transparent';
    }
    const chartAreaLineStyle = resolveChartLineStyle(chartAreaShapeProperties, themePalette);
    if (chartAreaLineStyle.hidden) {
      chart.chartAreaBorderColor = 'transparent';
    } else if (chartAreaLineStyle.color) {
      chart.chartAreaBorderColor = chartAreaLineStyle.color;
    }
  }
  if (!chart.chartAreaFillColor && (styleAppearance.chartAreaNoFill === true || plotAreaNoFill)) {
    chart.chartAreaFillColor = 'transparent';
  }

  const categoryAxisNodes = [
    ...getLocalChildren(plotArea, 'catAx'),
    ...getLocalChildren(plotArea, 'dateAx'),
  ];
  const valueAxisNodes = getLocalChildren(plotArea, 'valAx');
  const seriesAxisNode = getLocalChildren(plotArea, 'serAx')[0] ?? null;
  const isScatterLike = isScatterLikeChartType(chart.chartType);
  let categoryAxisNode = categoryAxisNodes[0] ?? null;
  let valueAxisNode = valueAxisNodes[0] ?? null;
  if (!categoryAxisNode && isScatterLike && valueAxisNodes.length >= 2) {
    categoryAxisNode =
      valueAxisNodes.find((axisNode) => {
        const position = getFirstLocalChild(axisNode, 'axPos')?.getAttribute('val');
        return position === 'b' || position === 't';
      }) ?? valueAxisNodes[0] ?? null;
    valueAxisNode =
      valueAxisNodes.find((axisNode) => {
        const position = getFirstLocalChild(axisNode, 'axPos')?.getAttribute('val');
        return position === 'l' || position === 'r';
      }) ??
      valueAxisNodes[1] ??
      valueAxisNodes[0] ??
      null;
  }
  chart.categoryAxis = mergeChartAxis(chart.categoryAxis, readChartAxisFromXml(categoryAxisNode));
  chart.valueAxis = mergeChartAxis(chart.valueAxis, readChartAxisFromXml(valueAxisNode));
  chart.seriesAxis = mergeChartAxis(chart.seriesAxis, readChartAxisFromXml(seriesAxisNode));
  chart.axes =
    chart.axes.length > 0
      ? chart.axes.map((axis, index) =>
          index === 0 && categoryAxisNode
            ? { ...axis, ...readChartAxisFromXml(categoryAxisNode) }
            : index === 1 && valueAxisNode
              ? { ...axis, ...readChartAxisFromXml(valueAxisNode) }
              : axis,
        )
      : chart.axes;
  if (seriesAxisNode) {
    const seriesAxis = readChartAxisFromXml(seriesAxisNode);
    if (seriesAxis && !chart.axes.some((axis) => axis.id != null && axis.id === seriesAxis.id)) {
      chart.axes = [...chart.axes, seriesAxis as XlsxChartAxis];
    }
  }

  applyChartSeriesStyleFromXml(chart, chartTypeNode, themePalette);
  applyFallbackSeriesStyles();
  applyBuiltinChartDefaults(chart, themePalette);
}

// ---------------------------------------------------------------------------
// Entry points.
// ---------------------------------------------------------------------------

const LEGACY_CHART_TYPE_MAP: Record<string, string> = {
  bar: 'ColumnClustered',
  line: 'Line',
  pie: 'Pie',
  doughnut: 'Doughnut',
  area: 'Area',
  scatter: 'Scatter',
  radar: 'Radar',
  bubble: 'Bubble',
  stock: 'Stock',
  surface: 'Surface',
};

/** Builds a renderable chart from the parsed summary when no XML is available. */
function buildFallbackChart(node: ChartNode): XlsxChart {
  const chart = createEmptyChart(node.id);
  chart.chartType = LEGACY_CHART_TYPE_MAP[node.chartType] ?? 'ColumnClustered';
  chart.title = node.title;
  chart.name = node.title;
  chart.legend = node.hasLegend ? { position: 'b', raw: {} } : null;
  chart.series = node.series.map((series, index) => ({
    bubbleSizeRef: null,
    bubbleSizes: [],
    categories: (series.categories ?? []).map((value) =>
      typeof value === 'string' || typeof value === 'number' ? value : null,
    ),
    categoriesRef: null,
    color: series.color?.value ? (normalizeHexColor(series.color.value) ?? undefined) : undefined,
    dataPoints: [],
    id: `${node.id}-series-${index}`,
    name: series.name,
    raw: {},
    values: series.values.map((value) =>
      typeof value === 'number' && Number.isFinite(value) ? value : null,
    ),
    valuesRef: null,
  }));
  return chart;
}

function isChartExDocument(chartDocument: XMLDocument): boolean {
  return getFirstLocalDescendant(chartDocument, 'chartData') != null;
}

/**
 * Builds a fully styled chart model for a PowerPoint chart node using the same
 * pipeline as react-xlsx: base model, chart XML styles, companion style and
 * color parts, and built-in Office defaults driven by the slide theme.
 */
export function buildPptxChart(node: ChartNode, themePalette: XlsxThemePalette | null): XlsxChart {
  const chartXml = node.chartXml;
  if (chartXml) {
    const chartDocument = parseXml(chartXml);
    if (chartDocument) {
      if (isChartExDocument(chartDocument)) {
        const chart = buildChartExBase(node.id, chartDocument, themePalette);
        const palette = readChartColorPalette(node.chartColorsXml, themePalette);
        if (palette.length > 0) chart.chartColorPalette = palette;
        const styleAppearance = readChartStyleAppearance(node.chartStyleXml, themePalette);
        chart.axisLabelColor = styleAppearance.axisLabelColor ?? chart.axisLabelColor;
        chart.axisLineColor = styleAppearance.axisLineColor ?? chart.axisLineColor;
        chart.chartAreaBorderColor = styleAppearance.chartAreaBorderColor ?? chart.chartAreaBorderColor;
        chart.chartAreaFillColor = styleAppearance.chartAreaFillColor ?? chart.chartAreaFillColor;
        chart.chartColorPaletteOffset = styleAppearance.paletteOffset ?? chart.chartColorPaletteOffset;
        chart.textColor = styleAppearance.textColor ?? chart.textColor;
        chart.titleColor = styleAppearance.titleColor ?? chart.titleColor;
        if (styleAppearance.chartAreaNoFill === true && !chart.chartAreaFillColor) {
          chart.chartAreaFillColor = 'transparent';
        }
        applyBuiltinChartDefaults(chart, themePalette);
        return chart;
      }
      const chart = buildClassicChartBase(node.id, chartDocument, themePalette);
      applyChartStyleFromXml(chart, chartXml, node.chartStyleXml, node.chartColorsXml, themePalette);
      return chart;
    }
  }
  const chart = buildFallbackChart(node);
  applyBuiltinChartDefaults(chart, themePalette);
  return chart;
}
