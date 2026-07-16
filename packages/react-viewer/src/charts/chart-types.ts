/**
 * Chart model types vendored from @extend-ai/react-xlsx so PowerPoint charts
 * render with the exact same model, styles, and colors as Excel charts.
 * Keep the `Xlsx*` names intact to make future syncs with react-xlsx diffable.
 */

export interface XlsxThemePalette {
  colorsByIndex: Record<number, string>;
  majorLatinFont?: string | undefined;
  minorLatinFont?: string | undefined;
}

export interface XlsxImageMarker {
  col: number;
  colOffsetEmu: number;
  row: number;
  rowOffsetEmu: number;
}

export type XlsxImageAnchor =
  | {
      from: XlsxImageMarker;
      kind: 'one-cell';
      sizeEmu: {
        cx: number;
        cy: number;
      };
    }
  | {
      kind: 'absolute';
      positionEmu: {
        x: number;
        y: number;
      };
      sizeEmu: {
        cx: number;
        cy: number;
      };
    }
  | {
      from: XlsxImageMarker;
      kind: 'two-cell';
      to: XlsxImageMarker;
    };

export interface XlsxImageRect {
  height: number;
  left: number;
  top: number;
  width: number;
}

export interface XlsxChartReference {
  formula?: string | undefined;
  refType?: string | undefined;
  values?: Array<number | string | null> | undefined;
}

export interface XlsxChartDataLabels {
  pointLabels?: XlsxChartPointDataLabel[] | undefined;
  raw?: Record<string, unknown> | undefined;
  showBubbleSize?: boolean | undefined;
  showCategoryName?: boolean | undefined;
  showLegendKey?: boolean | undefined;
  showPercent?: boolean | undefined;
  showSeriesName?: boolean | undefined;
  showValue?: boolean | undefined;
}

export interface XlsxChartPointDataLabel {
  deleted?: boolean | undefined;
  fontSizePt?: number | undefined;
  index: number;
  showBubbleSize?: boolean | undefined;
  showCategoryName?: boolean | undefined;
  showPercent?: boolean | undefined;
  showSeriesName?: boolean | undefined;
  showValue?: boolean | undefined;
  x?: number | undefined;
  y?: number | undefined;
}

export interface XlsxChartLegend {
  overlay?: boolean | undefined;
  position?: string | undefined;
  raw?: Record<string, unknown> | undefined;
}

export interface XlsxChartAxis {
  crossId?: number | undefined;
  crosses?: string | undefined;
  crossBetween?: string | undefined;
  delete?: boolean | undefined;
  id?: number | undefined;
  labelPosition?: string | undefined;
  logBase?: number | undefined;
  orientation?: string | undefined;
  majorUnit?: number | undefined;
  max?: number | undefined;
  min?: number | undefined;
  majorGridlines?: boolean | undefined;
  majorTickMark?: string | undefined;
  minorUnit?: number | undefined;
  minorGridlines?: boolean | undefined;
  minorTickMark?: string | undefined;
  numberFormat?: {
    formatCode?: string | undefined;
    sourceLinked?: boolean | undefined;
  };
  position?: string | undefined;
  raw?: Record<string, unknown> | undefined;
  shapeProperties?: Record<string, unknown> | undefined;
  tickLabelSkip?: number | undefined;
  tickMarkSkip?: number | undefined;
}

export interface XlsxChartPointStyle {
  color?: string | undefined;
  explosion?: number | undefined;
  index: number;
  lineColor?: string | undefined;
}

export interface XlsxChartSeries {
  bubbleSizeRef?: XlsxChartReference | null | undefined;
  bubbleSizes?: Array<number | null> | undefined;
  categories: Array<number | string | null>;
  categoriesRef?: XlsxChartReference | null | undefined;
  color?: string | undefined;
  dataPoints: unknown[];
  dataPointStyles?: XlsxChartPointStyle[] | undefined;
  formatIdx?: number | undefined;
  hidden?: boolean | undefined;
  id: string;
  invertIfNegative?: boolean | undefined;
  lineColor?: string | undefined;
  lineWidthPx?: number | undefined;
  marker?: Record<string, unknown> | undefined;
  markerColor?: string | undefined;
  markerLineColor?: string | undefined;
  markerSize?: number | undefined;
  markerSymbol?: string | undefined;
  name?: string | undefined;
  negativeColor?: string | undefined;
  negativeLineColor?: string | undefined;
  raw?: Record<string, unknown> | undefined;
  shapeProperties?: Record<string, unknown> | undefined;
  smooth?: boolean | undefined;
  values: Array<number | null>;
  valuesRef?: XlsxChartReference | null | undefined;
}

export type XlsxChartElementSelection =
  | { kind: 'chart'; chartId: string }
  | { kind: 'series'; chartId: string; seriesId: string; seriesIndex: number }
  | { kind: 'point'; chartId: string; seriesId: string; seriesIndex: number; pointIndex: number }
  | { kind: 'legendEntry'; chartId: string; seriesId: string; seriesIndex: number };

export interface XlsxChartTypeGroup {
  axisIds?: number[] | undefined;
  chartType: string;
  dataLabels?: XlsxChartDataLabels | null | undefined;
  gapWidth?: number | undefined;
  is3d?: boolean | undefined;
  overlap?: number | undefined;
  raw?: Record<string, unknown> | undefined;
  series: XlsxChartSeries[];
  varyColors?: boolean | undefined;
}

export interface XlsxChartWall {
  fillColor?: string | undefined;
  hidden?: boolean | undefined;
  lineColor?: string | undefined;
  thickness?: number | undefined;
}

export interface XlsxChart {
  anchor: XlsxImageAnchor;
  autoTitleDeleted?: boolean | undefined;
  axes: XlsxChartAxis[];
  axisLabelColor?: string | undefined;
  axisLineColor?: string | undefined;
  categoryAxis?: XlsxChartAxis | null | undefined;
  chartExLayout?: string | undefined;
  chartAreaBorderColor?: string | undefined;
  chartAreaFillColor?: string | undefined;
  chartColorPalette?: string[] | undefined;
  chartColorPaletteOffset?: number | undefined;
  chartPath?: string | undefined;
  chartStyleId?: number | undefined;
  chartType: string;
  dataLabels?: XlsxChartDataLabels | null | undefined;
  displayBlanksAs?: string | undefined;
  editable?: boolean | undefined;
  firstSliceAngle?: number | undefined;
  fontFamily?: string | undefined;
  gapWidth?: number | undefined;
  holeSize?: number | undefined;
  id: string;
  is3d?: boolean | undefined;
  legend?: XlsxChartLegend | null | undefined;
  name?: string | undefined;
  overlap?: number | undefined;
  plotVisibleOnly?: boolean | undefined;
  raw?: Record<string, unknown> | undefined;
  radarStyle?: string | undefined;
  scatterStyle?: string | undefined;
  roundedCorners?: boolean | undefined;
  shape3d?: string | undefined;
  seriesAxis?: XlsxChartAxis | null | undefined;
  series: XlsxChartSeries[];
  sheetIndex: number;
  showDlblsOverMax?: boolean | undefined;
  sideWall?: XlsxChartWall | null | undefined;
  backWall?: XlsxChartWall | null | undefined;
  bubbleScale?: number | undefined;
  bubble3d?: boolean | undefined;
  floor?: XlsxChartWall | null | undefined;
  surfaceMaterial?: string | undefined;
  textColor?: string | undefined;
  title?: string | undefined;
  titleColor?: string | undefined;
  titleFontFamily?: string | undefined;
  typeGroups?: XlsxChartTypeGroup[] | undefined;
  valueAxis?: XlsxChartAxis | null | undefined;
  varyColors?: boolean | undefined;
  view3d?: {
    depthPercent?: number | undefined;
    perspective?: number | undefined;
    rAngAx?: boolean | undefined;
    rotX?: number | undefined;
    rotY?: number | undefined;
  };
  wireframe?: boolean | undefined;
  workbookSheetIndex: number;
  zIndex: number;
}
