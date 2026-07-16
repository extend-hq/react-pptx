import { describe, expect, it } from 'vitest';
import type { ChartNode } from '@extend-ai/react-pptx-model';
import { buildPptxChart, buildThemePaletteFromPresentationTheme } from './chart-model';

const OFFICE_THEME = {
  id: 'ppt/theme/theme1.xml',
  colors: {
    dk1: '#000000',
    lt1: '#FFFFFF',
    dk2: '#44546A',
    lt2: '#E7E6E6',
    accent1: '#4472C4',
    accent2: '#ED7D31',
    accent3: '#A5A5A5',
    accent4: '#FFC000',
    accent5: '#5B9BD5',
    accent6: '#70AD47',
    hlink: '#0563C1',
    folHlink: '#954F72',
  },
  minorFonts: { latin: 'Calibri' },
  majorFonts: { latin: 'Calibri Light' },
};

function chartNode(chartXml: string, extra: Partial<ChartNode> = {}): ChartNode {
  return {
    id: 'chart-1',
    type: 'chart',
    transform: { x: 0, y: 0, width: 3_000_000, height: 2_000_000 },
    chartType: 'bar',
    series: [],
    chartXml,
    ...extra,
  };
}

const CLASSIC_BAR_XML = `<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title><c:tx><c:rich><a:p><a:r><a:t>Quarterly</a:t></a:r></a:p></c:rich></c:tx></c:title>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:gapWidth val="150"/>
        <c:overlap val="-27"/>
        <c:ser>
          <c:idx val="0"/>
          <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Explicit</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:spPr><a:solidFill><a:srgbClr val="D8663A"/></a:solidFill></c:spPr>
          <c:cat><c:strRef><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="1"/>
          <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Scheme</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:spPr><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></c:spPr>
          <c:val><c:numRef><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>5</c:v></c:pt><c:pt idx="1"><c:v>8</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="2"/>
          <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Themed</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:val><c:numRef><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>7</c:v></c:pt><c:pt idx="1"><c:v>9</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
        <c:axId val="111"/>
        <c:axId val="222"/>
      </c:barChart>
      <c:catAx>
        <c:axId val="111"/>
        <c:axPos val="b"/>
        <c:crossAx val="222"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="222"/>
        <c:scaling><c:max val="25"/><c:min val="0"/></c:scaling>
        <c:majorGridlines/>
        <c:numFmt formatCode="#,##0" sourceLinked="0"/>
        <c:axPos val="l"/>
        <c:crossAx val="111"/>
      </c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>
    <c:dispBlanksAs val="gap"/>
  </c:chart>
</c:chartSpace>`;

const SCATTER_WITH_BUILTIN_MARKERS_XML = `<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart">
  <mc:AlternateContent><mc:Choice Requires="c14"><c14:style val="118"/></mc:Choice><mc:Fallback><c:style val="18"/></mc:Fallback></mc:AlternateContent>
  <c:chart><c:plotArea><c:scatterChart>
    <c:scatterStyle val="lineMarker"/>
    <c:ser>
      <c:idx val="0"/><c:order val="0"/>
      <c:spPr><a:ln><a:noFill/></a:ln></c:spPr>
      <c:xVal><c:numRef><c:numCache><c:ptCount val="3"/><c:pt idx="0"><c:v>0.7</c:v></c:pt><c:pt idx="1"><c:v>1.8</c:v></c:pt><c:pt idx="2"><c:v>2.6</c:v></c:pt></c:numCache></c:numRef></c:xVal>
      <c:yVal><c:numRef><c:numCache><c:ptCount val="3"/><c:pt idx="0"><c:v>2.7</c:v></c:pt><c:pt idx="1"><c:v>3.2</c:v></c:pt><c:pt idx="2"><c:v>0.8</c:v></c:pt></c:numCache></c:numRef></c:yVal>
    </c:ser>
  </c:scatterChart></c:plotArea></c:chart>
</c:chartSpace>`;

describe('buildPptxChart', () => {
  const palette = buildThemePaletteFromPresentationTheme(OFFICE_THEME);

  it('maps the slide theme to the Excel-style palette', () => {
    expect(palette).not.toBeNull();
    expect(palette?.colorsByIndex[1]).toBe('#000000');
    expect(palette?.colorsByIndex[4]).toBe('#4472c4');
    expect(palette?.colorsByIndex[9]).toBe('#70ad47');
    expect(palette?.minorLatinFont).toBe('Calibri');
  });

  it('builds a fully styled classic bar chart with PowerPoint colors', () => {
    const chart = buildPptxChart(chartNode(CLASSIC_BAR_XML), palette);
    expect(chart.chartType).toBe('ColumnClustered');
    expect(chart.title).toBe('Quarterly');
    expect(chart.gapWidth).toBe(150);
    expect(chart.overlap).toBe(-27);
    expect(chart.legend?.position).toBe('b');
    expect(chart.displayBlanksAs).toBe('gap');
    expect(chart.series).toHaveLength(3);
    expect(chart.series[0]?.name).toBe('Explicit');
    expect(chart.series[0]?.values).toEqual([10, 20]);
    expect(chart.series[0]?.categories).toEqual(['Q1', 'Q2']);
    // Explicit sRGB fill wins.
    expect(chart.series[0]?.color).toBe('#d8663a');
    // schemeClr resolves through the slide theme.
    expect(chart.series[1]?.color).toBe('#ed7d31');
    // Unstyled series fall back to the theme accent cycle by index.
    expect(chart.series[2]?.color).toBe('#a5a5a5');
    // Axes carry ids, scaling, and number formats for the renderer.
    expect(chart.valueAxis?.max).toBe(25);
    expect(chart.valueAxis?.majorGridlines).toBe(true);
    expect(chart.valueAxis?.numberFormat?.formatCode).toBe('#,##0');
    expect(chart.axes.some((axis) => axis.id === 111)).toBe(true);
    expect(chart.axes.some((axis) => axis.id === 222)).toBe(true);
    // Chart text falls back to the theme minor font and dk1 text color.
    expect(chart.fontFamily).toBe('Calibri');
    expect(chart.textColor).toBe('#000000');
  });

  it('resolves omitted scatter markers through the built-in chart style', () => {
    const chart = buildPptxChart(chartNode(SCATTER_WITH_BUILTIN_MARKERS_XML), palette);

    expect(chart.chartType).toBe('ScatterLines');
    expect(chart.scatterStyle).toBe('lineMarker');
    expect(chart.series[0]?.categories).toEqual([0.7, 1.8, 2.6]);
    expect(chart.series[0]?.values).toEqual([2.7, 3.2, 0.8]);
    expect(chart.series[0]?.shapeProperties?.xmlLineHidden).toBe(true);
    expect(chart.series[0]?.markerSymbol).toBe('diamond');
  });

  it('preserves an explicit none scatter marker', () => {
    const chart = buildPptxChart(
      chartNode(
        SCATTER_WITH_BUILTIN_MARKERS_XML.replace(
          '<c:idx val="0"/><c:order val="0"/>',
          '<c:idx val="0"/><c:order val="0"/><c:marker><c:symbol val="none"/></c:marker>',
        ),
      ),
      palette,
    );

    expect(chart.series[0]?.markerSymbol).toBe('none');
  });

  it('reads the Microsoft chart color style part for the series palette', () => {
    const colorsXml = `<?xml version="1.0"?>
<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" meth="cycle" id="10">
  <a:schemeClr val="accent6"/>
  <a:srgbClr val="112233"/>
  <cs:variation/>
</cs:colorStyle>`;
    const chart = buildPptxChart(
      chartNode(CLASSIC_BAR_XML, { chartColorsXml: colorsXml }),
      palette,
    );
    expect(chart.chartColorPalette).toEqual(['#70ad47', '#112233']);
    // The unstyled third series now cycles the chart color style palette.
    expect(chart.series[2]?.color).toBe('#70ad47');
  });

  it('builds doughnut charts with hole size and vary-colors point styling', () => {
    const doughnutXml = `<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart><c:plotArea>
    <c:doughnutChart>
      <c:varyColors val="1"/>
      <c:ser>
        <c:idx val="0"/>
        <c:dPt><c:idx val="1"/><c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></c:spPr></c:dPt>
        <c:cat><c:strRef><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strCache></c:strRef></c:cat>
        <c:val><c:numRef><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>60</c:v></c:pt><c:pt idx="1"><c:v>40</c:v></c:pt></c:numCache></c:numRef></c:val>
      </c:ser>
      <c:holeSize val="62"/>
    </c:doughnutChart>
  </c:plotArea></c:chart>
</c:chartSpace>`;
    const chart = buildPptxChart(chartNode(doughnutXml), palette);
    expect(chart.chartType).toBe('Doughnut');
    expect(chart.holeSize).toBe(62);
    expect(chart.varyColors).toBe(true);
    expect(chart.series[0]?.dataPointStyles).toEqual([
      { color: '#ff0000', explosion: undefined, index: 1, lineColor: undefined },
    ]);
  });

  it('builds modern chartEx models from cached dimensions', () => {
    const chartExXml = `<?xml version="1.0"?>
<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <cx:chartData>
    <cx:data id="0">
      <cx:strDim type="cat"><cx:f>Sheet1!$A$2:$A$4</cx:f><cx:lvl ptCount="3"><cx:pt idx="0">Bronze</cx:pt><cx:pt idx="1">Silver</cx:pt><cx:pt idx="2">Gold</cx:pt></cx:lvl></cx:strDim>
      <cx:numDim type="val"><cx:f>Sheet1!$B$2:$B$4</cx:f><cx:lvl ptCount="3"><cx:pt idx="0">120</cx:pt><cx:pt idx="1">80</cx:pt><cx:pt idx="2">40</cx:pt></cx:lvl></cx:numDim>
    </cx:data>
  </cx:chartData>
  <cx:chart>
    <cx:title><cx:tx><cx:txData><cx:v>Pipeline</cx:v></cx:txData></cx:tx></cx:title>
    <cx:plotArea><cx:plotAreaRegion>
      <cx:series layoutId="funnel" uniqueId="{1}"><cx:dataId val="0"/></cx:series>
    </cx:plotAreaRegion></cx:plotArea>
    <cx:legend pos="t"/>
  </cx:chart>
</cx:chartSpace>`;
    const chart = buildPptxChart(chartNode(chartExXml, { chartType: 'chartEx' }), palette);
    expect(chart.chartType).toBe('Funnel');
    expect(chart.chartExLayout).toBe('funnel');
    expect(chart.title).toBe('Pipeline');
    expect(chart.legend?.position).toBe('t');
    expect(chart.series).toHaveLength(1);
    expect(chart.series[0]?.values).toEqual([120, 80, 40]);
    expect(chart.series[0]?.categories).toEqual(['Bronze', 'Silver', 'Gold']);
  });

  it('falls back to the parsed summary when no chart XML is available', () => {
    const chart = buildPptxChart(
      {
        id: 'legacy',
        type: 'chart',
        transform: { x: 0, y: 0, width: 1, height: 1 },
        chartType: 'pie',
        title: 'Legacy',
        hasLegend: true,
        series: [
          { name: 'S1', categories: ['A', 'B'], values: [1, 2], color: { value: '#336699' } },
        ],
      },
      palette,
    );
    expect(chart.chartType).toBe('Pie');
    expect(chart.title).toBe('Legacy');
    expect(chart.legend?.position).toBe('b');
    expect(chart.series[0]?.color).toBe('#336699');
    expect(chart.series[0]?.values).toEqual([1, 2]);
  });
});
