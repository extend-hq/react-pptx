import { describe, expect, it } from 'vitest';
import type { ChartNode } from '@extend-ai/react-pptx-model';
import { renderChartInto } from './chart-host';
import { loadRegionMapAtlas } from './region-map-data';

const REGION_MAP_CHART_EX_XML = `<?xml version="1.0"?>
<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <cx:chartData>
    <cx:data id="0">
      <cx:strDim type="cat"><cx:f>Sheet1!$A$2:$A$4</cx:f><cx:lvl ptCount="3"><cx:pt idx="0">France</cx:pt><cx:pt idx="1">Germany</cx:pt><cx:pt idx="2">Spain</cx:pt></cx:lvl></cx:strDim>
      <cx:numDim type="colorVal"><cx:f>Sheet1!$B$2:$B$4</cx:f><cx:lvl ptCount="3"><cx:pt idx="0">10</cx:pt><cx:pt idx="1">40</cx:pt><cx:pt idx="2">25</cx:pt></cx:lvl></cx:numDim>
    </cx:data>
  </cx:chartData>
  <cx:chart>
    <cx:plotArea><cx:plotAreaRegion>
      <cx:series layoutId="regionMap" uniqueId="{1}"><cx:dataId val="0"/></cx:series>
    </cx:plotAreaRegion></cx:plotArea>
  </cx:chart>
</cx:chartSpace>`;

const DEFAULT_MARKER_SCATTER_XML = `<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:style val="18"/>
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

function regionMapNode(): ChartNode {
  return {
    id: 'map-chart',
    type: 'chart',
    transform: { x: 0, y: 0, width: 6_000_000, height: 4_000_000 },
    chartType: 'chartEx',
    series: [],
    chartXml: REGION_MAP_CHART_EX_XML,
  };
}

describe('renderChartInto region maps', () => {
  it('renders immediately and fills in geography once the atlas chunk loads', async () => {
    const container = document.createElement('div');
    document.body.append(container);

    expect(renderChartInto(container, regionMapNode(), undefined)).toBe(true);
    const pathsBeforeAtlas = container.querySelectorAll('path').length;
    expect(container.querySelector('svg')).not.toBeNull();

    // The host kicks off the lazy atlas load itself; awaiting the shared
    // loader plus a microtask gives its re-render a chance to run.
    await loadRegionMapAtlas();
    await new Promise((resolve) => setTimeout(resolve, 0));

    const pathsAfterAtlas = container.querySelectorAll('path').length;
    // The world map alone contributes hundreds of country outlines.
    expect(pathsAfterAtlas).toBeGreaterThan(pathsBeforeAtlas + 100);
    container.remove();
  });
});

describe('renderChartInto scatter markers', () => {
  it('renders points when the marker symbol is supplied by the built-in style', () => {
    const container = document.createElement('div');
    const node: ChartNode = {
      id: 'scatter-chart',
      type: 'chart',
      transform: { x: 0, y: 0, width: 6_096_000, height: 4_064_000 },
      chartType: 'scatter',
      series: [],
      chartXml: DEFAULT_MARKER_SCATTER_XML,
    };

    expect(renderChartInto(container, node, undefined)).toBe(true);
    expect(container.querySelectorAll('path[data-xlsx-chart-point-index]')).toHaveLength(3);
  });
});
