#!/usr/bin/env node
import { copyFile, mkdir } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import PptxGenJS from 'pptxgenjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const outputDirectory = path.join(root, 'tests', 'fixtures', 'generated');
await mkdir(outputDirectory, { recursive: true });

const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'react-pptx test suite';
pptx.subject = 'Deterministic viewer fixture';
pptx.title = 'Rendering systems';
pptx.company = 'Extend AI';
pptx.lang = 'en-US';
pptx.theme = {
  headFontFace: 'Aptos Display',
  bodyFontFace: 'Aptos',
  lang: 'en-US',
};

const first = pptx.addSlide();
first.background = { color: 'F4F0E8' };
first.addShape(pptx.ShapeType.rect, {
  x: 0,
  y: 0,
  w: 3.25,
  h: 7.5,
  line: { color: '14362C', transparency: 100 },
  fill: { color: '14362C' },
});
first.addText('RENDERING\nSYSTEMS', {
  x: 0.55,
  y: 0.72,
  w: 2.2,
  h: 1.7,
  fontFace: 'Aptos Display',
  fontSize: 27,
  bold: true,
  color: 'F4F0E8',
  breakLine: false,
  margin: 0,
});
first.addText('A deterministic PowerPoint fixture', {
  x: 3.85,
  y: 1.15,
  w: 7.85,
  h: 0.6,
  fontFace: 'Aptos Display',
  fontSize: 27,
  bold: true,
  color: '18201D',
  margin: 0,
});
first.addText('Text, shapes, grouping, tables, charts, links, and slide navigation.', {
  x: 3.88,
  y: 2.05,
  w: 7.1,
  h: 0.8,
  fontFace: 'Aptos',
  fontSize: 16,
  color: '4C5752',
  margin: 0,
});
first.addShape(pptx.ShapeType.arc, {
  x: 8.9,
  y: 3.45,
  w: 2.3,
  h: 2.3,
  adjustPoint: 0.32,
  rotate: 18,
  line: { color: 'D8663A', width: 3 },
  fill: { color: 'D8663A', transparency: 100 },
});

const second = pptx.addSlide();
second.background = { color: 'F4F0E8' };
second.addText('Measured output', {
  x: 0.7,
  y: 0.55,
  w: 5,
  h: 0.5,
  fontFace: 'Aptos Display',
  fontSize: 26,
  bold: true,
  color: '18201D',
  margin: 0,
});
second.addTable(
  [
    [
      { text: 'Surface', options: { bold: true, color: 'F4F0E8' } },
      { text: 'Target', options: { bold: true, color: 'F4F0E8' } },
      { text: 'Result', options: { bold: true, color: 'F4F0E8' } },
    ],
    ['First slide', '< 1 second', 'Ready'],
    ['Cached navigation', '< 100 ms', 'Ready'],
    ['Unmount cleanup', '0 retained handles', 'Ready'],
  ],
  {
    x: 0.75,
    y: 1.45,
    w: 6.4,
    h: 3.3,
    border: { color: 'B8B0A4', width: 1 },
    fill: 'FBF8F2',
    color: '18201D',
    fontFace: 'Aptos',
    fontSize: 14,
    margin: 0.12,
    rowH: 0.6,
  },
);
second.addChart(
  pptx.ChartType.bar,
  [
    { name: 'Baseline', labels: ['Parse', 'Render', 'Navigate'], values: [700, 280, 70] },
    { name: 'Budget', labels: ['Parse', 'Render', 'Navigate'], values: [900, 400, 100] },
  ],
  {
    x: 7.65,
    y: 1.2,
    w: 4.9,
    h: 4.8,
    catAxisLabelFontFace: 'Aptos',
    valAxisLabelFontFace: 'Aptos',
    showLegend: true,
    showTitle: false,
    chartColors: ['D8663A', '14362C'],
    showValue: false,
  },
);

const fixturePath = path.join(outputDirectory, 'viewer-smoke.pptx');
await pptx.writeFile({ fileName: fixturePath });
const playgroundPublic = path.join(root, 'apps', 'playground', 'public');
await mkdir(playgroundPublic, { recursive: true });
await copyFile(fixturePath, path.join(playgroundPublic, 'viewer-smoke.pptx'));
console.log(fixturePath);

// Chart showcase: one deck covering the chart types the viewer renders with
// react-xlsx parity, for manual QA in the playground.
const charts = new PptxGenJS();
charts.defineLayout({ name: 'WIDE', width: 13.33, height: 7.5 });
charts.layout = 'WIDE';

const chartLabels = ['Q1', 'Q2', 'Q3', 'Q4'];
const north = { name: 'North', labels: chartLabels, values: [120, 180, 90, 210] };
const south = { name: 'South', labels: chartLabels, values: [80, 140, 160, 60] };
const west = { name: 'West', labels: chartLabels, values: [45, 95, 130, 175] };

const columnsSlide = charts.addSlide();
columnsSlide.addText('Column + Bar', { x: 0.4, y: 0.1, fontSize: 18, bold: true });
columnsSlide.addChart(charts.ChartType.bar, [north, south, west], {
  x: 0.4,
  y: 0.7,
  w: 6.1,
  h: 4.4,
  barDir: 'col',
  showLegend: true,
  legendPos: 'b',
  showTitle: true,
  title: 'Revenue by Quarter',
});
columnsSlide.addChart(charts.ChartType.bar, [north, south], {
  x: 6.8,
  y: 0.7,
  w: 6.1,
  h: 4.4,
  barDir: 'bar',
  showLegend: true,
  legendPos: 'r',
  chartColors: ['C0504D', '4F81BD'],
  showValue: true,
});

const stackedSlide = charts.addSlide();
stackedSlide.addText('Stacked + 100% stacked', { x: 0.4, y: 0.1, fontSize: 18, bold: true });
stackedSlide.addChart(charts.ChartType.bar, [north, south, west], {
  x: 0.4,
  y: 0.7,
  w: 6.1,
  h: 4.4,
  barDir: 'col',
  barGrouping: 'stacked',
  showLegend: true,
  legendPos: 'b',
});
stackedSlide.addChart(charts.ChartType.bar, [north, south, west], {
  x: 6.8,
  y: 0.7,
  w: 6.1,
  h: 4.4,
  barDir: 'col',
  barGrouping: 'percentStacked',
  showLegend: true,
  legendPos: 'b',
});

const pieSlide = charts.addSlide();
pieSlide.addText('Pie + Doughnut', { x: 0.4, y: 0.1, fontSize: 18, bold: true });
pieSlide.addChart(
  charts.ChartType.pie,
  [{ name: 'Share', labels: ['Alpha', 'Beta', 'Gamma', 'Delta'], values: [35, 25, 25, 15] }],
  { x: 0.4, y: 0.7, w: 6.1, h: 4.6, showLegend: true, legendPos: 'r', showPercent: true },
);
pieSlide.addChart(
  charts.ChartType.doughnut,
  [{ name: 'Share', labels: ['Alpha', 'Beta', 'Gamma', 'Delta'], values: [40, 30, 20, 10] }],
  { x: 6.8, y: 0.7, w: 6.1, h: 4.6, holeSize: 55, showLegend: true, legendPos: 'r' },
);

const lineSlide = charts.addSlide();
lineSlide.addText('Line + Area', { x: 0.4, y: 0.1, fontSize: 18, bold: true });
lineSlide.addChart(charts.ChartType.line, [north, south, west], {
  x: 0.4,
  y: 0.7,
  w: 6.1,
  h: 4.4,
  lineSmooth: true,
  lineSize: 2.5,
  lineDataSymbol: 'circle',
  lineDataSymbolSize: 7,
  showLegend: true,
  legendPos: 'b',
});
lineSlide.addChart(charts.ChartType.area, [north, south], {
  x: 6.8,
  y: 0.7,
  w: 6.1,
  h: 4.4,
  showLegend: true,
  legendPos: 'b',
});

const scatterSlide = charts.addSlide();
scatterSlide.addText('Scatter + Radar', { x: 0.4, y: 0.1, fontSize: 18, bold: true });
scatterSlide.addChart(
  charts.ChartType.scatter,
  [
    { name: 'X-Axis', values: [1, 2, 3, 4, 5, 6] },
    { name: 'Trial A', values: [3, 5, 4, 6, 8, 7] },
    { name: 'Trial B', values: [2, 4, 6, 5, 7, 9] },
  ],
  { x: 0.4, y: 0.7, w: 6.1, h: 4.4, showLegend: true, legendPos: 'b', lineSize: 0, lineDataSymbolSize: 8 },
);
scatterSlide.addChart(charts.ChartType.radar, [north, south], {
  x: 6.8,
  y: 0.7,
  w: 6.1,
  h: 4.4,
  radarStyle: 'marker',
  showLegend: true,
  legendPos: 'b',
});

const chartsFixturePath = path.join(outputDirectory, 'charts-showcase.pptx');
await charts.writeFile({ fileName: chartsFixturePath });
await copyFile(chartsFixturePath, path.join(playgroundPublic, 'charts-showcase.pptx'));
console.log(chartsFixturePath);
