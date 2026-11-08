import { describe, expect, it } from 'vitest';
import type { ChartNode } from '@extend-ai/react-pptx-model';
import { mountChartSvg, renderChartInto, sanitizeChartSvg } from './chart-svg';

const PNG_PIXEL =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';

function chartNode(chartSvg: string | undefined): ChartNode {
  return {
    id: 'chart',
    type: 'chart',
    transform: { x: 0, y: 0, width: 6_000_000, height: 4_000_000 },
    chartType: 'bar',
    series: [],
    ...(chartSvg === undefined ? {} : { chartSvg }),
  };
}

describe('sanitizeChartSvg', () => {
  it('keeps picture fills, descriptions, markers, and filters from the Rust renderer', () => {
    const markup = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">
      <desc>A chart drawing element is unsupported.</desc>
      <defs>
        <clipPath id="ci1"><rect x="0" y="0" width="50" height="50"/></clipPath>
        <pattern id="cp1" width="1" height="1" patternContentUnits="objectBoundingBox">
          <g clip-path="url(#ci1)"><image x="0" y="0" width="1" height="1" href="${PNG_PIXEL}" preserveAspectRatio="none"/></g>
        </pattern>
        <filter id="shadow"><feDropShadow dx="1" dy="1" stdDeviation="2" flood-color="#000000" flood-opacity="0.4"/></filter>
        <marker id="dot" markerWidth="4" markerHeight="4" refX="2" refY="2"><circle cx="2" cy="2" r="2" fill="#ff0000"/></marker>
      </defs>
      <rect data-chart-kind="bar" data-chart-series="0" data-chart-point="1" x="10" y="10" width="20" height="40" fill="url(#cp1)" filter="url(#shadow)"/>
      <path d="M0 0L10 10" stroke="#4472c4" marker-end="url(#dot)"/>
    </svg>`;

    const svg = sanitizeChartSvg(markup, 'p-');
    expect(svg).not.toBeNull();
    expect(svg!.querySelector('desc')?.textContent).toContain('unsupported');
    expect(svg!.querySelector('image')?.getAttribute('href')).toBe(PNG_PIXEL);
    expect(svg!.querySelector('pattern')?.id).toBe('p-cp1');
    expect(svg!.querySelector('rect[data-chart-kind="bar"]')?.getAttribute('fill')).toBe('url(#p-cp1)');
    expect(svg!.querySelector('rect[data-chart-kind="bar"]')?.getAttribute('filter')).toBe('url(#p-shadow)');
    expect(svg!.querySelector('path')?.getAttribute('marker-end')).toBe('url(#p-dot)');
    expect(svg!.querySelector('g')?.getAttribute('clip-path')).toBe('url(#p-ci1)');
  });

  it('rejects scripts, foreign elements, and non-SVG documents', () => {
    expect(sanitizeChartSvg('<svg xmlns="http://www.w3.org/2000/svg"><script>1</script></svg>', 'p-')).toBeNull();
    expect(sanitizeChartSvg('<svg xmlns="http://www.w3.org/2000/svg"><style>text{fill:red}</style></svg>', 'p-')).toBeNull();
    expect(
      sanitizeChartSvg('<svg xmlns="http://www.w3.org/2000/svg"><foreignObject><div/></foreignObject></svg>', 'p-'),
    ).toBeNull();
    expect(sanitizeChartSvg('<div>not svg</div>', 'p-')).toBeNull();
    expect(sanitizeChartSvg('<svg xmlns="http://www.w3.org/2000/svg"><rect', 'p-')).toBeNull();
  });

  it('strips event handlers, remote references, and unknown paint servers', () => {
    const markup = `<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" onload="alert(1)">
      <image href="https://evil.example/pixel.png" width="1" height="1"/>
      <image xlink:href="javascript:alert(1)" width="1" height="1"/>
      <rect fill="url(https://evil.example/paint.svg#p)" stroke="url(#missing)" width="1" height="1" onclick="alert(1)"/>
      <linearGradient id="g" href="#other"/>
      <use href="https://evil.example/sprite.svg#icon"/>
    </svg>`;

    const svg = sanitizeChartSvg(markup, 'p-');
    expect(svg).not.toBeNull();
    const serialized = new XMLSerializer().serializeToString(svg!);
    expect(serialized).not.toMatch(/evil\.example|javascript:|onload|onclick|url\(#missing\)/);
    expect(svg!.querySelector('linearGradient')?.hasAttribute('href')).toBe(false);
  });

  it('sanitizes nested SVG data URIs recursively', () => {
    const inner = '<svg xmlns="http://www.w3.org/2000/svg"><rect id="r" width="1" height="1" onclick="alert(1)"/></svg>';
    const malicious = '<svg xmlns="http://www.w3.org/2000/svg"><script>1</script></svg>';
    const encode = (value: string) => `data:image/svg+xml;base64,${btoa(value)}`;

    const good = sanitizeChartSvg(
      `<svg xmlns="http://www.w3.org/2000/svg"><image href="${encode(inner)}" width="1" height="1"/></svg>`,
      'p-',
    );
    const href = good!.querySelector('image')!.getAttribute('href')!;
    const decoded = atob(href.slice('data:image/svg+xml;base64,'.length));
    expect(decoded).toContain('id="p-image0-r"');
    expect(decoded).not.toContain('onclick');

    expect(
      sanitizeChartSvg(
        `<svg xmlns="http://www.w3.org/2000/svg"><image href="${encode(malicious)}" width="1" height="1"/></svg>`,
        'p-',
      ),
    ).toBeNull();
  });
});

describe('mountChartSvg', () => {
  it('mounts sanitized markup, prefixes ids per chart, and tags selectable geometry', () => {
    const markup = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 10 10">
      <defs><linearGradient id="g"><stop offset="0" stop-color="#fff"/></linearGradient></defs>
      <rect data-chart-kind="bar" data-chart-series="1" data-chart-point="2" fill="url(#g)" width="1" height="1"/>
      <g data-chart-effect="shadow"><rect data-chart-series="1" width="1" height="1"/></g>
      <path data-chart-kind="point" data-chart-series="0" data-chart-point="0" data-chart-x="4" data-chart-y="5" d="M4 5h1"/>
    </svg>`;
    const first = document.createElement('div');
    const second = document.createElement('div');

    expect(mountChartSvg(first, markup, 'Sales')).toBe(true);
    expect(mountChartSvg(second, markup, 'Sales')).toBe(true);

    const firstGradient = first.querySelector('linearGradient')!.id;
    const secondGradient = second.querySelector('linearGradient')!.id;
    expect(firstGradient).not.toBe(secondGradient);
    expect(first.querySelector('rect[data-chart-kind="bar"]')?.getAttribute('fill')).toBe(`url(#${firstGradient})`);

    const svg = first.querySelector('svg')!;
    expect(svg.getAttribute('role')).toBe('img');
    expect(svg.getAttribute('aria-label')).toBe('Sales');
    const bar = first.querySelector('rect[data-chart-kind="bar"]')!;
    expect(bar.getAttribute('data-xlsx-chart-series-index')).toBe('1');
    expect(bar.getAttribute('data-xlsx-chart-point-index')).toBe('2');
    expect(bar.getAttribute('data-xlsx-chart-selection-mode')).toBe('seriesFirst');
    expect(first.querySelector('[data-chart-effect] rect')?.hasAttribute('data-chart-series')).toBe(false);
    const hit = first.querySelector('path[data-chart-kind="point"] circle');
    expect(hit?.getAttribute('cx')).toBe('4');
    expect(hit?.getAttribute('r')).toBe('5');
  });
});

describe('renderChartInto', () => {
  it('mounts parser SVG and reports false when the parser produced none', () => {
    const container = document.createElement('div');
    expect(renderChartInto(container, chartNode(undefined))).toBe(false);
    expect(container.childElementCount).toBe(0);

    expect(
      renderChartInto(
        container,
        chartNode('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1 1"><rect width="1" height="1"/></svg>'),
      ),
    ).toBe(true);
    expect(container.querySelector('svg rect')).not.toBeNull();

    expect(renderChartInto(container, chartNode('<svg xmlns="http://www.w3.org/2000/svg"><script/></svg>'))).toBe(false);
  });
});
