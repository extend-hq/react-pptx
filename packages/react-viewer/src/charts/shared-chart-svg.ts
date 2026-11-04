const svgNamespace = 'http://www.w3.org/2000/svg';
const elements = new Set(
  'svg g rect circle ellipse path polygon polyline line text tspan defs clipPath linearGradient radialGradient stop pattern filter feDropShadow'.split(
    ' ',
  ),
);
const attributes = new Set(
  (
    'width height viewBox x y x1 y1 x2 y2 cx cy rx ry r d points transform id dx dy ' +
    'fill fill-opacity fill-rule stroke stroke-opacity stroke-width stroke-dasharray stroke-linecap stroke-linejoin ' +
    'clip-path filter opacity font-family font-size font-weight font-style text-anchor textLength lengthAdjust ' +
    'offset stop-color stop-opacity gradientUnits gradientTransform fx fy fr patternUnits patternContentUnits ' +
    'shape-rendering stdDeviation flood-color flood-opacity overflow ' +
    'data-chart-kind data-chart-series data-chart-point data-chart-x data-chart-y data-docx-series data-docx-point'
  ).split(' '),
);
const paints = new Set(['fill', 'stroke', 'filter', 'clip-path', 'stop-color', 'flood-color']);
let sequence = 0;

export function mountSharedChart(container: HTMLElement, markup: string, label: string): boolean {
  if (markup.includes('<!') || markup.includes('<?')) return false;
  const svg = new DOMParser().parseFromString(markup, 'image/svg+xml').documentElement;
  if (svg.localName !== 'svg') return false;

  const prefix = `rpv-chart-${++sequence}-`;
  for (const element of [svg, ...Array.from(svg.querySelectorAll('*'))]) {
    if (element.namespaceURI !== svgNamespace || !elements.has(element.localName)) return false;
    for (const attribute of Array.from(element.attributes)) {
      const { name, value } = attribute;
      if (name === 'xmlns' && element === svg && value === svgNamespace) continue;
      if (attribute.namespaceURI || !attributes.has(name)) return false;
      if (name === 'id') {
        if (!/^[a-zA-Z_][\w.-]*$/.test(value)) return false;
        element.id = prefix + value;
      } else if (paints.has(name)) {
        const reference = /^url\(#([\w.-]+)\)$/.exec(value);
        if (reference) element.setAttribute(name, `url(#${prefix}${reference[1]})`);
        else if (!/^(?:#[\da-fA-F]{3,8}|[a-zA-Z]+|rgba?\([\d.% ,]+\))$/.test(value)) return false;
      }
    }
    const series = element.getAttribute('data-chart-series');
    if (series === null) continue;
    element.setAttribute('data-xlsx-chart-series-index', series);
    const point = element.getAttribute('data-chart-point');
    if (point !== null) element.setAttribute('data-xlsx-chart-point-index', point);
    const kind = element.getAttribute('data-chart-kind');
    if (kind === 'bar') element.setAttribute('data-xlsx-chart-selection-mode', 'seriesFirst');
    if (kind === 'legendEntry') element.setAttribute('data-xlsx-chart-element-kind', 'legendEntry');
    if (kind === 'point') {
      const hit = svg.ownerDocument.createElementNS(svgNamespace, 'circle');
      hit.setAttribute('cx', element.getAttribute('data-chart-x') ?? '0');
      hit.setAttribute('cy', element.getAttribute('data-chart-y') ?? '0');
      hit.setAttribute('r', '5');
      hit.setAttribute('fill', 'transparent');
      element.append(hit);
    }
    element.setAttribute('cursor', 'pointer');
  }
  const view = container.ownerDocument.importNode(svg, true) as unknown as SVGSVGElement;
  view.style.cssText = 'width:100%;height:100%;display:block;';
  view.setAttribute('role', 'img');
  view.setAttribute('aria-label', label);
  container.replaceChildren(view);
  return true;
}
