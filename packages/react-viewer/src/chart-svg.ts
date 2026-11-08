/**
 * Mounts the chart artwork that the Rust parser renders with
 * `rust-office-charts` into a slide. The markup is parsed as XML and copied
 * into the host document through an allow-list, so only local paint servers,
 * inert geometry, text, and data-URI images reach the DOM. Ids are prefixed so
 * several charts on one page never collide.
 */
import type { ChartNode } from '@extend-ai/react-pptx-model';

const svgNamespace = 'http://www.w3.org/2000/svg';
const xlinkNamespace = 'http://www.w3.org/1999/xlink';
const maxMarkupBytes = 128 * 1024 * 1024;
const maxNestedDepth = 3;

const allowedElements = new Set([
  'svg', 'g', 'defs', 'clipPath', 'rect', 'circle', 'ellipse', 'path', 'line', 'polyline', 'polygon',
  'title', 'desc', 'text', 'tspan', 'linearGradient', 'radialGradient', 'stop', 'pattern', 'image',
  'filter', 'feGaussianBlur', 'feFlood', 'feComposite', 'feMorphology', 'feDiffuseLighting',
  'feSpecularLighting', 'feDistantLight', 'feDropShadow', 'feOffset', 'feBlend', 'feColorMatrix',
  'feComponentTransfer', 'feFuncR', 'feFuncG', 'feFuncB', 'feFuncA', 'use', 'mask', 'marker', 'symbol',
]);

const plainAttributes = new Set([
  'x', 'y', 'width', 'height', 'x1', 'x2', 'y1', 'y2', 'cx', 'cy', 'fx', 'fy', 'r', 'rx', 'ry', 'fr',
  'viewBox', 'preserveAspectRatio', 'transform', 'd', 'points', 'dx', 'dy', 'textLength', 'lengthAdjust',
  'font-family', 'font-weight', 'font-style', 'font-size', 'text-decoration', 'text-anchor',
  'dominant-baseline', 'letter-spacing', 'xml:space',
  'stroke-linecap', 'stroke-linejoin', 'stroke-width', 'stroke-dasharray', 'stroke-miterlimit',
  'opacity', 'fill-opacity', 'stroke-opacity', 'fill-rule', 'clip-rule', 'shape-rendering', 'overflow',
  'offset', 'stop-opacity', 'gradientUnits', 'gradientTransform', 'patternUnits', 'patternContentUnits',
  'patternTransform', 'maskUnits', 'maskContentUnits', 'markerUnits', 'markerWidth', 'markerHeight',
  'refX', 'refY', 'orient', 'color-interpolation', 'color-interpolation-filters',
  'type', 'mode', 'values', 'slope', 'intercept', 'tableValues', 'amplitude', 'exponent',
  'filterUnits', 'primitiveUnits', 'in', 'in2', 'result', 'stdDeviation', 'flood-opacity', 'operator',
  'radius', 'surfaceScale', 'diffuseConstant', 'specularConstant', 'specularExponent', 'azimuth',
  'elevation', 'k1', 'k2', 'k3', 'k4', 'pointer-events',
  'data-chart-kind', 'data-chart-series', 'data-chart-point', 'data-chart-x', 'data-chart-y',
  'data-chart-effect', 'data-chart-drawing', 'data-chart-overflow',
]);

const paintAttributes = new Set([
  'fill', 'stroke', 'stop-color', 'flood-color', 'lighting-color', 'clip-path', 'filter', 'mask',
  'marker-start', 'marker-mid', 'marker-end',
]);

const colorPattern = /^(?:#[\da-fA-F]{3,8}|[a-zA-Z]+|rgba?\(\s*[\d.%]+(?:\s*[, ]\s*[\d.%]+){2,3}\s*\))$/;
const rasterDataUrl = /^data:image\/(?:png|jpeg|gif|webp);base64,[A-Za-z0-9+/]+={0,2}$/;
const svgDataUrlPrefix = 'data:image/svg+xml;base64,';

let sequence = 0;

function isHref(name: string): boolean {
  return name === 'href' || name === 'xlink:href';
}

function decodeBase64(value: string): string {
  const bytes = Uint8Array.from(atob(value), (character) => character.charCodeAt(0));
  return new TextDecoder().decode(bytes);
}

function encodeBase64(value: string): string {
  const bytes = new TextEncoder().encode(value);
  const chunks: string[] = [];
  for (let offset = 0; offset < bytes.length; offset += 0x8000) {
    chunks.push(String.fromCharCode(...bytes.subarray(offset, offset + 0x8000)));
  }
  return btoa(chunks.join(''));
}

/**
 * Parses chart SVG markup and returns a detached, sanitized `<svg>` element,
 * or null when the markup is not an SVG document or uses a disallowed element.
 */
export function sanitizeChartSvg(markup: string, prefix: string, depth = 0): SVGSVGElement | null {
  if (depth > maxNestedDepth || markup.length > maxMarkupBytes) return null;
  const root = new DOMParser().parseFromString(markup, 'image/svg+xml').documentElement;
  if (root.localName !== 'svg' || root.namespaceURI !== svgNamespace) return null;
  const svg = root as unknown as SVGSVGElement;
  const ids = new Set(Array.from(svg.querySelectorAll('[id]'), (node) => node.id));
  let imageIndex = 0;

  for (const node of [svg, ...Array.from(svg.querySelectorAll('*'))]) {
    if (node.namespaceURI !== svgNamespace || !allowedElements.has(node.localName)) return null;
    for (const attribute of Array.from(node.attributes)) {
      const { name, value } = attribute;
      if (name === 'xmlns' || name === 'xmlns:xlink') continue;
      if (attribute.namespaceURI && attribute.namespaceURI !== xlinkNamespace && name !== 'xml:space') {
        node.removeAttributeNode(attribute);
        continue;
      }
      if (plainAttributes.has(name)) continue;
      if (name === 'id') {
        if (/^[\w.-]+$/.test(value)) node.setAttribute('id', prefix + value);
        else node.removeAttribute('id');
        continue;
      }
      if (isHref(name) && node.localName === 'image') {
        if (rasterDataUrl.test(value)) {
          if (name !== 'href') {
            node.removeAttributeNode(attribute);
            node.setAttribute('href', value);
          }
          continue;
        }
        if (value.startsWith(svgDataUrlPrefix)) {
          try {
            const nested = sanitizeChartSvg(
              decodeBase64(value.slice(svgDataUrlPrefix.length)),
              `${prefix}image${imageIndex++}-`,
              depth + 1,
            );
            if (!nested) return null;
            node.removeAttributeNode(attribute);
            node.setAttribute(
              'href',
              svgDataUrlPrefix + encodeBase64(new XMLSerializer().serializeToString(nested)),
            );
            continue;
          } catch {
            return null;
          }
        }
        node.removeAttributeNode(attribute);
        continue;
      }
      if (isHref(name) && /^(?:linearGradient|radialGradient|pattern|use)$/.test(node.localName)) {
        if (/^#[\w.-]+$/.test(value) && ids.has(value.slice(1))) {
          node.removeAttributeNode(attribute);
          node.setAttribute('href', `#${prefix}${value.slice(1)}`);
        } else {
          node.removeAttributeNode(attribute);
        }
        continue;
      }
      if (paintAttributes.has(name)) {
        const reference = /^url\(#([\w.-]+)\)$/.exec(value);
        if (reference && ids.has(reference[1]!)) {
          node.setAttribute(name, `url(#${prefix}${reference[1]})`);
          continue;
        }
        if (colorPattern.test(value)) continue;
      }
      node.removeAttributeNode(attribute);
    }
  }
  return svg;
}

/**
 * Marks the series and point geometry the renderer tagged so hosts can style
 * or select them, and gives tiny markers a hit area. Geometry inside `defs`,
 * effect groups, and drawing overlays stays inert.
 */
function annotateInteractiveGeometry(svg: SVGSVGElement): void {
  for (const element of Array.from(svg.querySelectorAll('[data-chart-series]'))) {
    if (element.closest('defs,[data-chart-effect],[data-chart-drawing]')) {
      element.removeAttribute('data-chart-series');
      element.removeAttribute('data-chart-point');
      element.removeAttribute('data-chart-kind');
      continue;
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
}

/**
 * Replaces the contents of `container` with the sanitized chart SVG.
 * Returns false when the markup cannot be mounted safely.
 */
export function mountChartSvg(container: HTMLElement, markup: string, label: string): boolean {
  const svg = sanitizeChartSvg(markup, `rpv-chart-${++sequence}-`);
  if (!svg) return false;
  annotateInteractiveGeometry(svg);
  const view = container.ownerDocument.importNode(svg, true) as unknown as SVGSVGElement;
  view.style.cssText = 'width:100%;height:100%;display:block;';
  view.setAttribute('role', 'img');
  view.setAttribute('aria-label', label);
  container.replaceChildren(view);
  return true;
}

/**
 * Renders a chart node's parser-produced artwork into `container`. Returns
 * false when the parser produced no SVG for this chart (callers then fall back
 * to a plain placeholder built from the cached series summary).
 */
export function renderChartInto(container: HTMLElement, node: ChartNode): boolean {
  if (!node.chartSvg) return false;
  try {
    return mountChartSvg(container, node.chartSvg, node.title ?? node.name ?? 'Chart');
  } catch {
    return false;
  }
}
