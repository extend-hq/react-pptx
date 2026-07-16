/**
 * Bridges the DOM-based normalized viewer and the React chart renderer
 * vendored from react-xlsx. Charts are rendered to static SVG markup
 * synchronously, which keeps slide construction deterministic and avoids
 * nesting live React roots inside host applications' render cycles.
 */
import { createElement } from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import type { ChartNode, PresentationTheme } from '@extend-ai/react-pptx-model';
import { buildPptxChart, buildThemePaletteFromPresentationTheme } from './chart-model';
import { MemoChartSvg } from './chart-renderer';
import { getRegionMapAtlas, loadRegionMapAtlas } from './region-map-data';
import type { XlsxChart } from './chart-types';

/** Neutral light palette matching react-xlsx viewer chrome defaults. */
export const CHART_HOST_PALETTE = {
  border: '#e4e4e7',
  mutedText: '#71717a',
  surface: '#ffffff',
  text: '#18181b',
} as const;

/**
 * Renders a PowerPoint chart node into `container` with react-xlsx fidelity.
 * Returns false when the chart model cannot be built or rendered (callers
 * then fall back to a plain placeholder).
 */
function chartNeedsRegionMapAtlas(chart: XlsxChart): boolean {
  return (
    chart.chartType === 'RegionMap' ||
    (chart.typeGroups ?? []).some((group) => group.chartType === 'RegionMap')
  );
}

export function renderChartInto(
  container: HTMLElement,
  node: ChartNode,
  theme: PresentationTheme | undefined,
): boolean {
  try {
    const themePalette = buildThemePaletteFromPresentationTheme(theme);
    const chart = buildPptxChart(node, themePalette);
    const render = () => {
      const element = createElement(MemoChartSvg, {
        chart,
        palette: CHART_HOST_PALETTE,
        rect: {
          left: 0,
          top: 0,
          width: Math.max(1, node.transform.width / 9525),
          height: Math.max(1, node.transform.height / 9525),
        },
      });
      return renderToStaticMarkup(element);
    };
    const markup = render();
    if (!markup) return false;
    container.innerHTML = markup;
    // Excel map charts need the TopoJSON boundary atlas, which lives in a
    // lazily loaded chunk so ordinary decks never pay for it. Re-render the
    // chart in place once the geometry arrives.
    if (chartNeedsRegionMapAtlas(chart) && !getRegionMapAtlas()) {
      void loadRegionMapAtlas()
        .then(() => {
          if (!container.isConnected) return;
          const refreshed = render();
          if (refreshed) container.innerHTML = refreshed;
        })
        .catch(() => undefined);
    }
    return true;
  } catch {
    return false;
  }
}
