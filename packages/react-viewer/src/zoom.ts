import type { ViewerZoomLevel } from './types';

export const MIN_VIEWER_ZOOM = 10;
export const MAX_VIEWER_ZOOM = 400;

export function normalizeViewerZoomLevel(level: ViewerZoomLevel): ViewerZoomLevel {
  if (typeof level !== 'number') return level;
  if (Number.isNaN(level)) return 100;
  return Math.max(MIN_VIEWER_ZOOM, Math.min(MAX_VIEWER_ZOOM, level));
}

export function resolveViewerZoom(
  level: ViewerZoomLevel,
  viewportWidth: number,
  viewportHeight: number,
  contentWidth: number,
  contentHeight: number,
): number {
  const normalized = normalizeViewerZoomLevel(level);
  if (typeof normalized === 'number') return normalized;

  const safeContentWidth = Number.isFinite(contentWidth) && contentWidth > 0 ? contentWidth : 1;
  const safeContentHeight = Number.isFinite(contentHeight) && contentHeight > 0 ? contentHeight : 1;
  const fitWidth =
    ((Number.isFinite(viewportWidth) && viewportWidth > 0 ? viewportWidth : safeContentWidth) /
      safeContentWidth) *
    100;
  const fitHeight =
    ((Number.isFinite(viewportHeight) && viewportHeight > 0 ? viewportHeight : safeContentHeight) /
      safeContentHeight) *
    100;

  const resolved =
    normalized === 'fit-width'
      ? fitWidth
      : normalized === 'fit-page'
        ? Math.min(fitWidth, fitHeight)
        : Math.min(fitWidth, 100);
  return Math.max(MIN_VIEWER_ZOOM, Math.min(MAX_VIEWER_ZOOM, resolved));
}
