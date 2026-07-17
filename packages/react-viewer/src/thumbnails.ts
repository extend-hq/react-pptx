import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type {
  PptxSlideThumbnailItem,
  PptxSlideThumbnailResolution,
  PptxSlideThumbnailStatus,
  PptxViewerController,
  PptxViewerThumbnails,
  UsePptxViewerThumbnailsOptions,
} from './types';

const EMU_PER_CSS_PIXEL = 9_525;
const DEFAULT_MAX_DIMENSION = 160;

interface ThumbnailState {
  error?: Error;
  status: PptxSlideThumbnailStatus;
}

interface PrefetchedThumbnail {
  cleanup: () => void;
  element: HTMLElement;
  width: number;
}

function normalizeSlideIndexes(
  indexes: readonly number[] | undefined,
  slideCount: number,
  sort = true,
): number[] {
  if (!indexes?.length || slideCount <= 0) return [];
  const normalized = [...new Set(indexes)].filter(
    (index) => Number.isInteger(index) && index >= 0 && index < slideCount,
  );
  return sort ? normalized.sort((left, right) => left - right) : normalized;
}

function finitePositive(value: number | undefined): number | undefined {
  return value !== undefined && Number.isFinite(value) && value > 0 ? value : undefined;
}

function thumbnailBounds(resolution?: PptxSlideThumbnailResolution): {
  maxHeight: number;
  maxWidth: number;
} {
  if (typeof resolution === 'number') {
    const maximum = finitePositive(resolution) ?? DEFAULT_MAX_DIMENSION;
    return { maxHeight: maximum, maxWidth: maximum };
  }
  const maxHeight = finitePositive(resolution?.maxHeight);
  const maxWidth = finitePositive(resolution?.maxWidth);
  if (maxHeight === undefined && maxWidth === undefined) {
    return { maxHeight: DEFAULT_MAX_DIMENSION, maxWidth: DEFAULT_MAX_DIMENSION };
  }
  return {
    maxHeight: maxHeight ?? Number.POSITIVE_INFINITY,
    maxWidth: maxWidth ?? Number.POSITIVE_INFINITY,
  };
}

function thumbnailSize(
  sourceWidth: number,
  sourceHeight: number,
  resolution?: PptxSlideThumbnailResolution,
): { height: number; width: number } {
  const { maxHeight, maxWidth } = thumbnailBounds(resolution);
  const scale = Math.min(maxWidth / sourceWidth, maxHeight / sourceHeight);
  return {
    height: Math.max(1, sourceHeight * scale),
    width: Math.max(1, sourceWidth * scale),
  };
}

function asError(reason: unknown): Error {
  return reason instanceof Error ? reason : new Error(String(reason));
}

/**
 * Exposes detached slide previews for consumer-owned thumbnail rails.
 *
 * PPTX slides remain live DOM/SVG surfaces to preserve the main viewer's
 * rendering fidelity. Attach `thumbnail.containerRef` to a sized element.
 */
export function usePptxViewerThumbnails(
  controller: PptxViewerController | null | undefined,
  options: UsePptxViewerThumbnailsOptions = {},
): PptxViewerThumbnails {
  const presentation = controller?.getDocument() ?? null;
  const disabled = options.disabled ?? false;
  const resolution = options.resolution;
  const attachedElementsRef = useRef(new Map<number, HTMLElement>());
  const cleanupsRef = useRef(new Map<number, () => void>());
  const prefetchedRef = useRef(new Map<number, PrefetchedThumbnail>());
  const prefetchInFlightRef = useRef(new Map<number, Promise<void>>());
  const renderGenerationRef = useRef(new Map<number, number>());
  const refCallbacksRef = useRef(new Map<number, (element: HTMLElement | null) => void>());
  const attachRef = useRef<(index: number, element: HTMLElement | null) => void>(() => {});
  const mountedRef = useRef(true);
  const [states, setStates] = useState<Record<number, ThumbnailState>>({});

  const slideCount = presentation?.slides.length ?? 0;
  const visibleSlideIndexes = useMemo(
    () => normalizeSlideIndexes(options.renderWindow?.visibleSlideIndexes, slideCount),
    [options.renderWindow?.visibleSlideIndexes, slideCount],
  );
  const prefetchSlideIndexes = useMemo(
    () => normalizeSlideIndexes(options.renderWindow?.prefetchSlideIndexes, slideCount, false),
    [options.renderWindow?.prefetchSlideIndexes, slideCount],
  );
  const visibleSlideIndexesKey = visibleSlideIndexes.join(',');
  const prefetchSlideIndexesKey = prefetchSlideIndexes.join(',');
  const visibleSlideIndexSet = useMemo(
    () => new Set(visibleSlideIndexes),
    [visibleSlideIndexesKey],
  );
  const visibleSlideIndexSetRef = useRef(visibleSlideIndexSet);
  visibleSlideIndexSetRef.current = visibleSlideIndexSet;
  const requestedPrefetchSlideIndexSetRef = useRef(new Set(prefetchSlideIndexes));
  requestedPrefetchSlideIndexSetRef.current = new Set(prefetchSlideIndexes);

  const sourceWidth = Math.max(1, (presentation?.size.widthEmu ?? 1) / EMU_PER_CSS_PIXEL);
  const sourceHeight = Math.max(1, (presentation?.size.heightEmu ?? 1) / EMU_PER_CSS_PIXEL);
  const size = thumbnailSize(sourceWidth, sourceHeight, resolution);

  const setThumbnailState = useCallback((index: number, state: ThumbnailState) => {
    if (!mountedRef.current) return;
    setStates((current) => {
      const previous = current[index];
      if (previous?.status === state.status && previous.error === state.error) return current;
      return { ...current, [index]: state };
    });
  }, []);

  const invalidateThumbnail = useCallback((index: number) => {
    renderGenerationRef.current.set(index, (renderGenerationRef.current.get(index) ?? 0) + 1);
  }, []);

  const disposeAttachedThumbnail = useCallback((index: number) => {
    cleanupsRef.current.get(index)?.();
    cleanupsRef.current.delete(index);
  }, []);

  const disposePrefetchedThumbnail = useCallback((index: number) => {
    prefetchedRef.current.get(index)?.cleanup();
    prefetchedRef.current.delete(index);
  }, []);

  const disposeThumbnail = useCallback(
    (index: number) => {
      invalidateThumbnail(index);
      disposeAttachedThumbnail(index);
      disposePrefetchedThumbnail(index);
    },
    [disposeAttachedThumbnail, disposePrefetchedThumbnail, invalidateThumbnail],
  );

  const renderThumbnail = useCallback(
    async (slideIndex: number, element: HTMLElement): Promise<void> => {
      disposeThumbnail(slideIndex);
      attachedElementsRef.current.set(slideIndex, element);
      if (disabled || !controller?.isReady()) {
        setThumbnailState(slideIndex, { status: 'idle' });
        return;
      }
      const generation = (renderGenerationRef.current.get(slideIndex) ?? 0) + 1;
      renderGenerationRef.current.set(slideIndex, generation);
      setThumbnailState(slideIndex, { status: 'rendering' });
      try {
        const cleanup = await controller.renderThumbnail(slideIndex, element, {
          width: size.width,
        });
        if (
          !mountedRef.current ||
          renderGenerationRef.current.get(slideIndex) !== generation ||
          attachedElementsRef.current.get(slideIndex) !== element
        ) {
          cleanup();
          return;
        }
        cleanupsRef.current.set(slideIndex, cleanup);
        setThumbnailState(slideIndex, { status: 'ready' });
      } catch (reason: unknown) {
        if (!mountedRef.current || renderGenerationRef.current.get(slideIndex) !== generation) {
          return;
        }
        setThumbnailState(slideIndex, { error: asError(reason), status: 'error' });
      }
    },
    [controller, disabled, disposeThumbnail, setThumbnailState, size.width],
  );

  attachRef.current = (index, element) => {
    const previous = attachedElementsRef.current.get(index);
    if (previous === element) return;
    if (!element) {
      invalidateThumbnail(index);
      disposeAttachedThumbnail(index);
      attachedElementsRef.current.delete(index);
      setThumbnailState(index, { status: 'idle' });
      return;
    }

    invalidateThumbnail(index);
    disposeAttachedThumbnail(index);
    attachedElementsRef.current.set(index, element);
    const prefetched = prefetchedRef.current.get(index);
    if (prefetched && !disabled && controller?.isReady() && prefetched.width === size.width) {
      prefetchedRef.current.delete(index);
      element.replaceChildren(...Array.from(prefetched.element.childNodes));
      cleanupsRef.current.set(index, prefetched.cleanup);
      setThumbnailState(index, { status: 'ready' });
      return;
    }
    disposePrefetchedThumbnail(index);
    void renderThumbnail(index, element);
  };

  const refForIndex = useCallback((index: number) => {
    let callback = refCallbacksRef.current.get(index);
    if (!callback) {
      callback = (element) => attachRef.current(index, element);
      refCallbacksRef.current.set(index, callback);
    }
    return callback;
  }, []);

  const rerenderAttachedThumbnails = useCallback(async (): Promise<void> => {
    const attached = [...attachedElementsRef.current].sort(([leftIndex], [rightIndex]) => {
      const leftPriority = visibleSlideIndexSetRef.current.has(leftIndex) ? 0 : 1;
      const rightPriority = visibleSlideIndexSetRef.current.has(rightIndex) ? 0 : 1;
      return leftPriority - rightPriority || leftIndex - rightIndex;
    });
    await Promise.all(attached.map(([index, element]) => renderThumbnail(index, element)));
  }, [renderThumbnail]);

  const prefetchThumbnail = useCallback(
    async (slideIndex: number): Promise<void> => {
      while (true) {
        const existing = prefetchInFlightRef.current.get(slideIndex);
        if (existing) {
          await existing;
          continue;
        }
        if (
          disabled ||
          !controller?.isReady() ||
          attachedElementsRef.current.has(slideIndex) ||
          prefetchedRef.current.has(slideIndex) ||
          !requestedPrefetchSlideIndexSetRef.current.has(slideIndex)
        ) {
          return;
        }

        const generation = (renderGenerationRef.current.get(slideIndex) ?? 0) + 1;
        renderGenerationRef.current.set(slideIndex, generation);
        const element = globalThis.document.createElement('div');
        setThumbnailState(slideIndex, { status: 'rendering' });
        const pending = (async () => {
          try {
            const cleanup = await controller.renderThumbnail(slideIndex, element, {
              width: size.width,
            });
            if (
              !mountedRef.current ||
              renderGenerationRef.current.get(slideIndex) !== generation ||
              attachedElementsRef.current.has(slideIndex) ||
              !requestedPrefetchSlideIndexSetRef.current.has(slideIndex)
            ) {
              cleanup();
              return;
            }
            prefetchedRef.current.set(slideIndex, { cleanup, element, width: size.width });
            setThumbnailState(slideIndex, { status: 'ready' });
          } catch (reason: unknown) {
            if (!mountedRef.current || renderGenerationRef.current.get(slideIndex) !== generation) {
              return;
            }
            setThumbnailState(slideIndex, { error: asError(reason), status: 'error' });
          }
        })();
        prefetchInFlightRef.current.set(slideIndex, pending);
        try {
          await pending;
        } finally {
          if (prefetchInFlightRef.current.get(slideIndex) === pending) {
            prefetchInFlightRef.current.delete(slideIndex);
          }
        }
        return;
      }
    },
    [controller, disabled, setThumbnailState, size.width],
  );

  useEffect(() => {
    mountedRef.current = true;
    return () => {
      mountedRef.current = false;
      for (const index of renderGenerationRef.current.keys()) {
        renderGenerationRef.current.set(index, (renderGenerationRef.current.get(index) ?? 0) + 1);
      }
      for (const cleanup of cleanupsRef.current.values()) cleanup();
      cleanupsRef.current.clear();
      for (const prefetched of prefetchedRef.current.values()) prefetched.cleanup();
      prefetchedRef.current.clear();
      prefetchInFlightRef.current.clear();
    };
  }, []);

  useEffect(
    () => () => {
      for (const index of prefetchInFlightRef.current.keys()) invalidateThumbnail(index);
      // The invalidated promises may belong to a previous controller or
      // resolution and must not gate replacement work for the same slide.
      prefetchInFlightRef.current.clear();
      for (const prefetched of prefetchedRef.current.values()) prefetched.cleanup();
      prefetchedRef.current.clear();
    },
    [controller, invalidateThumbnail, size.width],
  );

  useEffect(() => {
    void rerenderAttachedThumbnails();
  }, [rerenderAttachedThumbnails]);

  useEffect(() => {
    for (const index of [...prefetchedRef.current.keys()]) {
      if (disabled || !requestedPrefetchSlideIndexSetRef.current.has(index)) {
        disposePrefetchedThumbnail(index);
        if (!attachedElementsRef.current.has(index)) setThumbnailState(index, { status: 'idle' });
      }
    }
    for (const index of prefetchInFlightRef.current.keys()) {
      if (
        disabled ||
        attachedElementsRef.current.has(index) ||
        !requestedPrefetchSlideIndexSetRef.current.has(index)
      ) {
        invalidateThumbnail(index);
        prefetchInFlightRef.current.delete(index);
        if (!attachedElementsRef.current.has(index)) setThumbnailState(index, { status: 'idle' });
      }
    }
    if (disabled) return;

    let cancelled = false;
    void (async () => {
      for (const index of prefetchSlideIndexes) {
        if (cancelled) return;
        await prefetchThumbnail(index);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [
    disabled,
    disposePrefetchedThumbnail,
    invalidateThumbnail,
    prefetchSlideIndexesKey,
    prefetchThumbnail,
    setThumbnailState,
  ]);

  const thumbnails = useMemo<PptxSlideThumbnailItem[]>(
    () =>
      (presentation?.slides ?? []).map((slide, slideIndex) => {
        const state = states[slideIndex] ?? { status: 'idle' as const };
        return {
          aspectRatio: sourceWidth / sourceHeight,
          contentHeight: sourceHeight,
          contentWidth: sourceWidth,
          containerRef: refForIndex(slideIndex),
          ...(state.error ? { error: state.error } : {}),
          height: size.height,
          slide,
          slideIndex,
          slideNumber: slideIndex + 1,
          status: state.status,
          renderToContainer: (element: HTMLElement) => renderThumbnail(slideIndex, element),
          width: size.width,
        };
      }),
    [
      presentation,
      refForIndex,
      renderThumbnail,
      size.height,
      size.width,
      sourceHeight,
      sourceWidth,
      states,
    ],
  );

  return useMemo(
    () => ({ renderThumbnail, rerenderAttachedThumbnails, thumbnails }),
    [renderThumbnail, rerenderAttachedThumbnails, thumbnails],
  );
}
