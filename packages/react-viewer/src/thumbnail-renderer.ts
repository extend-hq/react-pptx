import { PptxViewerError, throwIfAborted } from './errors';
import { PptxFontManager } from './fonts';
import { NormalizedPresentationViewer } from './normalized-viewer';
import { performanceMeasurement, performanceNow } from './performance';
import type {
  ParsedPresentation,
  PptxPerformanceMeasurement,
  PptxThumbnailRenderOptions,
  PptxThumbnailRenderer,
  PptxThumbnailRendererOptions,
  PptxThumbnailRenderResult,
} from './types';

const EMU_PER_CSS_PIXEL = 9_525;
const DEFAULT_MAX_DIMENSION = 160;

type ThumbnailData = Blob | HTMLCanvasElement | ImageBitmap | string;

interface QueueWaiter {
  reject: (reason: unknown) => void;
  resolve: (release: () => void) => void;
  signal?: AbortSignal;
  abort?: () => void;
}

class RenderQueue {
  private active = 0;
  private destroyed = false;
  private readonly waiters: QueueWaiter[] = [];

  constructor(private readonly limit: number) {}

  acquire(signal?: AbortSignal): Promise<() => void> {
    throwIfAborted(signal);
    if (this.destroyed) return Promise.reject(abortedError());
    return new Promise((resolve, reject) => {
      const waiter: QueueWaiter = { resolve, reject, ...(signal ? { signal } : {}) };
      if (signal) {
        waiter.abort = () => {
          const index = this.waiters.indexOf(waiter);
          if (index >= 0) this.waiters.splice(index, 1);
          reject(abortedError(signal.reason));
        };
        signal.addEventListener('abort', waiter.abort, { once: true });
      }
      this.waiters.push(waiter);
      this.drain();
    });
  }

  destroy(): void {
    this.destroyed = true;
    for (const waiter of this.waiters.splice(0)) {
      if (waiter.abort && waiter.signal) {
        waiter.signal.removeEventListener('abort', waiter.abort);
      }
      waiter.reject(abortedError());
    }
  }

  private drain(): void {
    while (!this.destroyed && this.active < this.limit && this.waiters.length) {
      const waiter = this.waiters.shift()!;
      if (waiter.abort && waiter.signal) {
        waiter.signal.removeEventListener('abort', waiter.abort);
      }
      if (waiter.signal?.aborted) {
        waiter.reject(abortedError(waiter.signal.reason));
        continue;
      }
      this.active += 1;
      let released = false;
      waiter.resolve(() => {
        if (released) return;
        released = true;
        this.active -= 1;
        this.drain();
      });
    }
  }
}

function abortedError(reason?: unknown): PptxViewerError {
  return new PptxViewerError('aborted', 'Thumbnail rendering was cancelled.', { cause: reason });
}

function finitePositive(value: number | undefined, fallback: number): number {
  return value !== undefined && Number.isFinite(value) && value > 0 ? value : fallback;
}

function outputSize(
  presentation: ParsedPresentation,
  options: PptxThumbnailRenderOptions,
): { height: number; naturalHeight: number; naturalWidth: number; width: number } {
  const naturalWidth = Math.max(1, presentation.document.size.widthEmu / EMU_PER_CSS_PIXEL);
  const naturalHeight = Math.max(1, presentation.document.size.heightEmu / EMU_PER_CSS_PIXEL);
  const hasWidth =
    options.maxWidth !== undefined && Number.isFinite(options.maxWidth) && options.maxWidth > 0;
  const hasHeight =
    options.maxHeight !== undefined && Number.isFinite(options.maxHeight) && options.maxHeight > 0;
  const maxWidth = hasWidth
    ? options.maxWidth!
    : hasHeight
      ? Number.POSITIVE_INFINITY
      : DEFAULT_MAX_DIMENSION;
  const maxHeight = hasHeight
    ? options.maxHeight!
    : hasWidth
      ? Number.POSITIVE_INFINITY
      : DEFAULT_MAX_DIMENSION;
  const scale = Math.min(maxWidth / naturalWidth, maxHeight / naturalHeight);
  return {
    height: Math.max(1, Math.round(naturalHeight * scale)),
    naturalHeight,
    naturalWidth,
    width: Math.max(1, Math.round(naturalWidth * scale)),
  };
}

function blobToDataUrl(blob: Blob, signal?: AbortSignal): Promise<string> {
  return abortable(
    new Promise<string>((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(String(reader.result));
      reader.onerror = () => reject(reader.error ?? new Error('Could not read a slide asset.'));
      reader.readAsDataURL(blob);
    }),
    signal,
  );
}

async function inlineResourceUrl(url: string, signal?: AbortSignal): Promise<string> {
  if (!/^(?:blob:|https?:)/i.test(url)) return url;
  try {
    const response = await fetch(url, signal ? { signal } : undefined);
    if (!response.ok) return url;
    return await blobToDataUrl(await response.blob(), signal);
  } catch {
    throwIfAborted(signal);
    return url;
  }
}

async function inlineResources(root: HTMLElement, signal?: AbortSignal): Promise<void> {
  const cache = new Map<string, Promise<string>>();
  const inline = (url: string): Promise<string> => {
    let pending = cache.get(url);
    if (!pending) {
      pending = inlineResourceUrl(url, signal);
      cache.set(url, pending);
    }
    return pending;
  };
  const elements = [root, ...root.querySelectorAll<HTMLElement | SVGElement>('*')];
  await Promise.all(
    elements.flatMap((element) => {
      const attributeNames = ['src', 'poster', ...(element.localName === 'image' ? ['href'] : [])];
      const attributes = attributeNames
        .map((name) => ({ name, value: element.getAttribute(name) }))
        .filter((entry): entry is { name: string; value: string } => Boolean(entry.value));
      const attributeTasks = attributes.map(async ({ name, value }) => {
        element.setAttribute(name, await inline(value));
      });
      const styleTasks = Array.from({ length: element.style.length }, (_, index) =>
        element.style.item(index),
      )
        .map((property) => ({ property, value: element.style.getPropertyValue(property) }))
        .filter(({ value }) => value.includes('url('))
        .map(async ({ property, value }) => {
          const urls = [...value.matchAll(/url\(["']?([^"')]+)["']?\)/g)].map((match) => match[1]!);
          let next = value;
          for (const url of urls) next = next.replace(url, await inline(url));
          element.style.setProperty(property, next);
        });
      return [...attributeTasks, ...styleTasks];
    }),
  );
}

async function serializeSlide(
  slide: HTMLElement,
  naturalWidth: number,
  naturalHeight: number,
  width: number,
  height: number,
  signal?: AbortSignal,
): Promise<string> {
  const clone = slide.cloneNode(true) as HTMLElement;
  clone.style.transform = 'none';
  clone.style.transformOrigin = 'top left';
  clone.setAttribute('xmlns', 'http://www.w3.org/1999/xhtml');

  const sourceCanvases = slide.querySelectorAll('canvas');
  const clonedCanvases = clone.querySelectorAll('canvas');
  sourceCanvases.forEach((canvas, index) => {
    const cloned = clonedCanvases[index];
    if (!cloned) return;
    try {
      const image = document.createElement('img');
      image.src = canvas.toDataURL('image/png');
      image.setAttribute('style', cloned.getAttribute('style') ?? '');
      cloned.replaceWith(image);
    } catch {
      cloned.remove();
    }
  });

  await inlineResources(clone, signal);

  const content = new XMLSerializer().serializeToString(clone);
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${naturalWidth} ${naturalHeight}"><foreignObject width="${naturalWidth}" height="${naturalHeight}">${content}</foreignObject></svg>`;
}

async function abortable<T>(promise: Promise<T>, signal?: AbortSignal): Promise<T> {
  throwIfAborted(signal);
  if (!signal) return promise;
  return new Promise<T>((resolve, reject) => {
    const abort = () => reject(abortedError(signal.reason));
    signal.addEventListener('abort', abort, { once: true });
    void promise.then(
      (value) => {
        signal.removeEventListener('abort', abort);
        resolve(value);
      },
      (error: unknown) => {
        signal.removeEventListener('abort', abort);
        reject(error);
      },
    );
  });
}

async function decodeElementImages(root: HTMLElement, signal?: AbortSignal): Promise<void> {
  const images = [...root.querySelectorAll<HTMLImageElement>('img')];
  const svgImages = [...root.querySelectorAll<SVGImageElement>('image')]
    .map((image) => image.href.baseVal || image.getAttribute('href') || '')
    .filter(Boolean);
  await abortable(
    Promise.all([
      ...images.map(async (image) => {
        if (typeof image.decode === 'function') await image.decode().catch(() => undefined);
      }),
      ...svgImages.map(
        (source) =>
          new Promise<void>((resolve) => {
            const image = new Image();
            image.onload = () => resolve();
            image.onerror = () => resolve();
            image.src = source;
          }),
      ),
    ]).then(() => undefined),
    signal,
  );
}

async function decodeSvg(svg: string, signal?: AbortSignal): Promise<HTMLImageElement> {
  return abortable(
    new Promise<HTMLImageElement>((resolve, reject) => {
      const image = new Image();
      image.onload = () => resolve(image);
      image.onerror = () => reject(new Error('The rendered slide SVG could not be decoded.'));
      image.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
    }),
    signal,
  );
}

function canvasToBlob(canvas: HTMLCanvasElement, quality?: number): Promise<Blob> {
  return new Promise((resolve, reject) => {
    canvas.toBlob(
      (blob) => {
        if (blob) resolve(blob);
        else reject(new Error('The browser could not encode the thumbnail.'));
      },
      'image/png',
      quality,
    );
  });
}

class ThumbnailRenderer implements PptxThumbnailRenderer {
  readonly ready: Promise<void>;
  private readonly fontManager: PptxFontManager;
  private readonly queue: RenderQueue;
  private readonly root: HTMLDivElement;
  private readonly viewer: NormalizedPresentationViewer;
  private destroyed = false;

  constructor(
    readonly presentation: ParsedPresentation,
    private readonly options: PptxThumbnailRendererOptions,
  ) {
    if (typeof document === 'undefined') {
      throw new PptxViewerError(
        'unsupported-environment',
        'Thumbnail rendering requires a browser DOM.',
      );
    }
    const concurrency = Math.max(1, Math.floor(finitePositive(options.concurrency, 2)));
    this.queue = new RenderQueue(concurrency);
    this.root = document.createElement('div');
    this.root.setAttribute('aria-hidden', 'true');
    this.root.style.position = 'fixed';
    this.root.style.left = '-100000px';
    this.root.style.top = '0';
    this.root.style.pointerEvents = 'none';
    this.root.style.opacity = '0';
    document.body.append(this.root);
    this.fontManager = new PptxFontManager(
      {
        loadEmbeddedFonts: false,
        reportMissingFonts: false,
        waitForFonts: false,
        ...options.fonts,
      },
      (warning) => options.onWarning?.(warning),
    );
    this.viewer = new NormalizedPresentationViewer(this.root, presentation.document, {
      onNodeError: (nodeId, error) =>
        options.onWarning?.({
          code: 'degraded-rendering',
          severity: 'warning',
          nodeId,
          feature: 'thumbnail-renderer',
          message: `A slide element could not be rendered: ${
            error instanceof Error ? error.message : String(error)
          }`,
        }),
    });
    const startedAt = performanceNow();
    this.ready = this.fontManager.prepare(presentation.document).then(() => {
      this.report(performanceMeasurement('embedded-font-preparation', startedAt));
    });
  }

  renderSlide(
    slideIndex: number,
    options: PptxThumbnailRenderOptions & { output: 'canvas' },
  ): Promise<PptxThumbnailRenderResult<HTMLCanvasElement>>;
  renderSlide(
    slideIndex: number,
    options: PptxThumbnailRenderOptions & { output: 'imageBitmap' },
  ): Promise<PptxThumbnailRenderResult<ImageBitmap>>;
  renderSlide(
    slideIndex: number,
    options: PptxThumbnailRenderOptions & { output: 'svg' },
  ): Promise<PptxThumbnailRenderResult<string>>;
  renderSlide(
    slideIndex: number,
    options?: PptxThumbnailRenderOptions & { output?: 'blob' },
  ): Promise<PptxThumbnailRenderResult<Blob>>;
  renderSlide(
    slideIndex: number,
    options: PptxThumbnailRenderOptions,
  ): Promise<PptxThumbnailRenderResult<ThumbnailData>>;
  async renderSlide(
    slideIndex: number,
    options: PptxThumbnailRenderOptions = {},
  ): Promise<PptxThumbnailRenderResult<ThumbnailData>> {
    const release = await this.queue.acquire(options.signal);
    try {
      this.assertAvailable(options.signal);
      await abortable(this.ready, options.signal);
      this.assertAvailable(options.signal);
      if (
        !Number.isInteger(slideIndex) ||
        slideIndex < 0 ||
        slideIndex >= this.presentation.document.slides.length
      ) {
        throw new RangeError(`Slide ${slideIndex} does not exist.`);
      }

      const measurements: PptxPerformanceMeasurement[] = [];
      const size = outputSize(this.presentation, options);
      const target = document.createElement('div');
      this.root.append(target);
      const domStartedAt = performanceNow();
      const handle = this.viewer.renderThumbnailToContainer(slideIndex, target, {
        width: size.naturalWidth,
      });
      try {
        await abortable(handle.ready, options.signal);
        this.fontManager.applyTo(handle.element, slideIndex);
        const domDurationMs = performanceNow() - domStartedAt;

        const decodeStartedAt = performanceNow();
        await decodeElementImages(handle.element, options.signal);
        let decodeDurationMs = performanceNow() - decodeStartedAt;
        const serializationStartedAt = performanceNow();
        const svg = await serializeSlide(
          handle.element,
          size.naturalWidth,
          size.naturalHeight,
          size.width,
          size.height,
          options.signal,
        );
        this.record(
          {
            durationMs: Math.max(0, domDurationMs + performanceNow() - serializationStartedAt),
            phase: 'slide-dom-svg-rendering',
            slideIndex,
          },
          measurements,
        );
        const output = options.output ?? 'blob';
        if (output === 'svg') {
          this.record(
            { durationMs: Math.max(0, decodeDurationMs), phase: 'image-decoding', slideIndex },
            measurements,
          );
          return this.result(svg, output, slideIndex, size, measurements);
        }

        const svgDecodeStartedAt = performanceNow();
        const image = await decodeSvg(svg, options.signal);
        decodeDurationMs += performanceNow() - svgDecodeStartedAt;
        this.record(
          { durationMs: Math.max(0, decodeDurationMs), phase: 'image-decoding', slideIndex },
          measurements,
        );
        this.assertAvailable(options.signal);
        const encodeStartedAt = performanceNow();
        const pixelRatio = finitePositive(options.pixelRatio, 1);
        const canvas = document.createElement('canvas');
        canvas.width = Math.max(1, Math.round(size.width * pixelRatio));
        canvas.height = Math.max(1, Math.round(size.height * pixelRatio));
        canvas.style.width = `${size.width}px`;
        canvas.style.height = `${size.height}px`;
        const context = canvas.getContext('2d');
        if (!context) throw new Error('The browser does not provide a 2D canvas renderer.');
        context.drawImage(image, 0, 0, canvas.width, canvas.height);

        let data: Blob | HTMLCanvasElement | ImageBitmap;
        if (output === 'canvas') data = canvas;
        else if (output === 'imageBitmap') {
          if (typeof createImageBitmap !== 'function') {
            throw new PptxViewerError(
              'unsupported-environment',
              'ImageBitmap output is not supported by this browser.',
            );
          }
          data = await abortable(createImageBitmap(canvas), options.signal);
        } else data = await abortable(canvasToBlob(canvas, options.quality), options.signal);
        this.measure('raster-encoding', encodeStartedAt, slideIndex, measurements);
        return this.result(data, output, slideIndex, size, measurements);
      } finally {
        handle.dispose();
        target.remove();
      }
    } finally {
      release();
    }
  }

  renderSlides(
    slideIndexes: readonly number[],
    options: PptxThumbnailRenderOptions = {},
  ): Promise<PptxThumbnailRenderResult<ThumbnailData>[]> {
    return Promise.all(slideIndexes.map((slideIndex) => this.renderDynamic(slideIndex, options)));
  }

  destroy(): void {
    if (this.destroyed) return;
    this.destroyed = true;
    this.queue.destroy();
    this.viewer.destroy();
    this.fontManager.destroy();
    this.root.remove();
  }

  private renderDynamic(
    slideIndex: number,
    options: PptxThumbnailRenderOptions,
  ): Promise<PptxThumbnailRenderResult<ThumbnailData>> {
    return this.renderSlide(slideIndex, options);
  }

  private assertAvailable(signal?: AbortSignal): void {
    throwIfAborted(signal);
    if (this.destroyed) throw abortedError();
  }

  private measure(
    phase: PptxPerformanceMeasurement['phase'],
    startedAt: number,
    slideIndex: number,
    measurements: PptxPerformanceMeasurement[],
  ): void {
    this.record(performanceMeasurement(phase, startedAt, slideIndex), measurements);
  }

  private record(
    measurement: PptxPerformanceMeasurement,
    measurements: PptxPerformanceMeasurement[],
  ): void {
    measurements.push(measurement);
    this.report(measurement);
  }

  private report(measurement: PptxPerformanceMeasurement): void {
    this.options.onPerformanceMeasurement?.(measurement);
  }

  private result<T extends ThumbnailData>(
    data: T,
    output: PptxThumbnailRenderResult<T>['output'],
    slideIndex: number,
    size: { height: number; width: number },
    measurements: PptxPerformanceMeasurement[],
  ): PptxThumbnailRenderResult<T> {
    return {
      data,
      height: size.height,
      measurements,
      output,
      slideIndex,
      width: size.width,
    };
  }
}

export function createPptxThumbnailRenderer(
  presentation: ParsedPresentation,
  options: PptxThumbnailRendererOptions = {},
): PptxThumbnailRenderer {
  return new ThumbnailRenderer(presentation, options);
}
