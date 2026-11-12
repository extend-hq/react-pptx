import { isPresentationFormat } from '@extend-ai/react-pptx-model';
import type { PresentationDocument, PresentationFormat } from '@extend-ai/react-pptx-model';
import { PptxViewerError, throwIfAborted, toPptxViewerError } from './errors';
import type {
  BinaryPresentationSource,
  ParsePresentationOptions,
  ParsedPresentation,
  PresentationSource,
} from './types';
import { parseWithWasm } from './wasm';
import { performanceMeasurement, performanceNow } from './performance';

const DEFAULT_MAX_INPUT_BYTES = 100 * 1024 * 1024;

function isParsedPresentation(source: PresentationSource): source is ParsedPresentation {
  return (
    !!source &&
    typeof source === 'object' &&
    'kind' in source &&
    source.kind === 'parsed-presentation' &&
    'document' in source
  );
}

function isPresentationDocument(source: PresentationSource): source is PresentationDocument {
  return (
    !!source &&
    typeof source === 'object' &&
    'format' in source &&
    isPresentationFormat(source.format) &&
    'slides' in source &&
    Array.isArray(source.slides)
  );
}

async function readSource(
  source: BinaryPresentationSource,
  options: ParsePresentationOptions,
): Promise<ArrayBuffer> {
  throwIfAborted(options.signal);

  let buffer: ArrayBuffer;
  if (source instanceof ArrayBuffer) {
    buffer = source.slice(0);
  } else if (source instanceof Uint8Array) {
    buffer = source.buffer.slice(
      source.byteOffset,
      source.byteOffset + source.byteLength,
    ) as ArrayBuffer;
  } else if (typeof Blob !== 'undefined' && source instanceof Blob) {
    buffer = await source.arrayBuffer();
  } else if (typeof source === 'string' || source instanceof URL) {
    let response: Response;
    try {
      const requestInit: RequestInit = { ...options.fetchInit };
      if (options.signal) requestInit.signal = options.signal;
      response = await fetch(source, requestInit);
    } catch (error) {
      throw toPptxViewerError(error, 'fetch-failed', 'Could not fetch the presentation.');
    }
    if (!response.ok) {
      throw new PptxViewerError(
        'fetch-failed',
        `Could not fetch the presentation (${response.status} ${response.statusText}).`,
        { details: { status: response.status, url: response.url } },
      );
    }
    const declaredLength = Number(response.headers.get('content-length'));
    const maximum = options.maxInputBytes ?? DEFAULT_MAX_INPUT_BYTES;
    if (Number.isFinite(declaredLength) && declaredLength > maximum) {
      throw new PptxViewerError(
        'resource-limit',
        `The presentation exceeds the ${Math.round(maximum / 1024 / 1024)} MB input limit.`,
        { details: { byteLength: declaredLength, maximum } },
      );
    }
    buffer = await response.arrayBuffer();
  } else {
    throw new PptxViewerError('invalid-source', 'Unsupported presentation source.');
  }

  throwIfAborted(options.signal);
  const maximum = options.maxInputBytes ?? DEFAULT_MAX_INPUT_BYTES;
  if (buffer.byteLength > maximum) {
    throw new PptxViewerError(
      'resource-limit',
      `The presentation exceeds the ${Math.round(maximum / 1024 / 1024)} MB input limit.`,
      { details: { byteLength: buffer.byteLength, maximum } },
    );
  }
  if (buffer.byteLength === 0) {
    throw new PptxViewerError('invalid-source', 'The presentation is empty.');
  }
  return buffer;
}

const OPENDOCUMENT_PREFIX = 'application/vnd.oasis.opendocument.';

/**
 * The OpenDocument format a media type names, or `undefined` for an
 * OpenDocument type that is not a presentation or drawing (text, spreadsheet).
 */
function openDocumentFormat(mediaType: string): PresentationFormat | undefined {
  const kind = mediaType.slice(OPENDOCUMENT_PREFIX.length);
  if (kind.startsWith('presentation')) return 'odp';
  if (kind.startsWith('graphics')) return 'odg';
  return undefined;
}

function unsupportedOpenDocument(mediaType: string): PptxViewerError {
  return new PptxViewerError(
    'unsupported-format',
    `This OpenDocument file is not a presentation or drawing (${mediaType}).`,
    { details: { mediaType } },
  );
}

/**
 * Identifies the container from its first bytes. The parser makes the final
 * decision; this only rejects files that cannot be a presentation before
 * they reach it and labels errors.
 */
function detectFormat(buffer: ArrayBuffer, hint?: PresentationFormat): PresentationFormat {
  if (hint) return hint;
  const bytes = new Uint8Array(buffer, 0, Math.min(4096, buffer.byteLength));
  if (bytes[0] === 0x50 && bytes[1] === 0x4b) {
    // An OpenDocument package stores its media type, uncompressed, as the
    // first entry: a 30-byte local header, the name `mimetype`, then the type.
    const head = new TextDecoder('latin1').decode(bytes.subarray(0, Math.min(bytes.length, 160)));
    if (head.slice(30, 38) === 'mimetype' && head.startsWith(OPENDOCUMENT_PREFIX, 38)) {
      // Media types are lower case, so the next entry's `PK` signature ends the match.
      const mediaType = /^[a-z0-9.+\/-]+/.exec(head.slice(38))?.[0] ?? OPENDOCUMENT_PREFIX;
      const format = openDocumentFormat(mediaType);
      if (!format) throw unsupportedOpenDocument(mediaType);
      return format;
    }
    return 'pptx';
  }
  if (
    bytes[0] === 0xd0 &&
    bytes[1] === 0xcf &&
    bytes[2] === 0x11 &&
    bytes[3] === 0xe0 &&
    bytes[4] === 0xa1 &&
    bytes[5] === 0xb1 &&
    bytes[6] === 0x1a &&
    bytes[7] === 0xe1
  ) {
    return 'ppt';
  }
  // A flat OpenDocument file is one XML document that names its media type
  // on the root element.
  const text = new TextDecoder('utf-8').decode(bytes).replace(/^\uFEFF/, '').trimStart();
  if (text.startsWith('<')) {
    const mediaType = /office:mimetype\s*=\s*["'](application\/vnd\.oasis\.opendocument\.[a-z.+-]+)["']/.exec(
      text,
    )?.[1];
    if (mediaType) {
      const format = openDocumentFormat(mediaType);
      if (!format) throw unsupportedOpenDocument(mediaType);
      return format;
    }
  }
  throw new PptxViewerError(
    'unsupported-format',
    'The file is not a PowerPoint presentation (.pptx, .ppt) or an OpenDocument presentation or drawing (.odp, .odg).',
  );
}

function presentationWarnings(document: PresentationDocument) {
  return [
    ...document.warnings,
    ...document.slides.flatMap((slide) =>
      (slide.warnings ?? []).map((warning) => ({
        ...warning,
        slideIndex: warning.slideIndex ?? slide.index,
      })),
    ),
  ];
}

/**
 * Parse a PowerPoint or OpenDocument presentation source into the
 * format-neutral public model used by the package-owned renderer.
 */
export async function parsePresentation(
  source: PresentationSource,
  options: ParsePresentationOptions = {},
): Promise<ParsedPresentation> {
  if (isParsedPresentation(source)) return source;
  if (isPresentationDocument(source)) {
    return {
      kind: 'parsed-presentation',
      document: source,
      warnings: presentationWarnings(source),
    };
  }

  const downloadStartedAt = performanceNow();
  const buffer = await readSource(source, options);
  if (typeof source === 'string' || source instanceof URL) {
    options.onPerformanceMeasurement?.(performanceMeasurement('download', downloadStartedAt));
  }
  const format = detectFormat(buffer, options.formatHint);
  const bytes = new Uint8Array(buffer);
  throwIfAborted(options.signal);

  const parsingStartedAt = performanceNow();
  const nativeDocument = await parseWithWasm(bytes, {
    formatHint: format,
    ...(options.include3dModels !== undefined ? { include3dModels: options.include3dModels } : {}),
    ...(options.signal ? { signal: options.signal } : {}),
  });
  options.onPerformanceMeasurement?.(performanceMeasurement('wasm-parsing', parsingStartedAt));
  throwIfAborted(options.signal);

  if (!nativeDocument) {
    throw new PptxViewerError(
      format === 'ppt' ? 'unsupported-format' : 'parse-failed',
      format === 'ppt'
        ? 'This legacy .ppt file could not be decoded by the native parser.'
        : format === 'odp' || format === 'odg'
          ? 'The OpenDocument file could not be decoded by the native parser.'
          : 'The PowerPoint presentation could not be decoded by the native parser.',
    );
  }
  return {
    kind: 'parsed-presentation',
    document: nativeDocument,
    warnings: presentationWarnings(nativeDocument),
  };
}

export const presentationParsingDefaults = Object.freeze({
  maxInputBytes: DEFAULT_MAX_INPUT_BYTES,
  include3dModels: false,
});
