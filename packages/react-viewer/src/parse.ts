import type { PresentationDocument, PresentationFormat } from '@extend-ai/react-pptx-model';
import { PptxViewerError, throwIfAborted, toPptxViewerError } from './errors';
import type {
  BinaryPresentationSource,
  ParsePresentationOptions,
  ParsedPresentation,
  PresentationSource,
} from './types';
import { parseWithWasm } from './wasm';

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
    (source.format === 'ppt' || source.format === 'pptx') &&
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

function detectFormat(buffer: ArrayBuffer, hint?: PresentationFormat): PresentationFormat {
  if (hint) return hint;
  const bytes = new Uint8Array(buffer, 0, Math.min(8, buffer.byteLength));
  if (bytes[0] === 0x50 && bytes[1] === 0x4b) return 'pptx';
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
  throw new PptxViewerError(
    'unsupported-format',
    'The file is neither an OOXML PowerPoint presentation (.pptx) nor a legacy PowerPoint compound document (.ppt).',
  );
}

/**
 * Parse a PowerPoint source into the format-neutral public model used by the
 * package-owned renderer.
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
      warnings: source.warnings,
    };
  }

  const buffer = await readSource(source, options);
  const format = detectFormat(buffer, options.formatHint);
  const bytes = new Uint8Array(buffer);
  throwIfAborted(options.signal);

  const nativeDocument = await parseWithWasm(bytes, {
    formatHint: format,
    ...(options.signal ? { signal: options.signal } : {}),
  });
  throwIfAborted(options.signal);

  if (!nativeDocument) {
    throw new PptxViewerError(
      format === 'ppt' ? 'unsupported-format' : 'parse-failed',
      format === 'ppt'
        ? 'This legacy .ppt file could not be decoded by the native parser.'
        : 'The PowerPoint presentation could not be decoded by the native parser.',
    );
  }
  return {
    kind: 'parsed-presentation',
    document: nativeDocument,
    warnings: nativeDocument.warnings,
  };
}

export const presentationParsingDefaults = Object.freeze({
  maxInputBytes: DEFAULT_MAX_INPUT_BYTES,
});
