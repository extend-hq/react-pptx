export type PptxViewerErrorCode =
  | 'aborted'
  | 'fetch-failed'
  | 'invalid-source'
  | 'unsupported-environment'
  | 'unsupported-format'
  | 'encrypted-document'
  | 'resource-limit'
  | 'parse-failed'
  | 'render-failed';

export class PptxViewerError extends Error {
  readonly code: PptxViewerErrorCode;
  readonly cause?: unknown;
  readonly details?: Readonly<Record<string, unknown>>;

  constructor(
    code: PptxViewerErrorCode,
    message: string,
    options: { cause?: unknown; details?: Readonly<Record<string, unknown>> } = {},
  ) {
    super(message);
    this.name = 'PptxViewerError';
    this.code = code;
    if (options.cause !== undefined) this.cause = options.cause;
    if (options.details !== undefined) this.details = options.details;
  }
}

export function toPptxViewerError(
  error: unknown,
  fallbackCode: PptxViewerErrorCode,
  fallbackMessage: string,
): PptxViewerError {
  if (error instanceof PptxViewerError) return error;
  if (error instanceof DOMException && error.name === 'AbortError') {
    return new PptxViewerError('aborted', 'Presentation loading was cancelled.', { cause: error });
  }
  if (error instanceof Error && error.name === 'AbortError') {
    return new PptxViewerError('aborted', 'Presentation loading was cancelled.', { cause: error });
  }
  return new PptxViewerError(fallbackCode, fallbackMessage, { cause: error });
}

export function throwIfAborted(signal?: AbortSignal): void {
  if (signal?.aborted) {
    throw new PptxViewerError('aborted', 'Presentation loading was cancelled.', {
      cause: signal.reason,
    });
  }
}
