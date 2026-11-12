import type { WasmParseOptions } from '@extend-ai/react-pptx-wasm';
import { isPresentationFormat } from '@extend-ai/react-pptx-model';
import type {
  ExtractPresentationBlocksOptions,
  StructuredPresentation,
  PresentationDocument,
  PresentationFormat,
} from '@extend-ai/react-pptx-model';
import { PptxViewerError } from './errors';

export type WasmSource =
  | string
  | URL
  | Request
  | Response
  | BufferSource
  | WebAssembly.Module;

export type WorkerWasmSource = string | ArrayBuffer | WebAssembly.Module;

interface WasmRuntimeModule {
  extractPresentationBlocks?: (document: PresentationDocument, options?: ExtractPresentationBlocksOptions) => unknown;
  default?: (source?: unknown) => Promise<unknown> | unknown;
  initWasm?: (source?: unknown) => Promise<unknown> | unknown;
  init?: (source?: unknown) => Promise<unknown> | unknown;
  setWasmSource?: (source: unknown) => void;
  parsePresentation?: (bytes: Uint8Array, options?: unknown) => Promise<unknown> | unknown;
  parse_presentation?: (bytes: Uint8Array, options?: unknown) => Promise<unknown> | unknown;
}

let configuredSource: WasmSource | undefined;
let hasConfiguredSource = false;
let configuredWorkerSource: WorkerWasmSource | undefined;
let runtimePromise: Promise<WasmRuntimeModule> | undefined;
let initialized = false;

async function loadRuntime(): Promise<WasmRuntimeModule> {
  runtimePromise ??= import('@extend-ai/react-pptx-wasm') as unknown as Promise<WasmRuntimeModule>;
  return runtimePromise;
}

export function setWasmSource(source: WasmSource): void {
  hasConfiguredSource = true;
  configuredSource = source;
  configuredWorkerSource = sourceToWorkerSource(source);
  initialized = false;
  void runtimePromise?.then((runtime) => runtime.setWasmSource?.(source));
}

function bufferSourceToArrayBuffer(
  source: ArrayBuffer | ArrayBufferView<ArrayBufferLike>,
): ArrayBuffer {
  if (source instanceof ArrayBuffer) return source.slice(0);
  const bytes = new Uint8Array(source.buffer, source.byteOffset, source.byteLength);
  return new Uint8Array(bytes).buffer;
}

function sourceToWorkerSource(source: WasmSource): WorkerWasmSource | undefined {
  if (typeof source === 'string') return source;
  if (typeof URL !== 'undefined' && source instanceof URL) return source.href;
  if (typeof Request !== 'undefined' && source instanceof Request) return source.url;
  if (source instanceof ArrayBuffer || ArrayBuffer.isView(source)) {
    return bufferSourceToArrayBuffer(source);
  }
  if (typeof WebAssembly !== 'undefined' && source instanceof WebAssembly.Module) return source;
  return undefined;
}

function canUseConfiguredSourceInWorker(): boolean {
  return !hasConfiguredSource || configuredWorkerSource !== undefined;
}

export async function initWasm(source?: WasmSource): Promise<void> {
  if (source !== undefined) setWasmSource(source);
  if (initialized) return;

  try {
    const runtime = await loadRuntime();
    if (configuredSource !== undefined) runtime.setWasmSource?.(configuredSource);
    const initialize = runtime.initWasm ?? runtime.init ?? runtime.default;
    if (initialize) await initialize(configuredSource);
    initialized = true;
  } catch (error) {
    runtimePromise = undefined;
    throw new PptxViewerError('parse-failed', 'Could not initialize the PowerPoint parser.', {
      cause: error,
    });
  }
}

function isPresentationDocument(value: unknown): value is PresentationDocument {
  if (!value || typeof value !== 'object') return false;
  const candidate = value as Partial<PresentationDocument>;
  return (
    isPresentationFormat(candidate.format) &&
    !!candidate.size &&
    Array.isArray(candidate.slides) &&
    Array.isArray(candidate.warnings)
  );
}

function parserFailure(error: unknown): PptxViewerError | null {
  if (error instanceof PptxViewerError) return error;
  const message = error instanceof Error ? error.message : String(error);
  const normalized = message.toLowerCase();
  if (normalized.includes('encrypted presentation')) {
    return new PptxViewerError(
      'encrypted-document',
      'Encrypted presentations are not supported.',
      { cause: error },
    );
  }
  if (
    normalized.includes('resource limit exceeded') ||
    normalized.includes('input exceeds the configured')
  ) {
    return new PptxViewerError('resource-limit', message, { cause: error });
  }
  return null;
}

export async function parseWithWasm(
  bytes: Uint8Array,
  options?: { signal?: AbortSignal; formatHint?: PresentationFormat; include3dModels?: boolean },
): Promise<PresentationDocument | null> {
  try {
    if (
      canUseConfiguredSourceInWorker() &&
      typeof window !== 'undefined' &&
      typeof Worker !== 'undefined'
    ) {
      const result = await parseInWorker(bytes, options?.signal, configuredWorkerSource, { include3dModels: options?.include3dModels ?? false });
      return isPresentationDocument(result) ? result : null;
    }
    await initWasm();
    const runtime = await loadRuntime();
    const parse = runtime.parsePresentation ?? runtime.parse_presentation;
    if (!parse) return null;
    const result = await parse(bytes, {
      include3dModels: options?.include3dModels ?? false,
    });
    const value = typeof result === 'string' ? (JSON.parse(result) as unknown) : result;
    return isPresentationDocument(value) ? value : null;
  } catch (error) {
    if (options?.signal?.aborted) throw error;
    const failure = parserFailure(error);
    if (failure) throw failure;
    return null;
  }
}

let workerRequestId = 0;

async function parseInWorker(
  bytes: Uint8Array,
  signal?: AbortSignal,
  wasmSource?: WorkerWasmSource,
  parserOptions?: WasmParseOptions,
): Promise<unknown> {
  if (signal?.aborted) throw new DOMException('Parsing was aborted.', 'AbortError');

  const worker = new Worker(new URL('./native-parser-worker.js', import.meta.url), {
    type: 'module',
    name: 'react-pptx-parser',
  });
  const id = ++workerRequestId;
  const owned = bytes.slice().buffer as ArrayBuffer;
  return await new Promise((resolve, reject) => {
    let settled = false;
    const cleanup = () => {
      signal?.removeEventListener('abort', abort);
      worker.removeEventListener('message', handleMessage);
      worker.removeEventListener('error', handleError);
      worker.removeEventListener('messageerror', handleMessageError);
      worker.terminate();
    };
    const settle = (callback: () => void) => {
      if (settled) return;
      settled = true;
      cleanup();
      callback();
    };
    const abort = () => {
      settle(() => reject(new DOMException('Parsing was aborted.', 'AbortError')));
    };
    const handleMessage = (
      event: MessageEvent<{ id: number; result?: unknown; error?: string }>,
    ) => {
      if (event.data.id !== id) return;
      settle(() => {
        if (event.data.error) reject(new Error(event.data.error));
        else resolve(event.data.result);
      });
    };
    const handleError = (event: ErrorEvent) => {
      settle(() => reject(event.error ?? new Error(event.message)));
    };
    const handleMessageError = () => {
      settle(() => reject(new Error('The parser worker returned an unreadable response.')));
    };

    signal?.addEventListener('abort', abort, { once: true });
    worker.addEventListener('message', handleMessage);
    worker.addEventListener('error', handleError);
    worker.addEventListener('messageerror', handleMessageError);

    if (signal?.aborted) {
      abort();
      return;
    }

    try {
      worker.postMessage({ id, bytes: owned, wasmSource, parserOptions }, [owned]);
    } catch (error) {
      settle(() => reject(error));
    }
  });
}

export async function extractPresentationBlocks(document: PresentationDocument, options: ExtractPresentationBlocksOptions = {}): Promise<StructuredPresentation> {
  await initWasm();
  const runtime = await loadRuntime();
  if (!runtime.extractPresentationBlocks) throw new Error('This parser runtime does not support presentation block extraction.');
  return await runtime.extractPresentationBlocks(document, options) as StructuredPresentation;
}
