import type { WasmParseOptions } from './options';
export type { WasmParseOptions } from './options';

/** The source formats the parser detects. */
export type WasmPresentationFormat = 'ppt' | 'pptx' | 'odp' | 'odg';

type WasmModule = typeof import('../pkg/pptx_wasm.js');

export type WasmSource = string | URL | Request | Response | BufferSource | WebAssembly.Module;

let source: WasmSource | undefined;
let runtimePromise: Promise<WasmModule> | undefined;

export function setWasmSource(nextSource: WasmSource): void {
  source = nextSource;
  runtimePromise = undefined;
}

export async function initWasm(nextSource?: WasmSource): Promise<void> {
  if (nextSource !== undefined) setWasmSource(nextSource);
  runtimePromise ??= (async () => {
    const runtime = await import('../pkg/pptx_wasm.js');
    await runtime.default(source as Parameters<typeof runtime.default>[0]);
    return runtime;
  })();
  await runtimePromise;
}

export async function parsePresentation(bytes: Uint8Array, options?: WasmParseOptions): Promise<unknown> {
  await initWasm();
  const runtime = await runtimePromise!;
  return runtime.parse_presentation(bytes, options);
}

export async function detectPresentationFormat(bytes: Uint8Array): Promise<WasmPresentationFormat> {
  await initWasm();
  const runtime = await runtimePromise!;
  return runtime.detect_presentation_format(bytes) as WasmPresentationFormat;
}

export { parsePresentation as parse_presentation };

export async function extractPresentationBlocks(document: unknown, options?: { includeWords?: boolean; includeNotes?: boolean }): Promise<unknown> {
  await initWasm();
  const runtime = await runtimePromise!;
  const extract = (runtime as unknown as { extract_presentation_blocks?: (document: unknown, options?: unknown) => unknown }).extract_presentation_blocks;
  if (!extract) throw new Error('This runtime does not support block extraction; build with the updated Rust parser.');
  return extract(document, options);
}
