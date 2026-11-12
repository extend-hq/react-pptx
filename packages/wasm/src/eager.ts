import * as runtime from '../pkg/pptx_wasm.js';
import type { WasmParseOptions } from './options';
export type { WasmParseOptions } from './options';

import initialize, { detect_presentation_format, parse_presentation } from '../pkg/pptx_wasm.js';

/** The source formats the parser detects. */
export type WasmPresentationFormat = 'ppt' | 'pptx' | 'odp' | 'odg';

export type WasmSource =
  | string
  | URL
  | Request
  | Response
  | BufferSource
  | WebAssembly.Module;

let source: WasmSource | undefined;
let initialized: Promise<unknown> | undefined;

export function setWasmSource(nextSource: WasmSource): void {
  source = nextSource;
  initialized = undefined;
}

export async function initWasm(nextSource?: WasmSource): Promise<void> {
  if (nextSource !== undefined) setWasmSource(nextSource);
  initialized ??= initialize(source === undefined ? undefined : { module_or_path: source });
  await initialized;
}

export async function parsePresentation(bytes: Uint8Array, options?: WasmParseOptions): Promise<unknown> {
  await initWasm();
  return parse_presentation(bytes, options);
}

export async function detectPresentationFormat(bytes: Uint8Array): Promise<WasmPresentationFormat> {
  await initWasm();
  return detect_presentation_format(bytes) as WasmPresentationFormat;
}

export async function extractPresentationBlocks(document: unknown, options?: { includeWords?: boolean; includeNotes?: boolean }): Promise<unknown> {
  await initWasm();
  const extract = (runtime as unknown as { extract_presentation_blocks?: (document: unknown, options?: unknown) => unknown }).extract_presentation_blocks;
  if (!extract) throw new Error('This runtime does not support block extraction; build with the updated Rust parser.');
  return extract(document, options);
}
