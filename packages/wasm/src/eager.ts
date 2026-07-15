import initialize, { detect_presentation_format, parse_presentation } from '../pkg/pptx_wasm.js';

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

export async function parsePresentation(bytes: Uint8Array): Promise<unknown> {
  await initWasm();
  return parse_presentation(bytes);
}

export async function detectPresentationFormat(bytes: Uint8Array): Promise<'ppt' | 'pptx'> {
  await initWasm();
  return detect_presentation_format(bytes) as 'ppt' | 'pptx';
}
