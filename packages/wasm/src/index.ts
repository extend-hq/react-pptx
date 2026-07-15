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

export async function parsePresentation(bytes: Uint8Array): Promise<unknown> {
  await initWasm();
  const runtime = await runtimePromise!;
  return runtime.parse_presentation(bytes);
}

export async function detectPresentationFormat(bytes: Uint8Array): Promise<'ppt' | 'pptx'> {
  await initWasm();
  const runtime = await runtimePromise!;
  return runtime.detect_presentation_format(bytes) as 'ppt' | 'pptx';
}

export { parsePresentation as parse_presentation };
