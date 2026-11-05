/// <reference lib="webworker" />

import { parsePresentation, setWasmSource, type WasmParseOptions } from '@extend-ai/react-pptx-wasm/eager';
import type { WorkerWasmSource } from './wasm';

interface ParseRequest {
  id: number;
  bytes: ArrayBuffer;
  wasmSource?: WorkerWasmSource;
  parserOptions?: WasmParseOptions;
}

interface ParseResponse {
  id: number;
  result?: unknown;
  error?: string;
}

self.addEventListener('message', (event: MessageEvent<ParseRequest>) => {
  const { id, bytes, wasmSource, parserOptions } = event.data;
  if (wasmSource !== undefined) setWasmSource(wasmSource);
  void parsePresentation(new Uint8Array(bytes), parserOptions)
    .then((result) => {
      const response: ParseResponse = { id, result };
      self.postMessage(response);
    })
    .catch((error: unknown) => {
      const response: ParseResponse = {
        id,
        error: error instanceof Error ? error.message : String(error),
      };
      self.postMessage(response);
    });
});
