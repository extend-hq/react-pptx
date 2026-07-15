import type { PresentationDocument } from '@extend-ai/react-pptx-model';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

const parsedDocument: PresentationDocument = {
  format: 'pptx',
  size: { widthEmu: 9_144_000, heightEmu: 6_858_000 },
  slides: [],
  masters: [],
  layouts: [],
  themes: [],
  assets: {},
  warnings: [],
};

class MockWorker extends EventTarget {
  static instances: MockWorker[] = [];
  static errorMessage: string | undefined;

  readonly messages: Array<{
    id: number;
    bytes: ArrayBuffer;
    wasmSource?: string | ArrayBuffer | WebAssembly.Module;
  }> = [];
  terminated = false;

  constructor(
    readonly url: URL,
    readonly options?: WorkerOptions,
  ) {
    super();
    MockWorker.instances.push(this);
  }

  postMessage(message: (typeof this.messages)[number]): void {
    this.messages.push(message);
    queueMicrotask(() => {
      this.dispatchEvent(
        new MessageEvent('message', {
          data: MockWorker.errorMessage
            ? { id: message.id, error: MockWorker.errorMessage }
            : { id: message.id, result: parsedDocument },
        }),
      );
    });
  }

  terminate(): void {
    this.terminated = true;
  }
}

describe('Wasm worker configuration', () => {
  beforeEach(() => {
    vi.resetModules();
    MockWorker.instances = [];
    MockWorker.errorMessage = undefined;
    vi.stubGlobal('Worker', MockWorker);
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('forwards a configured Wasm URL to the parser worker', async () => {
    const { parseWithWasm, setWasmSource } = await import('./wasm');
    setWasmSource(new URL('https://cdn.example.com/pptx_wasm_bg.wasm'));

    await expect(parseWithWasm(new Uint8Array([1, 2, 3]))).resolves.toEqual(parsedDocument);

    expect(MockWorker.instances).toHaveLength(1);
    expect(MockWorker.instances[0]?.messages[0]?.wasmSource).toBe(
      'https://cdn.example.com/pptx_wasm_bg.wasm',
    );
    expect(MockWorker.instances[0]?.terminated).toBe(true);
  });

  it('copies a configured Wasm buffer before sending it to the worker', async () => {
    const { parseWithWasm, setWasmSource } = await import('./wasm');
    const source = new Uint8Array([0, 97, 115, 109]).buffer;
    setWasmSource(source);

    await parseWithWasm(new Uint8Array([1, 2, 3]));

    const forwarded = MockWorker.instances[0]?.messages[0]?.wasmSource;
    expect(forwarded).toBeInstanceOf(ArrayBuffer);
    expect(forwarded).not.toBe(source);
    expect(new Uint8Array(forwarded as ArrayBuffer)).toEqual(new Uint8Array(source));
  });

  it('does not start a worker for an already-aborted parse', async () => {
    const { parseWithWasm } = await import('./wasm');
    const controller = new AbortController();
    controller.abort();

    await expect(
      parseWithWasm(new Uint8Array([1, 2, 3]), { signal: controller.signal }),
    ).rejects.toMatchObject({ name: 'AbortError' });
    expect(MockWorker.instances).toHaveLength(0);
  });

  it('surfaces encrypted native files with the public error code', async () => {
    MockWorker.errorMessage = 'encrypted presentations are not supported';
    const { parseWithWasm } = await import('./wasm');

    await expect(parseWithWasm(new Uint8Array([1, 2, 3]))).rejects.toMatchObject({
      code: 'encrypted-document',
    });
    expect(MockWorker.instances[0]?.terminated).toBe(true);
  });
});
