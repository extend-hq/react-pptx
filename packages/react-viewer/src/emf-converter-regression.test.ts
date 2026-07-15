import { convertEmfToDataUrl } from 'emf-converter';
import { afterEach, describe, expect, it, vi } from 'vitest';

interface TextDraw {
  text: string;
  x: number;
  y: number;
  font: string;
}

class FakeCanvasContext {
  font = '';
  fillStyle: string | CanvasGradient | CanvasPattern = '#000000';
  textAlign: CanvasTextAlign = 'start';
  textBaseline: CanvasTextBaseline = 'alphabetic';
  readonly textDraws: TextDraw[] = [];

  save() {}
  restore() {}
  setTransform() {}
  fillText(text: string, x: number, y: number) {
    this.textDraws.push({ text, x, y, font: this.font });
  }
}

class FakeOffscreenCanvas {
  readonly context = new FakeCanvasContext();

  constructor(
    readonly width: number,
    readonly height: number,
  ) {
    canvases.push(this);
  }

  getContext() {
    return this.context;
  }

  async convertToBlob(options?: { type?: string }) {
    return new Blob([new Uint8Array([0x89, 0x50, 0x4e, 0x47])], {
      type: options?.type ?? 'image/png',
    });
  }
}

const canvases: FakeOffscreenCanvas[] = [];

function emfRecord(type: number, dataLength: number, write?: (view: DataView) => void) {
  const bytes = new Uint8Array(8 + dataLength);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, type, true);
  view.setUint32(4, bytes.byteLength, true);
  write?.(new DataView(bytes.buffer, 8));
  return bytes;
}

function writeUtf16(view: DataView, offset: number, value: string) {
  for (let index = 0; index < value.length; index += 1) {
    view.setUint16(offset + index * 2, value.charCodeAt(index), true);
  }
}

function mappingModeEmf(explicitViewport: boolean): ArrayBuffer {
  const records = [
    emfRecord(1, 100, (view) => {
      view.setInt32(0, 0, true);
      view.setInt32(4, 0, true);
      view.setInt32(8, 100, true);
      view.setInt32(12, 100, true);
      view.setInt32(16, 0, true);
      view.setInt32(20, 0, true);
      view.setInt32(24, 2_646, true);
      view.setInt32(28, 2_646, true);
    }),
    emfRecord(17, 4, (view) => view.setInt32(0, 8, true)),
    emfRecord(9, 8, (view) => {
      view.setInt32(0, 1_000, true);
      view.setInt32(4, 1_000, true);
    }),
    ...(explicitViewport
      ? [
          emfRecord(11, 8, (view) => {
            view.setInt32(0, 100, true);
            view.setInt32(4, 100, true);
          }),
        ]
      : []),
    emfRecord(82, 324, (view) => {
      view.setUint32(0, 3, true);
      view.setInt32(4, -100, true);
      view.setInt32(20, 400, true);
      writeUtf16(view, 32, 'Liberation Sans');
    }),
    emfRecord(37, 4, (view) => view.setUint32(0, 3, true)),
    emfRecord(84, 72, (view) => {
      view.setInt32(28, 500, true);
      view.setInt32(32, 500, true);
      view.setUint32(36, 1, true);
      view.setUint32(40, 76, true);
      writeUtf16(view, 68, 'A');
    }),
    emfRecord(14, 12),
  ];
  const byteLength = records.reduce((total, record) => total + record.byteLength, 0);
  const bytes = new Uint8Array(byteLength);
  let offset = 0;
  for (const record of records) {
    bytes.set(record, offset);
    offset += record.byteLength;
  }
  const header = new DataView(bytes.buffer);
  header.setUint32(48, byteLength, true);
  header.setUint32(52, records.length, true);
  return bytes.buffer;
}

describe('patched emf-converter mapping mode', () => {
  afterEach(() => {
    canvases.length = 0;
    vi.unstubAllGlobals();
  });

  it.each([
    ['an explicit viewport extent', true],
    ['the header-sized default viewport extent', false],
  ])('scales text and coordinates at 2x DPI with %s', async (_label, explicitViewport) => {
    vi.stubGlobal('OffscreenCanvas', FakeOffscreenCanvas);

    const url = await convertEmfToDataUrl(mappingModeEmf(explicitViewport), undefined, undefined, {
      dpiScale: 2,
    });

    expect(url).toMatch(/^data:image\/png;base64,/);
    expect(canvases).toHaveLength(1);
    expect([canvases[0]?.width, canvases[0]?.height]).toEqual([200, 200]);
    expect(canvases[0]?.context.textDraws).toEqual([
      { text: 'A', x: 100, y: 100, font: '20px "Liberation Sans"' },
    ]);
  });
});
