import { beforeEach, describe, expect, it, vi } from 'vitest';
import { PptxViewerError } from './errors';
import { parsePresentation, presentationParsingDefaults } from './parse';
import type { PresentationDocument } from '@extend-ai/react-pptx-model';
import type { ParsePresentationOptions } from './types';
import { parseWithWasm } from './wasm';

vi.mock('./wasm', () => ({ parseWithWasm: vi.fn() }));

const nativeDocument: PresentationDocument = {
  format: 'pptx',
  size: { widthEmu: 12_192_000, heightEmu: 6_858_000 },
  slides: [{ id: 'slide-1', index: 0, nodes: [] }],
  masters: [],
  layouts: [],
  themes: [],
  assets: {},
  warnings: [],
};

describe('parsePresentation', () => {
  beforeEach(() => vi.mocked(parseWithWasm).mockReset());

  it('accepts an already normalized presentation without a DOM parser', async () => {
    const document: PresentationDocument = {
      format: 'ppt',
      size: { widthEmu: 9_144_000, heightEmu: 6_858_000 },
      slides: [],
      masters: [],
      layouts: [],
      themes: [],
      assets: {},
      warnings: [],
    };
    const result = await parsePresentation(document);
    expect(result.document).toBe(document);
    expect(result.kind).toBe('parsed-presentation');
  });

  it('rejects unknown magic bytes before loading the renderer', async () => {
    await expect(parsePresentation(new Uint8Array([1, 2, 3, 4]))).rejects.toMatchObject({
      code: 'unsupported-format',
    } satisfies Partial<PptxViewerError>);
  });

  it('uses the owned Wasm model for PPTX input', async () => {
    vi.mocked(parseWithWasm).mockResolvedValueOnce(nativeDocument);

    const result = await parsePresentation(new Uint8Array([0x50, 0x4b, 0x03, 0x04]));

    expect(parseWithWasm).toHaveBeenCalledWith(
      expect.any(Uint8Array),
      expect.objectContaining({ formatHint: 'pptx' }),
    );
    expect(result).toEqual({
      kind: 'parsed-presentation',
      document: nativeDocument,
      warnings: [],
    });
    expect(result).not.toHaveProperty('rendererData');
  });

  it('reports a native PPTX decode failure without loading a second parser', async () => {
    vi.mocked(parseWithWasm).mockResolvedValueOnce(null);

    await expect(parsePresentation(new Uint8Array([0x50, 0x4b, 0x03, 0x04]))).rejects.toMatchObject(
      { code: 'parse-failed' },
    );
  });

  it('enforces the default input ceiling', () => {
    expect(presentationParsingDefaults).toEqual({ maxInputBytes: 100 * 1024 * 1024 });
  });

  it('exposes only parser options implemented by the owned parser boundary', () => {
    const supportedOptions: Record<keyof ParsePresentationOptions, true> = {
      signal: true,
      formatHint: true,
      maxInputBytes: true,
      fetchInit: true,
    };

    expect(Object.keys(supportedOptions)).toEqual([
      'signal',
      'formatHint',
      'maxInputBytes',
      'fetchInit',
    ]);
  });
});
