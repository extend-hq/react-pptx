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

/** The start of an OpenDocument package: a stored `mimetype` entry first. */
function openDocumentPackage(kind: string): Uint8Array {
  const mediaType = `application/vnd.oasis.opendocument.${kind}`;
  const bytes = new Uint8Array(30 + 8 + mediaType.length + 16);
  bytes.set([0x50, 0x4b, 0x03, 0x04]);
  bytes[26] = 8; // file name length
  bytes.set(new TextEncoder().encode(`mimetype${mediaType}`), 30);
  // The next local header follows the stored media type.
  bytes.set([0x50, 0x4b, 0x03, 0x04], 38 + mediaType.length);
  return bytes;
}

function flatOpenDocument(kind: string, prefix = ''): Uint8Array {
  // Copied into this realm's Uint8Array, which is what the source check accepts.
  return new Uint8Array(new TextEncoder().encode(
    `${prefix}<?xml version="1.0" encoding="UTF-8"?>\n<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.${kind}"><office:body/></office:document>`,
  ));
}

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

  it('recognizes OpenDocument packages and flat XML and passes them to the parser', async () => {
    const cases: Array<[string, Uint8Array, 'odp' | 'odg']> = [
      ['presentation package', openDocumentPackage('presentation'), 'odp'],
      ['presentation template', openDocumentPackage('presentation-template'), 'odp'],
      ['drawing package', openDocumentPackage('graphics'), 'odg'],
      ['flat presentation', flatOpenDocument('presentation'), 'odp'],
      ['flat drawing with a byte-order mark', flatOpenDocument('graphics', '\uFEFF'), 'odg'],
    ];
    for (const [label, bytes, format] of cases) {
      vi.mocked(parseWithWasm).mockResolvedValueOnce({ ...nativeDocument, format });
      const result = await parsePresentation(bytes);
      expect(result.document.format, label).toBe(format);
      expect(vi.mocked(parseWithWasm).mock.lastCall?.[1]?.formatHint, label).toBe(format);
    }
  });

  it('rejects OpenDocument files that are not presentations or drawings', async () => {
    for (const bytes of [openDocumentPackage('text'), flatOpenDocument('spreadsheet')]) {
      await expect(parsePresentation(bytes)).rejects.toMatchObject({
        code: 'unsupported-format',
        message: expect.stringContaining('not a presentation or drawing'),
      });
    }
    expect(parseWithWasm).not.toHaveBeenCalled();
  });

  it('accepts an already normalized OpenDocument presentation', async () => {
    const document: PresentationDocument = { ...nativeDocument, format: 'odp' };
    expect((await parsePresentation(document)).document).toBe(document);
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

  it('reports native parsing performance', async () => {
    vi.mocked(parseWithWasm).mockResolvedValueOnce(nativeDocument);
    const measurements: string[] = [];

    await parsePresentation(new Uint8Array([0x50, 0x4b, 0x03, 0x04]), {
      onPerformanceMeasurement: (measurement) => measurements.push(measurement.phase),
    });

    expect(measurements).toEqual(['wasm-parsing']);
  });

  it('reports a native PPTX decode failure without loading a second parser', async () => {
    vi.mocked(parseWithWasm).mockResolvedValueOnce(null);

    await expect(parsePresentation(new Uint8Array([0x50, 0x4b, 0x03, 0x04]))).rejects.toMatchObject(
      { code: 'parse-failed' },
    );
  });

  it('enforces the default input ceiling', () => {
    expect(presentationParsingDefaults).toEqual({ maxInputBytes: 100 * 1024 * 1024, include3dModels: false });
  });

  it('exposes only parser options implemented by the owned parser boundary', () => {
    const supportedOptions: Record<keyof ParsePresentationOptions, true> = {
      signal: true,
      include3dModels: true,
      formatHint: true,
      maxInputBytes: true,
      fetchInit: true,
      onPerformanceMeasurement: true,
    };

    expect(Object.keys(supportedOptions)).toEqual([
      'signal',
      'include3dModels',
      'formatHint',
      'maxInputBytes',
      'fetchInit',
      'onPerformanceMeasurement',
    ]);
  });
});
