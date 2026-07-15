import { existsSync, readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { PresentationDocument } from '@extend-ai/react-pptx-model';
import { describe, expect, it, vi } from 'vitest';

vi.mock('./wasm', () => ({ parseWithWasm: vi.fn() }));

import { PptxViewerError } from './errors';
import { parsePresentation } from './parse';
import { parseWithWasm } from './wasm';

const fixturePath = [
  resolve(process.cwd(), 'tests/fixtures/legacy/file-example-250kb.ppt'),
  resolve(process.cwd(), '../../tests/fixtures/legacy/file-example-250kb.ppt'),
].find(existsSync);

if (!fixturePath) throw new Error('Legacy PPT fixture was not found from the workspace root.');

const fixture = new Uint8Array(readFileSync(fixturePath));

const legacyDocument: PresentationDocument = {
  format: 'ppt',
  size: { widthEmu: 10_080_625, heightEmu: 7_559_675 },
  slides: [
    {
      id: 'slide-1',
      index: 0,
      nodes: [
        {
          id: 'legacy-text-1',
          type: 'shape',
          transform: { x: 0, y: 0, width: 5_000_000, height: 500_000 },
          geometry: { preset: 'rect' },
          paragraphs: [{ runs: [{ text: 'Lorem ipsum' }] }],
        },
      ],
    },
    { id: 'slide-2', index: 1, nodes: [] },
    { id: 'slide-3', index: 2, nodes: [] },
  ],
  masters: [],
  layouts: [],
  themes: [],
  assets: {},
  warnings: [
    {
      code: 'degraded-rendering',
      severity: 'warning',
      message: 'Unsupported binary PowerPoint records use normalized fallbacks.',
    },
  ],
};

describe('legacy PPT parsing pipeline', () => {
  it('routes a real OLE fixture through the owned Wasm model parser', async () => {
    vi.mocked(parseWithWasm).mockResolvedValueOnce(legacyDocument);

    const parsed = await parsePresentation(fixture);

    expect([...fixture.subarray(0, 8)]).toEqual([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
    expect(parseWithWasm).toHaveBeenCalledWith(
      expect.any(Uint8Array),
      expect.objectContaining({ formatHint: 'ppt' }),
    );
    expect(parsed.document).toBe(legacyDocument);
    expect(parsed.document.slides).toHaveLength(3);
    expect(parsed.document.slides[0]?.nodes[0]).toMatchObject({
      type: 'shape',
      paragraphs: [{ runs: [{ text: 'Lorem ipsum' }] }],
    });
    expect(parsed.warnings[0]?.code).toBe('degraded-rendering');
  });

  it('preserves the native encrypted-document failure', async () => {
    vi.mocked(parseWithWasm).mockRejectedValueOnce(
      new PptxViewerError('encrypted-document', 'Encrypted presentations are not supported.'),
    );

    await expect(parsePresentation(fixture)).rejects.toMatchObject({
      code: 'encrypted-document',
    });
  });
});
