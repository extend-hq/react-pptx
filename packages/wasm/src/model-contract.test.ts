import { existsSync } from 'node:fs';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';

import initialize, { parse_presentation } from '../pkg/pptx_wasm.js';

interface SerializedImageNode {
  type: 'image';
  assetId?: string;
  preserveAspectRatio?: boolean;
  asset_id?: string;
  preserve_aspect_ratio?: boolean;
}

describe('Rust to Wasm presentation model contract', () => {
  it('returns Record-compatible assets and camel-cased image fields', async () => {
    const workspaceRoot = [process.cwd(), resolve(process.cwd(), '../..')].find((directory) =>
      existsSync(resolve(directory, 'packages/wasm/pkg/pptx_wasm_bg.wasm')),
    );
    if (!workspaceRoot) throw new Error('Could not locate the react-pptx workspace root.');

    const wasm = await readFile(resolve(workspaceRoot, 'packages/wasm/pkg/pptx_wasm_bg.wasm'));
    await initialize({ module_or_path: wasm });

    const fixture = await readFile(
      resolve(workspaceRoot, 'tests/fixtures/legacy/file-example-250kb.ppt'),
    );
    const document = parse_presentation(fixture) as {
      assets: unknown;
      slides: Array<{ nodes: SerializedImageNode[] }>;
    };

    expect(document.assets).not.toBeInstanceOf(Map);
    expect(Object.getPrototypeOf(document.assets)).toBe(Object.prototype);
    expect(Object.keys(document.assets as Record<string, unknown>).length).toBeGreaterThan(0);
    const assets = Object.values(document.assets as Record<string, { data?: unknown }>);
    expect(assets.some((asset) => asset.data instanceof Uint8Array)).toBe(true);

    const image = document.slides
      .flatMap((slide) => slide.nodes)
      .find((node) => node.type === 'image');
    expect(image).toBeDefined();
    expect(image?.assetId).toMatch(/^legacy-picture-/);
    expect(image?.preserveAspectRatio).toBeTypeOf('boolean');
    expect(image).not.toHaveProperty('asset_id');
    expect(image).not.toHaveProperty('preserve_aspect_ratio');
  });
});
