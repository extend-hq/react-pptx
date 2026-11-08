import { defineConfig } from 'tsup';
import { copyFile } from 'node:fs/promises';
import { createRequire } from 'node:module';

// Charts arrive as SVG from the Rust parser, so the published package has no
// runtime dependencies: the workspace packages, the Office font fallbacks, and
// TanStack Virtual are all bundled. Only React stays external.
const bundledPackages = [
  "@extend-ai/office-fonts",
  /^@extend-ai\/react-pptx-(?:model|wasm)(?:\/.*)?$/,
  /^@tanstack\/(?:react-virtual|virtual-core)(?:\/.*)?$/,
];

export default defineConfig({
  entry: {
    index: 'src/index.ts',
    'native-parser-worker': 'src/native-parser-worker.ts',
  },
  format: ['esm'],
  dts: {
    resolve: ['@extend-ai/react-pptx-model'],
  },
  sourcemap: true,
  clean: true,
  splitting: true,
  treeshake: true,
  noExternal: bundledPackages,
  onSuccess: async () => {
    await copyFile('../wasm/dist/pptx_wasm_bg.wasm', 'dist/pptx_wasm_bg.wasm');
    await copyFile(createRequire(import.meta.url).resolve('@extend-ai/office-fonts/FONT-LICENSES.txt'), 'dist/Office-Font-Licenses.txt');
  },
  external: ['react', 'react-dom', 'react/jsx-runtime'],
});
