import { defineConfig } from 'tsup';
import { copyFile } from 'node:fs/promises';
import { createRequire } from 'node:module';

// d3/regl/topojson/TanStack Virtual ship as regular dependencies; only the
// TopoJSON atlas data stays bundled, inside a dynamically imported chunk so
// consumers never download map geometry unless a deck contains a map chart.
const bundledPackages = [
  "@extend-ai/office-fonts",
  /^@extend-ai\/react-pptx-(?:model|wasm)(?:\/.*)?$/,
  /^us-atlas(?:\/.*)?$/,
  /^world-atlas(?:\/.*)?$/,
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
