import { defineConfig } from 'tsup';

// d3/regl/topojson/virtual-core ship as regular dependencies; only the
// TopoJSON atlas data stays bundled, inside a dynamically imported chunk so
// consumers never download map geometry unless a deck contains a map chart.
const bundledPackages = [
  'emf-converter',
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
  onSuccess: 'cp ../wasm/dist/pptx_wasm_bg.wasm dist/pptx_wasm_bg.wasm',
  external: ['react', 'react-dom', 'react/jsx-runtime'],
});
