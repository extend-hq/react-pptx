import { defineConfig } from 'tsup';

const bundledPackages = ['emf-converter', /^@extend-ai\/react-pptx-(?:model|wasm)(?:\/.*)?$/];

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
