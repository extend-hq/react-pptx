import tailwindcss from '@tailwindcss/vite';
import react from '@vitejs/plugin-react';
import { defineConfig } from 'vite';

const srcDir = decodeURIComponent(new URL('./src/', import.meta.url).pathname);

export default defineConfig({
  plugins: [react(), tailwindcss()],
  resolve: {
    alias: [{ find: /^@\//, replacement: srcDir }],
  },
  server: {
    port: 4173,
  },
  preview: {
    port: 4173,
  },
});
