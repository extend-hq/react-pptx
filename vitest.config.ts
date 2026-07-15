import { defineConfig } from 'vitest/config';

export default defineConfig({
  resolve: { tsconfigPaths: true },
  test: {
    environment: 'jsdom',
    include: ['packages/*/src/**/*.test.{ts,tsx}'],
    coverage: { reporter: ['text', 'json', 'html'] },
  },
});
