import { defineConfig } from 'vitest/config';
import { resolve } from 'node:path';

const packages = [
  'design-system',
  'shared',
  'ui',
  'platform-adapter',
  'document-model',
  'ooxml',
  'editor-shell',
  'editor-word',
  'editor-sheet',
  'editor-slides',
  'export-service',
];

export default defineConfig({
  resolve: {
    alias: Object.fromEntries(
      packages.map((name) => [`@mini-office/${name}`, resolve(__dirname, `packages/${name}/src/index.ts`)]),
    ),
  },
  test: {
    environment: 'jsdom',
    include: ['tests/**/*.test.ts'],
    globals: true,
    testTimeout: 30000,
  },
});
