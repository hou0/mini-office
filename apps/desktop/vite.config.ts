import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { resolve } from 'node:path';

const pkg = (name: string) => resolve(__dirname, `../../packages/${name}/src/index.ts`);

export default defineConfig({
  plugins: [react()],
  resolve: {
    alias: {
      '@mini-office/design-system': pkg('design-system'),
      '@mini-office/shared': pkg('shared'),
      '@mini-office/ui': pkg('ui'),
      '@mini-office/platform-adapter': pkg('platform-adapter'),
      '@mini-office/document-model': pkg('document-model'),
      '@mini-office/ooxml': pkg('ooxml'),
      '@mini-office/editor-shell': pkg('editor-shell'),
      '@mini-office/editor-word': pkg('editor-word'),
      '@mini-office/editor-sheet': pkg('editor-sheet'),
      '@mini-office/editor-slides': pkg('editor-slides'),
      '@mini-office/export-service': pkg('export-service'),
      '@app': resolve(__dirname, 'src'),
    },
  },
  clearScreen: false,
  server: {
    port: 1420,
    strictPort: true,
    host: '127.0.0.1',
    watch: { ignored: ['**/src-tauri/**'] },
  },
  envPrefix: ['VITE_', 'TAURI_ENV_'],
  build: {
    target: 'es2022',
    sourcemap: false,
    chunkSizeWarningLimit: 1600,
  },
});
