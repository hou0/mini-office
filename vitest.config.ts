import { defineConfig } from 'vitest/config';
import { dirname, resolve } from 'node:path';
import { createRequire } from 'node:module';

// 功能：让根目录组件测试复用桌面应用的 React 运行时，适配 pnpm 依赖隔离；时间：2026-10-19；作者：lq。
const desktopRequire = createRequire(resolve(__dirname, 'apps/desktop/package.json'));

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
  'plugin-api',
  'plugin-host',
];

export default defineConfig({
  resolve: {
    alias: {
      ...Object.fromEntries(
        packages.map((name) => [`@mini-office/${name}`, resolve(__dirname, `packages/${name}/src/index.ts`)]),
      ),
      react: dirname(desktopRequire.resolve('react/package.json')),
      'react-dom': dirname(desktopRequire.resolve('react-dom/package.json')),
    },
  },
  test: {
    environment: 'jsdom',
    include: ['tests/**/*.test.ts'],
    globals: true,
    testTimeout: 30000,
  },
});
