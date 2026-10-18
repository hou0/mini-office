/** @type {import('tailwindcss').Config} */
export default {
  darkMode: ['class', '[data-theme="dark"]'],
  content: [
    './index.html',
    './src/**/*.{ts,tsx}',
    '../../packages/*/src/**/*.{ts,tsx}',
  ],
  theme: {
    extend: {
      colors: {
        app: 'var(--mo-bg)',
        surface: 'var(--mo-surface)',
        'surface-2': 'var(--mo-surface-2)',
        'surface-3': 'var(--mo-surface-3)',
        accent: 'var(--mo-accent)',
        'accent-soft': 'var(--mo-accent-soft)',
        'text-primary': 'var(--mo-text-primary)',
        'text-muted': 'var(--mo-text-muted)',
        'text-faint': 'var(--mo-text-faint)',
        line: 'var(--mo-line)',
        'line-strong': 'var(--mo-line-strong)',
        danger: 'var(--mo-danger)',
        success: 'var(--mo-success)',
        warning: 'var(--mo-warning)',
        canvas: 'var(--mo-canvas)',
      },
      borderRadius: {
        md: 'var(--mo-radius-md)',
        lg: 'var(--mo-radius-lg)',
        xl: 'var(--mo-radius-xl)',
      },
      boxShadow: {
        panel: 'var(--mo-shadow-panel)',
        pop: 'var(--mo-shadow-pop)',
        page: 'var(--mo-shadow-page)',
      },
      transitionDuration: {
        fast: 'var(--mo-motion-fast)',
        base: 'var(--mo-motion-base)',
      },
      fontFamily: {
        ui: 'var(--mo-font-ui)',
        doc: 'var(--mo-font-doc)',
      },
    },
  },
  plugins: [],
};
