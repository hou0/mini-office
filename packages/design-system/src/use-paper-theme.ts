import { useEffect, useState } from 'react';

export type PaperTheme = 'light' | 'dark';

function readPaperTheme(): PaperTheme {
  if (typeof document === 'undefined') return 'light';
  return document.documentElement.dataset.theme === 'dark' ? 'dark' : 'light';
}

/**
 * 功能：订阅当前生效的纸张主题（浅色/深色）；时间：2026-10-31；作者：lq。
 *
 * 文档画布的颜色换算需要知道主题（同一份 #FFF2CC 底纹，浅色纸上是浅黄、深色纸上要压暗），
 * 纯 CSS 变量做不到「按色相压暗」，因此由 React 侧读取 data-theme 并跟随切换。
 * 主题切换是低频操作，用 MutationObserver 订阅属性变化即可，不引入全局状态。
 */
export function usePaperTheme(): PaperTheme {
  const [theme, setTheme] = useState<PaperTheme>(readPaperTheme);
  useEffect(() => {
    if (typeof document === 'undefined') return;
    const root = document.documentElement;
    const sync = () => setTheme(readPaperTheme());
    sync();
    // 功能：测试环境可能没有 MutationObserver / matchMedia，缺失时只做一次性读取；时间：2026-10-31；作者：lq。
    const observer = typeof MutationObserver === 'function' ? new MutationObserver(sync) : null;
    observer?.observe(root, { attributes: true, attributeFilter: ['data-theme'] });
    // 系统主题跟随模式下 data-theme 也会被改写，这里再兜一层媒体查询变化。
    const media = typeof window.matchMedia === 'function' ? window.matchMedia('(prefers-color-scheme: dark)') : null;
    media?.addEventListener('change', sync);
    return () => {
      observer?.disconnect();
      media?.removeEventListener('change', sync);
    };
  }, []);
  return theme;
}
