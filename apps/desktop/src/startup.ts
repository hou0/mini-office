/** 功能：管理首页首帧与启动面板的交接；时间：2026-10-24；作者：lq。 */
import { useWorkspaceStore } from '@mini-office/document-model';

/** 功能：在已提交的主界面完成绘制后移除静态启动面板，并记录可交互耗时；时间：2026-10-24；作者：lq。 */
export function finishStartup(): () => void {
  let secondFrame = 0;
  let idleTask: number | undefined;
  let fontTimer: ReturnType<typeof setTimeout> | undefined;
  const firstFrame = requestAnimationFrame(() => {
    secondFrame = requestAnimationFrame(() => {
      document.getElementById('startup-splash')?.remove();
      window.dispatchEvent(new Event('mini-office:ready'));
      if (!performance.getEntriesByName('mini-office:ready').length) {
        performance.mark('mini-office:ready');
      }
      // 功能：首页呈现后在空闲时预热字体，慢扫描不占用启动首帧；时间：2026-10-24；作者：lq。
      const warmFonts = () => {
        void useWorkspaceStore.getState().loadFonts();
      };
      if ('requestIdleCallback' in window) {
        idleTask = window.requestIdleCallback(warmFonts, { timeout: 1500 });
      } else {
        fontTimer = setTimeout(warmFonts, 100);
      }
    });
  });
  return () => {
    cancelAnimationFrame(firstFrame);
    cancelAnimationFrame(secondFrame);
    if (idleTask !== undefined) window.cancelIdleCallback(idleTask);
    if (fontTimer !== undefined) clearTimeout(fontTimer);
  };
}

/** 功能：React 错误页替换启动面板，避免异常被无限 loading 遮挡；时间：2026-10-24；作者：lq。 */
export function dismissStartup(): void {
  document.getElementById('startup-splash')?.remove();
  window.dispatchEvent(new Event('mini-office:ready'));
}
