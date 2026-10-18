import { useEffect } from 'react';
import { useWorkspaceStore } from '@mini-office/document-model';
import {
  applyWindowVibrancy,
  isTauriRuntime,
  supportsWindowVibrancy,
} from '@mini-office/platform-adapter';
import { AppShell } from '@mini-office/editor-shell';
import { applyAppearance } from '@mini-office/design-system';

export default function App() {
  const hydrated = useWorkspaceStore((state) => state.hydrated);
  const initialize = useWorkspaceStore((state) => state.initialize);
  const settings = useWorkspaceStore((state) => state.settings);
  const openPath = useWorkspaceStore((state) => state.openPath);

  useEffect(() => {
    applyAppearance(settings);
    void initialize();
    document.documentElement.dataset.platform = isTauriRuntime() ? 'tauri' : 'web';
  }, [initialize, settings.theme, settings.accent, settings.customAccent]);

  useEffect(() => {
    if (!supportsWindowVibrancy()) {
      document.documentElement.dataset.vibrancy = 'off';
      return;
    }
    if (!settings.vibrancy) {
      document.documentElement.dataset.vibrancy = 'off';
      void applyWindowVibrancy(false);
      return;
    }
    let cancelled = false;
    void applyWindowVibrancy(true).then((applied) => {
      if (cancelled) return;
      document.documentElement.dataset.vibrancy = applied ? 'on' : 'off';
      if (!applied) {
        useWorkspaceStore
          .getState()
          .notify('warning', '毛玻璃效果不可用', '当前窗口无法启用原生磨砂材质，已保持不透明背景。');
      }
    });
    return () => {
      cancelled = true;
    };
  }, [settings.vibrancy]);

  useEffect(() => {
    if (!isTauriRuntime()) return;
    let unlisten: (() => void) | undefined;
    let cancelled = false;
    void (async () => {
      try {
        const { getCurrentWebview } = await import('@tauri-apps/api/webview');
        const stop = await getCurrentWebview().onDragDropEvent((event) => {
          const payload = event.payload;
          if (payload.type === 'drop' && payload.paths.length > 0) {
            void openPath(payload.paths[0]);
          }
        });
        if (cancelled) stop();
        else unlisten = stop;
      } catch {
        // Web fallback handles drag and drop when the Tauri API is unavailable.
      }
    })();
    return () => {
      cancelled = true;
      unlisten?.();
    };
  }, [openPath]);

  if (!hydrated) {
    return (
      <div className="flex h-full flex-col items-center justify-center gap-4 bg-[var(--mo-bg)]">
        <div className="mo-brand-mark flex h-14 w-14 items-center justify-center rounded-2xl text-[24px] font-semibold shadow-pop">
          M
        </div>
        <div className="text-[13px] font-medium text-[var(--mo-text-primary)]">Mini Office</div>
        <div className="flex items-center gap-2 text-[11px] text-[var(--mo-text-muted)]">
          <span className="h-3 w-3 animate-[mo-spin_0.9s_linear_infinite] rounded-full border-2 border-[var(--mo-line-strong)] border-t-[var(--mo-accent)]" />
          正在准备本地工作区…
        </div>
      </div>
    );
  }

  return <AppShell />;
}
