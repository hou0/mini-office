import { Component, StrictMode, type ErrorInfo, type ReactNode } from 'react';
import { createRoot } from 'react-dom/client';
import { useWorkspaceStore } from '@mini-office/document-model';
import { applyAppearance } from '@mini-office/design-system';
// 功能：首屏只加载格式代理和启动状态，移除三种编辑器的同步依赖；时间：2026-10-24；作者：lq。
import { registerDeferredFormats } from './format-handlers';
import { dismissStartup } from './startup';
import App from './App';
import './styles.css';

registerDeferredFormats();

class ErrorBoundary extends Component<{ children: ReactNode }, { error: Error | null }> {
  state = { error: null as Error | null };

  static getDerivedStateFromError(error: Error) {
    return { error };
  }

  componentDidCatch(error: Error, info: ErrorInfo) {
    // 功能：发生界面异常时展示可操作的错误页；时间：2026-10-24；作者：lq。
    dismissStartup();
    console.error('[mini-office] uncaught error', error, info);
  }

  render() {
    if (this.state.error) {
      return (
        <div className="flex h-full flex-col items-center justify-center gap-4 bg-[var(--mo-bg)] p-8">
          <div className="text-[15px] font-semibold text-[var(--mo-text-primary)]">界面遇到问题</div>
          <p className="max-w-md text-center text-[12px] leading-5 text-[var(--mo-text-muted)]">
            {this.state.error.message}
          </p>
          <p className="text-[11px] text-[var(--mo-text-muted)]">
            已保存的文档不会丢失。刷新页面可从自动恢复快照继续。
          </p>
          <button
            type="button"
            onClick={() => window.location.reload()}
            className="rounded-md bg-[var(--mo-accent)] px-4 py-2 text-[12px] font-medium text-[var(--mo-accent-ink)]"
          >
            刷新页面
          </button>
        </div>
      );
    }
    return this.props.children;
  }
}

// Apply the persisted appearance before the first React paint to avoid a
// flash of the default accent while the workspace hydrates.
applyAppearance(useWorkspaceStore.getState().settings);

const container = document.getElementById('root');
if (container) {
  createRoot(container).render(
    <StrictMode>
      <ErrorBoundary>
        <App />
      </ErrorBoundary>
    </StrictMode>,
  );
}
