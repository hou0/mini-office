import { invoke } from '@tauri-apps/api/core';
import { getCurrentWindow } from '@tauri-apps/api/window';
import { isTauriRuntime, supportsWindowVibrancy } from './types';
import { WebPlatformAdapter } from './web';
import { TauriPlatformAdapter } from './tauri';
import type { PlatformAdapter } from './types';

export * from './types';
export * from './fonts';
export * from './snapshots';
export { WebPlatformAdapter } from './web';
export { TauriPlatformAdapter } from './tauri';

let adapter: PlatformAdapter | null = null;

export function getPlatformAdapter(): PlatformAdapter {
  if (adapter) return adapter;
  adapter = isTauriRuntime() ? new TauriPlatformAdapter() : new WebPlatformAdapter();
  return adapter;
}

/** Test/embedded hosts can inject a fake adapter. */
export function setPlatformAdapter(next: PlatformAdapter | null): void {
  adapter = next;
}

export function platformName(): 'web' | 'tauri' {
  return getPlatformAdapter().kind;
}

export async function startWindowDrag(): Promise<void> {
  if (!isTauriRuntime()) return;
  try {
    await getCurrentWindow().startDragging();
  } catch (error) {
    console.warn('窗口拖动失败', error);
  }
}

export async function setPresentationFullscreen(enabled: boolean): Promise<boolean> {
  if (isTauriRuntime()) {
    try {
      await getCurrentWindow().setFullscreen(enabled);
      return true;
    } catch (error) {
      console.warn('切换窗口全屏失败', error);
      return false;
    }
  }
  try {
    if (enabled) {
      if (!document.fullscreenElement && document.documentElement.requestFullscreen) {
        await document.documentElement.requestFullscreen({ navigationUI: 'hide' });
      }
      return true;
    }
    if (document.fullscreenElement) await document.exitFullscreen();
    return true;
  } catch (error) {
    console.warn('浏览器全屏请求失败', error);
    return false;
  }
}

let windowVibrancyApplied = false;

/**
 * Apply (or hide) the native macOS frosted-glass material behind the webview.
 *
 * The window itself is created transparent on macOS, so turning the effect off
 * only needs to restore opaque CSS surfaces. The native view is intentionally
 * kept alive for an instant re-enable and because older Tauri runtimes cannot
 * reliably remove it at runtime.
 */
export async function applyWindowVibrancy(enabled: boolean): Promise<boolean> {
  if (!supportsWindowVibrancy()) return false;
  if (!enabled) return true;
  if (windowVibrancyApplied) return true;
  try {
    const { Effect, EffectState } = await import('@tauri-apps/api/window');
    await getCurrentWindow().setEffects({
      effects: [Effect.UnderWindowBackground],
      state: EffectState.Active,
    });
    windowVibrancyApplied = true;
    return true;
  } catch (error) {
    console.warn('启用毛玻璃效果失败', error);
    return false;
  }
}

export async function toggleWindowMaximize(): Promise<void> {
  if (!isTauriRuntime()) return;
  try {
    await getCurrentWindow().toggleMaximize();
  } catch (error) {
    console.warn('窗口最大化失败', error);
  }
}

/** Print the current view. Tauri uses the native WKWebView/WebView2 print API. */
export async function printCurrentView(): Promise<boolean> {
  if (isTauriRuntime()) {
    try {
      await invoke('print_window');
      return true;
    } catch (error) {
      console.warn('Tauri 原生打印失败，回退到 window.print()', error);
    }
  }
  if (typeof window.print === 'function') {
    window.print();
    return true;
  }
  return false;
}

// 功能：导出桌面文档关联和退出保护桥接；时间：2026-10-18；作者：lq。
export * from './desktop-events';
