/** 功能：封装桌面文件关联队列和原生退出通知，保持业务组件与 Tauri 隔离；时间：2026-10-18；作者：lq。 */
import { invoke } from '@tauri-apps/api/core';
import { listen } from '@tauri-apps/api/event';
import { isTauriRuntime } from './types';

/** 功能：先订阅再领取队列，串行打开全部文件，覆盖冷启动和运行中打开；时间：2026-10-18；作者：lq。 */
export async function watchDesktopFiles(openPath: (path: string) => Promise<unknown>): Promise<() => void> {
  if (!isTauriRuntime()) return () => {};
  let stopped = false;
  let running = false;
  let requested = false;
  const drain = async () => {
    requested = true;
    if (running || stopped) return;
    running = true;
    try {
      while (requested && !stopped) {
        requested = false;
        const paths = await invoke<string[]>('take_pending_files');
        // 一旦领取便完成本批，防止 React 重挂载期间丢失已领取文件；时间：2026-10-18；作者：lq。
        for (const path of paths) await openPath(path);
      }
    } catch (error) {
      console.error('读取系统打开文件队列失败', error);
    } finally {
      running = false;
    }
  };
  const stop = await listen('mo:open-files', () => { void drain(); });
  void drain();
  return () => { stopped = true; stop(); };
}

/** 功能：接收关窗及 Cmd+Q 请求，由业务层逐份确认后授权退出；时间：2026-10-18；作者：lq。 */
export async function watchDesktopClose(onRequest: () => void): Promise<() => void> {
  if (!isTauriRuntime()) return () => {};
  return listen('mo:request-close', onRequest);
}

/** 功能：所有文档已安全处理后退出原生应用；时间：2026-10-18；作者：lq。 */
export async function finishDesktopClose(): Promise<void> {
  if (isTauriRuntime()) await invoke('confirm_app_close');
}
