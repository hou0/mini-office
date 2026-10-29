/** 功能：桌面端默认应用状态查询与申请桥接，Web 环境返回只读占位状态；时间：2026-10-29；作者：lq。 */
import { invoke } from '@tauri-apps/api/core';
import { isTauriRuntime } from './types';

export type DefaultAppPlatform = 'windows' | 'macos' | 'linux' | 'unknown' | 'web';

/** applied 已生效；settings 已打开系统设置；manual 需用户手动确认；unsupported/error 失败。 */
export type DefaultAppOutcomeKind = 'applied' | 'settings' | 'manual' | 'unsupported' | 'error';

export interface AssociationStatus {
  extension: string;
  label: string;
  isDefault: boolean;
}

export interface DefaultAppStatus {
  platform: DefaultAppPlatform;
  appName: string;
  /** Windows 恒为 false：系统禁止程序静默修改默认应用。 */
  canSetProgrammatically: boolean;
  entries: AssociationStatus[];
  allDefault: boolean;
  defaultCount: number;
  detail: string;
}

export interface DefaultAppOutcome {
  outcome: DefaultAppOutcomeKind;
  message: string;
  status: DefaultAppStatus;
}

/** 与 Rust 端 AssociationTarget 保持一致的展示顺序，Web 占位与前端兜底共用。 */
export const DEFAULT_APP_EXTENSIONS: Array<{ extension: string; label: string }> = [
  { extension: 'docx', label: 'Word 文档' },
  { extension: 'xlsx', label: 'Excel 工作簿' },
  { extension: 'pptx', label: 'PowerPoint 演示文稿' },
  { extension: 'csv', label: 'CSV 文本' },
];

export const WEB_DEFAULT_APP_STATUS: DefaultAppStatus = {
  platform: 'web',
  appName: 'Mini Office',
  canSetProgrammatically: false,
  entries: DEFAULT_APP_EXTENSIONS.map((item) => ({ ...item, isDefault: false })),
  allDefault: false,
  defaultCount: 0,
  detail: '默认应用设置仅在桌面客户端可用。',
};

function fallbackStatus(detail: string): DefaultAppStatus {
  return { ...WEB_DEFAULT_APP_STATUS, platform: isTauriRuntime() ? 'unknown' : 'web', detail };
}

/** 功能：读取当前系统的默认应用状态，失败时返回可展示的兜底结果；时间：2026-10-29；作者：lq。 */
export async function getDefaultAppStatus(): Promise<DefaultAppStatus> {
  if (!isTauriRuntime()) return WEB_DEFAULT_APP_STATUS;
  try {
    return await invoke<DefaultAppStatus>('default_app_status');
  } catch (error) {
    return fallbackStatus(`无法读取默认应用状态：${error instanceof Error ? error.message : String(error)}`);
  }
}

/** 功能：申请默认应用；Windows 会先登记候选应用再打开系统设置页；时间：2026-10-29；作者：lq。 */
export async function requestDefaultApp(): Promise<DefaultAppOutcome> {
  if (!isTauriRuntime()) {
    return {
      outcome: 'unsupported',
      message: '默认应用设置仅在桌面客户端可用。',
      status: WEB_DEFAULT_APP_STATUS,
    };
  }
  try {
    return await invoke<DefaultAppOutcome>('request_default_app');
  } catch (error) {
    return {
      outcome: 'error',
      message: `设置默认应用失败：${error instanceof Error ? error.message : String(error)}`,
      status: fallbackStatus('设置默认应用失败，可在系统设置中手动选择。'),
    };
  }
}

/** 功能：打开系统默认应用设置页，目前仅 Windows 提供直达入口；时间：2026-10-29；作者：lq。 */
export async function openDefaultAppSettings(): Promise<void> {
  if (!isTauriRuntime()) throw new Error('默认应用设置仅在桌面客户端可用。');
  await invoke('open_default_app_settings');
}

/**
 * 功能：重新登记文件关联，修复右键“打开方式”选择本应用后不打开文档的问题；时间：2026-10-29；作者：lq。
 * 只登记“本应用支持这些类型”，不会改动系统默认应用，因此可安全重复执行。
 */
export async function repairFileAssociations(): Promise<string> {
  if (!isTauriRuntime()) throw new Error('文件关联修复仅在桌面客户端可用。');
  return invoke<string>('repair_file_associations');
}
