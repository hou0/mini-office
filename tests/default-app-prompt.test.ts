/** 功能：验证首次启动默认应用提示只弹一次，且不抢占已有弹窗；时间：2026-10-29；作者：lq。 */
import { beforeEach, describe, expect, it, vi } from 'vitest';

const mocks = vi.hoisted(() => ({
  runtime: vi.fn(),
  status: vi.fn(),
  request: vi.fn(),
  openSettings: vi.fn(),
}));

vi.mock('@mini-office/platform-adapter', () => ({
  isTauriRuntime: mocks.runtime,
  getDefaultAppStatus: mocks.status,
  requestDefaultApp: mocks.request,
  openDefaultAppSettings: mocks.openSettings,
  getPlatformAdapter: () => null,
}));

import { useWorkspaceStore } from '@mini-office/document-model';
import { maybePromptDefaultApp } from '../packages/editor-shell/src/default-app';

function statusReport(overrides: Record<string, unknown> = {}) {
  return {
    platform: 'windows',
    appName: 'Mini Office',
    canSetProgrammatically: false,
    entries: [
      { extension: 'docx', label: 'Word 文档', isDefault: false },
      { extension: 'xlsx', label: 'Excel 工作簿', isDefault: false },
      { extension: 'pptx', label: 'PowerPoint 演示文稿', isDefault: false },
      { extension: 'csv', label: 'CSV 文本', isDefault: false },
    ],
    allDefault: false,
    defaultCount: 0,
    detail: 'Windows 不允许程序静默修改默认应用。',
    ...overrides,
  };
}

beforeEach(() => {
  vi.resetAllMocks();
  localStorage.clear();
  mocks.runtime.mockReturnValue(true);
  mocks.status.mockResolvedValue(statusReport());
  useWorkspaceStore.setState({ dialog: null });
});

describe('首次启动默认应用提示', () => {
  it('未设置默认应用时弹出提示', async () => {
    await maybePromptDefaultApp();
    expect(useWorkspaceStore.getState().dialog).toEqual({ type: 'default-app', payload: { prompt: true } });
  });

  it('已全部设为默认时静默跳过并记录标记', async () => {
    mocks.status.mockResolvedValue(statusReport({ allDefault: true, defaultCount: 4 }));
    await maybePromptDefaultApp();
    expect(useWorkspaceStore.getState().dialog).toBeNull();
    expect(localStorage.getItem('mini-office.default-app-prompt')).not.toBeNull();
  });

  it('已提示过时不再查询系统状态', async () => {
    localStorage.setItem('mini-office.default-app-prompt', '1');
    await maybePromptDefaultApp();
    expect(mocks.status).not.toHaveBeenCalled();
    expect(useWorkspaceStore.getState().dialog).toBeNull();
  });

  it('已有弹窗时不覆盖用户正在进行的操作', async () => {
    useWorkspaceStore.setState({ dialog: { type: 'settings' } });
    await maybePromptDefaultApp();
    expect(useWorkspaceStore.getState().dialog).toEqual({ type: 'settings' });
  });

  it('Web 环境不提示默认应用', async () => {
    mocks.runtime.mockReturnValue(false);
    await maybePromptDefaultApp();
    expect(mocks.status).not.toHaveBeenCalled();
    expect(useWorkspaceStore.getState().dialog).toBeNull();
  });
});
