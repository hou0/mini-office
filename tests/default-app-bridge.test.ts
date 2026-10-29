/** 功能：验证默认应用桥接的状态查询、申请结果透传与 Web 降级；时间：2026-10-29；作者：lq。 */
import { beforeEach, describe, expect, it, vi } from 'vitest';

const mocks = vi.hoisted(() => ({ invoke: vi.fn(), runtime: vi.fn() }));
vi.mock('../packages/platform-adapter/node_modules/@tauri-apps/api/core.js', () => ({ invoke: mocks.invoke }));
vi.mock('../packages/platform-adapter/src/types', () => ({ isTauriRuntime: mocks.runtime }));

import {
  WEB_DEFAULT_APP_STATUS,
  getDefaultAppStatus,
  openDefaultAppSettings,
  requestDefaultApp,
} from '../packages/platform-adapter/src/default-app';

const STATUS = {
  platform: 'windows' as const,
  appName: 'Mini Office',
  canSetProgrammatically: false,
  entries: [{ extension: 'docx', label: 'Word 文档', isDefault: true }],
  allDefault: false,
  defaultCount: 1,
  detail: 'Windows 不允许程序静默修改默认应用。',
};

beforeEach(() => {
  vi.resetAllMocks();
  mocks.runtime.mockReturnValue(true);
});

describe('默认应用桥接', () => {
  it('Web 环境返回只读占位状态，不调用原生命令', async () => {
    mocks.runtime.mockReturnValue(false);
    await expect(getDefaultAppStatus()).resolves.toBe(WEB_DEFAULT_APP_STATUS);
    const outcome = await requestDefaultApp();
    expect(outcome.outcome).toBe('unsupported');
    await expect(openDefaultAppSettings()).rejects.toThrow('桌面客户端');
    expect(mocks.invoke).not.toHaveBeenCalled();
  });

  it('桌面端透传系统状态', async () => {
    mocks.invoke.mockResolvedValue(STATUS);
    await expect(getDefaultAppStatus()).resolves.toEqual(STATUS);
    expect(mocks.invoke).toHaveBeenCalledWith('default_app_status');
  });

  it('状态查询失败时返回可展示的兜底结果', async () => {
    mocks.invoke.mockRejectedValue('permission denied');
    const status = await getDefaultAppStatus();
    expect(status.entries.length).toBe(4);
    expect(status.allDefault).toBe(false);
    expect(status.detail).toContain('无法读取默认应用状态');
  });

  it('申请默认应用返回结果与最新状态', async () => {
    mocks.invoke.mockResolvedValue({ outcome: 'settings', message: '已打开系统设置', status: STATUS });
    const result = await requestDefaultApp();
    expect(mocks.invoke).toHaveBeenCalledWith('request_default_app');
    expect(result.outcome).toBe('settings');
    expect(result.status).toEqual(STATUS);
  });

  it('申请异常时收敛为 error 结果而不是抛出', async () => {
    mocks.invoke.mockRejectedValue(new Error('boom'));
    const result = await requestDefaultApp();
    expect(result.outcome).toBe('error');
    expect(result.message).toContain('boom');
  });

  it('打开系统设置走独立命令', async () => {
    mocks.invoke.mockResolvedValue(null);
    await openDefaultAppSettings();
    expect(mocks.invoke).toHaveBeenCalledWith('open_default_app_settings');
  });
});
