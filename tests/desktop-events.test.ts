/** 功能：验证冷启动文件领取、运行中多文件串行打开及退出桥接；时间：2026-10-18；作者：lq。 */
import { beforeEach, describe, expect, it, vi } from 'vitest';
const mocks = vi.hoisted(() => ({ invoke: vi.fn(), listen: vi.fn(), runtime: vi.fn(), stop: vi.fn() }));
vi.mock('../packages/platform-adapter/node_modules/@tauri-apps/api/core.js', () => ({ invoke: mocks.invoke }));
vi.mock('../packages/platform-adapter/node_modules/@tauri-apps/api/event.js', () => ({ listen: mocks.listen }));
vi.mock('../packages/platform-adapter/src/types', () => ({ isTauriRuntime: mocks.runtime }));
import { finishDesktopClose, watchDesktopClose, watchDesktopFiles } from '../packages/platform-adapter/src/desktop-events';

beforeEach(() => {
  vi.resetAllMocks();
  mocks.runtime.mockReturnValue(true);
  mocks.listen.mockResolvedValue(mocks.stop);
  mocks.invoke.mockResolvedValue([]);
});

describe('桌面事件桥接', () => {
  it('先监听再领取冷启动文件，支持中文及空格路径', async () => {
    mocks.invoke.mockResolvedValueOnce(['/合同/中文 合同.docx', '/表格/数据.xlsx']);
    const open = vi.fn().mockResolvedValue(null);
    const stop = await watchDesktopFiles(open);
    await vi.waitFor(() => expect(open).toHaveBeenCalledTimes(2));
    expect(mocks.listen.mock.invocationCallOrder[0]).toBeLessThan(mocks.invoke.mock.invocationCallOrder[0]);
    expect(open.mock.calls.map(([path]) => path)).toEqual(['/合同/中文 合同.docx', '/表格/数据.xlsx']);
    stop();
    expect(mocks.stop).toHaveBeenCalledOnce();
  });
  it('打开过程又收到事件时继续领取，文件打开保持串行', async () => {
    let signal: () => void = () => {};
    mocks.listen.mockImplementation(async (_name, callback) => { signal = callback; return mocks.stop; });
    mocks.invoke.mockResolvedValueOnce(['/a.docx']).mockResolvedValueOnce(['/b.xlsx']);
    let release!: () => void;
    const open = vi.fn().mockImplementationOnce(() => new Promise<void>((resolve) => { release = resolve; })).mockResolvedValue(null);
    const stop = await watchDesktopFiles(open);
    await vi.waitFor(() => expect(open).toHaveBeenCalledTimes(1));
    signal();
    expect(mocks.invoke).toHaveBeenCalledTimes(1);
    release();
    await vi.waitFor(() => expect(open).toHaveBeenCalledTimes(2));
    expect(open.mock.calls[1][0]).toBe('/b.xlsx');
    stop();
    signal();
    expect(mocks.invoke).toHaveBeenCalledTimes(2);
  });
  it('订阅退出事件，确认后才发送退出授权', async () => {
    const request = vi.fn();
    await watchDesktopClose(request);
    expect(mocks.listen).toHaveBeenCalledWith('mo:request-close', request);
    expect(mocks.invoke).not.toHaveBeenCalled();
    await finishDesktopClose();
    expect(mocks.invoke).toHaveBeenCalledWith('confirm_app_close');
  });
  it('Web 模式不调用桌面 API', async () => {
    mocks.runtime.mockReturnValue(false);
    await watchDesktopFiles(vi.fn());
    await watchDesktopClose(vi.fn());
    await finishDesktopClose();
    expect(mocks.invoke).not.toHaveBeenCalled();
    expect(mocks.listen).not.toHaveBeenCalled();
  });
});
