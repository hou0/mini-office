/** 功能：验证原生退出在多文档、失败、取消和并发编辑时不丢内容；时间：2026-10-18；作者：lq。 */
import { describe, expect, it, vi } from 'vitest';
import { createBlankDoc, type DocumentSession } from '@mini-office/document-model';
import { confirmWorkspaceClose } from '../packages/editor-shell/src/close-workspace';

/** 功能：创建最小未保存文档供退出保护回归；时间：2026-10-18；作者：lq。 */
function dirty(id: string): DocumentSession {
  return { id, data: createBlankDoc(), dirty: true, state: 'DIRTY' } as DocumentSession;
}

describe('原生退出保护', () => {
  it('无修改时直接完成，不弹窗', async () => {
    const ask = vi.fn();
    expect(await confirmWorkspaceClose(() => [], ask, vi.fn())).toBe(true);
    expect(ask).not.toHaveBeenCalled();
  });
  it('多文档逐份确认，全部完成前保留标签', async () => {
    const sessions = [dirty('a'), dirty('b')];
    const ask = vi.fn().mockResolvedValueOnce('save').mockResolvedValueOnce('discard');
    const save = vi.fn(async (id) => { sessions.find((item) => item.id === id)!.dirty = false; return true; });
    expect(await confirmWorkspaceClose(() => sessions, ask, save)).toBe(true);
    expect(ask).toHaveBeenCalledTimes(2);
    expect(save).toHaveBeenCalledWith('a');
    expect(sessions).toHaveLength(2);
    expect(sessions[1].dirty).toBe(true);
  });
  it.each(['cancel', 'save'] as const)('取消或保存失败保持打开：%s', async (choice) => {
    expect(await confirmWorkspaceClose(() => [dirty('a')], async () => choice, async () => false)).toBe(false);
  });
  it('保存过程中新增编辑，必须再次确认', async () => {
    const session = dirty('a');
    const ask = vi.fn().mockResolvedValueOnce('save').mockResolvedValueOnce('cancel');
    const save = vi.fn(async () => { session.data = createBlankDoc(); return true; });
    expect(await confirmWorkspaceClose(() => [session], ask, save)).toBe(false);
    expect(ask).toHaveBeenCalledTimes(2);
  });
  it('先前选择不保存的文档再次编辑后也须重新确认', async () => {
    const sessions = [dirty('a'), dirty('b')];
    const ask = vi.fn().mockResolvedValueOnce('discard').mockImplementationOnce(async () => {
      sessions[0].data = createBlankDoc(); return 'discard';
    }).mockResolvedValueOnce('cancel');
    expect(await confirmWorkspaceClose(() => sessions, ask, vi.fn())).toBe(false);
    expect(ask).toHaveBeenCalledTimes(3);
  });
  it('已有保存/导出任务时不退出', async () => {
    const session = dirty('a');
    session.state = 'SAVING';
    const ask = vi.fn();
    expect(await confirmWorkspaceClose(() => [session], ask, vi.fn())).toBe(false);
    expect(ask).not.toHaveBeenCalled();
  });
});
