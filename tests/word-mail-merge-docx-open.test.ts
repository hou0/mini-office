// 功能：输出方式选「合并为一个 Word 文档」时，按 0.1.4 行为直接在新标签打开合并结果，不弹保存框；时间：2026-10-30；作者：lq。
import { act, createElement, StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { expect, it, vi } from 'vitest';
import { createBlankDoc, createBlankWorkbook, useWorkspaceStore, type WordSession, type SheetSession } from '@mini-office/document-model';
import { setPlatformAdapter } from '@mini-office/platform-adapter';
import { useMailMergeController, type MailMergeController } from '../packages/editor-word/src/mail-merge-controller';

function renderBlocks(blocks: any[]): any {
  return createElement('div', null, blocks.map((block: any) => (
    block.type === 'paragraph'
      ? createElement('p', { key: block.id }, block.runs.map((run: any, index: number) => createElement('span', { key: index }, run.text)))
      : createElement('div', { key: block.id }, 'table')
  )));
}

it('合并为一个 Word 文档时直接打开新文档，不弹保存框', async () => {
  (globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
  vi.stubGlobal('ResizeObserver', class { observe() {} disconnect() {} });
  const saved: string[] = [];
  setPlatformAdapter({
    kind: 'web',
    saveBytes: async (_bytes: Uint8Array, options?: { suggestedName?: string }) => { saved.push(options?.suggestedName ?? ''); },
    openFile: async () => null,
  } as any);
  const doc = createBlankDoc();
  if (doc.blocks[0].type === 'paragraph') doc.blocks[0].runs = [{ mergeField: '姓名', text: '«姓名»', style: {} }];
  const workbook = createBlankWorkbook();
  workbook.sheets[0].cells = { '0:0': { value: '姓名' }, '1:0': { value: '张三' }, '2:0': { value: '李四' } };
  const session = { id: 'docx-open-doc', kind: 'docx', displayName: '模板.docx', data: doc, readonly: false } as WordSession;
  const sheet = { id: 'docx-open-sheet', kind: 'xlsx', displayName: '数据.xlsx', data: workbook } as SheetSession;
  useWorkspaceStore.setState({ sessions: { [session.id]: session, [sheet.id]: sheet } });
  const host = document.createElement('div'); document.body.appendChild(host);
  const root = createRoot(host);
  const ref: { current: MailMergeController | null } = { current: null };
  const opened: Array<{ doc: any; name: string }> = [];
  const Harness = (props: { session: WordSession; doc: typeof doc }) => {
    const c = useMailMergeController({
      session: props.session,
      doc: props.doc,
      renderBlocks,
      onOpenDocument: async (mergedDoc, name) => { opened.push({ doc: mergedDoc, name }); },
    });
    ref.current = c;
    return null;
  };
  const waitFor = async (check: () => boolean, message: string) => { for (let i = 0; i < 60; i++) { if (check()) return; await act(async () => { await new Promise((resolve) => setTimeout(resolve, 50)); }); } throw new Error(message); };
  try {
    await act(async () => { root.render(createElement(StrictMode, null, createElement(Harness, { session, doc }))); });
    await act(async () => { ref.current!.chooseSource(workbook, '数据.xlsx'); });
    await waitFor(() => ref.current!.chosen.length === 2, '未自动勾选记录');
    await act(async () => { ref.current!.setOutput('docx'); });
    // 不能把 generate 放进 act：它内部用独立 React root 挂载分页，act 作用域会与之互相等待
    void ref.current!.generate();
    await waitFor(() => opened.length === 1, '未打开合并后的新文档');
    // 直接打开新文档，不走保存对话框
    expect(saved).toHaveLength(0);
    expect(opened).toHaveLength(1);
    expect(opened[0].name).toContain('_合并报告');
    const text = JSON.stringify(opened[0].doc.blocks);
    expect(text).toContain('张三');
    expect(text).toContain('李四');
  } finally {
    await act(async () => root.unmount()); host.remove();
    useWorkspaceStore.setState({ sessions: {} }); vi.unstubAllGlobals();
    setPlatformAdapter(null);
  }
});
