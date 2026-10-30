// 验证主画布预览模式的进入、勾选联动与 ESC 退出
import { Fragment, act, createElement, StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { expect, it, vi } from 'vitest';
import { createBlankDoc, createBlankWorkbook, useWorkspaceStore, type WordSession, type SheetSession } from '@mini-office/document-model';
import { MailMergePanel } from '../packages/editor-word/src/MailMergePanel';
import { MailMergeCanvas, MergePreviewBar } from '../packages/editor-word/src/MailMergePreview';
import { useMailMergeController } from '../packages/editor-word/src/mail-merge-controller';

it('进入预览显示画布与控制条，取消勾选联动计数，ESC 退出', async () => {
  (globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
  vi.stubGlobal('ResizeObserver', class { observe() {} disconnect() {} });
  const doc = createBlankDoc();
  if (doc.blocks[0].type === 'paragraph') doc.blocks[0].runs = [{ mergeField: '姓名', text: '«姓名»', style: {} }];
  const workbook = createBlankWorkbook();
  workbook.sheets[0].cells = { '0:0': { value: '姓名' }, '1:0': { value: '张三' }, '2:0': { value: '李四' } };
  const session = { id: 'preview-test-doc', kind: 'docx', displayName: '模板.docx', data: doc, readonly: false } as WordSession;
  const sheet = { id: 'preview-test-sheet', kind: 'xlsx', displayName: '数据.xlsx', data: workbook } as SheetSession;
  useWorkspaceStore.setState({ sessions: { [session.id]: session, [sheet.id]: sheet } });
  const host = document.createElement('div'); document.body.appendChild(host);
  const root = createRoot(host);
  const Harness = (props: { session: WordSession; doc: typeof doc }) => {
    const c = useMailMergeController({ session: props.session, doc: props.doc, renderBlocks: () => null });
    return createElement(Fragment, null,
      createElement(MailMergePanel, { open: true, session: props.session, doc: props.doc, onClose: () => {}, onInsertField: () => {}, renderBlocks: () => null, controller: c }),
      c.previewOn ? createElement(Fragment, null,
        createElement(MailMergeCanvas, { controller: c, doc: props.doc }),
        createElement(MergePreviewBar, { controller: c }),
      ) : null,
    );
  };
  const click = async (text: string) => { const button = [...document.querySelectorAll('button')].find((item) => item.textContent === text); expect(button).toBeTruthy(); await act(async () => { button!.click(); }); };
  const waitFor = async (check: () => boolean, message: string) => { for (let i = 0; i < 40; i++) { if (check()) return; await act(async () => { await new Promise((resolve) => setTimeout(resolve, 50)); }); } throw new Error(message); };
  try {
    await act(async () => { root.render(createElement(StrictMode, null, createElement(Harness, { session, doc }))); });
    const select = document.querySelector<HTMLSelectElement>('[aria-label="已打开的数据表"]')!;
    await act(async () => { select.value = sheet.id; select.dispatchEvent(new Event('change', { bubbles: true })); });
    await click('使用');
    await click('展开筛选窗口（可调整大小）');
    await click('预览合并效果');
    await waitFor(() => !!document.querySelector('[data-mail-merge-canvas]') && !!document.querySelector('[data-mail-merge-bar]'), '预览模式未进入');
    expect(document.querySelector('[data-mail-merge-bar]')!.textContent).toContain('第 1 / 2 页');
    const row2 = document.querySelector<HTMLInputElement>('[aria-label="选择第 2 行"]')!;
    await act(async () => { row2.click(); });
    expect(document.querySelector('[data-mail-merge-bar]')!.textContent).toContain('第 1 / 1 页');
    expect(document.querySelector('[data-mail-merge-bar]')!.textContent).toContain('数据第 3 行');
    await act(async () => { window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' })); });
    await waitFor(() => !document.querySelector('[data-mail-merge-bar]'), 'ESC 未退出预览模式');
  } finally {
    await act(async () => root.unmount()); host.remove();
    useWorkspaceStore.setState({ sessions: {} }); vi.unstubAllGlobals();
  }
});
