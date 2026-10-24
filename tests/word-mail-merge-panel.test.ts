// 功能：验证批量报告在严格模式、切换文档标签后保留数据和勾选；时间：2026-10-24；作者：lq。
import { act, createElement, StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { expect, it, vi } from 'vitest';
import { createBlankDoc, createBlankWorkbook, useWorkspaceStore, type WordSession, type SheetSession } from '@mini-office/document-model';
import { MailMergePanel } from '../packages/editor-word/src/MailMergePanel';

it('切换标签重建组件时保留数据快照和记录勾选，StrictMode 不会重复重置', async () => {
  (globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
  vi.stubGlobal('ResizeObserver', class { observe() {} disconnect() {} });
  const doc = createBlankDoc();
  if (doc.blocks[0].type === 'paragraph') doc.blocks[0].runs = [{ mergeField: '姓名', text: '«姓名»', style: {} }];
  const workbook = createBlankWorkbook();
  workbook.sheets[0].cells = { '0:0': { value: '姓名' }, '1:0': { value: '张三' }, '2:0': { value: '李四' } };
  const session = { id: 'panel-test-doc', kind: 'docx', displayName: '模板.docx', data: doc, readonly: false } as WordSession;
  const sheet = { id: 'panel-test-sheet', kind: 'xlsx', displayName: '数据.xlsx', data: workbook } as SheetSession;
  useWorkspaceStore.setState({ sessions: { [session.id]: session, [sheet.id]: sheet } });
  const host = document.createElement('div'); document.body.appendChild(host);
  let root = createRoot(host);
  const props = { open: true, session, doc, onClose: vi.fn(), onInsertField: vi.fn(), renderBlocks: () => null };
  const mount = async () => { await act(async () => { root.render(createElement(StrictMode, null, createElement(MailMergePanel, props))); }); };
  const click = async (text: string) => { const button = [...document.querySelectorAll('button')].find((item) => item.textContent === text); expect(button).toBeTruthy(); await act(async () => { button!.click(); }); };
  try {
    await mount();
    const select = document.querySelector<HTMLSelectElement>('[aria-label="已打开的数据表"]')!;
    await act(async () => { select.value = sheet.id; select.dispatchEvent(new Event('change', { bubbles: true })); });
    await click('使用'); await click('选择记录并预览报告');
    const first = document.querySelector<HTMLInputElement>('[aria-label="选择第 2 行"]')!;
    expect(first.checked).toBe(true);
    await act(async () => { first.click(); });
    expect(first.checked).toBe(false);
    await act(async () => root.unmount());
    root = createRoot(host);
    await mount(); await click('选择记录并预览报告');
    expect(document.querySelector<HTMLInputElement>('[aria-label="选择第 2 行"]')?.checked).toBe(false);
    expect(document.querySelector<HTMLInputElement>('[aria-label="选择第 3 行"]')?.checked).toBe(true);
  } finally {
    await act(async () => root.unmount()); host.remove();
    useWorkspaceStore.setState({ sessions: {} }); vi.unstubAllGlobals();
  }
});
