// 功能：NEXT 模板预览按页显示——渲染当前记录所在的那一页模板（每页几个块就填几条），用「上一页 / 下一页」逐页切换；时间：2026-10-30；作者：lq。
import { Fragment, act, createElement, StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { expect, it, vi } from 'vitest';
import { createBlankDoc, createBlankWorkbook, useWorkspaceStore, type WordSession, type SheetSession } from '@mini-office/document-model';
import { MailMergePanel } from '../packages/editor-word/src/MailMergePanel';
import { MailMergeCanvas, MergePreviewBar } from '../packages/editor-word/src/MailMergePreview';
import { useMailMergeController, type MailMergeController } from '../packages/editor-word/src/mail-merge-controller';

function makeRun(text: string, extra: Record<string, unknown> = {}): any {
  return { text, style: {}, ...extra };
}
function makeCell(runs: any[]): any {
  return {
    id: `c${Math.random().toString(36).slice(2)}`,
    blocks: [{ type: 'paragraph', id: `p${Math.random().toString(36).slice(2)}`, runs, style: {} }],
    colSpan: 1,
    rowSpan: 1,
  };
}
/** 每个记录块两行：一行带 NEXT 域，一行普通数据行 */
function recordBlock(tag: string): any[] {
  return [
    {
      id: `r${tag}a`,
      cells: [makeCell([makeRun('«NEXT»', { mergeNext: true }), makeRun(`«名称${tag}»`, { mergeField: '名称' })])],
    },
    {
      id: `r${tag}b`,
      cells: [makeCell([makeRun(`«数量${tag}»`, { mergeField: '数量' })])],
    },
  ];
}
function makeNextTemplate(): any {
  const doc = createBlankDoc();
  doc.blocks = [
    {
      type: 'table',
      id: 't1',
      columnWidthsPx: [200],
      rows: [
        { id: 'h', cells: [makeCell([makeRun('名称', {})])] },
        ...recordBlock('1'),
        ...recordBlock('2'),
      ],
    },
  ];
  return doc;
}
function renderRuns(paragraph: any): any {
  return paragraph.runs.map((run: any, index: number) => createElement('span', { key: index }, run.text));
}
function renderCell(cell: any, cellIndex: number): any {
  return createElement('td', { key: cellIndex }, cell.blocks.map((paragraph: any) => createElement('p', { key: paragraph.id }, renderRuns(paragraph))));
}
function renderRow(row: any): any {
  return createElement('tr', { key: row.id }, row.cells.map((cell: any, cellIndex: number) => renderCell(cell, cellIndex)));
}
function renderBlock(block: any): any {
  if (block.type === 'paragraph') return createElement('p', { key: block.id }, renderRuns(block));
  if (block.type === 'table') return createElement('table', { key: block.id }, createElement('tbody', null, block.rows.map((row: any) => renderRow(row))));
  return null;
}

it('NEXT 模板预览按页显示当前记录所在的那一页模板', async () => {
  (globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
  vi.stubGlobal('ResizeObserver', class { observe() {} disconnect() {} });
  const doc = makeNextTemplate();
  const workbook = createBlankWorkbook();
  // 3 条记录：模板每页 2 个记录块，全量合并会分成 2 页
  workbook.sheets[0].cells = {
    '0:0': { value: '名称' }, '0:1': { value: '数量' },
    '1:0': { value: '张三' }, '1:1': { value: '1' },
    '2:0': { value: '李四' }, '2:1': { value: '2' },
    '3:0': { value: '王五' }, '3:1': { value: '3' },
  };
  const session = { id: 'next-preview-doc', kind: 'docx', displayName: '模板.docx', data: doc, readonly: false } as WordSession;
  const sheet = { id: 'next-preview-sheet', kind: 'xlsx', displayName: '数据.xlsx', data: workbook } as SheetSession;
  useWorkspaceStore.setState({ sessions: { [session.id]: session, [sheet.id]: sheet } });
  const host = document.createElement('div'); document.body.appendChild(host);
  const root = createRoot(host);
  const ref: { current: MailMergeController | null } = { current: null };
  // 固定引用：每次渲染新建函数会让分页 effect 反复重跑，预览永远算不出结果
  const renderBlocks = (blocks: any[]): any => createElement('div', null, blocks.map((block: any) => renderBlock(block)));
  const Harness = (props: { session: WordSession; doc: typeof doc }) => {
    const c = useMailMergeController({ session: props.session, doc: props.doc, renderBlocks });
    ref.current = c;
    return createElement(Fragment, null,
      createElement(MailMergePanel, { open: true, session: props.session, doc: props.doc, onClose: () => {}, onInsertField: () => {}, renderBlocks, controller: c }),
      c.previewOn ? createElement(Fragment, null,
        createElement(MailMergeCanvas, { controller: c, doc: props.doc }),
        createElement(MergePreviewBar, { controller: c }),
      ) : null,
    );
  };
  const click = async (text: string) => { const button = [...document.querySelectorAll('button')].find((item) => item.textContent === text); expect(button).toBeTruthy(); await act(async () => { button!.click(); }); };
  const waitFor = async (check: () => boolean, message: string) => { for (let i = 0; i < 60; i++) { if (check()) return; await act(async () => { await new Promise((resolve) => setTimeout(resolve, 50)); }); } throw new Error(message); };
  try {
    await act(async () => { root.render(createElement(StrictMode, null, createElement(Harness, { session, doc }))); });
    const select = document.querySelector<HTMLSelectElement>('[aria-label="已打开的数据表"]')!;
    await act(async () => { select.value = sheet.id; select.dispatchEvent(new Event('change', { bubbles: true })); });
    await click('使用');
    // 全量合并 3 条记录会分成 2 页，预览只应显示其中 1 页
    const { mergeNextDocument } = await import('../packages/editor-word/src/mail-merge');
    const full = mergeNextDocument(doc, ref.current!.chosen);
    expect(full.blocks.filter((block) => block.type === 'pageBreak')).toHaveLength(1);
    await click('预览合并效果');
    await waitFor(() => !!document.querySelector('[data-mail-merge-canvas]')?.textContent?.includes('张三'), '预览未渲染当前记录');
    // 预览只渲染一页模板：没有分页符，只有一个表格
    const blocks = ref.current!.previewDoc!.blocks;
    expect(blocks.filter((block) => block.type === 'pageBreak')).toHaveLength(0);
    expect(blocks.filter((block) => block.type === 'table')).toHaveLength(1);
    // 每页 2 个记录块：第 1 页同时显示第 1、2 条记录，而不是只填当前一条
    expect(document.querySelector('[data-mail-merge-canvas]')!.textContent).toContain('李四');
    expect(document.querySelector('[data-mail-merge-bar]')!.textContent).toContain('第 1 / 2 页');
    expect(document.querySelector('[data-mail-merge-bar]')!.textContent).toContain('本页第 1~2 条');
    // 下一页翻到第 2 页：该页只有第 3 条记录
    await click('下一页');
    expect(document.querySelector('[data-mail-merge-bar]')!.textContent).toContain('第 2 / 2 页');
    expect(document.querySelector('[data-mail-merge-bar]')!.textContent).toContain('本页第 3~3 条');
    await waitFor(() => !!document.querySelector('[data-mail-merge-canvas]')?.textContent?.includes('王五'), '跨页后预览未更新');
    expect(document.querySelector('[data-mail-merge-canvas]')!.textContent).not.toContain('张三');
    // 上一页回到第 1 页，3 条记录重新填满两个块
    await click('上一页');
    expect(document.querySelector('[data-mail-merge-bar]')!.textContent).toContain('第 1 / 2 页');
    await waitFor(() => !!document.querySelector('[data-mail-merge-canvas]')?.textContent?.includes('张三'), '返回上一页后预览未更新');
    expect(document.querySelector('[data-mail-merge-canvas]')!.textContent).toContain('李四');
    expect(ref.current!.previewDoc!.blocks.filter((block) => block.type === 'table')).toHaveLength(1);
  } finally {
    await act(async () => root.unmount()); host.remove();
    useWorkspaceStore.setState({ sessions: {} }); vi.unstubAllGlobals();
  }
});
