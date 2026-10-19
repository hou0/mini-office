/** 功能：验证跨段落选区提取和模型替换，不破坏样式与表格；时间：2026-10-18；作者：lq。 */
import { describe, expect, it } from 'vitest';
import { createBlankDoc, type DocTable } from '@mini-office/document-model';
import { makeParagraph, collectParagraphs } from '../packages/editor-word/src/doc-utils';
import { selectedParagraphSlices, runsPlainText } from '../packages/editor-word/src/editing';
import { replaceDocumentSelection, selectionHtml, selectionText } from '../packages/editor-word/src/document-selection';

/** 功能：创建多段正文并含不同 run 样式；时间：2026-10-18；作者：lq。 */
function sample() {
  const doc = createBlankDoc();
  doc.blocks = [makeParagraph([{ text: '第一段内容', style: { bold: true } }]), makeParagraph([{ text: '第二段内容', style: {} }]), makeParagraph([{ text: '第三段内容', style: { italic: true } }])];
  return doc;
}

describe('Word 多段选区', () => {
  it('跨段 DOM 选区按原子偏移提取，排除页眉', () => {
    const root = document.createElement('div');
    root.innerHTML = '<div contenteditable="false">页眉</div><div data-paragraph-id="a"><span>第一段</span></div><div data-paragraph-id="b"><b>第二段</b></div>';
    document.body.append(root);
    const range = document.createRange();
    range.setStart(root.querySelector('[data-paragraph-id="a"] span')!.firstChild!, 1);
    range.setEnd(root.querySelector('[data-paragraph-id="b"] b')!.firstChild!, 2);
    expect(selectedParagraphSlices(root, range)).toEqual([{ paragraphId: 'a', start: 1, end: 3 }, { paragraphId: 'b', start: 0, end: 2 }]);
    root.remove();
  });
  it('跨段复制只包含选中部分，排除段外文字', () => {
    const doc = sample();
    const paragraphs = collectParagraphs(doc);
    expect(selectionText(doc, [{ paragraphId: paragraphs[0].id, start: 2, end: 5 }, { paragraphId: paragraphs[1].id, start: 0, end: 2 }])).toBe('段内容\n第二');
  });
  it('富文本剪贴板保留选中文字样式，排除分页容器', () => {
    const doc = sample();
    const paragraphs = collectParagraphs(doc);
    const html = selectionHtml(doc, [{ paragraphId: paragraphs[0].id, start: 2, end: 5 }]);
    expect(html).toContain('font-weight:700');
    expect(html).toContain('段内容');
    expect(html).not.toContain('第一');
    expect(html).not.toContain('data-word-page');
  });
  it('跨段替换合并首尾并保留两端样式，原模型不变', () => {
    const doc = sample();
    const paragraphs = collectParagraphs(doc);
    const slices = [{ paragraphId: paragraphs[0].id, start: 2, end: 5 }, { paragraphId: paragraphs[1].id, start: 0, end: 5 }, { paragraphId: paragraphs[2].id, start: 0, end: 2 }];
    const result = replaceDocumentSelection(doc, slices, '新文字')!;
    const changed = collectParagraphs(result.doc);
    expect(changed).toHaveLength(1);
    expect(runsPlainText(changed[0].runs)).toBe('第一新文字段内容');
    expect(changed[0].runs[0].style.bold).toBe(true);
    expect(changed[0].runs.at(-1)?.style.italic).toBe(true);
    expect(result.offset).toBe(5);
    expect(collectParagraphs(doc)).toHaveLength(3);
  });
  it('全选删除后留可编辑空段落，页面配置不变', () => {
    const doc = sample();
    const slices = collectParagraphs(doc).map((p) => ({ paragraphId: p.id, start: 0, end: runsPlainText(p.runs).length }));
    const result = replaceDocumentSelection(doc, slices, '')!;
    expect(result.doc.blocks).toHaveLength(1);
    expect(runsPlainText(collectParagraphs(result.doc)[0].runs)).toBe('');
    expect(result.doc.page).toBe(doc.page);
  });
  it('跨段粘贴多行内容生成真实段落', () => {
    const doc = sample();
    const slices = collectParagraphs(doc).map((p) => ({ paragraphId: p.id, start: 0, end: runsPlainText(p.runs).length }));
    const result = replaceDocumentSelection(doc, slices, '甲\n乙\n丙')!;
    expect(collectParagraphs(result.doc).map((p) => runsPlainText(p.runs))).toEqual(['甲', '乙', '丙']);
    expect(result.paragraphId).toBe(result.doc.blocks.at(-1)!.id);
  });
  it('跨单元格部分替换保留表格、未选中的后缀及其他正文', () => {
    const doc = sample();
    const p1 = makeParagraph([{ text: '甲方姓名', style: {} }]);
    const p2 = makeParagraph([{ text: '乙方姓名', style: {} }]);
    const table: DocTable = { type: 'table', id: 'table', columnWidthsPx: [100, 100], rows: [{ id: 'r', cells: [p1, p2].map((p, i) => ({ id: `c${i}`, colSpan: 1, rowSpan: 1, blocks: [p] })) }] };
    doc.blocks.push(table);
    const result = replaceDocumentSelection(doc, [{ paragraphId: p1.id, start: 1, end: 4 }, { paragraphId: p2.id, start: 0, end: 1 }], '新')!;
    const changed = result.doc.blocks.at(-1) as DocTable;
    expect(changed.rows[0].cells).toHaveLength(2);
    expect(collectParagraphs({ ...doc, blocks: [changed] }).map((p) => runsPlainText(p.runs))).toEqual(['甲新', '方姓名']);
    expect(result.doc.blocks[0]).toBe(doc.blocks[0]);
  });
});
