// 功能：验证页眉页脚动态域、起始页码和 DOCX 新建/修改/清空往返；时间：2026-10-24；作者：lq。
import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { createBlankDoc } from '@mini-office/document-model';
import { loadOoxmlPackage } from '@mini-office/ooxml';
import { makeHeaderFooter, headerFooterText, PAGE_NUMBER_TEMPLATES } from '../packages/editor-word/src/header-footer';
import { parseDocxPackage } from '../packages/editor-word/src/parse';
import { serializeDocx } from '../packages/editor-word/src/serialize';
import { renderParagraphHtml } from '../packages/editor-word/src/html';

/** 功能：通过真实 OOXML 包解析回读保存结果；时间：2026-10-24；作者：lq。 */
async function reopen(bytes: Uint8Array) {
  const pkg = await loadOoxmlPackage(bytes);
  try { return await parseDocxPackage(pkg); } finally { pkg.dispose(); }
}

describe('页眉页脚与页码', () => {
  it('空区域先选对齐和字号仍能保存设置，避免输入文字后恢复默认值', () => {
    const value = makeHeaderFooter('', 'right', 16, 38);
    expect(value.blocks[0]).toMatchObject({ style: { alignment: 'right' }, runs: [{ text: '', style: { fontSizePt: 16 } }] });
    expect(headerFooterText(value)).toBe('');
  });
  it('所有预设都使用动态域，预览按当前页与总页数解析并转义文字', () => {
    for (const template of PAGE_NUMBER_TEMPLATES) {
      const value = makeHeaderFooter(template.value, 'center', 10.5, 48);
      expect(headerFooterText(value)).toBe(template.value);
      for (const block of value.blocks) if (block.type === 'paragraph') {
        const html = renderParagraphHtml(block, { pageNumber: 7, totalPages: 12 });
        expect(html).not.toContain('{PAGE}');
        if (template.value.includes('{NUMPAGES}')) expect(html).toContain('12');
        if (template.value.includes('{PAGE}')) expect(html).toContain('7');
      }
    }
    const value = makeHeaderFooter('<script> & {PAGE}', 'left', 12, 30);
    const block = value.blocks[0];
    expect(block.type === 'paragraph' && renderParagraphHtml(block)).toContain('&lt;script&gt;');
  });

  it('新文档保存有完整关系、内容类型和可回读的页码域', async () => {
    const doc = createBlankDoc();
    doc.header = makeHeaderFooter('公司 & 报告\n内部资料', 'left', 12, 38);
    doc.footer = makeHeaderFooter('第 {PAGE} 页 / 共 {NUMPAGES} 页', 'right', 10.5, 48);
    doc.page.pageNumberStart = 5;
    doc.page.columns = 2;
    doc.page.margins.header = 38;
    const bytes = await serializeDocx(doc, { name: 'headers.docx' });
    const zip = await JSZip.loadAsync(bytes);
    expect(await zip.file('[Content_Types].xml')!.async('text')).toContain('wordprocessingml.header+xml');
    expect(await zip.file('word/_rels/document.xml.rels')!.async('text')).toContain('mini-footer.xml');
    const saved = await reopen(bytes);
    expect(headerFooterText(saved.header)).toBe('公司 & 报告\n内部资料');
    expect(headerFooterText(saved.footer)).toBe('第 {PAGE} 页 / 共 {NUMPAGES} 页');
    expect(saved.page.pageNumberStart).toBe(5);
    expect(saved.page.columns).toBe(2);
    expect(saved.page.margins.header).toBeCloseTo(38, 1);
    expect(saved.footer?.blocks[0]).toMatchObject({ style: { alignment: 'right' } });
  });

  it('修改导入合同的页脚，保留原页眉与图片；再清空后不会复活旧页脚', async () => {
    URL.createObjectURL ??= () => 'blob:test';
    URL.revokeObjectURL ??= () => {};
    const sourceBytes = new Uint8Array(readFileSync('tests/fixtures/docx/contract.docx'));
    const doc = await reopen(sourceBytes);
    const header = headerFooterText(doc.header);
    doc.footer = makeHeaderFooter('— {PAGE} —', 'center', 9, 48);
    const bytes = await serializeDocx(doc, { sourceBytes, name: 'edited.docx' });
    const saved = await reopen(bytes);
    expect(headerFooterText(saved.header)).toBe(header);
    expect(headerFooterText(saved.footer)).toBe('— {PAGE} —');
    expect(Object.keys(saved.images).length).toBeGreaterThan(0);
    saved.footer = makeHeaderFooter('', 'center', 9, 48);
    const cleared = await reopen(await serializeDocx(saved, { sourceBytes: bytes, name: 'clear.docx' }));
    expect(headerFooterText(cleared.footer)).toBe('');
    expect(headerFooterText(cleared.header)).toBe(header);
  });
});
