// 功能：验证结构化合并域、数据完整性、过滤及批量报告 DOCX 分节保存；时间：2026-10-24；作者：lq。
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { createBlankDoc, createBlankWorkbook, type DocModel } from '@mini-office/document-model';
import { loadOoxmlPackage } from '@mini-office/ooxml';
import { collectMergeFields, extractMergeData, filterMergeRecords, mergeDocument, mergeFileName, mergeRecordIssues, readMergeWorkbook } from '../packages/editor-word/src/mail-merge';
import { parseDocxPackage } from '../packages/editor-word/src/parse';
import { serializeDocx } from '../packages/editor-word/src/serialize';
import { makeHeaderFooter } from '../packages/editor-word/src/header-footer';
import { renderRunsHtml, renderParagraphHtml } from '../packages/editor-word/src/html';
import { paginateDocument } from '../packages/editor-word/src/pagination';
import { documentSections } from '../packages/editor-word/src/sections';
import { getCaretOffset, runsPlainText, serializeDomRuns, setCaretAtOffset } from '../packages/editor-word/src/editing';
import { splitRunsAtOffset } from '../packages/editor-word/src/doc-utils';
import { serializeXlsx } from '../packages/editor-sheet/src/serialize';

/** 功能：复用真实 OOXML 读取路径；时间：2026-10-24；作者：lq。 */
async function reopen(bytes: Uint8Array) { const pkg = await loadOoxmlPackage(bytes); try { return await parseDocxPackage(pkg); } finally { pkg.dispose(); } }
/** 功能：构造含富文本和合并字段的模板；时间：2026-10-24；作者：lq。 */
function template(): DocModel {
  const doc = createBlankDoc();
  doc.blocks = [{ type: 'paragraph', id: 'a', style: {}, runs: [{ text: '报告：', style: {} }, { mergeField: '编号', text: '«编号»', style: { bold: true, color: '#FF0000' } }, { text: '；', style: {} }, { mergeField: '姓名', text: '«姓名»', style: {} }] }];
  doc.footer = makeHeaderFooter('第 {PAGE} 页 / 共 {NUMPAGES} 页', 'center', 10.5, 48);
  return doc;
}

describe('Word 邮件合并', () => {
  it('CSV 保留前导零、长编号、换行、逗号和等号开头文本，过滤空行', async () => {
    const csv = '\uFEFF编号,姓名,结果\r\n0001,张三,"合格,正常"\r\n12345678901234567890,李四,"两行\n内容"\r\n, ,\r\n=1+1,王五,0';
    const book = await readMergeWorkbook(new TextEncoder().encode(csv), '数据.csv');
    const data = extractMergeData(book, book.sheets[0].id, 1);
    expect(data.records).toHaveLength(3);
    expect(data.records[0].values).toMatchObject({ 编号: '0001', 结果: '合格,正常' });
    expect(data.records[1].values).toMatchObject({ 编号: '12345678901234567890', 结果: '两行\n内容' });
    expect(data.records[2].values.编号).toBe('=1+1');
    expect(data.records[2].rowNumber).toBe(5);
    await expect(readMergeWorkbook(new TextEncoder().encode('a,b\n"bad,x'), 'bad.csv')).rejects.toThrow('引号');
  });

  it('XLSX 日期金额显示格式被保留，数值筛选使用原值', async () => {
    const book = createBlankWorkbook(), sheet = book.sheets[0];
    sheet.cells = { '0:0': { value: '编号' }, '0:1': { value: '金额' }, '0:2': { value: '日期' }, '1:0': { value: '0009' }, '1:1': { value: 1234.5, styleId: 'money' }, '1:2': { value: 45292, styleId: 'date' } };
    book.styles.money = { numFmt: '#,##0.00' }; book.styles.date = { numFmt: 'yyyy-MM-dd' };
    const saved = await readMergeWorkbook(await serializeXlsx(book, { name: 'data.xlsx' }), 'data.xlsx');
    const data = extractMergeData(saved, saved.sheets[0].id, 1);
    expect(data.records[0].values).toMatchObject({ 编号: '0009', 金额: '1,234.50', 日期: '2024-01-01' });
    expect(filterMergeRecords(data.records, [{ field: '金额', op: 'gt', value: '1000' }])).toHaveLength(1);
    expect(filterMergeRecords(data.records, [{ field: '金额', op: 'lt', value: '1000' }])).toHaveLength(0);
  });

  it('拒绝重复/空表头和合并区域，支持指定非首行表头', () => {
    const book = createBlankWorkbook(), sheet = book.sheets[0];
    sheet.cells = { '1:0': { value: '姓名' }, '1:1': { value: '姓名' }, '2:0': { value: '张三' }, '2:1': { value: '001' } };
    expect(() => extractMergeData(book, sheet.id, 2)).toThrow('重复');
    sheet.cells['1:1'].value = null;
    expect(() => extractMergeData(book, sheet.id, 2)).toThrow('缺少表头');
    sheet.cells['1:1'].value = '编号';
    expect(extractMergeData(book, sheet.id, 2).records[0].rowNumber).toBe(3);
    sheet.merges = [{ r1: 1, r2: 2, c1: 0, c2: 0 }];
    expect(() => extractMergeData(book, sheet.id, 2)).toThrow('合并单元格');
  });

  it('多条件筛选、空值和错误校验不会将数值零当成空值', async () => {
    const book = await readMergeWorkbook(new TextEncoder().encode('编号,姓名,结果\n001,张三,0\n002,李四,10\n003,,20'), 'a.csv');
    const data = extractMergeData(book, book.sheets[0].id, 1);
    expect(filterMergeRecords(data.records, [{ field: '结果', op: 'gte', value: '10' }, { field: '姓名', op: 'notEmpty', value: '' }])[0].values.编号).toBe('002');
    expect(mergeRecordIssues(template(), data, data.records[0])).toEqual([]);
    expect(mergeRecordIssues(template(), data, data.records[2])).toContain('字段为空：姓名');
    expect(mergeRecordIssues(template(), data, data.records[2], true)).toEqual([]);
    const missing = { ...data, fields: ['编号'] };
    expect(mergeRecordIssues(template(), missing, data.records[0], true)).toContain('缺少数据列：姓名');
    data.records[0].errors.push('结果：计算失败');
    expect(mergeRecordIssues(template(), data, data.records[0], true)).toContain('结果：计算失败');
  });

  it('字段 DOM 往返保留字形、不拆分字段，不将占位块颜色写入报告', () => {
    const run = { mergeField: '姓名', text: '«姓名»', style: { bold: true, fontSizePt: 16, color: '#FF0000' } };
    const root = document.createElement('div'); document.body.appendChild(root);
    root.innerHTML = renderRunsHtml([{ text: '前', style: {} }, run, { text: '后', style: {} }]);
    const runs = serializeDomRuns(root, {}, new Map());
    expect(runs[1]).toMatchObject(run);
    expect(runs[1].style.highlight).toBeUndefined();
    expect(runsPlainText(runs)).toBe('前￼后');
    setCaretAtOffset(root, 2);
    expect(getCaretOffset(root)).toBe(2);
    expect(window.getSelection()?.anchorNode?.parentElement?.closest('[data-merge-field]')).toBeNull();
    const [left, right] = splitRunsAtOffset(runs, 2, {});
    expect(left.some((item) => item.mergeField === '姓名')).toBe(true);
    expect(right.some((item) => item.mergeField)).toBe(false);
    root.remove();
  });

  it('简单和跨 run 的复杂 Word 合并域可读取、保存和再打开', async () => {
    const doc = template();
    const bytes = await serializeDocx(doc, { name: 'template.docx' });
    expect(collectMergeFields(await reopen(bytes))).toEqual(['编号', '姓名']);
    const zip = await JSZip.loadAsync(bytes);
    let xml = await zip.file('word/document.xml')!.async('text');
    xml = xml.replace(/<w:fldSimple[^>]*>[\s\S]*?<\/w:fldSimple>/, '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> MERGE</w:instrText></w:r><w:r><w:instrText>FIELD "编号" </w:instrText><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>«编</w:t></w:r><w:r><w:t>号»</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>');
    zip.file('word/document.xml', xml);
    const parsed = await reopen(await zip.generateAsync({ type: 'uint8array' }));
    expect(parsed.blocks[0].type === 'paragraph' && parsed.blocks[0].runs.filter((run) => run.mergeField === '编号')).toHaveLength(1);
    expect(collectMergeFields(await reopen(await serializeDocx(parsed, { name: 'again.docx' })))).toEqual(['编号', '姓名']);
  });

  it('每条记录生成副本，保留样式和原模板，并产生安全独立文件名', async () => {
    const doc = template();
    const book = await readMergeWorkbook(new TextEncoder().encode('编号,姓名\n001,张三\n002,李四'), 'a.csv');
    const records = extractMergeData(book, book.sheets[0].id, 1).records;
    const first = mergeDocument(doc, records[0]), second = mergeDocument(doc, records[1]);
    expect(collectMergeFields(doc)).toEqual(['编号', '姓名']);
    expect(collectMergeFields(first)).toEqual([]);
    expect(first.blocks[0].type === 'paragraph' && first.blocks[0].runs[1]).toMatchObject({ text: '001', style: { bold: true, color: '#FF0000' } });
    expect(second.blocks[0].type === 'paragraph' && second.blocks[0].runs[1].text).toBe('002');
    expect(mergeFileName('../../{姓名}:报告', records[0], 0)).not.toMatch(/[/:]/);
    expect(mergeFileName('报告', records[0], 0)).not.toBe(mergeFileName('报告', records[1], 1));
  });

  it('合并 DOCX 保留每份的独立节、页码重启及节总页数，成品不含合并占位符', async () => {
    const doc = template();
    const book = await readMergeWorkbook(new TextEncoder().encode('编号,姓名\n001,张三\n002,李四'), 'a.csv');
    const records = extractMergeData(book, book.sheets[0].id, 1).records;
    const sections = records.map((record) => { const value = mergeDocument(doc, record); if (value.footer?.blocks[0].type === 'paragraph') value.footer.blocks[0].runs = value.footer.blocks[0].runs.map((run) => run.field === 'NUMPAGES' ? { ...run, field: 'SECTIONPAGES' } : run); return value; });
    const bytes = await serializeDocx(sections[0], { name: 'merged.docx', sections, materialized: true });
    const zip = await JSZip.loadAsync(bytes), xml = await zip.file('word/document.xml')!.async('text');
    expect(xml.match(/<w:sectPr\b/g)).toHaveLength(2);
    expect(xml.match(/w:start="1"/g)).toHaveLength(2);
    expect(xml).toContain('张三'); expect(xml).toContain('李四'); expect(xml).not.toContain('MERGEFIELD');
    expect(await zip.file('word/mini-footer-1.xml')!.async('text')).toContain('SECTIONPAGES');
    expect(await zip.file('word/mini-footer-2.xml')!.async('text')).toContain('SECTIONPAGES');
    expect(await zip.file('word/settings.xml')!.async('text')).toContain('updateFields');
    const reopened = await reopen(bytes);
    expect(documentSections(reopened)).toHaveLength(2);
    const pages = paginateDocument(reopened, new Map());
    expect(pages.map((page) => page.sectionPage)).toEqual([0, 0]);
    expect(pages.map((page) => page.sectionTotal)).toEqual([1, 1]);
    for (const page of pages) {
      const footer = page.section?.footer?.blocks[0];
      expect(footer?.type === 'paragraph' && renderParagraphHtml(footer, { pageNumber: 1, totalPages: 2, sectionPages: page.sectionTotal })).toContain('1');
    }
    const again = await reopen(await serializeDocx(reopened, { sourceBytes: bytes, name: 'again.docx' }));
    expect(documentSections(again)).toHaveLength(2);
    expect(paginateDocument(again, new Map()).map((page) => page.sectionTotal)).toEqual([1, 1]);
  });

  it('常见零填充编号格式被保留，超过 15 位的数值和无缓存公式会明确报错', () => {
    const book = createBlankWorkbook(), sheet = book.sheets[0];
    book.styles.code = { numFmt: '00000' };
    sheet.cells = { '0:0': { value: '编号' }, '1:0': { value: 42, styleId: 'code' }, '2:0': { value: 1234567890123456 }, '3:0': { formula: '=UNKNOWN()', value: null } };
    const data = extractMergeData(book, sheet.id, 1);
    expect(data.records[0].values.编号).toBe('00042');
    expect(data.records[1].errors.join()).toContain('精度');
    expect(data.records[2].errors.join()).toContain('公式');
  });

  it('合并单元格内的换行与制表符写成真正的 Word 元素', async () => {
    const doc = template();
    const record = { rowNumber: 2, values: { 编号: '0001', 姓名: '第一行\n第二行\t备注' }, raw: {}, errors: [] };
    const saved = await serializeDocx(mergeDocument(doc, record), { name: 'multiline.docx', materialized: true });
    const zip = await JSZip.loadAsync(saved);
    const xml = await zip.file('word/document.xml')!.async('text');
    expect(xml).toContain('<w:br/>'); expect(xml).toContain('<w:tab/>');
    const reopened = await reopen(saved);
    expect(reopened.blocks[0].type === 'paragraph' && reopened.blocks[0].runs.some((run) => run.lineBreak)).toBe(true);
  });

  it('分页标记保存重开不会引入额外空段与空白页', async () => {
    const doc = createBlankDoc();
    doc.blocks = [
      { type: 'paragraph', id: 'first', style: {}, runs: [{ text: '第一页', style: {} }] },
      { type: 'pageBreak', id: 'break' },
      { type: 'paragraph', id: 'second', style: {}, runs: [{ text: '第二页', style: {} }] },
    ];
    const bytes = await serializeDocx(doc, { name: 'pages.docx' });
    const reopened = await reopen(bytes);
    expect(reopened.blocks.map((block) => block.type)).toEqual(['paragraph', 'pageBreak', 'paragraph']);
    const heights = new Map(reopened.blocks.filter((block) => block.type !== 'pageBreak').map((block) => [block.id, 920]));
    expect(paginateDocument(reopened, heights)).toHaveLength(2);
  });

  it('替换带图片页眉中的合并字段后，图片关系仍可保存并重新打开', async () => {
    URL.createObjectURL ??= () => 'blob:test'; URL.revokeObjectURL ??= () => {};
    const doc = template();
    doc.header = { edited: true, distancePx: 48, blocks: [{ type: 'paragraph', id: 'header', style: {}, runs: [
      { style: {}, image: { src: '', dataUrl: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', widthPx: 10, heightPx: 10 } },
      { text: '«编号»', mergeField: '编号', style: {} },
    ] }] };
    const original = await serializeDocx(doc, { name: 'template.docx' });
    const parsed = await reopen(original);
    const merged = mergeDocument(parsed, { rowNumber: 2, values: { 编号: '0007', 姓名: '张三' }, raw: {}, errors: [] });
    const saved = await reopen(await serializeDocx(merged, { sourceBytes: original, name: 'report.docx', materialized: true }));
    const header = saved.header?.blocks[0];
    expect(header?.type === 'paragraph' && header.runs.some((run) => run.image)).toBe(true);
    expect(header?.type === 'paragraph' && header.runs.some((run) => run.text === '0007')).toBe(true);
  });
});
