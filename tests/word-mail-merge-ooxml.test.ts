// 功能：验证筛选条件翻译成 Word 查询、写出标准 w:mailMerge 与外部数据源关系；时间：2026-10-29；作者：lq。
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { createBlankDoc } from '@mini-office/document-model';
import { buildWordMailMergeXml, conditionsToWordQuery, fileUriTarget, mergeConditionToSql } from '../packages/editor-word/src/mail-merge-ooxml';
import { serializeDocx } from '../packages/editor-word/src/serialize';

describe('Word / WPS 邮件合并兼容输出', () => {
  it('操作符翻译成 SQL 片段，数值比较遇非数值保持恒假', () => {
    expect(mergeConditionToSql({ field: '姓名', op: 'contains', value: '张' })).toBe("`姓名` LIKE '%张%'");
    expect(mergeConditionToSql({ field: '姓名', op: 'notContains', value: '张' })).toBe("`姓名` NOT LIKE '%张%'");
    expect(mergeConditionToSql({ field: '姓名', op: 'eq', value: "O'Brien" })).toBe("`姓名` = 'O''Brien'");
    expect(mergeConditionToSql({ field: '姓名', op: 'ne', value: '张三' })).toBe("`姓名` <> '张三'");
    expect(mergeConditionToSql({ field: '结果', op: 'gt', value: '10' })).toBe('`结果` > 10');
    expect(mergeConditionToSql({ field: '结果', op: 'gte', value: '10' })).toBe('`结果` >= 10');
    expect(mergeConditionToSql({ field: '结果', op: 'lt', value: '10' })).toBe('`结果` < 10');
    expect(mergeConditionToSql({ field: '结果', op: 'lte', value: '10' })).toBe('`结果` <= 10');
    expect(mergeConditionToSql({ field: '结果', op: 'gt', value: 'abc' })).toBe('(1 = 0)');
    expect(mergeConditionToSql({ field: '姓名', op: 'empty', value: '' })).toBe('`姓名` IS NULL');
    expect(mergeConditionToSql({ field: '姓名', op: 'notEmpty', value: '' })).toBe('`姓名` IS NOT NULL');
  });

  it('标识符反引号与 LIKE 通配符被转义', () => {
    expect(mergeConditionToSql({ field: 'a`b', op: 'eq', value: 'x' })).toBe("`a``b` = 'x'");
    expect(mergeConditionToSql({ field: '备注', op: 'contains', value: '50%_x' })).toBe("`备注` LIKE '%50[%][_]x%'");
  });

  it('且/或按从左到右求值，与面板过滤结果一致', () => {
    const conditions = [
      { field: '姓名', op: 'eq' as const, value: '张三', conjunction: 'and' as const },
      { field: '结果', op: 'eq' as const, value: '10', conjunction: 'or' as const },
      { field: '编号', op: 'eq' as const, value: '001', conjunction: 'and' as const },
    ];
    expect(conditionsToWordQuery(conditions, 'Sheet1')).toBe("SELECT * FROM `Sheet1$` WHERE (((`姓名` = '张三') OR (`结果` = '10')) AND (`编号` = '001'))");
    expect(conditionsToWordQuery([], 'Sheet1')).toBe('SELECT * FROM `Sheet1$`');
  });

  it('磁盘路径转成外部关系使用的 file URI（Windows 路径规范化为正斜杠）', () => {
    expect(fileUriTarget('C:\\data\\原料.xlsx')).toBe('file:///C:/data/原料.xlsx');
    expect(fileUriTarget('Z:\\康美\\原料检验流水.xlsx')).toBe('file:///Z:/康美/原料检验流水.xlsx');
    expect(fileUriTarget('/home/me/data.xlsx')).toBe('file:///home/me/data.xlsx');
  });

  it('w:mailMerge 片段含数据源关系、查询与字段映射', () => {
    const xml = buildWordMailMergeXml({ dataSourcePath: 'C:\\data\\原料.xlsx', sheetName: 'Sheet1', fields: ['编号', '姓名'], conditions: [{ field: '姓名', op: 'eq', value: '张三' }] });
    expect(xml).toContain('<w:mailMerge>');
    expect(xml).toContain('<w:mainDocumentType w:val="formLetters"/>');
    expect(xml).toContain('<w:dataSource r:id="rIdMailMergeSource"/>');
    expect(xml).toContain('<w:src r:id="rIdMailMergeOdso"/>');
    expect(xml).toContain('SELECT * FROM `Sheet1$`');
    expect(xml).toContain('<w:fieldMapData>');
  });

  it('保存模板写出 w:mailMerge 与外部关系，成品输出不保留', async () => {
    const doc = createBlankDoc();
    const info = { dataSourcePath: 'C:\\data\\原料.xlsx', sheetName: 'Sheet1', fields: ['编号'], conditions: [] };
    const bytes = await serializeDocx(doc, { name: 'template.docx', mailMerge: info });
    const zip = await JSZip.loadAsync(bytes);
    const settings = await zip.file('word/settings.xml')?.async('text') ?? '';
    const rels = await zip.file('word/_rels/settings.xml.rels')?.async('text') ?? '';
    expect(settings).toContain('<w:mailMerge>');
    expect(settings).toContain('r:id="rIdMailMergeSource"');
    // w:mailMerge 必须排在 defaultTabStop 之前（CT_Settings 的 schema 顺序）
    expect(settings.indexOf('<w:mailMerge>')).toBeGreaterThanOrEqual(0);
    expect(settings.indexOf('<w:mailMerge>')).toBeLessThan(settings.indexOf('<w:defaultTabStop'));
    expect(rels).toContain('TargetMode="External"');
    expect(rels).toContain('file:///C:/data/原料.xlsx');
    const output = await serializeDocx(doc, { name: 'output.docx', materialized: true, mailMerge: info });
    const outputZip = await JSZip.loadAsync(output);
    const outputSettings = await outputZip.file('word/settings.xml')?.async('text') ?? '';
    expect(outputSettings).not.toContain('<w:mailMerge>');
  });

  it('覆盖 Word 原有的邮件合并设置并清除缓存的收件人名单', async () => {
    const base = await serializeDocx(createBlankDoc(), { name: 'base.docx' });
    const zip = await JSZip.loadAsync(base);
    const settings = (await zip.file('word/settings.xml')?.async('text') ?? '')
      .replace('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"', 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"')
      .replace(
        '<w:defaultTabStop',
        '<w:mailMerge><w:mainDocumentType w:val="formLetters"/><w:dataSource r:id="rId1"/><w:odso><w:table w:val="旧表$"/><w:src r:id="rId2"/><w:recipientData r:id="rId3"/></w:odso></w:mailMerge><w:defaultTabStop');
    zip.file('word/settings.xml', settings);
    zip.file('word/_rels/settings.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/mailMergeSource" Target="file:///Z:\\old\\old.xlsx" TargetMode="External"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/recipientData" Target="recipientData.xml"/></Relationships>');
    zip.file('word/recipientData.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><root/>');
    const patched = await zip.generateAsync({ type: 'uint8array' });
    const output = await serializeDocx(createBlankDoc(), {
      name: 'out.docx',
      sourceBytes: patched,
      mailMerge: { dataSourcePath: 'Z:\\new\\new.xlsx', sheetName: '新表', fields: ['编号'], conditions: [] },
    });
    const outZip = await JSZip.loadAsync(output);
    const outSettings = await outZip.file('word/settings.xml')?.async('text') ?? '';
    const outRels = await outZip.file('word/_rels/settings.xml.rels')?.async('text') ?? '';
    expect(outSettings).toContain('r:id="rIdMailMergeSource"');
    expect(outSettings).toContain('`新表$`');
    expect(outSettings).not.toContain('w:recipientData');
    expect(outSettings).not.toContain('旧表');
    expect(outRels).not.toContain('recipientData');
    expect(outRels).not.toContain('old.xlsx');
    expect(outZip.file('word/recipientData.xml')).toBeNull();
  });
});
