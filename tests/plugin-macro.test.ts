import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import {
  MacroError,
  macroKindOf,
  parsePackage,
  runMacro,
  validateMacroSteps,
  validateManifest,
} from '@mini-office/plugin-host';
import { createBlankDoc, type DocModel, type DocTable, type WorkbookModel } from '@mini-office/document-model';

const BASE_MANIFEST = {
  id: 'mini-office.pack-gov-doc',
  name: '公文排版工具',
  version: '1.0.0',
  publisher: 'mini-office',
  tier: 'declarative',
  engines: { miniOffice: '>=0.1.5 <0.5.0', api: '^1.0.0' },
  permissions: [],
};

function withCommands(steps: unknown[]) {
  return {
    ...BASE_MANIFEST,
    contributes: { commands: [{ id: 'apply', title: '套用公文版式', steps }] },
  };
}

function makeTable(): DocTable {
  const cell = (id: string) => ({ id, blocks: [], colSpan: 1, rowSpan: 1 });
  return {
    type: 'table',
    id: 't-1',
    columnWidthsPx: [100, 100],
    rows: [
      { id: 'r-1', cells: [cell('c-1'), cell('c-2')] },
      { id: 'r-2', cells: [cell('c-3'), cell('c-4')] },
      { id: 'r-3', cells: [cell('c-5'), cell('c-6')] },
    ],
  };
}

function docWithTable(): DocModel {
  const doc = createBlankDoc();
  return { ...doc, blocks: [...doc.blocks, makeTable()] };
}

function makeWorkbook(): WorkbookModel {
  return {
    version: 1,
    sheets: [
      {
        id: 'sheet-1',
        name: 'Sheet1',
        state: 'visible',
        order: 0,
        cells: {
          '0:0': { value: '姓名' },
          '0:1': { value: '金额' },
          '1:0': { value: '张三' },
          '1:1': { value: 12 },
        },
        colWidths: {},
        rowHeights: {},
        merges: [],
        hiddenRows: {},
        hiddenCols: {},
        dimension: { rows: 100, cols: 26 },
        defaultColWidth: 88,
        defaultRowHeight: 24,
        showGridLines: true,
      },
    ],
    activeSheetId: 'sheet-1',
    styles: { default: {} },
    themeColors: {},
    themeFonts: { major: '等线 Light', minor: '等线' },
    defaultFont: '等线',
    defaultFontSizePt: 11,
    meta: {},
  };
}

describe('宏步骤校验', () => {
  it('接受合法的 word 宏', () => {
    expect(
      validateMacroSteps([
        { op: 'word.page.margins', top: 120, bottom: 96 },
        { op: 'word.font.apply', fontFamily: '仿宋_GB2312', fontSizePt: 16 },
      ]),
    ).toEqual([]);
  });

  it('拒绝未知操作', () => {
    const issues = validateMacroSteps([{ op: 'word.magic' }]);
    expect(issues[0].message).toContain('未知操作');
  });

  it('拒绝空步骤与超长步骤', () => {
    expect(validateMacroSteps([])[0].message).toContain('非空数组');
    const many = Array.from({ length: 33 }, () => ({ op: 'word.page.margins', top: 10 }));
    expect(validateMacroSteps(many)[0].message).toContain('32');
  });

  it('拒绝越界数值与非法枚举', () => {
    expect(validateMacroSteps([{ op: 'word.page.margins', top: 99999 }])[0].message).toContain('之间');
    expect(validateMacroSteps([{ op: 'word.page.orientation', orientation: 'sideways' }])[0].message).toContain(
      'portrait',
    );
  });

  it('拒绝非法颜色', () => {
    expect(validateMacroSteps([{ op: 'word.font.apply', color: 'red' }])[0].message).toContain('#RRGGBB');
  });

  it('拒绝没有任何参数的步骤', () => {
    expect(validateMacroSteps([{ op: 'word.paragraph.format', scope: 'all' }])[0].message).toContain('不会有变化');
  });

  it('拒绝混用文档类型的步骤', () => {
    const issues = validateMacroSteps([
      { op: 'word.page.margins', top: 10 },
      { op: 'sheet.freezeHeader' },
    ]);
    expect(issues[0].message).toContain('同一种文档类型');
  });

  it('能从步骤推断文档类型', () => {
    expect(macroKindOf([{ op: 'word.font.apply', bold: true }])).toBe('docx');
    expect(macroKindOf([{ op: 'sheet.freezeHeader' }])).toBe('xlsx');
    expect(
      macroKindOf([
        { op: 'word.font.apply', bold: true },
        { op: 'sheet.freezeHeader' },
      ]),
    ).toBeNull();
  });
});

describe('宏执行：Word', () => {
  it('修改页边距且不改动其他页设置', () => {
    const doc = createBlankDoc();
    const next = runMacro('docx', doc, [{ op: 'word.page.margins', top: 120, left: 140 }]);
    expect(next.page.margins.top).toBe(120);
    expect(next.page.margins.left).toBe(140);
    expect(next.page.margins.right).toBe(doc.page.margins.right);
    expect(doc.page.margins.top).toBe(doc.page.margins.top);
  });

  it('切换纸张方向时同步交换宽高', () => {
    const doc = createBlankDoc();
    const next = runMacro('docx', doc, [{ op: 'word.page.orientation', orientation: 'landscape' }]);
    expect(next.page.orientation).toBe('landscape');
    expect(next.page.widthPx).toBe(doc.page.heightPx);
    expect(next.page.heightPx).toBe(doc.page.widthPx);
  });

  it('把段落格式与字体应用到正文段落', () => {
    const doc = createBlankDoc();
    const next = runMacro('docx', doc, [
      { op: 'word.paragraph.format', lineHeight: 1.5, firstLineIndentPx: 32, alignment: 'justify' },
      { op: 'word.font.apply', fontFamily: '仿宋_GB2312', fontSizePt: 16, color: '#000000' },
    ]);
    const paragraph = next.blocks[0];
    expect(paragraph.type).toBe('paragraph');
    if (paragraph.type !== 'paragraph') return;
    expect(paragraph.style.lineHeight).toBe(1.5);
    expect(paragraph.style.firstLineIndentPx).toBe(32);
    expect(paragraph.style.alignment).toBe('justify');
    expect(paragraph.runs[0].style.fontFamily).toBe('仿宋_GB2312');
    expect(paragraph.runs[0].style.fontSizePt).toBe(16);
  });

  it('默认跳过居中标题与右对齐落款，scope=all 时才会改动它们', () => {
    const doc = createBlankDoc();
    const paragraph = (id: string, alignment: 'center' | 'right' | 'left') => ({
      type: 'paragraph' as const,
      id,
      runs: [{ text: id, style: {} }],
      style: { alignment },
    });
    const withBlocks: DocModel = {
      ...doc,
      blocks: [paragraph('title', 'center'), paragraph('body', 'left'), paragraph('sign', 'right')],
    };

    const scoped = runMacro('docx', withBlocks, [
      { op: 'word.paragraph.format', lineHeight: 1.5, firstLineIndentPx: 32, alignment: 'justify' },
      { op: 'word.font.apply', fontFamily: '仿宋_GB2312', fontSizePt: 16 },
    ]);
    const pick = (target: DocModel, id: string) => {
      const block = target.blocks.find((item) => item.id === id);
      return block && block.type === 'paragraph' ? block : undefined;
    };
    expect(pick(scoped, 'body')?.style.firstLineIndentPx).toBe(32);
    expect(pick(scoped, 'body')?.style.lineHeight).toBe(1.5);
    expect(pick(scoped, 'body')?.runs[0].style.fontFamily).toBe('仿宋_GB2312');
    // 标题保持居中、落款保持右对齐，且没有被改字号
    expect(pick(scoped, 'title')?.style.alignment).toBe('center');
    expect(pick(scoped, 'title')?.runs[0].style.fontSizePt).toBeUndefined();
    expect(pick(scoped, 'sign')?.style.alignment).toBe('right');
    expect(pick(scoped, 'sign')?.style.firstLineIndentPx).toBeUndefined();

    const everything = runMacro('docx', withBlocks, [
      { op: 'word.paragraph.format', scope: 'all', alignment: 'justify', firstLineIndentPx: 32 },
    ]);
    expect(pick(everything, 'title')?.style.alignment).toBe('justify');
    expect(pick(everything, 'title')?.style.firstLineIndentPx).toBe(32);
  });

  it('三线表：首行上下有线、末行下方粗线、中间无线', () => {
    const doc = docWithTable();
    const next = runMacro('docx', doc, [{ op: 'word.table.borders', preset: 'three-line' }]);
    const table = next.blocks.find((block): block is DocTable => block.type === 'table');
    expect(table).toBeDefined();
    if (!table) return;
    expect(table.borders).toBeUndefined();
    const firstRowCell = table.rows[0].cells[0];
    const middleRowCell = table.rows[1].cells[0];
    const lastRowCell = table.rows[2].cells[0];
    expect(firstRowCell.borders?.top?.style).toBe('single');
    expect(firstRowCell.borders?.bottom?.style).toBe('single');
    expect(firstRowCell.borders?.left?.style).toBe('none');
    expect(middleRowCell.borders?.top?.style).toBe('none');
    expect(middleRowCell.borders?.bottom?.style).toBe('none');
    expect(lastRowCell.borders?.bottom?.widthPx).toBe(1.5);
  });

  it('设置默认字体', () => {
    const next = runMacro('docx', createBlankDoc(), [
      { op: 'word.defaultFont.set', fontFamily: '方正小标宋', fontSizePt: 22 },
    ]);
    expect(next.defaultFont).toBe('方正小标宋');
    expect(next.defaultFontSizePt).toBe(22);
  });
});

describe('宏执行：Excel', () => {
  it('冻结首行', () => {
    const next = runMacro('xlsx', makeWorkbook(), [{ op: 'sheet.freezeHeader' }]);
    expect(next.sheets[0].freeze).toEqual({ rows: 1, cols: 0 });
  });

  it('给表头行套用样式，空单元格也会被填充', () => {
    const workbook = makeWorkbook();
    workbook.sheets[0].cells['0:2'] = { value: null };
    const next = runMacro('xlsx', workbook, [{ op: 'sheet.headerStyle', bold: true, fill: '#F2F2F2' }]);
    const sheet = next.sheets[0];
    const styleId = sheet.cells['0:0'].styleId;
    expect(styleId).toBeDefined();
    if (!styleId) return;
    expect(next.styles[styleId].bold).toBe(true);
    expect(next.styles[styleId].fill).toBe('#F2F2F2');
    // 第 2 行不受影响
    expect(sheet.cells['1:0'].styleId).toBeUndefined();
    // 同一份样式被复用而不是每次新建
    expect(sheet.cells['0:1'].styleId).toBe(styleId);
  });
});

describe('宏执行的安全边界', () => {
  it('文档类型不匹配时拒绝执行', () => {
    expect(() => runMacro('xlsx', makeWorkbook(), [{ op: 'word.page.margins', top: 10 }])).toThrow(MacroError);
  });

  it('校验失败时抛错且不改动原文档', () => {
    const doc = createBlankDoc();
    const snapshot = JSON.stringify(doc);
    expect(() => runMacro('docx', doc, [{ op: 'word.page.margins', top: -5 }])).toThrow(MacroError);
    expect(JSON.stringify(doc)).toBe(snapshot);
  });

  it('不修改入参，返回新对象', () => {
    const doc = createBlankDoc();
    const next = runMacro('docx', doc, [{ op: 'word.page.margins', top: 120 }]);
    expect(next).not.toBe(doc);
    expect(doc.page.margins.top).not.toBe(120);
  });
});

describe('清单校验：命令与模板', () => {
  it('接受只有命令的插件', () => {
    const result = validateManifest(withCommands([{ op: 'word.page.margins', top: 100 }]));
    expect(result.ok).toBe(true);
    expect(result.manifest?.contributes.commands?.[0].title).toBe('套用公文版式');
  });

  it('拒绝未实现的贡献点而不是静默忽略', () => {
    const result = validateManifest({ ...BASE_MANIFEST, contributes: { panels: [{ id: 'x' }] } });
    expect(result.ok).toBe(false);
    expect(result.issues.some((issue) => issue.path === 'contributes.panels')).toBe(true);
  });

  it('拒绝混用文档类型的命令', () => {
    const result = validateManifest(
      withCommands([
        { op: 'word.page.margins', top: 10 },
        { op: 'sheet.freezeHeader' },
      ]),
    );
    expect(result.ok).toBe(false);
  });

  it('拒绝重复命令 id', () => {
    const result = validateManifest({
      ...BASE_MANIFEST,
      contributes: {
        commands: [
          { id: 'a', title: 'A', steps: [{ op: 'word.page.margins', top: 10 }] },
          { id: 'a', title: 'B', steps: [{ op: 'word.page.margins', top: 20 }] },
        ],
      },
    });
    expect(result.ok).toBe(false);
    expect(result.issues.some((issue) => issue.message.includes('重复'))).toBe(true);
  });

  it('模板扩展名必须与 kind 一致', () => {
    const result = validateManifest({
      ...BASE_MANIFEST,
      contributes: { templates: [{ id: 'letter', label: '公文模板', kind: 'docx', file: 'templates/a.xlsx' }] },
    });
    expect(result.ok).toBe(false);
    expect(result.issues.some((issue) => issue.path.endsWith('.file'))).toBe(true);
  });

  it('模板路径不能穿越', () => {
    const result = validateManifest({
      ...BASE_MANIFEST,
      contributes: { templates: [{ id: 'letter', label: '公文模板', kind: 'docx', file: '../a.docx' }] },
    });
    expect(result.ok).toBe(false);
  });
});

describe('插件包：模板文件必须存在', () => {
  async function pack(manifest: unknown, extra: Record<string, string> = {}): Promise<Uint8Array> {
    const zip = new JSZip();
    zip.file('manifest.json', JSON.stringify(manifest));
    for (const [path, content] of Object.entries(extra)) zip.file(path, content);
    return zip.generateAsync({ type: 'uint8array' });
  }

  it('缺少模板文件时安装失败', async () => {
    const bytes = await pack({
      ...BASE_MANIFEST,
      contributes: { templates: [{ id: 'letter', label: '公文模板', kind: 'docx', file: 'templates/a.docx' }] },
    });
    await expect(parsePackage(bytes)).rejects.toThrow(/模板文件缺失/);
  });

  it('模板文件存在时安装通过', async () => {
    const bytes = await pack(
      {
        ...BASE_MANIFEST,
        contributes: { templates: [{ id: 'letter', label: '公文模板', kind: 'docx', file: 'templates/a.docx' }] },
      },
      { 'templates/a.docx': 'not-a-real-docx-but-present' },
    );
    const parsed = await parsePackage(bytes);
    expect(parsed.manifest.contributes.templates).toHaveLength(1);
    expect(parsed.files.has('templates/a.docx')).toBe(true);
  });
});
