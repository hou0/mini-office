/**
 * 插件宏：校验与执行。
 *
 * 宏是"用数据描述的一串文档操作"。声明式插件因此仍然**不执行任何代码**——
 * 宿主拿清单里的 `steps` 解释执行，能做与不能做完全由这里的白名单决定。
 *
 * 两条硬规则：
 * 1. 先整体校验、再整体执行。任何一步不合法，整条命令都不执行，
 *    不会出现"改了一半"的文档。
 * 2. 一条命令的所有步骤必须属于同一种文档类型，避免在 xlsx 上跑 word 操作。
 */
import {
  MACRO_OPS,
  type MacroOp,
  type MacroScope,
  type MacroStep,
  type PluginDocumentKind,
  type TableBorderPreset,
  type ValidationIssue,
} from '@mini-office/plugin-api';
import type {
  DocBlock,
  DocModel,
  DocParagraph,
  DocTable,
  DocTableCell,
  SheetModel,
  WorkbookModel,
} from '@mini-office/document-model';

export class MacroError extends Error {
  readonly issues: ValidationIssue[];
  constructor(message: string, issues: ValidationIssue[] = []) {
    super(message);
    this.name = 'MacroError';
    this.issues = issues;
  }
}

const HEX_COLOR = /^#[0-9a-fA-F]{6}$/;
const ORIENTATIONS = ['portrait', 'landscape'] as const;
const ALIGNMENTS = ['left', 'center', 'right', 'justify'] as const;
const SCOPES = ['all', 'body'] as const;
const SHEET_TARGETS = ['active', 'all'] as const;
const BORDER_PRESETS: TableBorderPreset[] = ['three-line', 'grid', 'none'];

/** 单个操作的参数规格。用声明式表格描述，避免为每个操作手写一遍校验。 */
interface OpSpec {
  kind: PluginDocumentKind;
  /** 数值字段：[最小值, 最大值] */
  numbers?: Record<string, [number, number]>;
  /** 枚举字段 */
  enums?: Record<string, readonly string[]>;
  /** 字符串字段：[最大长度] */
  strings?: Record<string, number>;
  /** 颜色字段，必须为 `#RRGGBB` */
  colors?: string[];
  booleans?: string[];
  required?: string[];
  /** 至少需要提供其中一个可选字段，否则步骤没有意义 */
  atLeastOne?: boolean;
}

const OP_SPECS: Record<MacroOp, OpSpec> = {
  'word.page.margins': {
    kind: 'docx',
    numbers: { top: [0, 1440], right: [0, 1440], bottom: [0, 1440], left: [0, 1440] },
    atLeastOne: true,
  },
  'word.page.orientation': {
    kind: 'docx',
    enums: { orientation: ORIENTATIONS },
    required: ['orientation'],
  },
  'word.paragraph.format': {
    kind: 'docx',
    enums: { scope: SCOPES, alignment: ALIGNMENTS },
    numbers: {
      lineHeight: [0.5, 10],
      firstLineIndentPx: [-960, 960],
      spaceBeforePx: [0, 960],
      spaceAfterPx: [0, 960],
      indentLeftPx: [0, 1440],
      indentRightPx: [0, 1440],
    },
    atLeastOne: true,
  },
  'word.font.apply': {
    kind: 'docx',
    enums: { scope: SCOPES },
    numbers: { fontSizePt: [1, 400] },
    strings: { fontFamily: 128 },
    colors: ['color'],
    booleans: ['bold', 'italic'],
    atLeastOne: true,
  },
  'word.table.borders': {
    kind: 'docx',
    enums: { preset: BORDER_PRESETS },
    required: ['preset'],
  },
  'word.defaultFont.set': {
    kind: 'docx',
    numbers: { fontSizePt: [1, 400] },
    strings: { fontFamily: 128 },
    required: ['fontFamily'],
  },
  'sheet.freezeHeader': {
    kind: 'xlsx',
    enums: { sheet: SHEET_TARGETS },
  },
  'sheet.headerStyle': {
    kind: 'xlsx',
    enums: { sheet: SHEET_TARGETS },
    numbers: { fontSizePt: [1, 400] },
    strings: { fontFamily: 128 },
    colors: ['fill', 'color'],
    booleans: ['bold'],
    atLeastOne: true,
  },
};

/** 由步骤前缀推断文档类型；混合类型返回 null。 */
export function macroKindOf(steps: readonly MacroStep[]): PluginDocumentKind | null {
  let kind: PluginDocumentKind | null = null;
  for (const step of steps) {
    const spec = isMacroOp(step?.op) ? OP_SPECS[step.op] : undefined;
    if (!spec) return null;
    if (!kind) kind = spec.kind;
    else if (kind !== spec.kind) return null;
  }
  return kind;
}

function isMacroOp(value: unknown): value is MacroOp {
  return typeof value === 'string' && (MACRO_OPS as readonly string[]).includes(value);
}

/** 校验单个步骤；返回空数组表示通过。 */
function validateStep(step: unknown, path: string): ValidationIssue[] {
  const issues: ValidationIssue[] = [];
  if (!step || typeof step !== 'object' || Array.isArray(step)) {
    return [{ path, message: '步骤必须是对象' }];
  }
  const record = step as Record<string, unknown>;
  if (!isMacroOp(record.op)) {
    issues.push({ path: `${path}.op`, message: `未知操作 ${String(record.op)}，可用操作见文档` });
    return issues;
  }
  const spec = OP_SPECS[record.op];

  for (const key of spec.required ?? []) {
    if (record[key] === undefined) issues.push({ path: `${path}.${key}`, message: '缺少必填字段' });
  }

  for (const [key, [min, max]] of Object.entries(spec.numbers ?? {})) {
    const value = record[key];
    if (value === undefined) continue;
    if (typeof value !== 'number' || !Number.isFinite(value)) {
      issues.push({ path: `${path}.${key}`, message: '必须是数字' });
    } else if (value < min || value > max) {
      issues.push({ path: `${path}.${key}`, message: `需在 ${min} 到 ${max} 之间` });
    }
  }

  for (const [key, allowed] of Object.entries(spec.enums ?? {})) {
    const value = record[key];
    if (value === undefined) continue;
    if (typeof value !== 'string' || !allowed.includes(value)) {
      issues.push({ path: `${path}.${key}`, message: `取值必须是 ${allowed.join(' / ')}` });
    }
  }

  for (const [key, maxLength] of Object.entries(spec.strings ?? {})) {
    const value = record[key];
    if (value === undefined) continue;
    if (typeof value !== 'string' || !value.trim()) {
      issues.push({ path: `${path}.${key}`, message: '必须是非空字符串' });
    } else if (value.length > maxLength) {
      issues.push({ path: `${path}.${key}`, message: `长度不能超过 ${maxLength} 个字符` });
    }
  }

  for (const key of spec.colors ?? []) {
    const value = record[key];
    if (value === undefined) continue;
    if (typeof value !== 'string' || !HEX_COLOR.test(value)) {
      issues.push({ path: `${path}.${key}`, message: '颜色必须是 #RRGGBB 形式' });
    }
  }

  for (const key of spec.booleans ?? []) {
    const value = record[key];
    if (value === undefined) continue;
    if (typeof value !== 'boolean') issues.push({ path: `${path}.${key}`, message: '必须是布尔值' });
  }

  if (spec.atLeastOne) {
    const hasAny = Object.keys(record).some((key) => key !== 'op' && key !== 'scope' && key !== 'sheet');
    if (!hasAny) issues.push({ path, message: '该步骤没有给出任何参数，执行后不会有变化' });
  }

  return issues;
}

/** 校验整条命令的步骤列表。 */
export function validateMacroSteps(steps: unknown, path = 'steps'): ValidationIssue[] {
  if (!Array.isArray(steps) || steps.length === 0) {
    return [{ path, message: 'steps 必须是非空数组' }];
  }
  if (steps.length > 32) {
    return [{ path, message: '单条命令最多 32 个步骤' }];
  }
  const issues: ValidationIssue[] = [];
  steps.forEach((step, index) => {
    issues.push(...validateStep(step, `${path}[${index}]`));
  });
  if (issues.length > 0) return issues;

  const kinds = new Set(
    steps.map((step) => OP_SPECS[(step as MacroStep).op].kind),
  );
  if (kinds.size > 1) {
    issues.push({
      path,
      message: `同一条命令的步骤必须属于同一种文档类型，当前混用了 ${[...kinds].join(' 与 ')}`,
    });
  }
  return issues;
}

/* -------------------------------------------------------------------------- */
/* 执行                                                                        */
/* -------------------------------------------------------------------------- */

function borderSide(style: 'none' | 'single', widthPx: number, color = '#000000') {
  return { style, widthPx, color } as const;
}

function tableBorderPreset(preset: TableBorderPreset, rowIndex: number, isLastRow: boolean) {
  if (preset === 'none') {
    return {
      top: borderSide('none', 0),
      bottom: borderSide('none', 0),
      left: borderSide('none', 0),
      right: borderSide('none', 0),
    };
  }
  if (preset === 'grid') {
    return {
      top: borderSide('single', 1),
      bottom: borderSide('single', 1),
      left: borderSide('single', 1),
      right: borderSide('single', 1),
    };
  }
  // 三线表：首行上下为框线，中间无线，末行下方为粗线。
  // 无线时必须用 `none` 而不是宽度 0——渲染层会把 0 抬到 0.75px，反而画出线来。
  const top = rowIndex === 0 ? borderSide('single', 1.5) : borderSide('none', 0);
  const bottom =
    rowIndex === 0 ? borderSide('single', 0.75) : isLastRow ? borderSide('single', 1.5) : borderSide('none', 0);
  return {
    top,
    bottom,
    left: borderSide('none', 0),
    right: borderSide('none', 0),
  };
}

function applyTablePreset(table: DocTable, preset: TableBorderPreset): DocTable {
  const lastRow = table.rows.length - 1;
  const rows = table.rows.map((row, rowIndex) => ({
    ...row,
    cells: row.cells.map((cell) => {
      const cellLastRow = Math.min(lastRow, rowIndex + cell.rowSpan - 1);
      const sides = tableBorderPreset(preset, rowIndex, cellLastRow >= lastRow);
      const next: DocTableCell = {
        ...cell,
        borders: {
          top: { ...sides.top },
          bottom: { ...sides.bottom },
          left: { ...sides.left },
          right: { ...sides.right },
        },
      };
      return next;
    }),
  }));
  // 表格级边框会让单元格继承出多余线条，这里清空，完全由单元格自己决定。
  return { ...table, rows, borders: undefined };
}

/** 只遍历正文顶层段落，不进入表格单元格——公文排版通常不希望改动表内文字。 */
function mapBodyParagraphs(doc: DocModel, fn: (paragraph: DocParagraph) => DocParagraph): DocModel {
  let changed = false;
  const blocks: DocBlock[] = doc.blocks.map((block) => {
    if (block.type !== 'paragraph') return block;
    const next = fn(block);
    if (next !== block) changed = true;
    return next;
  });
  return changed ? { ...doc, blocks } : doc;
}

/**
 * `scope: 'body'` 时跳过居中/右对齐段落。
 *
 * 标题通常是居中、落款是右对齐，它们需要保持原样；如果一视同仁地套用正文格式，
 * 标题会被拉成两端对齐、落款会多出首行缩进而换行——这正是宏最容易出错的地方。
 */
function isBodyParagraph(paragraph: DocParagraph): boolean {
  const alignment = paragraph.style.alignment;
  return alignment !== 'center' && alignment !== 'right';
}

function mapScopedParagraphs(
  doc: DocModel,
  scope: MacroScope | undefined,
  fn: (paragraph: DocParagraph) => DocParagraph,
): DocModel {
  const bodyOnly = (scope ?? 'body') === 'body';
  return mapBodyParagraphs(doc, (paragraph) => (bodyOnly && !isBodyParagraph(paragraph) ? paragraph : fn(paragraph)));
}

function applyWordStep(doc: DocModel, step: MacroStep): DocModel {
  switch (step.op) {
    case 'word.page.margins': {
      const margins = { ...doc.page.margins };
      if (step.top !== undefined) margins.top = step.top;
      if (step.right !== undefined) margins.right = step.right;
      if (step.bottom !== undefined) margins.bottom = step.bottom;
      if (step.left !== undefined) margins.left = step.left;
      return { ...doc, page: { ...doc.page, margins } };
    }
    case 'word.page.orientation': {
      if (doc.page.orientation === step.orientation) return doc;
      return {
        ...doc,
        page: {
          ...doc.page,
          orientation: step.orientation,
          widthPx: doc.page.heightPx,
          heightPx: doc.page.widthPx,
        },
      };
    }
    case 'word.paragraph.format': {
      const patch = {
        ...(step.alignment !== undefined ? { alignment: step.alignment } : {}),
        ...(step.lineHeight !== undefined ? { lineHeight: step.lineHeight } : {}),
        ...(step.firstLineIndentPx !== undefined ? { firstLineIndentPx: step.firstLineIndentPx } : {}),
        ...(step.spaceBeforePx !== undefined ? { spaceBeforePx: step.spaceBeforePx } : {}),
        ...(step.spaceAfterPx !== undefined ? { spaceAfterPx: step.spaceAfterPx } : {}),
        ...(step.indentLeftPx !== undefined ? { indentLeftPx: step.indentLeftPx } : {}),
        ...(step.indentRightPx !== undefined ? { indentRightPx: step.indentRightPx } : {}),
      };
      return mapScopedParagraphs(doc, step.scope, (paragraph) => ({
        ...paragraph,
        style: { ...paragraph.style, ...patch },
      }));
    }
    case 'word.font.apply': {
      const patch = {
        ...(step.fontFamily !== undefined ? { fontFamily: step.fontFamily } : {}),
        ...(step.fontSizePt !== undefined ? { fontSizePt: step.fontSizePt } : {}),
        ...(step.bold !== undefined ? { bold: step.bold } : {}),
        ...(step.italic !== undefined ? { italic: step.italic } : {}),
        ...(step.color !== undefined ? { color: step.color } : {}),
      };
      return mapScopedParagraphs(doc, step.scope, (paragraph) => ({
        ...paragraph,
        runs: paragraph.runs.map((run) => ({ ...run, style: { ...run.style, ...patch } })),
      }));
    }
    case 'word.table.borders': {
      let changed = false;
      const blocks = doc.blocks.map((block) => {
        if (block.type !== 'table') return block;
        changed = true;
        return applyTablePreset(block, step.preset);
      });
      return changed ? { ...doc, blocks } : doc;
    }
    case 'word.defaultFont.set': {
      return {
        ...doc,
        defaultFont: step.fontFamily,
        ...(step.fontSizePt !== undefined ? { defaultFontSizePt: step.fontSizePt } : {}),
      };
    }
    default:
      throw new MacroError(`宏步骤与文档类型不匹配：${step.op}`);
  }
}

function cellKey(row: number, col: number): string {
  return `${row}:${col}`;
}

/** 找到工作表中已使用的最右列；上限 256，避免在空表上铺满整行。 */
function usedMaxCol(sheet: SheetModel): number {
  let maxCol = 0;
  for (const key of Object.keys(sheet.cells)) {
    const separator = key.indexOf(':');
    if (separator < 0) continue;
    const col = Number(key.slice(separator + 1));
    if (Number.isFinite(col) && col > maxCol) maxCol = col;
  }
  return Math.min(maxCol, 255);
}

function styleKey(style: Record<string, unknown>): string {
  const { xfIndex: _ignored, ...rest } = style;
  return JSON.stringify(rest);
}

function findOrCreateStyle(
  styles: WorkbookModel['styles'],
  next: Record<string, unknown>,
): { styles: WorkbookModel['styles']; styleId: string } {
  const key = styleKey(next);
  for (const [id, existing] of Object.entries(styles)) {
    if (styleKey(existing as Record<string, unknown>) === key) return { styles, styleId: id };
  }
  let counter = Object.keys(styles).length + 1;
  let styleId = `style-macro-${counter}`;
  while (styles[styleId]) {
    counter += 1;
    styleId = `style-macro-${counter}`;
  }
  return { styles: { ...styles, [styleId]: next as never }, styleId };
}

function applyHeaderStyle(workbook: WorkbookModel, sheet: SheetModel, patch: Record<string, unknown>): WorkbookModel {
  const maxCol = usedMaxCol(sheet);
  const cells = { ...sheet.cells };
  let styles = workbook.styles;
  for (let col = 0; col <= maxCol; col += 1) {
    const key = cellKey(0, col);
    const existing = cells[key];
    const base = (existing?.styleId ? styles[existing.styleId] : undefined) ?? styles.default ?? {};
    const merged = { ...base, ...patch };
    const created = findOrCreateStyle(styles, merged);
    styles = created.styles;
    cells[key] = { value: existing?.value ?? null, ...(existing?.formula ? { formula: existing.formula } : {}), styleId: created.styleId };
  }
  const sheets = workbook.sheets.map((item) => (item.id === sheet.id ? { ...item, cells } : item));
  return { ...workbook, sheets, styles };
}

function applySheetStep(workbook: WorkbookModel, step: MacroStep): WorkbookModel {
  const targets = step.op === 'sheet.freezeHeader' || step.op === 'sheet.headerStyle' ? step.sheet ?? 'active' : 'active';
  const selected = targets === 'all' ? workbook.sheets : workbook.sheets.filter((sheet) => sheet.id === workbook.activeSheetId);

  switch (step.op) {
    case 'sheet.freezeHeader': {
      if (selected.length === 0) return workbook;
      const ids = new Set(selected.map((sheet) => sheet.id));
      return {
        ...workbook,
        sheets: workbook.sheets.map((sheet) =>
          ids.has(sheet.id) ? { ...sheet, freeze: { rows: 1, cols: 0 } } : sheet,
        ),
      };
    }
    case 'sheet.headerStyle': {
      const patch = {
        ...(step.fontFamily !== undefined ? { fontFamily: step.fontFamily } : {}),
        ...(step.fontSizePt !== undefined ? { fontSizePt: step.fontSizePt } : {}),
        ...(step.bold !== undefined ? { bold: step.bold } : {}),
        ...(step.fill !== undefined ? { fill: step.fill } : {}),
        ...(step.color !== undefined ? { color: step.color } : {}),
      };
      return selected.reduce((current, sheet) => applyHeaderStyle(current, sheet, patch), workbook);
    }
    default:
      throw new MacroError(`宏步骤与文档类型不匹配：${step.op}`);
  }
}

/**
 * 执行一条宏，返回**新的**文档数据，不修改入参。
 *
 * @throws MacroError 步骤校验失败或类型不匹配时抛出，此时文档不会被改动。
 */
export function runMacro<T extends DocModel | WorkbookModel>(kind: PluginDocumentKind, data: T, steps: readonly MacroStep[]): T {
  const issues = validateMacroSteps(steps);
  if (issues.length > 0) throw new MacroError('宏校验未通过', issues);

  const inferred = macroKindOf(steps);
  if (!inferred) throw new MacroError('宏的步骤混用了不同文档类型');
  if (inferred !== kind) throw new MacroError('该命令不适用于当前文档类型');

  let current = data;
  for (const step of steps) {
    current =
      kind === 'docx'
        ? (applyWordStep(current as DocModel, step) as T)
        : (applySheetStep(current as WorkbookModel, step) as T);
  }
  return current;
}
