/** 功能：识别 Word 合并域的引号字段名和可选开关；时间：2026-10-24；作者：lq。 */
export function parseMergeField(instruction: string): string | undefined {
  const match = /^\s*MERGEFIELD\s+(?:"((?:[^"\\]|\\.)*)"|([^\s\\]+))/i.exec(instruction);
  return match ? (match[1]?.replace(/\\(["\\])/g, '$1') ?? match[2]) : undefined;
}

export type MergeFieldFormat = {
  num?: string;        // \# 格式，如 "#,##0.00"
  date?: string;       // \@ 格式，如 "yyyy年M月d日"
  case?: 'upper' | 'lower' | 'caps' | 'firstcap'; // \* 开关
  before?: string;     // \b 非空前缀
  after?: string;      // \f 非空后缀
};

export type MergeDirective =
  | { kind: 'mergeField'; name: string; format: MergeFieldFormat }
  | { kind: 'next' }
  | { kind: 'nextIf'; condition: string }
  | { kind: 'skipIf'; condition: string }
  | { kind: 'record' }   // MERGEREC
  | { kind: 'seq' };     // MERGESEQ

function parseQuotedSwitch(instruction: string, flag: string): string | undefined {
  const re = new RegExp(`\\\\${flag}\\s+"((?:[^"\\\\]|\\\\.)*)"`, 'i');
  const m = re.exec(instruction);
  return m ? m[1].replace(/\\(["\\\\])/g, '$1') : undefined;
}

function parseCaseSwitch(instruction: string): MergeFieldFormat['case'] | undefined {
  const m = /\\\*\s+(upper|lower|caps|firstcap)\b/i.exec(instruction);
  return m ? (m[1].toLowerCase() as MergeFieldFormat['case']) : undefined;
}

export function parseMergeDirective(instruction: string): MergeDirective | undefined {
  const trimmed = instruction.trim();
  const upper = trimmed.toUpperCase();

  if (/^NEXTIF\b/i.test(trimmed)) {
    const m = /^NEXTIF\s+(.+)$/i.exec(trimmed);
    if (m) return { kind: 'nextIf', condition: m[1].trim() };
  }
  if (/^SKIPIF\b/i.test(trimmed)) {
    const m = /^SKIPIF\s+(.+)$/i.exec(trimmed);
    if (m) return { kind: 'skipIf', condition: m[1].trim() };
  }
  if (/^NEXT\b/i.test(trimmed)) return { kind: 'next' };
  if (/^MERGEREC\b/i.test(trimmed)) return { kind: 'record' };
  if (/^MERGESEQ\b/i.test(trimmed)) return { kind: 'seq' };

  const field = /^\s*MERGEFIELD\s+(?:"((?:[^"\\]|\\.)*)"|([^\s\\]+))/i.exec(trimmed);
  if (!field) return undefined;

  const name = (field[1]?.replace(/\\(["\\])/g, '$1') ?? field[2])!;
  const format: MergeFieldFormat = {};
  const num = parseQuotedSwitch(instruction, '#');
  if (num) format.num = num;
  const date = parseQuotedSwitch(instruction, '@');
  if (date) format.date = date;
  const case_ = parseCaseSwitch(instruction);
  if (case_) format.case = case_;
  const before = parseQuotedSwitch(instruction, 'b');
  if (before) format.before = before;
  const after = parseQuotedSwitch(instruction, 'f');
  if (after) format.after = after;

  return { kind: 'mergeField', name, format };
}