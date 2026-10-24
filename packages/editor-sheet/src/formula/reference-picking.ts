// 功能：定位公式光标处的引用替换范围，并协调网格点选/拖选与输入框；时间：2026-10-24；作者：lq。
export interface ReferenceSpan { start: number; end: number }
export type ReferencePickPhase = 'start' | 'update' | 'end';
type ReferencePicker = (phase: ReferencePickPhase, reference: string) => boolean;
const pickers = new WeakMap<HTMLInputElement, ReferencePicker>();

// 功能：只在公式操作数位置插入引用，避免改写字符串、函数名或已结束的表达式；时间：2026-10-24；作者：lq。
export function formulaReferenceSpan(text: string, start: number, end = start): ReferenceSpan | null {
  if (!text.startsWith('=') || start < 1) return null;
  let quoted = '';
  for (let i = 1; i < start; i++) {
    const char = text[i];
    if (quoted) { if (char === quoted) { if (text[i + 1] === quoted) i++; else quoted = ''; } }
    else if (char === '"' || char === "'") quoted = char;
  }
  if (quoted) return null;
  if (end > start) return { start, end };
  const reference = /\$?[A-Za-z]{1,3}\$?[1-9]\d*(?::\$?[A-Za-z]{1,3}\$?[1-9]\d*)?/g;
  let match: RegExpExecArray | null;
  while ((match = reference.exec(text))) {
    const last = match.index + match[0].length;
    if (match.index <= start && start <= last && !/[A-Za-z0-9_.!]/.test(text[match.index - 1] ?? '') && !/[A-Za-z0-9_(]/.test(text[last] ?? '')) return { start: match.index, end: last };
  }
  if (/[=+\-*/^&(<>,;:]\s*$/.test(text.slice(0, start))) return { start, end };
  return null;
}

// 功能：按输入框实例注册引用接收器，公式栏与单元格编辑使用同一套交互；时间：2026-10-24；作者：lq。
export function registerReferencePicker(input: HTMLInputElement, picker: ReferencePicker): () => void {
  pickers.set(input, picker);
  return () => { if (pickers.get(input) === picker) pickers.delete(input); };
}
export function pickFormulaReference(input: HTMLInputElement, phase: ReferencePickPhase, reference: string): boolean {
  return pickers.get(input)?.(phase, reference) ?? false;
}

// 功能：提交点选后的函数公式时补齐末尾右括号，忽略字符串和带引号的工作表名称；时间：2026-10-24；作者：lq。
export function completeFormulaParentheses(text: string): string {
  if (!text.startsWith('=') || /[+\-*/^&,;:(=<>]\s*$/.test(text)) return text;
  let quoted = ''; let balance = 0;
  for (let i = 1; i < text.length; i++) {
    const char = text[i];
    if (quoted) { if (char === quoted) { if (text[i + 1] === quoted) i++; else quoted = ''; } continue; }
    if (char === '"' || char === "'") quoted = char;
    else if (char === '(') balance++;
    else if (char === ')' && --balance < 0) return text;
  }
  return !quoted && balance > 0 ? text + ')'.repeat(balance) : text;
}
