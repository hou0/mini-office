import type { DocParagraph, HeaderFooter, WordAlignment } from '@mini-office/document-model';

/** 功能：常用页码样式使用真正的 PAGE / NUMPAGES 域；时间：2026-10-24；作者：lq。 */
export const PAGE_NUMBER_TEMPLATES = [
  { label: '仅文字', value: '' },
  { label: '1', value: '{PAGE}' },
  { label: '— 1 —', value: '— {PAGE} —' },
  { label: '第 1 页', value: '第 {PAGE} 页' },
  { label: '第 1 页 / 共 3 页', value: '第 {PAGE} 页 / 共 {NUMPAGES} 页' },
  { label: '1 / 3', value: '{PAGE} / {NUMPAGES}' },
  { label: 'Page 1 of 3', value: 'Page {PAGE} of {NUMPAGES}' },
];

/** 功能：从文字和页码占位符构造可保存的页眉页脚，支持多行；时间：2026-10-24；作者：lq。 */
export function makeHeaderFooter(text: string, alignment: WordAlignment, fontSizePt: number, distancePx: number): HeaderFooter {
  const blocks: DocParagraph[] = text.split('\n').map((line, index) => ({
    type: 'paragraph', id: `hf-${index}`, style: { alignment, lineHeight: 1.2, spaceAfterPx: 0 },
    runs: (line ? line.split(/(\{PAGE\}|\{NUMPAGES\})/g).filter(Boolean) : ['']).map((part) => ({
      style: { fontSizePt },
      ...(part === '{PAGE}' || part === '{NUMPAGES}' ? { field: part.slice(1, -1), text: '1' } : { text: part }),
    })),
  }));
  return { blocks, distancePx, pageNumbers: text.includes('{PAGE}'), edited: true };
}

/** 功能：将简单页眉页脚还原为可编辑模板，复杂内容需用户主动替换；时间：2026-10-24；作者：lq。 */
export function headerFooterText(value?: HeaderFooter): string {
  return value?.blocks.map((block) => block.type === 'paragraph'
    ? block.runs.map((run) => run.field ? `{${run.field}}` : run.tab ? '\t' : run.lineBreak ? '\n' : run.text ?? '').join('')
    : '').join('\n') ?? '';
}
