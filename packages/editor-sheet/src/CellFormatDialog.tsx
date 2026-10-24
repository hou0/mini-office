// 功能：提供 Excel 风格的单元格格式设置，草稿确认后一次应用到整个选区；时间：2026-10-24；作者：lq。
import { useState } from 'react';
import type { BorderSide, CellStyle } from '@mini-office/document-model';
import { Button, ColorField, Dialog, Field, Input, Select } from '@mini-office/ui';
import { FONT_CHOICES } from '@mini-office/design-system';
import { formatCellValue } from './format';
import type { BorderPreset } from './SheetToolbar';

const FORMATS = [
  ['常规', 'General'], ['数值', '0.00'], ['货币', '¥#,##0.00'], ['会计专用', '_("¥"* #,##0.00_);_("¥"* (#,##0.00);_("¥"* "-"??_);_(@_)'],
  ['日期', 'yyyy-MM-dd'], ['时间', 'hh:mm:ss'], ['百分比', '0.00%'], ['分数', '# ?/?'], ['科学记数', '0.00E+00'], ['文本', '@'], ['自定义', ''],
];
const TABS = ['数字', '对齐', '字体', '边框', '填充'] as const;

export function CellFormatDialog({ style, value, initialTab = '数字', onApply, onClose }: {
  initialTab?: typeof TABS[number];
  style: CellStyle; value: string | number | boolean | null;
  onApply: (patch: Partial<CellStyle>, border?: { preset: BorderPreset; side: BorderSide }) => void;
  onClose: () => void;
}) {
  const [tab, setTab] = useState<typeof TABS[number]>(initialTab);
  const [patch, setPatch] = useState<Partial<CellStyle>>({});
  const [category, setCategory] = useState(FORMATS.find(([_, format]) => format === (style.numFmt ?? 'General'))?.[0] ?? '自定义');
  const [border, setBorder] = useState<{ preset: BorderPreset; side: BorderSide }>();
  const [side, setSide] = useState<BorderSide>(style.borders?.bottom ?? { style: 'thin', color: '#000000' });
  const draft = { ...style, ...patch };
  const update = (change: Partial<CellStyle>) => setPatch((current) => ({ ...current, ...change }));
  const chooseBorder = (preset: BorderPreset) => setBorder({ preset, side });
  return <Dialog open title="设置单元格格式" width="md" onClose={onClose} footer={<><Button onClick={onClose}>取消</Button><Button variant="primary" onClick={() => onApply(patch, border)}>确定</Button></>}>
    <div role="tablist" aria-label="单元格格式类别" className="mb-4 flex gap-1 border-b border-[var(--mo-line)]">
      {TABS.map((item) => <button key={item} role="tab" aria-selected={tab === item} className={`px-4 py-2 text-[12px] ${tab === item ? 'border-b-2 border-[var(--mo-accent)] font-semibold text-[var(--mo-accent)]' : ''}`} onClick={() => setTab(item)}>{item}</button>)}
    </div>
    <div role="tabpanel" aria-label={tab} className="min-h-[285px] space-y-4 text-[12px]">
      {tab === '数字' ? <div className="grid grid-cols-[130px_1fr] gap-5">
        <div role="listbox" aria-label="数字格式分类" className="max-h-[285px] overflow-auto rounded border border-[var(--mo-line)] p-1">{FORMATS.map(([label, format]) => <button key={label} role="option" aria-selected={category === label} className={`block w-full rounded px-2 py-1.5 text-left ${category === label ? 'bg-[var(--mo-accent-soft)] text-[var(--mo-accent)]' : ''}`} onClick={() => { setCategory(label); if (format) update({ numFmt: format, numFmtId: undefined }); }}>{label}</button>)}</div>
        <div className="space-y-4"><Field label="示例"><div className="rounded border border-[var(--mo-line)] p-4 text-[16px]">{formatCellValue(value ?? 1234.56, draft)}</div></Field><Field label="格式代码"><Input aria-label="格式代码" value={draft.numFmt ?? 'General'} onChange={(event) => { setCategory('自定义'); update({ numFmt: event.target.value, numFmtId: undefined }); }} /></Field><p className="text-[var(--mo-text-muted)]">选择分类或输入格式代码，如 #,##0.00、yyyy-MM-dd。</p></div>
      </div> : null}
      {tab === '对齐' ? <><div className="grid grid-cols-2 gap-4"><Field label="水平对齐"><Select aria-label="水平对齐方式" value={draft.horizontalAlign ?? 'general'} onChange={(e) => update({ horizontalAlign: e.target.value as CellStyle['horizontalAlign'] })}><option value="general">常规</option><option value="left">左对齐</option><option value="center">居中</option><option value="right">右对齐</option></Select></Field><Field label="垂直对齐"><Select aria-label="垂直对齐方式" value={draft.verticalAlign ?? 'bottom'} onChange={(e) => update({ verticalAlign: e.target.value as CellStyle['verticalAlign'] })}><option value="top">靠上</option><option value="middle">居中</option><option value="bottom">靠下</option></Select></Field><Field label="缩进"><Input aria-label="缩进" type="number" min={0} max={250} value={draft.indent ?? 0} onChange={(e) => update({ indent: Math.max(0, Math.min(250, Number(e.target.value))) })} /></Field></div><label className="flex items-center gap-2"><input type="checkbox" checked={!!draft.wrapText} onChange={(e) => update({ wrapText: e.target.checked })} />自动换行</label><label className="flex items-center gap-2"><input type="checkbox" checked={!!draft.shrinkToFit} onChange={(e) => update({ shrinkToFit: e.target.checked })} />缩小字体填充</label></> : null}
      {tab === '字体' ? <><div className="grid grid-cols-2 gap-4"><Field label="字体"><Select aria-label="格式字体" value={draft.fontFamily ?? '等线'} onChange={(e) => update({ fontFamily: e.target.value })}>{[...new Set([draft.fontFamily ?? '等线', ...FONT_CHOICES])].map((font) => <option key={font}>{font}</option>)}</Select></Field><Field label="字号"><Input aria-label="格式字号" type="number" min={1} max={409} value={draft.fontSizePt ?? 11} onChange={(e) => update({ fontSizePt: Math.max(1, Math.min(409, Number(e.target.value))) })} /></Field></div><div className="flex gap-4">{([['bold', '加粗'], ['italic', '斜体'], ['underline', '下划线'], ['strike', '删除线']] as const).map(([key, label]) => <label key={key} className="flex items-center gap-1"><input type="checkbox" checked={!!draft[key]} onChange={(e) => update({ [key]: e.target.checked })} />{label}</label>)}</div><ColorField label="文字颜色" value={draft.color} onChange={(color) => update({ color })} /></> : null}
      {tab === '边框' ? <><div className="grid grid-cols-2 gap-4"><Field label="线型"><Select aria-label="边框线型" value={side.style} onChange={(e) => { const next = { ...side, style: e.target.value as BorderSide['style'] }; setSide(next); if (border) setBorder({ ...border, side: next }); }}>{[['thin', '细实线'], ['medium', '中实线'], ['thick', '粗实线'], ['dashed', '虚线'], ['dotted', '点线'], ['double', '双线']].map(([key, label]) => <option key={key} value={key}>{label}</option>)}</Select></Field><Field label="线条颜色"><ColorField label="边框颜色" value={side.color} onChange={(color) => { const next = { ...side, color }; setSide(next); if (border) setBorder({ ...border, side: next }); }} /></Field></div><div className="flex gap-2">{([['none', '无'], ['outline', '外边框'], ['all', '所有框线'], ['bottom', '下边框']] as const).map(([preset, label]) => <Button key={preset} variant={border?.preset === preset ? 'primary' : 'secondary'} onClick={() => chooseBorder(preset)}>{label}</Button>)}</div><p className="text-[var(--mo-text-muted)]">外边框仅应用到选区四周，所有框线同时应用内部边框。</p></> : null}
      {tab === '填充' ? <><Field label="背景色"><ColorField label="单元格背景色" value={draft.fill} onChange={(fill) => update({ fill })} /></Field><Button onClick={() => update({ fill: undefined })}>无颜色</Button></> : null}
    </div>
    {tab !== '数字' ? <div className="mt-3 rounded border border-[var(--mo-line)] p-4 text-center" style={{ background: draft.fill ?? 'var(--mo-surface)', color: draft.color, fontFamily: draft.fontFamily, fontSize: `${draft.fontSizePt ?? 11}pt`, fontWeight: draft.bold ? 700 : 400, fontStyle: draft.italic ? 'italic' : undefined, textDecoration: [draft.underline && 'underline', draft.strike && 'line-through'].filter(Boolean).join(' ') }}>{formatCellValue(value ?? '文本示例 AaBb', draft)}</div> : null}
  </Dialog>;
}
