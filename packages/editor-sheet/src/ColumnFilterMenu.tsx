// 功能：提供 Excel 风格的表头排序、搜索和勾选筛选浮层；时间：2026-10-24；作者：lq。
import { useEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import { ArrowDownAZ, ArrowUpAZ, X } from 'lucide-react';
import { Button, Input } from '@mini-office/ui';

export interface ColumnFilterMenuProps {
  anchor: { left: number; bottom: number };
  title: string;
  values: string[];
  selected: Set<string>;
  onSelected: (selected: Set<string>) => void;
  onSort: (ascending: boolean) => void;
  onApply: (selected: Set<string>) => void;
  onClear: () => void;
  onClose: () => void;
}

export function ColumnFilterMenu(props: ColumnFilterMenuProps) {
  const [query, setQuery] = useState('');
  const menuRef = useRef<HTMLDivElement>(null);
  const allRef = useRef<HTMLInputElement>(null);
  const visible = props.values.filter((value) => (value || '(空白)').toLocaleLowerCase().includes(query.toLocaleLowerCase()));
  const count = visible.filter((value) => props.selected.has(value)).length;
  useEffect(() => {
    if (allRef.current) allRef.current.indeterminate = count > 0 && count < visible.length;
  }, [count, visible.length]);
  // 功能：点击外部、Esc 或移动视图时关闭浮层，取消未提交条件；时间：2026-10-24；作者：lq。
  useEffect(() => {
    const outside = (event: MouseEvent) => { if (!menuRef.current?.contains(event.target as Node)) props.onClose(); };
    const key = (event: KeyboardEvent) => { if (event.key === 'Escape') { event.stopPropagation(); props.onClose(); } };
    const scroll = (event: Event) => { if (!menuRef.current?.contains(event.target as Node)) props.onClose(); };
    document.addEventListener('mousedown', outside);
    document.addEventListener('keydown', key);
    window.addEventListener('resize', props.onClose);
    document.addEventListener('scroll', scroll, true);
    return () => {
      document.removeEventListener('mousedown', outside);
      document.removeEventListener('keydown', key);
      window.removeEventListener('resize', props.onClose);
      document.removeEventListener('scroll', scroll, true);
    };
  }, [props.onClose]);
  return createPortal(
    <div ref={menuRef} role="dialog" aria-label={`筛选 ${props.title}`} className="mo-glass fixed z-[100] flex w-[280px] flex-col gap-2 rounded-xl border border-[var(--mo-line)] bg-[var(--mo-surface)] p-3 text-[12px] text-[var(--mo-text-primary)] shadow-2xl"
      style={{ left: Math.max(8, Math.min(props.anchor.left, window.innerWidth - 288)), top: Math.max(8, Math.min(props.anchor.bottom + 4, window.innerHeight - 480)), maxHeight: 'calc(100vh - 16px)' }}>
      <div className="flex items-center justify-between"><strong className="truncate">{props.title}</strong><button aria-label="关闭筛选菜单" onClick={props.onClose}><X size={14} /></button></div>
      <div className="flex gap-2"><Button className="flex-1" onClick={() => props.onSort(true)}><ArrowUpAZ size={14} />升序</Button><Button className="flex-1" onClick={() => props.onSort(false)}><ArrowDownAZ size={14} />降序</Button></div>
      <Button variant="ghost" onClick={props.onClear}>清除此列筛选</Button>
      <Input autoFocus aria-label="搜索筛选值" placeholder="搜索" value={query} onChange={(event) => setQuery(event.target.value)} />
      <label className="flex items-center gap-2 border-b border-[var(--mo-line)] pb-2">
        <input ref={allRef} type="checkbox" className="accent-[var(--mo-accent)]" checked={visible.length > 0 && count === visible.length} disabled={!visible.length}
          onChange={(event) => {
            const next = new Set(props.selected);
            visible.forEach((value) => event.target.checked ? next.add(value) : next.delete(value));
            props.onSelected(next);
          }} />{query ? '全选搜索结果' : '(全选)'}
      </label>
      <div className="min-h-0 max-h-[240px] flex-1 overflow-auto">
        {visible.length ? visible.map((value) => <label key={value} className="flex items-center gap-2 rounded px-1 py-1 hover:bg-[var(--mo-surface-2)]">
          <input type="checkbox" className="accent-[var(--mo-accent)]" checked={props.selected.has(value)} onChange={(event) => {
            const next = new Set(props.selected);
            if (event.target.checked) next.add(value); else next.delete(value);
            props.onSelected(next);
          }} /><span className="truncate" title={value}>{value || '(空白)'}</span>
        </label>) : <p className="py-5 text-center text-[var(--mo-text-muted)]">没有匹配项</p>}
      </div>
      <div className="flex justify-end gap-2 border-t border-[var(--mo-line)] pt-2"><Button onClick={props.onClose}>取消</Button><Button variant="primary" disabled={query ? count === 0 : !props.selected.size} onClick={() => props.onApply(query ? new Set(visible.filter((value) => props.selected.has(value))) : props.selected)}>确定</Button></div>
    </div>, document.body,
  );
}
