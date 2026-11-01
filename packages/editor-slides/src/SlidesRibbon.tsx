import { Ribbon, type RibbonTab } from '@mini-office/ui';
import {
  Bold,
  ChevronLeft,
  ChevronRight,
  Copy,
  Download,
  FileWarning,
  Italic,
  PanelLeft,
  Play,
  Plus,
  Redo2,
  StickyNote,
  Trash2,
  Type,
  Undo2,
} from 'lucide-react';

/** Slides 工具栏/Ribbon 共用的动作集合，由 SlidesEditor 注入。 */
export interface SlidesToolbarProps {
  readonly?: boolean;
  canUndo: boolean;
  canRedo: boolean;
  thumbnailsOpen: boolean;
  notesOpen: boolean;
  selectedId?: string | null;
  slideIndex: number;
  slideCount: number;
  onToggleThumbnails: () => void;
  onToggleNotes: () => void;
  onPrev: () => void;
  onNext: () => void;
  onBold: () => void;
  onItalic: () => void;
  onUndo: () => void;
  onRedo: () => void;
  onAddText: () => void;
  onDuplicate: () => void;
  onAddSlide: () => void;
  onDeleteSlide: () => void;
  onPresent: () => void;
  onExport: () => void;
  onCompatibility: () => void;
}

/** 功能：Slides 功能区（Ribbon）布局；时间：2026-10-31；作者：lq。 */
export function SlidesRibbon(props: SlidesToolbarProps) {
  const disabled = props.readonly;
  const tabs: RibbonTab[] = [
    {
      id: 'home',
      label: '开始',
      groups: [
        {
          id: 'clipboard',
          label: '剪贴板',
          items: [
            { id: 'undo', label: '撤销', icon: <Undo2 size={15} />, disabled: !props.canUndo || disabled, onClick: props.onUndo },
            { id: 'redo', label: '重做', icon: <Redo2 size={15} />, disabled: !props.canRedo || disabled, onClick: props.onRedo },
          ],
        },
        {
          id: 'font',
          label: '字体',
          items: [
            { id: 'bold', label: '粗体', icon: <Bold size={16} />, hideLabel: true, disabled, onClick: props.onBold },
            { id: 'italic', label: '斜体', icon: <Italic size={16} />, hideLabel: true, disabled, onClick: props.onItalic },
          ],
        },
        {
          id: 'slides',
          label: '幻灯片',
          items: [
            { id: 'prev', label: '上一页', icon: <ChevronLeft size={16} />, hideLabel: true, disabled: props.slideIndex <= 0, onClick: props.onPrev },
            { id: 'next', label: '下一页', icon: <ChevronRight size={16} />, hideLabel: true, disabled: props.slideIndex >= props.slideCount - 1, onClick: props.onNext },
            {
              id: 'index',
              kind: 'inline',
              content: (
                <span className="px-1 text-center text-[11px] tabular-nums text-[var(--mo-text-muted)]">
                  {props.slideIndex + 1} / {props.slideCount}
                </span>
              ),
            },
            { id: 'duplicate', label: '复制', icon: <Copy size={15} />, disabled, onClick: props.onDuplicate },
            { id: 'add', label: '新建', icon: <Plus size={15} />, disabled, onClick: props.onAddSlide },
            { id: 'delete', label: '删除', icon: <Trash2 size={15} />, disabled: props.slideCount <= 1, onClick: props.onDeleteSlide },
          ],
        },
      ],
    },
    {
      id: 'insert',
      label: '插入',
      groups: [
        {
          id: 'objects',
          label: '对象',
          items: [{ id: 'text', label: '文本框', icon: <Type size={15} />, disabled, onClick: props.onAddText }],
        },
      ],
    },
    {
      id: 'view',
      label: '视图',
      groups: [
        {
          id: 'show',
          label: '显示与播放',
          items: [
            { id: 'thumbnails', label: '缩略图', icon: <PanelLeft size={15} />, active: props.thumbnailsOpen, onClick: props.onToggleThumbnails },
            { id: 'notes', label: '备注', icon: <StickyNote size={15} />, active: props.notesOpen, onClick: props.onToggleNotes },
            { id: 'present', label: '演示', icon: <Play size={15} />, onClick: props.onPresent },
            { id: 'export', label: '导出', icon: <Download size={15} />, onClick: props.onExport },
            { id: 'compat', label: '兼容报告', icon: <FileWarning size={15} />, onClick: props.onCompatibility },
          ],
        },
      ],
    },
  ];

  return <Ribbon tabs={tabs} defaultTab="home" />;
}
