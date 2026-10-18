import { useMemo, useState } from 'react';
import {
  ArrowRight,
  Clock3,
  FilePlus2,
  FileText,
  FolderOpen,
  HardDrive,
  Presentation,
  Search,
  ShieldCheck,
  Sparkles,
  Table2,
  Trash2,
  UploadCloud,
} from 'lucide-react';
import { useWorkspaceStore } from '@mini-office/document-model';
import { Button, IconButton, Input } from '@mini-office/ui';
import { OFFICE_META, formatBytes, formatDateTime, type OfficeKind } from '@mini-office/shared';
import { contractTemplate, expenseTemplate } from './templates';

const KIND_ICONS: Record<OfficeKind, typeof FileText> = {
  docx: FileText,
  xlsx: Table2,
  pptx: Presentation,
};

const KIND_TINTS: Record<OfficeKind, { bg: string; color: string }> = {
  docx: { bg: '#EEF3FF', color: '#2B579A' },
  xlsx: { bg: '#E9F7EF', color: '#217346' },
  pptx: { bg: '#FCEFE9', color: '#C43E1C' },
};

export function HomePage() {
  const recent = useWorkspaceStore((state) => state.recent);
  const createDocument = useWorkspaceStore((state) => state.createDocument);
  const openFileDialog = useWorkspaceStore((state) => state.openFileDialog);
  const openSnapshot = useWorkspaceStore((state) => state.openSnapshot);
  const deleteSnapshot = useWorkspaceStore((state) => state.deleteSnapshot);
  const updateData = useWorkspaceStore((state) => state.updateData);
  const notify = useWorkspaceStore((state) => state.notify);
  const openDialog = useWorkspaceStore((state) => state.openDialog);
  const [query, setQuery] = useState('');
  const [busy, setBusy] = useState(false);

  const filtered = useMemo(() => {
    const keyword = query.trim().toLowerCase();
    if (!keyword) return recent;
    return recent.filter(
      (item) =>
        item.displayName.toLowerCase().includes(keyword) ||
        OFFICE_META[item.kind].label.toLowerCase().includes(keyword),
    );
  }, [query, recent]);

  const dirtyCount = recent.filter((item) => item.dirty).length;

  const createBlank = async (kind: OfficeKind) => {
    if (busy) return;
    setBusy(true);
    try {
      await createDocument(kind);
    } finally {
      setBusy(false);
    }
  };

  const createContract = async () => {
    if (busy) return;
    setBusy(true);
    try {
      const id = await createDocument('docx');
      if (!id) return;
      updateData(id, contractTemplate(), { dirty: true });
      notify('success', '已创建合同模板', '可按需修改条款并另存为正式合同。');
    } finally {
      setBusy(false);
    }
  };

  const createExpense = async () => {
    if (busy) return;
    setBusy(true);
    try {
      const id = await createDocument('xlsx');
      if (!id) return;
      updateData(id, expenseTemplate(), { dirty: true });
      notify('success', '已创建费用报销模板', '公式已自动计算，可直接填写金额。');
    } finally {
      setBusy(false);
    }
  };

  const openFile = async () => {
    if (busy) return;
    setBusy(true);
    try {
      await openFileDialog();
    } finally {
      setBusy(false);
    }
  };

  const openRecent = async (id: string) => {
    if (busy) return;
    setBusy(true);
    try {
      await openSnapshot(id);
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className="mo-home mo-scroll h-full overflow-y-auto bg-[var(--mo-bg)]">
      <div className="mx-auto w-full max-w-[1120px] px-6 pb-16 pt-8 md:px-10">
        <section className="mo-glass relative overflow-hidden rounded-2xl border border-[var(--mo-line)] bg-[var(--mo-surface)] px-7 py-8 shadow-panel md:px-10 md:py-10">
          <div
            aria-hidden
            className="pointer-events-none absolute -right-20 -top-28 h-72 w-72 rounded-full opacity-60 blur-3xl"
            style={{
              background:
                'radial-gradient(circle, color-mix(in srgb, var(--mo-accent) 20%, transparent), transparent 70%)',
            }}
          />
          <div
            aria-hidden
            className="pointer-events-none absolute -bottom-32 right-40 h-64 w-64 rounded-full opacity-50 blur-3xl"
            style={{ background: 'radial-gradient(circle, rgba(18,183,106,0.16), transparent 70%)' }}
          />
          <div className="relative flex flex-col gap-7 md:flex-row md:items-end md:justify-between">
            <div className="max-w-[620px]">
              <div className="mb-4 flex items-center gap-2.5">
                <LogoMark />
                <span className="text-[11px] font-semibold uppercase tracking-[0.18em] text-[var(--mo-text-muted)]">
                  Mini Office
                </span>
              </div>
              <h1 className="text-[30px] font-semibold leading-tight tracking-tight text-[var(--mo-text-primary)] md:text-[36px]">
                只保留真正需要的
                <span className="text-[var(--mo-accent)]"> Office</span> 功能
              </h1>
              <p className="mt-3 max-w-[540px] text-[13px] leading-6 text-[var(--mo-text-muted)]">
                离线优先、轻量快速的跨平台办公套件。Word、Excel、PPT 在同一处打开，
                重视真实文件兼容与文档安全。
              </p>
              <div className="mt-5 flex flex-wrap items-center gap-2">
                <Button variant="primary" icon={<FolderOpen size={15} />} onClick={() => void openFile()} loading={busy}>
                  打开文件
                </Button>
                <Button variant="secondary" icon={<FilePlus2 size={15} />} onClick={() => void createBlank('docx')}>
                  新建文档
                </Button>
                <button
                  type="button"
                  onClick={() => void openFile()}
                  className="ml-1 inline-flex items-center gap-1.5 rounded-md px-2 py-1.5 text-[12px] text-[var(--mo-text-muted)] hover:bg-[var(--mo-surface-2)]"
                >
                  <UploadCloud size={14} /> 或将 DOCX / XLSX / PPTX 拖入窗口
                </button>
              </div>
            </div>
            <div className="flex shrink-0 flex-wrap gap-2 md:max-w-[280px] md:justify-end">
              {[
                { icon: ShieldCheck, label: '离线优先' },
                { icon: HardDrive, label: '本地保存' },
                { icon: Sparkles, label: '中文友好' },
              ].map(({ icon: Icon, label }) => (
                <span
                  key={label}
                  className="inline-flex items-center gap-1.5 rounded-full border border-[var(--mo-line)] bg-[var(--mo-surface-2)]/60 px-2.5 py-1 text-[11px] text-[var(--mo-text-muted)]"
                >
                  <Icon size={12} /> {label}
                </span>
              ))}
            </div>
          </div>
        </section>

        <section className="mt-6 grid gap-4 lg:grid-cols-[1.4fr_1fr]">
          <div className="grid gap-3 sm:grid-cols-3">
            <CreateCard
              kind="docx"
              title="Word 文档"
              description="合同、公文与常用文档编辑"
              actionLabel="新建文档"
              onAction={() => void createBlank('docx')}
            />
            <CreateCard
              kind="xlsx"
              title="Excel 表格"
              description="数据录入、公式计算与打印"
              actionLabel="新建表格"
              onAction={() => void createBlank('xlsx')}
            />
            <CreateCard
              kind="pptx"
              title="PowerPoint"
              description="高保真静态预览与全屏演示"
              actionLabel="打开演示"
              onAction={() => void openFile()}
            />
          </div>

          <div className="rounded-xl border border-[var(--mo-line)] bg-[var(--mo-surface)] p-4 shadow-panel">
            <div className="mb-3 flex items-center justify-between">
              <h2 className="text-[13px] font-semibold">从模板开始</h2>
              <span className="text-[10px] text-[var(--mo-text-faint)]">内置模板</span>
            </div>
            <div className="space-y-2">
              <TemplateRow
                icon={<FileText size={15} />}
                tint="#2B579A"
                title="房屋租赁合同"
                description="标准条款、首行缩进与签章区"
                onClick={() => void createContract()}
              />
              <TemplateRow
                icon={<Table2 size={15} />}
                tint="#217346"
                title="费用报销单"
                description="含 SUM 公式、日期与金额格式"
                onClick={() => void createExpense()}
              />
            </div>
            <div className="mt-3 rounded-lg bg-[var(--mo-surface-2)] px-3 py-2 text-[10px] leading-4 text-[var(--mo-text-muted)]">
              模板创建为未保存副本，保存时会提示选择位置，不会覆盖任何模板文件。
            </div>
          </div>
        </section>

        {dirtyCount > 0 ? (
          <section className="mt-5 flex items-center gap-3 rounded-xl border border-[color-mix(in_srgb,var(--mo-warning)_35%,transparent)] bg-[var(--mo-warning-soft)] px-4 py-3">
            <Clock3 size={16} className="shrink-0 text-[var(--mo-warning)]" />
            <div className="min-w-0 flex-1">
              <div className="text-[12px] font-medium text-[var(--mo-text-primary)]">
                发现 {dirtyCount} 个未保存的自动恢复快照
              </div>
              <div className="text-[10px] text-[var(--mo-text-muted)]">
                快照保存在本机；恢复后可继续编辑并另存为文件。
              </div>
            </div>
            <Button
              size="sm"
              variant="soft"
              onClick={() => {
                const first = recent.find((item) => item.dirty);
                if (first) void openRecent(first.id);
              }}
            >
              恢复最近一个
            </Button>
          </section>
        ) : null}

        <section className="mt-8">
          <div className="mb-3 flex flex-wrap items-center justify-between gap-3">
            <div className="flex items-baseline gap-2">
              <h2 className="text-[15px] font-semibold">最近文件</h2>
              <span className="text-[10px] text-[var(--mo-text-faint)]">
                本机保留 {recent.length} 个恢复快照
              </span>
            </div>
            <div className="relative w-[240px]">
              <Search size={13} className="pointer-events-none absolute left-2.5 top-1/2 -translate-y-1/2 text-[var(--mo-text-faint)]" />
              <Input
                value={query}
                onChange={(event) => setQuery(event.target.value)}
                placeholder="搜索最近文件…"
                className="h-8 pl-8 text-[12px]"
              />
            </div>
          </div>
          <div className="overflow-hidden rounded-xl border border-[var(--mo-line)] bg-[var(--mo-surface)] shadow-panel">
            {filtered.length === 0 ? (
              <div className="flex flex-col items-center justify-center gap-3 px-6 py-14 text-center">
                <div className="flex h-12 w-12 items-center justify-center rounded-2xl bg-[var(--mo-surface-2)] text-[var(--mo-text-muted)]">
                  <FolderOpen size={22} />
                </div>
                <div>
                  <div className="text-[13px] font-medium">{recent.length === 0 ? '还没有最近文件' : '没有匹配的文件'}</div>
                  <div className="mt-1 text-[11px] text-[var(--mo-text-muted)]">
                    {recent.length === 0
                      ? '打开一个 Office 文件，或从上方模板开始。'
                      : '尝试更换搜索关键词。'}
                  </div>
                </div>
                {recent.length === 0 ? (
                  <Button variant="primary" size="sm" icon={<FolderOpen size={14} />} onClick={() => void openFile()}>
                    打开文件
                  </Button>
                ) : null}
              </div>
            ) : (
              <ul className="divide-y divide-[var(--mo-line)]">
                {filtered.map((item) => {
                  const Icon = KIND_ICONS[item.kind];
                  const tint = KIND_TINTS[item.kind];
                  return (
                    <li key={item.id}>
                      <div className="group flex items-center gap-3 px-4 py-3 transition-colors hover:bg-[var(--mo-surface-2)]">
                        <button
                          type="button"
                          onClick={() => void openRecent(item.id)}
                          className="flex min-w-0 flex-1 items-center gap-3 text-left"
                        >
                          <span
                            className="flex h-9 w-9 shrink-0 items-center justify-center rounded-lg"
                            style={{ background: tint.bg, color: tint.color }}
                          >
                            <Icon size={17} />
                          </span>
                          <span className="min-w-0 flex-1">
                            <span className="flex items-center gap-2">
                              <span className="truncate text-[12.5px] font-medium text-[var(--mo-text-primary)]">
                                {item.displayName}
                              </span>
                              {item.dirty ? (
                                <span className="shrink-0 rounded bg-[var(--mo-warning-soft)] px-1 py-0.5 text-[9px] text-[var(--mo-warning)]">
                                  未保存
                                </span>
                              ) : null}
                            </span>
                            <span className="mt-0.5 flex items-center gap-2 text-[10px] text-[var(--mo-text-muted)]">
                              <span>{OFFICE_META[item.kind].label}</span>
                              <span className="text-[var(--mo-text-faint)]">·</span>
                              <span>{formatBytes(item.size)}</span>
                              <span className="text-[var(--mo-text-faint)]">·</span>
                              <span>{formatDateTime(item.updatedAt)}</span>
                              {item.path ? (
                                <>
                                  <span className="text-[var(--mo-text-faint)]">·</span>
                                  <span className="max-w-[260px] truncate">{item.path}</span>
                                </>
                              ) : null}
                            </span>
                          </span>
                          <ArrowRight
                            size={14}
                            className="shrink-0 text-[var(--mo-text-faint)] opacity-0 transition-opacity group-hover:opacity-100"
                          />
                        </button>
                        <IconButton
                          label="从最近文件移除"
                          onClick={() => void deleteSnapshot(item.id)}
                          className="opacity-0 transition-opacity group-hover:opacity-100"
                        >
                          <Trash2 size={13} />
                        </IconButton>
                      </div>
                    </li>
                  );
                })}
              </ul>
            )}
          </div>
        </section>

        <section className="mt-8 flex flex-wrap items-center justify-between gap-3 border-t border-[var(--mo-line)] pt-5 text-[10px] text-[var(--mo-text-faint)]">
          <div className="flex flex-wrap items-center gap-4">
            <span>Mini Office 0.1.0</span>
            <span>离线优先 · 不上传文件</span>
            <span>DOCX / XLSX / PPTX</span>
          </div>
          <button
            type="button"
            onClick={() => openDialog({ type: 'shortcuts' })}
            className="rounded px-1.5 py-1 hover:bg-[var(--mo-surface-2)] hover:text-[var(--mo-text-muted)]"
          >
            键盘快捷键
          </button>
        </section>
      </div>
    </div>
  );
}

function LogoMark({ size = 34 }: { size?: number }) {
  return (
    <span
      className="mo-brand-mark relative inline-flex items-center justify-center rounded-[10px] shadow-sm"
      style={{ width: size, height: size }}
    >
      <span className="font-semibold" style={{ fontSize: size * 0.46 }}>
        M
      </span>
      <span className="absolute right-[5px] top-[5px] h-1 w-1 rounded-full bg-white/80" />
    </span>
  );
}

function CreateCard({
  kind,
  title,
  description,
  actionLabel,
  onAction,
}: {
  kind: OfficeKind;
  title: string;
  description: string;
  actionLabel: string;
  onAction: () => void;
}) {
  const Icon = KIND_ICONS[kind];
  const tint = KIND_TINTS[kind];
  return (
    <button
      type="button"
      onClick={onAction}
      className="group flex flex-col items-start rounded-xl border border-[var(--mo-line)] bg-[var(--mo-surface)] p-4 text-left shadow-panel transition-all duration-fast hover:-translate-y-0.5 hover:border-[var(--mo-line-strong)] hover:shadow-pop"
    >
      <span
        className="mb-4 flex h-10 w-10 items-center justify-center rounded-lg"
        style={{ background: tint.bg, color: tint.color }}
      >
        <Icon size={19} />
      </span>
      <span className="text-[13px] font-semibold text-[var(--mo-text-primary)]">{title}</span>
      <span className="mt-1 text-[10.5px] leading-4 text-[var(--mo-text-muted)]">{description}</span>
      <span className="mt-4 inline-flex items-center gap-1 text-[11px] font-medium text-[var(--mo-accent)]">
        {actionLabel}
        <ArrowRight size={12} className="transition-transform duration-fast group-hover:translate-x-0.5" />
      </span>
    </button>
  );
}

function TemplateRow({
  icon,
  tint,
  title,
  description,
  onClick,
}: {
  icon: React.ReactNode;
  tint: string;
  title: string;
  description: string;
  onClick: () => void;
}) {
  return (
    <button
      type="button"
      onClick={onClick}
      className="flex w-full items-center gap-3 rounded-lg border border-transparent px-2 py-2 text-left transition-colors hover:border-[var(--mo-line)] hover:bg-[var(--mo-surface-2)]"
    >
      <span
        className="flex h-8 w-8 items-center justify-center rounded-md"
        style={{ background: `${tint}14`, color: tint }}
      >
        {icon}
      </span>
      <span className="min-w-0 flex-1">
        <span className="block truncate text-[11.5px] font-medium">{title}</span>
        <span className="block truncate text-[10px] text-[var(--mo-text-muted)]">{description}</span>
      </span>
      <ArrowRight size={13} className="text-[var(--mo-text-faint)]" />
    </button>
  );
}
