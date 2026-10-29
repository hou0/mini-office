import { useCallback, useMemo, useRef, useState } from 'react';
import {
  AlertTriangle,
  Check,
  Download,
  ExternalLink,
  Package,
  Palette,
  Puzzle,
  RefreshCw,
  ShieldCheck,
  Trash2,
  Upload,
} from 'lucide-react';
import {
  PLUGIN_PERMISSION_LABEL,
  PLUGIN_TIER_LABEL,
  type InstalledPlugin,
  type PluginIndexEntry,
} from '@mini-office/plugin-api';
import { collectThemePacks, collectUpdates, usePluginRegistry, type PluginUpdate } from '@mini-office/plugin-host';
import { APP_VERSION } from '@mini-office/shared';
import { Badge, Button, EmptyState, Spinner, Switch } from '@mini-office/ui';
import { ThemePreview } from './ThemePreview';

function formatDate(timestamp: number | null): string {
  if (!timestamp) return '尚未拉取';
  return new Date(timestamp).toLocaleString('zh-CN', { hour12: false });
}

function formatBytes(size?: number): string {
  if (!size) return '';
  if (size < 1024) return `${size} B`;
  if (size < 1024 * 1024) return `${(size / 1024).toFixed(0)} KB`;
  return `${(size / 1024 / 1024).toFixed(1)} MB`;
}

const TIER_TONE = {
  declarative: 'success',
  sandboxed: 'accent',
  trusted: 'warning',
} as const;

function TierBadge({ tier }: { tier: InstalledPlugin['tier'] }) {
  return <Badge tone={TIER_TONE[tier]}>{PLUGIN_TIER_LABEL[tier]}</Badge>;
}

function SourceBadge({ source }: { source: InstalledPlugin['source'] }) {
  const label = source === 'github' ? 'GitHub' : source === 'local' ? '本地安装' : '内置';
  return <Badge>{label}</Badge>;
}

export function PluginManager() {
  const installed = usePluginRegistry((state) => state.installed);
  const index = usePluginRegistry((state) => state.index);
  const indexLoading = usePluginRegistry((state) => state.indexLoading);
  const indexError = usePluginRegistry((state) => state.indexError);
  const indexOrigin = usePluginRegistry((state) => state.indexOrigin);
  const indexFetchedAt = usePluginRegistry((state) => state.indexFetchedAt);
  const installingId = usePluginRegistry((state) => state.installingId);
  const repository = usePluginRegistry((state) => state.repository);
  const activeThemePack = usePluginRegistry((state) => state.activeThemePack);
  const refreshIndex = usePluginRegistry((state) => state.refreshIndex);
  const installLocalPackage = usePluginRegistry((state) => state.installLocalPackage);
  const installFromIndex = usePluginRegistry((state) => state.installFromIndex);
  const uninstall = usePluginRegistry((state) => state.uninstall);
  const setEnabled = usePluginRegistry((state) => state.setEnabled);
  const setActiveThemePack = usePluginRegistry((state) => state.setActiveThemePack);
  const pluginWarnings = usePluginRegistry((state) => state.pluginWarnings);
  const pluginErrors = usePluginRegistry((state) => state.pluginErrors);
  const reportPluginError = usePluginRegistry((state) => state.reportPluginError);

  const fileInputRef = useRef<HTMLInputElement>(null);
  const [busyLocal, setBusyLocal] = useState(false);
  const [actionError, setActionError] = useState<{ message: string; hint: string } | null>(null);
  const [notice, setNotice] = useState<string | null>(null);

  const themePacks = useMemo(() => collectThemePacks(installed), [installed]);
  const updates = useMemo(() => collectUpdates(installed, index), [installed, index]);
  const updateMap = useMemo(() => new Map(updates.map((item) => [item.pluginId, item])), [updates]);
  const installedIds = useMemo(() => new Set(installed.map((item) => item.id)), [installed]);
  const available = useMemo(
    () => (index?.plugins ?? []).filter((entry) => !installedIds.has(entry.id)),
    [index, installedIds],
  );

  const handleLocalFile = useCallback(
    async (file: File | undefined) => {
      if (!file) return;
      setBusyLocal(true);
      setActionError(null);
      setNotice(null);
      try {
        const bytes = new Uint8Array(await file.arrayBuffer());
        const plugin = await installLocalPackage(bytes);
        setNotice(`已安装「${plugin.manifest.name}」${plugin.version}`);
      } catch (error) {
        setActionError({
          message: error instanceof Error ? error.message : '安装失败',
          hint: '请确认文件是有效的 .moext 插件包，且 manifest.json 符合规范。',
        });
      } finally {
        setBusyLocal(false);
        if (fileInputRef.current) fileInputRef.current.value = '';
      }
    },
    [installLocalPackage],
  );

  const handleRemoteInstall = useCallback(
    async (entry: PluginIndexEntry) => {
      setActionError(null);
      setNotice(null);
      try {
        const plugin = await installFromIndex(entry.id);
        setNotice(`已安装「${plugin.manifest.name}」${plugin.version}`);
      } catch (error) {
        setActionError({
          message: error instanceof Error ? error.message : '安装失败',
          hint: '请刷新插件列表后重试；若持续失败，可能是插件包已更新，索引尚未同步。',
        });
      }
    },
    [installFromIndex],
  );

  const handleUpdate = useCallback(
    async (update: PluginUpdate) => {
      setActionError(null);
      setNotice(null);
      try {
        const plugin = await installFromIndex(update.pluginId);
        setNotice(`已把「${plugin.manifest.name}」更新到 ${plugin.version}`);
      } catch (error) {
        setActionError({
          message: error instanceof Error ? error.message : '更新失败',
          hint: '旧版本仍可继续使用；若反复失败，请卸载后重新安装。',
        });
      }
    },
    [installFromIndex],
  );

  const handleUninstall = useCallback(
    async (plugin: InstalledPlugin) => {
      if (!window.confirm(`确定卸载「${plugin.manifest.name}」吗？`)) return;
      setActionError(null);
      try {
        await uninstall(plugin.id);
        setNotice(`已卸载「${plugin.manifest.name}」`);
      } catch (error) {
        setActionError({ message: error instanceof Error ? error.message : '卸载失败', hint: '请稍后重试。' });
      }
    },
    [uninstall],
  );

  return (
    <div className="space-y-5">
      <section className="space-y-2.5">
        <div className="flex items-center gap-1.5 text-[11px] font-semibold text-[var(--mo-text-muted)]">
          <Package size={13} />
          插件来源
        </div>
        <div className="rounded-md border border-[var(--mo-line)] px-3 py-2.5">
          <div className="flex flex-wrap items-center justify-between gap-2">
            <div className="min-w-0">
              <div className="truncate text-[12px] font-medium text-[var(--mo-text-primary)]">
                github.com/{repository.owner}/{repository.repo}
              </div>
              <div className="text-[10px] text-[var(--mo-text-muted)]">
                分支 {repository.branch} · {repository.indexPath} · 上次更新 {formatDate(indexFetchedAt)}
              </div>
            </div>
            <div className="flex shrink-0 items-center gap-1.5">
              <Button
                size="sm"
                variant="secondary"
                icon={indexLoading ? <Spinner size={12} /> : <RefreshCw size={13} />}
                disabled={indexLoading}
                onClick={() => void refreshIndex({ force: true })}
              >
                刷新列表
              </Button>
              <Button
                size="sm"
                variant="primary"
                icon={busyLocal ? <Spinner size={12} /> : <Upload size={13} />}
                disabled={busyLocal}
                onClick={() => fileInputRef.current?.click()}
              >
                安装本地插件
              </Button>
            </div>
          </div>
          <p className="mt-2 border-t border-[var(--mo-line)] pt-2 text-[10px] leading-4 text-[var(--mo-text-muted)]">
            插件以 .moext 包分发，本质是 zip。安装前会校验清单、主题令牌与包完整性，校验不通过不会写入。
          </p>
        </div>
        <input
          ref={fileInputRef}
          type="file"
          accept=".moext,.zip"
          className="hidden"
          onChange={(event) => void handleLocalFile(event.target.files?.[0])}
        />
        {indexOrigin === 'offline' ? (
          <p className="rounded-md bg-[var(--mo-warning-soft)] px-3 py-2 text-[10px] leading-4 text-[var(--mo-warning)]">
            当前显示的是本地缓存的插件列表，网络恢复后点「刷新列表」可获取最新版本。
          </p>
        ) : null}
      </section>

      {actionError ? (
        <div className="rounded-md border border-[var(--mo-danger)]/40 bg-[var(--mo-danger-soft)] px-3 py-2">
          <div className="flex items-start gap-1.5 text-[11px] font-medium text-[var(--mo-danger)]">
            <AlertTriangle size={13} className="mt-0.5 shrink-0" />
            {actionError.message}
          </div>
          <p className="mt-1 pl-5 text-[10px] leading-4 text-[var(--mo-text-muted)]">{actionError.hint}</p>
        </div>
      ) : null}

      {notice ? (
        <div className="flex items-center gap-1.5 rounded-md bg-[var(--mo-success-soft)] px-3 py-2 text-[11px] text-[var(--mo-success)]">
          <Check size={13} />
          {notice}
        </div>
      ) : null}

      <section className="space-y-2 border-t border-[var(--mo-line)] pt-4">
        <div className="flex flex-wrap items-center gap-1.5 text-[11px] font-semibold text-[var(--mo-text-muted)]">
          <Puzzle size={13} />
          已安装（{installed.length}）
          {updates.length > 0 ? <Badge tone="accent">{updates.length} 个可更新</Badge> : null}
        </div>
        {installed.length === 0 ? (
          <EmptyState
            title="还没有安装插件"
            description="可以从上方列表安装，或选择本地的 .moext 文件。"
          />
        ) : (
          <div className="space-y-1.5">
            {installed.map((plugin) => {
              const themes = plugin.manifest.contributes.themes ?? [];
              const commands = plugin.manifest.contributes.commands ?? [];
              const templates = plugin.manifest.contributes.templates ?? [];
              const update = updateMap.get(plugin.id);
              const warnings = pluginWarnings[plugin.id] ?? [];
              const lastError = pluginErrors[plugin.id];
              return (
                <div key={plugin.id} className="rounded-md border border-[var(--mo-line)] px-3 py-2.5">
                  <div className="flex flex-wrap items-start justify-between gap-2">
                    <div className="min-w-0">
                      <div className="flex flex-wrap items-center gap-1.5">
                        <span className="truncate text-[12px] font-medium text-[var(--mo-text-primary)]">
                          {plugin.manifest.name}
                        </span>
                        <Badge>{plugin.version}</Badge>
                        <TierBadge tier={plugin.tier} />
                        <SourceBadge source={plugin.source} />
                        {update ? (
                          <Badge tone="accent">
                            可更新 {update.from} → {update.to}
                          </Badge>
                        ) : null}
                      </div>
                      <div className="mt-0.5 text-[10px] text-[var(--mo-text-muted)]">
                        {plugin.manifest.publisher} · {plugin.id}
                      </div>
                      {plugin.manifest.description ? (
                        <p className="mt-1 text-[10px] leading-4 text-[var(--mo-text-muted)]">
                          {plugin.manifest.description}
                        </p>
                      ) : null}
                      <div className="mt-1.5 flex flex-wrap items-center gap-1.5">
                        {themes.length > 0 ? (
                          <Badge tone="accent">{themes.length} 个主题</Badge>
                        ) : null}
                        {commands.length > 0 ? <Badge>{commands.length} 个命令</Badge> : null}
                        {templates.length > 0 ? <Badge>{templates.length} 个模板</Badge> : null}
                        {plugin.manifest.permissions.length > 0 ? (
                          <Badge tone="warning">{plugin.manifest.permissions.length} 项权限</Badge>
                        ) : null}
                      </div>
                      {commands.length > 0 ? (
                        <p className="mt-1 text-[10px] leading-4 text-[var(--mo-text-faint)]">
                          命令为宏：由你手动触发，只执行白名单内的文档操作，不运行插件代码。
                        </p>
                      ) : null}
                    </div>
                    <div className="flex shrink-0 items-center gap-1.5">
                      {update ? (
                        <Button
                          size="sm"
                          variant="primary"
                          disabled={installingId === plugin.id}
                          icon={installingId === plugin.id ? <Spinner size={12} /> : <Download size={13} />}
                          onClick={() => void handleUpdate(update)}
                        >
                          {installingId === plugin.id ? '更新中' : '更新'}
                        </Button>
                      ) : null}
                      <Switch
                        checked={plugin.enabled}
                        onChange={(checked) => void setEnabled(plugin.id, checked)}
                        label={`启用 ${plugin.manifest.name}`}
                      />
                      <Button
                        size="sm"
                        variant="ghost"
                        icon={<Trash2 size={13} />}
                        className="text-[var(--mo-danger)]"
                        onClick={() => void handleUninstall(plugin)}
                      >
                        卸载
                      </Button>
                    </div>
                  </div>
                  {plugin.manifest.permissions.length > 0 ? (
                    <p className="mt-1.5 border-t border-[var(--mo-line)] pt-1.5 text-[10px] leading-4 text-[var(--mo-text-muted)]">
                      权限：{plugin.manifest.permissions.map((p) => PLUGIN_PERMISSION_LABEL[p]).join('、')}
                    </p>
                  ) : null}
                  {warnings.length > 0 ? (
                    <div className="mt-1.5 rounded bg-[var(--mo-warning-soft)] px-2 py-1.5">
                      <div className="flex items-center gap-1 text-[10px] font-medium text-[var(--mo-warning)]">
                        <AlertTriangle size={11} />
                        安装告警（插件已可用，但以下项未通过检查）
                      </div>
                      <ul className="mt-0.5 space-y-0.5">
                        {warnings.slice(0, 4).map((issue, i) => (
                          <li key={`${issue.path}-${i}`} className="text-[10px] leading-4 text-[var(--mo-text-muted)]">
                            {issue.path ? `${issue.path}：` : ''}
                            {issue.message}
                          </li>
                        ))}
                      </ul>
                    </div>
                  ) : null}
                  {lastError ? (
                    <div className="mt-1.5 flex items-start justify-between gap-2 rounded bg-[var(--mo-danger-soft)] px-2 py-1.5">
                      <div className="min-w-0">
                        <div className="text-[10px] font-medium text-[var(--mo-danger)]">最近一次运行失败</div>
                        <p className="mt-0.5 break-words text-[10px] leading-4 text-[var(--mo-text-muted)]">
                          {lastError}
                        </p>
                      </div>
                      <button
                        type="button"
                        onClick={() => reportPluginError(plugin.id, null)}
                        className="shrink-0 text-[10px] text-[var(--mo-text-muted)] hover:underline"
                      >
                        清除
                      </button>
                    </div>
                  ) : null}
                </div>
              );
            })}
          </div>
        )}
      </section>

      {themePacks.length > 0 ? (
        <section className="space-y-2 border-t border-[var(--mo-line)] pt-4">
          <div className="flex items-center gap-1.5 text-[11px] font-semibold text-[var(--mo-text-muted)]">
            <Palette size={13} />
            主题插件
          </div>
          <p className="text-[10px] leading-4 text-[var(--mo-text-muted)]">
            主题只改变界面外壳，不会改动文档画布的字体度量，因此不影响分页与打印结果。
          </p>
          <div className="space-y-1.5">
            {themePacks.map((pack) => {
              const active = activeThemePack === pack.key;
              return (
                <div
                  key={pack.key}
                  className="flex items-center justify-between gap-2 rounded-md border border-[var(--mo-line)] px-3 py-2"
                >
                  <div className="flex min-w-0 items-center gap-2">
                    <ThemePreview pack={pack} className="h-7 w-10 shrink-0 rounded" />
                    <div className="min-w-0">
                      <div className="truncate text-[12px] font-medium text-[var(--mo-text-primary)]">
                        {pack.label}
                      </div>
                      <div className="truncate text-[10px] text-[var(--mo-text-muted)]">
                        {pack.pluginName}
                        {pack.description ? ` · ${pack.description}` : ''}
                      </div>
                    </div>
                  </div>
                  <Button
                    size="sm"
                    variant={active ? 'soft' : 'secondary'}
                    icon={active ? <Check size={13} /> : undefined}
                    onClick={() => setActiveThemePack(active ? 'builtin' : pack.key)}
                  >
                    {active ? '使用中' : '应用'}
                  </Button>
                </div>
              );
            })}
          </div>
        </section>
      ) : null}

      <section className="space-y-2 border-t border-[var(--mo-line)] pt-4">
        <div className="flex items-center gap-1.5 text-[11px] font-semibold text-[var(--mo-text-muted)]">
          <ShieldCheck size={13} />
          可用插件
        </div>
        {indexError ? (
          <div className="rounded-md border border-[var(--mo-line)] px-3 py-2.5">
            <div className="flex items-start gap-1.5 text-[11px] font-medium text-[var(--mo-text-primary)]">
              <AlertTriangle size={13} className="mt-0.5 shrink-0 text-[var(--mo-warning)]" />
              {indexError.message}
            </div>
            <p className="mt-1 pl-5 text-[10px] leading-4 text-[var(--mo-text-muted)]">{indexError.hint}</p>
            {indexError.issues.length > 0 ? (
              <ul className="mt-1.5 space-y-0.5 pl-5">
                {indexError.issues.slice(0, 6).map((issue, i) => (
                  <li key={`${issue.path}-${i}`} className="text-[10px] leading-4 text-[var(--mo-text-muted)]">
                    {issue.path ? `${issue.path}：` : ''}
                    {issue.message}
                  </li>
                ))}
              </ul>
            ) : null}
          </div>
        ) : null}
        {indexLoading && !index ? (
          <div className="flex items-center justify-center gap-2 py-6 text-[11px] text-[var(--mo-text-muted)]">
            <Spinner size={14} />
            正在拉取插件列表…
          </div>
        ) : available.length === 0 ? (
          <EmptyState
            title={index ? '暂无可安装的插件' : '尚未拉取插件列表'}
            description={index ? '列表中已全部安装。' : '点击上方「刷新列表」从 GitHub 获取官方插件。'}
          />
        ) : (
          <div className="space-y-1.5">
            {available.map((entry) => (
              <div key={entry.id} className="rounded-md border border-[var(--mo-line)] px-3 py-2.5">
                <div className="flex flex-wrap items-start justify-between gap-2">
                  <div className="min-w-0">
                    <div className="flex flex-wrap items-center gap-1.5">
                      <span className="truncate text-[12px] font-medium text-[var(--mo-text-primary)]">
                        {entry.name}
                      </span>
                      <Badge>{entry.version}</Badge>
                      <TierBadge tier={entry.tier} />
                    </div>
                    <div className="mt-0.5 text-[10px] text-[var(--mo-text-muted)]">
                      {entry.publisher}
                      {entry.size ? ` · ${formatBytes(entry.size)}` : ''}
                    </div>
                    {entry.description ? (
                      <p className="mt-1 text-[10px] leading-4 text-[var(--mo-text-muted)]">{entry.description}</p>
                    ) : null}
                    {entry.tags && entry.tags.length > 0 ? (
                      <div className="mt-1.5 flex flex-wrap gap-1">
                        {entry.tags.map((tag) => (
                          <Badge key={tag}>{tag}</Badge>
                        ))}
                      </div>
                    ) : null}
                  </div>
                  <Button
                    size="sm"
                    variant="primary"
                    disabled={installingId === entry.id}
                    icon={installingId === entry.id ? <Spinner size={12} /> : undefined}
                    onClick={() => void handleRemoteInstall(entry)}
                  >
                    {installingId === entry.id ? '安装中' : '安装'}
                  </Button>
                </div>
              </div>
            ))}
          </div>
        )}
      </section>

      <div className="flex items-center justify-between border-t border-[var(--mo-line)] pt-3">
        <span className="text-[10px] text-[var(--mo-text-faint)]">
          Mini Office {APP_VERSION} · 插件 API 1.0
        </span>
        <a
          href={`https://github.com/${repository.owner}/${repository.repo}`}
          target="_blank"
          rel="noreferrer"
          className="inline-flex items-center gap-1 text-[10px] text-[var(--mo-accent)] hover:underline"
        >
          查看插件仓库
          <ExternalLink size={11} />
        </a>
      </div>
    </div>
  );
}
