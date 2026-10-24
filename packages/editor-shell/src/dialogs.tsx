import { BrandIcon } from '@mini-office/ui';
import { useEffect, useMemo, useState } from 'react';
import {
  AlertTriangle,
  BellRing,
  Check,
  CheckCircle2,
  FileWarning,
  ExternalLink,
  Info,
  Keyboard,
  Languages,
  Monitor,
  Moon,
  RefreshCw,
  Settings as SettingsIcon,
  ShieldCheck,
  Sun,
  Trash2,
  X,
  XCircle,
} from 'lucide-react';
import { useWorkspaceStore } from '@mini-office/document-model';
import {
  Button,
  Dialog,
  Field,
  Kbd,
  SegmentedControl,
  Select,
  Switch,
} from '@mini-office/ui';
import {
  APP_VERSION,
  DEFAULT_ACCENT_HEX,
  MOD_KEY,
  decodeDisplayName,
  type AccentId,
} from '@mini-office/shared';
import {
  ACCENT_PRESETS,
  applyAppearance,
  normalizeAccentHex,
  resolveAccentPalette,
  resolveTheme,
  type ThemeMode,
} from '@mini-office/design-system';
import { isTauriRuntime, openExternalUrl, supportsWindowVibrancy } from '@mini-office/platform-adapter';

// 功能：统一关于页的新版本发布地址；时间：2026-10-24；作者：lq。
const RELEASES_URL = 'https://github.com/lq0910/mini-office/releases';

export function WorkspaceDialogs() {
  const dialog = useWorkspaceStore((state) => state.dialog);
  const closeDialog = useWorkspaceStore((state) => state.closeDialog);
  const sessions = useWorkspaceStore((state) => state.sessions);

  if (!dialog) return null;
  if (dialog.type === 'settings') {
    return (
      <Dialog open onClose={closeDialog} title="设置" icon={<SettingsIcon size={16} />} width="md">
        <SettingsContent onClose={closeDialog} />
      </Dialog>
    );
  }
  if (dialog.type === 'compatibility') {
    const session = dialog.sessionId ? sessions[dialog.sessionId] : undefined;
    return (
      <Dialog
        open
        onClose={closeDialog}
        title="兼容性报告"
        icon={<FileWarning size={16} />}
        width="xl"
        description="Mini Office 按 L0–L3 分级处理 Office 文件，优先保证原文件不被破坏。"
        footer={
          <div className="flex w-full items-center justify-between gap-4">
            <p className="hidden min-w-0 text-[10px] leading-4 text-[var(--mo-text-muted)] sm:block">
              L0 无法打开 · L1 仅预览 · L2 需另存 · L3 安全往返
            </p>
            <Button variant="secondary" size="sm" onClick={closeDialog} className="ml-auto shrink-0">
              关闭
            </Button>
          </div>
        }
      >
        {session ? <CompatibilityContent sessionId={session.id} onClose={closeDialog} /> : null}
      </Dialog>
    );
  }
  if (dialog.type === 'shortcuts') {
    return (
      <Dialog open onClose={closeDialog} title="键盘快捷键" icon={<Keyboard size={16} />} width="lg">
        <ShortcutsContent />
      </Dialog>
    );
  }
  if (dialog.type === 'about') {
    return (
      <Dialog open onClose={closeDialog} title="关于 Mini Office" icon={<Info size={16} />} width="md">
        <AboutContent />
      </Dialog>
    );
  }
  return null;
}

function SettingsContent({ onClose }: { onClose: () => void }) {
  const settings = useWorkspaceStore((state) => state.settings);
  const updateSettings = useWorkspaceStore((state) => state.updateSettings);
  const clearSnapshots = useWorkspaceStore((state) => state.clearSnapshots);
  const notify = useWorkspaceStore((state) => state.notify);

  const resolvedTheme = resolveTheme(settings.theme);
  const canVibrancy = supportsWindowVibrancy();
  const [customHex, setCustomHex] = useState(settings.customAccent);
  const customShades = resolveAccentPalette('custom', settings.customAccent, resolvedTheme);
  const currentShades = resolveAccentPalette(settings.accent, settings.customAccent, resolvedTheme);
  const activeAccentLabel =
    ACCENT_PRESETS.find((preset) => preset.id === settings.accent)?.label ?? '自定义';

  useEffect(() => {
    setCustomHex(settings.customAccent);
  }, [settings.customAccent]);

  const previewAppearance = (patch: {
    theme?: ThemeMode;
    accent?: AccentId;
    customAccent?: string;
  }) => {
    applyAppearance({
      theme: patch.theme ?? settings.theme,
      accent: patch.accent ?? settings.accent,
      customAccent: patch.customAccent ?? settings.customAccent,
    });
  };

  const setTheme = (theme: ThemeMode) => {
    updateSettings({ theme });
    previewAppearance({ theme });
  };

  const chooseAccent = (accent: AccentId) => {
    updateSettings({ accent });
    previewAppearance({ accent });
  };

  const commitCustomHex = (value: string) => {
    const normalized = normalizeAccentHex(value);
    if (!normalized) return false;
    updateSettings({ accent: 'custom', customAccent: normalized });
    previewAppearance({ accent: 'custom', customAccent: normalized });
    return true;
  };

  return (
    <div className="space-y-5">
      <section className="space-y-3">
        <SectionTitle icon={<Sun size={13} />} title="外观" />
        <div className="flex flex-col gap-1.5">
          <span className="text-[12px] font-medium text-[var(--mo-text-primary)]">主题</span>
          <SegmentedControl
            value={settings.theme}
            onChange={setTheme}
            className="w-full"
            options={[
              { value: 'light', label: <span className="inline-flex items-center gap-1"><Sun size={12} /> 浅色</span> },
              { value: 'dark', label: <span className="inline-flex items-center gap-1"><Moon size={12} /> 深色</span> },
              { value: 'system', label: <span className="inline-flex items-center gap-1"><Monitor size={12} /> 跟随系统</span> },
            ]}
          />
        </div>

        <div className="flex flex-col gap-1.5">
          <span className="text-[12px] font-medium text-[var(--mo-text-primary)]">主题色</span>
          <div role="group" aria-label="主题色" className="grid grid-cols-2 gap-1.5 sm:grid-cols-3">
            {ACCENT_PRESETS.map((preset) => {
              const active = settings.accent === preset.id;
              const shades = preset[resolvedTheme];
              return (
                <button
                  key={preset.id}
                  type="button"
                  aria-label={preset.label}
                  aria-pressed={active}
                  title={preset.description}
                  onClick={() => chooseAccent(preset.id)}
                  className={`flex items-center gap-2 rounded-lg border px-2 py-1.5 text-left transition-colors duration-fast ${
                    active
                      ? 'border-[var(--mo-accent)] bg-[var(--mo-accent-soft)] text-[var(--mo-accent)]'
                      : 'border-[var(--mo-line)] text-[var(--mo-text-primary)] hover:bg-[var(--mo-surface-2)]'
                  }`}
                >
                  <span
                    className="flex h-5 w-5 shrink-0 items-center justify-center rounded-full ring-1 ring-black/10"
                    style={{ background: shades.accent }}
                  >
                    {active ? <Check size={11} style={{ color: shades.ink }} /> : null}
                  </span>
                  <span className="min-w-0 flex-1 truncate text-[11.5px] font-medium">
                    {preset.label}
                  </span>
                </button>
              );
            })}
            <button
              type="button"
              aria-label="自定义"
              aria-pressed={settings.accent === 'custom'}
              title="使用自定义颜色"
              onClick={() => chooseAccent('custom')}
              className={`flex items-center gap-2 rounded-lg border px-2 py-1.5 text-left transition-colors duration-fast ${
                settings.accent === 'custom'
                  ? 'border-[var(--mo-accent)] bg-[var(--mo-accent-soft)] text-[var(--mo-accent)]'
                  : 'border-[var(--mo-line)] text-[var(--mo-text-primary)] hover:bg-[var(--mo-surface-2)]'
              }`}
            >
              <span
                className="flex h-5 w-5 shrink-0 items-center justify-center rounded-full ring-1 ring-black/10"
                style={{
                  background:
                    settings.accent === 'custom'
                      ? customShades.accent
                      : 'conic-gradient(from 210deg, #2E7D72, #2F63E0, #6B5BD6, #B25F4A, #2E7D72)',
                }}
              >
                {settings.accent === 'custom' ? (
                  <Check size={11} style={{ color: customShades.ink }} />
                ) : null}
              </span>
              <span className="min-w-0 flex-1 truncate text-[11.5px] font-medium">自定义</span>
            </button>
          </div>
          <span className="text-[11px] leading-4 text-[var(--mo-text-muted)]">
            主题色用于按钮、选区、链接与高亮；浅色/深色会分别自动适配。
          </span>
        </div>

        {settings.accent === 'custom' ? (
          <div className="flex flex-col gap-1.5">
            <span className="text-[12px] font-medium text-[var(--mo-text-primary)]">自定义颜色</span>
            <div className="flex items-center gap-2">
              <input
                type="color"
                aria-label="选择自定义主题色"
                value={normalizeAccentHex(customHex) ?? DEFAULT_ACCENT_HEX}
                onChange={(event) => {
                  setCustomHex(event.target.value);
                  commitCustomHex(event.target.value);
                }}
                className="h-8 w-10 shrink-0 cursor-pointer rounded-md border border-[var(--mo-line-strong)] bg-transparent p-0.5"
              />
              <input
                value={customHex}
                aria-label="主题色十六进制值"
                spellCheck={false}
                onChange={(event) => {
                  const value = event.target.value;
                  setCustomHex(value);
                  if (normalizeAccentHex(value)) commitCustomHex(value);
                }}
                onBlur={() => {
                  if (!normalizeAccentHex(customHex)) setCustomHex(settings.customAccent);
                }}
                className={`mo-field h-8 w-[132px] font-mono text-[11px] ${
                  normalizeAccentHex(customHex) ? '' : 'border-[var(--mo-danger)]'
                }`}
              />
              <span className="text-[10px] text-[var(--mo-text-muted)]">
                当前：{activeAccentLabel}
              </span>
            </div>
            <span className="text-[11px] leading-4 text-[var(--mo-text-muted)]">
              支持 #RGB 或 #RRGGBB，深色模式会自动提亮。
            </span>
          </div>
        ) : null}

        <div className="flex flex-wrap items-center gap-2 rounded-lg border border-[var(--mo-line)] bg-[var(--mo-surface-2)]/45 px-3 py-2.5">
          <span className="rounded-md bg-[var(--mo-accent)] px-2.5 py-1 text-[11px] font-medium text-[var(--mo-accent-ink)]">
            主操作
          </span>
          <span className="rounded-md bg-[var(--mo-accent-soft)] px-2 py-1 text-[11px] font-medium text-[var(--mo-accent)]">
            标签
          </span>
          <span className="text-[11px] font-medium text-[var(--mo-accent)] underline decoration-1 underline-offset-2">
            链接
          </span>
          <span className="ml-auto text-[10px] text-[var(--mo-text-muted)]">
            {activeAccentLabel} · {currentShades.accent}
          </span>
        </div>

        {canVibrancy ? (
          <SettingRow
            title="毛玻璃效果"
            description="使用 macOS 原生磨砂材质，让标题栏与面板透出桌面；深色模式自动适配。"
            control={
              <Switch
                checked={settings.vibrancy}
                onChange={(checked) => updateSettings({ vibrancy: checked })}
                label="毛玻璃效果"
              />
            }
          />
        ) : null}
        {canVibrancy && settings.vibrancy ? (
          <p className="rounded-md bg-[var(--mo-surface-2)] px-3 py-2 text-[10px] leading-4 text-[var(--mo-text-muted)]">
            原生磨砂材质在窗口创建时生效。如果刚更新过应用或效果不明显，请完全退出 Mini Office 后重新打开。
          </p>
        ) : null}

        <Field label="界面语言">
          <Select value={settings.language} onChange={(event) => updateSettings({ language: event.target.value as 'zh-CN' })}>
            <option value="zh-CN">简体中文</option>
          </Select>
        </Field>
        <Field label="默认缩放">
          <Select
            value={String(settings.defaultZoom)}
            onChange={(event) => updateSettings({ defaultZoom: Number(event.target.value) })}
          >
            {[0.75, 0.9, 1, 1.1, 1.25, 1.5].map((value) => (
              <option key={value} value={value}>
                {Math.round(value * 100)}%
              </option>
            ))}
          </Select>
        </Field>
      </section>

      <section className="space-y-2 border-t border-[var(--mo-line)] pt-4">
        <SectionTitle icon={<RefreshCw size={13} />} title="保存与恢复" />
        <SettingRow
          title="自动恢复快照"
          description="编辑时在本地保存恢复快照，异常退出后可继续编辑。"
          control={<Switch checked={settings.autoSave} onChange={(checked) => updateSettings({ autoSave: checked })} label="自动恢复快照" />}
        />
        <Field label="自动恢复间隔">
          <Select
            value={String(settings.autoSaveDelayMs)}
            onChange={(event) => updateSettings({ autoSaveDelayMs: Number(event.target.value) })}
            disabled={!settings.autoSave}
          >
            <option value="800">0.8 秒</option>
            <option value="1400">1.4 秒（推荐）</option>
            <option value="2500">2.5 秒</option>
            <option value="4000">4 秒</option>
          </Select>
        </Field>
        <SettingRow
          title="启动时显示最近文件"
          description="首页展示本机快照与可恢复文件。"
          control={<Switch checked={settings.restoreSession} onChange={(checked) => updateSettings({ restoreSession: checked })} label="显示最近文件" />}
        />
        <SettingRow
          title="显示工具栏"
          description="隐藏后可用命令面板调用高频命令。"
          control={<Switch checked={settings.showToolbar} onChange={(checked) => updateSettings({ showToolbar: checked })} label="显示工具栏" />}
        />
      </section>

      <section className="space-y-2 border-t border-[var(--mo-line)] pt-4">
        <SectionTitle icon={<ShieldCheck size={13} />} title="安全与隐私" />
        <p className="rounded-md bg-[var(--mo-surface-2)] px-3 py-2 text-[11px] leading-5 text-[var(--mo-text-muted)]">
          Mini Office 离线运行，不将文档上传到任何服务器。最近文件与自动恢复数据保存在本机浏览器
          IndexedDB / 应用数据目录中。
        </p>
        <div className="flex items-center justify-between rounded-md border border-[var(--mo-line)] px-3 py-2">
          <div>
            <div className="text-[12px] font-medium">清除所有恢复快照</div>
            <div className="text-[10px] text-[var(--mo-text-muted)]">不会删除磁盘上的原始 Office 文件。</div>
          </div>
          <Button
            size="sm"
            variant="ghost"
            icon={<Trash2 size={13} />}
            className="text-[var(--mo-danger)]"
            onClick={() => {
              if (window.confirm('确定清除所有自动恢复快照与最近文件记录吗？')) {
                void clearSnapshots().then(() => notify('success', '已清除恢复快照'));
              }
            }}
          >
            清除
          </Button>
        </div>
      </section>

      <div className="flex items-center justify-between border-t border-[var(--mo-line)] pt-3">
        {/* 功能：设置页展示统一版本；时间：2026-10-24；作者：lq。 */}
        <span className="text-[10px] text-[var(--mo-text-faint)]">Mini Office {APP_VERSION} · 本地优先</span>
        <Button
          size="sm"
          variant="ghost"
          onClick={() => {
            updateSettings({
              theme: 'system',
              accent: 'mist',
              customAccent: DEFAULT_ACCENT_HEX,
              vibrancy: false,
              autoSave: true,
              autoSaveDelayMs: 1400,
              showProperties: false,
              showToolbar: true,
              restoreSession: true,
            });
            applyAppearance({
              theme: 'system',
              accent: 'mist',
              customAccent: DEFAULT_ACCENT_HEX,
            });
            setCustomHex(DEFAULT_ACCENT_HEX);
          }}
        >
          恢复默认设置
        </Button>
        <Button variant="primary" size="sm" onClick={onClose}>
          完成
        </Button>
      </div>
    </div>
  );
}

function SectionTitle({ icon, title }: { icon: React.ReactNode; title: string }) {
  return (
    <div className="flex items-center gap-1.5 text-[11px] font-semibold text-[var(--mo-text-muted)]">
      {icon}
      {title}
    </div>
  );
}

function SettingRow({
  title,
  description,
  control,
}: {
  title: string;
  description: string;
  control: React.ReactNode;
}) {
  return (
    <div className="flex items-center justify-between gap-4 rounded-md border border-[var(--mo-line)] px-3 py-2">
      <div className="min-w-0">
        <div className="text-[12px] font-medium">{title}</div>
        <div className="mt-0.5 text-[10px] leading-4 text-[var(--mo-text-muted)]">{description}</div>
      </div>
      {control}
    </div>
  );
}

function CompatibilityContent({ sessionId, onClose }: { sessionId: string; onClose: () => void }) {
  const session = useWorkspaceStore((state) => state.sessions[sessionId]);
  const saveSession = useWorkspaceStore((state) => state.saveSession);
  const [saving, setSaving] = useState(false);

  const grouped = useMemo(() => {
    if (!session) return { supported: [], unsupported: [] };
    return {
      supported: session.capability.features.filter((feature) => feature.supported),
      unsupported: session.capability.features.filter((feature) => !feature.supported),
    };
  }, [session]);

  if (!session) return <p className="text-[12px] text-[var(--mo-text-muted)]">文档已关闭。</p>;

  const level = session.capability.level;
  const displayName = decodeDisplayName(session.displayName);
  const tone =
    level >= 3
      ? {
          bg: 'var(--mo-success-soft)',
          color: 'var(--mo-success)',
          label: 'L3 安全往返保存',
          note: '原文件可安全保存；写入前会校验 OOXML 容器，并保留未知部件。',
        }
      : level === 2
        ? {
            bg: 'var(--mo-warning-soft)',
            color: 'var(--mo-warning)',
            label: 'L2 可编辑，建议另存副本',
            note: '文档包含未完全支持的特性，建议另存副本以保护原始文件。',
          }
        : level === 1
          ? {
              bg: 'var(--mo-surface-2)',
              color: 'var(--mo-text-muted)',
              label: 'L1 仅预览 / 只读',
              note: '当前为只读预览，原始文件不会被修改。',
            }
          : {
              bg: 'var(--mo-danger-soft)',
              color: 'var(--mo-danger)',
              label: 'L0 无法打开',
              note: '文件无法安全解析，Mini Office 不会写入原始文件。',
            };
  const warnings = session.capability.warnings.filter((warning) => warning.code !== 'fonts.missing');
  const missingFonts = session.capability.missingFonts;
  const warningTone = {
    blocking: {
      box: 'border-[color-mix(in_srgb,var(--mo-danger)_22%,var(--mo-line))] bg-[var(--mo-danger-soft)]/25',
      icon: 'bg-[var(--mo-danger-soft)] text-[var(--mo-danger)]',
      label: '阻断',
      labelClass: 'text-[var(--mo-danger)]',
    },
    warning: {
      box: 'border-[color-mix(in_srgb,var(--mo-warning)_24%,var(--mo-line))] bg-[var(--mo-warning-soft)]/22',
      icon: 'bg-[var(--mo-warning-soft)] text-[var(--mo-warning)]',
      label: '注意',
      labelClass: 'text-[var(--mo-warning)]',
    },
    info: {
      box: 'border-[var(--mo-line)] bg-[var(--mo-surface-2)]/45',
      icon: 'bg-[var(--mo-accent-soft)] text-[var(--mo-accent)]',
      label: '提示',
      labelClass: 'text-[var(--mo-accent)]',
    },
  } as const;

  return (
    <div className="space-y-4">
      <div className="overflow-hidden rounded-xl border border-[var(--mo-line)] bg-gradient-to-br from-[var(--mo-surface-2)]/70 via-[var(--mo-surface)] to-[var(--mo-surface)]">
        <div className="flex items-start gap-3 p-3.5">
          <div
            className="flex h-11 w-11 shrink-0 items-center justify-center rounded-xl"
            style={{ background: tone.bg, color: tone.color }}
          >
            {level >= 3 ? <CheckCircle2 size={22} /> : level === 1 ? <Info size={22} /> : <AlertTriangle size={22} />}
          </div>
          <div className="min-w-0 flex-1">
            <div className="flex flex-wrap items-center gap-2">
              <span className="text-[14px] font-semibold" style={{ color: tone.color }}>
                {tone.label}
              </span>
              <span className="rounded-full border border-[var(--mo-line)] bg-[var(--mo-surface)]/80 px-1.5 py-0.5 font-mono text-[9px] font-semibold text-[var(--mo-text-muted)]">
                L{level}
              </span>
            </div>
            <div className="mt-1 flex min-w-0 items-center gap-1.5 text-[11px] text-[var(--mo-text-muted)]">
              <span className="truncate" title={displayName}>
                {displayName}
              </span>
              <span className="shrink-0 text-[var(--mo-text-faint)]">·</span>
              <span className="shrink-0">{session.readonly ? '只读模式' : '可编辑'}</span>
            </div>
            <p className="mt-1.5 text-[10.5px] leading-4 text-[var(--mo-text-muted)]">{tone.note}</p>
          </div>
          {level === 2 && !session.readonly ? (
            <Button
              size="sm"
              variant="soft"
              loading={saving}
              className="shrink-0"
              onClick={async () => {
                setSaving(true);
                const ok = await saveSession(session.id, { asNew: true });
                setSaving(false);
                if (ok) onClose();
              }}
            >
              另存副本
            </Button>
          ) : null}
        </div>
      </div>

      {warnings.length > 0 ? (
        <section className="space-y-2">
          <SectionHeading title="兼容性提示" count={warnings.length} />
          {warnings.map((warning, index) => {
            const visual = warningTone[warning.severity];
            return (
              <div
                key={`${warning.code}-${index}`}
                className={`flex gap-2.5 rounded-xl border p-2.5 ${visual.box}`}
              >
                <span className={`flex h-7 w-7 shrink-0 items-center justify-center rounded-lg ${visual.icon}`}>
                  {warning.severity === 'blocking' ? (
                    <XCircle size={14} />
                  ) : warning.severity === 'warning' ? (
                    <AlertTriangle size={14} />
                  ) : (
                    <Info size={14} />
                  )}
                </span>
                <div className="min-w-0 flex-1">
                  <div className="flex flex-wrap items-center gap-1.5">
                    <span className="rounded bg-[var(--mo-surface)]/70 px-1.5 py-0.5 font-mono text-[9px] text-[var(--mo-text-muted)]">
                      {warning.code}
                    </span>
                    <span className={`text-[9px] font-semibold ${visual.labelClass}`}>{visual.label}</span>
                  </div>
                  <p className="mt-1 break-words text-[11.5px] leading-5 text-[var(--mo-text-primary)]">
                    {warning.message}
                  </p>
                  {warning.detail ? (
                    <p className="mt-1 break-words text-[10px] leading-4 text-[var(--mo-text-faint)]">
                      {warning.detail}
                    </p>
                  ) : null}
                </div>
              </div>
            );
          })}
        </section>
      ) : null}

      {missingFonts.length > 0 ? (
        <section className="space-y-2">
          <SectionHeading title="缺失字体" count={missingFonts.length} />
          <div className="rounded-xl border border-[var(--mo-line)] bg-[var(--mo-surface-2)]/35 p-2.5">
            <div className="flex flex-wrap gap-1.5">
              {missingFonts.map((font) => (
                <span
                  key={font}
                  title={font}
                  className="max-w-full truncate rounded-md border border-[var(--mo-line)] bg-[var(--mo-surface)] px-2 py-1 text-[10px] text-[var(--mo-text-primary)]"
                >
                  {font}
                </span>
              ))}
            </div>
            <p className="mt-2 text-[10px] leading-4 text-[var(--mo-text-muted)]">
              缺失字体将使用替代字体显示，版面可能与原始文件不同。建议在目标系统安装相应字体。
            </p>
          </div>
        </section>
      ) : null}

      <section className="space-y-2">
        <SectionHeading title="能力概览" />
        <div className="rounded-xl border border-[var(--mo-line)] bg-[var(--mo-surface-2)]/30 p-2.5">
          {grouped.supported.length === 0 && grouped.unsupported.length === 0 ? (
            <p className="text-[10px] text-[var(--mo-text-faint)]">暂无能力记录</p>
          ) : (
            <div className="flex flex-wrap gap-1.5">
              {grouped.supported.map((feature) => (
                <span
                  key={feature.code}
                  className="inline-flex max-w-full items-center gap-1 rounded-md border border-[color-mix(in_srgb,var(--mo-success)_18%,var(--mo-line))] bg-[var(--mo-success-soft)]/25 px-1.5 py-1 text-[10px] text-[var(--mo-text-primary)]"
                >
                  <Check size={10} className="shrink-0 text-[var(--mo-success)]" />
                  <span className="min-w-0 truncate" title={feature.label}>{feature.label}</span>
                </span>
              ))}
              {grouped.unsupported.map((feature) => (
                <span
                  key={feature.code}
                  className="inline-flex max-w-full items-center gap-1 rounded-md border border-[color-mix(in_srgb,var(--mo-warning)_20%,var(--mo-line))] bg-[var(--mo-warning-soft)]/30 px-1.5 py-1 text-[10px] text-[var(--mo-text-primary)]"
                >
                  <X size={10} className="shrink-0 text-[var(--mo-warning)]" />
                  <span className="min-w-0 truncate" title={feature.label}>{feature.label}</span>
                </span>
              ))}
            </div>
          )}
        </div>
      </section>

    </div>
  );
}

function SectionHeading({ title, count }: { title: string; count?: number }) {
  return (
    <div className="flex items-center gap-2">
      <h3 className="text-[10.5px] font-semibold uppercase tracking-[0.04em] text-[var(--mo-text-muted)]">
        {title}
      </h3>
      {typeof count === 'number' ? (
        <span className="rounded-full bg-[var(--mo-surface-2)] px-1.5 py-0.5 text-[9px] font-medium tabular-nums text-[var(--mo-text-muted)]">
          {count}
        </span>
      ) : null}
      <span className="h-px flex-1 bg-[var(--mo-line)]" />
    </div>
  );
}

const SHORTCUTS: Array<{ group: string; items: Array<{ label: string; keys: string[] }> }> = [
  {
    group: '文件',
    items: [
      { label: '打开文件', keys: [MOD_KEY, 'O'] },
      { label: '新建 Word 文档', keys: [MOD_KEY, 'N'] },
      { label: '保存', keys: [MOD_KEY, 'S'] },
      { label: '另存为', keys: ['⇧', MOD_KEY, 'S'] },
      { label: '关闭标签', keys: [MOD_KEY, 'W'] },
      { label: '打印', keys: [MOD_KEY, 'P'] },
    ],
  },
  {
    group: '编辑',
    items: [
      { label: '撤销', keys: [MOD_KEY, 'Z'] },
      { label: '重做', keys: ['⇧', MOD_KEY, 'Z'] },
      { label: '查找替换', keys: [MOD_KEY, 'F'] },
      { label: '复制 / 粘贴', keys: [MOD_KEY, 'C / V'] },
      { label: '全选', keys: [MOD_KEY, 'A'] },
    ],
  },
  {
    group: '视图',
    items: [
      { label: '命令面板', keys: [MOD_KEY, 'K'] },
      { label: '设置', keys: [MOD_KEY, ','] },
      // 功能：将 Word 文档缩放快捷键加入内置帮助；时间：2026-10-24；作者：lq。
      { label: '放大 Word 文档', keys: [MOD_KEY, '+'] },
      { label: '缩小 Word 文档', keys: [MOD_KEY, '-'] },
      { label: 'Word 文档恢复 100%', keys: [MOD_KEY, '0'] },
      { label: 'Word 文档滚轮缩放', keys: [MOD_KEY, '滚轮'] },
      // 功能：将 Excel 工作表缩放快捷键加入内置帮助；时间：2026-10-24；作者：lq。
      { label: 'Excel 工作表放大', keys: [MOD_KEY, '+'] },
      { label: 'Excel 工作表缩小', keys: [MOD_KEY, '-'] },
      { label: 'Excel 工作表恢复 100%', keys: [MOD_KEY, '0'] },
      { label: 'Excel 工作表滚轮缩放', keys: [MOD_KEY, '滚轮'] },
      { label: '全屏演示（PPT）', keys: ['F5'] },
      { label: '退出演示 / 关闭弹窗', keys: ['Esc'] },
    ],
  },
];

function ShortcutsContent() {
  return (
    <div className="space-y-4">
      {SHORTCUTS.map((section) => (
        <section key={section.group}>
          <h3 className="mb-2 text-[10px] font-semibold uppercase tracking-wide text-[var(--mo-text-faint)]">
            {section.group}
          </h3>
          <div className="space-y-1">
            {section.items.map((item) => (
              <div
                key={item.label}
                className="flex items-center justify-between rounded-md px-2 py-1.5 text-[12px] hover:bg-[var(--mo-surface-2)]"
              >
                <span>{item.label}</span>
                <span className="flex items-center gap-1">
                  {item.keys.map((key, index) => (
                    <Kbd key={`${key}-${index}`}>{key}</Kbd>
                  ))}
                </span>
              </div>
            ))}
          </div>
        </section>
      ))}
      <p className="text-[10px] leading-4 text-[var(--mo-text-muted)]">
        在表格中，方向键移动选区，Shift + 方向键扩展选区，输入任意字符开始编辑，F2 编辑当前单元格。
      </p>
    </div>
  );
}

function AboutContent() {
  return (
    <div className="space-y-4">
      <div className="flex items-center gap-3 rounded-lg bg-[var(--mo-surface-2)] p-3">
        {/* 功能：关于页使用正式图标并标明出品方；时间：2026-10-18；作者：lq。 */}
        <BrandIcon size={52} />
        <div>
          <div className="text-[15px] font-semibold">Mini Office</div>
          {/* 功能：更正出品方名称；时间：2026-10-24；作者：lq。 */}
          <div className="mt-0.5 text-[12px] font-medium text-[var(--mo-accent)]">非人类程序员出品</div>
          {/* 功能：关于页展示统一版本；时间：2026-10-24；作者：lq。 */}
          <div className="flex flex-wrap items-center gap-x-2 gap-y-1 text-[11px] text-[var(--mo-text-muted)]">
            <span>版本 {APP_VERSION} · 离线优先的轻量办公套件</span>
            <a
              href={RELEASES_URL}
              target="_blank"
              rel="noreferrer"
              // 功能：从关于页跳转到新版本发布页，桌面端交由系统浏览器打开；时间：2026-10-24；作者：lq。
              onClick={(event) => {
                if (!isTauriRuntime()) return;
                event.preventDefault();
                void openExternalUrl(RELEASES_URL);
              }}
              className="inline-flex items-center gap-1 rounded px-1 py-0.5 text-[var(--mo-accent)] hover:bg-[var(--mo-surface-2)] hover:underline"
            >
              新版本更新
              <ExternalLink size={11} />
            </a>
          </div>
        </div>
      </div>
      <p className="text-[12px] leading-6 text-[var(--mo-text-muted)]">
        使用 Rust + Tauri 2、React、TypeScript 构建。支持 DOCX 合同编辑、XLSX 表格与公式、
        PPTX 基础内容编辑与高保真演示，并通过能力扫描与安全另存策略保护原始文件。
      </p>
      <div className="grid grid-cols-2 gap-2 text-[11px]">
        <div className="rounded-md border border-[var(--mo-line)] p-2.5">
          <div className="text-[var(--mo-text-muted)]">文档模型</div>
          <div className="mt-0.5 font-medium">OOXML 双模式 · L0–L3</div>
        </div>
        <div className="rounded-md border border-[var(--mo-line)] p-2.5">
          <div className="text-[var(--mo-text-muted)]">数据存储</div>
          <div className="mt-0.5 font-medium">本地快照 · 原子写入</div>
        </div>
        <div className="rounded-md border border-[var(--mo-line)] p-2.5">
          <div className="text-[var(--mo-text-muted)]">公式引擎</div>
          <div className="mt-0.5 font-medium">60+ 常用函数</div>
        </div>
        <div className="rounded-md border border-[var(--mo-line)] p-2.5">
          <div className="text-[var(--mo-text-muted)]">导出</div>
          <div className="mt-0.5 font-medium">PDF · 图片 ZIP</div>
        </div>
      </div>
      {/* 功能：展示出品方公众号原始二维码，保留完整留白供微信扫码；时间：2026-10-24；作者：lq。 */}
      <div className="flex items-center justify-between gap-4 rounded-lg border border-[var(--mo-line)] bg-[var(--mo-surface-2)] p-3">
        <div className="min-w-0 space-y-1.5">
          <div className="text-[13px] font-semibold">非人类程序员</div>
          <div className="text-[12px] text-[var(--mo-text-muted)]">微信扫一扫，关注公众号</div>
        </div>
        <img src="/brand/non-human-programmer-wechat.jpg" alt="非人类程序员微信公众号二维码" width={172} height={172} draggable={false} className="shrink-0 bg-white" style={{ imageRendering: 'pixelated' }} />
      </div>
      <div className="flex items-center gap-2 text-[10px] text-[var(--mo-text-faint)]">
        <ShieldCheck size={12} /> 不收集文档内容 · 不上传文件 · 本地优先
      </div>
    </div>
  );
}
