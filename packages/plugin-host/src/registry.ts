/**
 * 插件注册表：已安装插件、远端索引、当前主题包的唯一状态源。
 *
 * 只持久化元数据（清单快照）到 localStorage，插件文件本体放在 IndexedDB。
 * 这样即使文件存储损坏，UI 仍能展示"已安装但文件缺失"并给出重装入口。
 */
import { create } from 'zustand';
import {
  BUILTIN_THEME_PACK,
  DEFAULT_PLUGIN_REPOSITORY,
  type CommandContribution,
  type InstalledPlugin,
  type PluginDocumentKind,
  type PluginIndex,
  type PluginIndexEntry,
  type PluginRepositoryRef,
  type TemplateContribution,
  type ThemeContribution,
  type ValidationIssue,
} from '@mini-office/plugin-api';
import { APP_VERSION } from '@mini-office/shared';
import { downloadPluginPackage, fetchPluginIndex, PluginIndexError } from './github';
import { installFromBytes, PluginPackageError } from './installer';
import { macroKindOf } from './macro';
import { compareVersions } from './manifest';
import { applyThemePack, themePackKey } from './theme';
import { createPluginFileStore, type PluginFileStore } from './storage';

const INSTALLED_KEY = 'mini-office.plugins.installed.v1';
const ACTIVE_THEME_KEY = 'mini-office.plugins.active-theme.v1';

export interface ThemePackOption {
  key: string;
  label: string;
  description?: string;
  preview?: string;
  pluginId: string;
  pluginName: string;
  theme: ThemeContribution;
}

/** 一条来自插件的宏命令，已解析出目标文档类型。 */
export interface PluginMacroCommand {
  /** `<pluginId>#<commandId>`，用作命令面板的稳定 key。 */
  key: string;
  pluginId: string;
  pluginName: string;
  kind: PluginDocumentKind;
  command: CommandContribution;
}

/** 一个来自插件的模板。 */
export interface PluginTemplateOption {
  key: string;
  pluginId: string;
  pluginName: string;
  template: TemplateContribution;
}

export interface PluginFailure {
  message: string;
  hint: string;
  issues: ValidationIssue[];
}

/** 已安装插件在远端出现了更高版本。 */
export interface PluginUpdate {
  pluginId: string;
  name: string;
  from: string;
  to: string;
  entry: PluginIndexEntry;
}

export interface PluginRegistryState {
  initialized: boolean;
  installed: InstalledPlugin[];
  index: PluginIndex | null;
  indexOrigin: 'network' | 'cache' | 'offline' | null;
  indexFetchedAt: number | null;
  indexLoading: boolean;
  indexError: PluginFailure | null;
  installingId: string | null;
  repository: PluginRepositoryRef;
  activeThemePack: string;
  /** 安装时的非阻断告警，按插件 id 归档，用于在卡片上提示"装了但可能不完整"。 */
  pluginWarnings: Record<string, ValidationIssue[]>;
  /** 最近一次运行失败的原因，按插件 id 归档。成功运行后清除。 */
  pluginErrors: Record<string, string>;

  initialize(): Promise<void>;
  refreshIndex(options?: { force?: boolean }): Promise<void>;
  installLocalPackage(bytes: Uint8Array): Promise<InstalledPlugin>;
  installFromIndex(pluginId: string): Promise<InstalledPlugin>;
  uninstall(pluginId: string): Promise<void>;
  setEnabled(pluginId: string, enabled: boolean): Promise<void>;
  setActiveThemePack(key: string): void;
  setRepository(repo: PluginRepositoryRef): void;
  themePacks(): ThemePackOption[];
  macroCommands(): PluginMacroCommand[];
  templates(): PluginTemplateOption[];
  updates(): PluginUpdate[];
  reportPluginError(pluginId: string, message: string | null): void;
}

const store: PluginFileStore = createPluginFileStore();

function readInstalled(): InstalledPlugin[] {
  if (typeof localStorage === 'undefined') return [];
  try {
    const raw = localStorage.getItem(INSTALLED_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw) as InstalledPlugin[];
    if (!Array.isArray(parsed)) return [];
    return parsed.filter((item) => item && typeof item.id === 'string' && item.manifest);
  } catch {
    return [];
  }
}

function writeInstalled(installed: InstalledPlugin[]): void {
  if (typeof localStorage === 'undefined') return;
  try {
    localStorage.setItem(INSTALLED_KEY, JSON.stringify(installed));
  } catch {
    // 元数据写入失败不阻断当前会话。
  }
}

function readActiveTheme(): string {
  if (typeof localStorage === 'undefined') return BUILTIN_THEME_PACK;
  return localStorage.getItem(ACTIVE_THEME_KEY) || BUILTIN_THEME_PACK;
}

function writeActiveTheme(key: string): void {
  if (typeof localStorage === 'undefined') return;
  try {
    localStorage.setItem(ACTIVE_THEME_KEY, key);
  } catch {
    // ignore
  }
}

function toFailure(error: unknown): PluginFailure {
  if (error instanceof PluginIndexError) {
    return { message: error.message, hint: error.hint, issues: [] };
  }
  if (error instanceof PluginPackageError) {
    return { message: error.message, hint: '请检查插件包是否完整、清单是否符合规范。', issues: error.issues };
  }
  return {
    message: error instanceof Error ? error.message : '未知错误',
    hint: '如果问题持续出现，请卸载该插件后重试。',
    issues: [],
  };
}

/** 从按 id 归档的记录里移除一项，没有该键时返回原对象以避免无谓渲染。 */
function withoutKey<T>(record: Record<string, T>, key: string): Record<string, T> {
  if (!(key in record)) return record;
  const next = { ...record };
  delete next[key];
  return next;
}

function withWarnings(
  current: Record<string, ValidationIssue[]>,
  pluginId: string,
  warnings: ValidationIssue[],
): Record<string, ValidationIssue[]> {
  if (warnings.length === 0) return withoutKey(current, pluginId);
  return { ...current, [pluginId]: warnings };
}

/** 收集所有启用插件中的主题。 */
export function collectThemePacks(installed: InstalledPlugin[]): ThemePackOption[] {
  const packs: ThemePackOption[] = [];
  for (const plugin of installed) {
    if (!plugin.enabled) continue;
    for (const theme of plugin.manifest.contributes.themes ?? []) {
      packs.push({
        key: themePackKey(plugin.id, theme.id),
        label: theme.label,
        description: theme.description,
        preview: theme.preview,
        pluginId: plugin.id,
        pluginName: plugin.manifest.name,
        theme,
      });
    }
  }
  return packs;
}

/** 收集所有启用插件中的宏命令。类型无法推断的命令在安装期就已被拒绝，这里只是防御。 */
export function collectMacroCommands(installed: InstalledPlugin[]): PluginMacroCommand[] {
  const commands: PluginMacroCommand[] = [];
  for (const plugin of installed) {
    if (!plugin.enabled) continue;
    for (const command of plugin.manifest.contributes.commands ?? []) {
      const kind = macroKindOf(command.steps ?? []);
      if (!kind) continue;
      commands.push({
        key: `${plugin.id}#${command.id}`,
        pluginId: plugin.id,
        pluginName: plugin.manifest.name,
        kind,
        command,
      });
    }
  }
  return commands;
}

/** 收集所有启用插件中的模板。 */
export function collectTemplates(installed: InstalledPlugin[]): PluginTemplateOption[] {
  const templates: PluginTemplateOption[] = [];
  for (const plugin of installed) {
    if (!plugin.enabled) continue;
    for (const template of plugin.manifest.contributes.templates ?? []) {
      templates.push({
        key: `${plugin.id}#${template.id}`,
        pluginId: plugin.id,
        pluginName: plugin.manifest.name,
        template,
      });
    }
  }
  return templates;
}

/** 读取插件包内的文件，用于加载模板。 */
export async function readPluginFile(pluginId: string, path: string): Promise<Uint8Array | null> {
  return store.read(pluginId, path);
}

/**
 * 对比已安装插件与远端索引，找出可更新的项。
 *
 * 只认"远端版本更高"。本地装的比远端新（侧载开发版）时不出现在列表里，
 * 避免用旧包覆盖新包——那才是真正会让用户丢功能的行为。
 */
export function collectUpdates(installed: InstalledPlugin[], index: PluginIndex | null): PluginUpdate[] {
  if (!index) return [];
  const updates: PluginUpdate[] = [];
  for (const plugin of installed) {
    const entry = index.plugins.find((item) => item.id === plugin.id);
    if (!entry) continue;
    const diff = compareVersions(entry.version, plugin.version);
    if (diff === null || diff <= 0) continue;
    updates.push({
      pluginId: plugin.id,
      name: plugin.manifest.name,
      from: plugin.version,
      to: entry.version,
      entry,
    });
  }
  return updates;
}

function resolveActivePack(installed: InstalledPlugin[], key: string): ThemePackOption | null {
  if (!key || key === BUILTIN_THEME_PACK) return null;
  return collectThemePacks(installed).find((pack) => pack.key === key) ?? null;
}

function applyActive(installed: InstalledPlugin[], key: string): void {
  const pack = resolveActivePack(installed, key);
  if (!pack) {
    applyThemePack(null, 'light');
    return;
  }
  applyThemePack({ key: pack.key, theme: pack.theme }, 'light');
}

export const usePluginRegistry = create<PluginRegistryState>((set, get) => ({
  initialized: false,
  installed: [],
  index: null,
  indexOrigin: null,
  indexFetchedAt: null,
  indexLoading: false,
  indexError: null,
  installingId: null,
  repository: DEFAULT_PLUGIN_REPOSITORY,
  activeThemePack: BUILTIN_THEME_PACK,
  pluginWarnings: {},
  pluginErrors: {},

  async initialize() {
    if (get().initialized) return;
    const installed = readInstalled();
    let activeThemePack = readActiveTheme();
    // 主题插件被卸载或禁用后，回退到内置外观，避免停留在空白令牌状态。
    if (activeThemePack !== BUILTIN_THEME_PACK && !resolveActivePack(installed, activeThemePack)) {
      activeThemePack = BUILTIN_THEME_PACK;
      writeActiveTheme(activeThemePack);
    }
    applyActive(installed, activeThemePack);
    set({ installed, activeThemePack, initialized: true });
  },

  async refreshIndex(options = {}) {
    if (get().indexLoading) return;
    set({ indexLoading: true, indexError: null });
    try {
      const result = await fetchPluginIndex(get().repository, options);
      set({
        index: result.index,
        indexOrigin: result.origin,
        indexFetchedAt: result.fetchedAt,
        indexLoading: false,
      });
    } catch (error) {
      set({ indexError: toFailure(error), indexLoading: false });
    }
  },

  async installLocalPackage(bytes) {
    const { plugin, warnings } = await installFromBytes(store, bytes, 'local', {
      currentAppVersion: APP_VERSION,
    });
    const next = [...get().installed.filter((item) => item.id !== plugin.id), plugin];
    writeInstalled(next);
    set({
      installed: next,
      pluginWarnings: withWarnings(get().pluginWarnings, plugin.id, warnings),
      // 重装即视为修好了之前的运行错误。
      pluginErrors: withoutKey(get().pluginErrors, plugin.id),
    });
    return plugin;
  },

  async installFromIndex(pluginId) {
    const index = get().index;
    const entry = index?.plugins.find((item) => item.id === pluginId);
    if (!entry) {
      throw new PluginIndexError('插件不在当前列表中', '请刷新插件列表后重试。');
    }
    set({ installingId: pluginId });
    try {
      const { bytes } = await downloadPluginPackage(entry);
      const { plugin, warnings } = await installFromBytes(store, bytes, 'github', {
        currentAppVersion: APP_VERSION,
        download: entry.download,
      });
      const next = [...get().installed.filter((item) => item.id !== plugin.id), plugin];
      writeInstalled(next);
      set({
        installed: next,
        installingId: null,
        pluginWarnings: withWarnings(get().pluginWarnings, plugin.id, warnings),
        pluginErrors: withoutKey(get().pluginErrors, plugin.id),
      });
      return plugin;
    } catch (error) {
      set({ installingId: null });
      throw error;
    }
  },

  async uninstall(pluginId) {
    await store.remove(pluginId);
    const next = get().installed.filter((item) => item.id !== pluginId);
    writeInstalled(next);
    let activeThemePack = get().activeThemePack;
    if (activeThemePack !== BUILTIN_THEME_PACK && !resolveActivePack(next, activeThemePack)) {
      activeThemePack = BUILTIN_THEME_PACK;
      writeActiveTheme(activeThemePack);
    }
    applyActive(next, activeThemePack);
    set({
      installed: next,
      activeThemePack,
      pluginWarnings: withoutKey(get().pluginWarnings, pluginId),
      pluginErrors: withoutKey(get().pluginErrors, pluginId),
    });
  },

  async setEnabled(pluginId, enabled) {
    const next = get().installed.map((item) => (item.id === pluginId ? { ...item, enabled } : item));
    writeInstalled(next);
    let activeThemePack = get().activeThemePack;
    if (activeThemePack !== BUILTIN_THEME_PACK && !resolveActivePack(next, activeThemePack)) {
      activeThemePack = BUILTIN_THEME_PACK;
      writeActiveTheme(activeThemePack);
    }
    applyActive(next, activeThemePack);
    set({ installed: next, activeThemePack });
  },

  setActiveThemePack(key) {
    const installed = get().installed;
    writeActiveTheme(key);
    applyActive(installed, key);
    set({ activeThemePack: key });
  },

  setRepository(repo) {
    set({ repository: repo, index: null, indexError: null, indexOrigin: null });
  },

  themePacks() {
    return collectThemePacks(get().installed);
  },

  macroCommands() {
    return collectMacroCommands(get().installed);
  },

  templates() {
    return collectTemplates(get().installed);
  },

  updates() {
    return collectUpdates(get().installed, get().index);
  },

  reportPluginError(pluginId, message) {
    const next = message ? { ...get().pluginErrors, [pluginId]: message } : withoutKey(get().pluginErrors, pluginId);
    set({ pluginErrors: next });
  },
}));

/** 供非 React 场景（启动脚本、测试）直接取用。 */
export const pluginRegistry = usePluginRegistry;
