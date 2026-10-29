/**
 * 插件包解析与安装。
 *
 * 包格式：`.moext`（本质是 zip）
 * ```
 * manifest.json      必填，清单与主题令牌都在这里
 * preview.png        可选，主题预览图
 * assets/*           可选，主题引用的本地资源
 * ```
 *
 * 安全约束：条目数、总解压体积、路径穿越、清单合法性、主题令牌白名单与对比度
 * 全部在写入存储之前完成校验。任何一项失败都不落地。
 */
import JSZip from 'jszip';
import {
  PLUGIN_MANIFEST_FILE,
  PLUGIN_MAX_ENTRIES,
  PLUGIN_MAX_UNCOMPRESSED_BYTES,
  type InstalledPlugin,
  type PluginManifest,
  type PluginSource,
  type ThemeContribution,
  type ValidationIssue,
} from '@mini-office/plugin-api';
import { parseManifestBytes } from './manifest';
import { validateTheme } from './theme';
import type { PluginFileStore } from './storage';

export class PluginPackageError extends Error {
  readonly issues: ValidationIssue[];
  constructor(message: string, issues: ValidationIssue[] = []) {
    super(message);
    this.name = 'PluginPackageError';
    this.issues = issues;
  }
}

export interface ParsedPackage {
  manifest: PluginManifest;
  themes: ThemeContribution[];
  /** 去掉包内公共前缀后的相对路径 → 字节。 */
  files: Map<string, Uint8Array>;
  warnings: ValidationIssue[];
}

/** 拒绝绝对路径、反斜杠与 `..` 穿越。 */
export function isSafeEntryPath(path: string): boolean {
  if (!path || path.startsWith('/') || path.includes('\\')) return false;
  return !path.split('/').some((segment) => segment === '..' || segment === '');
}

function commonPrefixOf(manifestPath: string): string {
  const index = manifestPath.lastIndexOf('/');
  return index < 0 ? '' : manifestPath.slice(0, index + 1);
}

/** 解析 `.moext` 包，完成全部静态校验。 */
export async function parsePackage(bytes: Uint8Array, currentAppVersion?: string): Promise<ParsedPackage> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(bytes);
  } catch {
    throw new PluginPackageError('无法读取插件包', [
      { path: '', message: '文件不是有效的 zip 包，或已损坏。插件包扩展名应为 .moext。' },
    ]);
  }

  const entries = Object.values(zip.files).filter((entry) => !entry.dir);
  if (entries.length === 0) {
    throw new PluginPackageError('插件包为空', [{ path: '', message: '压缩包内没有任何文件。' }]);
  }
  if (entries.length > PLUGIN_MAX_ENTRIES) {
    throw new PluginPackageError('插件包条目过多', [
      { path: '', message: `条目数 ${entries.length} 超过上限 ${PLUGIN_MAX_ENTRIES}。` },
    ]);
  }

  const unsafe = entries.map((entry) => entry.name).filter((name) => !isSafeEntryPath(name));
  if (unsafe.length > 0) {
    throw new PluginPackageError('插件包包含非法路径', unsafe.map((name) => ({ path: name, message: '路径非法' })));
  }

  const manifestEntry = entries
    .filter((entry) => entry.name === PLUGIN_MANIFEST_FILE || entry.name.endsWith(`/${PLUGIN_MANIFEST_FILE}`))
    .sort((a, b) => a.name.length - b.name.length)[0];
  if (!manifestEntry) {
    throw new PluginPackageError('插件包缺少清单', [
      { path: PLUGIN_MANIFEST_FILE, message: '包根目录必须包含 manifest.json。' },
    ]);
  }

  const prefix = commonPrefixOf(manifestEntry.name);
  const manifestBytes = await manifestEntry.async('uint8array');
  const parsedManifest = parseManifestBytes(manifestBytes, currentAppVersion);
  if (!parsedManifest.ok || !parsedManifest.manifest) {
    throw new PluginPackageError('插件清单校验未通过', parsedManifest.issues);
  }

  const manifest = parsedManifest.manifest;
  const files = new Map<string, Uint8Array>();
  let totalBytes = 0;

  for (const entry of entries) {
    const relative = prefix && entry.name.startsWith(prefix) ? entry.name.slice(prefix.length) : entry.name;
    if (!relative) continue;
    const content = await entry.async('uint8array');
    totalBytes += content.byteLength;
    if (totalBytes > PLUGIN_MAX_UNCOMPRESSED_BYTES) {
      throw new PluginPackageError('插件包解压后体积过大', [
        { path: '', message: `解压后超过 ${Math.round(PLUGIN_MAX_UNCOMPRESSED_BYTES / 1024 / 1024)} MB，已中止安装。` },
      ]);
    }
    files.set(relative, content);
  }

  const themes: ThemeContribution[] = [];
  const warnings = [...parsedManifest.warnings];
  const themeIssues: ValidationIssue[] = [];
  (manifest.contributes.themes ?? []).forEach((theme, index) => {
    const result = validateTheme(theme, index);
    themeIssues.push(...result.issues);
    warnings.push(...result.warnings);
    if (result.ok && result.theme) themes.push(result.theme);
  });

  if (themeIssues.length > 0) {
    throw new PluginPackageError('主题定义校验未通过', themeIssues);
  }
  if (themes.length === 0 && (manifest.contributes.themes?.length ?? 0) > 0) {
    throw new PluginPackageError('主题定义校验未通过', [{ path: 'contributes.themes', message: '没有任何可用主题。' }]);
  }

  // 模板必须真的在包里，否则用户点了模板才会失败——安装期就挡住。
  const missingTemplates: ValidationIssue[] = [];
  for (const template of manifest.contributes.templates ?? []) {
    if (!files.has(template.file)) {
      missingTemplates.push({
        path: `contributes.templates.${template.id}.file`,
        message: `包内找不到模板文件 ${template.file}`,
      });
    }
  }
  if (missingTemplates.length > 0) {
    throw new PluginPackageError('模板文件缺失', missingTemplates);
  }

  return { manifest, themes, files, warnings };
}

export interface InstallResult {
  plugin: InstalledPlugin;
  warnings: ValidationIssue[];
}

/** 把已解析的包写入存储，并返回安装记录。 */
export async function commitPackage(
  store: PluginFileStore,
  parsed: ParsedPackage,
  source: PluginSource,
  download?: string,
): Promise<InstallResult> {
  const { manifest } = parsed;
  // 覆盖安装前先清空旧文件，避免残留资源造成版本错乱。
  await store.remove(manifest.id);
  for (const [path, content] of parsed.files) {
    await store.write(manifest.id, path, content);
  }

  const plugin: InstalledPlugin = {
    id: manifest.id,
    version: manifest.version,
    tier: manifest.tier,
    enabled: true,
    source,
    installedAt: Date.now(),
    manifest,
  };
  if (download) plugin.download = download;

  return { plugin, warnings: parsed.warnings };
}

/** 从字节直接安装，供本地文件与下载两条路径共用。 */
export async function installFromBytes(
  store: PluginFileStore,
  bytes: Uint8Array,
  source: PluginSource,
  options: { currentAppVersion?: string; download?: string } = {},
): Promise<InstallResult> {
  const parsed = await parsePackage(bytes, options.currentAppVersion);
  return commitPackage(store, parsed, source, options.download);
}
