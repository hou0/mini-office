/**
 * Mini Office 插件公开契约。
 *
 * 这是插件唯一允许依赖的包。宿主内部包（document-model、ooxml、design-system 等）
 * 都**不得**被插件直接引用，否则内核重构会击穿插件。
 *
 * 版本策略：本包的 minor 版本号即插件 API 版本；破坏性变更必须升 major，
 * 并保证旧 major 至少保留两个 minor 版本的兼容期。
 */

/** 插件 API 版本。写入 manifest.engines.api。 */
export const PLUGIN_API_VERSION = '1.0.0';

/** 插件包扩展名（本质是 zip）。 */
export const PLUGIN_PACKAGE_EXT = 'moext';

/** 清单文件名，必须位于包根目录。 */
export const PLUGIN_MANIFEST_FILE = 'manifest.json';

/** 单包解压后大小上限，防止压缩炸弹。 */
export const PLUGIN_MAX_UNCOMPRESSED_BYTES = 8 * 1024 * 1024;

/** 单包条目数上限。 */
export const PLUGIN_MAX_ENTRIES = 256;

/* -------------------------------------------------------------------------- */
/* 信任等级                                                                    */
/* -------------------------------------------------------------------------- */

/**
 * - `declarative`：纯数据，不执行代码。主题、模板、快捷键。
 * - `sandboxed`：JS 跑在 Worker / iframe，仅通过 RPC 访问文档。
 * - `trusted`：可调用原生能力，需签名与显式授权。
 */
export type PluginTier = 'declarative' | 'sandboxed' | 'trusted';

export const PLUGIN_TIER_LABEL: Record<PluginTier, string> = {
  declarative: '声明式',
  sandboxed: '沙箱',
  trusted: '受信',
};

/* -------------------------------------------------------------------------- */
/* 权限                                                                        */
/* -------------------------------------------------------------------------- */

export type PluginPermission =
  | 'document:read'
  | 'document:write'
  | 'fs:read'
  | 'fs:write'
  | 'network'
  | 'clipboard'
  | 'print'
  | 'shell:exec';

export const PLUGIN_PERMISSION_LABEL: Record<PluginPermission, string> = {
  'document:read': '读取当前文档内容',
  'document:write': '修改当前文档',
  'fs:read': '读取文档以外的文件',
  'fs:write': '写入文档以外的文件',
  network: '访问网络',
  clipboard: '读写剪贴板',
  print: '触发打印',
  'shell:exec': '调用外部程序',
};

/* -------------------------------------------------------------------------- */
/* 主题                                                                        */
/* -------------------------------------------------------------------------- */

/** 外壳渲染变体。主题只能从枚举中选择，不能注入 DOM。 */
export interface ChromeVariant {
  /** 工具栏形态。`classic-two-row` 用于 Office 2003 风格。 */
  toolbar?: 'modern' | 'classic-two-row' | 'ribbon';
  /** 控件形状。`beveled` 为凸起/凹陷的经典边框。 */
  controlShape?: 'flat' | 'beveled' | 'rounded';
  /** 控件密度。 */
  density?: 'compact' | 'default' | 'comfortable';
  /** 标题栏渲染方式。 */
  titlebar?: 'flat' | 'gradient';
}

/** 主题令牌集合，键必须命中宿主白名单（`--mo-*`，且不含 `--mo-doc-*`）。 */
export type ThemeTokens = Record<string, string>;

export interface ThemeContribution {
  /** 主题 id，在同一插件内唯一。 */
  id: string;
  label: string;
  description?: string;
  /** 预览图，包内相对路径，如 `preview.png`。 */
  preview?: string;
  /** 浅色令牌；必填。 */
  tokens: ThemeTokens;
  /** 深色令牌；可缺省，缺省时回退到浅色。 */
  darkTokens?: ThemeTokens;
  chrome?: ChromeVariant;
  /** 允许主题提供的默认强调色 id；用户仍可覆盖。 */
  defaultAccent?: string;
  /**
   * 文档纸张底色（如护眼米黄、复古纸白）。
   * 仅影响画布背景，**不影响字体度量**，因此不会改变分页与打印结果。
   */
  docPaperTint?: string;
}

/* -------------------------------------------------------------------------- */
/* 快捷键                                                                      */
/* -------------------------------------------------------------------------- */

const MODIFIER_ALIASES: Record<string, 'mod' | 'shift' | 'alt'> = {
  mod: 'mod',
  cmd: 'mod',
  command: 'mod',
  ctrl: 'mod',
  control: 'mod',
  shift: 'shift',
  alt: 'alt',
  option: 'alt',
  opt: 'alt',
};

const MODIFIER_ORDER = ['mod', 'shift', 'alt'] as const;

function normalizeShortcutKey(raw: string): string | null {
  const key = raw.trim().toLowerCase();
  if (!key) return null;
  if (key === 'space') return 'space';
  if (key.length === 1) return key;
  if (/^f([1-9]|1[0-2])$/.test(key)) return key;
  return null;
}

/**
 * 归一化快捷键写法：`Mod+Shift+K` → `mod+shift+k`。
 *
 * 放在契约包里而不是宿主里，是因为**清单校验与运行时必须用同一套规则**——
 * 两边各写一份正则，迟早会出现"校验通过但注册不上"或反之的偏差。
 *
 * 规则：必须带修饰键（`Mod`/`Ctrl`/`Cmd`/`Alt`/`Shift` 之一），主键为单字符或 F1–F12。
 * 没有修饰键的组合会抢占正常输入，因此视为非法。
 */
export function normalizeShortcut(input: string): string | null {
  const parts = input.split('+').map((part) => part.trim()).filter(Boolean);
  if (parts.length === 0) return null;

  const modifiers = new Set<'mod' | 'shift' | 'alt'>();
  let key: string | null = null;
  for (const part of parts) {
    const modifier = MODIFIER_ALIASES[part.toLowerCase()];
    if (modifier) {
      modifiers.add(modifier);
      continue;
    }
    if (key !== null) return null;
    key = normalizeShortcutKey(part);
    if (key === null) return null;
  }

  if (key === null || modifiers.size === 0) return null;
  return [...MODIFIER_ORDER.filter((item) => modifiers.has(item)), key].join('+');
}

/* -------------------------------------------------------------------------- */
/* 贡献点                                                                      */
/* -------------------------------------------------------------------------- */

/**
 * 插件宏能作用的文档类型。
 *
 * 这里不复用 `@mini-office/shared` 的 `OfficeKind`，因为本包必须保持零依赖——
 * 插件作者只装一个 `@mini-office/plugin-api` 就能写插件。
 */
export type PluginDocumentKind = 'docx' | 'xlsx' | 'pptx';

/**
 * 宏的作用范围。
 *
 * - `body`（默认）：**跳过居中与右对齐段落**。标题、落款、印章行通常属于这两类，
 *   一视同仁地套正文格式会把标题拉成两端对齐、给落款加首行缩进而换行。
 * - `all`：字面意义上的全部段落，不做任何跳过。
 */
export type MacroScope = 'all' | 'body';

/** 表格边框预设。`three-line` 即中文公文常用的三线表。 */
export type TableBorderPreset = 'three-line' | 'grid' | 'none';

/**
 * 宏步骤：一条命令由若干步骤组成，宿主按顺序执行。
 *
 * 设计原则：**只有白名单里的操作，没有"执行任意脚本"的入口**。
 * 这样声明式插件仍然不执行代码，宏只是被宿主解释的数据。
 */
export type MacroStep =
  | {
      op: 'word.page.margins';
      /** 单位 CSS px；96px ≈ 1 英寸。缺省字段保持不变。 */
      top?: number;
      right?: number;
      bottom?: number;
      left?: number;
    }
  | { op: 'word.page.orientation'; orientation: 'portrait' | 'landscape' }
  | {
      op: 'word.paragraph.format';
      scope?: MacroScope;
      alignment?: 'left' | 'center' | 'right' | 'justify';
      /** 倍数行距，如 1.5。 */
      lineHeight?: number;
      firstLineIndentPx?: number;
      spaceBeforePx?: number;
      spaceAfterPx?: number;
      indentLeftPx?: number;
      indentRightPx?: number;
    }
  | {
      op: 'word.font.apply';
      scope?: MacroScope;
      fontFamily?: string;
      fontSizePt?: number;
      bold?: boolean;
      italic?: boolean;
      /** 十六进制颜色，如 `#C00000`。 */
      color?: string;
    }
  | { op: 'word.table.borders'; preset: TableBorderPreset }
  | { op: 'word.defaultFont.set'; fontFamily: string; fontSizePt?: number }
  | { op: 'sheet.freezeHeader'; sheet?: 'active' | 'all' }
  | {
      op: 'sheet.headerStyle';
      sheet?: 'active' | 'all';
      fontFamily?: string;
      fontSizePt?: number;
      bold?: boolean;
      /** 表头填充色，十六进制。 */
      fill?: string;
      color?: string;
    };

/** 全部受支持的宏操作名，供校验、文档与示例共用。 */
export const MACRO_OPS = [
  'word.page.margins',
  'word.page.orientation',
  'word.paragraph.format',
  'word.font.apply',
  'word.table.borders',
  'word.defaultFont.set',
  'sheet.freezeHeader',
  'sheet.headerStyle',
] as const;

export type MacroOp = (typeof MACRO_OPS)[number];

/**
 * 命令贡献点。
 *
 * v1 的命令**只有宏**一种形态：不跑代码，由宿主按 `steps` 顺序改文档。
 * 需要真正执行逻辑的功能（校对、格式转换）属于 T2 沙箱插件，不走这里。
 */
export interface CommandContribution {
  id: string;
  title: string;
  description?: string;
  /** 命令面板搜索关键词，空格分隔。 */
  keywords?: string;
  /** 形如 `Mod+Shift+P`；`Mod` 在 macOS 解析为 Cmd。 */
  shortcut?: string;
  /** 至少一步。所有步骤必须属于同一种文档类型。 */
  steps: MacroStep[];
}

export interface TemplateContribution {
  id: string;
  label: string;
  /** 目标格式。 */
  kind: PluginDocumentKind;
  description?: string;
  /** 模板文件，包内相对路径，如 `templates/letter.docx`。 */
  file: string;
}

export interface PluginContributions {
  themes?: ThemeContribution[];
  commands?: CommandContribution[];
  templates?: TemplateContribution[];
}

/**
 * v1 明确支持的贡献点键。
 *
 * 清单里出现其他键（如 `panels`、`menus`）会被**拒绝安装**而不是静默忽略——
 * 静默忽略会让插件作者以为生效了，排查成本远高于直接报错。
 */
export const SUPPORTED_CONTRIBUTION_KEYS = ['themes', 'commands', 'templates'] as const;


/* -------------------------------------------------------------------------- */
/* 清单                                                                        */
/* -------------------------------------------------------------------------- */

export interface PluginManifest {
  /** 反向域名式 id，如 `mini-office.theme-office-2003`。 */
  id: string;
  name: string;
  version: string;
  publisher: string;
  description?: string;
  homepage?: string;
  license?: string;
  tier: PluginTier;
  engines: {
    /** 兼容的 Mini Office 版本范围，如 `>=0.1.5 <0.5.0`。 */
    miniOffice: string;
    /** 兼容的插件 API 范围，如 `^1.0.0`。 */
    api: string;
  };
  permissions: PluginPermission[];
  contributes: PluginContributions;
  /** 图标，包内相对路径。 */
  icon?: string;
}

/* -------------------------------------------------------------------------- */
/* 插件索引（GitHub 仓库托管）                                                  */
/* -------------------------------------------------------------------------- */

export interface PluginIndexEntry {
  id: string;
  name: string;
  version: string;
  publisher: string;
  description?: string;
  tier: PluginTier;
  permissions: PluginPermission[];
  /** 主题类插件在列表里展示的预览图 URL。 */
  preview?: string;
  /** 包下载地址，支持 GitHub Release 资源或仓库原始文件。 */
  download: string;
  /** 包内容的 SHA-256（小写十六进制），安装后校验。 */
  sha256: string;
  /** 字节数，用于展示与下载进度。 */
  size?: number;
  tags?: string[];
  updatedAt?: string;
}

export interface PluginIndex {
  schemaVersion: 1;
  updatedAt?: string;
  plugins: PluginIndexEntry[];
}

/* -------------------------------------------------------------------------- */
/* 安装记录                                                                    */
/* -------------------------------------------------------------------------- */

export type PluginSource = 'bundled' | 'local' | 'github';

export interface InstalledPlugin {
  id: string;
  version: string;
  tier: PluginTier;
  enabled: boolean;
  source: PluginSource;
  installedAt: number;
  /** 安装时的完整清单快照，避免包被替换后 UI 与运行态不一致。 */
  manifest: PluginManifest;
  /** 来自 GitHub 索引时记录的下载地址，用于检查更新。 */
  download?: string;
}

/* -------------------------------------------------------------------------- */
/* 校验结果                                                                    */
/* -------------------------------------------------------------------------- */

export interface ValidationIssue {
  path: string;
  message: string;
}

export interface ValidationResult {
  ok: boolean;
  issues: ValidationIssue[];
  warnings: ValidationIssue[];
}

/* -------------------------------------------------------------------------- */
/* 宿主约定                                                                    */
/* -------------------------------------------------------------------------- */

/** 主题包注入到 <html> 上的属性名，供 CSS 选择器定位。 */
export const THEME_PACK_ATTR = 'data-theme-pack';

/** 外壳变体注入到 <html> 上的属性前缀。 */
export const CHROME_ATTR_PREFIX = 'data-chrome-';

/** 内置主题的保留 id。 */
export const BUILTIN_THEME_PACK = 'builtin';

/** 设置项中的插件仓库坐标，供 UI 展示与修改。 */
export interface PluginRepositoryRef {
  owner: string;
  repo: string;
  /** 索引文件所在分支。 */
  branch: string;
  /** 索引文件路径。 */
  indexPath: string;
}

export const DEFAULT_PLUGIN_REPOSITORY: PluginRepositoryRef = {
  owner: 'lq0910',
  repo: 'mini-office-plugins',
  branch: 'main',
  indexPath: 'index.json',
};
