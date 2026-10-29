/**
 * 主题引擎：令牌校验与主题包应用。
 *
 * 设计要点：
 * 1. 主题只能改**外壳令牌**，`--mo-doc-*` 一律拒绝。文档画布的字体度量决定分页，
 *    一旦被主题改写，同一份 DOCX 的页数与打印结果就会变化。
 * 2. 主题通过注入 `<style>` 生效，而非内联样式。这样 `applyAccent` 写入的内联
 *    `--mo-accent*` 天然优先级更高，强调色始终由用户设置决定。
 * 3. 校验失败即拒绝激活，不做静默兜底。
 */
import {
  CHROME_ATTR_PREFIX,
  THEME_PACK_ATTR,
  type ChromeVariant,
  type ThemeContribution,
  type ThemeTokens,
  type ValidationIssue,
} from '@mini-office/plugin-api';

export type ResolvedTheme = 'light' | 'dark';

/**
 * 允许主题覆盖的令牌白名单，与 `packages/design-system/src/tokens.css` 保持一致。
 * 新增令牌时需同时更新此处，否则主题写了也不会生效（校验阶段会报未知令牌）。
 */
export const THEMEABLE_TOKENS: readonly string[] = [
  '--mo-bg',
  '--mo-surface',
  '--mo-surface-2',
  '--mo-surface-3',
  '--mo-canvas',
  '--mo-accent',
  '--mo-accent-hover',
  '--mo-accent-soft',
  '--mo-accent-ink',
  '--mo-text-primary',
  '--mo-text-muted',
  '--mo-text-faint',
  '--mo-line',
  '--mo-line-strong',
  '--mo-danger',
  '--mo-danger-soft',
  '--mo-success',
  '--mo-success-soft',
  '--mo-warning',
  '--mo-warning-soft',
  '--mo-radius-md',
  '--mo-radius-lg',
  '--mo-radius-xl',
  '--mo-shadow-panel',
  '--mo-shadow-pop',
  '--mo-shadow-page',
  '--mo-motion-fast',
  '--mo-motion-base',
  '--mo-font-ui',
  '--mo-font-doc',
  '--mo-font-mono',
];

/** 缺失即视为不完整主题的令牌。 */
const REQUIRED_TOKENS: readonly string[] = [
  '--mo-bg',
  '--mo-surface',
  '--mo-surface-2',
  '--mo-text-primary',
  '--mo-text-muted',
  '--mo-line',
];

/** 文档画布保留前缀，任何主题都不得触碰。 */
const FORBIDDEN_PREFIX = '--mo-doc-';

const CHROME_ENUMS: Record<keyof ChromeVariant, readonly string[]> = {
  toolbar: ['modern', 'classic-two-row', 'ribbon'],
  controlShape: ['flat', 'beveled', 'rounded'],
  density: ['compact', 'default', 'comfortable'],
  titlebar: ['flat', 'gradient'],
};

/* -------------------------------------------------------------------------- */
/* 颜色与对比度                                                                */
/* -------------------------------------------------------------------------- */

export interface Rgb {
  r: number;
  g: number;
  b: number;
}

/** 解析 #RGB / #RRGGBB / #RRGGBBAA；非颜色值返回 null。 */
export function parseHexColor(value: string): Rgb | null {
  const text = value.trim();
  const short = /^#([0-9a-f]{3})$/i.exec(text);
  if (short) {
    const [r, g, b] = short[1].split('');
    return {
      r: parseInt(r + r, 16),
      g: parseInt(g + g, 16),
      b: parseInt(b + b, 16),
    };
  }
  const full = /^#([0-9a-f]{6})([0-9a-f]{2})?$/i.exec(text);
  if (!full) return null;
  return {
    r: parseInt(full[1].slice(0, 2), 16),
    g: parseInt(full[1].slice(2, 4), 16),
    b: parseInt(full[1].slice(4, 6), 16),
  };
}

function channelLuminance(channel: number): number {
  const normalized = channel / 255;
  return normalized <= 0.03928 ? normalized / 12.92 : ((normalized + 0.055) / 1.055) ** 2.4;
}

export function relativeLuminance(color: Rgb): number {
  return 0.2126 * channelLuminance(color.r) + 0.7152 * channelLuminance(color.g) + 0.0722 * channelLuminance(color.b);
}

/** WCAG 2.1 对比度，范围 1–21。 */
export function contrastRatio(a: Rgb, b: Rgb): number {
  const la = relativeLuminance(a);
  const lb = relativeLuminance(b);
  const lighter = Math.max(la, lb);
  const darker = Math.min(la, lb);
  return (lighter + 0.05) / (darker + 0.05);
}

/* -------------------------------------------------------------------------- */
/* 校验                                                                        */
/* -------------------------------------------------------------------------- */

export interface ThemeValidationResult {
  ok: boolean;
  issues: ValidationIssue[];
  warnings: ValidationIssue[];
}

function validateTokenBlock(
  tokens: unknown,
  scope: string,
  issues: ValidationIssue[],
  warnings: ValidationIssue[],
): ThemeTokens {
  if (tokens === undefined) return {};
  if (typeof tokens !== 'object' || tokens === null || Array.isArray(tokens)) {
    issues.push({ path: scope, message: '令牌必须是键值对象' });
    return {};
  }
  const clean: ThemeTokens = {};
  for (const [key, value] of Object.entries(tokens as Record<string, unknown>)) {
    if (key.startsWith(FORBIDDEN_PREFIX)) {
      issues.push({ path: `${scope}.${key}`, message: '主题无权修改文档画布令牌（--mo-doc-*），会影响分页与打印' });
      continue;
    }
    if (!THEMEABLE_TOKENS.includes(key)) {
      issues.push({ path: `${scope}.${key}`, message: '未知或不可主题化的令牌' });
      continue;
    }
    if (typeof value !== 'string' || !value.trim()) {
      issues.push({ path: `${scope}.${key}`, message: '令牌值必须是非空字符串' });
      continue;
    }
    clean[key] = value.trim();
  }
  return clean;
}

function checkContrast(
  tokens: ThemeTokens,
  scope: string,
  issues: ValidationIssue[],
  warnings: ValidationIssue[],
): void {
  const surface = parseHexColor(tokens['--mo-surface'] ?? '');
  if (!surface) return;

  const pairs: Array<{ fg: string; min: number; label: string }> = [
    { fg: '--mo-text-primary', min: 4.5, label: '正文文字' },
    { fg: '--mo-text-muted', min: 3, label: '次要文字' },
  ];

  for (const pair of pairs) {
    const raw = tokens[pair.fg];
    if (!raw) continue;
    const fg = parseHexColor(raw);
    if (!fg) continue;
    const ratio = contrastRatio(fg, surface);
    if (ratio < pair.min) {
      issues.push({
        path: `${scope}.${pair.fg}`,
        message: `${pair.label}与背景对比度仅 ${ratio.toFixed(2)}:1，低于 ${pair.min}:1，装上会看不清`,
      });
    } else if (ratio < 4.5 && pair.min < 4.5) {
      warnings.push({
        path: `${scope}.${pair.fg}`,
        message: `${pair.label}对比度 ${ratio.toFixed(2)}:1，建议提升到 4.5:1`,
      });
    }
  }
}

function validateChrome(chrome: unknown, issues: ValidationIssue[]): ChromeVariant | undefined {
  if (chrome === undefined) return undefined;
  if (typeof chrome !== 'object' || chrome === null || Array.isArray(chrome)) {
    issues.push({ path: 'chrome', message: 'chrome 必须是对象' });
    return undefined;
  }
  const clean: ChromeVariant = {};
  for (const [key, value] of Object.entries(chrome as Record<string, unknown>)) {
    const allowed = CHROME_ENUMS[key as keyof ChromeVariant];
    if (!allowed) {
      issues.push({ path: `chrome.${key}`, message: '未知的外壳变体' });
      continue;
    }
    if (typeof value !== 'string' || !allowed.includes(value)) {
      issues.push({ path: `chrome.${key}`, message: `取值必须是 ${allowed.join(' / ')}` });
      continue;
    }
    (clean as Record<string, string>)[key] = value;
  }
  return clean;
}

/** 校验单个主题贡献。 */
export function validateTheme(theme: unknown, index: number): ThemeValidationResult & { theme?: ThemeContribution } {
  const issues: ValidationIssue[] = [];
  const warnings: ValidationIssue[] = [];
  const scope = `themes[${index}]`;

  if (typeof theme !== 'object' || theme === null || Array.isArray(theme)) {
    return { ok: false, issues: [{ path: scope, message: '主题必须是对象' }], warnings };
  }
  const source = theme as Record<string, unknown>;
  const id = typeof source.id === 'string' ? source.id.trim() : '';
  const label = typeof source.label === 'string' ? source.label.trim() : '';
  if (!id) issues.push({ path: `${scope}.id`, message: '缺少主题 id' });
  if (!label) issues.push({ path: `${scope}.label`, message: '缺少主题名称' });

  const tokens = validateTokenBlock(source.tokens, `${scope}.tokens`, issues, warnings);
  const darkTokens = validateTokenBlock(source.darkTokens, `${scope}.darkTokens`, issues, warnings);

  for (const required of REQUIRED_TOKENS) {
    if (!(required in tokens)) {
      issues.push({ path: `${scope}.tokens.${required}`, message: '缺少必填令牌' });
    }
  }

  checkContrast(tokens, `${scope}.tokens`, issues, warnings);
  if (Object.keys(darkTokens).length > 0) {
    checkContrast({ ...tokens, ...darkTokens }, `${scope}.darkTokens`, issues, warnings);
  }

  const chrome = validateChrome(source.chrome, issues);

  if (typeof source.docPaperTint === 'string' && !parseHexColor(source.docPaperTint)) {
    issues.push({ path: `${scope}.docPaperTint`, message: '纸张底色需为十六进制颜色' });
  }

  if (issues.length > 0) return { ok: false, issues, warnings };

  const result: ThemeContribution = { id, label, tokens };
  if (typeof source.description === 'string') result.description = source.description;
  if (typeof source.preview === 'string') result.preview = source.preview;
  if (Object.keys(darkTokens).length > 0) result.darkTokens = darkTokens;
  if (chrome && Object.keys(chrome).length > 0) result.chrome = chrome;
  if (typeof source.defaultAccent === 'string') result.defaultAccent = source.defaultAccent;
  if (typeof source.docPaperTint === 'string') result.docPaperTint = source.docPaperTint;

  return { ok: true, issues, warnings, theme: result };
}

/* -------------------------------------------------------------------------- */
/* 应用                                                                        */
/* -------------------------------------------------------------------------- */

const STYLE_ELEMENT_ID = 'mo-theme-pack';

/** 主题包的全局唯一键：`插件id#主题id`。 */
export function themePackKey(pluginId: string, themeId: string): string {
  return `${pluginId}#${themeId}`;
}

export function splitThemePackKey(key: string): { pluginId: string; themeId: string } | null {
  const index = key.indexOf('#');
  if (index <= 0) return null;
  return { pluginId: key.slice(0, index), themeId: key.slice(index + 1) };
}

function declarationBlock(tokens: ThemeTokens): string {
  return Object.entries(tokens)
    .map(([key, value]) => `  ${key}: ${value};`)
    .join('\n');
}

export interface ActiveTheme {
  /** 主题包唯一键，`builtin` 表示内置外观。 */
  key: string;
  theme?: ThemeContribution;
}

/**
 * 应用主题包。传 `null` 或不存在的包时回退到内置外观。
 *
 * 注意：不触碰 `--mo-accent*` 的最终值——强调色由 `applyAccent` 以内联样式
 * 覆盖，主题里的 accent 只作为初始值存在。
 */
export function applyThemePack(active: ActiveTheme | null, resolved: ResolvedTheme): void {
  if (typeof document === 'undefined') return;
  const root = document.documentElement;

  const existing = document.getElementById(STYLE_ELEMENT_ID);
  if (existing) existing.remove();

  if (!active?.theme) {
    delete root.dataset.themePack;
    for (const key of Object.keys(CHROME_ENUMS)) {
      root.removeAttribute(`${CHROME_ATTR_PREFIX}${kebab(key)}`);
    }
    return;
  }

  const theme = active.theme;
  const light = theme.tokens;
  const dark = theme.darkTokens && Object.keys(theme.darkTokens).length > 0 ? theme.darkTokens : theme.tokens;
  const scope = `[${THEME_PACK_ATTR}="${active.key}"]`;

  const blocks: string[] = [
    `${scope}[data-theme="light"] {\n${declarationBlock(light)}\n}`,
    `${scope}[data-theme="dark"] {\n${declarationBlock(dark)}\n}`,
  ];
  if (theme.docPaperTint) {
    blocks.push(`${scope} {\n  --mo-doc-paper: ${theme.docPaperTint};\n}`);
  }

  const style = document.createElement('style');
  style.id = STYLE_ELEMENT_ID;
  style.textContent = blocks.join('\n');
  document.head.appendChild(style);

  root.dataset.themePack = active.key;

  const chrome = theme.chrome ?? {};
  for (const key of Object.keys(CHROME_ENUMS) as Array<keyof ChromeVariant>) {
    const value = chrome[key];
    if (value) root.setAttribute(`${CHROME_ATTR_PREFIX}${kebab(key)}`, value);
    else root.removeAttribute(`${CHROME_ATTR_PREFIX}${kebab(key)}`);
  }

  void resolved;
}

function kebab(value: string): string {
  return value.replace(/[A-Z]/g, (char) => `-${char.toLowerCase()}`);
}
