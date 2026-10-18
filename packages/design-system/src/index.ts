import { DEFAULT_ACCENT_HEX, type AccentId, type AppSettings } from '@mini-office/shared';

export type ThemeMode = 'light' | 'dark' | 'system';
export type ResolvedTheme = 'light' | 'dark';

export const THEME_STORAGE_KEY = 'mini-office.theme';

export const UI_FONT_STACK =
  '-apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", "Hiragino Sans GB", "Microsoft YaHei", "Noto Sans CJK SC", "Source Han Sans SC", system-ui, sans-serif';

export const DOC_FONT_STACK =
  '"PingFang SC", "Songti SC", "SimSun", "Microsoft YaHei", "Noto Sans CJK SC", "Source Han Sans SC", serif';

/** Common Chinese-friendly font choices offered by editors. */
export const FONT_CHOICES = [
  '默认',
  '宋体',
  '仿宋',
  '黑体',
  '楷体',
  '微软雅黑',
  '等线',
  '思源宋体',
  '思源黑体',
  'PingFang SC',
  'Songti SC',
  'Helvetica Neue',
  'Arial',
  'Times New Roman',
  'Courier New',
] as const;

/** Map CSS-facing Chinese font names to concrete family stacks. */
export function fontStackFor(name?: string): string {
  if (!name || name === '默认') return DOC_FONT_STACK;
  const map: Record<string, string> = {
    宋体: '"Songti SC", "SimSun", "Noto Serif CJK SC", serif',
    仿宋: '"FangSong", "STFangsong", "FangSong_GB2312", serif',
    黑体: '"Heiti SC", "SimHei", "Noto Sans CJK SC", sans-serif',
    楷体: '"Kaiti SC", "KaiTi", "STKaiti", serif',
    微软雅黑: '"Microsoft YaHei", "PingFang SC", sans-serif',
    等线: '"DengXian", "PingFang SC", sans-serif',
    思源宋体: '"Source Han Serif SC", "Noto Serif CJK SC", "Songti SC", serif',
    思源黑体: '"Source Han Sans SC", "Noto Sans CJK SC", "PingFang SC", sans-serif',
  };
  return map[name] ?? `"${name}", ${DOC_FONT_STACK}`;
}

/* -------------------------------------------------------------------------- */
/* Accent palettes                                                            */
/* -------------------------------------------------------------------------- */

export interface AccentShades {
  /** Primary accent used for filled controls and links. */
  accent: string;
  /** Slightly deeper / brighter accent for hover states. */
  hover: string;
  /** Quiet tint used for selected rows, chips and soft buttons. */
  soft: string;
  /** Text/icon color that stays readable on top of `accent`. */
  ink: string;
}

export interface AccentPreset {
  id: Exclude<AccentId, 'custom'>;
  label: string;
  description: string;
  light: AccentShades;
  dark: AccentShades;
}

/**
 * Five calm, low-glare palettes. Each one ships a hand-tuned light and dark
 * variant so the UI never has to guess contrast from a single seed color.
 *
 * `mist` is the default: a deep, slightly desaturated teal that reads as
 * "quiet stationery" rather than a generic Office blue.
 */
export const ACCENT_PRESETS: AccentPreset[] = [
  {
    id: 'mist',
    label: '雾青',
    description: '沉静的松石绿，长时间阅读不刺眼',
    light: { accent: '#2E7D72', hover: '#25665D', soft: '#E6F2EF', ink: '#FFFFFF' },
    dark: { accent: '#6CC9BB', hover: '#8ADBD0', soft: '#1B3A36', ink: '#10221F' },
  },
  {
    id: 'ocean',
    label: '远山蓝',
    description: '更克制的靛蓝，保留熟悉的 Office 感',
    light: { accent: '#2F63E0', hover: '#2552C2', soft: '#E9EFFF', ink: '#FFFFFF' },
    dark: { accent: '#8FB2FF', hover: '#ADC7FF', soft: '#21304F', ink: '#111B2E' },
  },
  {
    id: 'iris',
    label: '紫藤',
    description: '柔和的紫罗兰，适合创意与演示场景',
    light: { accent: '#6B5BD6', hover: '#5949BD', soft: '#F0ECFC', ink: '#FFFFFF' },
    dark: { accent: '#B7A8FF', hover: '#CDC2FF', soft: '#2C2946', ink: '#171428' },
  },
  {
    id: 'pine',
    label: '松针',
    description: '自然深绿，安静而专注',
    light: { accent: '#2F7D4F', hover: '#256640', soft: '#E9F4EE', ink: '#FFFFFF' },
    dark: { accent: '#7BD3A0', hover: '#98E0B7', soft: '#1B3728', ink: '#0F2118' },
  },
  {
    id: 'clay',
    label: '陶土',
    description: '温暖的赤陶色，给文档一点手作感',
    light: { accent: '#B25F4A', hover: '#974C3A', soft: '#F9EDE9', ink: '#FFFFFF' },
    dark: { accent: '#F0A892', hover: '#F6C0AE', soft: '#3B2822', ink: '#2A1812' },
  },
];

export const DEFAULT_ACCENT: AccentId = 'mist';

export function normalizeAccentId(value: string | undefined | null): AccentId {
  if (value === 'custom') return 'custom';
  return ACCENT_PRESETS.some((preset) => preset.id === value) ? (value as AccentId) : DEFAULT_ACCENT;
}

/** Normalize `#abc` / `#aabbcc` to uppercase `#AABBCC`; returns null when invalid. */
export function normalizeAccentHex(value: unknown): string | null {
  if (typeof value !== 'string') return null;
  const input = value.trim();
  if (!/^#?[0-9a-f]{3}([0-9a-f]{3})?$/i.test(input)) return null;
  const hex = input.replace(/^#/, '');
  const expanded =
    hex.length === 3
      ? hex
          .split('')
          .map((char) => `${char}${char}`)
          .join('')
      : hex;
  return `#${expanded.toUpperCase()}`;
}

function hexToRgb(hex: string): [number, number, number] {
  const value = normalizeAccentHex(hex) ?? DEFAULT_ACCENT_HEX;
  return [
    Number.parseInt(value.slice(1, 3), 16),
    Number.parseInt(value.slice(3, 5), 16),
    Number.parseInt(value.slice(5, 7), 16),
  ];
}

function channelToHex(channel: number): string {
  return Math.round(Math.max(0, Math.min(255, channel)))
    .toString(16)
    .padStart(2, '0')
    .toUpperCase();
}

/** Mix `amount` (0–1) of `to` into `from`. */
export function mixHex(from: string, to: string, amount: number): string {
  const [r1, g1, b1] = hexToRgb(from);
  const [r2, g2, b2] = hexToRgb(to);
  const ratio = Math.max(0, Math.min(1, amount));
  return `#${channelToHex(r1 + (r2 - r1) * ratio)}${channelToHex(g1 + (g2 - g1) * ratio)}${channelToHex(
    b1 + (b2 - b1) * ratio,
  )}`;
}

function relativeLuminance(hex: string): number {
  const [r, g, b] = hexToRgb(hex).map((channel) => {
    const value = channel / 255;
    return value <= 0.03928 ? value / 12.92 : ((value + 0.055) / 1.055) ** 2.4;
  });
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

export function contrastRatio(a: string, b: string): number {
  const [light, dark] = [relativeLuminance(a), relativeLuminance(b)].sort((x, y) => y - x);
  return (light + 0.05) / (dark + 0.05);
}

/** Pick the ink with the better contrast against the accent. */
export function readableInk(accent: string): string {
  return contrastRatio(accent, '#FFFFFF') >= contrastRatio(accent, '#1D2939') ? '#FFFFFF' : '#1D2939';
}

function customShades(seed: string | undefined, theme: ResolvedTheme): AccentShades {
  const base = normalizeAccentHex(seed) ?? DEFAULT_ACCENT_HEX;
  if (theme === 'light') {
    return {
      accent: base,
      hover: mixHex(base, '#000000', 0.16),
      soft: mixHex(base, '#FFFFFF', 0.9),
      ink: readableInk(base),
    };
  }
  const accent = mixHex(base, '#FFFFFF', 0.34);
  return {
    accent,
    hover: mixHex(base, '#FFFFFF', 0.46),
    soft: mixHex(base, '#262C36', 0.8),
    ink: readableInk(accent),
  };
}

export function resolveAccentPalette(
  accent: AccentId | string | undefined | null,
  customAccent: string | undefined | null,
  theme: ResolvedTheme,
): AccentShades {
  const id = normalizeAccentId(accent);
  if (id === 'custom') return customShades(customAccent ?? undefined, theme);
  const preset = ACCENT_PRESETS.find((item) => item.id === id) ?? ACCENT_PRESETS[0];
  return preset[theme];
}

export function accentPresetById(accent: AccentId | string | undefined | null): AccentPreset | undefined {
  const id = normalizeAccentId(accent);
  return ACCENT_PRESETS.find((preset) => preset.id === id);
}

export type AppearanceConfig = Pick<AppSettings, 'theme' | 'accent' | 'customAccent'>;

let lastAccent: { accent: AccentId; customAccent: string } = {
  accent: DEFAULT_ACCENT,
  customAccent: DEFAULT_ACCENT_HEX,
};

/**
 * Apply an accent palette to the document root. All values are CSS custom
 * properties, so the rest of the UI stays theme-agnostic.
 */
export function applyAccent(
  accent: AccentId | string | undefined,
  customAccent: string | undefined,
  resolved: ResolvedTheme,
): AccentShades {
  const id = normalizeAccentId(accent);
  const shades = resolveAccentPalette(id, customAccent, resolved);
  lastAccent = {
    accent: id,
    customAccent: normalizeAccentHex(customAccent) ?? DEFAULT_ACCENT_HEX,
  };
  if (typeof document === 'undefined') return shades;
  const root = document.documentElement;
  root.dataset.accent = id;
  root.style.setProperty('--mo-accent', shades.accent);
  root.style.setProperty('--mo-accent-hover', shades.hover);
  root.style.setProperty('--mo-accent-soft', shades.soft);
  root.style.setProperty('--mo-accent-ink', shades.ink);
  return shades;
}

export function resolveTheme(mode: ThemeMode): ResolvedTheme {
  if (mode === 'dark') return 'dark';
  if (mode === 'light') return 'light';
  if (typeof window === 'undefined') return 'light';
  return window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
}

export function applyTheme(mode: ThemeMode): ResolvedTheme {
  if (typeof document === 'undefined') return 'light';
  const resolved = resolveTheme(mode);
  document.documentElement.dataset.theme = resolved;
  document.documentElement.style.colorScheme = resolved;
  // Keep the accent in sync when the resolved theme changes (system mode).
  applyAccent(lastAccent.accent, lastAccent.customAccent, resolved);
  return resolved;
}

/** Apply theme + accent together; this is the entry point used by the app. */
export function applyAppearance(config: AppearanceConfig): ResolvedTheme {
  lastAccent = {
    accent: normalizeAccentId(config.accent),
    customAccent: normalizeAccentHex(config.customAccent) ?? DEFAULT_ACCENT_HEX,
  };
  return applyTheme(config.theme);
}

let mediaCleanup: (() => void) | null = null;

export function watchSystemTheme(mode: ThemeMode): () => void {
  if (typeof window === 'undefined') return () => {};
  mediaCleanup?.();
  const media = window.matchMedia('(prefers-color-scheme: dark)');
  const onChange = () => {
    if (mode === 'system') applyTheme('system');
  };
  media.addEventListener('change', onChange);
  mediaCleanup = () => media.removeEventListener('change', onChange);
  return mediaCleanup;
}

export const MOTION = {
  fast: 'var(--mo-motion-fast)',
  base: 'var(--mo-motion-base)',
} as const;
