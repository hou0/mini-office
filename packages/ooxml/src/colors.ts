export interface Rgb {
  r: number;
  g: number;
  b: number;
}

export function normalizeHex(input: string | null | undefined): string | undefined {
  if (!input) return undefined;
  let hex = input.trim().replace(/^#/, '');
  if (/^[0-9a-f]{3}$/i.test(hex)) {
    hex = hex
      .split('')
      .map((char) => char + char)
      .join('');
  }
  if (/^[0-9a-f]{6}$/i.test(hex)) return `#${hex.toUpperCase()}`;
  if (/^[0-9a-f]{8}$/i.test(hex)) return `#${hex.slice(2, 8).toUpperCase()}`;
  return undefined;
}

export function hexToRgb(hex: string): Rgb {
  const normalized = normalizeHex(hex) ?? '#000000';
  return {
    r: parseInt(normalized.slice(1, 3), 16),
    g: parseInt(normalized.slice(3, 5), 16),
    b: parseInt(normalized.slice(5, 7), 16),
  };
}

export function rgbToHex({ r, g, b }: Rgb): string {
  const clamp = (v: number) => Math.max(0, Math.min(255, Math.round(v)));
  return `#${[clamp(r), clamp(g), clamp(b)]
    .map((v) => v.toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase()}`;
}

export function mixColors(a: string, b: string, amount: number): string {
  const ca = hexToRgb(a);
  const cb = hexToRgb(b);
  const t = Math.max(0, Math.min(1, amount));
  return rgbToHex({
    r: ca.r + (cb.r - ca.r) * t,
    g: ca.g + (cb.g - ca.g) * t,
    b: ca.b + (cb.b - ca.b) * t,
  });
}

export function tint(hex: string, amount: number): string {
  return mixColors(hex, '#FFFFFF', amount);
}

export function shade(hex: string, amount: number): string {
  return mixColors(hex, '#000000', amount);
}

function rgbToHsl({ r, g, b }: Rgb): { h: number; s: number; l: number } {
  const rn = r / 255;
  const gn = g / 255;
  const bn = b / 255;
  const max = Math.max(rn, gn, bn);
  const min = Math.min(rn, gn, bn);
  const l = (max + min) / 2;
  if (max === min) return { h: 0, s: 0, l };
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h = 0;
  if (max === rn) h = ((gn - bn) / d + (gn < bn ? 6 : 0)) / 6;
  else if (max === gn) h = ((bn - rn) / d + 2) / 6;
  else h = ((rn - gn) / d + 4) / 6;
  return { h, s, l };
}

function hslToRgb(h: number, s: number, l: number): Rgb {
  if (s === 0) {
    const v = Math.round(l * 255);
    return { r: v, g: v, b: v };
  }
  const hue2rgb = (p: number, q: number, t: number) => {
    let tt = t;
    if (tt < 0) tt += 1;
    if (tt > 1) tt -= 1;
    if (tt < 1 / 6) return p + (q - p) * 6 * tt;
    if (tt < 1 / 2) return q;
    if (tt < 2 / 3) return p + (q - p) * (2 / 3 - tt) * 6;
    return p;
  };
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  return {
    r: Math.round(hue2rgb(p, q, h + 1 / 3) * 255),
    g: Math.round(hue2rgb(p, q, h) * 255),
    b: Math.round(hue2rgb(p, q, h - 1 / 3) * 255),
  };
}

export function adjustLuminance(hex: string, mod: number, off = 0): string {
  const hsl = rgbToHsl(hexToRgb(hex));
  const l = Math.max(0, Math.min(1, hsl.l * mod + off));
  return rgbToHex(hslToRgb(hsl.h, hsl.s, l));
}

export interface ColorSpec {
  hex?: string;
  theme?: string;
  tint?: number;
  shade?: number;
  lumMod?: number;
  lumOff?: number;
  alpha?: number;
  auto?: boolean;
}

export function cssColor(spec: ColorSpec, theme: Record<string, string> = {}): string | undefined {
  let hex = spec.hex;
  if (!hex && spec.theme) {
    hex = theme[spec.theme] ?? theme[spec.theme.replace(/^theme/i, '')] ?? undefined;
  }
  if (!hex) return undefined;
  if (spec.lumMod !== undefined || spec.lumOff !== undefined) {
    hex = adjustLuminance(hex, (spec.lumMod ?? 100000) / 100000, (spec.lumOff ?? 0) / 100000);
  }
  if (spec.tint !== undefined) {
    const amount = Math.max(0, Math.min(1, 1 - spec.tint / 100000));
    hex = tint(hex, amount);
  }
  if (spec.shade !== undefined) {
    const amount = Math.max(0, Math.min(1, 1 - spec.shade / 100000));
    hex = shade(hex, amount);
  }
  if (spec.alpha !== undefined && spec.alpha < 100000) {
    const { r, g, b } = hexToRgb(hex);
    return `rgba(${r}, ${g}, ${b}, ${Math.max(0, spec.alpha / 100000)})`;
  }
  return hex;
}

function parsePercentage(value: string | undefined): number | undefined {
  if (value === undefined || value === '') return undefined;
  const n = Number(value);
  return Number.isFinite(n) ? n : undefined;
}

/** Resolve a DrawingML color choice node (a:srgbClr, a:schemeClr, a:prstClr). */
export function resolveDrawingColor(
  value: unknown,
  theme: Record<string, string> = {},
): ColorSpec | undefined {
  if (!value || typeof value !== 'object') return undefined;
  const node = value as Record<string, unknown>;
  const pick = (key: string) => (node[key] as Record<string, unknown> | undefined) ?? undefined;
  const srgb = pick('a:srgbClr');
  const scheme = pick('a:schemeClr');
  const sys = pick('a:sysClr');
  const prst = pick('a:prstClr');
  const source = srgb ?? scheme ?? sys ?? prst;
  if (!source) return undefined;
  const get = (key: string) => source[`@_${key}`] as string | undefined;
  const child = (key: string) => source[`a:${key}`] as Record<string, unknown> | undefined;
  const spec: ColorSpec = {
    hex: get('val') ? (srgb ? normalizeHex(get('val')) : undefined) : undefined,
    theme: scheme ? get('val') : undefined,
    tint: parsePercentage(child('tint')?.['@_val'] as string | undefined),
    shade: parsePercentage(child('shade')?.['@_val'] as string | undefined),
    lumMod: parsePercentage(child('lumMod')?.['@_val'] as string | undefined),
    lumOff: parsePercentage(child('lumOff')?.['@_val'] as string | undefined),
    alpha: parsePercentage(child('alpha')?.['@_val'] as string | undefined),
  };
  if (sys) spec.hex = normalizeHex(get('lastClr'));
  if (prst) {
    const preset: Record<string, string> = {
      black: '#000000',
      white: '#FFFFFF',
      red: '#FF0000',
      green: '#008000',
      blue: '#0000FF',
      yellow: '#FFFF00',
      gray: '#808080',
      grey: '#808080',
    };
    spec.hex = preset[(get('val') ?? '').toLowerCase()];
  }
  return spec;
}

/** Resolve a WordprocessingML color (`w:color`) including themeColor/themeTint. */
export function resolveWordColor(
  node: unknown,
  theme: Record<string, string> = {},
): string | undefined {
  if (!node || typeof node !== 'object') return undefined;
  const obj = node as Record<string, unknown>;
  const val = obj['@_w:val'] as string | undefined;
  const themeColor = obj['@_w:themeColor'] as string | undefined;
  const themeTint = obj['@_w:themeTint'] as string | undefined;
  const themeShade = obj['@_w:themeShade'] as string | undefined;
  if (val && normalizeHex(val)) return normalizeHex(val);
  if (!themeColor) return undefined;
  const schemeMap: Record<string, string> = {
    text1: 'tx1',
    background1: 'bg1',
    text2: 'tx2',
    background2: 'bg2',
    dark1: 'dk1',
    light1: 'lt1',
    dark2: 'dk2',
    light2: 'lt2',
    accent1: 'accent1',
    accent2: 'accent2',
    accent3: 'accent3',
    accent4: 'accent4',
    accent5: 'accent5',
    accent6: 'accent6',
    hyperlink: 'hlink',
    followedHyperlink: 'folHlink',
  };
  const scheme = schemeMap[themeColor] ?? themeColor;
  const base = theme[scheme];
  if (!base) return undefined;
  const parseHexPercent = (value?: string) => {
    if (!value) return undefined;
    const n = parseInt(value, 16);
    return Number.isFinite(n) ? (n / 255) * 100000 : undefined;
  };
  return cssColor(
    { hex: base, tint: parseHexPercent(themeTint), shade: parseHexPercent(themeShade) },
    theme,
  );
}

export function autoTextColor(backgroundHex: string): string {
  const { r, g, b } = hexToRgb(backgroundHex);
  const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
  return luminance > 0.6 ? '#1D2939' : '#FFFFFF';
}
