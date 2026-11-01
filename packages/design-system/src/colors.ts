/**
 * 功能：文档纸张相关的颜色换算——把文档里写死的颜色映射到当前主题可读的等价色；
 * 时间：2026-10-31；作者：lq。
 *
 * 背景：Word / Excel 文档里的颜色来自 OOXML（多为纯黑字、白底、浅色底纹、黑框线）。
 * 深色模式下如果原样渲染，黑字会消失在深色纸上、浅底纹会变成刺眼的白块。
 * 这里提供一组纯函数，把「文档色」换算成「纸张色」：
 *   - 中性深色文字 → 纸张文字色（--mo-paper-ink）
 *   - 高饱和但偏暗的文字 → 同色相提亮，保留强调色语义
 *   - 近白底纹 → 同色相压暗，保留单元格底色语义（Excel 里颜色就是信息）
 *   - 黑色框线 → --mo-table-line（浅色/打印时是纯黑，深色时自动提亮）
 * 打印与导出时主题变量被强制重置为白纸，因此这些换算在输出上等价于原色。
 */

export interface RgbColor {
  r: number;
  g: number;
  b: number;
  a: number;
}

/** 功能：解析 #rgb / #rrggbb / #rrggbbaa / rgb() / rgba()；无法识别时返回 null；时间：2026-10-31；作者：lq。 */
export function parseColor(input: string | undefined | null): RgbColor | null {
  if (!input) return null;
  const value = input.trim().toLowerCase();
  if (!value) return null;
  if (value.startsWith('#')) {
    const hex = value.slice(1);
    if (!/^[0-9a-f]+$/.test(hex)) return null;
    if (hex.length === 3 || hex.length === 4) {
      const expand = (char: string) => Number.parseInt(char + char, 16);
      return {
        r: expand(hex[0]),
        g: expand(hex[1]),
        b: expand(hex[2]),
        a: hex.length === 4 ? expand(hex[3]) / 255 : 1,
      };
    }
    if (hex.length === 6 || hex.length === 8) {
      return {
        r: Number.parseInt(hex.slice(0, 2), 16),
        g: Number.parseInt(hex.slice(2, 4), 16),
        b: Number.parseInt(hex.slice(4, 6), 16),
        a: hex.length === 8 ? Number.parseInt(hex.slice(6, 8), 16) / 255 : 1,
      };
    }
    return null;
  }
  const functional = /^rgba?\(([^)]+)\)$/.exec(value);
  if (!functional) return null;
  const parts = functional[1].split(/[,/\s]+/).filter(Boolean);
  if (parts.length < 3) return null;
  const channel = (raw: string) =>
    raw.endsWith('%') ? Math.round((Number.parseFloat(raw) / 100) * 255) : Number.parseFloat(raw);
  const [r, g, b, alpha] = parts;
  const parsed = { r: channel(r), g: channel(g), b: channel(b), a: alpha === undefined ? 1 : Number.parseFloat(alpha) };
  if ([parsed.r, parsed.g, parsed.b].some((item) => !Number.isFinite(item))) return null;
  return {
    r: clamp(parsed.r, 0, 255),
    g: clamp(parsed.g, 0, 255),
    b: clamp(parsed.b, 0, 255),
    a: Number.isFinite(parsed.a) ? clamp(parsed.a, 0, 1) : 1,
  };
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value));
}

/** 功能：0–255 的感知亮度（Rec.601），用于判断文字在深色纸上是否还看得见；时间：2026-10-31；作者：lq。 */
function luma(rgb: RgbColor): number {
  return 0.299 * rgb.r + 0.587 * rgb.g + 0.114 * rgb.b;
}

function chroma(rgb: RgbColor): number {
  return Math.max(rgb.r, rgb.g, rgb.b) - Math.min(rgb.r, rgb.g, rgb.b);
}

/**
 * 功能：判断一个颜色是否为「深到在深色纸张上不可读的中性色」；时间：2026-10-31；作者：lq。
 *
 * 只认接近黑/深灰的低彩度颜色：深红、深蓝这类强调色本身有足够辨识度，
 * 一律改成浅灰会把文档里的红色标题、蓝色链接全部抹平。
 */
export function isDarkInk(color: string | undefined): boolean {
  const rgb = parseColor(color);
  if (!rgb) return false;
  return chroma(rgb) <= 40 && luma(rgb) < 110;
}

/**
 * 功能：判断是否为「浅色底纹」；时间：2026-10-31；作者：lq。
 *
 * 用亮度而不是最小通道判定：Office 标准浅底纹（#FFF2CC / #E2EFDA / #FCE4D6 / #EBF1DE）
 * 的最小通道只有 204~235，用「三个通道都不低于 232」会漏掉大半，
 * 深色纸张上就会留下刺眼的白块。
 */
export function isLightShade(color: string | undefined): boolean {
  const rgb = parseColor(color);
  if (!rgb) return false;
  return luma(rgb) >= 200;
}

/** 功能：段落底纹——近白色走主题变量，深色纸上换成深色底，打印时回到白底；时间：2026-10-31；作者：lq。 */
export function shadeColor(color: string | undefined): string | undefined {
  if (!color) return undefined;
  return isLightShade(color) ? 'var(--mo-shade-light, #ffffff)' : color;
}

interface HslColor {
  h: number;
  s: number;
  l: number;
  a: number;
}

function toHsl(rgb: RgbColor): HslColor {
  const r = rgb.r / 255;
  const g = rgb.g / 255;
  const b = rgb.b / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  const delta = max - min;
  if (delta === 0) return { h: 0, s: 0, l, a: rgb.a };
  const s = delta / (1 - Math.abs(2 * l - 1));
  let h: number;
  if (max === r) h = ((g - b) / delta) % 6;
  else if (max === g) h = (b - r) / delta + 2;
  else h = (r - g) / delta + 4;
  return { h: (h * 60 + 360) % 360, s, l, a: rgb.a };
}

function hslCss(hsl: HslColor): string {
  const round = (value: number) => Math.round(value * 10) / 10;
  const alpha = hsl.a < 1 ? ` / ${Math.round(hsl.a * 1000) / 1000}` : '';
  return `hsl(${round(hsl.h)} ${round(hsl.s * 100)}% ${round(hsl.l * 100)}%${alpha})`;
}

/*
 * 功能：换算结果缓存；时间：2026-10-31；作者：lq。
 *
 * 一张工作表可见区域可能有上千个单元格，每个都要走一次解析与色相换算。
 * 文档里的颜色取值高度重复（黑字、白底、几种浅色底纹），按颜色字符串缓存后
 * 单帧只做哈希查找，滚动时不会有额外开销。
 */
const memo = {
  ink: new Map<string, string>(),
  fill: new Map<string, string>(),
};
const MEMO_LIMIT = 512;

function memoGet(kind: 'ink' | 'fill', color: string): string | undefined {
  return memo[kind].get(color);
}

function memoSet(kind: 'ink' | 'fill', color: string, value: string): void {
  if (memo[kind].size >= MEMO_LIMIT) memo[kind].clear();
  memo[kind].set(color, value);
}

/**
 * 功能：把任意颜色换算成深色纸张上的「文字色」；时间：2026-10-31；作者：lq。
 *
 * 1. 未指定颜色 → 纸张默认文字色；
 * 2. 中性深色 → 纸张文字色（最稳，不会出现偏色）；
 * 3. 已经够亮的颜色 → 原样保留；
 * 4. 高饱和但偏暗（如深红 C00000、深蓝 1F4E79）→ 同色相提亮到可读亮度。
 */
export function paperInk(color: string | undefined, dark: boolean): string {
  const fallback = 'var(--mo-paper-ink, #1d2939)';
  if (!color) return fallback;
  if (!dark) return color;
  const cached = memoGet('ink', color);
  if (cached) return cached;
  const rgb = parseColor(color);
  let result = color;
  if (rgb) {
    const isNeutralDark = chroma(rgb) <= 40 && luma(rgb) < 110;
    if (isNeutralDark) {
      result = fallback;
    } else {
      const hsl = toHsl(rgb);
      if (hsl.l < 0.42) result = hslCss({ ...hsl, l: clamp(0.42 + hsl.l * 0.5, 0.42, 0.72), s: Math.min(hsl.s, 0.7) });
    }
  }
  memoSet('ink', color, result);
  return result;
}

/**
 * 功能：把单元格底色换算成深色纸张上的底色；时间：2026-10-31；作者：lq。
 *
 * Excel 里底色往往承载语义（黄色=待确认、绿色=已完成），所以不能像 Word 段落底纹那样
 * 统一压成灰，而是「保留色相、只把亮度压到深色纸附近」，让颜色关系仍然可辨。
 */
export function paperFill(color: string | undefined, dark: boolean): string {
  const fallback = 'var(--mo-paper, #ffffff)';
  if (!color) return fallback;
  if (!dark) return color;
  const cached = memoGet('fill', color);
  if (cached) return cached;
  const rgb = parseColor(color);
  let result = color;
  /*
   * 只压暗「浅色底纹」；中深色底纹（表头蓝、强调橙）本身在深色纸上仍然可读，保持原色更忠于文档。
   * 这里按色相压暗而不是统一换成灰：Excel 的底色往往承载语义（黄=待确认、绿=已完成）。
   */
  if (rgb && luma(rgb) >= 200) {
    const hsl = toHsl(rgb);
    result = hslCss({ ...hsl, l: clamp(0.14 + hsl.l * 0.16, 0.14, 0.3), s: Math.min(hsl.s, 0.55) });
  }
  memoSet('fill', color, result);
  return result;
}

/**
 * 功能：把框线颜色换算成当前主题的线条色；时间：2026-10-31；作者：lq。
 *
 * 黑色/深灰框线统一走 --mo-table-line：浅色主题下它是纯黑，深色主题下自动提亮，
 * 打印与导出时被强制重置回纯黑，因此屏幕适配不会影响输出。
 */
export function paperLine(color: string | undefined, fallback = '#000000'): string {
  const value = color ?? fallback;
  if (!parseColor(value)) return value;
  return isDarkInk(value) ? 'var(--mo-table-line, #000000)' : value;
}
