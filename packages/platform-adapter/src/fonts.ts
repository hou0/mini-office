import type { FontInfo } from '@mini-office/shared';

const CHINESE_FONT_CANDIDATES = [
  '宋体',
  'Songti SC',
  '仿宋',
  'FangSong',
  '黑体',
  'Heiti SC',
  '楷体',
  'Kaiti SC',
  '微软雅黑',
  'Microsoft YaHei',
  '等线',
  'DengXian',
  '思源宋体',
  'Source Han Serif SC',
  '思源黑体',
  'Source Han Sans SC',
  'Noto Serif CJK SC',
  'Noto Sans CJK SC',
  'PingFang SC',
  'Hiragino Sans GB',
  'STSong',
  'STHeiti',
  'WenQuanYi Micro Hei',
  'SimSun',
  'SimHei',
];

const LATIN_FONT_CANDIDATES = [
  'Helvetica Neue',
  'Arial',
  'Times New Roman',
  'Courier New',
  'Georgia',
  'Verdana',
  'Tahoma',
  'Trebuchet MS',
  'Segoe UI',
  'SF Pro Text',
  'Menlo',
  'Consolas',
];

function hasChineseName(font: string): boolean {
  return /[\u3400-\u9fff]/.test(font) || CHINESE_FONT_CANDIDATES.includes(font);
}

/** Best-effort font availability probe using canvas metric comparison. */
export function detectWebFonts(): FontInfo[] {
  if (typeof document === 'undefined') return [];
  const canvas = document.createElement('canvas');
  const context = canvas.getContext('2d');
  if (!context) return [];
  const testText = 'mmmmmmmmmmlli中文测试';
  const bases = ['monospace', 'serif', 'sans-serif'];
  const baseSizes = new Map<string, number[]>();
  for (const base of bases) {
    context.font = `72px ${base}`;
    baseSizes.set(base, bases.map(() => context.measureText(testText).width));
  }
  const available: FontInfo[] = [];
  for (const font of [...CHINESE_FONT_CANDIDATES, ...LATIN_FONT_CANDIDATES]) {
    let detected = false;
    for (const [index, base] of bases.entries()) {
      context.font = `72px "${font}", ${base}`;
      const width = context.measureText(testText).width;
      const baseWidth = baseSizes.get(base)?.[index] ?? width;
      if (Math.abs(width - baseWidth) > 0.5) {
        detected = true;
        break;
      }
    }
    if (detected) {
      available.push({
        family: font,
        source: 'system',
        hasChinese: hasChineseName(font),
      });
    }
  }
  return available;
}

export const DEFAULT_AVAILABLE_FONTS: FontInfo[] = [
  'PingFang SC',
  'Songti SC',
  'Heiti SC',
  'Kaiti SC',
  'Helvetica Neue',
  'Arial',
  'Times New Roman',
  'Courier New',
].map((family) => ({ family, source: 'assumed', hasChinese: hasChineseName(family) }));
