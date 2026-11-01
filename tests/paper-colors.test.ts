import { describe, expect, it } from 'vitest';
import {
  isDarkInk,
  isLightShade,
  paperFill,
  paperInk,
  paperLine,
  parseColor,
  shadeColor,
} from '@mini-office/design-system';

describe('parseColor', () => {
  it('识别 #rgb / #rrggbb / #rrggbbaa', () => {
    expect(parseColor('#fff')).toEqual({ r: 255, g: 255, b: 255, a: 1 });
    expect(parseColor('#1D2939')).toEqual({ r: 29, g: 41, b: 57, a: 1 });
    expect(parseColor('#00000080')?.a).toBeCloseTo(0.5, 1);
  });

  it('识别 rgb() / rgba() 与百分比写法', () => {
    expect(parseColor('rgb(0, 0, 0)')).toEqual({ r: 0, g: 0, b: 0, a: 1 });
    expect(parseColor('rgba(255, 255, 255, 0.5)')?.a).toBe(0.5);
    expect(parseColor('rgb(100% 0% 0%)')).toEqual({ r: 255, g: 0, b: 0, a: 1 });
  });

  it('无法识别的颜色返回 null，交给调用方原样透传', () => {
    expect(parseColor('var(--mo-paper)')).toBeNull();
    expect(parseColor('hsl(0 0% 0%)')).toBeNull();
    expect(parseColor(undefined)).toBeNull();
  });
});

describe('isDarkInk', () => {
  it('只把中性深色判为深墨色', () => {
    expect(isDarkInk('#000000')).toBe(true);
    expect(isDarkInk('#1D2939')).toBe(true);
    expect(isDarkInk('#7F7F7F')).toBe(false);
  });

  it('高饱和强调色不参与提亮，保留原有语义', () => {
    expect(isDarkInk('#C00000')).toBe(false);
    expect(isDarkInk('#1F4E79')).toBe(false);
  });
});

describe('isLightShade', () => {
  it('覆盖 Office 标准浅底纹（最小通道不足 232 也算浅色）', () => {
    expect(isLightShade('#FFFFFF')).toBe(true);
    expect(isLightShade('#FFF2CC')).toBe(true);
    expect(isLightShade('#E2EFDA')).toBe(true);
    expect(isLightShade('#EBF1DE')).toBe(true);
    expect(isLightShade('#FCE4D6')).toBe(true);
  });

  it('中深色底纹不算浅色', () => {
    expect(isLightShade('#4472C4')).toBe(false);
    expect(isLightShade('#FFC000')).toBe(false);
    expect(isLightShade('#000000')).toBe(false);
  });

  it('段落底纹仍走统一变量，打印与导出可还原', () => {
    expect(shadeColor('#FFF2CC')).toBe('var(--mo-shade-light, #ffffff)');
    expect(shadeColor('#4472C4')).toBe('#4472C4');
    expect(shadeColor(undefined)).toBeUndefined();
  });
});

describe('paperInk', () => {
  it('浅色主题原样返回文档颜色', () => {
    expect(paperInk('#1D2939', false)).toBe('#1D2939');
    expect(paperInk('#C00000', false)).toBe('#C00000');
  });

  it('深色主题下中性深色走纸张文字色', () => {
    expect(paperInk(undefined, true)).toBe('var(--mo-paper-ink, #1d2939)');
    expect(paperInk('#1D2939', true)).toBe('var(--mo-paper-ink, #1d2939)');
  });

  it('深色主题下偏暗的强调色同色相提亮到可读亮度', () => {
    const lifted = paperInk('#C00000', true);
    expect(lifted).toMatch(/^hsl\(/);
    const lightness = Number.parseFloat(/hsl\([\d.]+ [\d.]+% ([\d.]+)%/.exec(lifted)?.[1] ?? '0');
    expect(lightness).toBeGreaterThanOrEqual(42);
    expect(lightness).toBeLessThanOrEqual(72);
  });

  it('已经够亮的颜色保持不变', () => {
    expect(paperInk('#FFD966', true)).toBe('#FFD966');
  });
});

describe('paperFill', () => {
  it('未指定底色时跟随纸张底色（打印/导出由变量重置回白纸）', () => {
    expect(paperFill(undefined, false)).toBe('var(--mo-paper, #ffffff)');
    expect(paperFill(undefined, true)).toBe('var(--mo-paper, #ffffff)');
  });

  it('深色主题把浅色底纹按色相压暗，而不是统一变灰', () => {
    const yellow = paperFill('#FFF2CC', true);
    const green = paperFill('#E2EFDA', true);
    expect(yellow).toMatch(/^hsl\(/);
    expect(green).toMatch(/^hsl\(/);
    expect(yellow).not.toBe(green);
    const hue = (value: string) => Number.parseFloat(/hsl\(([\d.]+)/.exec(value)?.[1] ?? '0');
    // 黄色底纹的色相仍落在暖色区，绿色底纹仍落在绿色区
    expect(hue(yellow)).toBeGreaterThan(30);
    expect(hue(yellow)).toBeLessThan(70);
    expect(hue(green)).toBeGreaterThan(80);
    expect(hue(green)).toBeLessThan(160);
  });

  it('中深色底纹保持原色，不被抹平', () => {
    expect(paperFill('#4472C4', true)).toBe('#4472C4');
    expect(paperFill('#FFC000', true)).toBe('#FFC000');
  });
});

describe('paperLine', () => {
  it('黑色框线走主题变量，浅色/打印时是纯黑、深色时提亮', () => {
    expect(paperLine('#000000')).toBe('var(--mo-table-line, #000000)');
    expect(paperLine(undefined)).toBe('var(--mo-table-line, #000000)');
  });

  it('彩色框线保持原色', () => {
    expect(paperLine('#C00000')).toBe('#C00000');
  });
});
