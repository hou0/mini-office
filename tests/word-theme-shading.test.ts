import { describe, expect, it } from 'vitest';
import { resolveThemeSlot, resolveWordShading } from '@mini-office/ooxml';

const THEME = {
  tx1: '#000000',
  bg1: '#FFFFFF',
  accent1: '#4472C4',
  accent2: '#ED7D31',
  lt1: '#FFFFFF',
};

describe('Word 主题色解析', () => {
  it('主题色槽位名映射到主题表键名', () => {
    expect(resolveThemeSlot('accent1', THEME)).toBe('#4472C4');
    expect(resolveThemeSlot('text1', THEME)).toBe('#000000');
    expect(resolveThemeSlot('background1', THEME)).toBe('#FFFFFF');
  });

  it('主题色叠加变浅百分比后仍返回具体颜色', () => {
    // w:themeTint 是单字节十六进制：0x33 = 51 → 20% 保留 → 向白色混 80%。
    const tinted = resolveThemeSlot('accent1', THEME, '33');
    expect(tinted).toBeDefined();
    expect(tinted).not.toBe('#4472C4');
    expect(tinted!.startsWith('#')).toBe(true);
    // 变浅后亮度必须高于原色，否则深色纸上仍是一块暗色。
    const brightness = (hex: string) => {
      const n = Number.parseInt(hex.slice(1), 16);
      return ((n >> 16) & 255) + ((n >> 8) & 255) + (n & 255);
    };
    expect(brightness(tinted!)).toBeGreaterThan(brightness('#4472C4'));
  });

  it('themeFill 在没有 w:fill 时也能解析出底纹', () => {
    expect(resolveWordShading({ '@_w:themeFill': 'accent1' }, THEME)).toBe('#4472C4');
  });

  it('w:fill 优先于 themeFill，保持与 Word 一致的覆盖关系', () => {
    expect(resolveWordShading({ '@_w:fill': 'FFF2CC', '@_w:themeFill': 'accent1' }, THEME)).toBe('#FFF2CC');
  });

  it('显式无底纹时返回 undefined，不会退回主题色', () => {
    expect(resolveWordShading({ '@_w:val': 'nil', '@_w:themeFill': 'accent1' }, THEME)).toBeUndefined();
    expect(resolveWordShading({ '@_w:val': 'none', '@_w:fill': 'FFFFFF' }, THEME)).toBeUndefined();
  });

  it('auto 填充与缺失主题表时安全退化', () => {
    expect(resolveWordShading({ '@_w:fill': 'auto', '@_w:themeFill': 'accent9' }, THEME)).toBeUndefined();
    expect(resolveWordShading({ '@_w:themeFill': 'accent1' }, {})).toBeUndefined();
    expect(resolveWordShading(null, THEME)).toBeUndefined();
  });
});
