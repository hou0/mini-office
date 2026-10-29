/**
 * 功能：验证中文字号映射、行距/缩进换算与格式刷的样式套用，覆盖用户反馈的字号与行距问题。
 * 时间：2026-10-29；作者：lq
 */
import { describe, expect, it } from 'vitest';
import type { DocRun } from '../packages/document-model/src/word';
import {
  CN_FONT_SIZES,
  applyRunStyleToRuns,
  cnFontSizeName,
  fontSizeLabel,
  lineSpacingLabel,
  mergeParagraphStyle,
  parseFontSizeInput,
  pickParagraphStyle,
  pxToChars,
  readLineSpacing,
  readSpecialIndent,
  writeLineSpacing,
  writeSpecialIndent,
} from '../packages/editor-word/src/typography';

describe('中文字号', () => {
  it('覆盖 Word / WPS 的 16 档中文字号', () => {
    expect(CN_FONT_SIZES).toHaveLength(16);
    const map = Object.fromEntries(CN_FONT_SIZES.map((item) => [item.name, item.pt]));
    expect(map['初号']).toBe(42);
    expect(map['三号']).toBe(16);
    expect(map['小四']).toBe(12);
    expect(map['五号']).toBe(10.5);
    expect(map['八号']).toBe(5);
  });

  it('解析中文名、纯数字与带单位输入', () => {
    expect(parseFontSizeInput('三号')).toBe(16);
    expect(parseFontSizeInput(' 小四 ')).toBe(12);
    expect(parseFontSizeInput('16')).toBe(16);
    expect(parseFontSizeInput('16磅')).toBe(16);
    expect(parseFontSizeInput('13.5pt')).toBe(13.5);
    expect(parseFontSizeInput('0')).toBeUndefined();
    expect(parseFontSizeInput('三号半')).toBeUndefined();
    expect(parseFontSizeInput('')).toBeUndefined();
    // 上限按 Word 的 1638 磅收敛，避免写入非法字号撑爆版面。
    expect(parseFontSizeInput('2000')).toBe(1638);
  });

  it('命中中文字号时显示名称，否则只显示磅值', () => {
    expect(cnFontSizeName(16)).toBe('三号');
    expect(cnFontSizeName(13.5)).toBeUndefined();
    expect(fontSizeLabel(16)).toBe('三号 16');
    expect(fontSizeLabel(13.5)).toBe('13.5');
    expect(fontSizeLabel(undefined)).toBe('默认');
  });
});

describe('行距', () => {
  it('读取三种行距模式', () => {
    expect(readLineSpacing({ lineHeight: 1.5 })).toEqual({ mode: 'multiple', value: 1.5 });
    expect(readLineSpacing({ lineHeightExactPx: 24 })).toEqual({ mode: 'exact', value: 18 });
    expect(readLineSpacing({ lineHeightAtLeastPx: 20 })).toEqual({ mode: 'atLeast', value: 15 });
    expect(readLineSpacing(undefined)).toEqual({ mode: 'multiple', value: 1.5 });
  });

  it('写回时清理其它模式字段，支持任意磅值', () => {
    const exact = writeLineSpacing({ mode: 'exact', value: 22 });
    expect(exact.lineHeight).toBeUndefined();
    expect(exact.lineHeightAtLeastPx).toBeUndefined();
    expect(exact.lineHeightExactPx).toBeCloseTo(29.33, 1);

    const multiple = writeLineSpacing({ mode: 'multiple', value: 2.5 });
    expect(multiple.lineHeight).toBe(2.5);
    expect(multiple.lineHeightExactPx).toBeUndefined();

    const atLeast = writeLineSpacing({ mode: 'atLeast', value: 30 });
    expect(atLeast.lineHeightAtLeastPx).toBeCloseTo(40, 1);
  });

  it('工具栏标签区分倍数、固定值与最小值', () => {
    expect(lineSpacingLabel({ mode: 'multiple', value: 1.5 })).toBe('1.5 倍');
    expect(lineSpacingLabel({ mode: 'exact', value: 18 })).toBe('固定 18 磅');
    expect(lineSpacingLabel({ mode: 'atLeast', value: 20 })).toBe('最小 20 磅');
  });
});

describe('特殊缩进', () => {
  it('正值为首行缩进，负值为悬挂缩进', () => {
    expect(readSpecialIndent({ firstLineIndentPx: 32 }, 12)).toEqual({ mode: 'firstLine', chars: 2 });
    expect(readSpecialIndent({ firstLineIndentPx: -32 }, 12)).toEqual({ mode: 'hanging', chars: 2 });
    expect(readSpecialIndent({}, 12)).toEqual({ mode: 'none', chars: 0 });
    expect(pxToChars(32, 12)).toBeCloseTo(2, 1);
  });

  it('按字号换算字符宽度并写回正负号', () => {
    expect(writeSpecialIndent('none', 2, 12)).toEqual({ firstLineIndentPx: 0 });
    const firstLine = writeSpecialIndent('firstLine', 2, 12);
    expect(firstLine.firstLineIndentPx).toBeGreaterThan(0);
    const hanging = writeSpecialIndent('hanging', 2, 12);
    expect(hanging.firstLineIndentPx).toBeLessThan(0);
    expect(Math.abs(hanging.firstLineIndentPx!)).toBeCloseTo(firstLine.firstLineIndentPx!, 5);
  });
});

describe('格式刷样式复制', () => {
  it('只复制排版相关字段，不带走边框与底纹', () => {
    const painted = pickParagraphStyle({
      alignment: 'center',
      firstLineIndentPx: 24,
      lineHeightExactPx: 24,
      backgroundColor: '#FFF1F0',
      borders: { top: { style: 'single', widthPx: 1, color: '#000000' } },
    });
    expect(painted.alignment).toBe('center');
    expect(painted.firstLineIndentPx).toBe(24);
    expect(painted.lineHeightExactPx).toBe(24);
    expect(painted.backgroundColor).toBeUndefined();
    expect(painted.borders).toBeUndefined();
  });

  it('套用时覆盖已复制字段并保留目标原有内容型格式', () => {
    const merged = mergeParagraphStyle(
      { alignment: 'left', backgroundColor: '#EEF3FF', spaceAfterPx: 8 },
      { alignment: 'right', firstLineIndentPx: 0 },
    );
    expect(merged.alignment).toBe('right');
    expect(merged.firstLineIndentPx).toBe(0);
    expect(merged.backgroundColor).toBe('#EEF3FF');
    // 未复制的字段应被清空，保证“格式完全一致”。
    expect(merged.spaceAfterPx).toBeUndefined();
  });

  it('按原子偏移套用字体格式且不破坏图片与文字内容', () => {
    const runs: DocRun[] = [
      { text: 'abc', style: {} },
      { image: { src: 'data:image/png;base64,xx', widthPx: 10, heightPx: 10 }, style: {} },
      { text: 'def', style: { bold: true } },
    ];
    const painted = { bold: false, fontSizePt: 16, fontFamily: '宋体' };
    // 覆盖 "bc"（偏移 1-3）。
    const result = applyRunStyleToRuns(runs, 1, 3, painted);
    const text = result.map((run) => run.text ?? '').join('');
    expect(text).toBe('abcdef');
    expect(result.some((run) => run.image)).toBe(true);
    // 切分结果是 "a" + "bc"（区间内） + 图片 + "def"。
    const styled = result.filter((run) => run.text === 'bc');
    expect(styled).toHaveLength(1);
    expect(styled.every((run) => run.style.fontSizePt === 16 && run.style.fontFamily === '宋体')).toBe(true);
    expect(styled.every((run) => run.style.bold === false)).toBe(true);
    // 区间外文字保持原样式。
    expect(result[0].style.fontSizePt).toBeUndefined();
    expect(result.find((run) => run.text === 'def')?.style.bold).toBe(true);
  });

  it('空区间不产生任何修改', () => {
    const runs: DocRun[] = [{ text: 'abc', style: {} }];
    expect(applyRunStyleToRuns(runs, 2, 2, { bold: true })).toBe(runs);
  });
});
