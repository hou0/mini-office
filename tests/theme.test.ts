import { beforeEach, describe, expect, it } from 'vitest';
import {
  ACCENT_PRESETS,
  applyAccent,
  applyAppearance,
  contrastRatio,
  normalizeAccentHex,
  normalizeAccentId,
  resolveAccentPalette,
} from '@mini-office/design-system';
import { DEFAULT_ACCENT_HEX, DEFAULT_SETTINGS } from '@mini-office/shared';

function cssVar(name: string): string {
  return document.documentElement.style.getPropertyValue(name);
}

describe('theme accents', () => {
  beforeEach(() => {
    document.documentElement.removeAttribute('data-theme');
    document.documentElement.removeAttribute('data-accent');
    document.documentElement.removeAttribute('style');
  });

  it('ships five built-in palettes plus a custom option', () => {
    expect(ACCENT_PRESETS.map((preset) => preset.id)).toEqual([
      'mist',
      'ocean',
      'iris',
      'pine',
      'clay',
    ]);
    expect(normalizeAccentId('custom')).toBe('custom');
    expect(normalizeAccentId('not-a-theme')).toBe('mist');
  });

  it('keeps every preset readable in both themes', () => {
    for (const preset of ACCENT_PRESETS) {
      expect(contrastRatio(preset.light.accent, preset.light.ink)).toBeGreaterThanOrEqual(4.5);
      expect(contrastRatio(preset.dark.accent, preset.dark.ink)).toBeGreaterThanOrEqual(4.5);
    }
  });

  it('applies the default mist palette through CSS custom properties', () => {
    applyAccent('mist', DEFAULT_ACCENT_HEX, 'light');
    expect(document.documentElement.dataset.accent).toBe('mist');
    expect(cssVar('--mo-accent')).toBe('#2E7D72');
    expect(cssVar('--mo-accent-hover')).toBe('#25665D');
    expect(cssVar('--mo-accent-soft')).toBe('#E6F2EF');
    expect(cssVar('--mo-accent-ink')).toBe('#FFFFFF');
  });

  it('uses the dark variant with a dark ink for filled controls', () => {
    applyAccent('mist', DEFAULT_ACCENT_HEX, 'dark');
    expect(cssVar('--mo-accent')).toBe('#6CC9BB');
    expect(cssVar('--mo-accent-ink')).toBe('#10221F');
  });

  it('derives readable shades for a custom color', () => {
    const shades = resolveAccentPalette('custom', '#369', 'light');
    expect(shades.accent).toBe('#336699');
    expect(contrastRatio(shades.accent, shades.ink)).toBeGreaterThanOrEqual(4.5);
    expect(shades.hover).not.toBe(shades.accent);
    expect(shades.soft).not.toBe(shades.accent);
  });

  it('normalizes hex input and rejects malformed values', () => {
    expect(normalizeAccentHex('#abc')).toBe('#AABBCC');
    expect(normalizeAccentHex('2e7d72')).toBe('#2E7D72');
    expect(normalizeAccentHex('#12345')).toBeNull();
    expect(normalizeAccentHex('teal')).toBeNull();
  });

  it('keeps accent and theme in sync when appearance changes', () => {
    applyAppearance({ theme: 'light', accent: 'iris', customAccent: DEFAULT_ACCENT_HEX });
    expect(document.documentElement.dataset.theme).toBe('light');
    expect(document.documentElement.dataset.accent).toBe('iris');
    expect(cssVar('--mo-accent')).toBe('#6B5BD6');

    applyAppearance({ theme: 'dark', accent: 'clay', customAccent: DEFAULT_ACCENT_HEX });
    expect(document.documentElement.dataset.theme).toBe('dark');
    expect(document.documentElement.dataset.accent).toBe('clay');
    expect(cssVar('--mo-accent')).toBe('#F0A892');
  });

  // 功能：护眼模式只切换 data-eyecare，未传该字段时应回落为关闭；时间：2026-10-30；作者：lq。
  it('toggles the eye-care flag on the root element', () => {
    applyAppearance({ theme: 'light', accent: 'iris', customAccent: DEFAULT_ACCENT_HEX, eyeCare: true });
    expect(document.documentElement.dataset.eyecare).toBe('on');

    applyAppearance({ theme: 'light', accent: 'iris', customAccent: DEFAULT_ACCENT_HEX, eyeCare: false });
    expect(document.documentElement.dataset.eyecare).toBe('off');

    applyAppearance({ theme: 'light', accent: 'iris', customAccent: DEFAULT_ACCENT_HEX });
    expect(document.documentElement.dataset.eyecare).toBe('off');
  });

  it('defaults new settings to mist and native vibrancy off', () => {
    expect(DEFAULT_SETTINGS.accent).toBe('mist');
    expect(DEFAULT_SETTINGS.customAccent).toBe(DEFAULT_ACCENT_HEX);
    expect(DEFAULT_SETTINGS.vibrancy).toBe(false);
  });
});
