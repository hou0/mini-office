/**
 * 功能：验证低频排版设置的默认值与一次性迁移，避免「改了默认值但对老存档无效」。
 * 时间：2026-10-30；作者：lq。
 *
 * 背景：Word 标尺最初默认显示，用户反馈与 Word / WPS 不一致（原版默认不显示标尺）。
 * 默认值改为隐藏后，已把 true 写进 localStorage 的用户仍然会看到标尺，
 * 因此用一次性迁移标记把这些存档改回隐藏；用户之后手动开启必须能正常保留。
 */
import { beforeEach, describe, expect, it, vi } from 'vitest';

const SETTINGS_KEY = 'mini-office.settings';
const MIGRATION_KEY = 'mini-office.settings.migrated.v1';

async function loadFreshStore() {
  vi.resetModules();
  return import('@mini-office/document-model');
}

describe('设置默认值与迁移', () => {
  beforeEach(() => {
    localStorage.clear();
  });

  it('标尺与表格阅读模式默认关闭，护眼模式默认关闭', async () => {
    const { DEFAULT_SETTINGS } = await import('@mini-office/shared');
    expect(DEFAULT_SETTINGS.wordRuler).toBe(false);
    expect(DEFAULT_SETTINGS.sheetCrosshair).toBe(false);
    expect(DEFAULT_SETTINGS.eyeCare).toBe(false);
  });

  it('旧存档里被写成 true 的标尺会一次性改回隐藏', async () => {
    localStorage.setItem(SETTINGS_KEY, JSON.stringify({ wordRuler: true, theme: 'dark' }));
    const { useWorkspaceStore } = await loadFreshStore();
    const settings = useWorkspaceStore.getState().settings;
    expect(settings.wordRuler).toBe(false);
    // 迁移只动标尺，其它已保存的设置保持原样。
    expect(settings.theme).toBe('dark');
    expect(localStorage.getItem(MIGRATION_KEY)).toBe('1');
  });

  it('迁移只执行一次，用户之后手动开启的标尺不会被重置', async () => {
    localStorage.setItem(MIGRATION_KEY, '1');
    localStorage.setItem(SETTINGS_KEY, JSON.stringify({ wordRuler: true }));
    const { useWorkspaceStore } = await loadFreshStore();
    expect(useWorkspaceStore.getState().settings.wordRuler).toBe(true);
  });

  it('从未保存过设置时按新默认值启动', async () => {
    const { useWorkspaceStore } = await loadFreshStore();
    expect(useWorkspaceStore.getState().settings.wordRuler).toBe(false);
    // 没有旧存档也要落迁移标记，避免下次启动再走一遍。
    expect(localStorage.getItem(MIGRATION_KEY)).toBe('1');
  });
});
