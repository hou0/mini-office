import { describe, expect, it } from 'vitest';
import type { InstalledPlugin, PluginIndex, PluginManifest } from '@mini-office/plugin-api';
import { collectUpdates, compareVersions, type PluginMacroCommand } from '@mini-office/plugin-host';
import {
  RESERVED_SHORTCUTS,
  collectShortcutBindings,
  formatShortcut,
  matchShortcutFromEvent,
  normalizeShortcut,
  shortcutFromEvent,
} from '../packages/editor-shell/src/shortcuts';

function macro(pluginId: string, commandId: string, shortcut?: string): PluginMacroCommand {
  return {
    key: `${pluginId}#${commandId}`,
    pluginId,
    pluginName: `插件 ${pluginId}`,
    kind: 'docx',
    command: {
      id: commandId,
      title: commandId,
      shortcut,
      steps: [{ op: 'word.paragraph.format', lineHeight: 1.5 }],
    },
  };
}

function keydown(init: KeyboardEventInit): KeyboardEvent {
  return new KeyboardEvent('keydown', { bubbles: true, cancelable: true, ...init });
}

function manifest(id: string, version: string): PluginManifest {
  return {
    id,
    name: id,
    version,
    publisher: 'mini-office',
    tier: 'declarative',
    engines: { miniOffice: '>=0.1.5', api: '^1.0.0' },
    permissions: [],
    contributes: {
      commands: [{ id: 'c', title: 'c', steps: [{ op: 'word.page.orientation', orientation: 'landscape' }] }],
    },
  };
}

function installed(id: string, version: string): InstalledPlugin {
  return {
    id,
    version,
    tier: 'declarative',
    enabled: true,
    source: 'github',
    installedAt: 0,
    manifest: manifest(id, version),
  };
}

function indexEntry(id: string, version: string) {
  return {
    id,
    name: id,
    version,
    publisher: 'mini-office',
    tier: 'declarative' as const,
    permissions: [],
    download: `https://example.com/${id}-${version}.moext`,
    sha256: 'a'.repeat(64),
  };
}

describe('快捷键归一化', () => {
  it('接受常见写法并统一修饰键顺序', () => {
    expect(normalizeShortcut('Mod+Shift+K')).toBe('mod+shift+k');
    expect(normalizeShortcut('Shift+Cmd+K')).toBe('mod+shift+k');
    expect(normalizeShortcut('Ctrl+Alt+P')).toBe('mod+alt+p');
    expect(normalizeShortcut('Mod+,')).toBe('mod+,');
  });

  it('拒绝没有修饰键的组合，避免抢占正常输入', () => {
    expect(normalizeShortcut('K')).toBeNull();
    expect(normalizeShortcut('F5')).toBeNull();
    expect(normalizeShortcut('Shift')).toBeNull();
  });

  it('拒绝非法写法', () => {
    expect(normalizeShortcut('Mod+K+J')).toBeNull();
    expect(normalizeShortcut('Mod+Hyper+K')).toBeNull();
    expect(normalizeShortcut('')).toBeNull();
  });

  it('从事件还原出同样的字符串', () => {
    expect(shortcutFromEvent(keydown({ key: 'k', metaKey: true, shiftKey: true }))).toBe('mod+shift+k');
    // Windows/Linux 上 Ctrl 与 Cmd 都归一到 mod。
    expect(shortcutFromEvent(keydown({ key: 'K', ctrlKey: true }))).toBe('mod+k');
    // 只按修饰键本身不构成组合。
    expect(shortcutFromEvent(keydown({ key: 'Shift', shiftKey: true }))).toBeNull();
    expect(shortcutFromEvent(keydown({ key: 'k' }))).toBeNull();
  });

  it('展示写法带上平台修饰键', () => {
    const formatted = formatShortcut('mod+shift+k');
    expect(formatted.toUpperCase()).toContain('K');
    expect(formatted).not.toContain('mod');
  });
});

describe('快捷键冲突检测', () => {
  it('与内置快捷键冲突时标记原因而不是静默失效', () => {
    const bindings = collectShortcutBindings([macro('a.b', 'save', 'Mod+S')]);
    expect(bindings[0].shortcut).toBe('mod+s');
    expect(bindings[0].conflict).toContain('内置快捷键');
    expect(RESERVED_SHORTCUTS).toContain('mod+s');
  });

  it('两个插件声明同一组合时先到先得，后者标记冲突', () => {
    const bindings = collectShortcutBindings([
      macro('a.b', 'first', 'Mod+Shift+G'),
      macro('c.d', 'second', 'Mod+Shift+G'),
    ]);
    expect(bindings[0].conflict).toBeNull();
    expect(bindings[1].conflict).toContain('相同快捷键');
  });

  it('格式错误的快捷键给出可读原因', () => {
    const bindings = collectShortcutBindings([macro('a.b', 'bad', 'G')]);
    expect(bindings[0].shortcut).toBeNull();
    expect(bindings[0].conflict).toContain('格式无效');
  });

  it('没有声明快捷键的命令不参与冲突判定', () => {
    const bindings = collectShortcutBindings([macro('a.b', 'plain')]);
    expect(bindings[0].shortcut).toBeNull();
    expect(bindings[0].conflict).toBeNull();
  });

  it('事件只命中无冲突的绑定', () => {
    const bindings = collectShortcutBindings([
      macro('a.b', 'save', 'Mod+S'),
      macro('c.d', 'apply', 'Mod+Shift+G'),
    ]);
    expect(matchShortcutFromEvent(keydown({ key: 's', metaKey: true }), bindings)).toBeNull();
    expect(matchShortcutFromEvent(keydown({ key: 'g', metaKey: true, shiftKey: true }), bindings)?.macro.command.id).toBe(
      'apply',
    );
    expect(matchShortcutFromEvent(keydown({ key: 'g', metaKey: true }), bindings)).toBeNull();
  });
});

describe('版本比较与更新检测', () => {
  it('按主次修订号比较', () => {
    expect(compareVersions('1.1.0', '1.0.0')).toBeGreaterThan(0);
    expect(compareVersions('1.0.0', '1.0.0')).toBe(0);
    expect(compareVersions('1.0.9', '1.1.0')).toBeLessThan(0);
    expect(compareVersions('2.0.0', '1.99.99')).toBeGreaterThan(0);
  });

  it('无法解析的版本返回 null 而不是当作相等', () => {
    expect(compareVersions('abc', '1.0.0')).toBeNull();
    expect(compareVersions('1.0.0', '')).toBeNull();
  });

  it('只有远端版本更高时才提示更新', () => {
    const index: PluginIndex = {
      schemaVersion: 1,
      updatedAt: '2026-10-29T00:00:00.000Z',
      plugins: [indexEntry('a.one', '1.1.0'), indexEntry('b.two', '2.0.0'), indexEntry('c.three', '1.0.0')],
    };
    const updates = collectUpdates(
      [installed('a.one', '1.0.0'), installed('b.two', '3.0.0'), installed('d.four', '1.0.0')],
      index,
    );
    expect(updates.map((item) => item.pluginId)).toEqual(['a.one']);
    expect(updates[0]).toMatchObject({ from: '1.0.0', to: '1.1.0' });
  });

  it('没有索引时不报更新', () => {
    expect(collectUpdates([installed('a.one', '1.0.0')], null)).toEqual([]);
  });
});
