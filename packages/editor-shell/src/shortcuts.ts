/**
 * 插件快捷键的解析、归一化与冲突检测。
 *
 * 设计取舍：
 * 1. **只做纯函数**，不碰 React 与 store，便于单测直接覆盖。
 * 2. **必须带修饰键**。裸字母/F 键不做分发——插件一旦声明 `K` 就能在用户打字时
 *    抢走按键，这类"装上就坏"的体验比不支持快捷键严重得多。
 * 3. **冲突显式化**。与内置命令或其它插件撞车时不静默失效，而是标记 `conflict`，
 *    由命令面板展示原因并禁用——用户看到"为什么按不动"，而不是以为功能坏了。
 */
import { MOD_KEY } from '@mini-office/shared';
import { normalizeShortcut } from '@mini-office/plugin-api';
import type { PluginMacroCommand } from '@mini-office/plugin-host';

type Modifier = 'mod' | 'shift' | 'alt';

const MODIFIER_ORDER: Modifier[] = ['mod', 'shift', 'alt'];

/**
 * 内置命令已占用的快捷键。
 *
 * 与 `AppShell.tsx` 的 keydown 处理保持一致——新增内置快捷键时要同步这里，
 * 否则插件可以声明一个"看起来能用、实际被内置命令截胡"的组合。
 * 运行时内置命令优先（插件分发在它之后），所以遗漏只会导致插件命令失效，
 * 不会造成两个动作同时触发。
 */
export const RESERVED_SHORTCUTS: readonly string[] = [
  'mod+k',
  'mod+s',
  'mod+shift+s',
  'mod+o',
  'mod+n',
  'mod+,',
  'mod+w',
  'mod+p',
  'mod+f',
  ...Array.from({ length: 9 }, (_, index) => `mod+${index + 1}`),
];

/** 归一化单个主键；不认识时返回 null。 */
function normalizeKey(raw: string): string | null {
  const key = raw.trim().toLowerCase();
  if (!key) return null;
  if (key === 'space') return 'space';
  if (key.length === 1) return key;
  if (/^f([1-9]|1[0-2])$/.test(key)) return key;
  return null;
}

/** 归一化实现来自契约包，这里只做转出，保证校验与运行时用同一套规则。 */
export { normalizeShortcut };

/** 事件里的主键归一化；`event.key` 可能是 'Shift'、' ' 这类特殊值。 */
function normalizeEventKey(raw: string): string | null {
  const key = raw.toLowerCase();
  if (key === ' ' || key === 'spacebar') return 'space';
  if (key.length === 1) return key;
  if (/^f([1-9]|1[0-2])$/.test(key)) return key;
  return null;
}

/** 从键盘事件得到同样的归一化字符串；修饰键本身或缺少修饰键时返回 null。 */
export function shortcutFromEvent(event: KeyboardEvent): string | null {
  const modifiers: Modifier[] = [];
  if (event.metaKey || event.ctrlKey) modifiers.push('mod');
  if (event.shiftKey) modifiers.push('shift');
  if (event.altKey) modifiers.push('alt');
  if (modifiers.length === 0) return null;

  const key = normalizeEventKey(event.key);
  if (!key) return null;
  return [...modifiers, key].join('+');
}

/** 把归一化结果转成界面展示用的写法。 */
export function formatShortcut(shortcut: string): string {
  const parts = shortcut.split('+');
  const key = parts[parts.length - 1] ?? '';
  const modifiers = parts.slice(0, -1);
  const modPart = MOD_KEY;
  const altPart = MOD_KEY === '⌘' ? '⌥' : 'Alt';
  const shiftPart = '⇧';
  const labels: string[] = [];
  for (const modifier of MODIFIER_ORDER) {
    if (!modifiers.includes(modifier)) continue;
    labels.push(modifier === 'mod' ? modPart : modifier === 'alt' ? altPart : shiftPart);
  }
  const keyLabel = key === 'space' ? 'Space' : key.toUpperCase();
  return MOD_KEY === '⌘' ? `${labels.join('')}${keyLabel}` : `${labels.join('+')}+${keyLabel}`;
}

export interface ShortcutBinding {
  macro: PluginMacroCommand;
  /** 归一化后的快捷键；无法解析时为 null。 */
  shortcut: string | null;
  /** 非空表示这条绑定不可用，值为可读原因。 */
  conflict: string | null;
}

/**
 * 为一批插件命令建立快捷键绑定，并标出冲突项。
 *
 * 同一快捷键被多个插件声明时，先声明者生效，其余标记冲突——不随机挑选，
 * 避免"有时能用有时不能用"这种更难排查的表现。
 */
export function collectShortcutBindings(macros: PluginMacroCommand[]): ShortcutBinding[] {
  const reserved = new Set(RESERVED_SHORTCUTS);
  const owners = new Map<string, string>();

  return macros.map((macro) => {
    const declared = macro.command.shortcut?.trim();
    if (!declared) return { macro, shortcut: null, conflict: null };

    const shortcut = normalizeShortcut(declared);
    if (!shortcut) {
      return { macro, shortcut: null, conflict: '快捷键格式无效（需形如 Mod+Shift+K）' };
    }
    if (reserved.has(shortcut)) {
      return { macro, shortcut, conflict: `与内置快捷键 ${formatShortcut(shortcut)} 冲突` };
    }
    const owner = owners.get(shortcut);
    if (owner) {
      return { macro, shortcut, conflict: `与「${owner}」使用了相同快捷键` };
    }
    owners.set(shortcut, macro.pluginName);
    return { macro, shortcut, conflict: null };
  });
}

/** 命中事件对应的绑定；冲突项不参与匹配。 */
export function matchShortcutFromEvent(
  event: KeyboardEvent,
  bindings: ShortcutBinding[],
): ShortcutBinding | null {
  const pressed = shortcutFromEvent(event);
  if (!pressed) return null;
  return bindings.find((binding) => binding.shortcut === pressed && !binding.conflict) ?? null;
}
