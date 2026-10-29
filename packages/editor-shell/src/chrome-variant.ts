/**
 * 外壳变体读取。
 *
 * 主题插件只通过 `data-chrome-*` 属性影响样式，但某些变体需要改变 DOM 结构
 * （例如 Office 2003 风格需要在标题栏下多出一行菜单栏）。这类变体必须在
 * React 里读到，因此这里把「当前主题包的 chrome 配置」解析成可订阅的 hook。
 */
import { useMemo } from 'react';
import { BUILTIN_THEME_PACK, type ChromeVariant } from '@mini-office/plugin-api';
import { usePluginRegistry } from '@mini-office/plugin-host';

/** 未启用任何主题插件时的外壳形态，等同于当前版本的外观。 */
export const DEFAULT_CHROME: Required<ChromeVariant> = {
  toolbar: 'modern',
  controlShape: 'flat',
  density: 'default',
  titlebar: 'flat',
};

/**
 * 解析当前生效的 chrome 变体。
 * 主题插件被禁用或卸载后自动回退到默认值，不会残留经典外壳。
 */
export function useChromeVariant(): Required<ChromeVariant> {
  const installed = usePluginRegistry((state) => state.installed);
  const activeKey = usePluginRegistry((state) => state.activeThemePack);

  return useMemo(() => {
    if (!activeKey || activeKey === BUILTIN_THEME_PACK) return DEFAULT_CHROME;
    const separator = activeKey.indexOf('#');
    if (separator <= 0) return DEFAULT_CHROME;
    const pluginId = activeKey.slice(0, separator);
    const themeId = activeKey.slice(separator + 1);

    const plugin = installed.find((item) => item.id === pluginId && item.enabled);
    const theme = plugin?.manifest.contributes.themes?.find((item) => item.id === themeId);
    if (!theme?.chrome) return DEFAULT_CHROME;

    return { ...DEFAULT_CHROME, ...theme.chrome };
  }, [installed, activeKey]);
}
