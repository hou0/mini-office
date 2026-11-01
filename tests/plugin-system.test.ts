import { describe, expect, it } from 'vitest';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import {
  applyThemePack,
  contrastRatio,
  isSafeEntryPath,
  parseHexColor,
  parsePluginIndex,
  parsePackage,
  satisfiesRange,
  themePackKey,
  validateManifest,
  validateTheme,
} from '@mini-office/plugin-host';
import { createMemoryStore } from '@mini-office/plugin-host';
import { installFromBytes } from '@mini-office/plugin-host';

const VALID_MANIFEST = {
  id: 'mini-office.theme-classic',
  name: '经典外观',
  version: '1.0.0',
  publisher: 'mini-office',
  tier: 'declarative',
  engines: { miniOffice: '>=0.1.5 <0.5.0', api: '^1.0.0' },
  permissions: [],
  contributes: {
    themes: [
      {
        id: 'classic',
        label: '经典',
        tokens: {
          '--mo-bg': '#ECE9D8',
          '--mo-surface': '#FFFFFF',
          '--mo-surface-2': '#F1EFE2',
          '--mo-text-primary': '#000000',
          '--mo-text-muted': '#555555',
          '--mo-line': '#ACA899',
        },
      },
    ],
  },
};

async function buildPackage(manifest: unknown, extra: Record<string, string> = {}): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('manifest.json', JSON.stringify(manifest));
  for (const [path, content] of Object.entries(extra)) zip.file(path, content);
  return zip.generateAsync({ type: 'uint8array' });
}

describe('插件清单校验', () => {
  it('接受合法的声明式主题插件', () => {
    const result = validateManifest(VALID_MANIFEST, '0.1.5');
    expect(result.ok).toBe(true);
    expect(result.manifest?.id).toBe('mini-office.theme-classic');
  });

  it('拒绝非反向域名格式的 id', () => {
    const result = validateManifest({ ...VALID_MANIFEST, id: 'classic' });
    expect(result.ok).toBe(false);
    expect(result.issues.some((issue) => issue.path === 'id')).toBe(true);
  });

  it('拒绝声明式插件携带权限', () => {
    const result = validateManifest({ ...VALID_MANIFEST, permissions: ['document:read'] });
    expect(result.ok).toBe(false);
    expect(result.issues.some((issue) => issue.message.includes('权限必须为空'))).toBe(true);
  });

  it('拒绝没有任何贡献点的插件', () => {
    const result = validateManifest({ ...VALID_MANIFEST, contributes: {} });
    expect(result.ok).toBe(false);
    expect(result.issues.some((issue) => issue.path === 'contributes')).toBe(true);
  });

  it('宿主版本不满足 engines 时拒绝加载', () => {
    const result = validateManifest(VALID_MANIFEST, '0.6.0');
    expect(result.ok).toBe(false);
    expect(result.issues.some((issue) => issue.path === 'engines.miniOffice')).toBe(true);
  });
});

describe('semver 范围', () => {
  it('支持 caret 只允许同一 major', () => {
    expect(satisfiesRange('1.4.0', '^1.0.0')).toBe(true);
    expect(satisfiesRange('2.0.0', '^1.0.0')).toBe(false);
  });

  it('支持上下界组合', () => {
    expect(satisfiesRange('0.1.5', '>=0.1.5 <0.5.0')).toBe(true);
    expect(satisfiesRange('0.5.0', '>=0.1.5 <0.5.0')).toBe(false);
  });
});

describe('主题令牌校验', () => {
  const theme = VALID_MANIFEST.contributes.themes[0];

  it('接受完整且对比度达标的主题', () => {
    const result = validateTheme(theme, 0);
    expect(result.ok).toBe(true);
    expect(result.theme?.tokens['--mo-bg']).toBe('#ECE9D8');
  });

  it('拒绝修改文档画布令牌', () => {
    const result = validateTheme(
      { ...theme, tokens: { ...theme.tokens, '--mo-doc-font': 'Tahoma' } },
      0,
    );
    expect(result.ok).toBe(false);
    expect(result.issues.some((issue) => issue.message.includes('文档画布'))).toBe(true);
  });

  it('拒绝未知令牌', () => {
    const result = validateTheme({ ...theme, tokens: { ...theme.tokens, '--mo-nope': '#fff' } }, 0);
    expect(result.ok).toBe(false);
    expect(result.issues.some((issue) => issue.message.includes('未知'))).toBe(true);
  });

  it('拒绝缺少必填令牌的主题', () => {
    const result = validateTheme({ ...theme, tokens: { '--mo-bg': '#ffffff' } }, 0);
    expect(result.ok).toBe(false);
    expect(result.issues.filter((issue) => issue.message === '缺少必填令牌').length).toBeGreaterThan(0);
  });

  it('拒绝正文与背景对比度过低（装上会看不清）', () => {
    const result = validateTheme(
      {
        ...theme,
        tokens: { ...theme.tokens, '--mo-surface': '#F5F5F5', '--mo-text-primary': '#DDDDDD' },
      },
      0,
    );
    expect(result.ok).toBe(false);
    expect(result.issues.some((issue) => issue.message.includes('对比度'))).toBe(true);
  });

  it('拒绝非法的 chrome 变体取值', () => {
    const result = validateTheme({ ...theme, chrome: { controlShape: 'wavy' } }, 0);
    expect(result.ok).toBe(false);
    expect(result.issues.some((issue) => issue.path === 'chrome.controlShape')).toBe(true);
  });
});

describe('对比度计算', () => {
  it('黑白对比度为 21:1', () => {
    const black = parseHexColor('#000000');
    const white = parseHexColor('#ffffff');
    expect(black && white && contrastRatio(black, white)).toBeCloseTo(21, 1);
  });

  it('非颜色值返回 null，不参与对比度检查', () => {
    expect(parseHexColor('0px')).toBeNull();
    expect(parseHexColor('var(--x)')).toBeNull();
  });
});

describe('插件索引解析', () => {
  const entry = {
    id: 'mini-office.theme-classic',
    name: '经典外观',
    version: '1.0.0',
    publisher: 'mini-office',
    tier: 'declarative',
    permissions: [],
    download: 'https://example.com/a.moext',
    sha256: 'a'.repeat(64),
  };

  it('接受合法索引', () => {
    const index = parsePluginIndex({ schemaVersion: 1, plugins: [entry] });
    expect(index?.plugins).toHaveLength(1);
  });

  it('丢弃缺少 sha256 的条目，避免安装未校验的包', () => {
    const index = parsePluginIndex({ schemaVersion: 1, plugins: [{ ...entry, sha256: 'abc' }] });
    expect(index?.plugins).toHaveLength(0);
  });

  it('拒绝未知 schemaVersion', () => {
    expect(parsePluginIndex({ schemaVersion: 2, plugins: [] })).toBeNull();
  });
});

describe('插件包解析与安装', () => {
  it('解析带顶层目录的包并剥离前缀', async () => {
    const zip = new JSZip();
    zip.file('pkg/manifest.json', JSON.stringify(VALID_MANIFEST));
    zip.file('pkg/preview.png', 'fake');
    const bytes = await zip.generateAsync({ type: 'uint8array' });

    const parsed = await parsePackage(bytes, '0.1.5');
    expect(parsed.manifest.id).toBe('mini-office.theme-classic');
    expect([...parsed.files.keys()].sort()).toEqual(['manifest.json', 'preview.png']);
  });

  it('拒绝路径穿越、绝对路径与反斜杠条目', () => {
    expect(isSafeEntryPath('manifest.json')).toBe(true);
    expect(isSafeEntryPath('assets/icon.png')).toBe(true);
    expect(isSafeEntryPath('../evil.txt')).toBe(false);
    expect(isSafeEntryPath('assets/../../evil.txt')).toBe(false);
    expect(isSafeEntryPath('/etc/passwd')).toBe(false);
    expect(isSafeEntryPath('assets\\icon.png')).toBe(false);
  });

  it('拒绝缺少清单的包', async () => {
    const zip = new JSZip();
    zip.file('readme.txt', 'hi');
    const bytes = await zip.generateAsync({ type: 'uint8array' });
    await expect(parsePackage(bytes, '0.1.5')).rejects.toThrow(/缺少清单/);
  });

  it('安装后文件写入存储并生成安装记录', async () => {
    const store = createMemoryStore();
    const bytes = await buildPackage(VALID_MANIFEST, { 'preview.png': 'fake' });
    const { plugin } = await installFromBytes(store, bytes, 'local', { currentAppVersion: '0.1.5' });
    expect(plugin.id).toBe('mini-office.theme-classic');
    expect(plugin.enabled).toBe(true);
    expect(await store.list(plugin.id)).toEqual(['manifest.json', 'preview.png']);
  });
});

describe('插件仓库中的官方主题', () => {
  const repoRoot = resolve(__dirname, '../plugin-repo/plugins');

  it('每个官方主题包都能通过清单与主题校验', async () => {
    const { readdir } = await import('node:fs/promises');
    const dirs = (await readdir(repoRoot, { withFileTypes: true }))
      .filter((entry) => entry.isDirectory())
      .map((entry) => entry.name);
    expect(dirs.length).toBeGreaterThanOrEqual(3);

    for (const dir of dirs) {
      const raw = await readFile(resolve(repoRoot, dir, 'manifest.json'), 'utf8');
      const manifestResult = validateManifest(JSON.parse(raw), '0.1.5');
      expect(manifestResult.issues, `${dir} 清单`).toEqual([]);

      // 功能插件（宏 + 模板）没有主题，只校验清单即可。
      const themes = manifestResult.manifest?.contributes.themes ?? [];
      if (themes.length === 0) continue;
      for (const [index, theme] of themes.entries()) {
        const themeResult = validateTheme(theme, index);
        expect(themeResult.issues, `${dir} 主题 ${index}`).toEqual([]);
      }
    }
  });

  it('主题插件必须真的带主题，功能插件必须带命令或模板', async () => {
    const { readdir } = await import('node:fs/promises');
    const dirs = (await readdir(repoRoot, { withFileTypes: true }))
      .filter((entry) => entry.isDirectory())
      .map((entry) => entry.name);
    for (const dir of dirs) {
      const raw = await readFile(resolve(repoRoot, dir, 'manifest.json'), 'utf8');
      const manifest = validateManifest(JSON.parse(raw), '0.1.5').manifest;
      expect(manifest, `${dir} 清单`).toBeDefined();
      if (!manifest) continue;
      const total =
        (manifest.contributes.themes?.length ?? 0) +
        (manifest.contributes.commands?.length ?? 0) +
        (manifest.contributes.templates?.length ?? 0);
      expect(total, `${dir} 贡献点`).toBeGreaterThan(0);
    }
  });

  it('packages/ 里的每个包都能通过完整解析（含模板文件存在性）', async () => {
    const { readdir } = await import('node:fs/promises');
    const packagesDir = resolve(__dirname, '../plugin-repo/packages');
    const files = (await readdir(packagesDir)).filter((name) => name.endsWith('.moext'));
    expect(files.length).toBeGreaterThanOrEqual(3);
    for (const file of files) {
      const bytes = new Uint8Array(await readFile(resolve(packagesDir, file)));
      const parsed = await parsePackage(bytes, '0.1.5');
      expect(parsed.manifest.id, file).toBeTruthy();
      for (const template of parsed.manifest.contributes.templates ?? []) {
        expect(parsed.files.has(template.file), `${file} 模板 ${template.file}`).toBe(true);
      }
    }
  });

  it('打包产物可重现：所有条目时间戳固定为 1980，且不含目录条目', async () => {
    // 这是真实踩过的坑：JSZip 默认会给 `templates/` 这类父目录补一个目录条目，
    // 时间用的是当前时间而不是固定值。单文件插件不会触发，带子目录的插件每次
    // 打包哈希都不一样，索引里的校验值就和仓库里的包对不上。
    const { readdir } = await import('node:fs/promises');
    const packagesDir = resolve(__dirname, '../plugin-repo/packages');
    const files = (await readdir(packagesDir)).filter((name) => name.endsWith('.moext'));
    for (const file of files) {
      const zip = await JSZip.loadAsync(await readFile(resolve(packagesDir, file)));
      for (const entry of Object.values(zip.files)) {
        expect(entry.dir, `${file} 不应包含目录条目 ${entry.name}`).toBe(false);
        expect(entry.date.getUTCFullYear(), `${file} ${entry.name} 时间戳`).toBe(1980);
      }
    }
  });

  it('经典外观必须声明 beveled + compact，否则只是换了配色', async () => {
    const raw = await readFile(resolve(repoRoot, 'theme-office-2003/manifest.json'), 'utf8');
    const result = validateManifest(JSON.parse(raw), '0.1.5');
    expect(result.manifest?.contributes.themes?.[0]?.chrome).toMatchObject({
      controlShape: 'beveled',
      density: 'compact',
      toolbar: 'classic-two-row',
    });
  });

  it('插件仓库的 index.json 与官方主题一一对应且校验值合法', async () => {
    const raw = await readFile(resolve(__dirname, '../plugin-repo/index.json'), 'utf8');
    const index = parsePluginIndex(JSON.parse(raw));
    expect(index).not.toBeNull();
    const ids = (index?.plugins ?? []).map((entry) => entry.id).sort();
    expect(ids).toEqual([
      'mini-office.pack-gov-doc',
      'mini-office.theme-office-2003',
      'mini-office.theme-office-365',
      'mini-office.theme-wps',
      'mini-office.theme-wps-ribbon',
    ]);
    for (const entry of index?.plugins ?? []) {
      expect(entry.sha256).toMatch(/^[0-9a-f]{64}$/);
      expect(entry.download).toMatch(/^https:\/\//);
    }
  });
});

describe('主题包应用', () => {
  it('注入 style 元素并写入 chrome 属性，回退时清理干净', () => {
    const theme = VALID_MANIFEST.contributes.themes[0] as never;
    const key = themePackKey('mini-office.theme-classic', 'classic');

    applyThemePack({ key, theme }, 'light');
    expect(document.documentElement.dataset.themePack).toBe(key);
    const style = document.getElementById('mo-theme-pack');
    expect(style?.textContent).toContain('[data-theme="light"]');
    expect(style?.textContent).toContain('--mo-bg: #ECE9D8');

    applyThemePack(null, 'light');
    expect(document.getElementById('mo-theme-pack')).toBeNull();
    expect(document.documentElement.dataset.themePack).toBeUndefined();
  });

  it('chrome 变体写入 data-chrome-* 属性', () => {
    const theme = {
      ...VALID_MANIFEST.contributes.themes[0],
      chrome: { controlShape: 'beveled', density: 'compact' },
    } as never;
    applyThemePack({ key: 'p#t', theme }, 'light');
    expect(document.documentElement.getAttribute('data-chrome-control-shape')).toBe('beveled');
    expect(document.documentElement.getAttribute('data-chrome-density')).toBe('compact');
    applyThemePack(null, 'light');
    expect(document.documentElement.getAttribute('data-chrome-control-shape')).toBeNull();
  });
});
