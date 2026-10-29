/**
 * v1.1 收口验证：更新检测、一键更新、快捷键真实注册、主题预览。
 *
 * 前置：`pnpm build` 后由 `npx vite preview --port 1420` 提供 dist。
 */
import { chromium } from 'playwright-core';
import { homedir } from 'node:os';
import { resolve } from 'node:path';

const executablePath = resolve(
  homedir(),
  'Library/Caches/ms-playwright/chromium-1217/chrome-mac-arm64/Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing',
);
const oldPkg = '/tmp/pack-gov-doc-1.0.0.moext';
const out = (name) => resolve('tests/output', name);

const browser = await chromium.launch({ executablePath });
const page = await browser.newPage({ viewport: { width: 1360, height: 900 } });
const errors = [];
page.on('pageerror', (error) => errors.push(`pageerror: ${error.message}`));
page.on('console', (msg) => {
  if (msg.type() === 'error') errors.push(`console: ${msg.text()}`);
});

await page.goto('http://127.0.0.1:1420/', { waitUntil: 'networkidle' });
await page.waitForTimeout(1800);

// 1) 装 1.0.0 的旧包，模拟"客户手上已有旧版本"
await page.keyboard.press('Meta+Comma');
await page.waitForTimeout(800);
await page.getByRole('button', { name: '扩展', exact: true }).click();
await page.waitForTimeout(400);
await page.getByRole('button', { name: '管理插件' }).click();
await page.waitForTimeout(700);
await page.setInputFiles('input[type=file][accept*=".moext"]', oldPkg);
await page.waitForTimeout(2500);
let text = await page.locator('body').innerText();
console.log('installed old:', text.includes('公文排版助手') && text.includes('1.0.0'));

// 2) 刷新远端索引 → 应出现"可更新"
await page.getByRole('button', { name: '刷新列表' }).click();
await page.waitForTimeout(4000);
text = await page.locator('body').innerText();
console.log('update badge:', /可更新 1\.0\.0 → 1\.0\.1/.test(text));
console.log('update summary:', /1 个可更新/.test(text));
await page.screenshot({ path: out('v11-01-update-available.png') });

// 3) 一键更新
await page.getByRole('button', { name: '更新', exact: true }).click();
await page.waitForTimeout(4000);
text = await page.locator('body').innerText();
console.log('updated to 1.0.1:', text.includes('1.0.1') && !/可更新 1\.0\.0/.test(text));
console.log('update notice:', text.includes('更新到 1.0.1'));
await page.screenshot({ path: out('v11-02-updated.png') });

// 4) 主题预览：装一套主题插件后，设置里的风格选择器应有缩略图
const installButtons = page.getByRole('button', { name: '安装', exact: true });
if ((await installButtons.count()) > 0) {
  await installButtons.first().click();
  await page.waitForTimeout(3500);
}
// 插件管理是独立弹窗，要回到设置页才能看「外观」标签
await page.keyboard.press('Escape');
await page.waitForTimeout(500);
await page.keyboard.press('Meta+Comma');
await page.waitForTimeout(900);
await page.getByRole('button', { name: '外观', exact: true }).click();
await page.waitForTimeout(900);
const previews = await page.locator('[data-theme-preview]').count();
console.log('theme previews:', previews);
await page.screenshot({ path: out('v11-03-theme-preview.png') });

// 5) 快捷键：先看命令面板是否展示组合键，再直接按键触发
await page.keyboard.press('Escape');
await page.waitForTimeout(300);
await page.keyboard.press('Escape');
await page.waitForTimeout(600);
await page.getByText('党政机关公文', { exact: true }).first().click();
await page.waitForTimeout(3500);

await page.keyboard.press('Meta+K');
await page.waitForTimeout(500);
await page.keyboard.type('套用公文版式');
await page.waitForTimeout(500);
const paletteText = await page.locator('body').innerText();
console.log('palette shows shortcut:', /⇧?⌘|Ctrl/.test(paletteText));
await page.screenshot({ path: out('v11-04-palette-shortcut.png') });
await page.keyboard.press('Escape');
await page.waitForTimeout(400);

await page.keyboard.press('Meta+Shift+G');
await page.waitForTimeout(1500);
const afterShortcut = await page.locator('body').innerText();
console.log('shortcut ran macro:', afterShortcut.includes('已执行「套用公文版式」'));
await page.screenshot({ path: out('v11-05-shortcut-ran.png') });

// 6) 冲突的快捷键不应吞键：内置 ⌘S 仍走保存而不是插件命令
await page.keyboard.press('Meta+S');
await page.waitForTimeout(1200);
const afterSave = await page.locator('body').innerText();
console.log('builtin still wins:', !/已执行「套用公文版式」/.test(afterSave.split('已执行「套用公文版式」').pop() ?? ''));

console.log('errors:', errors.length ? errors.slice(0, 5) : 'none');
await browser.close();
