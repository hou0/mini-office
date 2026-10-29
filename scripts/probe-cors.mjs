import { chromium } from 'playwright-core';
import { homedir } from 'node:os';
import { resolve } from 'node:path';
const executablePath = resolve(homedir(), 'Library/Caches/ms-playwright/chromium-1217/chrome-mac-arm64/Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing');
const browser = await chromium.launch({ executablePath });
const page = await browser.newPage();
await page.goto('http://127.0.0.1:1420/', { waitUntil: 'domcontentloaded' });
const targets = {
  'raw main index': 'https://raw.githubusercontent.com/lq0910/mini-office-plugins/main/index.json',
  'raw tag v1 manifest': 'https://raw.githubusercontent.com/lq0910/mini-office-plugins/v1/plugins/theme-wps/manifest.json',
  'release asset': 'https://github.com/lq0910/mini-office-plugins/releases/download/v1/mini-office.theme-wps-1.0.0.moext',
  'api contents': 'https://api.github.com/repos/lq0910/mini-office-plugins/contents/index.json?ref=main',
};
for (const [label, url] of Object.entries(targets)) {
  const result = await page.evaluate(async (u) => {
    try {
      const r = await fetch(u, { redirect: 'follow' });
      const buf = await r.arrayBuffer();
      return { ok: r.ok, status: r.status, bytes: buf.byteLength };
    } catch (e) {
      return { ok: false, error: String(e).slice(0, 90) };
    }
  }, url);
  console.log(label, '->', JSON.stringify(result));
}
await browser.close();
