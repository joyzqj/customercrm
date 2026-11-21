/**
 * 客户通 CustomerCRM — 细节视觉巡检（Lead 用，输出到 screenshots/review/）
 * 运行：node tests/visual-review.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { serveStatic, launchChrome, connectToPage, PageSession, sleep } from './cdp.mjs';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const OUT = path.join(ROOT, 'screenshots', 'review');
fs.mkdirSync(OUT, { recursive: true });

const server = await serveStatic(ROOT);
const chrome = await launchChrome({ windowSize: { width: 1440, height: 900 } });
const { client } = await connectToPage(chrome);
const page = await PageSession.create(client);

async function shot(name) {
  await page.screenshot(path.join(OUT, name));
  console.log('shot', name);
}

try {
  await page.setViewport(1440, 900);
  await page.navigate(server.url + '/index.html');
  await page.waitUntil('!!window.CustomerCRM', { timeout: 20000 });
  await page.evaluate('localStorage.clear(); location.reload()');
  await page.waitUntil('!!window.CustomerCRM', { timeout: 20000 });
  await sleep(400);

  await page.appNavigate('#/dashboard');
  await sleep(200);
  await page.evaluate('window.scrollTo(0, 520)');
  await sleep(200);
  await shot('dashboard-lower.png');

  await page.appNavigate('#/settings');
  await sleep(200);
  await page.evaluate('window.scrollTo(0, 460)');
  await sleep(200);
  await shot('settings-lower.png');

  await page.appNavigate('#/customers');
  await page.waitUntil('document.querySelectorAll("tr.row").length > 0', { timeout: 10000 });
  await page.click('tr.row .row-toggle');
  await sleep(200);
  await page.evaluate('window.scrollTo(0, 300)');
  await sleep(150);
  await shot('customers-expanded.png');

  await page.appNavigate('#/recognize');
  await sleep(200);
  await page.click('[data-action="recognize-sample"][data-index="0"]').catch(() => {});
  await sleep(600);
  await page.evaluate('window.scrollTo(0, 320)');
  await sleep(150);
  await shot('recognize-lower.png');

  await page.emulateMedia('print');
  await page.appNavigate('#/customers');
  await sleep(300);
  await shot('customers-print.png');
  await page.emulateMedia('');
  await page.setViewport(390, 844, { mobile: true });
  await page.appNavigate('#/dashboard');
  await sleep(250);
  await page.evaluate('window.scrollTo(0, 400)');
  await sleep(200);
  await shot('mobile-dashboard-lower.png');

  console.log('errors:', JSON.stringify(page.errors().slice(0, 5)));
} finally {
  page.close();
  await chrome.close();
  await server.close();
}
