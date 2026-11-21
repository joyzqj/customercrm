/**
 * 客户通 CustomerCRM — 线上部署验证（Lead 用）
 * 运行：node tests/live-check.mjs [url]
 * 默认检查 https://joyzqj.github.io/customercrm/
 */
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { launchChrome, connectToPage, PageSession, sleep } from './cdp.mjs';

const url = process.argv[2] || 'https://joyzqj.github.io/customercrm/';
const downloadDir = fs.mkdtempSync(path.join(os.tmpdir(), 'crm-live-'));
const failures = [];
function check(name, cond, detail) {
  console.log(`${cond ? 'PASS' : 'FAIL'}  ${name}${cond ? '' : '  → ' + detail}`);
  if (!cond) failures.push(name);
}

const chrome = await launchChrome();
const { client } = await connectToPage(chrome);
const page = await PageSession.create(client);

try {
  await page.setDownloadBehavior(downloadDir);
  await page.navigate(url, { timeout: 60000 });
  await page.waitUntil('!!(window.CustomerCRM && window.CustomerCRM.app)', { timeout: 30000 });
  await page.waitUntil('!!document.querySelector("#view-inner .page")', { timeout: 15000 });
  const info = await page.evaluate(`(() => ({
    title: document.title,
    total: window.CustomerCRM.store.list().length,
    route: window.CustomerCRM.app.getRoute(),
    parser: !!window.CustomerCRM.parser,
    xlsx: !!window.CustomerCRM.xlsx,
    version: window.CustomerCRM.version
  }))()`);
  console.log('online =', JSON.stringify(info));
  check('线上 JS 加载正常', info.version === '1.0.0', JSON.stringify(info));
  check('示例数据 12 条', info.total === 12, info.total);
  check('识别引擎在线可用', info.parser === true);
  check('Excel 引擎在线可用', info.xlsx === true);

  // 巡检所有路由
  for (const route of ['#/dashboard', '#/customers', '#/recognize', '#/settings', '#/help']) {
    await page.appNavigate(route);
    await sleep(120);
    check(`路由 ${route} 渲染`, await page.evaluate('!!document.querySelector("#view-inner .page")'));
  }

  // 线上导出 Excel
  await page.appNavigate('#/customers');
  await page.waitUntil('document.querySelectorAll("tr.row").length > 0', { timeout: 10000 });
  const res = await page.evaluate('window.CustomerCRM.app.exportExcel()');
  check('线上导出返回文件', res && res.bytes > 3000, JSON.stringify(res));
  let file = null;
  for (let i = 0; i < 50; i += 1) {
    const files = fs.readdirSync(downloadDir).filter((f) => f.endsWith('.xlsx'));
    if (files.length) { file = path.join(downloadDir, files[0]); break; }
    await sleep(100);
  }
  check('线上下载落盘', !!file, downloadDir);

  // 线上识别
  const recog = await page.evaluate(`(() => {
    const r = window.CustomerCRM.app.recognize(window.CrmParser.SAMPLE_TEXTS[0].text);
    return r ? { filled: r.filled, name: r.fields.customerName, country: r.fields.country } : null;
  })()`);
  console.log('recognize =', JSON.stringify(recog));
  check('线上识别样例 A', recog && recog.filled >= 9 && recog.name === 'Michael Braun', JSON.stringify(recog));

  check('线上无 console 错误', page.errors().length === 0, JSON.stringify(page.errors().slice(0, 3)));
  await page.screenshot(path.join('screenshots', 'live-home.png'));
} finally {
  page.close();
  await chrome.close();
  try { fs.rmSync(downloadDir, { recursive: true, force: true }); } catch { /* ignore */ }
}

console.log(failures.length ? `\n线上校验失败 ${failures.length} 项` : '\n线上校验全部通过：' + url);
process.exit(failures.length ? 1 : 0);
