/**
 * 客户通 CustomerCRM — file:// 离线打开验证（Lead 用）
 * 运行：node tests/file-check.mjs
 * 验证「双击 index.html」与「dist 单文件版」两种离线用法。
 */
import fs from 'node:fs';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { fileURLToPath } from 'node:url';
import { launchChrome, connectToPage, PageSession, sleep } from './cdp.mjs';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const targets = [
  { name: 'index.html（多文件版）', file: path.join(ROOT, 'index.html') },
  { name: 'dist 单文件版', file: path.join(ROOT, 'dist', '客户通-外贸客户管理系统.html') }
];

const chrome = await launchChrome();
const failures = [];
function check(name, cond, detail) {
  console.log(`${cond ? 'PASS' : 'FAIL'}  ${name}${cond ? '' : '  → ' + detail}`);
  if (!cond) failures.push(name);
}

for (const target of targets) {
  if (!fs.existsSync(target.file)) {
    check(target.name + ' 存在', false, target.file);
    continue;
  }
  const { client } = await connectToPage(chrome);
  const page = await PageSession.create(client);
  try {
    const url = pathToFileURL(target.file).href;
    await page.navigate(url, { timeout: 30000 });
    await page.waitUntil('!!(window.CustomerCRM && window.CustomerCRM.app)', { timeout: 20000 });
    await page.waitUntil('!!document.querySelector("#view-inner .page")', { timeout: 10000 });
    const info = await page.evaluate(`(() => ({
      total: window.CustomerCRM.store.list().length,
      route: window.CustomerCRM.app.getRoute(),
      table: document.querySelectorAll('tr.row').length,
      xlsx: !!(window.CustomerCRM.xlsx && window.CustomerCRM.xlsx.exportCustomers),
      parser: !!window.CustomerCRM.parser
    }))()`);
    console.log(target.name, '→', JSON.stringify(info));
    check(target.name + ' 离线可用', info.total === 12 && info.xlsx && info.parser, JSON.stringify(info));
    const errors = page.errors().filter((e) => !/favicon/i.test(e.text));
    check(target.name + ' 无 console 错误', errors.length === 0, JSON.stringify(errors.slice(0, 3)));
    // 离线导出
    const res = await page.evaluate('window.CustomerCRM.app.exportExcel()');
    check(target.name + ' 离线导出 Excel', res && res.bytes > 3000, JSON.stringify(res));
    // file:// 下 localStorage 持久化
    await page.reload();
    await page.waitUntil('!!window.CustomerCRM', { timeout: 15000 });
    await sleep(200);
    const after = await page.evaluate('window.CustomerCRM.store.list().length');
    check(target.name + ' 刷新后数据仍在', after === 12, after);
  } finally {
    page.close();
  }
}

await chrome.close();
console.log(failures.length ? `\n离线校验失败 ${failures.length} 项` : '\n离线校验全部通过（file:// 双击可用）');
process.exit(failures.length ? 1 : 0);
