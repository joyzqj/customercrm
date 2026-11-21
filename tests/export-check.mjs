/**
 * 客户通 CustomerCRM — 浏览器端 Excel 下载联调检查（Lead 用）
 * 运行：node tests/export-check.mjs
 */
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { serveStatic, launchChrome, connectToPage, PageSession, sleep } from './cdp.mjs';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const downloadDir = fs.mkdtempSync(path.join(os.tmpdir(), 'crm-dl-'));

const server = await serveStatic(ROOT);
const chrome = await launchChrome();
const { client } = await connectToPage(chrome);
const page = await PageSession.create(client);
const failures = [];
function check(name, cond, detail) {
  console.log(`${cond ? 'PASS' : 'FAIL'}  ${name}${cond ? '' : '  → ' + detail}`);
  if (!cond) failures.push(name);
}

try {
  await page.setDownloadBehavior(downloadDir);
  await page.navigate(server.url + '/index.html');
  await page.waitUntil('!!window.CustomerCRM', { timeout: 20000 });
  await page.appNavigate('#/customers');
  await page.waitUntil('document.querySelectorAll("tr.row").length > 0', { timeout: 10000 });
  const res = await page.evaluate('window.CustomerCRM.app.exportExcel()');
  console.log('export result =', JSON.stringify(res));
  check('exportExcel 返回文件名', res && /^客户跟进台账_\d{8}_\d{4}\.xlsx$/.test(res.filename || ''), JSON.stringify(res));
  check('exportExcel 返回字节数 > 3000', res && res.bytes > 3000, res && res.bytes);

  let file = null;
  for (let i = 0; i < 60; i += 1) {
    const files = fs.readdirSync(downloadDir).filter((f) => f.endsWith('.xlsx'));
    if (files.length) { file = path.join(downloadDir, files[0]); break; }
    await sleep(100);
  }
  check('浏览器下载到 xlsx 文件', !!file, downloadDir);
  if (file) {
    const buf = fs.readFileSync(file);
    console.log('downloaded =', path.basename(file), buf.length, 'bytes');
    check('文件头 PK', buf[0] === 0x50 && buf[1] === 0x4b, buf.subarray(0, 4).toString('hex'));
    check('文件大小与返回值一致', buf.length === res.bytes, `${buf.length} vs ${res.bytes}`);
    check('zip 中含 workbook.xml', buf.includes(Buffer.from('xl/workbook.xml')), 'missing');
  }
  const errors = page.errors();
  check('无 console 错误', errors.length === 0, JSON.stringify(errors.slice(0, 3)));
} finally {
  page.close();
  await chrome.close();
  await server.close();
  try { fs.rmSync(downloadDir, { recursive: true, force: true }); } catch { /* ignore */ }
}

console.log(failures.length ? `\n失败 ${failures.length} 项` : '\n全部通过');
process.exit(failures.length ? 1 : 0);
