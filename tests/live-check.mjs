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
  check('线上 JS 加载正常', info.version === '1.2.0', JSON.stringify(info));
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
    return r ? {
      filled: r.filled, total: r.missing.length + r.filled,
      name: r.fields.customerName, country: r.fields.country,
      model: r.fields.productModel, biz: r.fields.mainBusiness,
      email: r.fields.email, phone: r.fields.phone, address: r.fields.address
    } : null;
  })()`);
  console.log('recognize =', JSON.stringify(recog));
  check('线上识别样例 A', recog && recog.filled >= 15 && recog.name === 'Michael Braun', JSON.stringify(recog));
  check('线上识别产品型号/主营业务', recog && /500ml/i.test(recog.model || '') && /户外用品/.test(recog.biz || ''), JSON.stringify(recog));
  check('线上识别邮箱/电话/地址', recog && /braun-trading\.de$/.test(recog.email || '') && /\+49/.test(recog.phone || '') && /Stuttgart/.test(recog.address || ''), JSON.stringify(recog));

  // v1.1：扇形图 / 宽版台账 / 产品库
  await page.appNavigate('#/dashboard');
  await page.waitForApp();
  await sleep(300);
  const pie = await page.evaluate(`(() => {
    const box = document.querySelector('[data-testid="country-pie"]');
    return box ? {
      slices: box.querySelectorAll('.pie-slice').length,
      legend: box.querySelectorAll('.legend-item').length,
      pct: Array.from(box.querySelectorAll('.legend-pct')).reduce((a, el) => a + parseInt(el.textContent, 10), 0)
    } : null;
  })()`);
  console.log('online pie =', JSON.stringify(pie));
  check('线上扇形图渲染', !!pie && pie.slices > 0 && pie.slices === pie.legend, JSON.stringify(pie));
  check('线上扇形图百分比合计 100', !!pie && pie.pct === 100, pie && pie.pct);

  await page.appNavigate('#/customers');
  await page.waitUntil('document.querySelectorAll("tr.row").length > 0', { timeout: 10000 });
  const wide = await page.evaluate(`(() => {
    const inner = document.querySelector('.view-inner');
    const table = document.querySelector('.data-table--customers');
    const firstRow = document.querySelector('tr.row');
    const contact = firstRow ? firstRow.querySelector('.cell-contact') : null;
    const addressHead = document.querySelector('th[data-col="address"]');
    return {
      bodyClass: document.body.className,
      innerWidth: Math.round(inner.getBoundingClientRect().width),
      tableWidth: Math.round(table.getBoundingClientRect().width),
      cols: firstRow ? firstRow.querySelectorAll('td').length : 0,
      contact: contact ? contact.textContent.trim() : '',
      hasAddressCol: !!addressHead,
      products: window.CustomerCRM.store.listProducts().length
    };
  })()`);
  console.log('online layout =', JSON.stringify(wide));
  check('线上宽版台账生效', /route-customers/.test(wide.bodyClass) && wide.innerWidth >= 1120, JSON.stringify(wide));
  check('线上表格 16 列', wide.cols === 16, wide.cols);
  check('线上名称单元格显示邮箱/电话', /@/.test(wide.contact || ''), wide.contact);
  check('线上存在地址列', wide.hasAddressCol === true);
  check('线上产品库 12 条', wide.products === 12, wide.products);

  // 展开行联系方式链接
  await page.click('tr.row .row-toggle');
  await sleep(200);
  const contactLinks = await page.evaluate(`(() => {
    const meta = document.querySelector('.row-detail:not([hidden]) .row-detail-meta');
    if (!meta) return null;
    const mail = meta.querySelector('a[href^="mailto:"]');
    const tel = meta.querySelector('a[href^="tel:"]');
    return { mail: mail ? mail.getAttribute('href') : '', tel: tel ? tel.getAttribute('href') : '', text: meta.textContent.slice(0, 120) };
  })()`);
  console.log('online contact links =', JSON.stringify(contactLinks));
  check('线上展开行有 mailto 链接', !!contactLinks && /mailto:/.test(contactLinks.mail), JSON.stringify(contactLinks));
  check('线上展开行有 tel 链接', !!contactLinks && /tel:\+?\d/.test(contactLinks.tel), JSON.stringify(contactLinks));

  check('线上无 console 错误', page.errors().length === 0, JSON.stringify(page.errors().slice(0, 3)));
  await page.screenshot(path.join('screenshots', 'live-home.png'));
} finally {
  page.close();
  await chrome.close();
  try { fs.rmSync(downloadDir, { recursive: true, force: true }); } catch { /* ignore */ }
}

console.log(failures.length ? `\n线上校验失败 ${failures.length} 项` : '\n线上校验全部通过：' + url);
process.exit(failures.length ? 1 : 0);
