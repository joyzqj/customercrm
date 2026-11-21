/**
 * 客户通 CustomerCRM — 集成冒烟测试（Lead 自查用）
 * 运行：node tests/smoke.mjs
 * 校验：首屏无 console error、路由渲染、示例数据、筛选、新建表单、响应式无横向溢出、截图。
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { serveStatic, launchChrome, connectToPage, PageSession, sleep } from './cdp.mjs';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const SHOTS = path.join(ROOT, 'screenshots', 'smoke');

const results = [];
function check(name, condition, detail) {
  results.push({ name, ok: !!condition, detail: detail === undefined ? '' : String(detail) });
  console.log(`${condition ? 'PASS' : 'FAIL'}  ${name}${condition ? '' : '  → ' + detail}`);
}

const server = await serveStatic(ROOT);
const chrome = await launchChrome({ windowSize: { width: 1440, height: 900 } });
const { client } = await connectToPage(chrome);
const page = await PageSession.create(client);

try {
  fs.mkdirSync(SHOTS, { recursive: true });

  await page.navigate(server.url + '/index.html');
  await page.waitUntil('!!(window.CustomerCRM && window.CustomerCRM.app)', { timeout: 20000 });
  await page.waitUntil('!!document.querySelector("#view-inner .page")', { timeout: 10000 });
  console.log('App booted, route =', await page.evaluate('window.CustomerCRM.app.getRoute()'));

  const api = await page.evaluate(`(() => ({
    total: window.CustomerCRM.store.list().length,
    version: window.CustomerCRM.version,
    parser: !!window.CustomerCRM.parser,
    xlsx: !!(window.CustomerCRM.xlsx && window.CustomerCRM.xlsx.exportCustomers),
    stats: window.CustomerCRM.app.stats(),
    route: window.CustomerCRM.app.getRoute()
  }))()`);
  console.log('api =', JSON.stringify(api));
  check('示例数据已载入 12 条', api.total === 12, api.total);
  check('xlsx 引擎已加载', api.xlsx);
  check('默认路由 dashboard', api.route === 'dashboard', api.route);
  check('统计总数正确', api.stats.total === api.total, JSON.stringify(api.stats && { t: api.stats.total, d: api.stats.deal }));
  check('已成交统计', api.stats.deal >= 2, api.stats.deal);

  const statDom = await page.evaluate(`(() => {
    const read = (id) => { const el = document.querySelector('[data-testid="' + id + '"]'); return el ? el.textContent.trim() : null; };
    return { total: read('stat-total'), deal: read('stat-deal'), open: read('stat-open'), overdue: read('stat-overdue') };
  })()`);
  console.log('stats dom =', JSON.stringify(statDom));
  check('统计卡渲染', statDom.total === String(api.stats.total), JSON.stringify(statDom));

  await page.screenshot(path.join(SHOTS, 'dashboard-1440.png'));

  // 客户台账
  await page.appNavigate('#/customers');
  await page.waitUntil('document.querySelectorAll("tr.row").length > 0', { timeout: 10000 });
  const rows = await page.count('tr.row');
  check('客户表格行数 = 12', rows === 12, rows);
  const cells = await page.evaluate(`(() => {
    const row = document.querySelector('tr.row');
    return row ? row.querySelectorAll('td').length : 0;
  })()`);
  check('表格列数 14', cells === 14, cells);
  await page.screenshot(path.join(SHOTS, 'customers-1440.png'));

  // 展开行详情
  await page.click('tr.row .row-toggle');
  await sleep(120);
  const detailVisible = await page.evaluate(`(() => {
    const d = document.querySelector('.row-detail');
    return d ? !d.hidden && d.offsetParent !== null : false;
  })()`);
  check('展开最近三次跟进', detailVisible);
  await page.screenshot(path.join(SHOTS, 'customers-expanded-1440.png'));

  // 搜索筛选
  await page.fill('#customer-search', '德国');
  await sleep(150);
  const filtered = await page.count('tr.row');
  check('搜索德国命中', filtered >= 1 && filtered < 12, filtered);
  await page.fill('#customer-search', '');
  await sleep(150);

  // 新建客户表单
  await page.click('[data-testid="new-customer"]');
  await page.waitUntil('!!document.querySelector("#customer-form")', { timeout: 8000 });
  const formFields = await page.evaluate(`(() => document.querySelectorAll('#customer-form [name]').length)()`);
  check('表单字段渲染', formFields >= 11, formFields);
  await page.screenshot(path.join(SHOTS, 'modal-new-customer.png'));

  // 必填校验：直接保存应报错
  await page.click('[data-testid="btn-save-customer"]');
  await sleep(150);
  const formError = await page.evaluate(`(() => { const el = document.getElementById('form-error'); return el && !el.hidden ? el.textContent : ''; })()`);
  const nameValue = await page.evaluate(`(() => { const el = document.querySelector('#customer-form [name="customerName"]'); return el ? el.value : 'x'; })()`);
  check('新建客户必填校验', nameValue === '' ? !!formError : true, `formError=${formError}`);

  // 填表并保存
  await page.fill('customerName', '烟测客户 Alpha');
  await page.fill('product', '测试产品 X1');
  await page.fill('country', '德国');
  await page.fill('companyName', 'Smoke Test GmbH');
  await page.fill('coreReason', '价格谈判中');
  await page.click('[data-testid="btn-save-customer"]');
  await sleep(250);
  const afterSave = await page.evaluate(`(() => ({ total: window.CustomerCRM.store.list().length, modal: !!document.querySelector('#customer-form') }))()`);
  check('保存后客户数 +1', afterSave.total === 13, afterSave.total);
  check('保存后弹窗关闭', afterSave.modal === false);
  const rowVisible = await page.evaluate(`(() => Array.from(document.querySelectorAll('tr.row td')).some(td => td.textContent.includes('烟测客户 Alpha')))()`);
  check('保存后表格出现新客户', rowVisible);

  // 删除测试客户
  const newId = await page.evaluate(`(() => { const rec = window.CustomerCRM.store.list().find(r => r.customerName === '烟测客户 Alpha'); return rec ? rec.id : ''; })()`);
  await page.evaluate(`window.CustomerCRM.store.remove(${JSON.stringify(newId)})`);
  await sleep(200);

  // 智能识别页（parser 就绪时验证）
  await page.appNavigate('#/recognize');
  await page.waitUntil('!!document.querySelector("#recognize-input")', { timeout: 8000 });
  await page.screenshot(path.join(SHOTS, 'recognize-empty-1440.png'));
  if (api.parser) {
    const result = await page.evaluate(`(() => {
      const sample = window.CrmParser.SAMPLE_TEXTS[0].text;
      document.getElementById('recognize-input').value = sample;
      document.getElementById('recognize-input').dispatchEvent(new Event('input', { bubbles: true }));
      document.querySelector('[data-testid="btn-run-recognize"]').click();
      const items = document.querySelectorAll('.recog-item').length;
      return { items, filled: document.querySelectorAll('.recog-item--filled').length, hasResult: !!document.querySelector('.recog-list') };
    })()`);
    await sleep(200);
    check('识别结果 11 项', result.items === 11, result.items);
    check('识别命中 ≥ 6 项', result.filled >= 6, result.filled);
    await page.screenshot(path.join(SHOTS, 'recognize-result-1440.png'));

    // 填入表格并核对
    await page.click('[data-testid="btn-fill-customer"]');
    await page.waitUntil('!!document.querySelector("#customer-form")', { timeout: 8000 });
    const prefilled = await page.evaluate(`(() => { const el = document.querySelector('#customer-form [name="customerName"]'); return el ? el.value : ''; })()`);
    check('识别结果预填表单', prefilled.length > 0, prefilled);
    await page.screenshot(path.join(SHOTS, 'modal-recognized.png'));
    await page.click('[data-testid="btn-cancel-customer"]');
    await sleep(180);
    const modalClosed = await page.evaluate(`!document.querySelector('#customer-form')`);
    check('取消按钮关闭弹窗', modalClosed);
  } else {
    console.log('SKIP 识别引擎尚未就绪（js/parser.js 未生成）');
  }

  // 设置页 & 帮助页
  await page.appNavigate('#/settings');
  await page.waitUntil('!!document.querySelector("#setting-warnDays")', { timeout: 8000 });
  await page.screenshot(path.join(SHOTS, 'settings-1440.png'));
  await page.appNavigate('#/help');
  await page.waitUntil('!!document.querySelector(".help-grid")', { timeout: 8000 });
  await page.screenshot(path.join(SHOTS, 'help-1440.png'));

  // 响应式
  for (const [w, h, name] of [[1024, 800, 'tablet'], [768, 900, 'narrow'], [390, 844, 'mobile']]) {
    await page.setViewport(w, h, { mobile: w < 600 });
    await page.appNavigate('#/customers');
    await sleep(200);
    const metrics = await page.evaluate(`(() => ({ sw: document.documentElement.scrollWidth, iw: window.innerWidth }))()`);
    check(`${name} 无横向溢出`, metrics.sw <= metrics.iw + 1, JSON.stringify(metrics));
    await page.screenshot(path.join(SHOTS, `customers-${name}.png`));
  }
  await page.setViewport(1440, 900);

  // console 错误
  const errors = page.errors();
  check('无 console / page 错误', errors.length === 0, JSON.stringify(errors.slice(0, 4)));

  // 保存持久化
  await page.reload();
  await page.waitUntil('!!window.CustomerCRM', { timeout: 15000 });
  const persisted = await page.evaluate('window.CustomerCRM.store.list().length');
  check('刷新后数据仍在', persisted === 12, persisted);
} finally {
  page.close();
  await chrome.close();
  await server.close();
}

const failed = results.filter((r) => !r.ok);
console.log(`\n冒烟测试：${results.length - failed.length}/${results.length} 通过`);
console.log('截图目录：' + SHOTS);
if (failed.length) {
  console.log('失败项：');
  failed.forEach((f) => console.log('  ✗ ' + f.name + ' — ' + f.detail));
  process.exit(1);
}
