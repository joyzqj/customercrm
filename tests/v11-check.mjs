/**
 * 客户通 CustomerCRM — v1.1 功能自检（Lead 用，不写正式截图）
 * 运行：node tests/v11-check.mjs
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
const failures = [];
function check(name, cond, detail) {
  console.log(`${cond ? 'PASS' : 'FAIL'}  ${name}${cond ? '' : '  → ' + detail}`);
  if (!cond) failures.push(name);
}

try {
  await page.setViewport(1440, 900);
  await page.navigate(server.url + '/index.html');
  await page.waitUntil('!!window.CustomerCRM', { timeout: 20000 });
  await page.evaluate('localStorage.clear(); location.reload()');
  await page.waitUntil('!!window.CustomerCRM', { timeout: 20000 });
  await sleep(400);

  const meta = await page.evaluate(`(() => ({
    version: window.CustomerCRM.version,
    fields: window.CRM_SCHEMA.FIELDS.length,
    products: window.CustomerCRM.store.listProducts().length,
    names: window.CustomerCRM.store.productNames().length,
    firstProduct: window.CustomerCRM.store.listProducts()[0]
  }))()`);
  console.log('meta =', JSON.stringify(meta));
  check('版本 1.2.0', meta.version === '1.2.0', meta.version);
  check('字段 16 个', meta.fields === 16, meta.fields);
  check('产品库 12 条', meta.products === 12, meta.products);

  await page.appNavigate('#/dashboard');
  await sleep(250);
  await page.screenshot(path.join(OUT, 'v11-dashboard-pie.png'));
  const pie = await page.evaluate(`(() => {
    const slices = Array.from(document.querySelectorAll('[data-testid="country-pie"] .pie-slice'));
    return slices.map(s => ({ country: s.getAttribute('data-country'), fill: s.getAttribute('fill'), d: (s.getAttribute('d') || '').slice(0, 40) }));
  })()`);
  console.log('slices =', JSON.stringify(pie));
  check('扇形切片 ≥1 且都带路径', pie.length >= 1 && pie.every((s) => s.d && s.d.length > 10), pie.length);

  await page.appNavigate('#/customers');
  await page.waitUntil('document.querySelectorAll("tr.row").length > 0', { timeout: 10000 });
  await sleep(200);
  await page.screenshot(path.join(OUT, 'v11-customers-wide.png'));
  const layout = await page.evaluate(`(() => {
    const inner = document.querySelector('.view-inner');
    const table = document.querySelector('.data-table--customers');
    const business = document.querySelector('th[data-col="mainBusiness"]');
    const productCell = document.querySelector('td[data-col="product"]');
    return {
      bodyClass: document.body.className,
      innerWidth: inner ? Math.round(inner.getBoundingClientRect().width) : 0,
      tableWidth: table ? Math.round(table.getBoundingClientRect().width) : 0,
      hasBusiness: !!business,
      productHtml: productCell ? productCell.innerHTML.slice(0, 120) : '',
      cols: document.querySelectorAll('tr.row:first-of-type td').length
    };
  })()`);
  console.log('layout =', JSON.stringify(layout));
  check('台账宽版 body class', /route-customers/.test(layout.bodyClass), layout.bodyClass);
  check('台账内容宽度 ≥1140', layout.innerWidth >= 1140, layout.innerWidth);
  check('主营业务列存在', layout.hasBusiness);
  check('产品型号小字渲染', /cell-sub/.test(layout.productHtml), layout.productHtml);

  // 新建客户：产品联动
  await page.click('[data-testid="new-customer"]');
  await page.waitUntil('!!document.querySelector("#customer-form")', { timeout: 8000 });
  await page.screenshot(path.join(OUT, 'v11-form-product.png'));
  const sync = await page.evaluate(`(() => {
    const form = document.getElementById('customer-form');
    const nameSelect = form.querySelector('[name="productName"]');
    const target = Array.from(nameSelect.options).find(o => o.value === '不锈钢保温杯');
    nameSelect.value = '不锈钢保温杯';
    nameSelect.dispatchEvent(new Event('change', { bubbles: true }));
    const modelSelect = form.querySelector('[name="productModel"]');
    const before = modelSelect ? Array.from(modelSelect.options).map(o => o.value) : [];
    modelSelect.value = '750ml';
    modelSelect.dispatchEvent(new Event('change', { bubbles: true }));
    return { before, current: modelSelect.value };
  })()`);
  console.log('product sync =', JSON.stringify(sync));
  check('型号下拉联动 ≥3 个型号', sync.before.length >= 4, sync.before.length);
  check('可选中具体型号', sync.current === '750ml', sync.current);
  await page.click('[data-testid="btn-cancel-customer"]');
  await sleep(150);

  // 设置页产品库
  await page.appNavigate('#/settings');
  await page.waitUntil('!!document.querySelector("#product-form")', { timeout: 8000 });
  await sleep(200);
  await page.screenshot(path.join(OUT, 'v11-settings-products.png'));
  const lib = await page.evaluate(`(() => ({
    rows: document.querySelectorAll('[data-testid="product-row"]').length,
    hasForm: !!document.querySelector('#product-form'),
    usage: document.querySelectorAll('.product-usage').length
  }))()`);
  console.log('product library =', JSON.stringify(lib));
  check('产品库列表 12 行', lib.rows === 12, lib.rows);

  // 新增产品
  await page.fill('productNewName', '测试新品');
  await page.fill('productNewModel', 'X1');
  await page.click('[data-testid="btn-add-product"]');
  await sleep(300);
  const afterAdd = await page.evaluate(`(() => ({ rows: document.querySelectorAll('[data-testid="product-row"]').length, names: window.CustomerCRM.store.productNames() }))()`);
  check('产品库新增成功', afterAdd.rows === 13 && afterAdd.names.indexOf('测试新品') !== -1, JSON.stringify(afterAdd.rows));
  const newProduct = await page.evaluate(`(() => window.CustomerCRM.store.listProducts().find(p => p.name === '测试新品'))()`);
  await page.evaluate(`window.CustomerCRM.store.removeProduct(${JSON.stringify(newProduct && newProduct.id)})`);
  await sleep(150);

  const errors = page.errors();
  check('无 console 错误', errors.length === 0, JSON.stringify(errors.slice(0, 3)));
} finally {
  page.close();
  await chrome.close();
  await server.close();
}

console.log(failures.length ? `\nv1.1 自检失败 ${failures.length} 项` : '\nv1.1 自检全部通过');
process.exit(failures.length ? 1 : 0);
