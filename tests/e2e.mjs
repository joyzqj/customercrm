/**
 * 客户通 CustomerCRM — 端到端验收（E2E）
 *
 * 运行：node tests/e2e.mjs
 *   全绿退出码 0；任一用例失败打印失败用例与详情并退出码 1。
 *
 * 技术：零 npm 依赖，复用 tests/cdp.mjs（随机端口静态服务 + headless Chrome CDP）。
 * 数据：每次运行/每组用例开头 localStorage.clear() + reload，自动重建 12 条示例数据，
 *       保证可重复、组间隔离。
 *
 * 覆盖（CONTRACT v1.1 §8/§9 + task-8）：
 *   A 首屏 / 仪表盘 4 统计卡 / 国家分布扇形图 / 无 console error / 截图
 *   B 客户台账 12 行 15 列 / 搜索（含产品型号、主营业务）/ 筛选 / 排序 / 展开行
 *   C 新建（13 字段、产品名称+型号联动、必填校验、持久化）/ 编辑 / 跟进 / 删除
 *   P 产品库（设置页 12 行、新增/删除/恢复预置）与产品选择器联动
 *   D 键盘快捷键 / 聚焦搜索、N 新建、Esc 关闭弹窗
 *   E 智能识别（13 项、filled≥11、产品型号/主营业务、填入表单、直接新增一行、空文本兜底）
 *   F 一键导出 Excel（真实下载、文件名、PK 头、ZIP 部件、openpyxl 两个工作表 / 21+11 列 / 新列有值）
 *   G 设置（阈值持久化与仪表盘联动、JSON 备份导出/恢复含产品库、清空后空状态、载入示例数据）
 *   H 响应式 1440/1024/768/390 无横向溢出、宽版台账、侧栏抽屉、打印媒体隐藏与恢复、移动端截图
 *   I 数据损坏兜底（非法 JSON / v1.0 旧 payload 升级）
 *   S 8 张验收截图齐备、PNG 尺寸正确
 */
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import zlib from 'node:zlib';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { serveStatic, launchChrome, connectToPage, PageSession, sleep } from './cdp.mjs';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
/** 调试期可用 E2E_SHOTS_DIR 指向临时目录，避免与 ui-dev 抢 screenshots/ 根目录；最终验收不设置即写正式目录 */
const SHOTS = process.env.E2E_SHOTS_DIR ? path.resolve(process.env.E2E_SHOTS_DIR) : path.join(ROOT, 'screenshots');
const TEMP = fs.mkdtempSync(path.join(os.tmpdir(), 'customer-crm-e2e-'));
const DOWNLOADS = path.join(TEMP, 'downloads');
fs.mkdirSync(DOWNLOADS, { recursive: true });

/* ================================================================== */
/* 常量                                                                */
/* ================================================================== */

/** CONTRACT v1.1 §6.1：客户总表 21 列（顺序固定） */
const EXPECTED_TOTAL_HEADERS = [
  '序号', '首次询盘时间', '客户名称', '咨询产品', '产品型号', '国家', '公司名', '主营业务',
  '成立时间', '团队规模', '年营收', '跟进状态', '是否成交', '核心原因', '距上次跟进(天)',
  '最近跟进①日期', '最近跟进①内容', '最近跟进②日期', '最近跟进②内容', '最近跟进③日期', '最近跟进③内容'
];

/** CONTRACT v1.1 §6.1：跟进明细 11 列（顺序固定） */
const EXPECTED_DETAIL_HEADERS = [
  '序号', '首次询盘时间', '客户名称', '公司名', '主营业务', '国家',
  '跟进日期', '距今天数', '跟进内容', '跟进状态', '是否成交'
];

/** README 引用的 8 张验收截图（1440×900 基准 + 390×844 移动端） */
const SHOT_SPECS = [
  { file: 'dashboard.png', width: 1440, height: 900 },
  { file: 'customers.png', width: 1440, height: 900 },
  { file: 'customers-expanded.png', width: 1440, height: 900 },
  { file: 'recognize.png', width: 1440, height: 900 },
  { file: 'form.png', width: 1440, height: 900 },
  { file: 'settings.png', width: 1440, height: 900 },
  { file: 'mobile-dashboard.png', width: 390, height: 844 },
  { file: 'mobile-customers.png', width: 390, height: 844 }
];

/** openpyxl 校验脚本（由 Python 3.12 + openpyxl 3.1.5 执行） */
const PY_XLSX_VALIDATOR = `# -*- coding: utf-8 -*-
import sys, json
from openpyxl import load_workbook

path = sys.argv[1]
wb = load_workbook(path)
out = {"sheetnames": list(wb.sheetnames), "sheets": {}}
for name in wb.sheetnames:
    ws = wb[name]
    header = []
    for cell in next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ()):
        header.append("" if cell is None else str(cell))
    names = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row is None:
            continue
        names.append("" if len(row) < 3 or row[2] is None else str(row[2]))
    first_row = []
    for row in ws.iter_rows(min_row=2, max_row=2, values_only=True):
        for cell in row:
            first_row.append("" if cell is None else str(cell))
        break
    out["sheets"][name] = {
        "max_row": ws.max_row,
        "max_column": ws.max_column,
        "header": header,
        "names": names,
        "first_row": first_row,
        "freeze": ws.freeze_panes,
        "auto_filter": ws.auto_filter.ref if ws.auto_filter is not None else None,
    }
sys.stdout.write("\\n" + json.dumps(out, ensure_ascii=True) + "\\n")
`;

/* ================================================================== */
/* 断言与用例收集                                                       */
/* ================================================================== */

const results = [];
let group = { id: 'P', title: '准备' };
let python = null;
const scratch = {};

const J = (value) => JSON.stringify(value);

function formatValue(value) {
  if (typeof value === 'string') return value.length > 500 ? value.slice(0, 500) + '…' : value;
  try {
    const text = JSON.stringify(value);
    return text !== undefined && text.length > 500 ? text.slice(0, 500) + '…' : String(text);
  } catch {
    return String(value);
  }
}

function setGroup(id, title) {
  group = { id, title };
  console.log(`\n── ${id} ${title} ──`);
}

async function test(id, title, fn) {
  const started = Date.now();
  try {
    await fn();
    const ms = Date.now() - started;
    results.push({ id, group: group.id, title, ok: true, detail: '', ms });
    console.log(`  PASS  ${id}  ${title}  (${ms}ms)`);
  } catch (err) {
    const ms = Date.now() - started;
    const detail = (err && err.__detail !== undefined) ? String(err.__detail) : ((err && err.message) || String(err));
    results.push({ id, group: group.id, title, ok: false, detail, ms });
    console.log(`  FAIL  ${id}  ${title}  (${ms}ms)`);
    console.log('        ' + detail.split('\n').join('\n        '));
  }
}

function assert(condition, message, detail) {
  if (condition) return;
  const err = new Error(message);
  err.__detail = detail === undefined
    ? message
    : `${message} → ${formatValue(detail)}`;
  throw err;
}

function assertEq(actual, expected, message) {
  if (actual === expected) return;
  const err = new Error(message);
  err.__detail = `${message} → 实际: ${formatValue(actual)}；期望: ${formatValue(expected)}`;
  throw err;
}

function fail(message, detail) {
  assert(false, message, detail);
}

/* ================================================================== */
/* 页面工具                                                             */
/* ================================================================== */

async function waitExpr(page, expression, label, timeout = 10000) {
  try {
    return await page.waitUntil(expression, { timeout, label });
  } catch (err) {
    throw Object.assign(new Error(`等待超时：${label}`), {
      __detail: `等待超时（${timeout}ms）：${label}\n表达式：${expression}`
    });
  }
}

async function waitSelector(page, selector, label, timeout = 10000) {
  return waitExpr(page, `!!document.querySelector(${J(selector)})`, label || `元素出现 ${selector}`, timeout);
}

async function waitGone(page, selector, label, timeout = 10000) {
  return waitExpr(page, `!document.querySelector(${J(selector)})`, label || `元素消失 ${selector}`, timeout);
}

async function click(page, selector) {
  const ok = await page.click(selector);
  assert(ok, '点击失败：元素不存在', selector);
  return ok;
}

async function goto(page, hash) {
  await page.appNavigate(hash);
  await waitExpr(page, `!!document.querySelector('#view-inner .page')`, `路由 ${hash} 渲染完成`, 10000);
}

async function rowCount(page) {
  return page.count('tr[data-testid="customer-row"]');
}

async function waitRowCount(page, expected, timeout = 10000) {
  return waitExpr(
    page,
    `document.querySelectorAll('tr[data-testid="customer-row"]').length === ${Number(expected)}`,
    `客户表格行数 = ${expected}`,
    timeout
  );
}

async function waitStoreCount(page, expected, timeout = 10000) {
  return waitExpr(
    page,
    `window.CustomerCRM.store.list().length === ${Number(expected)}`,
    `store 客户记录数 = ${expected}`,
    timeout
  );
}

async function waitHash(page, hash, timeout = 10000) {
  return waitExpr(page, `location.hash === ${J(hash)}`, `路由 hash = ${hash}`, timeout);
}

async function clearToasts(page) {
  await page.evaluate(`document.querySelectorAll('#toast-stack .toast').forEach(function (node) { node.remove(); });`);
}

async function waitToast(page, kind, contains, timeout = 6000) {
  const deadline = Date.now() + timeout;
  let seen = [];
  while (Date.now() < deadline) {
    seen = await page.evaluate(
      `Array.from(document.querySelectorAll('#toast-stack .toast--${kind}')).map(function (node) { return node.textContent || ''; })`
    );
    if (seen.some((text) => text.indexOf(contains) !== -1)) return seen;
    await sleep(80);
  }
  fail(`未出现期望的 ${kind} Toast（含「${contains}」）`, seen);
}

async function readStatsDom(page) {
  return page.evaluate(`(() => {
    const read = function (id) {
      const el = document.querySelector('[data-testid="' + id + '"]');
      return el ? el.textContent.trim() : null;
    };
    return { total: read('stat-total'), deal: read('stat-deal'), open: read('stat-open'), overdue: read('stat-overdue') };
  })()`);
}

async function readStatHints(page) {
  return page.evaluate(`(() => {
    const out = {};
    document.querySelectorAll('.stat-card').forEach(function (card) {
      const value = card.querySelector('[data-testid]');
      const hint = card.querySelector('.stat-hint');
      if (value && hint) out[value.getAttribute('data-testid')] = hint.textContent.trim();
    });
    return out;
  })()`);
}

async function readRowIds(page) {
  return page.evaluate(`Array.from(document.querySelectorAll('tr[data-testid="customer-row"]')).map(function (tr) { return tr.getAttribute('data-id'); })`);
}

async function readCol(page, col) {
  return page.evaluate(
    `Array.from(document.querySelectorAll('tr[data-testid="customer-row"] td[data-col="${col}"]')).map(function (td) { return td.textContent.trim(); })`
  );
}

async function formValues(page) {
  return page.evaluate(`(() => {
    const form = document.getElementById('customer-form');
    if (!form) return null;
    const out = {};
    Array.from(form.querySelectorAll('[name]')).forEach(function (el) { out[el.getAttribute('name')] = el.value; });
    return out;
  })()`);
}

async function filteredIds(page) {
  return page.evaluate(`window.CustomerCRM.app.filtered().map(function (record) { return record.id; })`);
}

async function filteredNames(page) {
  return page.evaluate(`window.CustomerCRM.app.filtered().map(function (record) { return record.customerName; })`);
}

/** 选择下拉框并触发 change（v1.1 产品名称/型号联动） */
async function selectOption(page, selector, value) {
  const result = await page.evaluate(`(() => {
    const el = document.querySelector(${J(selector)});
    if (!el) return { ok: false, error: '元素不存在' };
    const options = Array.from(el.options || []).map(function (option) { return option.value; });
    if (options.indexOf(${J(value)}) === -1) return { ok: false, error: '选项不存在', options: options };
    const proto = Object.getPrototypeOf(el);
    const descriptor = Object.getOwnPropertyDescriptor(proto, 'value');
    if (descriptor && descriptor.set) descriptor.set.call(el, ${J(value)});
    else el.value = ${J(value)};
    for (const type of ['input', 'change']) el.dispatchEvent(new Event(type, { bubbles: true }));
    return { ok: true, value: el.value };
  })()`);
  assert(result.ok, '下拉选择失败：' + selector, result);
  return result.value;
}

/** 读取仪表盘国家分布扇形图状态 */
async function readPie(page) {
  return page.evaluate(`(() => {
    const root = document.querySelector('[data-testid="country-pie"]');
    if (!root) return null;
    const slices = Array.from(root.querySelectorAll('.pie-slice'));
    const legend = Array.from(root.querySelectorAll('.pie-legend .legend-item'));
    const pctTexts = Array.from(root.querySelectorAll('.legend-pct')).map(function (el) { return el.textContent.trim(); });
    return {
      slices: slices.length,
      sliceCountries: slices.map(function (el) { return el.getAttribute('data-country'); }),
      slicePaths: slices.map(function (el) { const d = el.getAttribute('d') || ''; return d.length; }),
      sliceTitles: slices.map(function (el) { const title = el.querySelector('title'); return title ? title.textContent : ''; }),
      legend: legend.length,
      legendCountries: legend.map(function (el) { return el.getAttribute('data-country'); }),
      pctTexts: pctTexts,
      pctSum: pctTexts.reduce(function (acc, text) { return acc + (parseInt(text, 10) || 0); }, 0),
      total: (root.querySelector('.pie-total') || {}).textContent || '',
      svg: !!root.querySelector('.pie-svg')
    };
  })()`);
}

/** 每组用例开头：清空 localStorage + 刷新，等待 12 条示例数据重建 */
async function resetDemo(page) {
  await page.evaluate('localStorage.clear()');
  await page.reload();
  await page.waitForApp({ timeout: 25000 });
  await waitExpr(page, `!!document.querySelector('#view-inner .page')`, '重置后页面渲染', 20000);
  await waitStoreCount(page, 12, 15000);
  page.clearErrors();
}

async function assertNoOverflow(page, label) {
  const metrics = await page.layoutMetrics();
  assert(
    metrics.docScrollWidth <= metrics.innerWidth + 1,
    `${label} 出现整页横向溢出（documentElement.scrollWidth > innerWidth）`,
    metrics
  );
  assert(
    metrics.bodyScrollWidth <= metrics.bodyClientWidth + 1,
    `${label} body 出现横向溢出（bodyScrollWidth > bodyClientWidth）`,
    metrics
  );
  return metrics;
}

async function sidebarRect(page) {
  return page.evaluate(`(() => {
    const el = document.querySelector('.sidebar');
    if (!el) return null;
    const rect = el.getBoundingClientRect();
    const style = getComputedStyle(el);
    return { left: rect.left, right: rect.right, width: rect.width, display: style.display, visibility: style.visibility, transform: style.transform };
  })()`);
}

async function columnVisible(page, key) {
  return page.evaluate(
    `(() => { const th = document.querySelector('th[data-col="${key}"]'); return th ? getComputedStyle(th).display !== 'none' : null; })()`
  );
}

async function assertNoErrors(page, id, title) {
  await test(id, title, async () => {
    const errors = page.errors();
    assert(errors.length === 0, '存在 console error / pageerror / log.error', errors.slice(0, 6));
  });
}

/* ---------------- 下载 / XLSX / 截图 ---------------- */

function clearDownloads() {
  try {
    for (const name of fs.readdirSync(DOWNLOADS)) fs.rmSync(path.join(DOWNLOADS, name), { force: true, recursive: true });
  } catch { /* ignore */ }
}

async function waitForDownload(extensionRe, timeout = 30000) {
  const deadline = Date.now() + timeout;
  while (Date.now() < deadline) {
    let files = [];
    try { files = fs.readdirSync(DOWNLOADS); } catch { files = []; }
    const candidates = files.filter((name) => extensionRe.test(name) && !name.endsWith('.crdownload'));
    for (const name of candidates) {
      const filePath = path.join(DOWNLOADS, name);
      let first = 0;
      let second = 0;
      try { first = fs.statSync(filePath).size; } catch { continue; }
      await sleep(150);
      try { second = fs.statSync(filePath).size; } catch { continue; }
      if (first > 0 && first === second) return filePath;
    }
    await sleep(120);
  }
  let existing = [];
  try { existing = fs.readdirSync(DOWNLOADS); } catch { existing = []; }
  fail('等待下载文件超时：' + extensionRe, existing);
}

/** 解析 ZIP（xlsx 为 stored 方式，同时兼容 deflate），返回 Map<文件名, Buffer> */
function readZip(buffer) {
  assert(buffer.length > 22, 'ZIP 文件过小', buffer.length);
  let eocd = -1;
  for (let i = buffer.length - 22; i >= Math.max(0, buffer.length - 22 - 65535); i -= 1) {
    if (buffer.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  assert(eocd >= 0, 'ZIP 缺少 EOCD 结束记录（文件可能未写完/损坏）');
  const count = buffer.readUInt16LE(eocd + 10);
  const cdOffset = buffer.readUInt32LE(eocd + 16);
  const entries = new Map();
  let cursor = cdOffset;
  for (let i = 0; i < count; i += 1) {
    assert(buffer.readUInt32LE(cursor) === 0x02014b50, 'ZIP 中央目录损坏', { index: i, offset: cursor });
    const method = buffer.readUInt16LE(cursor + 10);
    const compressedSize = buffer.readUInt32LE(cursor + 20);
    const nameLength = buffer.readUInt16LE(cursor + 28);
    const extraLength = buffer.readUInt16LE(cursor + 30);
    const commentLength = buffer.readUInt16LE(cursor + 32);
    const localOffset = buffer.readUInt32LE(cursor + 42);
    const name = buffer.subarray(cursor + 46, cursor + 46 + nameLength).toString('utf8');
    assert(buffer.readUInt32LE(localOffset) === 0x04034b50, 'ZIP 本地文件头损坏', name);
    const localNameLength = buffer.readUInt16LE(localOffset + 26);
    const localExtraLength = buffer.readUInt16LE(localOffset + 28);
    const dataStart = localOffset + 30 + localNameLength + localExtraLength;
    const raw = buffer.subarray(dataStart, dataStart + compressedSize);
    let data;
    if (method === 0) data = Buffer.from(raw);
    else if (method === 8) data = zlib.inflateRawSync(raw);
    else fail('不支持的 ZIP 压缩方式', { name, method });
    entries.set(name, data);
    cursor += 46 + nameLength + extraLength + commentLength;
  }
  return entries;
}

function countXmlRows(xmlText) {
  return (xmlText.match(/<row[\s>]/g) || []).length;
}

function pngSize(filePath) {
  const buffer = fs.readFileSync(filePath);
  if (buffer.length < 24 || buffer.readUInt32BE(0) !== 0x89504e47) return null;
  return { width: buffer.readUInt32BE(16), height: buffer.readUInt32BE(20), bytes: buffer.length };
}

function assetScreenshot(file, expectedWidth, expectedHeight, minBytes = 20000) {
  const filePath = path.join(SHOTS, file);
  assert(fs.existsSync(filePath), '截图文件不存在', filePath);
  const size = pngSize(filePath);
  assert(size, '截图不是合法 PNG', filePath);
  assert(size.bytes >= minBytes, '截图体积过小，可能是空白页', { file, bytes: size.bytes });
  assertEq(size.width, expectedWidth, `截图 ${file} 宽度`);
  assertEq(size.height, expectedHeight, `截图 ${file} 高度`);
  return size;
}

async function takeScreenshot(page, file) {
  await clearToasts(page);
  await sleep(120);
  await page.screenshot(path.join(SHOTS, file));
}

/* ---------------- Python / openpyxl ---------------- */

function findPython() {
  const probes = [
    { command: process.env.CRM_PYTHON, args: [] },
    {
      command: 'C:\\Users\\Administrator\\AppData\\Local\\Programs\\DSH Desktop\\resources\\office-runtime\\primary-runtime\\dependencies\\python\\python.exe',
      args: []
    },
    { command: 'python', args: [] },
    { command: 'python3', args: [] },
    { command: 'py', args: ['-3'] }
  ].filter((probe) => probe.command);
  for (const probe of probes) {
    const result = spawnSync(probe.command, [...probe.args, '-c', 'import openpyxl,sys; sys.stdout.write(openpyxl.__version__)'], {
      encoding: 'utf8',
      timeout: 30000,
      windowsHide: true
    });
    if (result && result.status === 0) {
      const version = String(result.stdout || '').trim();
      if (version) return { command: probe.command, args: probe.args, version };
    }
  }
  return null;
}

function runPythonXlsxCheck(filePath) {
  if (!python) return null;
  const scriptPath = path.join(TEMP, 'validate_xlsx.py');
  fs.writeFileSync(scriptPath, PY_XLSX_VALIDATOR, 'utf8');
  const result = spawnSync(python.command, [...python.args, scriptPath, filePath], {
    encoding: 'utf8',
    timeout: 60000,
    windowsHide: true,
    // Windows 下 Python 管道 stdout 默认 cp936，会让中文 JSON 乱码；强制 UTF-8
    env: Object.assign({}, process.env, { PYTHONIOENCODING: 'utf-8', PYTHONUTF8: '1' })
  });
  const stderrTail = String(result.stderr || '').trim().slice(-800);
  assert(result.status === 0, 'openpyxl 读取 xlsx 失败（文件可能损坏）', stderrTail || String(result.stdout || '').slice(-400));
  const lines = String(result.stdout || '').trim().split(/\r?\n/).filter(Boolean);
  const last = lines[lines.length - 1] || '';
  let parsed;
  try {
    parsed = JSON.parse(last);
  } catch {
    fail('openpyxl 校验脚本输出无法解析为 JSON', last.slice(0, 400));
  }
  return parsed;
}

/* ---------------- 键盘 ---------------- */

const KEY_DEFS = {
  '/': { key: '/', code: 'Slash', vk: 191 },
  n: { key: 'n', code: 'KeyN', vk: 78 },
  Escape: { key: 'Escape', code: 'Escape', vk: 27 }
};

async function pressKey(page, keyName) {
  const def = KEY_DEFS[keyName];
  assert(def, '未定义的按键', keyName);
  await page.client.send('Input.dispatchKeyEvent', {
    type: 'keyDown',
    key: def.key,
    code: def.code,
    windowsVirtualKeyCode: def.vk,
    nativeVirtualKeyCode: def.vk
  });
  await page.client.send('Input.dispatchKeyEvent', {
    type: 'keyUp',
    key: def.key,
    code: def.code,
    windowsVirtualKeyCode: def.vk,
    nativeVirtualKeyCode: def.vk
  });
}

function localDateStamp() {
  const now = new Date();
  const pad = (value) => (value < 10 ? '0' + value : String(value));
  return String(now.getFullYear()) + pad(now.getMonth() + 1) + pad(now.getDate());
}

/* ================================================================== */
/* A. 首屏与仪表盘                                                     */
/* ================================================================== */

async function groupA(page) {
  setGroup('A', '首屏启动与仪表盘');

  await test('A1', '首屏启动：#/dashboard 渲染，导航高亮、标题正确', async () => {
    await waitHash(page, '#/dashboard', 15000);
    assertEq(await page.evaluate('window.CustomerCRM.app.getRoute()'), 'dashboard', '默认路由应为 dashboard');
    assert(await page.exists('#view-inner .page'), '页面主体 #view-inner .page 未渲染');
    assert(await page.exists('.nav-item.is-active[data-nav="dashboard"]'), '侧栏「仪表盘」未高亮');
    const title = await page.title();
    assert(String(title).indexOf('仪表盘') !== -1, 'document.title 应包含「仪表盘」', title);
    assert(await page.exists('[data-testid="stat-total"]'), '缺少统计卡 stat-total');
  });

  await test('A2', '首屏无 console error / pageerror / log.error', async () => {
    const errors = page.errors();
    assert(errors.length === 0, '首屏存在错误级控制台输出', errors.slice(0, 6));
  });

  await test('A3', '首次访问自动写入 12 条示例数据，13 个字段完整（含产品型号/主营业务）', async () => {
    const meta = await page.evaluate(`(() => ({
      version: window.CustomerCRM.version,
      fields: window.CRM_SCHEMA.FIELDS.length,
      keys: window.CRM_SCHEMA.FIELD_KEYS.slice(),
      products: window.CustomerCRM.store.listProducts().length,
      records: window.CustomerCRM.store.list()
    }))()`);
    assertEq(meta.version, '1.1.0', 'CustomerCRM.version');
    assertEq(meta.fields, 13, 'CRM_SCHEMA.FIELDS 字段数');
    assertEq(meta.keys.length, 13, 'CRM_SCHEMA.FIELD_KEYS 字段数');
    assert(meta.keys.indexOf('productModel') !== -1, '缺少 productModel 字段', meta.keys);
    assert(meta.keys.indexOf('mainBusiness') !== -1, '缺少 mainBusiness 字段', meta.keys);
    assertEq(meta.products, 12, '首次访问应写入 12 条预置产品');
    const records = meta.records;
    assertEq(records.length, 12, '示例数据条数');
    for (const record of records) {
      for (const key of meta.keys) {
        assert(typeof record[key] === 'string', `客户「${record.customerName}」字段 ${key} 不是字符串`, record[key]);
      }
      assert(record.customerName.length > 0, '示例客户缺少客户名称', record.id);
      assert(record.productModel.length > 0, '示例客户缺少产品型号', record.customerName);
      assert(record.mainBusiness.length > 0, '示例客户缺少主营业务', record.customerName);
    }
    const names = records.map((record) => record.customerName);
    assert(names.indexOf('Michael Braun') !== -1, '示例数据应包含契约样例客户 Michael Braun', names);
    assertEq(records.filter((record) => record.customerName === 'Michael Braun')[0].productModel, '500ml', 'Michael Braun 产品型号');
  });

  await test('A4', '仪表盘 4 个统计卡数值 = CustomerCRM.app.stats()', async () => {
    const stats = await page.evaluate('window.CustomerCRM.app.stats()');
    const dom = await readStatsDom(page);
    assertEq(dom.total, String(stats.total), '客户总数卡');
    assertEq(dom.deal, String(stats.deal), '已成交卡');
    assertEq(dom.open, String(stats.open), '待跟进卡');
    assertEq(dom.overdue, String(stats.overdue), '超期未跟进卡');
    assertEq(Number(dom.total), 12, '客户总数应为 12');
    assert(stats.deal >= 1, '示例数据应至少 1 位已成交客户', stats.deal);
    assert(dom.overdue !== null && dom.open !== null, '统计卡未渲染', dom);
  });

  await test('A5', '仪表盘区块渲染：4 统计卡 / 国家分布 / 状态分布 / 最近动态 / 优先跟进', async () => {
    const layout = await page.evaluate(`(() => ({
      statCards: document.querySelectorAll('.stats-grid .stat-card').length,
      barRows: document.querySelectorAll('.bar-row').length,
      timelineItems: document.querySelectorAll('.timeline-item').length,
      priorityItems: document.querySelectorAll('.priority-item').length,
      pageSub: (document.querySelector('.page-sub') || {}).textContent || ''
    }))()`);
    assertEq(layout.statCards, 4, '统计卡数量');
    assert(layout.barRows >= 4, '状态分布条形图条目过少', layout.barRows);
    assertEq(layout.timelineItems, 8, '最近跟进动态应展示 8 条');
    assert(layout.pageSub.indexOf('共 12 位客户') !== -1, '页面摘要应显示 12 位客户', layout.pageSub);
    const stats = await page.evaluate('window.CustomerCRM.app.stats()');
    assertEq(layout.priorityItems, stats.overdueList.length, '优先跟进条数 = stats.overdueList');
  });

  await test('A6', '写状态指示器初始为 data-state="saved"', async () => {
    const state = await page.evaluate(`(document.getElementById('save-state') || { getAttribute: function () { return null; } }).getAttribute('data-state')`);
    assertEq(state, 'saved', '#save-state data-state');
    const text = await page.text('#save-state');
    assert(String(text).length > 0, '#save-state 文本为空');
  });

  await test('A7', '公共测试接口自检（CONTRACT v1.1 §8：app / store / schema / parser / xlsx / 产品库）', async () => {
    const api = await page.evaluate(`(() => {
      const crm = window.CustomerCRM;
      if (!crm) return { missing: 'window.CustomerCRM' };
      const methods = ['navigate', 'getRoute', 'refresh', 'newCustomer', 'editCustomer', 'saveCustomerForm', 'deleteCustomer', 'addFollowup', 'recognize', 'applyRecognition', 'exportExcel', 'stats', 'filtered', 'selectProduct', 'addProduct', 'products', 'toast', 'openFollowupModal'];
      const missingMethods = methods.filter(function (name) { return typeof crm.app[name] !== 'function'; });
      const storeMethods = ['list', 'get', 'add', 'update', 'remove', 'replaceAll', 'clear', 'seedDemo', 'settings', 'updateSettings', 'stats', 'exportJSON', 'importJSON', 'savedAt', 'listProducts', 'productNames', 'productModels', 'addProduct', 'removeProduct', 'seedProducts', 'productUsage'];
      const missingStore = storeMethods.filter(function (name) { return typeof crm.store[name] !== 'function'; });
      return {
        version: crm.version,
        schemaVersion: crm.schema.VERSION,
        parserVersion: crm.parser ? crm.parser.VERSION : null,
        key: crm.store.KEY,
        missingMethods: missingMethods,
        missingStore: missingStore,
        parser: !!(crm.parser && typeof crm.parser.recognize === 'function'),
        sampleCount: crm.parser && Array.isArray(crm.parser.SAMPLE_TEXTS) ? crm.parser.SAMPLE_TEXTS.length : 0,
        xlsx: !!(crm.xlsx && typeof crm.xlsx.exportCustomers === 'function' && typeof crm.xlsx.download === 'function'),
        hasState: !!crm.app.state
      };
    })()`);
    assertEq(api.version, '1.1.0', 'CustomerCRM.version');
    assertEq(api.schemaVersion, '1.1.0', 'CRM_SCHEMA.VERSION');
    assertEq(api.parserVersion, '1.1.0', 'CrmParser.VERSION');
    assertEq(api.key, 'customer-crm.v1', 'store.KEY');
    assertEq(api.missingMethods.join(','), '', 'app 缺少契约方法');
    assertEq(api.missingStore.join(','), '', 'store 缺少契约方法');
    assert(api.parser, 'CrmParser.recognize 未加载');
    assert(api.sampleCount >= 3, 'CrmParser.SAMPLE_TEXTS 应至少 3 段', api.sampleCount);
    assert(api.xlsx, 'CrmXlsx.exportCustomers / download 未加载');
    assert(api.hasState, 'app.state 不存在');
  });

  await test('A8', '1440×900 首屏无整页横向溢出', async () => {
    await assertNoOverflow(page, '1440 仪表盘');
  });

  await test('A9', '截图 screenshots/dashboard.png（1440×900）', async () => {
    await takeScreenshot(page, 'dashboard.png');
    assetScreenshot('dashboard.png', 1440, 900);
  });

  await test('A10', '仪表盘国家分布 SVG 扇形图（7 切片 = 7 图例、百分比合计 100、中心数=12）', async () => {
    await waitSelector(page, '[data-testid="country-pie"]', '国家分布扇形图容器');
    const pie = await readPie(page);
    assert(pie, '未渲染 [data-testid="country-pie"]');
    assert(pie.svg, '扇形图缺少 .pie-svg');
    assertEq(pie.slices, 7, '切片数量（前 6 国家 + 其他）');
    assertEq(pie.legend, pie.slices, '图例数量应等于切片数量');
    assertEq(pie.pctSum, 100, '图例百分比合计');
    assertEq(String(pie.total).trim(), '12', '扇形图中心客户总数');
    assertEq(pie.legendCountries.join('|'), pie.sliceCountries.join('|'), '图例国家顺序 = 切片 data-country');
    for (let i = 0; i < pie.slices; i += 1) {
      assert(pie.slicePaths[i] > 10, '切片缺少有效 path', { index: i, pathLength: pie.slicePaths[i] });
      assert(String(pie.sliceTitles[i]).length > 0, '切片缺少 hover <title>', { index: i });
    }
    assert(pie.sliceCountries.indexOf('其他 6 个国家') !== -1, '12 个国家应聚合出「其他 6 个国家」切片', pie.sliceCountries);
  });
}

/* ================================================================== */
/* B. 客户台账                                                          */
/* ================================================================== */

async function groupB(page) {
  setGroup('B', '客户台账：表格 / 搜索 / 筛选 / 排序 / 展开');

  await test('B1', '客户台账 12 行 × 15 列，产品型号小字、工具栏与 data-testid 齐全', async () => {
    await resetDemo(page);
    await goto(page, '#/customers');
    await waitRowCount(page, 12);
    const layout = await page.evaluate(`(() => {
      const row = document.querySelector('tr[data-testid="customer-row"]');
      const headers = Array.from(document.querySelectorAll('.data-table--customers thead th')).map(function (th) { return th.textContent.trim(); });
      const productCell = document.querySelector('tr[data-testid="customer-row"] td[data-col="product"]');
      return {
        cells: row ? row.querySelectorAll('td').length : 0,
        headers: headers,
        productHtml: productCell ? productCell.innerHTML : '',
        bodyClass: document.body.className
      };
    })()`);
    assertEq(layout.cells, 15, '数据行 td 数量（展开列 + 13 字段合并为 12 列 + 距上次跟进 + 操作列）');
    assertEq(layout.headers.length, 15, '表头列数');
    const headerText = layout.headers.join('|');
    for (const label of ['首次询盘时间', '距上次跟进', '客户名称', '咨询产品 / 型号', '国家', '公司名', '主营业务', '成立时间', '团队规模', '年营收', '跟进状态', '是否成交', '核心原因']) {
      assert(headerText.indexOf(label) !== -1, '表头缺少列：' + label, layout.headers);
    }
    assert(/route-customers/.test(layout.bodyClass), '客户台账页应带 body.route-customers（宽版布局）', layout.bodyClass);
    assert(/cell-sub/.test(layout.productHtml), '产品单元格应渲染型号小字 .cell-sub', layout.productHtml.slice(0, 160));
    for (const testid of ['btn-export-excel', 'customer-search', 'filter-country', 'filter-status', 'filter-deal', 'filter-reset', 'btn-edit-customer', 'btn-delete-customer', 'btn-add-followup', 'row-toggle']) {
      assert(await page.exists(`[data-testid="${testid}"]`), '缺少 data-testid：' + testid);
    }
    assert(await page.exists('[data-testid="new-customer"]'), '缺少 data-testid：new-customer');
  });

  await test('B2', '关键词搜索「德国」→ 命中 1 行，DOM 顺序 = app.filtered()', async () => {
    await page.fill('[data-testid="customer-search"]', '德国');
    await waitRowCount(page, 1);
    const state = await page.evaluate(`(() => ({
      domIds: Array.from(document.querySelectorAll('tr[data-testid="customer-row"]')).map(function (tr) { return tr.getAttribute('data-id'); }),
      filteredIds: window.CustomerCRM.app.filtered().map(function (record) { return record.id; }),
      filteredOn: !!document.querySelector('.filter-on')
    }))()`);
    assertEq(state.domIds.join(','), state.filteredIds.join(','), '搜索结果 DOM 顺序 = app.filtered()');
    assert(state.filteredOn, '筛选生效时应显示「已筛选」标记');
    const country = await page.text('tr[data-testid="customer-row"] td[data-col="country"]');
    assert(country.indexOf('德国') !== -1, '命中行国家列应含「德国」', country);
  });

  await test('B3', '搜索无结果 → 空状态引导，重置后恢复 12 行', async () => {
    await page.fill('[data-testid="customer-search"]', 'zzz-不存在的客户-zzz');
    await waitRowCount(page, 0);
    assert(await page.exists('.empty'), '无结果时应显示空状态 .empty');
    const emptyTitle = String(await page.text('.empty-title')).trim();
    assertEq(emptyTitle, '没有匹配的客户', '无结果空状态标题');
    await click(page, '[data-testid="filter-reset"]');
    await waitRowCount(page, 12);
    assertEq(await page.evaluate('window.CustomerCRM.app.state.filters.q'), '', '重置后关键词应清空');
  });

  await test('B4', '国家筛选「德国」→ 1 行（与 store 计算一致）', async () => {
    const expected = await page.evaluate(`window.CustomerCRM.store.list().filter(function (record) { return record.country === '德国'; }).length`);
    await page.fill('[data-testid="filter-country"]', '德国');
    await waitRowCount(page, expected);
    assertEq(expected, 1, 'store 中德国客户数');
    assertEq(await page.evaluate(`(document.querySelector('[data-testid="filter-country"]') || {}).value`), '德国', '国家下拉选中值');
    await click(page, '[data-testid="filter-reset"]');
    await waitRowCount(page, 12);
  });

  await test('B5', '跟进状态筛选「已报价」→ 3 行', async () => {
    const expected = await page.evaluate(`window.CustomerCRM.store.list().filter(function (record) { return record.followupStatus === '已报价'; }).length`);
    await page.fill('[data-testid="filter-status"]', '已报价');
    await waitRowCount(page, expected);
    assertEq(expected, 3, 'store 中已报价客户数');
    const statuses = await readCol(page, 'followupStatus');
    assertEq(statuses.join(','), new Array(3).fill('已报价').join(','), '筛选结果跟进状态均为已报价');
    await click(page, '[data-testid="filter-reset"]');
    await waitRowCount(page, 12);
  });

  await test('B6', '是否成交筛选「已成交」→ 2 行', async () => {
    const expected = await page.evaluate(`window.CustomerCRM.store.list().filter(function (record) { return record.dealStatus === '已成交'; }).length`);
    await page.fill('[data-testid="filter-deal"]', '已成交');
    await waitRowCount(page, expected);
    assertEq(expected, 2, 'store 中已成交客户数');
    const deals = await readCol(page, 'dealStatus');
    assertEq(deals.join(','), new Array(2).fill('已成交').join(','), '筛选结果成交状态均为已成交');
    await click(page, '[data-testid="filter-reset"]');
    await waitRowCount(page, 12);
  });

  await test('B7', '组合筛选：跟进状态=已报价 + 国家=德国 → 1 行', async () => {
    await page.fill('[data-testid="filter-status"]', '已报价');
    await waitRowCount(page, 3);
    await page.fill('[data-testid="filter-country"]', '德国');
    await waitRowCount(page, 1);
    const names = await readCol(page, 'customerName');
    assertEq(names.join(','), 'Michael Braun', '组合筛选命中客户');
    assertEq(await page.evaluate('window.CustomerCRM.app.filtered().length'), 1, 'app.filtered() 条数');
    await click(page, '[data-testid="filter-reset"]');
    await waitRowCount(page, 12);
  });

  await test('B8', '重置筛选恢复 12 行且筛选控件归位', async () => {
    assertEq(await page.evaluate('window.CustomerCRM.app.state.filters.q'), '', '关键词');
    assertEq(await page.evaluate('window.CustomerCRM.app.state.filters.country'), '', '国家筛选');
    assertEq(await page.evaluate('window.CustomerCRM.app.state.filters.status'), '', '状态筛选');
    assertEq(await page.evaluate('window.CustomerCRM.app.state.filters.deal'), '', '成交筛选');
    assertEq(await page.evaluate(`(document.querySelector('[data-testid="filter-country"]') || {}).value`), '', '国家下拉值');
    assertEq(await rowCount(page), 12, '重置后行数');
  });

  await test('B9', '排序：首次询盘时间 desc → asc → 恢复默认（DOM 顺序 = app.filtered()）', async () => {
    await click(page, '[data-sort="firstInquiryAt"]');
    await sleep(60);
    let ids = await readRowIds(page);
    let expected = await filteredIds(page);
    assertEq(ids.join(','), expected.join(','), '首次点击后 DOM 顺序 = app.filtered()');
    let dates = await readCol(page, 'firstInquiryAt');
    for (let i = 1; i < dates.length; i += 1) assert(dates[i - 1] >= dates[i], '首次点击应按首次询盘时间降序', dates);
    await click(page, '[data-sort="firstInquiryAt"]');
    await sleep(60);
    ids = await readRowIds(page);
    expected = await filteredIds(page);
    assertEq(ids.join(','), expected.join(','), '第二次点击后 DOM 顺序 = app.filtered()');
    dates = await readCol(page, 'firstInquiryAt');
    for (let i = 1; i < dates.length; i += 1) assert(dates[i - 1] <= dates[i], '第二次点击应按首次询盘时间升序', dates);
    await click(page, '[data-sort="firstInquiryAt"]');
    await sleep(60);
    assertEq(await page.evaluate('window.CustomerCRM.app.state.filters.sort'), 'updated-desc', '第三次点击应恢复默认排序 updated-desc');
    await click(page, '[data-testid="filter-reset"]');
    await waitRowCount(page, 12);
  });

  await test('B10', '排序：客户名称 desc → asc（DOM 顺序 = app.filtered()）', async () => {
    await click(page, '[data-sort="customerName"]');
    await sleep(60);
    let expected = await page.evaluate('window.CustomerCRM.app.filtered().map(function (record) { return record.customerName; })');
    let names = await readCol(page, 'customerName');
    assertEq(names.join('|'), expected.join('|'), '客户名称降序排序结果');
    await click(page, '[data-sort="customerName"]');
    await sleep(60);
    expected = await page.evaluate('window.CustomerCRM.app.filtered().map(function (record) { return record.customerName; })');
    names = await readCol(page, 'customerName');
    assertEq(names.join('|'), expected.join('|'), '客户名称升序排序结果');
    assertEq(await page.evaluate('window.CustomerCRM.app.state.filters.sort'), 'customerName-asc', '排序状态 customerName-asc');
    await click(page, '[data-testid="filter-reset"]');
    await waitRowCount(page, 12);
  });

  await test('B11', '排序：距上次跟进（daysSince desc）', async () => {
    await click(page, '[data-sort="daysSince"]');
    await sleep(60);
    assertEq(await page.evaluate('window.CustomerCRM.app.state.filters.sort'), 'daysSince-desc', '排序状态 daysSince-desc');
    const ids = await readRowIds(page);
    const expected = await filteredIds(page);
    assertEq(ids.join(','), expected.join(','), '距上次跟进排序 DOM 顺序 = app.filtered()');
    const derived = await page.evaluate('window.CustomerCRM.store.deriveAll()');
    const sorted = derived.slice().sort((a, b) => {
      if (a.daysSince === null && b.daysSince === null) return 0;
      if (a.daysSince === null) return 1;
      if (b.daysSince === null) return -1;
      return b.daysSince - a.daysSince;
    });
    assertEq(ids.join(','), sorted.map((record) => record.id).join(','), 'DOM 顺序 = daysSince 降序');
    await click(page, '[data-testid="filter-reset"]');
    await waitRowCount(page, 12);
  });

  await test('B12', '展开行：显示最近三次跟进（最新在前，与 store 一致）', async () => {
    const firstId = await page.evaluate(`document.querySelector('tr[data-testid="customer-row"]').getAttribute('data-id')`);
    assert(firstId, '找不到首行 data-id');
    const record = await page.evaluate(`window.CustomerCRM.store.get(${J(firstId)})`);
    assert(record, 'store 中找不到首行客户', firstId);
    await click(page, `tr[data-id="${firstId}"] [data-testid="row-toggle"]`);
    const detail = await page.evaluate(`(() => {
      const row = document.querySelector('tr[data-id="${firstId}"]');
      const detail = document.querySelector('tr.row-detail[data-detail="${firstId}"]');
      if (!row || !detail) return null;
      return {
        hidden: detail.hidden,
        expanded: row.classList.contains('row--expanded'),
        aria: row.querySelector('[data-testid="row-toggle"]').getAttribute('aria-expanded'),
        chips: Array.from(detail.querySelectorAll('.followup-chip')).map(function (chip) {
          return {
            date: (chip.querySelector('.followup-date') || {}).textContent || '',
            note: (chip.querySelector('.followup-note') || {}).textContent || ''
          };
        })
      };
    })()`);
    assert(detail, '展开后找不到行详情 tr.row-detail');
    assertEq(detail.hidden, false, '行详情 hidden=false');
    assertEq(detail.expanded, true, '数据行应有 row--expanded');
    assertEq(detail.aria, 'true', '展开按钮 aria-expanded');
    assertEq(detail.chips.length, Math.min(3, record.followups.length), '展开显示跟进条数（≤3）');
    const expectedNotes = record.followups.slice(0, detail.chips.length).map((item) => item.note);
    assertEq(detail.chips.map((chip) => chip.note).join('|'), expectedNotes.join('|'), '跟进内容与 store 一致（最新在前）');
    assert(detail.chips[0].date.indexOf('最近') !== -1, '第一条跟进应标注「最近」', detail.chips[0].date);
    const meta = await page.evaluate(`(() => {
      const el = document.querySelector('tr.row-detail[data-detail="${firstId}"] .row-detail-meta');
      return el ? el.textContent : '';
    })()`);
    for (const label of ['产品型号', '主营业务', '成立时间', '团队规模', '年营收', '首次询盘', '最近更新']) {
      assert(meta.indexOf(label) !== -1, '行展开 meta 应含「' + label + '」', meta);
    }
    assert(meta.indexOf(record.productModel) !== -1, '行展开 meta 应含产品型号值', { meta, productModel: record.productModel });
    assert(meta.indexOf(record.mainBusiness) !== -1, '行展开 meta 应含主营业务值', { meta, mainBusiness: record.mainBusiness });
    scratch.expandedId = firstId;
  });

  await test('B13', '再次点击收起行详情', async () => {
    const id = scratch.expandedId;
    assert(id, '前置展开行用例未成功');
    await click(page, `tr[data-id="${id}"] [data-testid="row-toggle"]`);
    const state = await page.evaluate(`(() => {
      const row = document.querySelector('tr[data-id="${id}"]');
      const detail = document.querySelector('tr.row-detail[data-detail="${id}"]');
      return { hidden: detail ? detail.hidden : null, expanded: row ? row.classList.contains('row--expanded') : null };
    })()`);
    assertEq(state.hidden, true, '收起后行详情 hidden=true');
    assertEq(state.expanded, false, '收起后无 row--expanded');
  });

  await test('B14', '截图 screenshots/customers.png 与 customers-expanded.png（1440×900）', async () => {
    await waitRowCount(page, 12);
    await takeScreenshot(page, 'customers.png');
    assetScreenshot('customers.png', 1440, 900);
    const firstId = await page.evaluate(`document.querySelector('tr[data-testid="customer-row"]').getAttribute('data-id')`);
    await click(page, `tr[data-id="${firstId}"] [data-testid="row-toggle"]`);
    await waitExpr(page, `!document.querySelector('tr.row-detail[data-detail="${firstId}"]').hidden`, '行详情展开');
    await takeScreenshot(page, 'customers-expanded.png');
    assetScreenshot('customers-expanded.png', 1440, 900);
    await click(page, `tr[data-id="${firstId}"] [data-testid="row-toggle"]`);
  });

  await test('B16', '搜索主营业务「户外用品」→ 命中 2 行且主营列有值', async () => {
    const expected = await page.evaluate(`(() => {
      const keyword = '户外用品';
      return window.CustomerCRM.store.list().filter(function (record) {
        const haystack = [record.customerName, record.companyName, record.product, record.productModel, record.mainBusiness, record.country, record.coreReason].join(' ');
        return haystack.indexOf(keyword) !== -1;
      }).length;
    })()`);
    await page.fill('[data-testid="customer-search"]', '户外用品');
    await waitRowCount(page, expected);
    assertEq(expected, 2, '示例数据中主营业务含「户外用品」的客户数');
    const ids = await readRowIds(page);
    const filtered = await filteredIds(page);
    assertEq(ids.join(','), filtered.join(','), 'DOM 顺序 = app.filtered()');
    const businesses = await readCol(page, 'mainBusiness');
    assert(businesses.length === 2 && businesses.every((text) => text.indexOf('户外用品') !== -1), '命中行主营业务列应含关键词', businesses);
    await click(page, '[data-testid="filter-reset"]');
    await waitRowCount(page, 12);
  });

  await test('B17', '搜索产品型号「500ml」→ 命中 2 行且产品单元格含型号小字', async () => {
    const expected = await page.evaluate(`(() => {
      const keyword = '500ml';
      return window.CustomerCRM.store.list().filter(function (record) {
        return (record.product + ' ' + record.productModel + ' ' + record.mainBusiness).toLowerCase().indexOf(keyword) !== -1;
      }).length;
    })()`);
    await page.fill('[data-testid="customer-search"]', '500ml');
    await waitRowCount(page, expected);
    assertEq(expected, 2, '示例数据中型号为 500ml 的客户数');
    const cells = await page.evaluate(`Array.from(document.querySelectorAll('tr[data-testid="customer-row"] td[data-col="product"]')).map(function (td) { return td.textContent; })`);
    assert(cells.length === 2 && cells.every((text) => text.indexOf('500ml') !== -1), '命中行产品单元格应含 500ml', cells);
    const subs = await page.evaluate(`Array.from(document.querySelectorAll('tr[data-testid="customer-row"] .cell-sub')).map(function (el) { return el.textContent; })`);
    assert(subs.some((text) => text.indexOf('500ml') !== -1), '产品型号应渲染在 .cell-sub 小字中', subs);
    await click(page, '[data-testid="filter-reset"]');
    await waitRowCount(page, 12);
  });

  await assertNoErrors(page, 'B18', '客户台账全流程无 console / page 错误');
}

/* ================================================================== */
/* C. 新建 / 编辑 / 跟进 / 删除                                         */
/* ================================================================== */

async function groupC(page) {
  setGroup('C', '客户增删改与跟进记录');

  const newValues = {
    firstInquiryAt: '',
    customerName: 'E2E 测试客户',
    mainBusiness: '测试主营业务：户外用品进口',
    country: '德国',
    companyName: 'E2E Test GmbH',
    foundedAt: '2015',
    teamSize: '20-50 人',
    annualRevenue: '500 万欧元',
    followupStatus: '谈判中',
    dealStatus: '未成交',
    coreReason: 'E2E 测试：价格与交期符合，等待客户确认订单'
  };

  await test('C1', '打开新建客户弹窗：13 字段映射 + 产品选择器 testid + role=dialog', async () => {
    await resetDemo(page);
    await goto(page, '#/customers');
    await waitRowCount(page, 12);
    await click(page, '[data-testid="new-customer"]');
    await waitSelector(page, '[data-testid="form-customer"]');
    const form = await page.evaluate(`(() => {
      const form = document.getElementById('customer-form');
      const names = Array.from(form.querySelectorAll('[name]')).map(function (el) { return el.getAttribute('name'); });
      const keys = window.CRM_SCHEMA.FIELD_KEYS.slice();
      const missingFields = keys.filter(function (key) {
        if (key === 'product' || key === 'productModel') return false; /* v1.1 由产品选择器承载 */
        return !form.elements[key];
      });
      const nameSelect = form.querySelector('[name="productName"]');
      const modelSelect = form.querySelector('[name="productModel"]');
      return {
        testid: form.getAttribute('data-testid'),
        mode: form.getAttribute('data-mode'),
        names: names,
        missingFields: missingFields,
        productName: !!document.querySelector('[data-testid="product-name-select"]'),
        productModel: !!document.querySelector('[data-testid="product-model-select"]'),
        productCustom: !!document.querySelector('[data-testid="product-custom-input"]'),
        productOptions: nameSelect ? Array.from(nameSelect.options).map(function (option) { return option.value; }) : [],
        modelOptions: modelSelect ? modelSelect.options.length : 0,
        save: !!document.querySelector('[data-testid="btn-save-customer"]'),
        cancel: !!document.querySelector('[data-testid="btn-cancel-customer"]'),
        error: !!document.querySelector('[data-testid="form-error"]'),
        dialog: !!document.querySelector('#modal-root [role="dialog"][aria-modal="true"]'),
        required: Array.from(form.querySelectorAll('.field-label .req')).length
      };
    })()`);
    assertEq(form.testid, 'form-customer', '表单 data-testid');
    assertEq(form.mode, 'create', '新建表单 data-mode');
    assertEq(form.missingFields.join(','), '', '13 个字段中缺少表单控件');
    assert(form.productName && form.productModel && form.productCustom, 'v1.1 产品选择器 data-testid 缺失', form);
    assert(form.productOptions.indexOf('不锈钢保温杯') !== -1, '产品名称下拉应含产品库产品', form.productOptions.slice(0, 16));
    assertEq(form.modelOptions, 1, '未选产品名称前型号下拉只有占位项');
    assert(form.save && form.cancel && form.error, '表单按钮/错误提示 data-testid 缺失', form);
    assert(form.dialog, '弹窗缺少 role="dialog" aria-modal="true"');
    assertEq(form.required, 2, '必填标记数量（首次询盘时间 / 客户名称）');
    const today = await page.evaluate('window.CRM_SCHEMA.todayISO()');
    assertEq(await page.evaluate(`document.querySelector('#customer-form [name="firstInquiryAt"]').value`), today, '新建表单首次询盘时间默认今天');
  });

  await test('C2', '截图 screenshots/form.png（1440×900，产品名称+型号联动已选）', async () => {
    await page.fill('customerName', 'E2E 验收客户');
    await page.fill('mainBusiness', '户外用品进口分销');
    await page.fill('country', '德国');
    await page.fill('companyName', 'QA Acceptance GmbH');
    await page.fill('coreReason', '价格与交期均符合，等待客户确认订单');
    await selectOption(page, '[data-testid="product-name-select"]', '不锈钢保温杯');
    await waitExpr(page, `Array.from(document.querySelectorAll('#customer-form [name="productModel"] option')).length >= 4`, '型号下拉联动出现 ≥3 个型号');
    await selectOption(page, '[data-testid="product-model-select"]', '500ml');
    const filled = await page.evaluate(`(() => {
      const form = document.getElementById('customer-form');
      return {
        customerName: form.querySelector('[name="customerName"]').value,
        mainBusiness: form.querySelector('[name="mainBusiness"]').value,
        productName: form.querySelector('[name="productName"]').value,
        productModel: form.querySelector('[name="productModel"]').value
      };
    })()`);
    assertEq(filled.customerName, 'E2E 验收客户', '截图前客户名称已填入');
    assertEq(filled.mainBusiness, '户外用品进口分销', '截图前主营业务已填入');
    assertEq(filled.productName, '不锈钢保温杯', '截图前产品名称已选择');
    assertEq(filled.productModel, '500ml', '截图前产品型号已选择');
    await page.evaluate(`(() => { if (document.activeElement && document.activeElement.blur) document.activeElement.blur(); })()`);
    await sleep(200);
    await takeScreenshot(page, 'form.png');
    assetScreenshot('form.png', 1440, 900);
  });

  await test('C3', '必填校验：清空客户名称与首次询盘时间 → form-error 出现且弹窗不关闭', async () => {
    await page.fill('customerName', '');
    await page.fill('firstInquiryAt', '');
    await click(page, '[data-testid="btn-save-customer"]');
    await waitExpr(
      page,
      `(() => { const el = document.querySelector('[data-testid="form-error"]'); return el && !el.hidden && el.textContent.trim().length > 0; })()`,
      '必填校验错误提示显示'
    );
    const errorText = String(await page.text('[data-testid="form-error"]'));
    assert(errorText.indexOf('客户名称') !== -1, '错误汇总应包含「客户名称」', errorText);
    assert(errorText.indexOf('首次询盘时间') !== -1, '错误汇总应包含「首次询盘时间」', errorText);
    assert(await page.exists('[data-testid="form-customer"]'), '校验失败时弹窗不应关闭');
    const fieldError = await page.evaluate(`(() => { const el = document.querySelector('[data-error-for="customerName"]'); return el ? (!el.hidden ? el.textContent : '') : ''; })()`);
    assert(String(fieldError).length > 0, '客户名称字段级错误未显示');
  });

  await test('C4', '填表保存（产品名称+型号下拉）→ store 13 条、13 字段落库、表格出现产品小字', async () => {
    const today = await page.evaluate('window.CRM_SCHEMA.todayISO()');
    const values = Object.assign({}, newValues, { firstInquiryAt: today });
    for (const [key, value] of Object.entries(values)) {
      const ok = await page.fill(key, value);
      assert(ok, '表单字段填充失败：' + key);
    }
    await selectOption(page, '[data-testid="product-name-select"]', '不锈钢保温杯');
    await waitExpr(page, `Array.from(document.querySelectorAll('#customer-form [name="productModel"] option')).length >= 4`, '型号下拉联动出现 ≥3 个型号');
    await selectOption(page, '[data-testid="product-model-select"]', '500ml');
    await clearToasts(page);
    await click(page, '[data-testid="btn-save-customer"]');
    await waitGone(page, '[data-testid="form-customer"]', '保存后弹窗关闭');
    await waitStoreCount(page, 13);
    await waitRowCount(page, 13);
    await waitToast(page, 'ok', '已新增客户');
    const record = await page.evaluate(`window.CustomerCRM.store.list().filter(function (item) { return item.customerName === 'E2E 测试客户'; })[0] || null`);
    assert(record, 'store 中找不到新建的 E2E 测试客户');
    scratch.customerId = record.id;
    assertEq(record.customerName, 'E2E 测试客户', 'store 客户名称');
    assertEq(record.companyName, 'E2E Test GmbH', 'store 公司名');
    assertEq(record.country, '德国', 'store 国家');
    assertEq(record.product, '不锈钢保温杯', 'store 咨询产品（产品库选择）');
    assertEq(record.productModel, '500ml', 'store 产品型号（型号下拉选择）');
    assertEq(record.mainBusiness, '测试主营业务：户外用品进口', 'store 主营业务');
    assertEq(record.followupStatus, '谈判中', 'store 跟进状态');
    assertEq(record.coreReason, 'E2E 测试：价格与交期符合，等待客户确认订单', 'store 核心原因');
    const inTable = await page.evaluate(`Array.from(document.querySelectorAll('tr[data-testid="customer-row"] td[data-col="customerName"]')).some(function (td) { return td.textContent.indexOf('E2E 测试客户') !== -1; })`);
    assert(inTable, '表格未出现新建客户');
    const productCell = await page.evaluate(`(() => {
      const row = Array.from(document.querySelectorAll('tr[data-testid="customer-row"]')).filter(function (tr) {
        const td = tr.querySelector('td[data-col="customerName"]');
        return td && td.textContent.indexOf('E2E 测试客户') !== -1;
      })[0];
      const cell = row ? row.querySelector('td[data-col="product"]') : null;
      return cell ? { text: cell.textContent, sub: cell.querySelector('.cell-sub') ? cell.querySelector('.cell-sub').textContent : '' } : null;
    })()`);
    assert(productCell, '表格中找不到新客户的产品单元格');
    assert(productCell.text.indexOf('不锈钢保温杯') !== -1, '产品单元格应显示产品名称', productCell);
    assertEq(productCell.sub, '500ml', '产品单元格 .cell-sub 应显示型号');
    const saveState = await page.evaluate(`document.querySelector('#save-state').getAttribute('data-state')`);
    assertEq(saveState, 'saved', '写操作后 #save-state data-state');
  });

  await test('C5', '刷新页面后新客户仍在（localStorage 持久化 13 字段）', async () => {
    const id = scratch.customerId;
    assert(id, '前置保存用例未成功');
    await page.reload();
    await page.waitForApp({ timeout: 20000 });
    await waitExpr(page, `!!document.querySelector('#view-inner .page')`, '刷新后页面渲染', 15000);
    await waitRowCount(page, 13);
    const persisted = await page.evaluate(`(() => {
      try {
        const payload = JSON.parse(localStorage.getItem('customer-crm.v1'));
        const record = payload.customers.filter(function (item) { return item.id === ${J(id)}; })[0];
        return {
          count: payload.customers.length,
          name: record ? record.customerName : null,
          product: record ? record.product : null,
          productModel: record ? record.productModel : null,
          mainBusiness: record ? record.mainBusiness : null,
          products: Array.isArray(payload.products) ? payload.products.length : null,
          error: ''
        };
      } catch (err) {
        return { count: null, name: null, error: String(err && err.message) };
      }
    })()`);
    assertEq(persisted.error, '', 'localStorage 数据应可解析');
    assertEq(persisted.count, 13, 'localStorage 客户记录数');
    assertEq(persisted.name, 'E2E 测试客户', 'localStorage 中客户名称');
    assertEq(persisted.product, '不锈钢保温杯', 'localStorage 中咨询产品');
    assertEq(persisted.productModel, '500ml', 'localStorage 中产品型号');
    assertEq(persisted.mainBusiness, '测试主营业务：户外用品进口', 'localStorage 中主营业务');
    assertEq(persisted.products, 12, 'localStorage 中产品库条数');
    const visible = await page.evaluate(`Array.from(document.querySelectorAll('tr[data-testid="customer-row"] td[data-col="customerName"]')).some(function (td) { return td.textContent.indexOf('E2E 测试客户') !== -1; })`);
    assert(visible, '刷新后表格未显示新客户');
  });

  await test('C6', '编辑客户：产品选择器预填正确 → 保存后 store / 表格同步更新', async () => {
    const id = scratch.customerId;
    assert(id, '前置保存用例未成功');
    await goto(page, '#/customers');
    await waitRowCount(page, 13);
    await click(page, `tr[data-id="${id}"] [data-testid="btn-edit-customer"]`);
    await waitSelector(page, '[data-testid="form-customer"]');
    assertEq(await page.evaluate(`document.getElementById('customer-form').getAttribute('data-mode')`), 'edit', '编辑表单 data-mode');
    const before = await formValues(page);
    assertEq(before.customerName, 'E2E 测试客户', '编辑弹窗预填客户名称');
    assertEq(before.companyName, 'E2E Test GmbH', '编辑弹窗预填公司名');
    assertEq(before.productName, '不锈钢保温杯', '编辑弹窗产品名称下拉预填');
    assertEq(before.productModel, '500ml', '编辑弹窗产品型号下拉预填');
    assertEq(before.mainBusiness, '测试主营业务：户外用品进口', '编辑弹窗预填主营业务');
    await page.fill('customerName', 'E2E 测试客户（已改名）');
    await page.fill('mainBusiness', '测试主营业务：家居用品分销');
    await clearToasts(page);
    await click(page, '[data-testid="btn-save-customer"]');
    await waitGone(page, '[data-testid="form-customer"]', '编辑保存后弹窗关闭');
    await waitExpr(page, `window.CustomerCRM.store.get(${J(id)}).customerName === 'E2E 测试客户（已改名）'`, 'store 客户名已更新');
    const updated = await page.evaluate(`window.CustomerCRM.store.get(${J(id)})`);
    assertEq(updated.customerName, 'E2E 测试客户（已改名）', 'store 客户名称');
    assertEq(updated.mainBusiness, '测试主营业务：家居用品分销', 'store 主营业务已更新');
    assertEq(updated.product, '不锈钢保温杯', '编辑后咨询产品保持');
    assertEq(updated.productModel, '500ml', '编辑后产品型号保持');
    const visible = await page.evaluate(`Array.from(document.querySelectorAll('tr[data-testid="customer-row"] td[data-col="customerName"]')).some(function (td) { return td.textContent.indexOf('E2E 测试客户（已改名）') !== -1; })`);
    assert(visible, '表格未显示编辑后的客户名称');
    await waitToast(page, 'ok', '客户资料已更新');
  });

  await test('C7', '添加跟进第 1~3 条：条数递增、最新在前', async () => {
    const id = scratch.customerId;
    assert(id, '前置保存用例未成功');
    const dates = await page.evaluate(`(() => { const schema = window.CRM_SCHEMA; const today = schema.todayISO(); return { today: today, m1: schema.addDays(today, -1), m2: schema.addDays(today, -2), m3: schema.addDays(today, -3) }; })()`);
    scratch.dates = dates;
    const plan = [
      { date: dates.m3, note: '第 1 次跟进：发送产品目录' },
      { date: dates.m2, note: '第 2 次跟进：确认试单数量' },
      { date: dates.m1, note: '第 3 次跟进：寄出样品' }
    ];
    for (let i = 0; i < plan.length; i += 1) {
      await click(page, `tr[data-id="${id}"] [data-testid="btn-add-followup"]`);
      await waitSelector(page, '#followup-note', '跟进弹窗出现');
      await page.fill('#followup-date', plan[i].date);
      await page.fill('#followup-note', plan[i].note);
      await click(page, '[data-action="save-followup"]');
      await waitGone(page, '#followup-note', '跟进保存后弹窗关闭');
      await waitExpr(page, `window.CustomerCRM.store.get(${J(id)}).followups.length === ${i + 1}`, '跟进条数 = ' + (i + 1));
      const followups = await page.evaluate(`window.CustomerCRM.store.get(${J(id)}).followups`);
      assertEq(followups[0].date, plan[i].date, '最新一条跟进日期应在最前');
      assertEq(followups[0].note, plan[i].note, '最新一条跟进内容');
    }
    const followups = await page.evaluate(`window.CustomerCRM.store.get(${J(id)}).followups`);
    assertEq(followups.length, 3, '三次跟进记录数');
    assertEq(followups.map((item) => item.date).join(','), [dates.m1, dates.m2, dates.m3].join(','), '跟进记录应按日期倒序（最新在前）');
  });

  await test('C8', '添加第 4 条跟进：只保留 3 条、最新在前、「距上次跟进」= 今天', async () => {
    const id = scratch.customerId;
    const dates = scratch.dates;
    assert(id && dates, '前置跟进用例未成功');
    await click(page, `tr[data-id="${id}"] [data-testid="btn-add-followup"]`);
    await waitSelector(page, '#followup-note', '跟进弹窗出现');
    await page.fill('#followup-date', dates.today);
    await page.fill('#followup-note', '第 4 次跟进（最新）');
    await click(page, '[data-action="save-followup"]');
    await waitGone(page, '#followup-note', '跟进保存后弹窗关闭');
    await waitExpr(
      page,
      `(() => { const record = window.CustomerCRM.store.get(${J(id)}); return record && record.followups.length === 3 && record.followups[0].date === ${J(dates.today)}; })()`,
      '第 4 条跟进写入且截断到 3 条'
    );
    const followups = await page.evaluate(`window.CustomerCRM.store.get(${J(id)}).followups`);
    assertEq(followups.length, 3, '只保留最近 3 条跟进');
    assertEq(followups[0].note, '第 4 次跟进（最新）', '最新一条跟进内容');
    assertEq(followups[0].date, dates.today, '最新一条跟进日期');
    assertEq(followups[1].date, dates.m1, '第 2 条应为原最新');
    assertEq(followups[2].date, dates.m2, '第 3 条应为原第 2 条');
    assert(!followups.some((item) => item.note === '第 1 次跟进：发送产品目录'), '最旧一条跟进应被截断');
    const age = await page.evaluate(`(() => {
      const cell = document.querySelector('tr[data-id="${id}"] td[data-col="age"]');
      if (!cell) return null;
      const el = cell.querySelector('.age');
      return { text: el ? el.textContent.trim() : '', cls: el ? el.className : '' };
    })()`);
    assert(age, '找不到该客户的距上次跟进单元格');
    assertEq(age.text, '今天', '距上次跟进应显示「今天」');
    assert(age.cls.indexOf('age--ok') !== -1, '距上次跟进色调应为 ok（绿）', age.cls);
  });

  await test('C9', '删除客户：取消不删；确认后回到 12 行', async () => {
    const id = scratch.customerId;
    assert(id, '前置保存用例未成功');
    await click(page, `tr[data-id="${id}"] [data-testid="btn-delete-customer"]`);
    await waitSelector(page, '[data-confirm-ok]', '删除确认弹窗');
    const confirmText = String(await page.text('[data-modal-id] .confirm-text'));
    assert(confirmText.indexOf('E2E 测试客户（已改名）') !== -1, '删除确认文案应包含客户名称', confirmText);
    await click(page, '[data-confirm-cancel]');
    await waitGone(page, '[data-confirm-ok]', '取消后确认弹窗关闭');
    await sleep(120);
    await waitStoreCount(page, 13);
    assert(await page.exists(`tr[data-id="${id}"]`), '取消删除后客户行应保留');
    await click(page, `tr[data-id="${id}"] [data-testid="btn-delete-customer"]`);
    await waitSelector(page, '[data-confirm-ok]', '删除确认弹窗');
    await clearToasts(page);
    await click(page, '[data-confirm-ok]');
    await waitStoreCount(page, 12);
    await waitRowCount(page, 12);
    await waitToast(page, 'ok', '已删除客户');
    assert(!(await page.exists(`tr[data-id="${id}"]`)), '确认删除后客户行应消失');
    assertEq(await page.evaluate(`window.CustomerCRM.store.get(${J(id)})`), null, 'store 中客户应被删除');
  });

  await test('C10', '行内「记录跟进」按钮具备契约 data-testid=btn-add-followup', async () => {
    await waitRowCount(page, 12);
    assert(
      await page.exists('tr[data-testid="customer-row"] [data-testid="btn-add-followup"]'),
      'CONTRACT §8 要求客户行提供 data-testid="btn-add-followup"'
    );
    assert(await page.exists('[data-testid="row-toggle"]'), 'CONTRACT §8 要求 row-toggle 提供 data-testid');
  });

  await assertNoErrors(page, 'C11', '客户增删改全流程无 console / page 错误');
}

/* ================================================================== */
/* P. 产品库与产品选择器（v1.1 新增）                                   */
/* ================================================================== */

async function groupP(page) {
  setGroup('P', '产品库与产品选择器（v1.1）');

  await test('P1', '设置页产品库渲染：12 行、名称/型号/使用次数、表单与按钮 testid 齐全', async () => {
    await resetDemo(page);
    await goto(page, '#/settings');
    await waitSelector(page, '#product-form', '产品库表单');
    const lib = await page.evaluate(`(() => {
      const rows = Array.from(document.querySelectorAll('[data-testid="product-row"]'));
      const cell = function (row, selector) { const el = row.querySelector(selector); return el ? el.textContent.trim() : ''; };
      const target = rows.filter(function (row) { return cell(row, '.product-name') === '不锈钢保温杯' && cell(row, '.product-model') === '500ml'; })[0];
      const card = document.querySelector('[data-testid="product-library"]');
      const usage = window.CustomerCRM.store.productUsage();
      return {
        rows: rows.length,
        names: rows.map(function (row) { return cell(row, '.product-name'); }),
        models: rows.map(function (row) { return cell(row, '.product-model'); }),
        usageTexts: rows.map(function (row) { return cell(row, '.product-usage'); }),
        targetUsage: target ? cell(target, '.product-usage') : null,
        cardSub: card ? (card.querySelector('.card-sub') || {}).textContent || '' : '',
        hasForm: !!document.querySelector('[data-testid="form-product"]'),
        nameInput: !!document.querySelector('#product-form [name="productNewName"]'),
        modelInput: !!document.querySelector('#product-form [name="productNewModel"]'),
        addBtn: !!document.querySelector('[data-testid="btn-add-product"]'),
        seedBtn: !!document.querySelector('[data-action="seed-products"]'),
        deleteBtns: document.querySelectorAll('[data-testid="btn-delete-product"]').length,
        usage500: usage['不锈钢保温杯|500ml'] || 0
      };
    })()`);
    assertEq(lib.rows, 12, '产品库列表行数');
    assert(lib.cardSub.indexOf('共 12 个产品') !== -1, '产品库卡片副标题应显示 12 个产品', lib.cardSub);
    assertEq(lib.deleteBtns, 12, '每行都应有删除按钮');
    assert(lib.hasForm && lib.nameInput && lib.modelInput && lib.addBtn, '产品库表单/输入/按钮 testid 缺失', lib);
    assert(lib.seedBtn, '缺少「恢复预置产品」按钮');
    assertEq(lib.names.indexOf('不锈钢保温杯') !== -1, true, '产品库应包含预置产品「不锈钢保温杯」', lib.names);
    assertEq(lib.models.indexOf('500ml') !== -1, true, '产品库应包含预置型号「500ml」', lib.models);
    assertEq(lib.usage500, 1, 'productUsage() 中 不锈钢保温杯|500ml 应被 1 位客户使用');
    assertEq(lib.targetUsage, '1 位客户使用', '产品行使用次数文案');
  });

  await test('P2', '新增产品「测试产品/X1」→ 13 行、API 同步、客户表单下拉包含', async () => {
    await page.fill('productNewName', '测试产品');
    await page.fill('productNewModel', 'X1');
    await clearToasts(page);
    await click(page, '[data-testid="btn-add-product"]');
    await waitExpr(page, `document.querySelectorAll('[data-testid="product-row"]').length === 13`, '产品库新增后 13 行');
    await waitToast(page, 'ok', '已加入产品库');
    const api = await page.evaluate(`(() => ({
      count: window.CustomerCRM.store.listProducts().length,
      names: window.CustomerCRM.store.productNames(),
      models: window.CustomerCRM.store.productModels('测试产品')
    }))()`);
    assertEq(api.count, 13, 'store.listProducts() 条数');
    assert(api.names.indexOf('测试产品') !== -1, 'productNames() 应包含新增产品', api.names);
    assertEq(api.models.join(','), 'X1', 'productModels(测试产品)');
    await goto(page, '#/customers');
    await waitRowCount(page, 12);
    await click(page, '[data-testid="new-customer"]');
    await waitSelector(page, '[data-testid="product-name-select"]');
    const options = await page.evaluate(`Array.from(document.querySelectorAll('#customer-form [name="productName"] option')).map(function (option) { return option.value; })`);
    assert(options.indexOf('测试产品') !== -1, '新建客户产品名称下拉应包含新产品', options);
    await selectOption(page, '[data-testid="product-name-select"]', '测试产品');
    const modelOptions = await page.evaluate(`Array.from(document.querySelectorAll('#customer-form [name="productModel"] option')).map(function (option) { return option.value; })`);
    assert(modelOptions.indexOf('X1') !== -1, '选中新产品后型号下拉应包含 X1', modelOptions);
    await click(page, '[data-testid="btn-cancel-customer"]');
    await waitGone(page, '[data-testid="form-customer"]');
  });

  await test('P3', '删除产品「测试产品/X1」→ 回 12 行、下拉同步移除', async () => {
    await goto(page, '#/settings');
    await waitSelector(page, '[data-testid="product-row"]');
    const productId = await page.evaluate(`(() => {
      const rows = Array.from(document.querySelectorAll('[data-testid="product-row"]'));
      const target = rows.filter(function (row) {
        const name = row.querySelector('.product-name');
        return name && name.textContent.trim() === '测试产品';
      })[0];
      return target ? target.getAttribute('data-id') : null;
    })()`);
    assert(productId, '找不到新增的测试产品行');
    await clearToasts(page);
    await click(page, `[data-testid="product-row"][data-id="${productId}"] [data-testid="btn-delete-product"]`);
    await waitExpr(page, `document.querySelectorAll('[data-testid="product-row"]').length === 12`, '删除产品后回 12 行');
    const names = await page.evaluate('window.CustomerCRM.store.productNames()');
    assert(names.indexOf('测试产品') === -1, 'productNames() 应移除已删除产品', names);
    await goto(page, '#/customers');
    await click(page, '[data-testid="new-customer"]');
    await waitSelector(page, '[data-testid="product-name-select"]');
    const options = await page.evaluate(`Array.from(document.querySelectorAll('#customer-form [name="productName"] option')).map(function (option) { return option.value; })`);
    assert(options.indexOf('测试产品') === -1, '删除后客户表单下拉不应再包含该产品', options);
    await click(page, '[data-testid="btn-cancel-customer"]');
    await waitGone(page, '[data-testid="form-customer"]');
  });

  await test('P4', '删除预置产品后「恢复预置产品」→ 回到 12 行且 500ml 恢复', async () => {
    await goto(page, '#/settings');
    await waitSelector(page, '[data-testid="product-row"]');
    const defaultId = await page.evaluate(`(() => {
      const rows = Array.from(document.querySelectorAll('[data-testid="product-row"]'));
      const target = rows.filter(function (row) {
        const name = row.querySelector('.product-name');
        const model = row.querySelector('.product-model');
        return name && name.textContent.trim() === '不锈钢保温杯' && model && model.textContent.trim() === '500ml';
      })[0];
      return target ? target.getAttribute('data-id') : null;
    })()`);
    assert(defaultId, '找不到预置产品「不锈钢保温杯/500ml」');
    await click(page, `[data-testid="product-row"][data-id="${defaultId}"] [data-testid="btn-delete-product"]`);
    await waitExpr(page, `document.querySelectorAll('[data-testid="product-row"]').length === 11`, '删除后 11 行');
    await clearToasts(page);
    await click(page, '[data-action="seed-products"]');
    await waitExpr(page, `document.querySelectorAll('[data-testid="product-row"]').length === 12`, '恢复预置产品后 12 行');
    const restored = await page.evaluate(`(() => ({
      names: window.CustomerCRM.store.productNames(),
      models: window.CustomerCRM.store.productModels('不锈钢保温杯'),
      names_with_500: window.CustomerCRM.store.listProducts().filter(function (item) { return item.model === '500ml'; }).length
    }))()`);
    assert(restored.names.indexOf('不锈钢保温杯') !== -1, '恢复后应包含预置产品名', restored.names);
    assert(restored.models.indexOf('500ml') !== -1, '恢复后 不锈钢保温杯 应包含 500ml', restored.models);
    assert(restored.names_with_500 >= 1, '恢复后应包含 500ml 型号');
  });

  await test('P5', '客户表单产品联动：≥3 型号、app.selectProduct、自定义产品加入产品库', async () => {
    await goto(page, '#/customers');
    await waitRowCount(page, 12);
    await click(page, '[data-testid="new-customer"]');
    await waitSelector(page, '[data-testid="product-name-select"]');
    await selectOption(page, '[data-testid="product-name-select"]', '不锈钢保温杯');
    const modelOptions = await page.evaluate(`Array.from(document.querySelectorAll('#customer-form [name="productModel"] option')).map(function (option) { return option.value; }).filter(Boolean)`);
    assert(modelOptions.length >= 3, '型号下拉应出现 ≥3 个型号', modelOptions);
    await selectOption(page, '[data-testid="product-model-select"]', '500ml');
    assertEq(await page.evaluate(`document.querySelector('#customer-form [name="productModel"]').value`), '500ml', '可选中 500ml');
    const apiResult = await page.evaluate(`window.CustomerCRM.app.selectProduct('不锈钢保温杯', '750ml')`);
    assertEq(apiResult, true, 'app.selectProduct 返回值');
    assertEq(await page.evaluate(`document.querySelector('#customer-form [name="productModel"]').value`), '750ml', 'app.selectProduct 联动型号');
    await selectOption(page, '[data-testid="product-name-select"]', '__custom__');
    const customState = await page.evaluate(`(() => ({
      customVisible: !document.querySelector('[data-testid="product-custom-input"]').hidden,
      modelHidden: document.querySelector('#customer-form [name="productModel"]').hidden,
      addVisible: !document.querySelector('.product-add-btn').hidden
    }))()`);
    assert(customState.customVisible, '选择「手动输入其他产品」后自定义输入框应显示', customState);
    assert(customState.modelHidden, '自定义模式下型号下拉应隐藏', customState);
    assert(customState.addVisible, '自定义模式下「加入产品库」按钮应显示', customState);
    await page.fill('productCustom', '测试定制产品');
    await page.fill('productModelCustom', 'S-01');
    await clearToasts(page);
    await click(page, '[data-action="add-product-from-form"]');
    await waitToast(page, 'ok', '已加入产品库');
    const names = await page.evaluate('window.CustomerCRM.store.productNames()');
    assert(names.indexOf('测试定制产品') !== -1, '自定义产品应加入产品库', names);
    await click(page, '[data-testid="btn-cancel-customer"]');
    await waitGone(page, '[data-testid="form-customer"]');
  });

  await assertNoErrors(page, 'P6', '产品库与产品选择器流程无 console / page 错误');
}

/* ================================================================== */
/* D. 键盘快捷键                                                        */
/* ================================================================== */

async function groupD(page) {
  setGroup('D', '键盘快捷键');

  await test('D1', '「/」：从仪表盘跳转客户台账并聚焦搜索框', async () => {
    await resetDemo(page);
    await goto(page, '#/dashboard');
    await waitSelector(page, '[data-testid="stat-total"]');
    await page.evaluate('document.activeElement && document.activeElement.blur && document.activeElement.blur()');
    await pressKey(page, '/');
    await waitHash(page, '#/customers');
    await waitExpr(page, `document.activeElement && document.activeElement.id === 'customer-search'`, '「/」聚焦搜索框');
    await waitRowCount(page, 12);
  });

  await test('D2', '「N」：打开新建客户弹窗', async () => {
    await page.evaluate('document.activeElement && document.activeElement.blur && document.activeElement.blur()');
    await pressKey(page, 'n');
    await waitSelector(page, '[data-testid="form-customer"]', '「N」打开新建客户弹窗');
    assertEq(await page.evaluate(`document.getElementById('customer-form').getAttribute('data-mode')`), 'create', '新建弹窗 data-mode');
  });

  await test('D3', '「Esc」：关闭弹窗并解除 body 滚动锁', async () => {
    await pressKey(page, 'Escape');
    await waitGone(page, '[data-testid="form-customer"]', '「Esc」关闭弹窗');
    assertEq(await page.evaluate(`document.body.classList.contains('is-modal-open')`), false, 'Esc 后 body 应移除 is-modal-open');
  });

  await assertNoErrors(page, 'D4', '快捷键流程无 console / page 错误');
}

/* ================================================================== */
/* E. 智能识别                                                          */
/* ================================================================== */

async function groupE(page) {
  setGroup('E', '智能识别');

  await test('E1', '粘贴示例询盘 → 识别 13 项、命中 ≥11、含产品型号/主营业务', async () => {
    await resetDemo(page);
    await goto(page, '#/recognize');
    await waitSelector(page, '[data-testid="recognize-input"]');
    const sample = await page.evaluate(`(() => {
      const source = window.CrmParser.SAMPLE_TEXTS[0];
      const area = document.getElementById('recognize-input');
      const proto = Object.getPrototypeOf(area);
      const descriptor = Object.getOwnPropertyDescriptor(proto, 'value');
      if (descriptor && descriptor.set) descriptor.set.call(area, source.text);
      else area.value = source.text;
      area.dispatchEvent(new Event('input', { bubbles: true }));
      return { label: source.label, chars: area.value.length };
    })()`);
    assert(sample.chars > 50, '示例文本写入失败', sample);
    await clearToasts(page);
    await click(page, '[data-testid="btn-run-recognize"]');
    await waitSelector(page, '.recog-list', '识别结果列表出现');
    const dom = await page.evaluate(`(() => ({
      items: document.querySelectorAll('.recog-item').length,
      filled: document.querySelectorAll('.recog-item--filled').length,
      missing: document.querySelectorAll('.recog-item--missing').length,
      summary: (document.querySelector('.recog-summary .recog-count') || {}).textContent || '',
      fields: (function () { const out = {}; document.querySelectorAll('[data-recog-key]').forEach(function (el) { out[el.getAttribute('data-recog-key')] = el.value; }); return out; })(),
      evidenceCount: document.querySelectorAll('.recog-evidence').length,
      followupChips: document.querySelectorAll('.recog-result-body .followup-strip .followup-chip').length
    }))()`);
    assertEq(dom.items, 13, '识别结果项数（13 个字段）');
    assert(dom.filled >= 11, '识别命中字段数应 ≥11', dom.filled);
    assertEq(dom.filled + dom.missing, 13, '命中 + 未识别 = 13');
    const expected = await page.evaluate('window.CustomerCRM.app.recognize(window.CrmParser.SAMPLE_TEXTS[0].text)');
    assert(expected && expected.ok, 'CrmParser.recognize 应返回 ok');
    assertEq(dom.filled, expected.filled, '界面命中数 = parser.filled');
    const summaryMatch = dom.summary.match(/已识别\s*(\d+)\s*\/\s*(\d+)/);
    assert(summaryMatch, '识别摘要文案格式异常', dom.summary);
    assertEq(Number(summaryMatch[1]), dom.filled, '摘要计数 = DOM 命中项数');
    assertEq(Number(summaryMatch[2]), 13, '摘要分母应为 13');
    assert(dom.evidenceCount >= 11, '应展示原文证据（≥11 条）', dom.evidenceCount);
    assert(dom.followupChips >= 1, '应展示识别到的跟进记录', dom.followupChips);
    assert(dom.fields.productModel.indexOf('500ml') !== -1, '识别产品型号应含 500ml', dom.fields.productModel);
    assert(dom.fields.mainBusiness.indexOf('户外用品') !== -1, '识别主营业务应含 户外用品', dom.fields.mainBusiness);
    scratch.recognized = { dom, expected, sample };
  });

  await test('E2', '识别字段值与 CrmParser.recognize 直接调用一致（13 字段）', async () => {
    const state = scratch.recognized;
    assert(state, '前置识别用例未成功');
    const fields = state.expected.fields;
    for (const key of Object.keys(fields)) {
      assertEq(state.dom.fields[key], fields[key], '识别字段 ' + key + ' 与 parser 一致');
    }
    assertEq(state.expected.filled, 13, 'parser 对样例 A 应命中 13/13');
    assertEq(state.dom.fields.customerName, 'Michael Braun', '识别客户名称');
    assertEq(state.dom.fields.country, '德国', '识别国家');
    assertEq(state.dom.fields.firstInquiryAt, '2026-09-12', '识别首次询盘时间');
    assertEq(state.dom.fields.productModel, '500ml', '识别产品型号');
    assertEq(state.dom.fields.mainBusiness, '户外用品', '识别主营业务');
    assert(state.dom.fields.product.indexOf('保温杯') !== -1, '识别咨询产品应含保温杯', state.dom.fields.product);
  });

  await test('E3', '「填入表格并核对」→ 新建表单预填，产品下拉显示名称/型号', async () => {
    const state = scratch.recognized;
    assert(state, '前置识别用例未成功');
    await click(page, '[data-testid="btn-fill-customer"]');
    await waitSelector(page, '[data-testid="form-customer"]', '预填表单弹窗');
    const values = await page.evaluate(`(() => {
      const form = document.getElementById('customer-form');
      const read = function (name) { const el = form.querySelector('[name="' + name + '"]'); return el ? el.value : null; };
      return {
        customerName: read('customerName'),
        firstInquiryAt: read('firstInquiryAt'),
        mainBusiness: read('mainBusiness'),
        productName: read('productName'),
        productModel: read('productModel')
      };
    })()`);
    assertEq(await page.evaluate(`document.getElementById('customer-form').getAttribute('data-mode')`), 'create', '预填表单为新建模式');
    assertEq(values.customerName, state.expected.fields.customerName, '预填客户名称');
    assertEq(values.firstInquiryAt, state.expected.fields.firstInquiryAt, '预填首次询盘时间');
    assertEq(values.mainBusiness, state.expected.fields.mainBusiness, '预填主营业务');
    assertEq(values.productName, '不锈钢保温杯', '产品名称下拉应对齐产品库（不锈钢保温杯）');
    assertEq(values.productModel, '500ml', '产品型号下拉应显示 500ml');
    assert(await page.exists('.form-notice'), '预填表单应显示「已根据识别结果自动填入」提示');
    await click(page, '[data-testid="btn-cancel-customer"]');
    await waitGone(page, '[data-testid="form-customer"]', '取消预填表单');
  });

  await test('E4', '截图 screenshots/recognize.png（1440×900，识别结果态）', async () => {
    await waitSelector(page, '.recog-list', '识别结果仍在');
    await takeScreenshot(page, 'recognize.png');
    assetScreenshot('recognize.png', 1440, 900);
  });

  await test('E5', '「直接新增一行」→ store +1 且字段正确', async () => {
    const state = scratch.recognized;
    assert(state, '前置识别用例未成功');
    const beforeIds = await page.evaluate('window.CustomerCRM.store.list().map(function (record) { return record.id; })');
    await clearToasts(page);
    await click(page, '[data-testid="btn-add-recognized"]');
    await waitStoreCount(page, beforeIds.length + 1);
    await waitToast(page, 'ok', '已新增客户');
    const added = await page.evaluate(`(() => {
      const before = ${J(beforeIds)};
      return window.CustomerCRM.store.list().filter(function (record) { return before.indexOf(record.id) === -1; })[0] || null;
    })()`);
    assert(added, '直接新增一行后 store 未增加记录');
    assertEq(added.customerName, state.expected.fields.customerName, '新增记录客户名称');
    assertEq(added.firstInquiryAt, state.expected.fields.firstInquiryAt, '新增记录首次询盘时间');
    assertEq(added.country, state.expected.fields.country, '新增记录国家');
    assert(added.product.indexOf('保温杯') !== -1, '新增记录咨询产品应含保温杯', added.product);
    assertEq(added.productModel, '500ml', '新增记录产品型号');
    assertEq(added.mainBusiness, '户外用品', '新增记录主营业务');
    assert(added.followups.length >= 1, '新增记录应带识别到的跟进记录', added.followups);
    scratch.recognizeAddedId = added.id;
  });

  await test('E6', '空文本点击开始识别 → 仅提示错误，不产生结果', async () => {
    await page.fill('[data-testid="recognize-input"]', '');
    await clearToasts(page);
    await click(page, '[data-testid="btn-run-recognize"]');
    await waitToast(page, 'error', '请先粘贴一段询盘文本');
    assert(await page.exists('[data-testid="recognize-input"]'), '识别输入框应保留');
  });

  await assertNoErrors(page, 'E7', '智能识别流程无 console / page 错误');
}

/* ================================================================== */
/* F. 一键导出 Excel                                                    */
/* ================================================================== */

async function groupF(page) {
  setGroup('F', '一键导出 Excel');

  await test('F1', '无筛选导出：触发真实下载，文件名 客户跟进台账_YYYYMMDD_HHmm.xlsx', async () => {
    await resetDemo(page);
    await goto(page, '#/customers');
    await waitRowCount(page, 12);
    const expectedNames = await filteredNames(page);
    const expectedDetailRows = await page.evaluate('window.CustomerCRM.app.filtered().reduce(function (n, record) { return n + Math.min(3, (record.followups || []).length); }, 0)');
    const expectedFirst = await page.evaluate(`(() => {
      const record = window.CustomerCRM.app.filtered()[0];
      return { customerName: record.customerName, product: record.product, productModel: record.productModel, mainBusiness: record.mainBusiness };
    })()`);
    clearDownloads();
    await page.setDownloadBehavior(DOWNLOADS);
    await clearToasts(page);
    await click(page, '[data-testid="btn-export-excel"]');
    const filePath = await waitForDownload(/\.xlsx$/i);
    const filename = path.basename(filePath);
    assert(/^客户跟进台账_\d{8}_\d{4}\.xlsx$/.test(filename), '导出文件名格式应为 客户跟进台账_YYYYMMDD_HHmm.xlsx', filename);
    assert(filename.indexOf('客户跟进台账_' + localDateStamp() + '_') === 0, '导出文件名日期应为当天', { filename, today: localDateStamp() });
    await waitToast(page, 'ok', '已导出 12 位客户');
    scratch.export = { filePath, filename, expectedNames, expectedDetailRows, expectedFirst };
  });

  await test('F2', '文件结构：PK 头、体积、ZIP 部件完整、两个工作表 XML、行数正确', async () => {
    const state = scratch.export;
    assert(state, '前置导出用例未成功');
    const buffer = fs.readFileSync(state.filePath);
    assert(buffer.length > 3000, 'xlsx 体积异常', { bytes: buffer.length });
    assert(
      buffer[0] === 0x50 && buffer[1] === 0x4b && buffer[2] === 0x03 && buffer[3] === 0x04,
      'xlsx 文件头不是 PK\\x03\\x04',
      buffer.subarray(0, 8).toString('hex')
    );
    const entries = readZip(buffer);
    const requiredParts = [
      '[Content_Types].xml', '_rels/.rels', 'xl/workbook.xml', 'xl/_rels/workbook.xml.rels',
      'xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml', 'xl/styles.xml'
    ];
    for (const part of requiredParts) {
      assert(entries.has(part), 'xlsx 缺少 ZIP 部件：' + part, Array.from(entries.keys()));
    }
    const workbookXml = entries.get('xl/workbook.xml').toString('utf8');
    assert(workbookXml.indexOf('客户总表') !== -1, 'xl/workbook.xml 缺少工作表名「客户总表」');
    assert(workbookXml.indexOf('跟进明细') !== -1, 'xl/workbook.xml 缺少工作表名「跟进明细」');
    const sheet1Xml = entries.get('xl/worksheets/sheet1.xml').toString('utf8');
    const sheet2Xml = entries.get('xl/worksheets/sheet2.xml').toString('utf8');
    assertEq(countXmlRows(sheet1Xml), state.expectedNames.length + 1, '客户总表行数（数据行 + 表头）');
    assertEq(countXmlRows(sheet2Xml), state.expectedDetailRows + 1, '跟进明细行数（跟进行 + 表头）');
    assert(sheet1Xml.indexOf(state.expectedNames[0]) !== -1, '客户总表应包含客户中文名', state.expectedNames[0]);
  });

  await test('F3', 'openpyxl 校验：21 列表头（含产品型号/主营业务）/ 11 列明细 / 行数 / 冻结 / 自动筛选 / 新列有值', async () => {
    const state = scratch.export;
    assert(state, '前置导出用例未成功');
    if (!python) {
      console.log('  SKIP  F3  openpyxl 不可用，已由 F2 Node ZIP/XML 断言覆盖');
      return;
    }
    const data = runPythonXlsxCheck(state.filePath);
    assertEq(data.sheetnames.join('|'), '客户总表|跟进明细', '工作表名与顺序');
    const sheet1 = data.sheets['客户总表'];
    const sheet2 = data.sheets['跟进明细'];
    assertEq(sheet1.max_row, state.expectedNames.length + 1, 'openpyxl 客户总表行数');
    assertEq(sheet1.max_column, 21, '客户总表列数（21）');
    assertEq(sheet1.header.join('|'), EXPECTED_TOTAL_HEADERS.join('|'), '客户总表 21 列表头文本');
    assertEq(sheet1.first_row[3], state.expectedFirst.product, 'E 列咨询产品 = 导出首条记录产品');
    assertEq(sheet1.first_row[4], state.expectedFirst.productModel, 'E/F 列产品型号有值且正确');
    assertEq(sheet1.first_row[7], state.expectedFirst.mainBusiness, 'H 列主营业务有值且正确');
    assert(String(sheet1.first_row[4]).length > 0, '产品型号列有值');
    assert(String(sheet1.first_row[7]).length > 0, '主营业务列有值');
    assertEq(sheet1.freeze, 'A2', '客户总表冻结首行');
    assert(!!sheet1.auto_filter, '客户总表未设置自动筛选');
    assertEq(sheet1.names.join('|'), state.expectedNames.join('|'), '客户总表客户名顺序 = 导出记录顺序');
    assertEq(sheet2.max_row, state.expectedDetailRows + 1, 'openpyxl 跟进明细行数');
    assertEq(sheet2.max_column, 11, '跟进明细列数（11）');
    assertEq(sheet2.header.join('|'), EXPECTED_DETAIL_HEADERS.join('|'), '跟进明细 11 列表头文本');
  });

  await test('F4', '有筛选时只导出筛选结果（已成交 2 条 → 工作表 3 行）', async () => {
    await page.fill('[data-testid="filter-status"]', '已成交');
    await waitRowCount(page, 2);
    const filtered = await filteredNames(page);
    assertEq(filtered.length, 2, '筛选后记录数');
    const filteredDetailRows = await page.evaluate('window.CustomerCRM.app.filtered().reduce(function (n, record) { return n + Math.min(3, (record.followups || []).length); }, 0)');
    clearDownloads();
    await page.setDownloadBehavior(DOWNLOADS);
    await clearToasts(page);
    await click(page, '[data-testid="btn-export-excel"]');
    const filePath = await waitForDownload(/\.xlsx$/i);
    const filename = path.basename(filePath);
    assert(/^客户跟进台账_\d{8}_\d{4}\.xlsx$/.test(filename), '筛选导出文件名格式', filename);
    await waitToast(page, 'ok', '已导出 2 位客户');
    const entries = readZip(fs.readFileSync(filePath));
    const sheet1Xml = entries.get('xl/worksheets/sheet1.xml').toString('utf8');
    const sheet2Xml = entries.get('xl/worksheets/sheet2.xml').toString('utf8');
    assertEq(countXmlRows(sheet1Xml), filtered.length + 1, '筛选导出客户总表行数 = 筛选结果 + 表头');
    assertEq(countXmlRows(sheet2Xml), filteredDetailRows + 1, '筛选导出跟进明细行数');
    if (python) {
      const data = runPythonXlsxCheck(filePath);
      assertEq(data.sheets['客户总表'].names.join('|'), filtered.join('|'), 'openpyxl 读回的客户名 = 当前筛选结果');
      assertEq(data.sheets['客户总表'].max_column, 21, '筛选导出客户总表列数（21）');
      assertEq(data.sheets['客户总表'].max_row, filtered.length + 1, '筛选导出客户总表行数（openpyxl）');
      assertEq(data.sheets['跟进明细'].max_column, 11, '筛选导出明细列数（11）');
    }
    await click(page, '[data-testid="filter-reset"]');
    await waitRowCount(page, 12);
  });

  await assertNoErrors(page, 'F5', '导出流程无 console / page 错误');
}

/* ================================================================== */
/* G. 设置与数据管理                                                    */
/* ================================================================== */

async function groupG(page) {
  setGroup('G', '设置、JSON 备份 / 恢复与数据管理');

  await test('G1', '设置页渲染：默认阈值 7/14，数据管理按钮齐全', async () => {
    await resetDemo(page);
    await goto(page, '#/settings');
    await waitSelector(page, '#setting-warnDays');
    const info = await page.evaluate(`(() => ({
      warn: (document.getElementById('setting-warnDays') || {}).value,
      stale: (document.getElementById('setting-staleDays') || {}).value,
      seed: !!document.querySelector('[data-testid="btn-seed-demo"]'),
      clear: !!document.querySelector('[data-testid="btn-clear-all"]'),
      exportJson: !!document.querySelector('[data-action="export-json"]'),
      importJson: !!document.querySelector('[data-action="import-json"]')
    }))()`);
    assertEq(info.warn, '7', '默认待跟进阈值');
    assertEq(info.stale, '14', '默认超期阈值');
    assert(info.seed && info.clear, '数据管理按钮 data-testid 缺失', info);
    assert(info.exportJson && info.importJson, 'JSON 备份 / 恢复按钮缺失', info);
  });

  await test('G2', '截图 screenshots/settings.png（1440×900）', async () => {
    await takeScreenshot(page, 'settings.png');
    assetScreenshot('settings.png', 1440, 900);
  });

  await test('G3', '保存阈值 warnDays=3 / staleDays=5 并持久化；非法阈值被拒绝', async () => {
    await page.fill('#setting-warnDays', '3');
    await page.fill('#setting-staleDays', '5');
    await clearToasts(page);
    await click(page, '[data-action="save-settings"]');
    await waitExpr(page, `window.CustomerCRM.store.settings().warnDays === 3 && window.CustomerCRM.store.settings().staleDays === 5`, '设置写入 store');
    await waitToast(page, 'ok', '设置已保存');
    await page.fill('#setting-warnDays', '30');
    await page.fill('#setting-staleDays', '10');
    await clearToasts(page);
    await click(page, '[data-action="save-settings"]');
    await waitToast(page, 'error', '超期阈值');
    assertEq(await page.evaluate('window.CustomerCRM.store.settings().warnDays'), 3, '非法阈值不应写入 store');
    assertEq(await page.evaluate('window.CustomerCRM.store.settings().staleDays'), 5, '非法阈值不应写入 store');
  });

  await test('G4', '仪表盘统计随阈值变化：4 卡 DOM = stats()，超期数增加，hint 文案更新', async () => {
    await goto(page, '#/dashboard');
    await waitSelector(page, '[data-testid="stat-overdue"]');
    const stats = await page.evaluate('window.CustomerCRM.app.stats()');
    const dom = await readStatsDom(page);
    assertEq(dom.total, String(stats.total), '客户总数卡 = stats.total');
    assertEq(dom.deal, String(stats.deal), '已成交卡 = stats.deal');
    assertEq(dom.open, String(stats.open), '待跟进卡 = stats.open');
    assertEq(dom.overdue, String(stats.overdue), '超期未跟进卡 = stats.overdue');
    const context = await page.evaluate(`(() => {
      const schema = window.CRM_SCHEMA;
      const today = schema.todayISO();
      return {
        today: today,
        records: window.CustomerCRM.store.list().map(function (record) {
          return {
            dealStatus: record.dealStatus,
            followupStatus: record.followupStatus,
            last: (record.followups && record.followups[0] && record.followups[0].date) || null
          };
        })
      };
    })()`);
    const daysBetween = (fromISO, toISO) => {
      const from = Date.UTC(Number(fromISO.slice(0, 4)), Number(fromISO.slice(5, 7)) - 1, Number(fromISO.slice(8, 10)));
      const to = Date.UTC(Number(toISO.slice(0, 4)), Number(toISO.slice(5, 7)) - 1, Number(toISO.slice(8, 10)));
      return Math.round((to - from) / 86400000);
    };
    const isOpen = (record) => record.dealStatus !== '已成交' && record.followupStatus !== '已流失';
    const expectedOpen = context.records.filter(isOpen).length;
    const expectedOverdue5 = context.records.filter((record) => isOpen(record) && record.last && daysBetween(record.last, context.today) >= 5).length;
    const expectedOverdue14 = context.records.filter((record) => isOpen(record) && record.last && daysBetween(record.last, context.today) >= 14).length;
    assertEq(stats.open, expectedOpen, 'stats.open = 未成交且未流失记录数');
    assertEq(Number(dom.overdue), expectedOverdue5, 'staleDays=5 时超期未跟进数量');
    assert(expectedOverdue5 > expectedOverdue14, 'staleDays 由 14 调至 5 后超期数量应增加', { expectedOverdue5, expectedOverdue14 });
    const hints = await readStatHints(page);
    assert(String(hints['stat-open']).indexOf('≥3') !== -1, '待跟进卡片 hint 应显示 warnDays=3', hints['stat-open']);
    assert(String(hints['stat-overdue']).indexOf('≥5') !== -1, '超期卡片 hint 应显示 staleDays=5', hints['stat-overdue']);
  });

  await test('G5', '客户台账 age 颜色分布 = store.deriveAll()（当前阈值）', async () => {
    await goto(page, '#/customers');
    await waitRowCount(page, 12);
    const domTones = await page.evaluate(`(() => {
      const counts = { ok: 0, warn: 0, danger: 0, none: 0 };
      document.querySelectorAll('tr[data-testid="customer-row"] td[data-col="age"] .age').forEach(function (el) {
        ['ok', 'warn', 'danger', 'none'].forEach(function (tone) {
          if (el.classList.contains('age--' + tone)) counts[tone] += 1;
        });
      });
      return counts;
    })()`);
    const expectedTones = await page.evaluate(`window.CustomerCRM.store.deriveAll().reduce(function (acc, record) {
      acc[record.ageTone] = (acc[record.ageTone] || 0) + 1;
      return acc;
    }, { ok: 0, warn: 0, danger: 0, none: 0 })`);
    assertEq(JSON.stringify(domTones), JSON.stringify(expectedTones), 'age 颜色分布 = deriveAll(当前阈值)');
    assertEq(domTones.ok + domTones.warn + domTones.danger + domTones.none, 12, 'age 单元格数量');
  });

  await test('G6', '导出 JSON 备份：文件名、合法 JSON、12 条客户、产品库与设置', async () => {
    await goto(page, '#/settings');
    await waitSelector(page, '[data-action="export-json"]');
    clearDownloads();
    await page.setDownloadBehavior(DOWNLOADS);
    await clearToasts(page);
    await click(page, '[data-action="export-json"]');
    const filePath = await waitForDownload(/\.json$/i);
    const filename = path.basename(filePath);
    assert(/^客户通备份_\d{8}_\d{4}\.json$/.test(filename), 'JSON 备份文件名格式', filename);
    const text = fs.readFileSync(filePath, 'utf8');
    let payload = null;
    try { payload = JSON.parse(text); } catch (err) { fail('JSON 备份无法解析', String(err && err.message)); }
    assertEq(payload.app, 'CustomerCRM', '备份 app 标识');
    assertEq(payload.version, '1.1.0', '备份版本号');
    assertEq(payload.customers.length, 12, '备份客户数');
    assert(payload.customers.some((record) => record.customerName === 'Michael Braun'), '备份应包含示例客户 Michael Braun');
    assertEq(Array.isArray(payload.products) ? payload.products.length : -1, 12, '备份应包含 12 条产品库数据');
    assert(payload.products.some((product) => product.name === '不锈钢保温杯' && product.model === '500ml'), '备份产品库应含 不锈钢保温杯/500ml');
    assert(payload.settings && typeof payload.settings.warnDays === 'number', '备份应包含设置', payload.settings);
    scratch.backupPath = path.join(TEMP, 'customer-crm-backup.json');
    fs.copyFileSync(filePath, scratch.backupPath);
  });

  await test('G7', '清空全部数据 → 空状态引导 + 仪表盘统计归零（产品库保留）', async () => {
    await clearToasts(page);
    await click(page, '[data-testid="btn-clear-all"]');
    await waitSelector(page, '[data-confirm-ok]', '清空确认弹窗');
    const confirmText = String(await page.text('[data-modal-id] .confirm-text'));
    assert(confirmText.indexOf('删除本机保存的全部客户记录') !== -1, '清空确认文案', confirmText);
    await click(page, '[data-confirm-ok]');
    await waitStoreCount(page, 0);
    assertEq(await page.evaluate('window.CustomerCRM.store.listProducts().length'), 12, '清空客户数据后产品库应保留 12 条');
    await goto(page, '#/customers');
    await waitSelector(page, '.empty', '空状态引导');
    const empty = await page.evaluate(`(() => ({
      title: (document.querySelector('.empty-title') || {}).textContent || '',
      text: (document.querySelector('.empty-text') || {}).textContent || '',
      actions: Array.from(document.querySelectorAll('.empty-actions .btn')).map(function (btn) { return btn.textContent.trim(); })
    }))()`);
    assertEq(empty.title.trim(), '还没有客户记录', '空状态标题');
    assert(empty.text.length > 0, '空状态说明文案为空');
    const actions = empty.actions.join('|');
    assert(actions.indexOf('智能识别') !== -1 && actions.indexOf('新建客户') !== -1 && actions.indexOf('载入示例数据') !== -1, '空状态引导按钮不完整', actions);
    assertEq(await rowCount(page), 0, '空状态下数据行数量');
    await goto(page, '#/dashboard');
    await waitSelector(page, '[data-testid="stat-total"]');
    assertEq(String(await page.text('[data-testid="stat-total"]')).trim(), '0', '清空后仪表盘客户总数 = 0');
  });

  await test('G8', '从 JSON 备份恢复 → 12 条客户回表', async () => {
    const backupPath = scratch.backupPath;
    assert(backupPath && fs.existsSync(backupPath), '前置 JSON 备份用例未成功');
    await goto(page, '#/settings');
    await waitSelector(page, '[data-action="import-json"]');
    await clearToasts(page);
    await click(page, '[data-action="import-json"]');
    await waitExpr(page, `!!document.querySelector('input[type="file"]')`, '隐藏文件选择框出现');
    const backupText = fs.readFileSync(backupPath, 'utf8');
    const injected = await page.evaluate(`(() => {
      const input = document.querySelector('input[type="file"]');
      if (!input) return false;
      const transfer = new DataTransfer();
      transfer.items.add(new File([${J(backupText)}], 'customer-crm-backup.json', { type: 'application/json' }));
      input.files = transfer.files;
      input.dispatchEvent(new Event('change', { bubbles: true }));
      return true;
    })()`);
    assert(injected, '无法向文件选择框注入备份文件');
    await waitStoreCount(page, 12, 15000);
    await waitToast(page, 'ok', '已从备份导入');
    assertEq(await page.evaluate('window.CustomerCRM.store.listProducts().length'), 12, '恢复后产品库应为 12 条');
    await goto(page, '#/customers');
    await waitRowCount(page, 12);
    const restored = await page.evaluate(`window.CustomerCRM.store.list().filter(function (record) { return record.customerName === 'Michael Braun'; }).length`);
    assertEq(restored, 1, '恢复后应包含示例客户 Michael Braun');
    await page.evaluate(`(() => { const input = document.querySelector('input[type="file"]'); if (input && input.parentNode) input.parentNode.removeChild(input); })()`);
  });

  await test('G9', '再次清空后「载入示例数据」恢复 12 条', async () => {
    await goto(page, '#/settings');
    await waitSelector(page, '[data-testid="btn-clear-all"]');
    await click(page, '[data-testid="btn-clear-all"]');
    await waitSelector(page, '[data-confirm-ok]', '清空确认弹窗');
    await click(page, '[data-confirm-ok]');
    await waitStoreCount(page, 0);
    await goto(page, '#/customers');
    await waitSelector(page, '.empty', '空状态引导');
    await clearToasts(page);
    await click(page, '[data-testid="btn-seed-demo"]');
    await waitStoreCount(page, 12);
    await waitRowCount(page, 12);
    await waitToast(page, 'ok', '已载入 12 位示例客户');
    const footer = String(await page.text('.table-count'));
    assert(footer.indexOf('12') !== -1, '表格底部计数应含 12', footer);
  });

  await assertNoErrors(page, 'G10', '设置与数据管理流程无 console / page 错误');
}

/* ================================================================== */
/* H. 响应式 / 抽屉 / 打印                                              */
/* ================================================================== */

async function groupH(page) {
  setGroup('H', '响应式、侧栏抽屉与打印媒体');

  await test('H1', '1440×900 宽版台账：内容宽 ≥1140、表格 min-width ≥1328、15 列、3 列折叠、产品/主营可见', async () => {
    await resetDemo(page);
    await page.setViewport(1440, 900);
    await goto(page, '#/customers');
    await waitRowCount(page, 12);
    await assertNoOverflow(page, '1440');
    const info = await sidebarRect(page);
    assert(info.left >= 0 && info.width > 200, '1440 侧栏应完整可见', info);
    const layout = await page.evaluate(`(() => {
      const inner = document.querySelector('.view-inner');
      const table = document.querySelector('.data-table--customers');
      const row = document.querySelector('tr[data-testid="customer-row"]');
      const visible = {};
      ['firstInquiryAt', 'age', 'customerName', 'product', 'country', 'companyName', 'mainBusiness', 'followupStatus', 'dealStatus', 'coreReason'].forEach(function (key) {
        const td = document.querySelector('tr[data-testid="customer-row"] td[data-col="' + key + '"]');
        visible[key] = td ? getComputedStyle(td).display !== 'none' : null;
      });
      const hidden = {};
      ['foundedAt', 'teamSize', 'annualRevenue'].forEach(function (key) {
        const td = document.querySelector('tr[data-testid="customer-row"] td[data-col="' + key + '"]');
        hidden[key] = td ? getComputedStyle(td).display === 'none' : null;
      });
      const actions = document.querySelector('tr[data-testid="customer-row"] td.col-actions');
      return {
        bodyClass: document.body.className,
        innerWidth: inner ? Math.round(inner.getBoundingClientRect().width) : 0,
        tableMinWidth: table ? parseFloat(getComputedStyle(table).minWidth) : 0,
        tds: row ? row.querySelectorAll('td').length : 0,
        visible: visible,
        hidden: hidden,
        actionsVisible: actions ? getComputedStyle(actions).display !== 'none' : null
      };
    })()`);
    assert(/route-customers/.test(layout.bodyClass), '台账页应带 body.route-customers', layout.bodyClass);
    assert(layout.innerWidth >= 1140, '1440 台账内容宽度应 ≥1140（契约实测 1152）', layout.innerWidth);
    assert(layout.tableMinWidth >= 1328, '1440 表格 min-width 应 ≥1328', layout.tableMinWidth);
    assertEq(layout.tds, 15, '数据行 td 数量（15 列仍在 DOM）');
    for (const key of ['firstInquiryAt', 'age', 'customerName', 'product', 'country', 'companyName', 'mainBusiness', 'followupStatus', 'dealStatus', 'coreReason']) {
      assertEq(layout.visible[key], true, '1440 应显示列 ' + key);
    }
    for (const key of ['foundedAt', 'teamSize', 'annualRevenue']) {
      assertEq(layout.hidden[key], true, '1440 应隐藏列（改在行展开详情展示）' + key);
    }
    assertEq(layout.actionsVisible, true, '1440 操作列应可见');
    // 低优先字段改在行展开详情的 .row-detail-meta 中展示
    const firstId = await page.evaluate(`document.querySelector('tr[data-testid="customer-row"]').getAttribute('data-id')`);
    const record = await page.evaluate(`window.CustomerCRM.store.get(${J(firstId)})`);
    await click(page, `tr[data-id="${firstId}"] [data-testid="row-toggle"]`);
    await waitExpr(page, `!document.querySelector('tr.row-detail[data-detail="${firstId}"]').hidden`, '行详情展开');
    const meta = await page.evaluate(`(() => { const el = document.querySelector('tr.row-detail[data-detail="${firstId}"] .row-detail-meta'); return el ? el.textContent : ''; })()`);
    for (const label of ['产品型号', '主营业务', '成立时间', '团队规模', '年营收', '首次询盘', '最近更新']) {
      assert(meta.indexOf(label) !== -1, '行展开详情应含「' + label + '」', meta);
    }
    assert(meta.indexOf(record.productModel) !== -1, '行展开详情应含产品型号值', record.productModel);
    assert(meta.indexOf(record.mainBusiness) !== -1, '行展开详情应含主营业务值', record.mainBusiness);
    await click(page, `tr[data-id="${firstId}"] [data-testid="row-toggle"]`);
  });

  await test('H1b', '宽版对比：台账 .view-inner 宽度明显大于仪表盘', async () => {
    const customersWidth = await page.evaluate(`Math.round(document.querySelector('.view-inner').getBoundingClientRect().width)`);
    await goto(page, '#/dashboard');
    await waitSelector(page, '[data-testid="stat-total"]');
    const dashboardWidth = await page.evaluate(`Math.round(document.querySelector('.view-inner').getBoundingClientRect().width)`);
    assert(customersWidth >= 1140, '台账内容宽度 ≥1140', customersWidth);
    assert(customersWidth > dashboardWidth, '台账宽版内容宽度应大于仪表盘', { customersWidth, dashboardWidth });
    await goto(page, '#/customers');
    await waitRowCount(page, 12);
  });

  await test('H2', '1024×800：无横向溢出、侧栏收起、隐藏 3 列，主营业务可见', async () => {
    await page.setViewport(1024, 800);
    await goto(page, '#/customers');
    await waitRowCount(page, 12);
    await waitExpr(page, `document.querySelector('.sidebar').getBoundingClientRect().right <= 1`, '1024 侧栏收起动画结束', 5000);
    await assertNoOverflow(page, '1024');
    const info = await sidebarRect(page);
    assert(info.right <= 1, '1024 侧栏应默认收起（translateX(-100%)）', info);
    for (const key of ['foundedAt', 'teamSize', 'annualRevenue']) {
      assertEq(await columnVisible(page, key), false, '1024 应隐藏列 ' + key);
    }
    for (const key of ['product', 'mainBusiness', 'companyName', 'coreReason']) {
      assertEq(await columnVisible(page, key), true, '1024 应显示列 ' + key);
    }
  });

  await test('H3', '768×900 边界 + 767（<768）：隐藏产品/主营业务/公司名/核心原因/距上次跟进', async () => {
    await page.setViewport(768, 900);
    await goto(page, '#/customers');
    await waitRowCount(page, 12);
    await waitExpr(page, `document.querySelector('.sidebar').getBoundingClientRect().right <= 1`, '768 侧栏收起动画结束', 5000);
    await assertNoOverflow(page, '768');
    const info = await sidebarRect(page);
    assert(info.right <= 1, '768 侧栏应默认收起', info);
    // 断点为 <768：768 本身仍显示产品/主营业务列（边界语义）
    assertEq(await columnVisible(page, 'product'), true, '768 应显示咨询产品列（媒体查询为 <768）');
    assertEq(await columnVisible(page, 'mainBusiness'), true, '768 应显示主营业务列');
    assertEq(await columnVisible(page, 'age'), true, '768 应显示距上次跟进列');
    // 收窄到 767，验证 <768 断点隐藏
    await page.setViewport(767, 900);
    await waitExpr(page, `getComputedStyle(document.querySelector('th[data-col="product"]')).display === 'none'`, '767 隐藏产品列', 5000);
    for (const key of ['product', 'mainBusiness', 'companyName', 'coreReason', 'age']) {
      assertEq(await columnVisible(page, key), false, '767 应隐藏列 ' + key);
    }
    assertEq(await columnVisible(page, 'customerName'), true, '767 应显示客户名称列');
    await assertNoOverflow(page, '767');
  });

  await test('H4', '390×844：无横向溢出、侧栏收起、菜单按钮可见、移动端折叠列', async () => {
    await page.setViewport(390, 844, { mobile: true });
    await goto(page, '#/customers');
    await waitRowCount(page, 12);
    await waitExpr(page, `document.querySelector('.sidebar').getBoundingClientRect().right <= 1`, '390 侧栏收起动画结束', 5000);
    await assertNoOverflow(page, '390');
    const info = await sidebarRect(page);
    assert(info.right <= 1, '390 侧栏应默认收起', info);
    const toggleVisible = await page.evaluate(`(function () {
      const el = document.querySelector('.menu-toggle');
      if (!el) return null;
      const style = getComputedStyle(el);
      const rect = el.getBoundingClientRect();
      return style.display !== 'none' && style.visibility !== 'hidden' && rect.width > 0;
    })()`);
    assertEq(toggleVisible, true, '390 菜单按钮应可见');
    for (const key of ['product', 'mainBusiness', 'companyName', 'coreReason']) {
      assertEq(await columnVisible(page, key), false, '390 应隐藏列 ' + key);
    }
    assertEq(await columnVisible(page, 'customerName'), true, '390 应显示客户名称列');
  });

  await test('H5', '390 侧栏抽屉：菜单按钮打开（遮罩可见）→ 点遮罩关闭', async () => {
    await waitExpr(page, `document.querySelector('.sidebar').getBoundingClientRect().right <= 1`, '抽屉初始收起', 5000);
    await click(page, '.menu-toggle');
    await waitExpr(page, `document.getElementById('sidebar').classList.contains('is-open')`, '侧栏抽屉打开');
    await waitExpr(page, `document.querySelector('.sidebar').getBoundingClientRect().left >= -1`, '抽屉滑入动画结束', 5000);
    const openState = await page.evaluate(`(() => {
      const sidebar = document.getElementById('sidebar');
      const backdrop = document.getElementById('sidebar-backdrop');
      const rect = sidebar.getBoundingClientRect();
      const backdropStyle = getComputedStyle(backdrop);
      return {
        left: rect.left,
        width: rect.width,
        backdropVisible: backdropStyle.visibility !== 'hidden' && Number(backdropStyle.opacity) > 0.1,
        ariaExpanded: document.querySelector('.menu-toggle').getAttribute('aria-expanded')
      };
    })()`);
    assert(openState.left >= -1 && openState.width > 200, '抽屉打开后侧栏应滑入视口', openState);
    assert(openState.backdropVisible, '抽屉打开后遮罩应可见', openState);
    assertEq(openState.ariaExpanded, 'true', '菜单按钮 aria-expanded');
    await click(page, '.sidebar-backdrop');
    await waitExpr(page, `!document.getElementById('sidebar').classList.contains('is-open')`, '点遮罩后抽屉关闭');
    await waitExpr(page, `document.querySelector('.sidebar').getBoundingClientRect().right <= 1`, '抽屉滑出动画结束', 5000);
    const closedState = await sidebarRect(page);
    assert(closedState.right <= 1, '抽屉关闭后侧栏应在视口外', closedState);
  });

  await test('H6', '打印媒体（台账）：侧栏/工具栏/操作列隐藏、全部业务列恢复（含主营业务）', async () => {
    await page.setViewport(1440, 900);
    await goto(page, '#/customers');
    await waitRowCount(page, 12);
    await page.emulateMedia('print');
    await waitExpr(page, `getComputedStyle(document.querySelector('.sidebar')).display === 'none'`, '打印样式生效');
    const printState = await page.evaluate(`(() => {
      const read = function (selector) {
        const el = document.querySelector(selector);
        if (!el) return { missing: true };
        const style = getComputedStyle(el);
        return { display: style.display, visibility: style.visibility, offsetParent: el.offsetParent !== null };
      };
      const thead = document.querySelector('thead');
      return {
        sidebar: read('.sidebar'),
        topbar: read('.topbar'),
        toolbar: read('.toolbar'),
        pageHeadActions: read('.page-head-actions'),
        colActions: read('th.col-actions'),
        toastStack: read('.toast-stack'),
        theadDisplay: thead ? getComputedStyle(thead).display : null,
        viewVisible: read('#view-inner')
      };
    })()`);
    for (const key of ['sidebar', 'topbar', 'toolbar', 'pageHeadActions', 'colActions', 'toastStack']) {
      assert(printState[key] && printState[key].display === 'none', '打印媒体下应隐藏：' + key, printState[key]);
    }
    assert(printState.viewVisible && printState.viewVisible.display !== 'none', '打印媒体下内容应保持可见', printState.viewVisible);
    // 打印时恢复所有业务列（含 <=1599 折叠的 3 列与主营业务）
    for (const key of ['foundedAt', 'teamSize', 'annualRevenue', 'mainBusiness', 'product', 'companyName', 'coreReason']) {
      assertEq(await columnVisible(page, key), true, '打印媒体下应恢复列 ' + key);
    }
    const expandHidden = await page.evaluate(`(function () { const th = document.querySelector('th.col-expand'); return th ? getComputedStyle(th).display === 'none' : null; })()`);
    assertEq(expandHidden, true, '打印媒体下展开列应隐藏');
    const actionHidden = await page.evaluate(`(function () { const th = document.querySelector('th.col-actions'); return th ? getComputedStyle(th).display === 'none' : null; })()`);
    assertEq(actionHidden, true, '打印媒体下操作列应隐藏');
    assertEq(printState.theadDisplay, 'table-header-group', '打印媒体下 thead 应重复表头');
    await page.emulateMedia('');
    await waitExpr(page, `getComputedStyle(document.querySelector('.sidebar')).display !== 'none'`, '恢复 screen 媒体');
    await clearToasts(page);
  });

  await test('H6b', '打印媒体（仪表盘）：国家分布扇形图保留可见', async () => {
    await goto(page, '#/dashboard');
    await waitSelector(page, '[data-testid="country-pie"]');
    await page.emulateMedia('print');
    await waitExpr(page, `getComputedStyle(document.querySelector('.sidebar')).display === 'none'`, '仪表盘打印样式生效');
    const piePrint = await page.evaluate(`(() => {
      const root = document.querySelector('[data-testid="country-pie"]');
      if (!root) return null;
      const style = getComputedStyle(root);
      const rect = root.getBoundingClientRect();
      return {
        display: style.display,
        visibility: style.visibility,
        width: Math.round(rect.width),
        height: Math.round(rect.height),
        slices: root.querySelectorAll('.pie-slice').length,
        total: (root.querySelector('.pie-total') || {}).textContent || ''
      };
    })()`);
    assert(piePrint, '打印媒体下找不到扇形图容器');
    assert(piePrint.display !== 'none' && piePrint.visibility !== 'hidden', '打印媒体下扇形图应保留可见', piePrint);
    assert(piePrint.width > 100 && piePrint.height > 100, '打印媒体下扇形图应有实际尺寸', piePrint);
    assertEq(piePrint.slices, 7, '打印媒体下扇形图切片数');
    assertEq(String(piePrint.total).trim(), '12', '打印媒体下扇形图中心数');
    await page.emulateMedia('');
    await waitExpr(page, `getComputedStyle(document.querySelector('.sidebar')).display !== 'none'`, '恢复 screen 媒体');
    await clearToasts(page);
  });

  await test('H7', '截图 mobile-dashboard.png 与 mobile-customers.png（390×844）', async () => {
    await page.setViewport(390, 844, { mobile: true });
    await goto(page, '#/dashboard');
    await waitSelector(page, '[data-testid="stat-total"]');
    await takeScreenshot(page, 'mobile-dashboard.png');
    assetScreenshot('mobile-dashboard.png', 390, 844);
    await goto(page, '#/customers');
    await waitRowCount(page, 12);
    await takeScreenshot(page, 'mobile-customers.png');
    assetScreenshot('mobile-customers.png', 390, 844);
    await page.setViewport(1440, 900);
  });

  await assertNoErrors(page, 'H8', '响应式 / 打印流程无 console / page 错误');
}

/* ================================================================== */
/* I. 数据损坏兜底                                                      */
/* ================================================================== */

async function groupI(page) {
  setGroup('I', '数据损坏与旧数据升级兜底');

  await test('I1', 'localStorage 非法 JSON → 刷新后不白屏，自动重建 12 条客户 + 12 条产品', async () => {
    await resetDemo(page);
    await page.setViewport(1440, 900);
    await page.evaluate(`localStorage.setItem('customer-crm.v1', '{bad json')`);
    page.clearErrors();
    await page.reload();
    await page.waitForApp({ timeout: 25000 });
    await waitExpr(page, `!!document.querySelector('#view-inner .page')`, '损坏数据后应用仍能渲染页面', 15000);
    await goto(page, '#/dashboard');
    await waitExpr(page, `!!document.querySelector('[data-testid="stat-total"]')`, '损坏数据后仍渲染仪表盘', 15000);
    const health = await page.evaluate(`(() => {
      let parsed = null;
      let error = '';
      try { parsed = JSON.parse(localStorage.getItem('customer-crm.v1')); } catch (err) { error = String(err && err.message); }
      return {
        total: window.CustomerCRM.store.list().length,
        products: window.CustomerCRM.store.listProducts().length,
        domTotal: (document.querySelector('[data-testid="stat-total"]') || {}).textContent || '',
        viewLength: (document.getElementById('view-inner') || {}).innerHTML.length || 0,
        storedCount: parsed && parsed.customers ? parsed.customers.length : null,
        storedProducts: parsed && parsed.products ? parsed.products.length : null,
        storedVersion: parsed ? parsed.version : null,
        error: error
      };
    })()`);
    assertEq(health.total, 12, '损坏后示例数据条数（store）');
    assertEq(health.products, 12, '损坏后产品库条数（store）');
    assertEq(health.domTotal.trim(), '12', '损坏后仪表盘客户总数');
    assert(health.viewLength > 1000, '页面疑似白屏（#view-inner 内容过短）', health.viewLength);
    assertEq(health.storedCount, 12, '损坏数据应被重写为合法 payload');
    assertEq(health.storedProducts, 12, 'payload 中产品库应重建为 12 条');
    assertEq(health.error, '', 'localStorage 应恢复为合法 JSON');
    const errors = page.errors();
    assert(errors.length === 0, '损坏兜底过程不应产生 pageerror / console error', errors.slice(0, 4));
  });

  await test('I2', 'v1.0 旧 payload（无 products）升级 → 自动补齐 12 条预置产品', async () => {
    const today = await page.evaluate('window.CRM_SCHEMA.todayISO()');
    const legacy = {
      version: 1,
      savedAt: new Date().toISOString(),
      settings: { warnDays: 7, staleDays: 14, seeded: true },
      customers: [{
        id: 'c_legacy001',
        firstInquiryAt: today,
        customerName: '旧版本客户',
        product: '不锈钢保温杯',
        productModel: '500ml',
        country: '德国',
        companyName: 'Legacy GmbH',
        mainBusiness: '旧数据主营业务',
        foundedAt: '2000',
        teamSize: '10 人',
        annualRevenue: '100 万欧元',
        followupStatus: '新询盘',
        dealStatus: '未成交',
        coreReason: '',
        followups: [],
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString()
      }]
    };
    await page.evaluate(`localStorage.setItem('customer-crm.v1', ${J(JSON.stringify(legacy))})`);
    page.clearErrors();
    await page.reload();
    await page.waitForApp({ timeout: 25000 });
    await waitExpr(page, `!!document.querySelector('#view-inner .page')`, '旧数据升级后页面渲染', 15000);
    const upgraded = await page.evaluate(`(() => {
      let parsed = null;
      try { parsed = JSON.parse(localStorage.getItem('customer-crm.v1')); } catch (err) { parsed = null; }
      return {
        customers: window.CustomerCRM.store.list().length,
        products: window.CustomerCRM.store.listProducts().length,
        names: window.CustomerCRM.store.productNames().length,
        storedProducts: parsed && parsed.products ? parsed.products.length : null
      };
    })()`);
    assertEq(upgraded.customers, 1, '旧客户数据应保留');
    assertEq(upgraded.products, 12, '旧 payload 应补齐 12 条预置产品');
    assert(upgraded.names >= 6, '补齐后的产品名称应可用', upgraded.names);
    // 升级后的 payload 会在首次写操作时持久化 products（v1.0 → v1.1 迁移）
    await goto(page, '#/settings');
    await waitSelector(page, '#setting-warnDays');
    await clearToasts(page);
    await click(page, '[data-action="save-settings"]');
    await waitToast(page, 'ok', '设置已保存');
    const persisted = await page.evaluate(`(() => {
      let parsed = null;
      try { parsed = JSON.parse(localStorage.getItem('customer-crm.v1')); } catch (err) { parsed = null; }
      return {
        products: parsed && parsed.products ? parsed.products.length : null,
        customers: parsed && parsed.customers ? parsed.customers.length : null
      };
    })()`);
    assertEq(persisted.products, 12, '首次写操作后 localStorage 应持久化 products');
    assertEq(persisted.customers, 1, '首次写操作后客户数据保持');
    const errors = page.errors();
    assert(errors.length === 0, '旧数据升级过程不应产生错误', errors.slice(0, 4));
  });
}

/* ================================================================== */
/* S. 交付物校验                                                        */
/* ================================================================== */

async function groupS() {
  setGroup('S', '交付物校验');

  await test('S1', '8 张验收截图齐备且尺寸正确（1440×900 / 390×844）', async () => {
    const problems = [];
    for (const spec of SHOT_SPECS) {
      const filePath = path.join(SHOTS, spec.file);
      if (!fs.existsSync(filePath)) { problems.push(spec.file + ' 缺失'); continue; }
      const size = pngSize(filePath);
      if (!size) { problems.push(spec.file + ' 不是合法 PNG'); continue; }
      if (size.bytes < 20000) problems.push(spec.file + ' 体积过小 ' + size.bytes + 'B');
      if (size.width !== spec.width || size.height !== spec.height) {
        problems.push(spec.file + ' 尺寸 ' + size.width + 'x' + size.height + ' 期望 ' + spec.width + 'x' + spec.height);
      }
    }
    assertEq(problems.join('；'), '', '截图检查失败');
  });

  await test('S2', '截图文件命名与 README 引用路径一致（screenshots/*.png）', async () => {
    for (const spec of SHOT_SPECS) {
      assert(/^[a-z-]+\.png$/.test(spec.file), '截图文件名不符合规范', spec.file);
      assert(fs.existsSync(path.join(SHOTS, spec.file)), '截图缺失', spec.file);
    }
  });
}

/* ================================================================== */
/* 主流程                                                               */
/* ================================================================== */

python = findPython();
console.log('客户通 CustomerCRM — 端到端验收（E2E）');
console.log('项目根：' + ROOT);
console.log('被测版本：CONTRACT v1.1（13 字段 / 表格 15 列 / Excel 21+11 列 / 产品库 / 扇形图 / 宽版台账）');
console.log('Node：' + process.version + ' / ' + process.platform + ' / ' + os.release());
console.log('截图目录：' + SHOTS);
console.log('openpyxl：' + (python ? python.version + ' @ ' + python.command : '不可用 → 降级为 Node ZIP/XML 断言'));

let chrome = null;
let server = null;
let page = null;

try {
  fs.mkdirSync(SHOTS, { recursive: true });
  server = await serveStatic(ROOT);
  chrome = await launchChrome({ windowSize: { width: 1440, height: 900 } });
  console.log('Chrome：' + (chrome.devtools && chrome.devtools.Browser ? chrome.devtools.Browser : chrome.chromePath));
  console.log('静态服务：' + server.url);
  const connection = await connectToPage(chrome);
  page = await PageSession.create(connection.client);

  await page.setViewport(1440, 900);
  await page.setDownloadBehavior(DOWNLOADS);

  await page.navigate(server.url + '/index.html');
  await page.waitForApp({ timeout: 25000 });
  await waitExpr(page, `!!document.querySelector('#view-inner .page')`, '首屏渲染', 20000);

  // 需求：每次 e2e 运行开头清空并重建示例数据，保证可重复
  await page.evaluate('localStorage.clear()');
  page.clearErrors();
  await page.reload();
  await page.waitForApp({ timeout: 25000 });
  await waitExpr(page, `!!document.querySelector('#view-inner .page')`, '重载后首屏渲染', 20000);
  await waitStoreCount(page, 12, 15000);

  await groupA(page);
  await groupB(page);
  await groupC(page);
  await groupP(page);
  await groupD(page);
  await groupE(page);
  await groupF(page);
  await groupG(page);
  await groupH(page);
  await groupI(page);
  await groupS();
} catch (err) {
  let bootDiagnostics = null;
  try {
    if (page) {
      bootDiagnostics = {
        url: await page.currentUrl(),
        readyState: await page.evaluate('document.readyState'),
        bodyHead: String(await page.evaluate('document.body ? document.body.innerHTML.slice(0, 500) : "（无 body）"')),
        errors: page.errors().slice(0, 6)
      };
    }
  } catch { /* ignore */ }
  results.push({
    id: 'FATAL',
    group: group.id,
    title: 'E2E 运行中断',
    ok: false,
    detail: ((err && err.stack) || String(err)) + (bootDiagnostics ? '\n启动诊断：' + JSON.stringify(bootDiagnostics) : ''),
    ms: 0
  });
  console.error('\nE2E 运行中断：', err && err.stack ? err.stack : err);
  if (bootDiagnostics) console.error('启动诊断：', JSON.stringify(bootDiagnostics, null, 2));
} finally {
  try { if (page) page.close(); } catch { /* ignore */ }
  try { if (chrome) await chrome.close(); } catch { /* ignore */ }
  try { if (server) await server.close(); } catch { /* ignore */ }
}

/* ---------------- 汇总 ---------------- */

const failed = results.filter((result) => !result.ok);
const passed = results.length - failed.length;
const groupIds = Array.from(new Set(results.map((result) => result.group)));

console.log('\n================ E2E 汇总 ================');
for (const groupId of groupIds) {
  const groupResults = results.filter((result) => result.group === groupId);
  const groupFailed = groupResults.filter((result) => !result.ok).length;
  console.log(`  ${groupId}：${groupResults.length - groupFailed}/${groupResults.length} PASS` + (groupFailed ? `（${groupFailed} FAIL）` : ''));
}
console.log(`\n用例：${results.length} 个，PASS ${passed}，FAIL ${failed.length}`);
console.log('截图目录：' + SHOTS);
console.log('下载/临时目录：' + TEMP);
if (python) console.log('openpyxl 校验：' + python.version);
else console.log('openpyxl 校验：不可用（使用 Node ZIP/XML 断言）');

if (failed.length) {
  console.log('\n失败用例与详情：');
  for (const result of failed) {
    console.log(`  ✗ ${result.id}  [${result.group}] ${result.title}`);
    console.log('      ' + String(result.detail).split('\n').join('\n      '));
  }
  process.exitCode = 1;
} else {
  console.log('\n全部用例通过 ✓');
  try { fs.rmSync(TEMP, { recursive: true, force: true }); } catch { /* ignore */ }
  process.exitCode = 0;
}
