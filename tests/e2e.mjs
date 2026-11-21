/**
 * 瀹㈡埛閫?CustomerCRM 鈥?绔埌绔獙鏀讹紙E2E锛? *
 * 杩愯锛歯ode tests/e2e.mjs
 *   鍏ㄧ豢閫€鍑虹爜 0锛涗换涓€鐢ㄤ緥澶辫触鎵撳嵃澶辫触鐢ㄤ緥涓庤鎯呭苟閫€鍑虹爜 1銆? *
 * 鎶€鏈細闆?npm 渚濊禆锛屽鐢?tests/cdp.mjs锛堥殢鏈虹鍙ｉ潤鎬佹湇鍔?+ headless Chrome CDP锛夈€? * 鏁版嵁锛氭瘡娆¤繍琛?姣忕粍鐢ㄤ緥寮€澶?localStorage.clear() + reload锛岃嚜鍔ㄩ噸寤?12 鏉＄ず渚嬫暟鎹紝
 *       淇濊瘉鍙噸澶嶃€佺粍闂撮殧绂汇€? *
 * 瑕嗙洊锛圕ONTRACT 搂8/搂9 + task-4锛夛細
 *   A 棣栧睆 / 浠〃鐩?4 缁熻鍗?/ 鏃?console error / 鎴浘
 *   B 瀹㈡埛鍙拌处 12 琛?14 鍒?/ 鎼滅储 / 鍥藉路鐘舵€伮锋垚浜ょ瓫閫?/ 缁勫悎绛涢€?/ 閲嶇疆 / 涓夌鎺掑簭 / 灞曞紑琛? *   C 鏂板缓锛堝繀濉牎楠屻€佷繚瀛樸€佸埛鏂版寔涔呭寲锛? 缂栬緫 / 璺熻繘锛堢 4 鏉℃埅鏂€佹渶鏂板湪鍓嶃€佽窛涓婃璺熻繘锛? 鍒犻櫎
 *   D 閿洏蹇嵎閿?/ 鑱氱劍鎼滅储銆丯 鏂板缓銆丒sc 鍏抽棴寮圭獥
 *   E 鏅鸿兘璇嗗埆锛?1 椤广€佲墺6 鍛戒腑銆佸～鍏ヨ〃鍗曘€佺洿鎺ユ柊澧炰竴琛屻€佺┖鏂囨湰鍏滃簳锛? *   F 涓€閿鍑?Excel锛堢湡瀹炰笅杞姐€佹枃浠跺悕銆丳K 澶淬€乑IP 閮ㄤ欢銆乷penpyxl 涓や釜宸ヤ綔琛?/ 19 鍒?/ 琛屾暟 / 绛涢€夊鍑猴級
 *   G 璁剧疆锛堥槇鍊兼寔涔呭寲涓庝华琛ㄧ洏鑱斿姩銆丣SON 澶囦唤瀵煎嚭/鎭㈠銆佹竻绌哄悗绌虹姸鎬併€佽浇鍏ョず渚嬫暟鎹級
 *   H 鍝嶅簲寮?1440/1024/768/390 鏃犳í鍚戞孩鍑恒€佷晶鏍忔娊灞夈€佹墦鍗板獟浣撻殣钘忋€佺Щ鍔ㄧ鎴浘
 *   I 鏁版嵁鎹熷潖鍏滃簳锛堥潪娉?JSON 鍒锋柊涓嶇櫧灞忋€侀噸寤虹ず渚嬫暟鎹級
 *   S 8 寮犻獙鏀舵埅鍥鹃綈澶囥€丳NG 灏哄姝ｇ‘
 */
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import zlib from 'node:zlib';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { serveStatic, launchChrome, connectToPage, PageSession, sleep } from './cdp.mjs';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const SHOTS = path.join(ROOT, 'screenshots');
const TEMP = fs.mkdtempSync(path.join(os.tmpdir(), 'customer-crm-e2e-'));
const DOWNLOADS = path.join(TEMP, 'downloads');
fs.mkdirSync(DOWNLOADS, { recursive: true });

/* ================================================================== */
/* 甯搁噺                                                                */
/* ================================================================== */

/** CONTRACT 搂6.1锛氬鎴锋€昏〃 19 鍒楋紙椤哄簭鍥哄畾锛?*/
const EXPECTED_TOTAL_HEADERS = [
  '搴忓彿', '棣栨璇㈢洏鏃堕棿', '瀹㈡埛鍚嶇О', '鍜ㄨ浜у搧', '鍥藉', '鍏徃鍚?, '鎴愮珛鏃堕棿', '鍥㈤槦瑙勬ā', '骞磋惀鏀?,
  '璺熻繘鐘舵€?, '鏄惁鎴愪氦', '鏍稿績鍘熷洜', '璺濅笂娆¤窡杩?澶?',
  '鏈€杩戣窡杩涒憼鏃ユ湡', '鏈€杩戣窡杩涒憼鍐呭', '鏈€杩戣窡杩涒憽鏃ユ湡', '鏈€杩戣窡杩涒憽鍐呭', '鏈€杩戣窡杩涒憿鏃ユ湡', '鏈€杩戣窡杩涒憿鍐呭'
];

/** README 寮曠敤鐨?8 寮犻獙鏀舵埅鍥撅紙1440脳900 鍩哄噯 + 390脳844 绉诲姩绔級 */
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

/** openpyxl 鏍￠獙鑴氭湰锛堢敱 Python 3.12 + openpyxl 3.1.5 鎵ц锛?*/
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
    out["sheets"][name] = {
        "max_row": ws.max_row,
        "max_column": ws.max_column,
        "header": header,
        "names": names,
        "freeze": ws.freeze_panes,
        "auto_filter": ws.auto_filter.ref if ws.auto_filter is not None else None,
    }
sys.stdout.write("\\n" + json.dumps(out, ensure_ascii=False) + "\\n")
`;

/* ================================================================== */
/* 鏂█涓庣敤渚嬫敹闆?                                                      */
/* ================================================================== */

const results = [];
let group = { id: 'P', title: '鍑嗗' };
let python = null;
const scratch = {};

const J = (value) => JSON.stringify(value);

function formatValue(value) {
  if (typeof value === 'string') return value.length > 500 ? value.slice(0, 500) + '鈥? : value;
  try {
    const text = JSON.stringify(value);
    return text !== undefined && text.length > 500 ? text.slice(0, 500) + '鈥? : String(text);
  } catch {
    return String(value);
  }
}

function setGroup(id, title) {
  group = { id, title };
  console.log(`\n鈹€鈹€ ${id} ${title} 鈹€鈹€`);
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
    : `${message} 鈫?${formatValue(detail)}`;
  throw err;
}

function assertEq(actual, expected, message) {
  if (actual === expected) return;
  const err = new Error(message);
  err.__detail = `${message} 鈫?瀹為檯: ${formatValue(actual)}锛涙湡鏈? ${formatValue(expected)}`;
  throw err;
}

function fail(message, detail) {
  assert(false, message, detail);
}

/* ================================================================== */
/* 椤甸潰宸ュ叿                                                             */
/* ================================================================== */

async function waitExpr(page, expression, label, timeout = 10000) {
  try {
    return await page.waitUntil(expression, { timeout, label });
  } catch (err) {
    throw Object.assign(new Error(`绛夊緟瓒呮椂锛?{label}`), {
      __detail: `绛夊緟瓒呮椂锛?{timeout}ms锛夛細${label}\n琛ㄨ揪寮忥細${expression}`
    });
  }
}

async function waitSelector(page, selector, label, timeout = 10000) {
  return waitExpr(page, `!!document.querySelector(${J(selector)})`, label || `鍏冪礌鍑虹幇 ${selector}`, timeout);
}

async function waitGone(page, selector, label, timeout = 10000) {
  return waitExpr(page, `!document.querySelector(${J(selector)})`, label || `鍏冪礌娑堝け ${selector}`, timeout);
}

async function click(page, selector) {
  const ok = await page.click(selector);
  assert(ok, '鐐瑰嚮澶辫触锛氬厓绱犱笉瀛樺湪', selector);
  return ok;
}

async function goto(page, hash) {
  await page.appNavigate(hash);
  await waitExpr(page, `!!document.querySelector('#view-inner .page')`, `璺敱 ${hash} 娓叉煋瀹屾垚`, 10000);
}

async function rowCount(page) {
  return page.count('tr[data-testid="customer-row"]');
}

async function waitRowCount(page, expected, timeout = 10000) {
  return waitExpr(
    page,
    `document.querySelectorAll('tr[data-testid="customer-row"]').length === ${Number(expected)}`,
    `瀹㈡埛琛ㄦ牸琛屾暟 = ${expected}`,
    timeout
  );
}

async function waitStoreCount(page, expected, timeout = 10000) {
  return waitExpr(
    page,
    `window.CustomerCRM.store.list().length === ${Number(expected)}`,
    `store 瀹㈡埛璁板綍鏁?= ${expected}`,
    timeout
  );
}

async function waitHash(page, hash, timeout = 10000) {
  return waitExpr(page, `location.hash === ${J(hash)}`, `璺敱 hash = ${hash}`, timeout);
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
  fail(`鏈嚭鐜版湡鏈涚殑 ${kind} Toast锛堝惈銆?{contains}銆嶏級`, seen);
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

/** 姣忕粍鐢ㄤ緥寮€澶达細娓呯┖ localStorage + 鍒锋柊锛岀瓑寰?12 鏉＄ず渚嬫暟鎹噸寤?*/
async function resetDemo(page) {
  await page.evaluate('localStorage.clear()');
  await page.reload();
  await page.waitForApp({ timeout: 25000 });
  await waitExpr(page, `!!document.querySelector('#view-inner .page')`, '閲嶇疆鍚庨〉闈㈡覆鏌?, 20000);
  await waitStoreCount(page, 12, 15000);
  page.clearErrors();
}

async function assertNoOverflow(page, label) {
  const metrics = await page.layoutMetrics();
  assert(
    metrics.docScrollWidth <= metrics.innerWidth + 1,
    `${label} 鍑虹幇鏁撮〉妯悜婧㈠嚭锛坉ocumentElement.scrollWidth > innerWidth锛塦,
    metrics
  );
  assert(
    metrics.bodyScrollWidth <= metrics.bodyClientWidth + 1,
    `${label} body 鍑虹幇妯悜婧㈠嚭锛坆odyScrollWidth > bodyClientWidth锛塦,
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
    assert(errors.length === 0, '瀛樺湪 console error / pageerror / log.error', errors.slice(0, 6));
  });
}

/* ---------------- 涓嬭浇 / XLSX / 鎴浘 ---------------- */

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
  fail('绛夊緟涓嬭浇鏂囦欢瓒呮椂锛? + extensionRe, existing);
}

/** 瑙ｆ瀽 ZIP锛坸lsx 涓?stored 鏂瑰紡锛屽悓鏃跺吋瀹?deflate锛夛紝杩斿洖 Map<鏂囦欢鍚? Buffer> */
function readZip(buffer) {
  assert(buffer.length > 22, 'ZIP 鏂囦欢杩囧皬', buffer.length);
  let eocd = -1;
  for (let i = buffer.length - 22; i >= Math.max(0, buffer.length - 22 - 65535); i -= 1) {
    if (buffer.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  assert(eocd >= 0, 'ZIP 缂哄皯 EOCD 缁撴潫璁板綍锛堟枃浠跺彲鑳芥湭鍐欏畬/鎹熷潖锛?);
  const count = buffer.readUInt16LE(eocd + 10);
  const cdOffset = buffer.readUInt32LE(eocd + 16);
  const entries = new Map();
  let cursor = cdOffset;
  for (let i = 0; i < count; i += 1) {
    assert(buffer.readUInt32LE(cursor) === 0x02014b50, 'ZIP 涓ぎ鐩綍鎹熷潖', { index: i, offset: cursor });
    const method = buffer.readUInt16LE(cursor + 10);
    const compressedSize = buffer.readUInt32LE(cursor + 20);
    const nameLength = buffer.readUInt16LE(cursor + 28);
    const extraLength = buffer.readUInt16LE(cursor + 30);
    const commentLength = buffer.readUInt16LE(cursor + 32);
    const localOffset = buffer.readUInt32LE(cursor + 42);
    const name = buffer.subarray(cursor + 46, cursor + 46 + nameLength).toString('utf8');
    assert(buffer.readUInt32LE(localOffset) === 0x04034b50, 'ZIP 鏈湴鏂囦欢澶存崯鍧?, name);
    const localNameLength = buffer.readUInt16LE(localOffset + 26);
    const localExtraLength = buffer.readUInt16LE(localOffset + 28);
    const dataStart = localOffset + 30 + localNameLength + localExtraLength;
    const raw = buffer.subarray(dataStart, dataStart + compressedSize);
    let data;
    if (method === 0) data = Buffer.from(raw);
    else if (method === 8) data = zlib.inflateRawSync(raw);
    else fail('涓嶆敮鎸佺殑 ZIP 鍘嬬缉鏂瑰紡', { name, method });
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
  assert(fs.existsSync(filePath), '鎴浘鏂囦欢涓嶅瓨鍦?, filePath);
  const size = pngSize(filePath);
  assert(size, '鎴浘涓嶆槸鍚堟硶 PNG', filePath);
  assert(size.bytes >= minBytes, '鎴浘浣撶Н杩囧皬锛屽彲鑳芥槸绌虹櫧椤?, { file, bytes: size.bytes });
  assertEq(size.width, expectedWidth, `鎴浘 ${file} 瀹藉害`);
  assertEq(size.height, expectedHeight, `鎴浘 ${file} 楂樺害`);
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
    windowsHide: true
  });
  const stderrTail = String(result.stderr || '').trim().slice(-800);
  assert(result.status === 0, 'openpyxl 璇诲彇 xlsx 澶辫触锛堟枃浠跺彲鑳芥崯鍧忥級', stderrTail || String(result.stdout || '').slice(-400));
  const lines = String(result.stdout || '').trim().split(/\r?\n/).filter(Boolean);
  const last = lines[lines.length - 1] || '';
  let parsed;
  try {
    parsed = JSON.parse(last);
  } catch {
    fail('openpyxl 鏍￠獙鑴氭湰杈撳嚭鏃犳硶瑙ｆ瀽涓?JSON', last.slice(0, 400));
  }
  return parsed;
}

/* ---------------- 閿洏 ---------------- */

const KEY_DEFS = {
  '/': { key: '/', code: 'Slash', vk: 191 },
  n: { key: 'n', code: 'KeyN', vk: 78 },
  Escape: { key: 'Escape', code: 'Escape', vk: 27 }
};

async function pressKey(page, keyName) {
  const def = KEY_DEFS[keyName];
  assert(def, '鏈畾涔夌殑鎸夐敭', keyName);
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
/* A. 棣栧睆涓庝华琛ㄧ洏                                                     */
/* ================================================================== */

async function groupA(page) {
  setGroup('A', '棣栧睆鍚姩涓庝华琛ㄧ洏');

  await test('A1', '棣栧睆鍚姩锛?/dashboard 娓叉煋锛屽鑸珮浜€佹爣棰樻纭?, async () => {
    await waitHash(page, '#/dashboard', 15000);
    assertEq(await page.evaluate('window.CustomerCRM.app.getRoute()'), 'dashboard', '榛樿璺敱搴斾负 dashboard');
    assert(await page.exists('#view-inner .page'), '椤甸潰涓讳綋 #view-inner .page 鏈覆鏌?);
    assert(await page.exists('.nav-item.is-active[data-nav="dashboard"]'), '渚ф爮銆屼华琛ㄧ洏銆嶆湭楂樹寒');
    const title = await page.title();
    assert(String(title).indexOf('浠〃鐩?) !== -1, 'document.title 搴斿寘鍚€屼华琛ㄧ洏銆?, title);
    assert(await page.exists('[data-testid="stat-total"]'), '缂哄皯缁熻鍗?stat-total');
  });

  await test('A2', '棣栧睆鏃?console error / pageerror / log.error', async () => {
    const errors = page.errors();
    assert(errors.length === 0, '棣栧睆瀛樺湪閿欒绾ф帶鍒跺彴杈撳嚭', errors.slice(0, 6));
  });

  await test('A3', '棣栨璁块棶鑷姩鍐欏叆 12 鏉＄ず渚嬫暟鎹笖 11 涓瓧娈靛畬鏁?, async () => {
    const records = await page.evaluate('window.CustomerCRM.store.list()');
    assertEq(records.length, 12, '绀轰緥鏁版嵁鏉℃暟');
    const fieldKeys = await page.evaluate('window.CRM_SCHEMA.FIELD_KEYS');
    assertEq(fieldKeys.length, 11, 'CRM_SCHEMA.FIELD_KEYS 瀛楁鏁?);
    for (const record of records) {
      for (const key of fieldKeys) {
        assert(typeof record[key] === 'string', `瀹㈡埛銆?{record.customerName}銆嶅瓧娈?${key} 涓嶆槸瀛楃涓瞏, record[key]);
      }
      assert(record.customerName.length > 0, '绀轰緥瀹㈡埛缂哄皯瀹㈡埛鍚嶇О', record.id);
    }
    const names = records.map((record) => record.customerName);
    assert(names.indexOf('Michael Braun') !== -1, '绀轰緥鏁版嵁搴斿寘鍚绾︽牱渚嬪鎴?Michael Braun', names);
  });

  await test('A4', '浠〃鐩?4 涓粺璁″崱鏁板€?= CustomerCRM.app.stats()', async () => {
    const stats = await page.evaluate('window.CustomerCRM.app.stats()');
    const dom = await readStatsDom(page);
    assertEq(dom.total, String(stats.total), '瀹㈡埛鎬绘暟鍗?);
    assertEq(dom.deal, String(stats.deal), '宸叉垚浜ゅ崱');
    assertEq(dom.open, String(stats.open), '寰呰窡杩涘崱');
    assertEq(dom.overdue, String(stats.overdue), '瓒呮湡鏈窡杩涘崱');
    assertEq(Number(dom.total), 12, '瀹㈡埛鎬绘暟搴斾负 12');
    assert(stats.deal >= 1, '绀轰緥鏁版嵁搴旇嚦灏?1 浣嶅凡鎴愪氦瀹㈡埛', stats.deal);
    assert(dom.overdue !== null && dom.open !== null, '缁熻鍗℃湭娓叉煋', dom);
  });

  await test('A5', '浠〃鐩樺尯鍧楁覆鏌擄細4 缁熻鍗?/ 鍥藉鍒嗗竷 / 鐘舵€佸垎甯?/ 鏈€杩戝姩鎬?/ 浼樺厛璺熻繘', async () => {
    const layout = await page.evaluate(`(() => ({
      statCards: document.querySelectorAll('.stats-grid .stat-card').length,
      barRows: document.querySelectorAll('.bar-row').length,
      timelineItems: document.querySelectorAll('.timeline-item').length,
      priorityItems: document.querySelectorAll('.priority-item').length,
      pageSub: (document.querySelector('.page-sub') || {}).textContent || ''
    }))()`);
    assertEq(layout.statCards, 4, '缁熻鍗℃暟閲?);
    assert(layout.barRows >= 4, '鍥藉/鐘舵€佸垎甯冩潯褰㈠浘鏉＄洰杩囧皯', layout.barRows);
    assertEq(layout.timelineItems, 8, '鏈€杩戣窡杩涘姩鎬佸簲灞曠ず 8 鏉?);
    assert(layout.pageSub.indexOf('鍏?12 浣嶅鎴?) !== -1, '椤甸潰鎽樿搴旀樉绀?12 浣嶅鎴?, layout.pageSub);
    const stats = await page.evaluate('window.CustomerCRM.app.stats()');
    assertEq(layout.priorityItems, stats.overdueList.length, '浼樺厛璺熻繘鏉℃暟 = stats.overdueList');
  });

  await test('A6', '鍐欑姸鎬佹寚绀哄櫒鍒濆涓?data-state="saved"', async () => {
    const state = await page.evaluate(`(document.getElementById('save-state') || { getAttribute: function () { return null; } }).getAttribute('data-state')`);
    assertEq(state, 'saved', '#save-state data-state');
    const text = await page.text('#save-state');
    assert(String(text).length > 0, '#save-state 鏂囨湰涓虹┖');
  });

  await test('A7', '鍏叡娴嬭瘯鎺ュ彛鑷锛圕ONTRACT 搂8锛歛pp / store / schema / parser / xlsx锛?, async () => {
    const api = await page.evaluate(`(() => {
      const crm = window.CustomerCRM;
      if (!crm) return { missing: 'window.CustomerCRM' };
      const methods = ['navigate', 'getRoute', 'refresh', 'newCustomer', 'editCustomer', 'saveCustomerForm', 'deleteCustomer', 'addFollowup', 'recognize', 'applyRecognition', 'exportExcel', 'stats', 'filtered', 'toast', 'openFollowupModal'];
      const missingMethods = methods.filter(function (name) { return typeof crm.app[name] !== 'function'; });
      const storeMethods = ['list', 'get', 'add', 'update', 'remove', 'replaceAll', 'clear', 'seedDemo', 'settings', 'updateSettings', 'stats', 'exportJSON', 'importJSON', 'savedAt'];
      const missingStore = storeMethods.filter(function (name) { return typeof crm.store[name] !== 'function'; });
      return {
        version: crm.version,
        key: crm.store.KEY,
        missingMethods: missingMethods,
        missingStore: missingStore,
        parser: !!(crm.parser && typeof crm.parser.recognize === 'function'),
        sampleCount: crm.parser && Array.isArray(crm.parser.SAMPLE_TEXTS) ? crm.parser.SAMPLE_TEXTS.length : 0,
        xlsx: !!(crm.xlsx && typeof crm.xlsx.exportCustomers === 'function' && typeof crm.xlsx.download === 'function'),
        hasState: !!crm.app.state
      };
    })()`);
    assertEq(api.version, '1.0.0', 'CustomerCRM.version');
    assertEq(api.key, 'customer-crm.v1', 'store.KEY');
    assertEq(api.missingMethods.join(','), '', 'app 缂哄皯濂戠害鏂规硶');
    assertEq(api.missingStore.join(','), '', 'store 缂哄皯濂戠害鏂规硶');
    assert(api.parser, 'CrmParser.recognize 鏈姞杞?);
    assert(api.sampleCount >= 3, 'CrmParser.SAMPLE_TEXTS 搴旇嚦灏?3 娈?, api.sampleCount);
    assert(api.xlsx, 'CrmXlsx.exportCustomers / download 鏈姞杞?);
    assert(api.hasState, 'app.state 涓嶅瓨鍦?);
  });

  await test('A8', '1440脳900 棣栧睆鏃犳暣椤垫í鍚戞孩鍑?, async () => {
    await assertNoOverflow(page, '1440 浠〃鐩?);
  });

  await test('A9', '鎴浘 screenshots/dashboard.png锛?440脳900锛?, async () => {
    await takeScreenshot(page, 'dashboard.png');
    assetScreenshot('dashboard.png', 1440, 900);
  });
}

/* ================================================================== */
/* B. 瀹㈡埛鍙拌处                                                          */
/* ================================================================== */

async function groupB(page) {
  setGroup('B', '瀹㈡埛鍙拌处锛氳〃鏍?/ 鎼滅储 / 绛涢€?/ 鎺掑簭 / 灞曞紑');

  await test('B1', '瀹㈡埛鍙拌处 12 琛?脳 14 鍒楋紝宸ュ叿鏍忎笌 data-testid 榻愬叏', async () => {
    await resetDemo(page);
    await goto(page, '#/customers');
    await waitRowCount(page, 12);
    const layout = await page.evaluate(`(() => {
      const row = document.querySelector('tr[data-testid="customer-row"]');
      const headers = Array.from(document.querySelectorAll('.data-table--customers thead th')).map(function (th) { return th.textContent.trim(); });
      return { cells: row ? row.querySelectorAll('td').length : 0, headers: headers };
    })()`);
    assertEq(layout.cells, 14, '鏁版嵁琛?td 鏁伴噺锛堝睍寮€鍒?+ 11 瀛楁 + 璺濅笂娆¤窡杩?+ 鎿嶄綔鍒楋級');
    assertEq(layout.headers.length, 14, '琛ㄥご鍒楁暟');
    const headerText = layout.headers.join('|');
    for (const label of ['棣栨璇㈢洏鏃堕棿', '瀹㈡埛鍚嶇О', '鍜ㄨ浜у搧', '鍥藉', '鍏徃鍚?, '鎴愮珛鏃堕棿', '鍥㈤槦瑙勬ā', '骞磋惀鏀?, '璺熻繘鐘舵€?, '鏄惁鎴愪氦', '鏍稿績鍘熷洜', '璺濅笂娆¤窡杩?]) {
      assert(headerText.indexOf(label) !== -1, '琛ㄥご缂哄皯鍒楋細' + label, layout.headers);
    }
    for (const testid of ['btn-export-excel', 'customer-search', 'filter-country', 'filter-status', 'filter-deal', 'filter-reset', 'btn-edit-customer', 'btn-delete-customer']) {
      assert(await page.exists(`[data-testid="${testid}"]`), '缂哄皯 data-testid锛? + testid);
    }
    assert(await page.exists('[data-testid="new-customer"]'), '缂哄皯 data-testid锛歯ew-customer');
  });

  await test('B2', '鍏抽敭璇嶆悳绱€屽痉鍥姐€嶁啋 鍛戒腑 1 琛岋紝DOM 椤哄簭 = app.filtered()', async () => {
    await page.fill('[data-testid="customer-search"]', '寰峰浗');
    await waitRowCount(page, 1);
    const state = await page.evaluate(`(() => ({
      domIds: Array.from(document.querySelectorAll('tr[data-testid="customer-row"]')).map(function (tr) { return tr.getAttribute('data-id'); }),
      filteredIds: window.CustomerCRM.app.filtered().map(function (record) { return record.id; }),
      filteredOn: !!document.querySelector('.filter-on')
    }))()`);
    assertEq(state.domIds.join(','), state.filteredIds.join(','), '鎼滅储缁撴灉 DOM 椤哄簭 = app.filtered()');
    assert(state.filteredOn, '绛涢€夌敓鏁堟椂搴旀樉绀恒€屽凡绛涢€夈€嶆爣璁?);
    const country = await page.text('tr[data-testid="customer-row"] td[data-col="country"]');
    assert(country.indexOf('寰峰浗') !== -1, '鍛戒腑琛屽浗瀹跺垪搴斿惈銆屽痉鍥姐€?, country);
  });

  await test('B3', '鎼滅储鏃犵粨鏋?鈫?绌虹姸鎬佸紩瀵硷紝閲嶇疆鍚庢仮澶?12 琛?, async () => {
    await page.fill('[data-testid="customer-search"]', 'zzz-涓嶅瓨鍦ㄧ殑瀹㈡埛-zzz');
    await waitRowCount(page, 0);
    assert(await page.exists('.empty'), '鏃犵粨鏋滄椂搴旀樉绀虹┖鐘舵€?.empty');
    const emptyTitle = String(await page.text('.empty-title')).trim();
    assertEq(emptyTitle, '娌℃湁鍖归厤鐨勫鎴?, '鏃犵粨鏋滅┖鐘舵€佹爣棰?);
    await click(page, '[data-testid="filter-reset"]');
    await waitRowCount(page, 12);
    assertEq(await page.evaluate('window.CustomerCRM.app.state.filters.q'), '', '閲嶇疆鍚庡叧閿瘝搴旀竻绌?);
  });

  await test('B4', '鍥藉绛涢€夈€屽痉鍥姐€嶁啋 1 琛岋紙涓?store 璁＄畻涓€鑷达級', async () => {
    const expected = await page.evaluate(`window.CustomerCRM.store.list().filter(function (record) { return record.country === '寰峰浗'; }).length`);
    await page.fill('[data-testid="filter-country"]', '寰峰浗');
    await waitRowCount(page, expected);
    assertEq(expected, 1, 'store 涓痉鍥藉鎴锋暟');
    assertEq(await page.evaluate(`(document.querySelector('[data-testid="filter-country"]') || {}).value`), '寰峰浗', '鍥藉涓嬫媺閫変腑鍊?);
    await click(page, '[data-testid="filter-reset"]');
    await waitRowCount(page, 12);
  });

  await test('B5', '璺熻繘鐘舵€佺瓫閫夈€屽凡鎶ヤ环銆嶁啋 3 琛?, async () => {
    const expected = await page.evaluate(`window.CustomerCRM.store.list().filter(function (record) { return record.followupStatus === '宸叉姤浠?; }).length`);
    await page.fill('[data-testid="filter-status"]', '宸叉姤浠?);
    await waitRowCount(page, expected);
    assertEq(expected, 3, 'store 涓凡鎶ヤ环瀹㈡埛鏁?);
    const statuses = await readCol(page, 'followupStatus');
    assertEq(statuses.join(','), new Array(3).fill('宸叉姤浠?).join(','), '绛涢€夌粨鏋滆窡杩涚姸鎬佸潎涓哄凡鎶ヤ环');
    await click(page, '[data-testid="filter-reset"]');
    await waitRowCount(page, 12);
  });

  await test('B6', '鏄惁鎴愪氦绛涢€夈€屽凡鎴愪氦銆嶁啋 2 琛?, async () => {
    const expected = await page.evaluate(`window.CustomerCRM.store.list().filter(function (record) { return record.dealStatus === '宸叉垚浜?; }).length`);
    await page.fill('[data-testid="filter-deal"]', '宸叉垚浜?);
    await waitRowCount(page, expected);
    assertEq(expected, 2, 'store 涓凡鎴愪氦瀹㈡埛鏁?);
    const deals = await readCol(page, 'dealStatus');
    assertEq(deals.join(','), new Array(2).fill('宸叉垚浜?).join(','), '绛涢€夌粨鏋滄垚浜ょ姸鎬佸潎涓哄凡鎴愪氦');
    await click(page, '[data-testid="filter-reset"]');
    await waitRowCount(page, 12);
  });

  await test('B7', '缁勫悎绛涢€夛細璺熻繘鐘舵€?宸叉姤浠?+ 鍥藉=寰峰浗 鈫?1 琛?, async () => {
    await page.fill('[data-testid="filter-status"]', '宸叉姤浠?);
    await waitRowCount(page, 3);
    await page.fill('[data-testid="filter-country"]', '寰峰浗');
    await waitRowCount(page, 1);
    const names = await readCol(page, 'customerName');
    assertEq(names.join(','), 'Michael Braun', '缁勫悎绛涢€夊懡涓鎴?);
    assertEq(await page.evaluate('window.CustomerCRM.app.filtered().length'), 1, 'app.filtered() 鏉℃暟');
    await click(page, '[data-testid="filter-reset"]');
    await waitRowCount(page, 12);
  });

  await test('B8', '閲嶇疆绛涢€夋仮澶?12 琛屼笖绛涢€夋帶浠跺綊浣?, async () => {
    assertEq(await page.evaluate('window.CustomerCRM.app.state.filters.q'), '', '鍏抽敭璇?);
    assertEq(await page.evaluate('window.CustomerCRM.app.state.filters.country'), '', '鍥藉绛涢€?);
    assertEq(await page.evaluate('window.CustomerCRM.app.state.filters.status'), '', '鐘舵€佺瓫閫?);
    assertEq(await page.evaluate('window.CustomerCRM.app.state.filters.deal'), '', '鎴愪氦绛涢€?);
    assertEq(await page.evaluate(`(document.querySelector('[data-testid="filter-country"]') || {}).value`), '', '鍥藉涓嬫媺鍊?);
    assertEq(await rowCount(page), 12, '閲嶇疆鍚庤鏁?);
  });

  await test('B9', '鎺掑簭锛氶娆¤鐩樻椂闂?desc 鈫?asc 鈫?鎭㈠榛樿锛圖OM 椤哄簭 = app.filtered()锛?, async () => {
    await click(page, '[data-sort="firstInquiryAt"]');
    await sleep(60);
    let ids = await readRowIds(page);
    let expected = await filteredIds(page);
    assertEq(ids.join(','), expected.join(','), '棣栨鐐瑰嚮鍚?DOM 椤哄簭 = app.filtered()');
    let dates = await readCol(page, 'firstInquiryAt');
    for (let i = 1; i < dates.length; i += 1) assert(dates[i - 1] >= dates[i], '棣栨鐐瑰嚮搴旀寜棣栨璇㈢洏鏃堕棿闄嶅簭', dates);
    await click(page, '[data-sort="firstInquiryAt"]');
    await sleep(60);
    ids = await readRowIds(page);
    expected = await filteredIds(page);
    assertEq(ids.join(','), expected.join(','), '绗簩娆＄偣鍑诲悗 DOM 椤哄簭 = app.filtered()');
    dates = await readCol(page, 'firstInquiryAt');
    for (let i = 1; i < dates.length; i += 1) assert(dates[i - 1] <= dates[i], '绗簩娆＄偣鍑诲簲鎸夐娆¤鐩樻椂闂村崌搴?, dates);
    await click(page, '[data-sort="firstInquiryAt"]');
    await sleep(60);
    assertEq(await page.evaluate('window.CustomerCRM.app.state.filters.sort'), 'updated-desc', '绗笁娆＄偣鍑诲簲鎭㈠榛樿鎺掑簭 updated-desc');
    await click(page, '[data-testid="filter-reset"]');
    await waitRowCount(page, 12);
  });

  await test('B10', '鎺掑簭锛氬鎴峰悕绉?desc 鈫?asc锛圖OM 椤哄簭 = app.filtered()锛?, async () => {
    await click(page, '[data-sort="customerName"]');
    await sleep(60);
    let expected = await page.evaluate('window.CustomerCRM.app.filtered().map(function (record) { return record.customerName; })');
    let names = await readCol(page, 'customerName');
    assertEq(names.join('|'), expected.join('|'), '瀹㈡埛鍚嶇О闄嶅簭鎺掑簭缁撴灉');
    await click(page, '[data-sort="customerName"]');
    await sleep(60);
    expected = await page.evaluate('window.CustomerCRM.app.filtered().map(function (record) { return record.customerName; })');
    names = await readCol(page, 'customerName');
    assertEq(names.join('|'), expected.join('|'), '瀹㈡埛鍚嶇О鍗囧簭鎺掑簭缁撴灉');
    assertEq(await page.evaluate('window.CustomerCRM.app.state.filters.sort'), 'customerName-asc', '鎺掑簭鐘舵€?customerName-asc');
    await click(page, '[data-testid="filter-reset"]');
    await waitRowCount(page, 12);
  });

  await test('B11', '鎺掑簭锛氳窛涓婃璺熻繘锛坉aysSince desc锛?, async () => {
    await click(page, '[data-sort="daysSince"]');
    await sleep(60);
    assertEq(await page.evaluate('window.CustomerCRM.app.state.filters.sort'), 'daysSince-desc', '鎺掑簭鐘舵€?daysSince-desc');
    const ids = await readRowIds(page);
    const expected = await filteredIds(page);
    assertEq(ids.join(','), expected.join(','), '璺濅笂娆¤窡杩涙帓搴?DOM 椤哄簭 = app.filtered()');
    const derived = await page.evaluate('window.CustomerCRM.store.deriveAll()');
    const sorted = derived.slice().sort((a, b) => {
      if (a.daysSince === null && b.daysSince === null) return 0;
      if (a.daysSince === null) return 1;
      if (b.daysSince === null) return -1;
      return b.daysSince - a.daysSince;
    });
    assertEq(ids.join(','), sorted.map((record) => record.id).join(','), 'DOM 椤哄簭 = daysSince 闄嶅簭');
    await click(page, '[data-testid="filter-reset"]');
    await waitRowCount(page, 12);
  });

  await test('B12', '灞曞紑琛岋細鏄剧ず鏈€杩戜笁娆¤窡杩涳紙鏈€鏂板湪鍓嶏紝涓?store 涓€鑷达級', async () => {
    const firstId = await page.evaluate(`(document.querySelector('tr[data-testid="customer-row"]') || {}).getAttribute ? document.querySelector('tr[data-testid="customer-row"]').getAttribute('data-id') : null`);
    assert(firstId, '鎵句笉鍒伴琛?data-id');
    const record = await page.evaluate(`window.CustomerCRM.store.get(${J(firstId)})`);
    assert(record, 'store 涓壘涓嶅埌棣栬瀹㈡埛', firstId);
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
    assert(detail, '灞曞紑鍚庢壘涓嶅埌琛岃鎯?tr.row-detail');
    assertEq(detail.hidden, false, '琛岃鎯?hidden=false');
    assertEq(detail.expanded, true, '鏁版嵁琛屽簲鏈?row--expanded');
    assertEq(detail.aria, 'true', '灞曞紑鎸夐挳 aria-expanded');
    assertEq(detail.chips.length, Math.min(3, record.followups.length), '灞曞紑鏄剧ず璺熻繘鏉℃暟锛堚墹3锛?);
    const expectedNotes = record.followups.slice(0, detail.chips.length).map((item) => item.note);
    assertEq(detail.chips.map((chip) => chip.note).join('|'), expectedNotes.join('|'), '璺熻繘鍐呭涓?store 涓€鑷达紙鏈€鏂板湪鍓嶏級');
    assert(detail.chips[0].date.indexOf('鏈€杩?) !== -1, '绗竴鏉¤窡杩涘簲鏍囨敞銆屾渶杩戙€?, detail.chips[0].date);
    scratch.expandedId = firstId;
  });

  await test('B13', '鍐嶆鐐瑰嚮鏀惰捣琛岃鎯?, async () => {
    const id = scratch.expandedId;
    assert(id, '鍓嶇疆灞曞紑琛岀敤渚嬫湭鎴愬姛');
    await click(page, `tr[data-id="${id}"] [data-testid="row-toggle"]`);
    const state = await page.evaluate(`(() => {
      const row = document.querySelector('tr[data-id="${id}"]');
      const detail = document.querySelector('tr.row-detail[data-detail="${id}"]');
      return { hidden: detail ? detail.hidden : null, expanded: row ? row.classList.contains('row--expanded') : null };
    })()`);
    assertEq(state.hidden, true, '鏀惰捣鍚庤璇︽儏 hidden=true');
    assertEq(state.expanded, false, '鏀惰捣鍚庢棤 row--expanded');
  });

  await test('B14', '鎴浘 screenshots/customers.png 涓?customers-expanded.png锛?440脳900锛?, async () => {
    await waitRowCount(page, 12);
    await takeScreenshot(page, 'customers.png');
    assetScreenshot('customers.png', 1440, 900);
    const firstId = await page.evaluate(`document.querySelector('tr[data-testid="customer-row"]').getAttribute('data-id')`);
    await click(page, `tr[data-id="${firstId}"] [data-testid="row-toggle"]`);
    await waitExpr(page, `!document.querySelector('tr.row-detail[data-detail="${firstId}"]').hidden`, '琛岃鎯呭睍寮€');
    await takeScreenshot(page, 'customers-expanded.png');
    assetScreenshot('customers-expanded.png', 1440, 900);
    await click(page, `tr[data-id="${firstId}"] [data-testid="row-toggle"]`);
  });

  await assertNoErrors(page, 'B15', '瀹㈡埛鍙拌处鍏ㄦ祦绋嬫棤 console / page 閿欒');
}

/* ================================================================== */
/* C. 鏂板缓 / 缂栬緫 / 璺熻繘 / 鍒犻櫎                                         */
/* ================================================================== */

async function groupC(page) {
  setGroup('C', '瀹㈡埛澧炲垹鏀逛笌璺熻繘璁板綍');

  const newValues = {
    firstInquiryAt: '',
    customerName: 'E2E 娴嬭瘯瀹㈡埛',
    product: '娴嬭瘯淇濇俯鏉?600ml',
    country: '寰峰浗',
    companyName: 'E2E Test GmbH',
    foundedAt: '2015',
    teamSize: '20-50 浜?,
    annualRevenue: '500 涓囨鍏?,
    followupStatus: '璋堝垽涓?,
    dealStatus: '鏈垚浜?,
    coreReason: 'E2E 娴嬭瘯锛氫环鏍间笌浜ゆ湡绗﹀悎锛岀瓑寰呭鎴风‘璁よ鍗?
  };

  await test('C1', '鎵撳紑鏂板缓瀹㈡埛寮圭獥锛?1 涓瓧娈?+ 鎸夐挳 data-testid + role=dialog', async () => {
    await resetDemo(page);
    await goto(page, '#/customers');
    await waitRowCount(page, 12);
    await click(page, '[data-testid="new-customer"]');
    await waitSelector(page, '[data-testid="form-customer"]');
    const form = await page.evaluate(`(() => {
      const form = document.getElementById('customer-form');
      const names = Array.from(form.querySelectorAll('[name]')).map(function (el) { return el.getAttribute('name'); });
      const keys = window.CRM_SCHEMA.FIELD_KEYS.slice();
      return {
        testid: form.getAttribute('data-testid'),
        mode: form.getAttribute('data-mode'),
        names: names,
        allFields: keys.every(function (key) { return names.indexOf(key) !== -1; }),
        save: !!document.querySelector('[data-testid="btn-save-customer"]'),
        cancel: !!document.querySelector('[data-testid="btn-cancel-customer"]'),
        error: !!document.querySelector('[data-testid="form-error"]'),
        dialog: !!document.querySelector('#modal-root [role="dialog"][aria-modal="true"]'),
        required: Array.from(form.querySelectorAll('.field-label .req')).length
      };
    })()`);
    assertEq(form.testid, 'form-customer', '琛ㄥ崟 data-testid');
    assertEq(form.mode, 'create', '鏂板缓琛ㄥ崟 data-mode');
    assert(form.allFields, '琛ㄥ崟 11 涓瓧娈?name 鏈綈鍏?, form.names);
    assert(form.save && form.cancel && form.error, '琛ㄥ崟鎸夐挳/閿欒鎻愮ず data-testid 缂哄け', form);
    assert(form.dialog, '寮圭獥缂哄皯 role="dialog" aria-modal="true"');
    assertEq(form.required, 2, '蹇呭～鏍囪鏁伴噺锛堥娆¤鐩樻椂闂?/ 瀹㈡埛鍚嶇О锛?);
    const today = await page.evaluate('window.CRM_SCHEMA.todayISO()');
    assertEq(await page.evaluate(`document.querySelector('#customer-form [name="firstInquiryAt"]').value`), today, '鏂板缓琛ㄥ崟棣栨璇㈢洏鏃堕棿榛樿浠婂ぉ');
  });

  await test('C2', '鎴浘 screenshots/form.png锛?440脳900锛屾柊寤哄鎴峰脊绐楋級', async () => {
    await page.fill('customerName', 'E2E 楠屾敹瀹㈡埛');
    await page.fill('product', '涓嶉攬閽繚娓╂澂 600ml');
    await page.fill('country', '寰峰浗');
    await page.fill('companyName', 'QA Acceptance GmbH');
    await page.fill('coreReason', '浠锋牸涓庝氦鏈熷潎绗﹀悎锛岀瓑寰呭鎴风‘璁よ鍗?);
    await sleep(120);
    await takeScreenshot(page, 'form.png');
    assetScreenshot('form.png', 1440, 900);
  });

  await test('C3', '蹇呭～鏍￠獙锛氭竻绌哄鎴峰悕绉颁笌棣栨璇㈢洏鏃堕棿 鈫?form-error 鍑虹幇涓斿脊绐椾笉鍏抽棴', async () => {
    await page.fill('customerName', '');
    await page.fill('firstInquiryAt', '');
    await click(page, '[data-testid="btn-save-customer"]');
    await waitExpr(
      page,
      `(() => { const el = document.querySelector('[data-testid="form-error"]'); return el && !el.hidden && el.textContent.trim().length > 0; })()`,
      '蹇呭～鏍￠獙閿欒鎻愮ず鏄剧ず'
    );
    const errorText = String(await page.text('[data-testid="form-error"]'));
    assert(errorText.indexOf('瀹㈡埛鍚嶇О') !== -1, '閿欒姹囨€诲簲鍖呭惈銆屽鎴峰悕绉般€?, errorText);
    assert(errorText.indexOf('棣栨璇㈢洏鏃ユ湡') !== -1, '閿欒姹囨€诲簲鍖呭惈銆岄娆¤鐩樻棩鏈熴€?, errorText);
    assert(await page.exists('[data-testid="form-customer"]'), '鏍￠獙澶辫触鏃跺脊绐椾笉搴斿叧闂?);
    const fieldError = await page.evaluate(`(() => { const el = document.querySelector('[data-error-for="customerName"]'); return el ? (!el.hidden ? el.textContent : '') : ''; })()`);
    assert(String(fieldError).length > 0, '瀹㈡埛鍚嶇О瀛楁绾ч敊璇湭鏄剧ず');
  });

  await test('C4', '濉〃淇濆瓨 鈫?store 13 鏉°€佸脊绐楀叧闂€佽〃鏍煎嚭鐜版柊琛屻€乻ave-state 宸蹭繚瀛?, async () => {
    const today = await page.evaluate('window.CRM_SCHEMA.todayISO()');
    const values = Object.assign({}, newValues, { firstInquiryAt: today });
    for (const [key, value] of Object.entries(values)) {
      const ok = await page.fill(key, value);
      assert(ok, '琛ㄥ崟瀛楁濉厖澶辫触锛? + key);
    }
    await clearToasts(page);
    await click(page, '[data-testid="btn-save-customer"]');
    await waitGone(page, '[data-testid="form-customer"]', '淇濆瓨鍚庡脊绐楀叧闂?);
    await waitStoreCount(page, 13);
    await waitRowCount(page, 13);
    await waitToast(page, 'ok', '宸叉柊澧炲鎴?);
    const record = await page.evaluate(`window.CustomerCRM.store.list().filter(function (item) { return item.customerName === 'E2E 娴嬭瘯瀹㈡埛'; })[0] || null`);
    assert(record, 'store 涓壘涓嶅埌鏂板缓鐨?E2E 娴嬭瘯瀹㈡埛');
    scratch.customerId = record.id;
    assertEq(record.customerName, 'E2E 娴嬭瘯瀹㈡埛', 'store 瀹㈡埛鍚嶇О');
    assertEq(record.companyName, 'E2E Test GmbH', 'store 鍏徃鍚?);
    assertEq(record.country, '寰峰浗', 'store 鍥藉');
    assertEq(record.followupStatus, '璋堝垽涓?, 'store 璺熻繘鐘舵€?);
    assertEq(record.coreReason, 'E2E 娴嬭瘯锛氫环鏍间笌浜ゆ湡绗﹀悎锛岀瓑寰呭鎴风‘璁よ鍗?, 'store 鏍稿績鍘熷洜');
    const inTable = await page.evaluate(`Array.from(document.querySelectorAll('tr[data-testid="customer-row"] td[data-col="customerName"]')).some(function (td) { return td.textContent.indexOf('E2E 娴嬭瘯瀹㈡埛') !== -1; })`);
    assert(inTable, '琛ㄦ牸鏈嚭鐜版柊寤哄鎴?);
    const saveState = await page.evaluate(`document.querySelector('#save-state').getAttribute('data-state')`);
    assertEq(saveState, 'saved', '鍐欐搷浣滃悗 #save-state data-state');
  });

  await test('C5', '鍒锋柊椤甸潰鍚庢柊瀹㈡埛浠嶅湪锛坙ocalStorage 鎸佷箙鍖栵級', async () => {
    const id = scratch.customerId;
    assert(id, '鍓嶇疆淇濆瓨鐢ㄤ緥鏈垚鍔?);
    await page.reload();
    await page.waitForApp({ timeout: 20000 });
    await waitExpr(page, `!!document.querySelector('#view-inner .page')`, '鍒锋柊鍚庨〉闈㈡覆鏌?, 15000);
    await waitRowCount(page, 13);
    const persisted = await page.evaluate(`(() => {
      try {
        const payload = JSON.parse(localStorage.getItem('customer-crm.v1'));
        const record = payload.customers.filter(function (item) { return item.id === ${J(id)}; })[0];
        return { count: payload.customers.length, name: record ? record.customerName : null, error: '' };
      } catch (err) {
        return { count: null, name: null, error: String(err && err.message) };
      }
    })()`);
    assertEq(persisted.error, '', 'localStorage 鏁版嵁搴斿彲瑙ｆ瀽');
    assertEq(persisted.count, 13, 'localStorage 瀹㈡埛璁板綍鏁?);
    assertEq(persisted.name, 'E2E 娴嬭瘯瀹㈡埛', 'localStorage 涓鎴峰悕绉?);
    const visible = await page.evaluate(`Array.from(document.querySelectorAll('tr[data-testid="customer-row"] td[data-col="customerName"]')).some(function (td) { return td.textContent.indexOf('E2E 娴嬭瘯瀹㈡埛') !== -1; })`);
    assert(visible, '鍒锋柊鍚庤〃鏍兼湭鏄剧ず鏂板鎴?);
  });

  await test('C6', '缂栬緫瀹㈡埛锛氶濉纭?鈫?淇濆瓨鍚?store / 琛ㄦ牸鍚屾鏇存柊', async () => {
    const id = scratch.customerId;
    assert(id, '鍓嶇疆淇濆瓨鐢ㄤ緥鏈垚鍔?);
    await goto(page, '#/customers');
    await waitRowCount(page, 13);
    await click(page, `tr[data-id="${id}"] [data-testid="btn-edit-customer"]`);
    await waitSelector(page, '[data-testid="form-customer"]');
    assertEq(await page.evaluate(`document.getElementById('customer-form').getAttribute('data-mode')`), 'edit', '缂栬緫琛ㄥ崟 data-mode');
    const before = await formValues(page);
    assertEq(before.customerName, 'E2E 娴嬭瘯瀹㈡埛', '缂栬緫寮圭獥棰勫～瀹㈡埛鍚嶇О');
    assertEq(before.companyName, 'E2E Test GmbH', '缂栬緫寮圭獥棰勫～鍏徃鍚?);
    await page.fill('customerName', 'E2E 娴嬭瘯瀹㈡埛锛堝凡鏀瑰悕锛?);
    await clearToasts(page);
    await click(page, '[data-testid="btn-save-customer"]');
    await waitGone(page, '[data-testid="form-customer"]', '缂栬緫淇濆瓨鍚庡脊绐楀叧闂?);
    await waitExpr(page, `window.CustomerCRM.store.get(${J(id)}).customerName === 'E2E 娴嬭瘯瀹㈡埛锛堝凡鏀瑰悕锛?`, 'store 瀹㈡埛鍚嶅凡鏇存柊');
    assertEq(await page.evaluate(`window.CustomerCRM.store.get(${J(id)}).customerName`), 'E2E 娴嬭瘯瀹㈡埛锛堝凡鏀瑰悕锛?, 'store 瀹㈡埛鍚嶇О');
    const visible = await page.evaluate(`Array.from(document.querySelectorAll('tr[data-testid="customer-row"] td[data-col="customerName"]')).some(function (td) { return td.textContent.indexOf('E2E 娴嬭瘯瀹㈡埛锛堝凡鏀瑰悕锛?) !== -1; })`);
    assert(visible, '琛ㄦ牸鏈樉绀虹紪杈戝悗鐨勫鎴峰悕绉?);
    await waitToast(page, 'ok', '瀹㈡埛璧勬枡宸叉洿鏂?);
  });

  await test('C7', '娣诲姞璺熻繘绗?1~3 鏉★細鏉℃暟閫掑銆佹渶鏂板湪鍓?, async () => {
    const id = scratch.customerId;
    assert(id, '鍓嶇疆淇濆瓨鐢ㄤ緥鏈垚鍔?);
    const dates = await page.evaluate(`(() => { const schema = window.CRM_SCHEMA; const today = schema.todayISO(); return { today: today, m1: schema.addDays(today, -1), m2: schema.addDays(today, -2), m3: schema.addDays(today, -3) }; })()`);
    scratch.dates = dates;
    const plan = [
      { date: dates.m3, note: '绗?1 娆¤窡杩涳細鍙戦€佷骇鍝佺洰褰? },
      { date: dates.m2, note: '绗?2 娆¤窡杩涳細纭璇曞崟鏁伴噺' },
      { date: dates.m1, note: '绗?3 娆¤窡杩涳細瀵勫嚭鏍峰搧' }
    ];
    for (let i = 0; i < plan.length; i += 1) {
      await click(page, `tr[data-id="${id}"] [data-testid="btn-add-followup"]`);
      await waitSelector(page, '#followup-note', '璺熻繘寮圭獥鍑虹幇');
      await page.fill('#followup-date', plan[i].date);
      await page.fill('#followup-note', plan[i].note);
      await click(page, '[data-action="save-followup"]');
      await waitGone(page, '#followup-note', '璺熻繘淇濆瓨鍚庡脊绐楀叧闂?);
      await waitExpr(page, `window.CustomerCRM.store.get(${J(id)}).followups.length === ${i + 1}`, '璺熻繘鏉℃暟 = ' + (i + 1));
      const followups = await page.evaluate(`window.CustomerCRM.store.get(${J(id)}).followups`);
      assertEq(followups[0].date, plan[i].date, '鏈€鏂颁竴鏉¤窡杩涙棩鏈熷簲鍦ㄦ渶鍓?);
      assertEq(followups[0].note, plan[i].note, '鏈€鏂颁竴鏉¤窡杩涘唴瀹?);
    }
    const followups = await page.evaluate(`window.CustomerCRM.store.get(${J(id)}).followups`);
    assertEq(followups.length, 3, '涓夋璺熻繘璁板綍鏁?);
    assertEq(followups.map((item) => item.date).join(','), [dates.m1, dates.m2, dates.m3].join(','), '璺熻繘璁板綍搴旀寜鏃ユ湡鍊掑簭锛堟渶鏂板湪鍓嶏級');
  });

  await test('C8', '娣诲姞绗?4 鏉¤窡杩涳細鍙繚鐣?3 鏉°€佹渶鏂板湪鍓嶃€併€岃窛涓婃璺熻繘銆? 浠婂ぉ', async () => {
    const id = scratch.customerId;
    const dates = scratch.dates;
    assert(id && dates, '鍓嶇疆璺熻繘鐢ㄤ緥鏈垚鍔?);
    await click(page, `tr[data-id="${id}"] [data-testid="btn-add-followup"]`);
    await waitSelector(page, '#followup-note', '璺熻繘寮圭獥鍑虹幇');
    await page.fill('#followup-date', dates.today);
    await page.fill('#followup-note', '绗?4 娆¤窡杩涳紙鏈€鏂帮級');
    await click(page, '[data-action="save-followup"]');
    await waitGone(page, '#followup-note', '璺熻繘淇濆瓨鍚庡脊绐楀叧闂?);
    await waitExpr(
      page,
      `(() => { const record = window.CustomerCRM.store.get(${J(id)}); return record && record.followups.length === 3 && record.followups[0].date === ${J(dates.today)}; })()`,
      '绗?4 鏉¤窡杩涘啓鍏ヤ笖鎴柇鍒?3 鏉?
    );
    const followups = await page.evaluate(`window.CustomerCRM.store.get(${J(id)}).followups`);
    assertEq(followups.length, 3, '鍙繚鐣欐渶杩?3 鏉¤窡杩?);
    assertEq(followups[0].note, '绗?4 娆¤窡杩涳紙鏈€鏂帮級', '鏈€鏂颁竴鏉¤窡杩涘唴瀹?);
    assertEq(followups[0].date, dates.today, '鏈€鏂颁竴鏉¤窡杩涙棩鏈?);
    assertEq(followups[1].date, dates.m1, '绗?2 鏉″簲涓哄師鏈€鏂?);
    assertEq(followups[2].date, dates.m2, '绗?3 鏉″簲涓哄師绗?2 鏉?);
    assert(!followups.some((item) => item.note === '绗?1 娆¤窡杩涳細鍙戦€佷骇鍝佺洰褰?), '鏈€鏃т竴鏉¤窡杩涘簲琚埅鏂?);
    const age = await page.evaluate(`(() => {
      const cell = document.querySelector('tr[data-id="${id}"] td[data-col="age"]');
      if (!cell) return null;
      const el = cell.querySelector('.age');
      return { text: el ? el.textContent.trim() : '', cls: el ? el.className : '' };
    })()`);
    assert(age, '鎵句笉鍒拌瀹㈡埛鐨勮窛涓婃璺熻繘鍗曞厓鏍?);
    assertEq(age.text, '浠婂ぉ', '璺濅笂娆¤窡杩涘簲鏄剧ず銆屼粖澶┿€?);
    assert(age.cls.indexOf('age--ok') !== -1, '璺濅笂娆¤窡杩涜壊璋冨簲涓?ok锛堢豢锛?, age.cls);
  });

  await test('C9', '鍒犻櫎瀹㈡埛锛氬彇娑堜笉鍒狅紱纭鍚庡洖鍒?12 琛?, async () => {
    const id = scratch.customerId;
    assert(id, '鍓嶇疆淇濆瓨鐢ㄤ緥鏈垚鍔?);
    await click(page, `tr[data-id="${id}"] [data-testid="btn-delete-customer"]`);
    await waitSelector(page, '[data-confirm-ok]', '鍒犻櫎纭寮圭獥');
    const confirmText = String(await page.text('[data-modal-id] .confirm-text'));
    assert(confirmText.indexOf('E2E 娴嬭瘯瀹㈡埛锛堝凡鏀瑰悕锛?) !== -1, '鍒犻櫎纭鏂囨搴斿寘鍚鎴峰悕绉?, confirmText);
    await click(page, '[data-confirm-cancel]');
    await waitGone(page, '[data-confirm-ok]', '鍙栨秷鍚庣‘璁ゅ脊绐楀叧闂?);
    await sleep(120);
    await waitStoreCount(page, 13);
    assert(await page.exists(`tr[data-id="${id}"]`), '鍙栨秷鍒犻櫎鍚庡鎴疯搴斾繚鐣?);
    await click(page, `tr[data-id="${id}"] [data-testid="btn-delete-customer"]`);
    await waitSelector(page, '[data-confirm-ok]', '鍒犻櫎纭寮圭獥');
    await clearToasts(page);
    await click(page, '[data-confirm-ok]');
    await waitStoreCount(page, 12);
    await waitRowCount(page, 12);
    await waitToast(page, 'ok', '宸插垹闄ゅ鎴?);
    assert(!(await page.exists(`tr[data-id="${id}"]`)), '纭鍒犻櫎鍚庡鎴疯搴旀秷澶?);
    assertEq(await page.evaluate(`window.CustomerCRM.store.get(${J(id)})`), null, 'store 涓鎴峰簲琚垹闄?);
  });

  await test('C10', '琛屽唴鎿嶄綔鎸夐挳鍏峰濂戠害 data-testid=btn-add-followup', async () => {
    await waitRowCount(page, 12);
    assert(
      await page.exists('tr[data-testid="customer-row"] [data-testid="btn-add-followup"]'),
      'CONTRACT 搂8 瑕佹眰瀹㈡埛琛屾彁渚?data-testid="btn-add-followup"锛堝綋鍓嶄粎鏈?data-action="add-followup"锛?
    );
  });

  await assertNoErrors(page, 'C11', '瀹㈡埛澧炲垹鏀瑰叏娴佺▼鏃?console / page 閿欒');
}

/* ================================================================== */
/* D. 閿洏蹇嵎閿?                                                       */
/* ================================================================== */

async function groupD(page) {
  setGroup('D', '閿洏蹇嵎閿?);

  await test('D1', '銆?銆嶏細浠庝华琛ㄧ洏璺宠浆瀹㈡埛鍙拌处骞惰仛鐒︽悳绱㈡', async () => {
    await resetDemo(page);
    await goto(page, '#/dashboard');
    await waitSelector(page, '[data-testid="stat-total"]');
    await page.evaluate('document.activeElement && document.activeElement.blur && document.activeElement.blur()');
    await pressKey(page, '/');
    await waitHash(page, '#/customers');
    await waitExpr(page, `document.activeElement && document.activeElement.id === 'customer-search'`, '銆?銆嶈仛鐒︽悳绱㈡');
    await waitRowCount(page, 12);
  });

  await test('D2', '銆孨銆嶏細鎵撳紑鏂板缓瀹㈡埛寮圭獥', async () => {
    await page.evaluate('document.activeElement && document.activeElement.blur && document.activeElement.blur()');
    await pressKey(page, 'n');
    await waitSelector(page, '[data-testid="form-customer"]', '銆孨銆嶆墦寮€鏂板缓瀹㈡埛寮圭獥');
    assertEq(await page.evaluate(`document.getElementById('customer-form').getAttribute('data-mode')`), 'create', '鏂板缓寮圭獥 data-mode');
  });

  await test('D3', '銆孍sc銆嶏細鍏抽棴寮圭獥骞惰В闄?body 婊氬姩閿?, async () => {
    await pressKey(page, 'Escape');
    await waitGone(page, '[data-testid="form-customer"]', '銆孍sc銆嶅叧闂脊绐?);
    assertEq(await page.evaluate(`document.body.classList.contains('is-modal-open')`), false, 'Esc 鍚?body 搴旂Щ闄?is-modal-open');
  });

  await assertNoErrors(page, 'D4', '蹇嵎閿祦绋嬫棤 console / page 閿欒');
}

/* ================================================================== */
/* E. 鏅鸿兘璇嗗埆                                                          */
/* ================================================================== */

async function groupE(page) {
  setGroup('E', '鏅鸿兘璇嗗埆');

  await test('E1', '绮樿创绀轰緥璇㈢洏 鈫?璇嗗埆 11 椤广€佸懡涓?鈮?銆佹憳瑕佽鏁颁竴鑷?, async () => {
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
    assert(sample.chars > 50, '绀轰緥鏂囨湰鍐欏叆澶辫触', sample);
    await clearToasts(page);
    await click(page, '[data-testid="btn-run-recognize"]');
    await waitSelector(page, '.recog-list', '璇嗗埆缁撴灉鍒楄〃鍑虹幇');
    const dom = await page.evaluate(`(() => ({
      items: document.querySelectorAll('.recog-item').length,
      filled: document.querySelectorAll('.recog-item--filled').length,
      missing: document.querySelectorAll('.recog-item--missing').length,
      summary: (document.querySelector('.recog-summary .recog-count') || {}).textContent || '',
      fields: (function () { const out = {}; document.querySelectorAll('[data-recog-key]').forEach(function (el) { out[el.getAttribute('data-recog-key')] = el.value; }); return out; })(),
      evidenceCount: document.querySelectorAll('.recog-evidence').length,
      followupChips: document.querySelectorAll('.recog-result-body .followup-strip .followup-chip').length
    }))()`);
    assertEq(dom.items, 11, '璇嗗埆缁撴灉椤规暟锛?1 涓瓧娈碉級');
    assert(dom.filled >= 6, '璇嗗埆鍛戒腑瀛楁鏁板簲 鈮?', dom.filled);
    assertEq(dom.filled + dom.missing, 11, '鍛戒腑 + 鏈瘑鍒?= 11');
    const expected = await page.evaluate('window.CustomerCRM.app.recognize(window.CrmParser.SAMPLE_TEXTS[0].text)');
    assert(expected && expected.ok, 'CrmParser.recognize 搴旇繑鍥?ok');
    assertEq(dom.filled, expected.filled, '鐣岄潰鍛戒腑鏁?= parser.filled');
    const summaryMatch = dom.summary.match(/宸茶瘑鍒玕s*(\d+)/);
    assert(summaryMatch, '璇嗗埆鎽樿鏂囨鏍煎紡寮傚父', dom.summary);
    assertEq(Number(summaryMatch[1]), dom.filled, '鎽樿璁℃暟 = DOM 鍛戒腑椤规暟');
    assert(dom.evidenceCount >= 6, '搴斿睍绀哄師鏂囪瘉鎹紙鈮? 鏉★級', dom.evidenceCount);
    assert(dom.followupChips >= 1, '搴斿睍绀鸿瘑鍒埌鐨勮窡杩涜褰?, dom.followupChips);
    scratch.recognized = { dom, expected, sample };
  });

  await test('E2', '璇嗗埆瀛楁鍊间笌 CrmParser.recognize 鐩存帴璋冪敤涓€鑷?, async () => {
    const state = scratch.recognized;
    assert(state, '鍓嶇疆璇嗗埆鐢ㄤ緥鏈垚鍔?);
    const fields = state.expected.fields;
    for (const key of Object.keys(fields)) {
      assertEq(state.dom.fields[key], fields[key], '璇嗗埆瀛楁 ' + key + ' 涓?parser 涓€鑷?);
    }
    assertEq(state.dom.fields.customerName, 'Michael Braun', '璇嗗埆瀹㈡埛鍚嶇О');
    assertEq(state.dom.fields.country, '寰峰浗', '璇嗗埆鍥藉');
    assertEq(state.dom.fields.firstInquiryAt, '2026-09-12', '璇嗗埆棣栨璇㈢洏鏃堕棿');
    assert(state.dom.fields.product.indexOf('淇濇俯鏉?) !== -1, '璇嗗埆鍜ㄨ浜у搧搴斿惈淇濇俯鏉?, state.dom.fields.product);
  });

  await test('E3', '銆屽～鍏ヨ〃鏍煎苟鏍稿銆嶁啋 鏂板缓琛ㄥ崟鎸夎瘑鍒粨鏋滈濉?, async () => {
    const state = scratch.recognized;
    assert(state, '鍓嶇疆璇嗗埆鐢ㄤ緥鏈垚鍔?);
    await click(page, '[data-testid="btn-fill-customer"]');
    await waitSelector(page, '[data-testid="form-customer"]', '棰勫～琛ㄥ崟寮圭獥');
    const values = await formValues(page);
    assertEq(await page.evaluate(`document.getElementById('customer-form').getAttribute('data-mode')`), 'create', '棰勫～琛ㄥ崟涓烘柊寤烘ā寮?);
    assertEq(values.customerName, state.expected.fields.customerName, '棰勫～瀹㈡埛鍚嶇О');
    assertEq(values.firstInquiryAt, state.expected.fields.firstInquiryAt, '棰勫～棣栨璇㈢洏鏃堕棿');
    assertEq(values.country, state.expected.fields.country, '棰勫～鍥藉');
    assertEq(values.companyName, state.expected.fields.companyName, '棰勫～鍏徃鍚?);
    assert(await page.exists('.form-notice'), '棰勫～琛ㄥ崟搴旀樉绀恒€屽凡鏍规嵁璇嗗埆缁撴灉鑷姩濉叆銆嶆彁绀?);
    await click(page, '[data-testid="btn-cancel-customer"]');
    await waitGone(page, '[data-testid="form-customer"]', '鍙栨秷棰勫～琛ㄥ崟');
  });

  await test('E4', '鎴浘 screenshots/recognize.png锛?440脳900锛岃瘑鍒粨鏋滄€侊級', async () => {
    await waitSelector(page, '.recog-list', '璇嗗埆缁撴灉浠嶅湪');
    await takeScreenshot(page, 'recognize.png');
    assetScreenshot('recognize.png', 1440, 900);
  });

  await test('E5', '銆岀洿鎺ユ柊澧炰竴琛屻€嶁啋 store +1 涓斿瓧娈垫纭?, async () => {
    const state = scratch.recognized;
    assert(state, '鍓嶇疆璇嗗埆鐢ㄤ緥鏈垚鍔?);
    const beforeIds = await page.evaluate('window.CustomerCRM.store.list().map(function (record) { return record.id; })');
    await clearToasts(page);
    await click(page, '[data-testid="btn-add-recognized"]');
    await waitStoreCount(page, beforeIds.length + 1);
    await waitToast(page, 'ok', '宸叉柊澧炲鎴?);
    const added = await page.evaluate(`(() => {
      const before = ${J(beforeIds)};
      return window.CustomerCRM.store.list().filter(function (record) { return before.indexOf(record.id) === -1; })[0] || null;
    })()`);
    assert(added, '鐩存帴鏂板涓€琛屽悗 store 鏈鍔犺褰?);
    assertEq(added.customerName, state.expected.fields.customerName, '鏂板璁板綍瀹㈡埛鍚嶇О');
    assertEq(added.firstInquiryAt, state.expected.fields.firstInquiryAt, '鏂板璁板綍棣栨璇㈢洏鏃堕棿');
    assertEq(added.country, state.expected.fields.country, '鏂板璁板綍鍥藉');
    assert(added.followups.length >= 1, '鏂板璁板綍搴斿甫璇嗗埆鍒扮殑璺熻繘璁板綍', added.followups);
    scratch.recognizeAddedId = added.id;
  });

  await test('E6', '绌烘枃鏈偣鍑诲紑濮嬭瘑鍒?鈫?浠呮彁绀洪敊璇紝涓嶄骇鐢熺粨鏋?, async () => {
    await page.fill('[data-testid="recognize-input"]', '');
    await clearToasts(page);
    await click(page, '[data-testid="btn-run-recognize"]');
    await waitToast(page, 'error', '璇峰厛绮樿创涓€娈佃鐩樻枃鏈?);
    assert(await page.exists('[data-testid="recognize-input"]'), '璇嗗埆杈撳叆妗嗗簲淇濈暀');
  });

  await assertNoErrors(page, 'E7', '鏅鸿兘璇嗗埆娴佺▼鏃?console / page 閿欒');
}

/* ================================================================== */
/* F. 涓€閿鍑?Excel                                                    */
/* ================================================================== */

async function groupF(page) {
  setGroup('F', '涓€閿鍑?Excel');

  await test('F1', '鏃犵瓫閫夊鍑猴細瑙﹀彂鐪熷疄涓嬭浇锛屾枃浠跺悕 瀹㈡埛璺熻繘鍙拌处_YYYYMMDD_HHmm.xlsx', async () => {
    await resetDemo(page);
    await goto(page, '#/customers');
    await waitRowCount(page, 12);
    const expectedNames = await filteredNames(page);
    const expectedDetailRows = await page.evaluate('window.CustomerCRM.app.filtered().reduce(function (n, record) { return n + Math.min(3, (record.followups || []).length); }, 0)');
    clearDownloads();
    await page.setDownloadBehavior(DOWNLOADS);
    await clearToasts(page);
    await click(page, '[data-testid="btn-export-excel"]');
    const filePath = await waitForDownload(/\.xlsx$/i);
    const filename = path.basename(filePath);
    assert(/^瀹㈡埛璺熻繘鍙拌处_\d{8}_\d{4}\.xlsx$/.test(filename), '瀵煎嚭鏂囦欢鍚嶆牸寮忓簲涓?瀹㈡埛璺熻繘鍙拌处_YYYYMMDD_HHmm.xlsx', filename);
    assert(filename.indexOf('瀹㈡埛璺熻繘鍙拌处_' + localDateStamp() + '_') === 0, '瀵煎嚭鏂囦欢鍚嶆棩鏈熷簲涓哄綋澶?, { filename, today: localDateStamp() });
    await waitToast(page, 'ok', '宸插鍑?12 浣嶅鎴?);
    scratch.export = { filePath, filename, expectedNames, expectedDetailRows };
  });

  await test('F2', '鏂囦欢缁撴瀯锛歅K 澶淬€佷綋绉€乑IP 閮ㄤ欢瀹屾暣銆佷袱涓伐浣滆〃 XML銆佽鏁版纭?, async () => {
    const state = scratch.export;
    assert(state, '鍓嶇疆瀵煎嚭鐢ㄤ緥鏈垚鍔?);
    const buffer = fs.readFileSync(state.filePath);
    assert(buffer.length > 3000, 'xlsx 浣撶Н寮傚父', { bytes: buffer.length });
    assert(
      buffer[0] === 0x50 && buffer[1] === 0x4b && buffer[2] === 0x03 && buffer[3] === 0x04,
      'xlsx 鏂囦欢澶翠笉鏄?PK\\x03\\x04',
      buffer.subarray(0, 8).toString('hex')
    );
    const entries = readZip(buffer);
    const requiredParts = [
      '[Content_Types].xml', '_rels/.rels', 'xl/workbook.xml', 'xl/_rels/workbook.xml.rels',
      'xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml', 'xl/styles.xml'
    ];
    for (const part of requiredParts) {
      assert(entries.has(part), 'xlsx 缂哄皯 ZIP 閮ㄤ欢锛? + part, Array.from(entries.keys()));
    }
    const workbookXml = entries.get('xl/workbook.xml').toString('utf8');
    assert(workbookXml.indexOf('瀹㈡埛鎬昏〃') !== -1, 'xl/workbook.xml 缂哄皯宸ヤ綔琛ㄥ悕銆屽鎴锋€昏〃銆?);
    assert(workbookXml.indexOf('璺熻繘鏄庣粏') !== -1, 'xl/workbook.xml 缂哄皯宸ヤ綔琛ㄥ悕銆岃窡杩涙槑缁嗐€?);
    const sheet1Xml = entries.get('xl/worksheets/sheet1.xml').toString('utf8');
    const sheet2Xml = entries.get('xl/worksheets/sheet2.xml').toString('utf8');
    assertEq(countXmlRows(sheet1Xml), state.expectedNames.length + 1, '瀹㈡埛鎬昏〃琛屾暟锛堟暟鎹 + 琛ㄥご锛?);
    assertEq(countXmlRows(sheet2Xml), state.expectedDetailRows + 1, '璺熻繘鏄庣粏琛屾暟锛堣窡杩涜 + 琛ㄥご锛?);
    assert(sheet1Xml.indexOf(state.expectedNames[0]) !== -1, '瀹㈡埛鎬昏〃搴斿寘鍚鎴蜂腑鏂囧悕', state.expectedNames[0]);
  });

  await test('F3', 'openpyxl 鏍￠獙锛氫袱涓伐浣滆〃 / 19 鍒楄〃澶?/ 琛屾暟 / 鍐荤粨 / 鑷姩绛涢€?/ 瀹㈡埛鍚嶉『搴?, async () => {
    const state = scratch.export;
    assert(state, '鍓嶇疆瀵煎嚭鐢ㄤ緥鏈垚鍔?);
    if (!python) {
      console.log('  SKIP  F3  openpyxl 涓嶅彲鐢紝宸茬敱 F2 Node ZIP/XML 鏂█瑕嗙洊');
      return;
    }
    const data = runPythonXlsxCheck(state.filePath);
    assertEq(data.sheetnames.join('|'), '瀹㈡埛鎬昏〃|璺熻繘鏄庣粏', '宸ヤ綔琛ㄥ悕涓庨『搴?);
    const sheet1 = data.sheets['瀹㈡埛鎬昏〃'];
    const sheet2 = data.sheets['璺熻繘鏄庣粏'];
    assertEq(sheet1.max_row, state.expectedNames.length + 1, 'openpyxl 瀹㈡埛鎬昏〃琛屾暟');
    assertEq(sheet1.max_column, 19, '瀹㈡埛鎬昏〃鍒楁暟锛?9锛?);
    assertEq(sheet1.header.join('|'), EXPECTED_TOTAL_HEADERS.join('|'), '瀹㈡埛鎬昏〃 19 鍒楄〃澶存枃鏈?);
    assertEq(sheet1.freeze, 'A2', '瀹㈡埛鎬昏〃鍐荤粨棣栬');
    assert(!!sheet1.auto_filter, '瀹㈡埛鎬昏〃鏈缃嚜鍔ㄧ瓫閫?);
    assertEq(sheet1.names.join('|'), state.expectedNames.join('|'), '瀹㈡埛鎬昏〃瀹㈡埛鍚嶉『搴?= 瀵煎嚭璁板綍椤哄簭');
    assertEq(sheet2.max_row, state.expectedDetailRows + 1, 'openpyxl 璺熻繘鏄庣粏琛屾暟');
    assertEq(sheet2.max_column, 10, '璺熻繘鏄庣粏鍒楁暟锛?0锛?);
  });

  await test('F4', '鏈夌瓫閫夋椂鍙鍑虹瓫閫夌粨鏋滐紙宸叉垚浜?2 鏉?鈫?宸ヤ綔琛?3 琛岋級', async () => {
    await page.fill('[data-testid="filter-status"]', '宸叉垚浜?);
    await waitRowCount(page, 2);
    const filtered = await filteredNames(page);
    assertEq(filtered.length, 2, '绛涢€夊悗璁板綍鏁?);
    const filteredDetailRows = await page.evaluate('window.CustomerCRM.app.filtered().reduce(function (n, record) { return n + Math.min(3, (record.followups || []).length); }, 0)');
    clearDownloads();
    await page.setDownloadBehavior(DOWNLOADS);
    await clearToasts(page);
    await click(page, '[data-testid="btn-export-excel"]');
    const filePath = await waitForDownload(/\.xlsx$/i);
    const filename = path.basename(filePath);
    assert(/^瀹㈡埛璺熻繘鍙拌处_\d{8}_\d{4}\.xlsx$/.test(filename), '绛涢€夊鍑烘枃浠跺悕鏍煎紡', filename);
    await waitToast(page, 'ok', '宸插鍑?2 浣嶅鎴?);
    const entries = readZip(fs.readFileSync(filePath));
    const sheet1Xml = entries.get('xl/worksheets/sheet1.xml').toString('utf8');
    const sheet2Xml = entries.get('xl/worksheets/sheet2.xml').toString('utf8');
    assertEq(countXmlRows(sheet1Xml), filtered.length + 1, '绛涢€夊鍑哄鎴锋€昏〃琛屾暟 = 绛涢€夌粨鏋?+ 琛ㄥご');
    assertEq(countXmlRows(sheet2Xml), filteredDetailRows + 1, '绛涢€夊鍑鸿窡杩涙槑缁嗚鏁?);
    if (python) {
      const data = runPythonXlsxCheck(filePath);
      assertEq(data.sheets['瀹㈡埛鎬昏〃'].names.join('|'), filtered.join('|'), 'openpyxl 璇诲洖鐨勫鎴峰悕 = 褰撳墠绛涢€夌粨鏋?);
    }
    await click(page, '[data-testid="filter-reset"]');
    await waitRowCount(page, 12);
  });

  await assertNoErrors(page, 'F5', '瀵煎嚭娴佺▼鏃?console / page 閿欒');
}

/* ================================================================== */
/* G. 璁剧疆涓庢暟鎹鐞?                                                   */
/* ================================================================== */

async function groupG(page) {
  setGroup('G', '璁剧疆銆丣SON 澶囦唤 / 鎭㈠涓庢暟鎹鐞?);

  await test('G1', '璁剧疆椤垫覆鏌擄細榛樿闃堝€?7/14锛屾暟鎹鐞嗘寜閽綈鍏?, async () => {
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
    assertEq(info.warn, '7', '榛樿寰呰窡杩涢槇鍊?);
    assertEq(info.stale, '14', '榛樿瓒呮湡闃堝€?);
    assert(info.seed && info.clear, '鏁版嵁绠＄悊鎸夐挳 data-testid 缂哄け', info);
    assert(info.exportJson && info.importJson, 'JSON 澶囦唤 / 鎭㈠鎸夐挳缂哄け', info);
  });

  await test('G2', '鎴浘 screenshots/settings.png锛?440脳900锛?, async () => {
    await takeScreenshot(page, 'settings.png');
    assetScreenshot('settings.png', 1440, 900);
  });

  await test('G3', '淇濆瓨闃堝€?warnDays=3 / staleDays=5 骞舵寔涔呭寲锛涢潪娉曢槇鍊艰鎷掔粷', async () => {
    await page.fill('#setting-warnDays', '3');
    await page.fill('#setting-staleDays', '5');
    await clearToasts(page);
    await click(page, '[data-action="save-settings"]');
    await waitExpr(page, `window.CustomerCRM.store.settings().warnDays === 3 && window.CustomerCRM.store.settings().staleDays === 5`, '璁剧疆鍐欏叆 store');
    await waitToast(page, 'ok', '璁剧疆宸蹭繚瀛?);
    await page.fill('#setting-warnDays', '30');
    await page.fill('#setting-staleDays', '10');
    await clearToasts(page);
    await click(page, '[data-action="save-settings"]');
    await waitToast(page, 'error', '瓒呮湡闃堝€?);
    assertEq(await page.evaluate('window.CustomerCRM.store.settings().warnDays'), 3, '闈炴硶闃堝€间笉搴斿啓鍏?store');
    assertEq(await page.evaluate('window.CustomerCRM.store.settings().staleDays'), 5, '闈炴硶闃堝€间笉搴斿啓鍏?store');
  });

  await test('G4', '浠〃鐩樼粺璁￠殢闃堝€煎彉鍖栵細4 鍗?DOM = stats()锛岃秴鏈熸暟澧炲姞锛宧int 鏂囨鏇存柊', async () => {
    await goto(page, '#/dashboard');
    await waitSelector(page, '[data-testid="stat-overdue"]');
    const stats = await page.evaluate('window.CustomerCRM.app.stats()');
    const dom = await readStatsDom(page);
    assertEq(dom.total, String(stats.total), '瀹㈡埛鎬绘暟鍗?= stats.total');
    assertEq(dom.deal, String(stats.deal), '宸叉垚浜ゅ崱 = stats.deal');
    assertEq(dom.open, String(stats.open), '寰呰窡杩涘崱 = stats.open');
    assertEq(dom.overdue, String(stats.overdue), '瓒呮湡鏈窡杩涘崱 = stats.overdue');
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
    const isOpen = (record) => record.dealStatus !== '宸叉垚浜? && record.followupStatus !== '宸叉祦澶?;
    const expectedOpen = context.records.filter(isOpen).length;
    const expectedOverdue5 = context.records.filter((record) => isOpen(record) && record.last && daysBetween(record.last, context.today) >= 5).length;
    const expectedOverdue14 = context.records.filter((record) => isOpen(record) && record.last && daysBetween(record.last, context.today) >= 14).length;
    assertEq(stats.open, expectedOpen, 'stats.open = 鏈垚浜や笖鏈祦澶辫褰曟暟');
    assertEq(Number(dom.overdue), expectedOverdue5, 'staleDays=5 鏃惰秴鏈熸湭璺熻繘鏁伴噺');
    assert(expectedOverdue5 > expectedOverdue14, 'staleDays 鐢?14 璋冭嚦 5 鍚庤秴鏈熸暟閲忓簲澧炲姞', { expectedOverdue5, expectedOverdue14 });
    const hints = await readStatHints(page);
    assert(String(hints['stat-open']).indexOf('鈮?') !== -1, '寰呰窡杩涘崱鐗?hint 搴旀樉绀?warnDays=3', hints['stat-open']);
    assert(String(hints['stat-overdue']).indexOf('鈮?') !== -1, '瓒呮湡鍗＄墖 hint 搴旀樉绀?staleDays=5', hints['stat-overdue']);
  });

  await test('G5', '瀹㈡埛鍙拌处 age 棰滆壊鍒嗗竷 = store.deriveAll()锛堝綋鍓嶉槇鍊硷級', async () => {
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
    assertEq(JSON.stringify(domTones), JSON.stringify(expectedTones), 'age 棰滆壊鍒嗗竷 = deriveAll(褰撳墠闃堝€?');
    assertEq(domTones.ok + domTones.warn + domTones.danger + domTones.none, 12, 'age 鍗曞厓鏍兼暟閲?);
  });

  await test('G6', '瀵煎嚭 JSON 澶囦唤锛氭枃浠跺悕銆佸悎娉?JSON銆?2 鏉″鎴?, async () => {
    await goto(page, '#/settings');
    await waitSelector(page, '[data-action="export-json"]');
    clearDownloads();
    await page.setDownloadBehavior(DOWNLOADS);
    await clearToasts(page);
    await click(page, '[data-action="export-json"]');
    const filePath = await waitForDownload(/\.json$/i);
    const filename = path.basename(filePath);
    assert(/^瀹㈡埛閫氬浠絖\d{8}_\d{4}\.json$/.test(filename), 'JSON 澶囦唤鏂囦欢鍚嶆牸寮?, filename);
    const text = fs.readFileSync(filePath, 'utf8');
    let payload = null;
    try { payload = JSON.parse(text); } catch (err) { fail('JSON 澶囦唤鏃犳硶瑙ｆ瀽', String(err && err.message)); }
    assertEq(payload.app, 'CustomerCRM', '澶囦唤 app 鏍囪瘑');
    assertEq(payload.version, '1.0.0', '澶囦唤鐗堟湰鍙?);
    assertEq(payload.customers.length, 12, '澶囦唤瀹㈡埛鏁?);
    assert(payload.customers.some((record) => record.customerName === 'Michael Braun'), '澶囦唤搴斿寘鍚ず渚嬪鎴?Michael Braun');
    assert(payload.settings && typeof payload.settings.warnDays === 'number', '澶囦唤搴斿寘鍚缃?, payload.settings);
    scratch.backupPath = path.join(TEMP, 'customer-crm-backup.json');
    fs.copyFileSync(filePath, scratch.backupPath);
  });

  await test('G7', '娓呯┖鍏ㄩ儴鏁版嵁 鈫?绌虹姸鎬佸紩瀵?+ 浠〃鐩樼粺璁″綊闆?, async () => {
    await clearToasts(page);
    await click(page, '[data-testid="btn-clear-all"]');
    await waitSelector(page, '[data-confirm-ok]', '娓呯┖纭寮圭獥');
    const confirmText = String(await page.text('[data-modal-id] .confirm-text'));
    assert(confirmText.indexOf('鍒犻櫎鏈満淇濆瓨鐨勫叏閮ㄥ鎴疯褰?) !== -1, '娓呯┖纭鏂囨', confirmText);
    await click(page, '[data-confirm-ok]');
    await waitStoreCount(page, 0);
    await goto(page, '#/customers');
    await waitSelector(page, '.empty', '绌虹姸鎬佸紩瀵?);
    const empty = await page.evaluate(`(() => ({
      title: (document.querySelector('.empty-title') || {}).textContent || '',
      text: (document.querySelector('.empty-text') || {}).textContent || '',
      actions: Array.from(document.querySelectorAll('.empty-actions .btn')).map(function (btn) { return btn.textContent.trim(); })
    }))()`);
    assertEq(empty.title.trim(), '杩樻病鏈夊鎴疯褰?, '绌虹姸鎬佹爣棰?);
    assert(empty.text.length > 0, '绌虹姸鎬佽鏄庢枃妗堜负绌?);
    const actions = empty.actions.join('|');
    assert(actions.indexOf('鏅鸿兘璇嗗埆') !== -1 && actions.indexOf('鏂板缓瀹㈡埛') !== -1 && actions.indexOf('杞藉叆绀轰緥鏁版嵁') !== -1, '绌虹姸鎬佸紩瀵兼寜閽笉瀹屾暣', actions);
    assertEq(await rowCount(page), 0, '绌虹姸鎬佷笅鏁版嵁琛屾暟閲?);
    await goto(page, '#/dashboard');
    await waitSelector(page, '[data-testid="stat-total"]');
    assertEq(String(await page.text('[data-testid="stat-total"]')).trim(), '0', '娓呯┖鍚庝华琛ㄧ洏瀹㈡埛鎬绘暟 = 0');
  });

  await test('G8', '浠?JSON 澶囦唤鎭㈠ 鈫?12 鏉″鎴峰洖琛?, async () => {
    const backupPath = scratch.backupPath;
    assert(backupPath && fs.existsSync(backupPath), '鍓嶇疆 JSON 澶囦唤鐢ㄤ緥鏈垚鍔?);
    await goto(page, '#/settings');
    await waitSelector(page, '[data-action="import-json"]');
    await clearToasts(page);
    await click(page, '[data-action="import-json"]');
    await waitExpr(page, `!!document.querySelector('input[type="file"]')`, '闅愯棌鏂囦欢閫夋嫨妗嗗嚭鐜?);
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
    assert(injected, '鏃犳硶鍚戞枃浠堕€夋嫨妗嗘敞鍏ュ浠芥枃浠?);
    await waitStoreCount(page, 12, 15000);
    await waitToast(page, 'ok', '宸蹭粠澶囦唤瀵煎叆');
    await goto(page, '#/customers');
    await waitRowCount(page, 12);
    const restored = await page.evaluate(`window.CustomerCRM.store.list().filter(function (record) { return record.customerName === 'Michael Braun'; }).length`);
    assertEq(restored, 1, '鎭㈠鍚庡簲鍖呭惈绀轰緥瀹㈡埛 Michael Braun');
    await page.evaluate(`(() => { const input = document.querySelector('input[type="file"]'); if (input && input.parentNode) input.parentNode.removeChild(input); })()`);
  });

  await test('G9', '鍐嶆娓呯┖鍚庛€岃浇鍏ョず渚嬫暟鎹€嶆仮澶?12 鏉?, async () => {
    await goto(page, '#/settings');
    await waitSelector(page, '[data-testid="btn-clear-all"]');
    await click(page, '[data-testid="btn-clear-all"]');
    await waitSelector(page, '[data-confirm-ok]', '娓呯┖纭寮圭獥');
    await click(page, '[data-confirm-ok]');
    await waitStoreCount(page, 0);
    await goto(page, '#/customers');
    await waitSelector(page, '.empty', '绌虹姸鎬佸紩瀵?);
    await clearToasts(page);
    await click(page, '[data-testid="btn-seed-demo"]');
    await waitStoreCount(page, 12);
    await waitRowCount(page, 12);
    await waitToast(page, 'ok', '宸茶浇鍏?12 浣嶇ず渚嬪鎴?);
    const footer = String(await page.text('.table-count'));
    assert(footer.indexOf('12') !== -1, '琛ㄦ牸搴曢儴璁℃暟搴斿惈 12', footer);
  });

  await assertNoErrors(page, 'G10', '璁剧疆涓庢暟鎹鐞嗘祦绋嬫棤 console / page 閿欒');
}

/* ================================================================== */
/* H. 鍝嶅簲寮?/ 鎶藉眽 / 鎵撳嵃                                              */
/* ================================================================== */

async function groupH(page) {
  setGroup('H', '鍝嶅簲寮忋€佷晶鏍忔娊灞変笌鎵撳嵃濯掍綋');

  await test('H1', '1440脳900锛氭棤妯悜婧㈠嚭銆佷晶鏍忓彲瑙併€佸叏閮ㄥ垪鏄剧ず', async () => {
    await resetDemo(page);
    await page.setViewport(1440, 900);
    await goto(page, '#/customers');
    await waitRowCount(page, 12);
    await assertNoOverflow(page, '1440');
    const info = await sidebarRect(page);
    assert(info.left >= 0 && info.width > 200, '1440 渚ф爮搴斿畬鏁村彲瑙?, info);
    for (const key of ['foundedAt', 'teamSize', 'annualRevenue', 'product', 'companyName', 'coreReason']) {
      assertEq(await columnVisible(page, key), true, '1440 搴旀樉绀哄垪 ' + key);
    }
  });

  await test('H2', '1024脳800锛氭棤妯悜婧㈠嚭銆佷晶鏍忔敹璧枫€侀殣钘?3 鍒?, async () => {
    await page.setViewport(1024, 800);
    await goto(page, '#/customers');
    await waitRowCount(page, 12);
    await assertNoOverflow(page, '1024');
    const info = await sidebarRect(page);
    assert(info.right <= 1, '1024 渚ф爮搴旈粯璁ゆ敹璧凤紙translateX(-100%)锛?, info);
    for (const key of ['foundedAt', 'teamSize', 'annualRevenue']) {
      assertEq(await columnVisible(page, key), false, '1024 搴旈殣钘忓垪 ' + key);
    }
    assertEq(await columnVisible(page, 'product'), true, '1024 搴旀樉绀哄挩璇骇鍝佸垪');
  });

  await test('H3', '768脳900锛氭棤妯悜婧㈠嚭銆侀殣钘忎骇鍝?鍏徃鍚?鏍稿績鍘熷洜', async () => {
    await page.setViewport(768, 900);
    await goto(page, '#/customers');
    await waitRowCount(page, 12);
    await assertNoOverflow(page, '768');
    const info = await sidebarRect(page);
    assert(info.right <= 1, '768 渚ф爮搴旈粯璁ゆ敹璧?, info);
    for (const key of ['product', 'companyName', 'coreReason']) {
      assertEq(await columnVisible(page, key), false, '768 搴旈殣钘忓垪 ' + key);
    }
    assertEq(await columnVisible(page, 'customerName'), true, '768 搴旀樉绀哄鎴峰悕绉板垪');
  });

  await test('H4', '390脳844锛氭棤妯悜婧㈠嚭銆佷晶鏍忔敹璧枫€佽彍鍗曟寜閽彲瑙?, async () => {
    await page.setViewport(390, 844, { mobile: true });
    await goto(page, '#/customers');
    await waitRowCount(page, 12);
    await assertNoOverflow(page, '390');
    const info = await sidebarRect(page);
    assert(info.right <= 1, '390 渚ф爮搴旈粯璁ゆ敹璧?, info);
    const toggleVisible = await page.evaluate(`(function () {
      const el = document.querySelector('.menu-toggle');
      if (!el) return null;
      const style = getComputedStyle(el);
      const rect = el.getBoundingClientRect();
      return style.display !== 'none' && style.visibility !== 'hidden' && rect.width > 0;
    })()`);
    assertEq(toggleVisible, true, '390 鑿滃崟鎸夐挳搴斿彲瑙?);
  });

  await test('H5', '390 渚ф爮鎶藉眽锛氳彍鍗曟寜閽墦寮€锛堥伄缃╁彲瑙侊級鈫?鐐归伄缃╁叧闂?, async () => {
    await click(page, '.menu-toggle');
    await waitExpr(page, `document.getElementById('sidebar').classList.contains('is-open')`, '渚ф爮鎶藉眽鎵撳紑');
    const openState = await page.evaluate(`(() => {
      const sidebar = document.getElementById('sidebar');
      const backdrop = document.getElementById('sidebar-backdrop');
      const rect = sidebar.getBoundingClientRect();
      const backdropStyle = getComputedStyle(backdrop);
      return {
        left: rect.left,
        width: rect.width,
        zIndex: getComputedStyle(sidebar).zIndex,
        backdropVisible: backdropStyle.visibility !== 'hidden' && Number(backdropStyle.opacity) > 0.1,
        ariaExpanded: document.querySelector('.menu-toggle').getAttribute('aria-expanded')
      };
    })()`);
    assert(openState.left >= -1 && openState.width > 200, '鎶藉眽鎵撳紑鍚庝晶鏍忓簲婊戝叆瑙嗗彛', openState);
    assert(openState.backdropVisible, '鎶藉眽鎵撳紑鍚庨伄缃╁簲鍙', openState);
    assertEq(openState.ariaExpanded, 'true', '鑿滃崟鎸夐挳 aria-expanded');
    await click(page, '.sidebar-backdrop');
    await waitExpr(page, `!document.getElementById('sidebar').classList.contains('is-open')`, '鐐归伄缃╁悗鎶藉眽鍏抽棴');
    const closedState = await sidebarRect(page);
    assert(closedState.right <= 1, '鎶藉眽鍏抽棴鍚庝晶鏍忓簲鍦ㄨ鍙ｅ', closedState);
  });

  await test('H6', '鎵撳嵃濯掍綋锛氫晶鏍?椤舵爮/宸ュ叿鏍?鎿嶄綔鍒楅殣钘忥紝琛ㄥご table-header-group', async () => {
    await page.setViewport(1440, 900);
    await goto(page, '#/customers');
    await waitRowCount(page, 12);
    await page.emulateMedia('print');
    await waitExpr(page, `getComputedStyle(document.querySelector('.sidebar')).display === 'none'`, '鎵撳嵃鏍峰紡鐢熸晥');
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
      assert(printState[key] && printState[key].display === 'none', '鎵撳嵃濯掍綋涓嬪簲闅愯棌锛? + key, printState[key]);
    }
    assert(printState.viewVisible && printState.viewVisible.display !== 'none', '鎵撳嵃濯掍綋涓嬪唴瀹瑰簲淇濇寔鍙', printState.viewVisible);
    assertEq(printState.theadDisplay, 'table-header-group', '鎵撳嵃濯掍綋涓?thead 搴旈噸澶嶈〃澶?);
    await page.emulateMedia('');
    await waitExpr(page, `getComputedStyle(document.querySelector('.sidebar')).display !== 'none'`, '鎭㈠ screen 濯掍綋');
    await clearToasts(page);
  });

  await test('H7', '鎴浘 mobile-dashboard.png 涓?mobile-customers.png锛?90脳844锛?, async () => {
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

  await assertNoErrors(page, 'H8', '鍝嶅簲寮?/ 鎵撳嵃娴佺▼鏃?console / page 閿欒');
}

/* ================================================================== */
/* I. 鏁版嵁鎹熷潖鍏滃簳                                                      */
/* ================================================================== */

async function groupI(page) {
  setGroup('I', '鏁版嵁鎹熷潖鍏滃簳');

  await test('I1', 'localStorage 闈炴硶 JSON 鈫?鍒锋柊鍚庝笉鐧藉睆锛岃嚜鍔ㄩ噸寤?12 鏉＄ず渚嬫暟鎹?, async () => {
    await resetDemo(page);
    await page.setViewport(1440, 900);
    await page.evaluate(`localStorage.setItem('customer-crm.v1', '{bad json')`);
    page.clearErrors();
    await page.reload();
    await page.waitForApp({ timeout: 25000 });
    await waitExpr(page, `!!document.querySelector('[data-testid="stat-total"]')`, '鎹熷潖鏁版嵁鍚庝粛娓叉煋浠〃鐩?, 15000);
    const health = await page.evaluate(`(() => {
      let parsed = null;
      let error = '';
      try { parsed = JSON.parse(localStorage.getItem('customer-crm.v1')); } catch (err) { error = String(err && err.message); }
      return {
        total: window.CustomerCRM.store.list().length,
        domTotal: (document.querySelector('[data-testid="stat-total"]') || {}).textContent || '',
        viewLength: (document.getElementById('view-inner') || {}).innerHTML.length || 0,
        storedCount: parsed && parsed.customers ? parsed.customers.length : null,
        error: error
      };
    })()`);
    assertEq(health.total, 12, '鎹熷潖鍚庣ず渚嬫暟鎹潯鏁帮紙store锛?);
    assertEq(health.domTotal.trim(), '12', '鎹熷潖鍚庝华琛ㄧ洏瀹㈡埛鎬绘暟');
    assert(health.viewLength > 1000, '椤甸潰鐤戜技鐧藉睆锛?view-inner 鍐呭杩囩煭锛?, health.viewLength);
    assertEq(health.storedCount, 12, '鎹熷潖鏁版嵁搴旇閲嶅啓涓哄悎娉?payload');
    assertEq(health.error, '', 'localStorage 搴旀仮澶嶄负鍚堟硶 JSON');
    const errors = page.errors();
    assert(errors.length === 0, '鎹熷潖鍏滃簳杩囩▼涓嶅簲浜х敓 pageerror / console error', errors.slice(0, 4));
  });
}

/* ================================================================== */
/* S. 浜や粯鐗╂牎楠?                                                       */
/* ================================================================== */

async function groupS() {
  setGroup('S', '浜や粯鐗╂牎楠?);

  await test('S1', '8 寮犻獙鏀舵埅鍥鹃綈澶囦笖灏哄姝ｇ‘锛?440脳900 / 390脳844锛?, async () => {
    const problems = [];
    for (const spec of SHOT_SPECS) {
      const filePath = path.join(SHOTS, spec.file);
      if (!fs.existsSync(filePath)) { problems.push(spec.file + ' 缂哄け'); continue; }
      const size = pngSize(filePath);
      if (!size) { problems.push(spec.file + ' 涓嶆槸鍚堟硶 PNG'); continue; }
      if (size.bytes < 20000) problems.push(spec.file + ' 浣撶Н杩囧皬 ' + size.bytes + 'B');
      if (size.width !== spec.width || size.height !== spec.height) {
        problems.push(spec.file + ' 灏哄 ' + size.width + 'x' + size.height + ' 鏈熸湜 ' + spec.width + 'x' + spec.height);
      }
    }
    assertEq(problems.join('锛?), '', '鎴浘妫€鏌ュけ璐?);
  });

  await test('S2', '鎴浘鏂囦欢鍛藉悕涓?README 寮曠敤璺緞涓€鑷达紙screenshots/*.png锛?, async () => {
    for (const spec of SHOT_SPECS) {
      assert(/^[a-z-]+\.png$/.test(spec.file), '鎴浘鏂囦欢鍚嶄笉绗﹀悎瑙勮寖', spec.file);
      assert(fs.existsSync(path.join(SHOTS, spec.file)), '鎴浘缂哄け', spec.file);
    }
  });
}

/* ================================================================== */
/* 涓绘祦绋?                                                              */
/* ================================================================== */

python = findPython();
console.log('瀹㈡埛閫?CustomerCRM 鈥?绔埌绔獙鏀讹紙E2E锛?);
console.log('椤圭洰鏍癸細' + ROOT);
console.log('Node锛? + process.version + ' / ' + process.platform + ' / ' + os.release());
console.log('鎴浘鐩綍锛? + SHOTS);
console.log('openpyxl锛? + (python ? python.version + ' @ ' + python.command : '涓嶅彲鐢?鈫?闄嶇骇涓?Node ZIP/XML 鏂█'));

let chrome = null;
let server = null;
let page = null;

try {
  fs.mkdirSync(SHOTS, { recursive: true });
  server = await serveStatic(ROOT);
  chrome = await launchChrome({ windowSize: { width: 1440, height: 900 } });
  console.log('Chrome锛? + (chrome.devtools && chrome.devtools.Browser ? chrome.devtools.Browser : chrome.chromePath));
  console.log('闈欐€佹湇鍔★細' + server.url);
  const connection = await connectToPage(chrome);
  page = await PageSession.create(connection.client);

  await page.setViewport(1440, 900);
  await page.setDownloadBehavior(DOWNLOADS);

  await page.navigate(server.url + '/index.html');
  await page.waitForApp({ timeout: 25000 });
  await waitExpr(page, `!!document.querySelector('#view-inner .page')`, '棣栧睆娓叉煋', 20000);

  // 闇€姹傦細姣忔 e2e 杩愯寮€澶存竻绌哄苟閲嶅缓绀轰緥鏁版嵁锛屼繚璇佸彲閲嶅
  await page.evaluate('localStorage.clear()');
  page.clearErrors();
  await page.reload();
  await page.waitForApp({ timeout: 25000 });
  await waitExpr(page, `!!document.querySelector('#view-inner .page')`, '閲嶈浇鍚庨灞忔覆鏌?, 20000);
  await waitStoreCount(page, 12, 15000);

  await groupA(page);
  await groupB(page);
  await groupC(page);
  await groupD(page);
  await groupE(page);
  await groupF(page);
  await groupG(page);
  await groupH(page);
  await groupI(page);
  await groupS();
} catch (err) {
  results.push({
    id: 'FATAL',
    group: group.id,
    title: 'E2E 杩愯涓柇',
    ok: false,
    detail: (err && err.stack) || String(err),
    ms: 0
  });
  console.error('\nE2E 杩愯涓柇锛?, err && err.stack ? err.stack : err);
} finally {
  try { if (page) page.close(); } catch { /* ignore */ }
  try { if (chrome) await chrome.close(); } catch { /* ignore */ }
  try { if (server) await server.close(); } catch { /* ignore */ }
}

/* ---------------- 姹囨€?---------------- */

const failed = results.filter((result) => !result.ok);
const passed = results.length - failed.length;
const groupIds = Array.from(new Set(results.map((result) => result.group)));
const groupTitles = {};
for (const result of results) groupTitles[result.group] = result.group;

console.log('\n================ E2E 姹囨€?================');
for (const groupId of groupIds) {
  const groupResults = results.filter((result) => result.group === groupId);
  const groupFailed = groupResults.filter((result) => !result.ok).length;
  console.log(`  ${groupId}锛?{groupResults.length - groupFailed}/${groupResults.length} PASS` + (groupFailed ? `锛?{groupFailed} FAIL锛塦 : ''));
}
console.log(`\n鐢ㄤ緥锛?{results.length} 涓紝PASS ${passed}锛孎AIL ${failed.length}`);
console.log('鎴浘鐩綍锛? + SHOTS);
console.log('涓嬭浇/涓存椂鐩綍锛? + TEMP);
if (python) console.log('openpyxl 鏍￠獙锛? + python.version);
else console.log('openpyxl 鏍￠獙锛氫笉鍙敤锛堜娇鐢?Node ZIP/XML 鏂█锛?);

if (failed.length) {
  console.log('\n澶辫触鐢ㄤ緥涓庤鎯咃細');
  for (const result of failed) {
    console.log(`  鉁?${result.id}  [${result.group}] ${result.title}`);
    console.log('      ' + String(result.detail).split('\n').join('\n      '));
  }
  process.exitCode = 1;
} else {
  console.log('\n鍏ㄩ儴鐢ㄤ緥閫氳繃 鉁?);
  try { fs.rmSync(TEMP, { recursive: true, force: true }); } catch { /* ignore */ }
  process.exitCode = 0;
}
