/**
 * 客户通 CustomerCRM — tests/test-xlsx.mjs
 * 运行：node tests/test-xlsx.mjs
 *
 * 校验 js/xlsx.js（window.CrmXlsx）：
 *   1. 真实生成 .xlsx 到系统临时目录（主产物 %TEMP%/crm-export-test.xlsx）；
 *   2. 用 Python + openpyxl 读回：工作表名、24/13 列表头、行数、中文值、冻结窗格、
 *      autoFilter 覆盖数据区、列宽、表头样式、换行、数字单元格、zipfile.testzip()；
 *   3. 至少 4 个场景：8 条示例客户、空记录、超长中文（500 字）、特殊字符与换行；
 *   4. 底层 buildWorkbook、浏览器 download 的 Node 行为、500 行 × 24 列性能。
 *
 * 任何断言失败 → exit 1 并打印失败明细；全部通过 → exit 0。
 */
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';

import schema from '../js/schema.js';
import xlsx from '../js/xlsx.js';

/* 与浏览器一致：把 schema 挂到全局，让导出引擎用界面同款派生口径（last3 / daysSince）。 */
globalThis.CRM_SCHEMA = schema;

const TODAY = '2026-11-21';

/* ================================================================== */
/* 0. 断言框架                                                         */
/* ================================================================== */

let assertionCount = 0;
const failures = [];
let currentCase = '';

function section(title) {
  currentCase = title;
  console.log('\n── ' + title + ' ──');
}

function ok(name, condition, detail) {
  assertionCount += 1;
  if (condition) {
    console.log('  ✓ ' + name);
    return true;
  }
  const text = detail === undefined || detail === '' ? '' : '  ← ' + String(detail);
  failures.push({ name: currentCase + ' / ' + name, detail: text });
  console.log('  ✗ ' + name + text);
  return false;
}

function eq(name, actual, expected) {
  const same = JSON.stringify(actual) === JSON.stringify(expected);
  return ok(
    name,
    same,
    same ? '' : '期望 ' + JSON.stringify(expected) + '，实际 ' + JSON.stringify(actual)
  );
}

function runCase(title, fn) {
  section(title);
  try {
    fn();
  } catch (err) {
    assertionCount += 1;
    failures.push({
      name: currentCase + ' / 未捕获异常',
      detail: (err && err.stack) || String(err)
    });
    console.log('  ✗ 未捕获异常：' + ((err && err.message) || err));
  }
}

/* ================================================================== */
/* 1. Python + openpyxl 环境（找不到时降级为纯 Node ZIP/XML 断言）      */
/* ================================================================== */

/* CI（Ubuntu + Node 20）优先 python3；本机 Windows 用 bundled python 兜底。 */
const WINDOWS_BUNDLED_PYTHON = 'C:\\Users\\Administrator\\AppData\\Local\\Programs\\DSH Desktop\\resources\\office-runtime\\primary-runtime\\dependencies\\python\\python.exe';

const PYTHON_CANDIDATES = [
  process.env.CRM_PYTHON,
  'python3',
  'python',
  WINDOWS_BUNDLED_PYTHON,
  'py',
  process.env.PYTHON
].filter((item) => item && String(item).trim());

const PYTHON_PACKAGES_CANDIDATES = [
  process.env.CRM_PYTHON_PACKAGES,
  'C:\\Users\\Administrator\\AppData\\Local\\Programs\\DSH Desktop\\resources\\office-runtime\\primary-runtime\\dependencies\\python\\Lib\\site-packages'
].filter(Boolean);

function pythonEnv() {
  const env = { ...process.env, PYTHONIOENCODING: 'utf-8', PYTHONUTF8: '1' };
  const packages = PYTHON_PACKAGES_CANDIDATES.find((item) => fs.existsSync(item));
  if (packages) {
    env.PYTHONPATH = env.PYTHONPATH ? packages + path.delimiter + env.PYTHONPATH : packages;
  }
  return env;
}

let PYTHON = null;
for (const candidate of PYTHON_CANDIDATES) {
  try {
    const version = execFileSync(
      candidate,
      ['-c', 'import openpyxl, sys; sys.stdout.write(openpyxl.__version__)'],
      { encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'], windowsHide: true, env: pythonEnv() }
    ).trim();
    PYTHON = candidate;
    console.log('Python: ' + candidate + '（openpyxl ' + version + '）');
    break;
  } catch (err) {
    /* 尝试下一个候选 */
  }
}

if (process.env.CRM_XLSX_SKIP_PYTHON) {
  /* 调试开关：强制走纯 Node 校验（用于验证 CI 降级路径） */
  PYTHON = null;
  console.log('SKIP openpyxl 校验（CRM_XLSX_SKIP_PYTHON：强制纯 Node ZIP/XML 断言）');
} else if (!PYTHON) {
  console.log('SKIP openpyxl 校验（未找到可用的 Python + openpyxl；继续运行纯 Node ZIP/XML 断言）');
}

const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'crm-xlsx-test-'));
const MAIN_XLSX = path.join(os.tmpdir(), 'crm-export-test.xlsx');
const VERIFY_SCRIPT = path.join(TMP, 'verify_xlsx.py');
const artifactSizes = {};

/* Python 校验脚本：把工作簿“实际长相”完整 dump 成 JSON，断言放在 Node 侧做。 */
const VERIFY_PY = `# -*- coding: utf-8 -*-
import json
import sys
import zipfile

from openpyxl import load_workbook


def scalar(value):
    if value is None:
        return None
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value
    return str(value)


path = sys.argv[1]
out = {}

with zipfile.ZipFile(path) as zf:
    out["testzip"] = zf.testzip()          # None 表示 CRC 全部正确
    out["entries"] = sorted(zf.namelist())

wb = load_workbook(path)
out["sheetnames"] = wb.sheetnames

sheets = []
for name in wb.sheetnames:
    ws = wb[name]
    max_row = ws.max_row
    max_col = ws.max_column
    headers = [scalar(ws.cell(row=1, column=c).value) for c in range(1, max_col + 1)]
    rows = []
    cell_types = []
    number_formats = []
    for r in range(2, max_row + 1):
        rows.append([scalar(ws.cell(row=r, column=c).value) for c in range(1, max_col + 1)])
        cell_types.append([ws.cell(row=r, column=c).data_type for c in range(1, max_col + 1)])
        number_formats.append([ws.cell(row=r, column=c).number_format for c in range(1, max_col + 1)])

    anchor = ws.cell(row=1, column=1)
    fill_rgb = None
    try:
        fill_rgb = anchor.fill.fgColor.rgb if anchor.fill is not None and anchor.fill.fgColor is not None else None
    except Exception:
        fill_rgb = None
    font_rgb = None
    try:
        font_rgb = anchor.font.color.rgb if anchor.font.color is not None else None
    except Exception:
        font_rgb = None

    widths = {}
    for key, dim in ws.column_dimensions.items():
        widths[key] = round(float(dim.width), 2) if dim.width is not None else None

    body_wrap = []
    if max_row >= 2:
        body_wrap = [bool(ws.cell(row=2, column=c).alignment.wrap_text) for c in range(1, max_col + 1)]

    sheets.append({
        "name": name,
        "max_row": max_row,
        "max_col": max_col,
        "headers": headers,
        "rows": rows,
        "cell_types": cell_types,
        "number_formats": number_formats,
        "freeze_panes": ws.freeze_panes,
        "auto_filter": ws.auto_filter.ref,
        "widths": widths,
        "header_height": ws.row_dimensions[1].height,
        "header_bold": bool(anchor.font.bold),
        "header_font_color": font_rgb,
        "header_fill": fill_rgb,
        "header_align": anchor.alignment.horizontal,
        "header_vertical": anchor.alignment.vertical,
        "body_wrap": body_wrap,
    })

out["sheets"] = sheets
out["properties"] = {"title": wb.properties.title, "creator": wb.properties.creator}
sys.stdout.write(json.dumps(out, ensure_ascii=False))
`;

fs.writeFileSync(VERIFY_SCRIPT, VERIFY_PY, 'utf8');

let pythonReads = 0;
let nodeReads = 0;

/* ================================================================== */
/* 1.1 纯 Node XLSX 检查器（无 Python / openpyxl 时的降级校验）         */
/*     自读 ZIP（store）→ 解析 workbook / styles / worksheet XML        */
/* ================================================================== */

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n += 1) {
    let c = n;
    for (let k = 0; k < 8; k += 1) {
      c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    }
    table[n] = c >>> 0;
  }
  return table;
})();

function crc32(buffer) {
  let c = 0xFFFFFFFF;
  for (let i = 0; i < buffer.length; i += 1) {
    c = CRC_TABLE[(c ^ buffer[i]) & 0xFF] ^ (c >>> 8);
  }
  return (c ^ 0xFFFFFFFF) >>> 0;
}

function readU16(bytes, offset) {
  return bytes[offset] | (bytes[offset + 1] << 8);
}

function readU32(bytes, offset) {
  return (bytes[offset] | (bytes[offset + 1] << 8) | (bytes[offset + 2] << 16) | (bytes[offset + 3] << 24)) >>> 0;
}

const TEXT_DECODER = new TextDecoder('utf-8');

/** 自读 store 方式 ZIP：校验 CRC/尺寸/store 方式/UTF-8 文件名标志位。 */
function unzipStored(bytes) {
  let eocd = -1;
  const lowest = Math.max(0, bytes.length - 22 - 65535);
  for (let i = bytes.length - 22; i >= lowest; i -= 1) {
    if (readU32(bytes, i) === 0x06054B50) { eocd = i; break; }
  }
  if (eocd < 0) { throw new Error('不是有效的 ZIP：找不到 EOCD'); }
  const count = readU16(bytes, eocd + 10);
  let cursor = readU32(bytes, eocd + 16);
  const files = new Map();
  let zipOk = true;
  for (let i = 0; i < count; i += 1) {
    if (readU32(bytes, cursor) !== 0x02014B50) { throw new Error('中央目录签名错误 @' + cursor); }
    const flags = readU16(bytes, cursor + 8);
    const method = readU16(bytes, cursor + 10);
    const crc = readU32(bytes, cursor + 16);
    const compSize = readU32(bytes, cursor + 20);
    const size = readU32(bytes, cursor + 24);
    const nameLen = readU16(bytes, cursor + 28);
    const extraLen = readU16(bytes, cursor + 30);
    const commentLen = readU16(bytes, cursor + 32);
    const localOffset = readU32(bytes, cursor + 42);
    const name = TEXT_DECODER.decode(bytes.slice(cursor + 46, cursor + 46 + nameLen));
    if (readU32(bytes, localOffset) !== 0x04034B50) { throw new Error('本地文件头签名错误：' + name); }
    const localNameLen = readU16(bytes, localOffset + 26);
    const localExtraLen = readU16(bytes, localOffset + 28);
    const dataStart = localOffset + 30 + localNameLen + localExtraLen;
    const data = bytes.slice(dataStart, dataStart + size);
    if (method !== 0) { throw new Error('条目不是 store 方式：' + name); }
    if ((flags & 0x0800) === 0) { zipOk = false; }
    if (compSize !== size || data.length !== size || crc32(data) !== crc) { zipOk = false; }
    files.set(name, data);
    cursor += 46 + nameLen + extraLen + commentLen;
  }
  return { testzip: zipOk ? null : 'ZIP/CRC 校验失败', entries: Array.from(files.keys()).sort(), files };
}

function xmlUnescape(text) {
  return String(text)
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (match, code) => String.fromCharCode(Number(code)))
    .replace(/&#x([0-9A-Fa-f]+);/g, (match, code) => String.fromCharCode(parseInt(code, 16)))
    .replace(/&amp;/g, '&');
}

function attrsOf(fragment) {
  const attrs = {};
  const re = /([A-Za-z_:][-A-Za-z0-9_:.]*)="([^"]*)"/g;
  let match;
  while ((match = re.exec(fragment)) !== null) { attrs[match[1]] = xmlUnescape(match[2]); }
  return attrs;
}

function colToIndex(letters) {
  let index = 0;
  for (const ch of String(letters).toUpperCase()) { index = index * 26 + (ch.charCodeAt(0) - 64); }
  return index;
}

function indexToCol(index) {
  let n = index;
  let text = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    text = String.fromCharCode(65 + m) + text;
    n = Math.floor((n - 1) / 26);
  }
  return text || 'A';
}

const BUILTIN_NUMFMT_TEXT = {
  0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00',
  9: '0%', 10: '0.00%', 11: '0.00E+00',
  14: 'm/d/yy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy',
  18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss', 22: 'm/d/yy h:mm'
};

function parseStylesXml(xml) {
  const customNumFmts = {};
  const numFmtBlock = /<numFmts\b[^>]*>([\s\S]*?)<\/numFmts>/.exec(xml);
  if (numFmtBlock) {
    const re = /<numFmt\b([^>]*)\/>/g;
    let match;
    while ((match = re.exec(numFmtBlock[1])) !== null) {
      const attrs = attrsOf(match[1]);
      customNumFmts[Number(attrs.numFmtId)] = attrs.formatCode;
    }
  }

  const fonts = [];
  const fontBlock = /<fonts\b[^>]*>([\s\S]*?)<\/fonts>/.exec(xml);
  if (fontBlock) {
    const re = /<font>([\s\S]*?)<\/font>/g;
    let match;
    while ((match = re.exec(fontBlock[1])) !== null) {
      const body = match[1];
      const colorMatch = /<color\b[^>]*rgb="([0-9A-Fa-f]{8})"/.exec(body);
      fonts.push({ bold: /<b\s*\/>/.test(body), color: colorMatch ? colorMatch[1].toUpperCase() : null });
    }
  }

  const fills = [];
  const fillBlock = /<fills\b[^>]*>([\s\S]*?)<\/fills>/.exec(xml);
  if (fillBlock) {
    const re = /<fill>([\s\S]*?)<\/fill>/g;
    let match;
    while ((match = re.exec(fillBlock[1])) !== null) {
      const colorMatch = /<fgColor\b[^>]*rgb="([0-9A-Fa-f]{8})"/.exec(match[1]);
      fills.push(colorMatch ? colorMatch[1].toUpperCase() : null);
    }
  }

  const xfs = [];
  const xfBlock = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(xml);
  if (xfBlock) {
    const re = /<xf\b([^>]*?)(?:\/>|>([\s\S]*?)<\/xf>)/g;
    let match;
    while ((match = re.exec(xfBlock[1])) !== null) {
      const attrs = attrsOf(match[1]);
      const inner = match[2] || '';
      const alignMatch = /<alignment\b([^>]*)\/>/.exec(inner);
      const align = alignMatch ? attrsOf(alignMatch[1]) : {};
      xfs.push({
        numFmtId: Number(attrs.numFmtId || 0),
        fontId: Number(attrs.fontId || 0),
        fillId: Number(attrs.fillId || 0),
        borderId: Number(attrs.borderId || 0),
        align: align.horizontal || null,
        vertical: align.vertical || null,
        wrap: align.wrapText === '1'
      });
    }
  }

  function numberFormat(id) {
    if (customNumFmts[id] !== undefined) { return customNumFmts[id]; }
    return BUILTIN_NUMFMT_TEXT[id] !== undefined ? BUILTIN_NUMFMT_TEXT[id] : 'General';
  }

  return { fonts, fills, xfs, numberFormat };
}

function parseWorksheetXml(xml, name, styles) {
  const dimensionMatch = /<dimension\b[^>]*ref="([^"]+)"/.exec(xml);
  const dimension = dimensionMatch ? dimensionMatch[1] : 'A1';
  const autoFilterMatch = /<autoFilter\b[^>]*ref="([^"]+)"/.exec(xml);
  const autoFilter = autoFilterMatch ? autoFilterMatch[1] : null;

  const paneMatch = /<pane\b([^>]*)\/>/.exec(xml);
  let freezePanes = null;
  if (paneMatch) {
    const paneAttrs = attrsOf(paneMatch[1]);
    if (paneAttrs.state === 'frozen' && paneAttrs.topLeftCell) { freezePanes = paneAttrs.topLeftCell; }
  }

  const widths = {};
  {
    const re = /<col\b([^>]*)\/>/g;
    let match;
    while ((match = re.exec(xml)) !== null) {
      const attrs = attrsOf(match[1]);
      if (attrs.min !== undefined) {
        widths[indexToCol(Number(attrs.min))] = attrs.width !== undefined
          ? Math.round(Number(attrs.width) * 100) / 100
          : null;
      }
    }
  }

  const rowsByNumber = new Map();
  {
    const rowRe = /<row\b([^>]*?)>([\s\S]*?)<\/row>/g;
    let rowMatch;
    while ((rowMatch = rowRe.exec(xml)) !== null) {
      const rowAttrs = attrsOf(rowMatch[1]);
      const rowNumber = Number(rowAttrs.r);
      const cells = new Map();
      const cellRe = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
      let cellMatch;
      outerCells:
      while ((cellMatch = cellRe.exec(rowMatch[2])) !== null) {
        const cellAttrs = attrsOf(cellMatch[1]);
        const inner = cellMatch[2] || '';
        const ref = cellAttrs.r || '';
        const letters = ref.replace(/\d+/g, '');
        const colIndex = letters ? colToIndex(letters) : 1;
        let value = null;
        let type = 'n';
        if (cellAttrs.t === 'inlineStr') {
          const textMatch = /<t\b[^>]*>([\s\S]*?)<\/t>/.exec(inner);
          value = textMatch ? xmlUnescape(textMatch[1]) : '';
          type = 's';
        } else {
          const valueMatch = /<v>([\s\S]*?)<\/v>/.exec(inner);
          if (valueMatch) { value = Number(valueMatch[1]); type = 'n'; }
        }
        cells.set(colIndex, {
          value,
          type,
          style: cellAttrs.s !== undefined ? Number(cellAttrs.s) : 0
        });
        if (cells.size > 16384) { break outerCells; }
      }
      rowsByNumber.set(rowNumber, {
        ht: rowAttrs.ht !== undefined ? Number(rowAttrs.ht) : null,
        cells
      });
    }
  }

  let dimensionLastCol = 1;
  let dimensionLastRow = 1;
  {
    const parts = dimension.split(':');
    const end = parts[parts.length - 1] || 'A1';
    const endLetters = end.replace(/\d+/g, '');
    const endRow = Number(end.replace(/\D+/g, ''));
    dimensionLastCol = endLetters ? colToIndex(endLetters) : 1;
    dimensionLastRow = endRow || 1;
  }

  let maxCellCol = 1;
  rowsByNumber.forEach((row) => {
    row.cells.forEach((unused, colIndex) => { maxCellCol = Math.max(maxCellCol, colIndex); });
  });
  const colCount = Math.max(dimensionLastCol, maxCellCol, 1);
  let maxRow = Math.max(dimensionLastRow, 1);
  rowsByNumber.forEach((unused, rowNumber) => { maxRow = Math.max(maxRow, rowNumber); });

  const rows = [];
  const cellTypes = [];
  const numberFormats = [];
  for (let r = 2; r <= maxRow; r += 1) {
    const row = rowsByNumber.get(r) || { cells: new Map() };
    const values = [];
    const types = [];
    const formats = [];
    for (let c = 1; c <= colCount; c += 1) {
      const cell = row.cells.get(c);
      values.push(cell ? cell.value : null);
      types.push(cell ? cell.type : 'n');
      formats.push(styles.numberFormat(cell ? Number((styles.xfs[cell.style] || {}).numFmtId || 0) : 0));
    }
    rows.push(values);
    cellTypes.push(types);
    numberFormats.push(formats);
  }

  const headerRow = rowsByNumber.get(1) || { ht: null, cells: new Map() };
  const headers = [];
  for (let c = 1; c <= colCount; c += 1) {
    const cell = headerRow.cells.get(c);
    headers.push(cell ? cell.value : null);
  }
  const anchor = headerRow.cells.get(1);
  const anchorXf = anchor ? (styles.xfs[anchor.style] || styles.xfs[0]) : (styles.xfs[0] || {});
  const anchorFont = styles.fonts[anchorXf.fontId] || null;
  const anchorFill = styles.fills[anchorXf.fillId] !== undefined ? styles.fills[anchorXf.fillId] : null;

  const bodyRow = rowsByNumber.get(2);
  const bodyWrap = [];
  for (let c = 1; c <= colCount; c += 1) {
    if (!bodyRow) { bodyWrap.push(false); continue; }
    const cell = bodyRow.cells.get(c);
    const xf = cell ? styles.xfs[cell.style] : null;
    bodyWrap.push(!!(xf && xf.wrap));
  }

  return {
    name,
    max_row: maxRow,
    max_col: colCount,
    headers,
    rows,
    cell_types: cellTypes,
    number_formats: numberFormats,
    freeze_panes: freezePanes,
    auto_filter: autoFilter,
    widths,
    header_height: headerRow.ht,
    header_bold: !!(anchorFont && anchorFont.bold),
    header_font_color: anchorFont ? anchorFont.color : null,
    header_fill: anchorFill,
    header_align: anchorXf.align || null,
    header_vertical: anchorXf.vertical || null,
    body_wrap: bodyWrap
  };
}

function partText(files, name) {
  const data = files.get(name);
  if (!data) { throw new Error('缺少 XLSX 部件：' + name); }
  return TEXT_DECODER.decode(data);
}

/** 与 Python dump 同构的纯 Node 检查结果，供同一批断言复用。 */
function nodeInspect(bytes) {
  const zip = unzipStored(bytes);
  const workbookXml = partText(zip.files, 'xl/workbook.xml');
  const sheetnames = [];
  {
    const re = /<sheet\b([^>]*?)\/>/g;
    let match;
    while ((match = re.exec(workbookXml)) !== null) {
      const attrs = attrsOf(match[1]);
      if (attrs.name) { sheetnames.push(attrs.name); }
    }
  }
  const styles = parseStylesXml(partText(zip.files, 'xl/styles.xml'));
  let title = null;
  if (zip.files.has('docProps/core.xml')) {
    const titleMatch = /<dc:title>([\s\S]*?)<\/dc:title>/.exec(partText(zip.files, 'docProps/core.xml'));
    title = titleMatch ? xmlUnescape(titleMatch[1]) : null;
  }
  const sheets = sheetnames.map((sheetName, index) => parseWorksheetXml(
    partText(zip.files, 'xl/worksheets/sheet' + (index + 1) + '.xml'),
    sheetName,
    styles
  ));
  return {
    testzip: zip.testzip,
    entries: zip.entries,
    sheetnames,
    properties: { title },
    sheets
  };
}

/**
 * 读回生成物：优先 Python + openpyxl（完整读回），
 * 无 Python / openpyxl 不可用时降级为纯 Node ZIP/XML 解析（同构结果）。
 */
function readWorkbookOrFail(file, label) {
  const bytes = fs.readFileSync(file);
  if (PYTHON) {
    try {
      const raw = execFileSync(PYTHON, [VERIFY_SCRIPT, file], {
        encoding: 'utf8',
        env: pythonEnv(),
        stdio: ['ignore', 'pipe', 'pipe'],
        windowsHide: true,
        maxBuffer: 64 * 1024 * 1024
      });
      pythonReads += 1;
      return JSON.parse(raw);
    } catch (err) {
      const stdout = err && err.stdout ? String(err.stdout) : '';
      const stderr = err && err.stderr ? String(err.stderr) : '';
      if (pythonReads === 0 && nodeReads === 0) {
        /* 环境问题（例如 CI 上 openpyxl 缺失）：降级为纯 Node 断言，整体仍 PASS */
        console.log('SKIP openpyxl 校验（Python + openpyxl 无法运行：' +
          (stderr || stdout || err.message || '').split('\n').filter(Boolean).slice(-2).join(' ').slice(0, 300) +
          '；继续运行纯 Node ZIP/XML 断言）');
        PYTHON = null;
      } else {
        ok(label + '：openpyxl 能正常打开', false, (stderr || stdout || err.message || '').slice(-1500));
        return null;
      }
    }
  }
  try {
    const info = nodeInspect(bytes);
    nodeReads += 1;
    return info;
  } catch (err) {
    ok(label + '：纯 Node 解析 ZIP/XML', false, (err && err.stack) || String(err));
    return null;
  }
}

function sheetOf(info, name) {
  if (!info || !Array.isArray(info.sheets)) {
    return {};
  }
  return info.sheets.find((sheet) => sheet.name === name) || {};
}

function writeXlsx(name, bytes) {
  const file = path.join(TMP, name);
  fs.writeFileSync(file, bytes);
  artifactSizes[name] = bytes.length;
  return file;
}

function expectPkHeader(label, bytes) {
  const isBytes = bytes instanceof Uint8Array;
  ok(label + '：返回 Uint8Array', isBytes, isBytes ? '' : '实际类型 ' + Object.prototype.toString.call(bytes));
  const head = isBytes ? Array.from(bytes.slice(0, 4)) : [];
  eq(label + '：文件头为 PK\\x03\\x04', head, [0x50, 0x4B, 0x03, 0x04]);
  return isBytes;
}

/* ================================================================== */
/* 2. 示例数据（8 条客户 / 15 条跟进）                                  */
/* ================================================================== */

function makeRecord(overrides) {
  return Object.assign({
    id: 'c_test',
    firstInquiryAt: '2026-09-12',
    customerName: '',
    email: '',
    phone: '',
    address: '',
    product: '',
    productModel: '',
    country: '',
    companyName: '',
    mainBusiness: '',
    foundedAt: '',
    teamSize: '',
    annualRevenue: '',
    followupStatus: '新询盘',
    dealStatus: '未成交',
    coreReason: '',
    followups: [],
    createdAt: '2026-09-12T02:10:00.000Z',
    updatedAt: '2026-11-18T09:30:00.000Z'
  }, overrides || {});
}

const SAMPLE_RECORDS = [
  makeRecord({
    firstInquiryAt: '2026-09-12',
    customerName: 'Michael Braun',
    email: 'm.braun@braun-trading.de',
    phone: '+49 711 4567 890',
    address: 'Industriestr. 12, 70565 Stuttgart, Germany',
    companyName: 'Braun Trading GmbH',
    country: '德国',
    product: '不锈钢保温杯 500ml（定制 LOGO）',
    productModel: '500ml',
    mainBusiness: '户外用品进口分销',
    foundedAt: '1998',
    teamSize: '50-100 人',
    annualRevenue: '1200 万欧元',
    followupStatus: '已报价',
    dealStatus: '未成交',
    coreReason: '报价高出目标价约 8%，客户等终端零售客户确认后回复',
    followups: [
      { date: '2026-11-18', note: '邮件确认样品尺寸与 LOGO 印刷位置，等寄样确认' },
      { date: '2026-11-10', note: '电话沟通，客户确认 2000pcs 试单意向' },
      { date: '2026-10-28', note: '首次邮件回复，发送价目表与产品册' }
    ]
  }),
  makeRecord({
    firstInquiryAt: '2026-10-05',
    customerName: 'Ahmed Hassan',
    email: 'a.hassan@elnile-import.com',
    phone: '+20 100 234 5678',
    address: '25 Nile Corniche, Cairo 11511, Egypt',
    companyName: 'El Nile Import & Export Co., Ltd',
    country: '埃及',
    product: '陶瓷餐具套装 16 件套',
    productModel: '16 件套',
    mainBusiness: '家居餐具进口分销',
    foundedAt: '2011',
    teamSize: '15 人',
    annualRevenue: '300 万美元',
    followupStatus: '已回复',
    dealStatus: '未成交',
    coreReason: '已发样册与目录，等待客户确认采购数量',
    followups: [
      { date: '2026-11-20', note: 'WhatsApp 回复，确认 300 万美元年采购额并索取 PI' },
      { date: '2026-11-02', note: '发送陶瓷餐具套装样册与 FOB 报价' },
      { date: '2026-10-20', note: '首次邮件联系，介绍工厂与认证' }
    ]
  }),
  makeRecord({
    firstInquiryAt: '2026-08-08',
    customerName: 'Sophie Martin',
    email: 's.martin@martincuisine.fr',
    phone: '+33 4 78 12 34 56',
    address: '22 Rue de la République, 69002 Lyon, France',
    companyName: 'Maison Martin SAS',
    country: '法国',
    product: '硅胶厨具套装 12 件套',
    productModel: '12 件套',
    mainBusiness: '厨房用品连锁零售',
    foundedAt: '2016',
    teamSize: '10-20 人',
    annualRevenue: '250 万欧元',
    followupStatus: '已成交',
    dealStatus: '已成交',
    coreReason: '价格与交期都满足，首单 5000 套，已收 30% 定金',
    followups: [
      { date: '2026-11-19', note: '收到 30% 定金，安排下周投产' },
      { date: '2026-11-01', note: '客户确认 PI 与包装稿' }
    ]
  }),
  makeRecord({
    firstInquiryAt: '2026-07-15',
    customerName: 'Yuki Tanaka',
    email: 'y.tanaka@tanaka-trading.jp',
    phone: '+81 3 1234 5678',
    address: '2-1-1 Uchisaiwaicho, Chiyoda-ku, Tokyo 100-0011, Japan',
    companyName: 'Tanaka Trading Co., Ltd.',
    country: '日本',
    product: '保温饭盒 1.2L',
    productModel: '1.2L',
    mainBusiness: '日用百货批发',
    foundedAt: '1985',
    teamSize: '30 人',
    annualRevenue: '5 亿日元',
    followupStatus: '谈判中',
    dealStatus: '未成交',
    coreReason: '客户要求 JIS 认证与日文包装，认证周期约 6 周',
    /* 故意打乱顺序 + 4 条，验证「倒序 + 最多 3 条」 */
    followups: [
      { date: '2026-10-01', note: '首次寄送样品与报价单（应被截断）' },
      { date: '2026-11-15', note: '邮件确认交期改善方案，客户内部评审' },
      { date: '2026-10-22', note: '视频会议讨论 JIS 认证与包装' },
      { date: '2026-11-05', note: '发送 45 天交期方案与补偿条款' }
    ]
  }),
  makeRecord({
    firstInquiryAt: '2026-09-30',
    customerName: 'Carlos Silva',
    email: 'lucas@silvacomercio.com.br',
    phone: '+55 11 98765 4321',
    address: 'Av. Paulista 1500, São Paulo 01310-100, Brazil',
    companyName: 'Silva Comércio Ltda',
    country: '巴西',
    product: '不锈钢水壶 750ml',
    productModel: '750ml',
    mainBusiness: '运动户外电商',
    foundedAt: '2009',
    teamSize: '100-200 人',
    annualRevenue: '2000 万雷亚尔',
    followupStatus: '寄样中',
    dealStatus: '未成交',
    coreReason: '样品已寄出，等待客户测试反馈后再谈价格',
    followups: [
      { date: '2026-11-16', note: 'DHL 单号已发客户，预计 4 天到圣保罗' }
    ]
  }),
  makeRecord({
    firstInquiryAt: '2026-11-20',
    customerName: 'Emily Chen',
    companyName: 'Chen Home Supply LLC',
    country: '美国',
    product: '玻璃保鲜盒 5 件套',
    foundedAt: '',
    teamSize: '',
    annualRevenue: '',
    followupStatus: '新询盘',
    dealStatus: '未成交',
    coreReason: '刚接触，需求数量与目标价待确认',
    followups: [] /* 刻意留空，验证空跟进列 */
  }),
  makeRecord({
    firstInquiryAt: '2026-06-01',
    customerName: 'Rajesh Kumar',
    email: 'r.kumar@kumarexports.in',
    phone: '+91 22 4000 1234',
    address: 'Plot 12, MIDC Andheri East, Mumbai 400093, India',
    companyName: 'Kumar Exports Pvt. Ltd',
    country: '印度',
    product: '不锈钢运动水壶 1L',
    productModel: '1L',
    mainBusiness: '酒店用品进口',
    foundedAt: '2007',
    teamSize: '200-500 人',
    annualRevenue: '8000 万美元',
    followupStatus: '已流失',
    dealStatus: '未成交',
    coreReason: '目标价低于我方成本 12%，首单仅 300 套无法开模',
    followups: [
      { date: '2026-09-20', note: '客户最终从越南供应商下单，标记流失' },
      { date: '2026-08-15', note: '二次报价仍低于成本，未达成' }
    ]
  }),
  makeRecord({
    firstInquiryAt: '2026-10-12',
    customerName: 'Lars Nielsen',
    email: 'lars@nielsennordic.dk',
    phone: '+45 33 12 34 56',
    address: 'Bredgade 24, 1260 København K, Denmark',
    companyName: 'Nielsen Nordic ApS',
    country: '丹麦',
    product: '保温杯 OEM（客户品牌）',
    productModel: '定制款',
    mainBusiness: '自有品牌 OEM 采购',
    foundedAt: '1994',
    teamSize: '50 人',
    annualRevenue: '1500 万欧元',
    followupStatus: '暂停跟进',
    dealStatus: '未成交',
    coreReason: '客户内部调整预算，项目推迟到 2027 年',
    followups: [
      { date: '2026-11-01', note: '客户确认项目暂停，2027 年 Q1 再联系' }
    ]
  })
];

const TOTAL_HEADERS = [
  '序号', '首次询盘时间', '客户名称', '邮箱', '电话', '地址', '咨询产品', '产品型号', '国家', '公司名', '主营业务',
  '成立时间', '团队规模', '年营收', '跟进状态', '是否成交', '核心原因', '距上次跟进(天)',
  '最近跟进①日期', '最近跟进①内容', '最近跟进②日期', '最近跟进②内容', '最近跟进③日期', '最近跟进③内容'
];

const DETAIL_HEADERS = [
  '序号', '首次询盘时间', '客户名称', '邮箱', '电话', '公司名', '主营业务', '国家',
  '跟进日期', '距今天数', '跟进内容', '跟进状态', '是否成交'
];

const TOTAL_FOLLOWUPS = 15; // 3+3+2+3+1+0+2+1

/* ================================================================== */
/* 3. 场景 1：8 条示例客户                                             */
/* ================================================================== */

runCase('场景 1｜8 条示例客户：客户总表 24 列 + 跟进明细 13 列', () => {
  const result = xlsx.exportCustomers(SAMPLE_RECORDS, { today: TODAY });
  ok('返回 { filename, bytes }', !!(result && typeof result.filename === 'string' && result.bytes));
  ok('filename 为 客户跟进台账_YYYYMMDD_HHmm.xlsx',
    /^客户跟进台账_\d{8}_\d{4}\.xlsx$/.test(result.filename), result.filename);
  ok('bytes.length > 3000', result.bytes.length > 3000, result.bytes.length + ' 字节');
  expectPkHeader('8 条示例', result.bytes);

  fs.writeFileSync(MAIN_XLSX, result.bytes);
  artifactSizes['crm-export-test.xlsx'] = result.bytes.length;

  const info = readWorkbookOrFail(MAIN_XLSX, '8 条示例');
  if (!info) { return; }

  eq('ZIP 完整性：zipfile.testzip() === None', info.testzip, null);
  const requiredParts = [
    '[Content_Types].xml', '_rels/.rels', 'xl/workbook.xml', 'xl/_rels/workbook.xml.rels',
    'xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml', 'xl/styles.xml'
  ];
  requiredParts.forEach((part) => {
    ok('ZIP 包含部件 ' + part, Array.isArray(info.entries) && info.entries.indexOf(part) !== -1);
  });

  eq('工作表名与顺序', info.sheetnames, ['客户总表', '跟进明细']);
  eq('文档属性标题', info.properties && info.properties.title, '客户跟进台账');

  const s1 = sheetOf(info, '客户总表');
  const s2 = sheetOf(info, '跟进明细');

  /* ---- 客户总表 ---- */
  eq('客户总表：表头 24 列文本完全一致', s1.headers, TOTAL_HEADERS);
  eq('客户总表：列数 24', s1.max_col, 24);
  eq('客户总表：行数 = 记录数 + 1', s1.max_row, SAMPLE_RECORDS.length + 1);
  eq('客户总表：冻结窗格 A2', s1.freeze_panes, 'A2');
  eq('客户总表：autoFilter 覆盖数据区 A1:X9', s1.auto_filter, 'A1:X9');
  eq('客户总表：表头行高 24', s1.header_height, 24);
  ok('客户总表：表头加粗', s1.header_bold === true);
  ok('客户总表：表头白字 FFFFFFFF', String(s1.header_font_color || '').toUpperCase().endsWith('FFFFFFFF'),
    String(s1.header_font_color));
  ok('客户总表：表头深蓝底 #0B2239', String(s1.header_fill || '').toUpperCase().endsWith('0B2239'),
    String(s1.header_fill));
  eq('客户总表：表头水平居中', s1.header_align, 'center');
  eq('客户总表：表头垂直居中', s1.header_vertical, 'center');

  eq('列宽：序号 A = 8', s1.widths && s1.widths.A, 8);
  eq('列宽：客户名称 C = 20', s1.widths && s1.widths.C, 20);
  eq('列宽：邮箱 D = 26', s1.widths && s1.widths.D, 26);
  eq('列宽：电话 E = 20', s1.widths && s1.widths.E, 20);
  eq('列宽：地址 F = 34', s1.widths && s1.widths.F, 34);
  eq('列宽：产品型号 H = 16', s1.widths && s1.widths.H, 16);
  eq('列宽：主营业务 K = 22', s1.widths && s1.widths.K, 22);
  eq('列宽：核心原因 Q = 46', s1.widths && s1.widths.Q, 46);
  eq('列宽：24 列全部自定义', Object.keys(s1.widths || {}).length, 24);

  eq('正文：核心原因列 wrapText', s1.body_wrap && s1.body_wrap[16], true);
  eq('正文：最近跟进①内容列 wrapText', s1.body_wrap && s1.body_wrap[19], true);
  eq('正文：客户名称列不换行', s1.body_wrap && s1.body_wrap[2], false);
  eq('正文：邮箱列不换行', s1.body_wrap && s1.body_wrap[3], false);
  eq('正文：地址列不换行', s1.body_wrap && s1.body_wrap[5], false);
  eq('数字列：距上次跟进(天) 为数字单元格', s1.cell_types && s1.cell_types[0] && s1.cell_types[0][17], 'n');
  eq('文本列：客户名称为字符串单元格', s1.cell_types && s1.cell_types[0] && s1.cell_types[0][2], 's');
  eq('文本列：邮箱/电话为字符串单元格', [s1.cell_types && s1.cell_types[0] && s1.cell_types[0][3],
    s1.cell_types && s1.cell_types[0] && s1.cell_types[0][4]], ['s', 's']);
  eq('文本列：地址为字符串单元格', s1.cell_types && s1.cell_types[0] && s1.cell_types[0][5], 's');
  eq('数字格式：距上次跟进(天) numFmt = 0', s1.number_formats && s1.number_formats[0] && s1.number_formats[0][17], '0');

  eq('第 1 行内容（24 列全字段）', s1.rows && s1.rows[0], [
    1,
    '2026-09-12',
    'Michael Braun',
    'm.braun@braun-trading.de',
    '+49 711 4567 890',
    'Industriestr. 12, 70565 Stuttgart, Germany',
    '不锈钢保温杯 500ml（定制 LOGO）',
    '500ml',
    '德国',
    'Braun Trading GmbH',
    '户外用品进口分销',
    '1998',
    '50-100 人',
    '1200 万欧元',
    '已报价',
    '未成交',
    '报价高出目标价约 8%，客户等终端零售客户确认后回复',
    3,
    '2026-11-18',
    '邮件确认样品尺寸与 LOGO 印刷位置，等寄样确认',
    '2026-11-10',
    '电话沟通，客户确认 2000pcs 试单意向',
    '2026-10-28',
    '首次邮件回复，发送价目表与产品册'
  ]);
  eq('新列取值：邮箱/电话/地址（第 2 条记录）',
    [s1.rows && s1.rows[1][3], s1.rows && s1.rows[1][4], s1.rows && s1.rows[1][5]],
    ['a.hassan@elnile-import.com', '+20 100 234 5678', '25 Nile Corniche, Cairo 11511, Egypt']);
  eq('新列取值：地址含中文（第 4 条记录）', [s1.rows && s1.rows[3][3], s1.rows && s1.rows[3][5]],
    ['y.tanaka@tanaka-trading.jp', '2-1-1 Uchisaiwaicho, Chiyoda-ku, Tokyo 100-0011, Japan']);
  eq('v1.1 列保持正确：产品型号/主营业务（第 2 条记录）',
    [s1.rows && s1.rows[1][7], s1.rows && s1.rows[1][10]], ['16 件套', '家居餐具进口分销']);
  eq('v1.1 列保持正确：第 4 条记录 产品型号/主营业务',
    [s1.rows && s1.rows[3][7], s1.rows && s1.rows[3][10]], ['1.2L', '日用百货批发']);
  eq('无跟进/缺联系方式客户：新列均为空单元格（旧备份缺字段同口径）',
    [s1.rows && s1.rows[5][3], s1.rows && s1.rows[5][4], s1.rows && s1.rows[5][5],
      s1.rows && s1.rows[5][7], s1.rows && s1.rows[5][10]],
    [null, null, null, null, null]);

  eq('跟进记录超过 3 条时截断为最近 3 条并按日期倒序', s1.rows && s1.rows[3].slice(18), [
    '2026-11-15', '邮件确认交期改善方案，客户内部评审',
    '2026-11-05', '发送 45 天交期方案与补偿条款',
    '2026-10-22', '视频会议讨论 JIS 认证与包装'
  ]);
  ok('最早第 4 条跟进（2026-10-01）未导出',
    !(s1.rows && s1.rows[3].slice(18).indexOf('2026-10-01') !== -1));

  eq('无跟进客户：距上次跟进及最近跟进三列为空', s1.rows && s1.rows[5].slice(17),
    [null, null, null, null, null, null, null]);
  eq('无跟进客户：跟进状态/成交状态正常',
    [s1.rows && s1.rows[5][2], s1.rows && s1.rows[5][14], s1.rows && s1.rows[5][15]],
    ['Emily Chen', '新询盘', '未成交']);
  eq('距上次跟进天数：第 2 条记录 = 1', s1.rows && s1.rows[1][17], 1);
  eq('距上次跟进天数：第 7 条记录 = 62', s1.rows && s1.rows[6][17], 62);
  eq('距上次跟进天数：第 8 条记录 = 20', s1.rows && s1.rows[7][17], 20);

  /* ---- 跟进明细 ---- */
  eq('跟进明细：表头 13 列文本完全一致', s2.headers, DETAIL_HEADERS);
  eq('跟进明细：列数 13', s2.max_col, 13);
  eq('跟进明细：行数 = 跟进总数 + 1', s2.max_row, TOTAL_FOLLOWUPS + 1);
  eq('跟进明细：冻结窗格 A2', s2.freeze_panes, 'A2');
  eq('跟进明细：autoFilter A1:M16', s2.auto_filter, 'A1:M16');
  eq('跟进明细：表头行高 24', s2.header_height, 24);
  eq('跟进明细：序号连续 1..15', (s2.rows || []).map((row) => row[0]),
    Array.from({ length: TOTAL_FOLLOWUPS }, (unused, index) => index + 1));
  eq('跟进明细：跟进内容列 wrapText', s2.body_wrap && s2.body_wrap[10], true);
  eq('跟进明细：主营业务列不换行', s2.body_wrap && s2.body_wrap[6], false);
  eq('跟进明细：邮箱列不换行', s2.body_wrap && s2.body_wrap[3], false);
  eq('跟进明细：列宽 K = 60', s2.widths && s2.widths.K, 60);
  eq('跟进明细：列宽 D（邮箱）= 26 / E（电话）= 20 / G（主营业务）= 22',
    [s2.widths && s2.widths.D, s2.widths && s2.widths.E, s2.widths && s2.widths.G], [26, 20, 22]);
  eq('跟进明细：13 列全部自定义', Object.keys(s2.widths || {}).length, 13);
  eq('跟进明细：第 1 行 = 迈克尔最近一条跟进', s2.rows && s2.rows[0], [
    1, '2026-09-12', 'Michael Braun', 'm.braun@braun-trading.de', '+49 711 4567 890',
    'Braun Trading GmbH', '户外用品进口分销', '德国',
    '2026-11-18', 3, '邮件确认样品尺寸与 LOGO 印刷位置，等寄样确认', '已报价', '未成交'
  ]);
  eq('跟进明细：第 4 行 = 艾哈迈德最近一条跟进', s2.rows && s2.rows[3], [
    4, '2026-10-05', 'Ahmed Hassan', 'a.hassan@elnile-import.com', '+20 100 234 5678',
    'El Nile Import & Export Co., Ltd', '家居餐具进口分销', '埃及',
    '2026-11-20', 1, 'WhatsApp 回复，确认 300 万美元年采购额并索取 PI', '已回复', '未成交'
  ]);
  eq('跟进明细：主营/邮箱/电话列取值（第 4 行）',
    [s2.rows && s2.rows[3][3], s2.rows && s2.rows[3][4], s2.rows && s2.rows[3][6]],
    ['a.hassan@elnile-import.com', '+20 100 234 5678', '家居餐具进口分销']);
  eq('跟进明细：每个客户一行/条且按客户分组', (s2.rows || []).map((row) => row[2]), [
    'Michael Braun', 'Michael Braun', 'Michael Braun',
    'Ahmed Hassan', 'Ahmed Hassan', 'Ahmed Hassan',
    'Sophie Martin', 'Sophie Martin',
    'Yuki Tanaka', 'Yuki Tanaka', 'Yuki Tanaka',
    'Carlos Silva',
    'Rajesh Kumar', 'Rajesh Kumar',
    'Lars Nielsen'
  ]);
  eq('跟进明细：距今天数正确（第 12 行 = 62）', s2.rows && s2.rows[12][9], 62);
  eq('跟进明细：无跟进客户不产生明细行（共 15 行）', s2.max_row, 16);
});

/* ================================================================== */
/* 4. 场景 2：空记录                                                   */
/* ================================================================== */

runCase('场景 2｜空记录：空数组 + 默认空记录', () => {
  /* 2.1 完全空的记录数组：只应剩表头 */
  const empty = xlsx.exportCustomers([], { today: TODAY });
  ok('空数组：bytes 非空', empty.bytes.length > 1000, empty.bytes.length + ' 字节');
  expectPkHeader('空数组', empty.bytes);
  const emptyFile = writeXlsx('crm-export-empty.xlsx', empty.bytes);
  const info = readWorkbookOrFail(emptyFile, '空数组');
  if (info) {
    eq('空数组：ZIP 完整', info.testzip, null);
    eq('空数组：工作表名', info.sheetnames, ['客户总表', '跟进明细']);
    const s1 = sheetOf(info, '客户总表');
    const s2 = sheetOf(info, '跟进明细');
    eq('空数组：客户总表仅表头（1 行）', s1.max_row, 1);
    eq('空数组：客户总表表头 24 列', s1.headers, TOTAL_HEADERS);
    eq('空数组：客户总表冻结 A2', s1.freeze_panes, 'A2');
    eq('空数组：客户总表 autoFilter 覆盖表头行', s1.auto_filter, 'A1:X1');
    eq('空数组：跟进明细仅表头（1 行）', s2.max_row, 1);
    eq('空数组：跟进明细表头 13 列', s2.headers, DETAIL_HEADERS);
    eq('空数组：跟进明细 autoFilter', s2.auto_filter, 'A1:M1');
  } else {
    return;
  }

  /* 2.2 默认空记录（11 个字段全默认值）：应有 1 行 */
  const blank = schema.emptyRecord(TODAY);
  const blankResult = xlsx.exportCustomers([blank], { today: TODAY });
  const blankFile = writeXlsx('crm-export-blank-record.xlsx', blankResult.bytes);
  const blankInfo = readWorkbookOrFail(blankFile, '默认空记录');
  if (!blankInfo) { return; }
  eq('默认空记录：ZIP 完整', blankInfo.testzip, null);
  const s1 = sheetOf(blankInfo, '客户总表');
  eq('默认空记录：客户总表 2 行（表头 + 1 行）', s1.max_row, 2);
  eq('默认空记录：序号 = 1', s1.rows && s1.rows[0][0], 1);
  eq('默认空记录：首次询盘时间 = today', s1.rows && s1.rows[0][1], TODAY);
  eq('默认空记录：客户名称等文本为空单元格',
    [s1.rows && s1.rows[0][2], s1.rows && s1.rows[0][3], s1.rows && s1.rows[0][4],
      s1.rows && s1.rows[0][5], s1.rows && s1.rows[0][6], s1.rows && s1.rows[0][7],
      s1.rows && s1.rows[0][9], s1.rows && s1.rows[0][10], s1.rows && s1.rows[0][16]],
    [null, null, null, null, null, null, null, null, null]);
  eq('默认空记录：跟进状态/成交状态默认枚举', [s1.rows && s1.rows[0][14], s1.rows && s1.rows[0][15]],
    ['新询盘', '未成交']);
  eq('默认空记录：距上次跟进与最近跟进列为空', s1.rows && s1.rows[0].slice(17),
    [null, null, null, null, null, null, null]);
  eq('默认空记录：跟进明细仅表头', sheetOf(blankInfo, '跟进明细').max_row, 1);
});

/* ================================================================== */
/* 5. 场景 3：超长中文（500 字）                                        */
/* ================================================================== */

runCase('场景 3｜超长中文 500 字：客户总表核心原因列', () => {
  const longText = '超长中文'.repeat(125); // 4 × 125 = 500
  eq('构造 500 字中文串', longText.length, 500);
  const record = makeRecord({
    firstInquiryAt: '2026-10-01',
    customerName: '超长文本客户',
    companyName: '超长文本测试有限公司',
    country: '德国',
    product: '保温杯',
    followupStatus: '已回复',
    dealStatus: '未成交',
    coreReason: longText,
    followups: [{ date: '2026-11-18', note: '常规跟进，客户已收到样册' }]
  });

  const result = xlsx.exportCustomers([record], { today: TODAY });
  expectPkHeader('500 字中文', result.bytes);
  const file = writeXlsx('crm-export-long-text.xlsx', result.bytes);
  const info = readWorkbookOrFail(file, '500 字中文');
  if (!info) { return; }

  eq('500 字场景：ZIP 完整', info.testzip, null);
  const s1 = sheetOf(info, '客户总表');
  eq('500 字场景：行数 2', s1.max_row, 2);
  eq('500 字场景：核心原因 500 字完整保留', s1.rows && s1.rows[0][16], longText);
  eq('500 字场景：核心原因长度 = 500', s1.rows && s1.rows[0][16] && s1.rows[0][16].length, 500);
  eq('500 字场景：中文值未被破坏', s1.rows && s1.rows[0][2], '超长文本客户');
  eq('500 字场景：核心原因列 wrapText', s1.body_wrap && s1.body_wrap[16], true);
  eq('500 字场景：跟进明细正常', sheetOf(info, '跟进明细').max_row, 2);
  eq('500 字场景：跟进内容正常', sheetOf(info, '跟进明细').rows && sheetOf(info, '跟进明细').rows[0][10],
    '常规跟进，客户已收到样册');
});

/* ================================================================== */
/* 6. 场景 4：特殊字符 + 换行 + 控制字符                                */
/* ================================================================== */

runCase('场景 4｜特殊字符 <&"\'>、换行、制表符与控制字符', () => {
  const SPECIAL_NAME = "A<&\"'>B\nC <tag> & \"quote\" 'single'";
  const SPECIAL_EMAIL = 'A<&"\'>B@example.com';
  const SPECIAL_PHONE = '+49 <&"\'> 123\n分机 9';
  const SPECIAL_ADDRESS = '地址 <&"\'> 第一行\n第二行 & "quote"';
  const SPECIAL_PRODUCT = 'line1\nline2\r\nline3';
  const SPECIAL_MODEL = '型号 <&"\' >';
  const SPECIAL_BUSINESS = '主营 <&"\' >';
  const SPECIAL_REASON = '控制\u0000字符\u0007测试\u000B制表\t符';
  const SPECIAL_NOTE = '第一行\n第二行 <&"\'> 结束';

  const record = makeRecord({
    firstInquiryAt: '2026-11-01',
    customerName: SPECIAL_NAME,
    email: SPECIAL_EMAIL,
    phone: SPECIAL_PHONE,
    address: SPECIAL_ADDRESS,
    companyName: 'Special <Chars> & "Quotes" Co.',
    country: '德国',
    product: SPECIAL_PRODUCT,
    productModel: SPECIAL_MODEL,
    mainBusiness: SPECIAL_BUSINESS,
    followupStatus: '已回复',
    dealStatus: '未成交',
    coreReason: SPECIAL_REASON,
    followups: [{ date: '2026-11-18', note: SPECIAL_NOTE }]
  });

  const result = xlsx.exportCustomers([record], { today: TODAY });
  expectPkHeader('特殊字符', result.bytes);
  const file = writeXlsx('crm-export-special-chars.xlsx', result.bytes);
  const info = readWorkbookOrFail(file, '特殊字符');
  if (!info) { return; }

  eq('特殊字符：ZIP 完整（XML 未损坏）', info.testzip, null);
  const s1 = sheetOf(info, '客户总表');
  const s2 = sheetOf(info, '跟进明细');
  eq('特殊字符：客户名称原样读回（<&"\'> 与换行）', s1.rows && s1.rows[0][2], SPECIAL_NAME);
  eq('特殊字符：邮箱特殊字符原样读回', s1.rows && s1.rows[0][3], SPECIAL_EMAIL);
  eq('特殊字符：电话换行归一 + 特殊字符原样读回', s1.rows && s1.rows[0][4], SPECIAL_PHONE);
  eq('特殊字符：地址换行与特殊字符原样读回', s1.rows && s1.rows[0][5], SPECIAL_ADDRESS);
  eq('特殊字符：\\r\\n 归一为 \\n（咨询产品）', s1.rows && s1.rows[0][6], 'line1\nline2\nline3');
  eq('特殊字符：产品型号原样读回', s1.rows && s1.rows[0][7], SPECIAL_MODEL);
  eq('特殊字符：公司名特殊字符原样读回', s1.rows && s1.rows[0][9], 'Special <Chars> & "Quotes" Co.');
  eq('特殊字符：主营业务原样读回', s1.rows && s1.rows[0][10], SPECIAL_BUSINESS);
  eq('特殊字符：XML 非法控制字符被剔除、制表符保留', s1.rows && s1.rows[0][16], '控制字符测试制表\t符');
  eq('特殊字符：最近跟进内容中的换行与特殊字符原样保留', s1.rows && s1.rows[0][19], SPECIAL_NOTE);
  eq('特殊字符：跟进明细内容读回一致', s2.rows && s2.rows[0][10], SPECIAL_NOTE);
  eq('特殊字符：跟进明细邮箱/电话读回一致', [s2.rows && s2.rows[0][3], s2.rows && s2.rows[0][4]],
    [SPECIAL_EMAIL, SPECIAL_PHONE]);
});

/* ================================================================== */
/* 7. 场景 5：入参兼容（原始记录 / 已派生记录）                         */
/* ================================================================== */

runCase('场景 5｜入参兼容：原始记录 / 已派生记录（last3 + daysSince）都能导出', () => {
  /* 5.1 已派生但没有 followups 的记录：沿用 last3 / daysSince，不丢信息 */
  const derivedRecord = makeRecord({
    firstInquiryAt: '2026-10-01',
    customerName: '派生字段客户',
    email: 'derived@example.com',
    phone: '+49 30 1234 5678',
    address: 'Derived Str. 1, 10115 Berlin, Germany',
    companyName: 'Derived Trading Ltd',
    country: '德国',
    product: '不锈钢保温杯',
    productModel: '500ml',
    mainBusiness: '户外用品进口分销',
    followupStatus: '已报价',
    dealStatus: '未成交'
  });
  delete derivedRecord.followups;
  derivedRecord.last3 = [
    { date: '2026-11-18', note: '派生：最近一次跟进', daysAgo: 3 },
    { date: '2026-11-10', note: '派生：第二次跟进', daysAgo: 11 },
    { date: '2026-10-28', note: '派生：第一次跟进', daysAgo: 24 }
  ];
  derivedRecord.daysSince = 3;
  derivedRecord.ageTone = 'ok';

  const derivedResult = xlsx.exportCustomers([derivedRecord], { today: TODAY });
  const derivedFile = writeXlsx('crm-export-derived-record.xlsx', derivedResult.bytes);
  const derivedInfo = readWorkbookOrFail(derivedFile, '已派生记录');
  if (derivedInfo) {
    const s1 = sheetOf(derivedInfo, '客户总表');
    const s2 = sheetOf(derivedInfo, '跟进明细');
    eq('已派生记录：距上次跟进(天) = 3', s1.rows && s1.rows[0][17], 3);
    eq('已派生记录：联系方式/产品型号/主营业务正常导出',
      [s1.rows && s1.rows[0][3], s1.rows && s1.rows[0][4], s1.rows && s1.rows[0][5],
        s1.rows && s1.rows[0][7], s1.rows && s1.rows[0][10]],
      ['derived@example.com', '+49 30 1234 5678', 'Derived Str. 1, 10115 Berlin, Germany',
        '500ml', '户外用品进口分销']);
    eq('已派生记录：最近跟进①②③ 完整导出', s1.rows && s1.rows[0].slice(18), [
      '2026-11-18', '派生：最近一次跟进',
      '2026-11-10', '派生：第二次跟进',
      '2026-10-28', '派生：第一次跟进'
    ]);
    eq('已派生记录：跟进明细 4 行（表头 + 3）', s2.max_row, 4);
    eq('已派生记录：明细距今天数', (s2.rows || []).map((row) => row[9]), [3, 11, 24]);
    eq('已派生记录：明细主营业务列', (s2.rows || []).map((row) => row[6]),
      ['户外用品进口分销', '户外用品进口分销', '户外用品进口分销']);
    eq('已派生记录：明细邮箱列', (s2.rows || []).map((row) => row[3]),
      ['derived@example.com', 'derived@example.com', 'derived@example.com']);
  }

  /* 5.2 同时带 followups 和派生字段（CrmStore.deriveAll() 的真实形态）：以 followups 为准 */
  const hybridRecord = makeRecord({
    firstInquiryAt: '2026-09-12',
    customerName: '混合形态客户',
    companyName: 'Hybrid GmbH',
    country: '德国',
    followups: [{ date: '2026-11-01', note: '真实跟进记录' }]
  });
  hybridRecord.last3 = [{ date: '2020-01-01', note: '过期派生值', daysAgo: 999 }];
  hybridRecord.daysSince = 999;

  const hybridResult = xlsx.exportCustomers([hybridRecord], { today: TODAY });
  const hybridFile = writeXlsx('crm-export-hybrid-record.xlsx', hybridResult.bytes);
  const hybridInfo = readWorkbookOrFail(hybridFile, '混合形态记录');
  if (hybridInfo) {
    const h1 = sheetOf(hybridInfo, '客户总表');
    eq('混合形态记录：以 followups 重算 daysSince', h1.rows && h1.rows[0][17], 20);
    eq('混合形态记录：最近跟进①日期取自 followups', h1.rows && h1.rows[0][18], '2026-11-01');
    eq('混合形态记录：最近跟进①内容取自 followups', h1.rows && h1.rows[0][19], '真实跟进记录');
  }

  /* 5.3 旧备份记录：完全没有 email / phone / address / productModel / mainBusiness → 新列为空单元格 */
  const legacyRecord = makeRecord({
    firstInquiryAt: '2026-09-12',
    customerName: '旧备份客户',
    companyName: 'Legacy Trading Ltd',
    country: '德国',
    product: '保温杯',
    followups: [{ date: '2026-11-01', note: '旧备份跟进' }]
  });
  delete legacyRecord.email;
  delete legacyRecord.phone;
  delete legacyRecord.address;
  delete legacyRecord.productModel;
  delete legacyRecord.mainBusiness;

  const legacyResult = xlsx.exportCustomers([legacyRecord], { today: TODAY });
  const legacyFile = writeXlsx('crm-export-legacy-record.xlsx', legacyResult.bytes);
  const legacyInfo = readWorkbookOrFail(legacyFile, '旧备份记录');
  if (legacyInfo) {
    const l1 = sheetOf(legacyInfo, '客户总表');
    const l2 = sheetOf(legacyInfo, '跟进明细');
    eq('旧备份记录（缺 v1.1/v1.2 新字段）：邮箱/电话/地址/产品型号/主营业务为空单元格',
      [l1.rows && l1.rows[0][3], l1.rows && l1.rows[0][4], l1.rows && l1.rows[0][5],
        l1.rows && l1.rows[0][7], l1.rows && l1.rows[0][10]],
      [null, null, null, null, null]);
    eq('旧备份记录：原有字段照常导出', [l1.rows && l1.rows[0][2], l1.rows && l1.rows[0][9]],
      ['旧备份客户', 'Legacy Trading Ltd']);
    eq('旧备份记录：明细邮箱/主营业务为空单元格',
      [l2.rows && l2.rows[0][3], l2.rows && l2.rows[0][6]], [null, null]);
  }
});

/* ================================================================== */
/* 8. 场景 6：buildWorkbook 底层接口 + download 的 Node 行为            */
/* ================================================================== */

runCase('场景 6｜buildWorkbook 底层接口 / download 的 Node 行为', () => {
  eq('VERSION = 1.0.0', xlsx.VERSION, '1.0.0');

  const built = xlsx.buildWorkbook([{
    name: '测试/表 1',
    columns: [
      { header: '序号', width: 8, align: 'center' },
      { header: '名称', width: 24 },
      { header: '金额', width: 14, align: 'right', numFmt: '0' }
    ],
    rows: [[1, '甲', 100], [2, '乙', 200]],
    freeze: { rows: 1 },
    autoFilter: true
  }]);
  ok('buildWorkbook 返回 Uint8Array', built instanceof Uint8Array);
  expectPkHeader('buildWorkbook', built);

  const file = writeXlsx('crm-export-build-workbook.xlsx', built);
  const info = readWorkbookOrFail(file, 'buildWorkbook');
  if (info) {
    eq('buildWorkbook：ZIP 完整', info.testzip, null);
    eq('buildWorkbook：非法工作表名被清洗为「测试 表 1」', info.sheetnames, ['测试 表 1']);
    const sheet = sheetOf(info, '测试 表 1');
    eq('buildWorkbook：max_row = 数据行 + 表头', sheet.max_row, 3);
    eq('buildWorkbook：表头', sheet.headers, ['序号', '名称', '金额']);
    eq('buildWorkbook：autoFilter A1:C3', sheet.auto_filter, 'A1:C3');
    eq('buildWorkbook：冻结 A2', sheet.freeze_panes, 'A2');
    eq('buildWorkbook：列宽 A=8 / B=24 / C=14', [sheet.widths.A, sheet.widths.B, sheet.widths.C], [8, 24, 14]);
    eq('buildWorkbook：数据行', sheet.rows, [[1, '甲', 100], [2, '乙', 200]]);
    eq('buildWorkbook：数字单元格 100 为数字', sheet.rows && sheet.rows[0][2], 100);
    eq('buildWorkbook：金额列 numFmt = 0', sheet.number_formats && sheet.number_formats[0][2], '0');
  }

  /* 无参数时也应产出可打开的默认工作簿 */
  const fallbackBytes = xlsx.buildWorkbook();
  ok('buildWorkbook() 无参数仍返回有效 ZIP', fallbackBytes instanceof Uint8Array && fallbackBytes.length > 500);

  /* Node 下 download 必须抛出明确错误（而不是静默失败或引用 document） */
  let downloadError = null;
  try {
    xlsx.download(built, 'demo.xlsx');
  } catch (err) {
    downloadError = err;
  }
  ok('Node 下 download() 抛明确错误', !!downloadError && /浏览器/.test(String(downloadError.message)),
    downloadError ? downloadError.message : '未抛错');
});

/* ================================================================== */
/* 9. 场景 7：性能（500 行 × 24 列 < 300ms）                            */
/* ================================================================== */

runCase('场景 7｜性能：500 行 × 24 列 < 300ms', () => {
  /* 8.1 契约口径：单个 500 行 × 24 列工作表 */
  const columns = Array.from({ length: 24 }, (unused, index) => ({
    header: '列' + (index + 1),
    width: index === 0 ? 8 : 14
  }));
  const rows = Array.from({ length: 500 }, (unused, rowIndex) =>
    Array.from({ length: 24 }, (ignored, colIndex) =>
      colIndex === 0 ? rowIndex + 1 : '值' + rowIndex + '-' + colIndex));
  const workSheet = { name: '压测', columns: columns, rows: rows, freeze: { rows: 1 }, autoFilter: true };

  const timings = [];
  let stressBytes = null;
  for (let run = 0; run < 3; run += 1) {
    const startedAt = performance.now();
    stressBytes = xlsx.buildWorkbook([workSheet]);
    timings.push(performance.now() - startedAt);
  }
  const best = Math.min.apply(null, timings);
  console.log('    耗时（500×24，三次）：' + timings.map((value) => value.toFixed(1)).join(' / ') + ' ms');
  ok('buildWorkbook 500 行 × 24 列 < 300ms（三次取最优）', best < 300,
    '耗时 ' + timings.map((value) => value.toFixed(1)).join(' / ') + ' ms，最优 ' + best.toFixed(1) + 'ms');
  ok('压测工作簿体积正常（> 200KB）', stressBytes.length > 200000, stressBytes.length + ' 字节');

  /* 8.2 完整导出（客户总表 500×24 + 跟进明细 1500×13） */
  const bigRecords = Array.from({ length: 500 }, (unused, index) => makeRecord({
    firstInquiryAt: '2026-09-12',
    customerName: '客户 ' + (index + 1),
    email: 'customer' + (index + 1) + '@example.com',
    phone: '+49 30 1234 ' + String(index + 1).padStart(4, '0'),
    address: '示例路 ' + (index + 1) + ' 号, 10115 Berlin, Germany',
    companyName: '示例进出口贸易有限公司 ' + (index + 1),
    country: index % 2 === 0 ? '德国' : '美国',
    product: '不锈钢保温杯 500ml / 型号 ' + (index + 1),
    productModel: '型号-' + (index + 1),
    mainBusiness: '户外用品进口分销（压测）',
    foundedAt: '2000',
    teamSize: '50-100 人',
    annualRevenue: '1200 万欧元',
    followupStatus: '已报价',
    dealStatus: '未成交',
    coreReason: '价格高出目标价 8%，等待终端客户确认；这是用于压测的核心原因文本。',
    followups: [
      { date: '2026-11-18', note: '邮件确认样品尺寸与 LOGO 印刷位置，等寄样确认（压测）' },
      { date: '2026-11-10', note: '电话沟通，确认 2000pcs 试单意向（压测）' },
      { date: '2026-10-28', note: '首次邮件回复，发送价目表与产品册（压测）' }
    ]
  }));

  xlsx.exportCustomers(bigRecords.slice(0, 20), { today: TODAY }); // 预热

  const exportTimings = [];
  let bigResult = null;
  for (let run = 0; run < 3; run += 1) {
    const startedAt = performance.now();
    bigResult = xlsx.exportCustomers(bigRecords, { today: TODAY });
    exportTimings.push(performance.now() - startedAt);
  }
  const bestExport = Math.min.apply(null, exportTimings);
  console.log('    耗时（500 客户 / 1500 条跟进，三次）：' + exportTimings.map((value) => value.toFixed(1)).join(' / ') + ' ms');
  ok('exportCustomers 500 条客户（含 1500 行跟进明细）< 800ms', bestExport < 800,
    '耗时 ' + exportTimings.map((value) => value.toFixed(1)).join(' / ') + ' ms，最优 ' + bestExport.toFixed(1) + 'ms');
  ok('大数据导出体积 > 100KB', bigResult.bytes.length > 100000, bigResult.bytes.length + ' 字节');

  /* 8.3 大数据文件同样要能被读回（ZIP 偏移 / XML 结构在 1MB 级别仍正确） */
  const bigFile = writeXlsx('crm-export-big-500.xlsx', bigResult.bytes);
  const bigInfo = readWorkbookOrFail(bigFile, '500 条大数据');
  if (bigInfo) {
    const bigSheet1 = sheetOf(bigInfo, '客户总表');
    const bigSheet2 = sheetOf(bigInfo, '跟进明细');
    eq('大数据：ZIP 完整', bigInfo.testzip, null);
    eq('大数据：客户总表 501 行 × 24 列', [bigSheet1.max_row, bigSheet1.max_col], [501, 24]);
    eq('大数据：客户总表 autoFilter 覆盖 A1:X501', bigSheet1.auto_filter, 'A1:X501');
    eq('大数据：跟进明细 1501 行 × 13 列', [bigSheet2.max_row, bigSheet2.max_col], [1501, 13]);
    eq('大数据：跟进明细 autoFilter 覆盖 A1:M1501', bigSheet2.auto_filter, 'A1:M1501');
    eq('大数据：最后一条客户名正确', bigSheet1.rows && bigSheet1.rows[499][2], '客户 500');
    eq('大数据：最后一条明细序号 = 1500', bigSheet2.rows && bigSheet2.rows[1499][0], 1500);
    eq('大数据：最后一条客户的新列（邮箱/电话/地址/产品型号/主营业务）',
      [bigSheet1.rows && bigSheet1.rows[499][3], bigSheet1.rows && bigSheet1.rows[499][4],
        bigSheet1.rows && bigSheet1.rows[499][5], bigSheet1.rows && bigSheet1.rows[499][7],
        bigSheet1.rows && bigSheet1.rows[499][10]],
      ['customer500@example.com', '+49 30 1234 0500', '示例路 500 号, 10115 Berlin, Germany',
        '型号-500', '户外用品进口分销（压测）']);
    eq('大数据：最后一条明细的邮箱/电话',
      [bigSheet2.rows && bigSheet2.rows[1499][3], bigSheet2.rows && bigSheet2.rows[1499][4]],
      ['customer500@example.com', '+49 30 1234 0500']);
  }
});

/* ================================================================== */
/* 10. 汇总                                                            */
/* ================================================================== */

console.log('\n' + '─'.repeat(66));
console.log('产物（%TEMP%）：');
console.log('  主产物      ' + MAIN_XLSX + '（' + (artifactSizes['crm-export-test.xlsx'] || 0) + ' 字节）');
Object.keys(artifactSizes).forEach((name) => {
  if (name === 'crm-export-test.xlsx') { return; }
  console.log('  ' + name + '  ' + artifactSizes[name] + ' 字节');
});
console.log('  校验方式：' + (PYTHON
  ? 'Python + openpyxl 读回 ' + pythonReads + ' 次'
  : '纯 Node ZIP/XML 解析 ' + nodeReads + ' 次（SKIP openpyxl 校验）'));

if (failures.length) {
  console.error('\n✗ FAIL — ' + failures.length + ' 项断言失败 / 共 ' + assertionCount + ' 项（退出码 1）');
  failures.forEach((failure, index) => {
    console.error('  ' + (index + 1) + '. ' + failure.name + (failure.detail ? '：' + failure.detail : ''));
  });
  process.exit(1);
}

console.log('\n✓ ALL PASS — ' + assertionCount + ' 项断言通过；4 个数据场景 + 入参兼容 + buildWorkbook + 性能 + 读回校验全部通过');
process.exit(0);
