/*!
 * 客户通 CustomerCRM — js/xlsx.js  v1.0.0
 * 零依赖 XLSX 导出引擎：自己实现 ZIP（store 方式 + CRC32）与 Office Open XML 工作簿。
 *
 * 对外接口（浏览器 window.CrmXlsx / Node module.exports）：
 *   buildWorkbook(sheets[, options])           -> Uint8Array
 *   exportCustomers(records, opts)             -> { filename, bytes }
 *   download(bytes, filename)                  -> 浏览器触发下载；Node 下抛明确错误
 *
 * 设计要点：
 *   - 不引用 document/window/Blob 于 Node 分支；字符串用 inlineStr，不需要 sharedStrings。
 *   - 日期一律写成文本 YYYY-MM-DD，规避 Excel 日期序列号与本地化坑。
 *   - XML 转义 & < > " '，剔除 XML 非法控制字符，保留换行/制表符。
 *   - 表头：深蓝底 #0B2239 + 白色加粗 + 居中，行高 24；首行冻结 + autoFilter + 自定义列宽。
 */
(function (root, factory) {
  'use strict';
  var api = factory(root);
  if (typeof module === 'object' && module.exports) { module.exports = api; }
  if (root) { root.CrmXlsx = api; }
})(typeof globalThis !== 'undefined' ? globalThis : this, function (root) {
  'use strict';

  var VERSION = '1.0.0';
  var MIME_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

  var NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  var NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  var NS_PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
  var NS_CONTENT = 'http://schemas.openxmlformats.org/package/2006/content-types';
  var NS_CP = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties';
  var NS_DC = 'http://purl.org/dc/elements/1.1/';
  var NS_DCTERMS = 'http://purl.org/dc/terms/';
  var NS_DCMITYPE = 'http://purl.org/dc/dcmitype/';
  var NS_XSI = 'http://www.w3.org/2001/XMLSchema-instance';

  var XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  var MAX_SHEET_NAME = 31;
  var HEADER_FILL_COLOR = 'FF0B2239'; // 深海军蓝（与全站主色一致）

  /* ================================================================== */
  /* 1. UTF-8 / XML 文本工具                                             */
  /* ================================================================== */

  var _encoder;

  function utf8Encode(text) {
    if (_encoder === undefined) {
      _encoder = typeof TextEncoder !== 'undefined' ? new TextEncoder() : null;
    }
    if (_encoder) { return _encoder.encode(String(text)); }
    return utf8Fallback(String(text));
  }

  /** 无 TextEncoder 的兜底实现（老环境 / 测试桩） */
  function utf8Fallback(text) {
    var bytes = [];
    for (var i = 0; i < text.length; i += 1) {
      var code = text.charCodeAt(i);
      if (code >= 0xD800 && code <= 0xDBFF && i + 1 < text.length) {
        var low = text.charCodeAt(i + 1);
        if (low >= 0xDC00 && low <= 0xDFFF) {
          code = 0x10000 + ((code - 0xD800) << 10) + (low - 0xDC00);
          i += 1;
        }
      }
      if (code < 0x80) {
        bytes.push(code);
      } else if (code < 0x800) {
        bytes.push(0xC0 | (code >> 6), 0x80 | (code & 0x3F));
      } else if (code < 0x10000) {
        bytes.push(0xE0 | (code >> 12), 0x80 | ((code >> 6) & 0x3F), 0x80 | (code & 0x3F));
      } else {
        bytes.push(0xF0 | (code >> 18), 0x80 | ((code >> 12) & 0x3F), 0x80 | ((code >> 6) & 0x3F), 0x80 | (code & 0x3F));
      }
    }
    return new Uint8Array(bytes);
  }

  /** 去掉 XML 1.0 非法字符（保留 \t \n），并规整换行、剔除孤立代理项。 */
  function cleanXmlText(value) {
    var text = String(value);
    text = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    text = text.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, '');
    return stripLoneSurrogates(text);
  }

  /** XML 不允许孤立 UTF-16 代理项，直接剔除避免生成非法文档。 */
  function stripLoneSurrogates(text) {
    var out = '';
    for (var i = 0; i < text.length; i += 1) {
      var code = text.charCodeAt(i);
      if (code >= 0xD800 && code <= 0xDBFF) {
        var next = i + 1 < text.length ? text.charCodeAt(i + 1) : 0;
        if (next >= 0xDC00 && next <= 0xDFFF) {
          out += text.charAt(i) + text.charAt(i + 1);
          i += 1;
        }
        /* 孤立高位代理：丢弃 */
      } else if (code >= 0xDC00 && code <= 0xDFFF) {
        /* 孤立低位代理：丢弃 */
      } else {
        out += text.charAt(i);
      }
    }
    return out;
  }

  /** 文本节点的 XML 转义（契约要求 & < > " ' 全部转义）。 */
  function xmlEscape(value) {
    return cleanXmlText(value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  /* ================================================================== */
  /* 2. CRC32 + ZIP（store，不压缩）                                     */
  /* ================================================================== */

  var CRC_TABLE = (function () {
    var table = new Uint32Array(256);
    for (var n = 0; n < 256; n += 1) {
      var c = n;
      for (var k = 0; k < 8; k += 1) {
        c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      }
      table[n] = c >>> 0;
    }
    return table;
  })();

  function crc32(bytes) {
    var c = 0xFFFFFFFF;
    for (var i = 0; i < bytes.length; i += 1) {
      c = CRC_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
    }
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  /** 合并多个 Uint8Array 为一块连续内存。 */
  function concatBytes(chunks) {
    var total = 0;
    var i;
    for (i = 0; i < chunks.length; i += 1) { total += chunks[i].length; }
    var out = new Uint8Array(total);
    var offset = 0;
    for (i = 0; i < chunks.length; i += 1) {
      out.set(chunks[i], offset);
      offset += chunks[i].length;
    }
    return out;
  }

  function dosDateTime(date) {
    var d = date instanceof Date ? date : new Date();
    var year = d.getFullYear();
    if (year < 1980) { year = 1980; }
    if (year > 2107) { year = 2107; }
    var time = ((d.getHours() & 0x1F) << 11) | ((d.getMinutes() & 0x3F) << 5) | (Math.floor(d.getSeconds() / 2) & 0x1F);
    var day = ((year - 1980) << 9) | (((d.getMonth() + 1) & 0x0F) << 5) | (d.getDate() & 0x1F);
    return { time: time & 0xFFFF, date: day & 0xFFFF };
  }

  /**
   * 极简 ZIP 写入器：compression method 0（stored）、UTF-8 文件名标记位 0x0800、
   * version needed 20，最后输出本地文件头 + 中央目录 + EOCD。
   */
  function ZipWriter() {
    this.parts = [];
    this.central = [];
    this.offset = 0;
    this.count = 0;
    var stamp = dosDateTime(new Date());
    this.dosTime = stamp.time;
    this.dosDate = stamp.date;
  }

  ZipWriter.prototype.add = function (name, data) {
    var nameBytes = utf8Encode(name);
    var crc = crc32(data);
    var localOffset = this.offset;

    var local = new Uint8Array(30 + nameBytes.length);
    var lv = new DataView(local.buffer);
    lv.setUint32(0, 0x04034B50, true);   // local file header signature
    lv.setUint16(4, 20, true);           // version needed to extract (2.0)
    lv.setUint16(6, 0x0800, true);       // general purpose flag：UTF-8 文件名
    lv.setUint16(8, 0, true);            // compression method：0 = stored
    lv.setUint16(10, this.dosTime, true);
    lv.setUint16(12, this.dosDate, true);
    lv.setUint32(14, crc, true);
    lv.setUint32(18, data.length, true); // compressed size
    lv.setUint32(22, data.length, true); // uncompressed size
    lv.setUint16(26, nameBytes.length, true);
    lv.setUint16(28, 0, true);           // extra field length
    local.set(nameBytes, 30);

    var central = new Uint8Array(46 + nameBytes.length);
    var cv = new DataView(central.buffer);
    cv.setUint32(0, 0x02014B50, true);   // central directory header signature
    cv.setUint16(4, 20, true);           // version made by
    cv.setUint16(6, 20, true);           // version needed
    cv.setUint16(8, 0x0800, true);       // flags（UTF-8）
    cv.setUint16(10, 0, true);           // method：stored
    cv.setUint16(12, this.dosTime, true);
    cv.setUint16(14, this.dosDate, true);
    cv.setUint32(16, crc, true);
    cv.setUint32(20, data.length, true);
    cv.setUint32(24, data.length, true);
    cv.setUint16(28, nameBytes.length, true);
    cv.setUint16(30, 0, true);           // extra
    cv.setUint16(32, 0, true);           // comment
    cv.setUint16(34, 0, true);           // disk number start
    cv.setUint16(36, 0, true);           // internal attributes
    cv.setUint32(38, 0, true);           // external attributes
    cv.setUint32(42, localOffset, true); // local header offset
    central.set(nameBytes, 46);

    this.parts.push(local);
    this.parts.push(data);
    this.central.push(central);
    this.offset += local.length + data.length;
    this.count += 1;
  };

  ZipWriter.prototype.finish = function () {
    var cdChunks = this.central;
    var cdSize = 0;
    var i;
    for (i = 0; i < cdChunks.length; i += 1) { cdSize += cdChunks[i].length; }
    var cdOffset = this.offset;

    var eocd = new Uint8Array(22);
    var ev = new DataView(eocd.buffer);
    ev.setUint32(0, 0x06054B50, true); // end of central directory signature
    ev.setUint16(4, 0, true);          // number of this disk
    ev.setUint16(6, 0, true);          // disk where CD starts
    ev.setUint16(8, this.count, true); // entries on this disk
    ev.setUint16(10, this.count, true);
    ev.setUint32(12, cdSize, true);
    ev.setUint32(16, cdOffset, true);
    ev.setUint16(20, 0, true);         // comment length

    var all = this.parts.slice();
    for (i = 0; i < cdChunks.length; i += 1) { all.push(cdChunks[i]); }
    all.push(eocd);
    return concatBytes(all);
  };

  /* ================================================================== */
  /* 3. 单元格位置 / 样式表                                               */
  /* ================================================================== */

  function colLetter(index) {
    var n = Math.floor(index) + 1;
    var text = '';
    while (n > 0) {
      var m = (n - 1) % 26;
      text = String.fromCharCode(65 + m) + text;
      n = Math.floor((n - 1) / 26);
    }
    return text || 'A';
  }

  function cellRef(colIndex, rowNumber) {
    return colLetter(colIndex) + String(rowNumber);
  }

  var BUILTIN_NUMFMT = {
    '0': 1,
    '0.00': 2,
    '#,##0': 3,
    '#,##0.00': 4,
    '0%': 9,
    '0.00%': 10,
    '0.00E+00': 11,
    'm/d/yy': 14,
    'd-mmm-yy': 15,
    'd-mmm': 16,
    'mmm-yy': 17,
    'h:mm AM/PM': 18,
    'h:mm:ss AM/PM': 19,
    'h:mm': 20,
    'h:mm:ss': 21,
    'm/d/yy h:mm': 22
  };
  var CUSTOM_NUMFMT_START = 164;

  /**
   * 样式登记表。cellXfs 约定：
   *   0 -> 正文（左对齐、垂直居中、细边框）
   *   1 -> 表头（白字加粗、深蓝底、居中）
   *   2+ -> 按需生成的「对齐 / 换行 / 数字格式」组合
   */
  function StyleBook() {
    this.xfs = [];
    this.xfMap = {};
    this.numFmts = [];
    this.numFmtMap = {};
    this.nextNumFmtId = CUSTOM_NUMFMT_START;
    this.bodyIndex = this.addXf({
      fontId: 0, fillId: 0, borderId: 1, numFmtId: 0,
      align: 'left', wrap: false, vertical: 'center'
    }, 'body|left|0|0');
    this.headerIndex = this.addXf({
      fontId: 1, fillId: 2, borderId: 1, numFmtId: 0,
      align: 'center', wrap: false, vertical: 'center'
    }, 'header');
  }

  StyleBook.prototype.addXf = function (xf, key) {
    if (key && Object.prototype.hasOwnProperty.call(this.xfMap, key)) { return this.xfMap[key]; }
    var index = this.xfs.length;
    this.xfs.push(xf);
    if (key) { this.xfMap[key] = index; }
    return index;
  };

  StyleBook.prototype.resolveNumFmt = function (code) {
    if (code === null || code === undefined || code === '') { return 0; }
    var key = String(code);
    if (Object.prototype.hasOwnProperty.call(BUILTIN_NUMFMT, key)) { return BUILTIN_NUMFMT[key]; }
    if (Object.prototype.hasOwnProperty.call(this.numFmtMap, key)) { return this.numFmtMap[key]; }
    var id = this.nextNumFmtId;
    this.nextNumFmtId += 1;
    this.numFmtMap[key] = id;
    this.numFmts.push({ id: id, code: key });
    return id;
  };

  StyleBook.prototype.bodyStyle = function (align, wrap, numFmt) {
    var a = align === 'center' || align === 'right' ? align : 'left';
    var w = wrap ? 1 : 0;
    var numFmtId = this.resolveNumFmt(numFmt);
    var key = 'body|' + a + '|' + w + '|' + numFmtId;
    if (Object.prototype.hasOwnProperty.call(this.xfMap, key)) { return this.xfMap[key]; }
    return this.addXf({
      fontId: 0, fillId: 0, borderId: 1, numFmtId: numFmtId,
      align: a, wrap: w === 1, vertical: 'center'
    }, key);
  };

  StyleBook.prototype.toXml = function () {
    var xml = [XML_DECL, '<styleSheet xmlns="', NS_MAIN, '">'];

    if (this.numFmts.length) {
      xml.push('<numFmts count="', String(this.numFmts.length), '">');
      this.numFmts.forEach(function (fmt) {
        xml.push('<numFmt numFmtId="', String(fmt.id), '" formatCode="', xmlEscape(fmt.code), '"/>');
      });
      xml.push('</numFmts>');
    }

    /* 字体：0 正文；1 表头（加粗、白色） */
    xml.push('<fonts count="2">');
    xml.push('<font><sz val="11"/><color rgb="FF1F2933"/><name val="微软雅黑"/><family val="2"/></font>');
    xml.push('<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="微软雅黑"/><family val="2"/></font>');
    xml.push('</fonts>');

    /* 填充：0 none（必须）、1 gray125（必须）、2 表头深蓝 */
    xml.push('<fills count="3">');
    xml.push('<fill><patternFill patternType="none"/></fill>');
    xml.push('<fill><patternFill patternType="gray125"/></fill>');
    xml.push('<fill><patternFill patternType="solid"><fgColor rgb="', HEADER_FILL_COLOR,
      '"/><bgColor indexed="64"/></patternFill></fill>');
    xml.push('</fills>');

    /* 边框：0 无边框；1 浅灰细边框 */
    var gridSides = '<left style="thin"><color rgb="FFD8DEE6"/></left>' +
      '<right style="thin"><color rgb="FFD8DEE6"/></right>' +
      '<top style="thin"><color rgb="FFD8DEE6"/></top>' +
      '<bottom style="thin"><color rgb="FFD8DEE6"/></bottom>' +
      '<diagonal/>';
    xml.push('<borders count="2">');
    xml.push('<border><left/><right/><top/><bottom/><diagonal/></border>');
    xml.push('<border>', gridSides, '</border>');
    xml.push('</borders>');

    xml.push('<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>');

    xml.push('<cellXfs count="', String(this.xfs.length), '">');
    this.xfs.forEach(function (xf) {
      xml.push('<xf numFmtId="', String(xf.numFmtId),
        '" fontId="', String(xf.fontId),
        '" fillId="', String(xf.fillId),
        '" borderId="', String(xf.borderId),
        '" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"',
        xf.fillId ? ' applyFill="1"' : '',
        xf.numFmtId ? ' applyNumberFormat="1"' : '',
        '><alignment horizontal="', xf.align,
        '" vertical="', (xf.vertical || 'center'),
        '"', xf.wrap ? ' wrapText="1"' : '', '/></xf>');
    });
    xml.push('</cellXfs>');

    xml.push('<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>');
    xml.push('<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>');
    xml.push('</styleSheet>');
    return xml.join('');
  };

  /* ================================================================== */
  /* 4. 单元格 / 工作表 XML                                               */
  /* ================================================================== */

  function formatNumber(value) {
    if (!isFinite(value)) { return '0'; }
    if (Number.isInteger(value)) { return String(value); }
    var text = String(value);
    if (/[eE]/.test(text)) {
      text = value.toFixed(10).replace(/0+$/, '').replace(/\.$/, '');
    }
    return text;
  }

  function textCellXml(ref, styleIndex, text) {
    var escaped = xmlEscape(text);
    if (!escaped) { return ''; }
    return '<c r="' + ref + '" s="' + styleIndex + '" t="inlineStr"><is><t xml:space="preserve">' +
      escaped + '</t></is></c>';
  }

  function numberCellXml(ref, styleIndex, value) {
    return '<c r="' + ref + '" s="' + styleIndex + '"><v>' + formatNumber(value) + '</v></c>';
  }

  function formatDateOnly(date) {
    var d = date instanceof Date ? date : new Date(date);
    if (isNaN(d.getTime())) { return ''; }
    return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate());
  }

  function pad2(n) {
    n = Number(n) || 0;
    return n < 10 ? '0' + n : String(n);
  }

  function normalizeAlign(align) {
    return align === 'center' || align === 'right' ? align : 'left';
  }

  function sanitizeSheetName(name, fallback) {
    var text = name === null || name === undefined ? '' : String(name);
    text = text.replace(/[[\]:*?\/\\]/g, ' ').replace(/\s+/g, ' ').trim();
    text = text.replace(/^'+|'+$/g, '').trim();
    if (!text) { text = fallback || 'Sheet'; }
    if (text.length > MAX_SHEET_NAME) { text = text.slice(0, MAX_SHEET_NAME); }
    return text;
  }

  function uniqueSheetName(name, used) {
    var candidate = name;
    var suffix = 2;
    while (used[candidate]) {
      var tail = ' (' + suffix + ')';
      candidate = name.slice(0, Math.max(1, MAX_SHEET_NAME - tail.length)) + tail;
      suffix += 1;
    }
    return candidate;
  }

  /**
   * 生成单个 worksheet 的 XML。返回 { name, xml }。
   * sheets[i]: { name, columns:[{header,width,numFmt,align,wrap}], rows:[[cell]], freeze:{rows,cols}, autoFilter }
   */
  function buildSheetXml(sheet, styles, index) {
    sheet = sheet && typeof sheet === 'object' ? sheet : {};
    var columns = Array.isArray(sheet.columns) ? sheet.columns : [];
    var rows = Array.isArray(sheet.rows) ? sheet.rows : [];

    var colCount = columns.length;
    var r;
    for (r = 0; r < rows.length; r += 1) {
      if (Array.isArray(rows[r]) && rows[r].length > colCount) { colCount = rows[r].length; }
    }

    var name = sanitizeSheetName(sheet.name, 'Sheet' + (index + 1));

    /* 4.1 列定义（默认宽度 12，限制 2~255） */
    var colDefs = [];
    var colsXml = '';
    if (colCount > 0) {
      var cols = ['<cols>'];
      var c;
      for (c = 0; c < colCount; c += 1) {
        var col = columns[c] || {};
        var width = Number(col.width);
        if (!isFinite(width) || width <= 0) { width = 12; }
        width = Math.min(255, Math.max(2, Math.round(width * 10) / 10));
        colDefs.push({
          width: width,
          align: normalizeAlign(col.align),
          wrap: !!col.wrap,
          numFmt: col.numFmt
        });
        cols.push('<col min="', String(c + 1), '" max="', String(c + 1),
          '" width="', String(width), '" customWidth="1"/>');
      }
      cols.push('</cols>');
      colsXml = cols.join('');
    }

    /* 4.2 表头行（深蓝底白字、行高 24、居中） */
    var headerCells = '';
    for (var hc = 0; hc < colCount; hc += 1) {
      var headerText = columns[hc] && columns[hc].header !== null && columns[hc].header !== undefined
        ? String(columns[hc].header) : '';
      if (!headerText) { continue; }
      headerCells += textCellXml(cellRef(hc, 1), styles.headerIndex, headerText);
    }
    var headerRowXml = '<row r="1" ht="24" customHeight="1">' + headerCells + '</row>';

    /* 4.3 数据行 */
    var rowXmlList = [];
    for (r = 0; r < rows.length; r += 1) {
      var row = Array.isArray(rows[r]) ? rows[r] : [];
      var cells = ['<row r="', String(r + 2), '">'];
      var hasCell = false;
      for (var ci = 0; ci < colCount; ci += 1) {
        var value = row[ci];
        if (value === null || value === undefined || value === '') { continue; }
        var def = colDefs[ci];
        var ref = cellRef(ci, r + 2);
        var cellXml;
        if (typeof value === 'number') {
          if (!isFinite(value)) { continue; }
          cellXml = numberCellXml(ref, styles.bodyStyle(def.align, false, def.numFmt), value);
        } else if (typeof value === 'boolean') {
          cellXml = textCellXml(ref, styles.bodyStyle(def.align, def.wrap, def.numFmt), value ? '是' : '否');
        } else if (value instanceof Date) {
          var dateText = formatDateOnly(value);
          cellXml = dateText ? textCellXml(ref, styles.bodyStyle(def.align, def.wrap, def.numFmt), dateText) : '';
        } else if (typeof value === 'object') {
          var jsonText = '';
          try { jsonText = JSON.stringify(value); } catch (e) { jsonText = ''; }
          cellXml = jsonText ? textCellXml(ref, styles.bodyStyle(def.align, def.wrap, def.numFmt), jsonText) : '';
        } else {
          cellXml = textCellXml(ref, styles.bodyStyle(def.align, def.wrap, def.numFmt), String(value));
        }
        if (cellXml) {
          cells.push(cellXml);
          hasCell = true;
        }
      }
      cells.push('</row>');
      if (hasCell || colCount === 0) { rowXmlList.push(cells.join('')); }
      else { rowXmlList.push('<row r="' + (r + 2) + '"></row>'); }
    }

    /* 4.4 冻结窗格 */
    var freeze = sheet.freeze && typeof sheet.freeze === 'object' ? sheet.freeze : {};
    var freezeRows = Math.max(0, Math.min(1048575, Math.floor(Number(freeze.rows) || 0)));
    var freezeCols = Math.max(0, Math.min(16383, Math.floor(Number(freeze.cols) || 0)));
    var viewParts = ['<sheetView workbookViewId="0"'];
    if (index === 0) { viewParts.push(' tabSelected="1"'); }
    viewParts.push('>');
    if (freezeRows > 0 || freezeCols > 0) {
      var topLeft = colLetter(freezeCols) + String(freezeRows + 1);
      var activePane = freezeRows > 0 && freezeCols > 0 ? 'bottomRight'
        : (freezeRows > 0 ? 'bottomLeft' : 'topRight');
      viewParts.push('<pane');
      if (freezeCols > 0) { viewParts.push(' xSplit="', String(freezeCols), '"'); }
      if (freezeRows > 0) { viewParts.push(' ySplit="', String(freezeRows), '"'); }
      viewParts.push(' topLeftCell="', topLeft, '" activePane="', activePane, '" state="frozen"/>');
      viewParts.push('<selection pane="', activePane, '" activeCell="', topLeft,
        '" sqref="', topLeft, '"/>');
    } else {
      viewParts.push('<selection activeCell="A1" sqref="A1"/>');
    }
    viewParts.push('</sheetView>');
    var sheetViewsXml = '<sheetViews>' + viewParts.join('') + '</sheetViews>';

    /* 4.5 尺寸 / 自动筛选 */
    var lastRow = rows.length + 1;
    var lastColLetter = colLetter(Math.max(0, colCount - 1));
    var dimension = colCount > 0 ? 'A1:' + lastColLetter + lastRow : 'A1';
    var autoFilterXml = '';
    if (sheet.autoFilter && colCount > 0) {
      autoFilterXml = '<autoFilter ref="A1:' + lastColLetter + lastRow + '"/>';
    }

    var xml = XML_DECL + '<worksheet xmlns="' + NS_MAIN + '" xmlns:r="' + NS_REL + '">' +
      '<dimension ref="' + dimension + '"/>' +
      sheetViewsXml +
      '<sheetFormatPr defaultRowHeight="18" baseColWidth="10"/>' +
      colsXml +
      '<sheetData>' + headerRowXml + rowXmlList.join('') + '</sheetData>' +
      autoFilterXml +
      '</worksheet>';

    return { name: name, xml: xml };
  }

  /* ================================================================== */
  /* 5. 工作簿打包                                                       */
  /* ================================================================== */

  function contentTypesXml(sheetCount) {
    var xml = [XML_DECL, '<Types xmlns="', NS_CONTENT, '">'];
    xml.push('<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>');
    xml.push('<Default Extension="xml" ContentType="application/xml"/>');
    xml.push('<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>');
    for (var i = 1; i <= sheetCount; i += 1) {
      xml.push('<Override PartName="/xl/worksheets/sheet', String(i),
        '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>');
    }
    xml.push('<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>');
    xml.push('<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>');
    xml.push('</Types>');
    return xml.join('');
  }

  function rootRelsXml() {
    return XML_DECL + '<Relationships xmlns="' + NS_PKG_REL + '">' +
      '<Relationship Id="rId1" Type="' + NS_REL + '/officeDocument" Target="xl/workbook.xml"/>' +
      '<Relationship Id="rId2" Type="' + NS_PKG_REL + '/metadata/core-properties" Target="docProps/core.xml"/>' +
      '</Relationships>';
  }

  function workbookXml(sheetFiles) {
    var parts = [XML_DECL, '<workbook xmlns="', NS_MAIN, '" xmlns:r="', NS_REL, '">'];
    parts.push('<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="14400" activeTab="0"/></bookViews>');
    parts.push('<sheets>');
    for (var i = 0; i < sheetFiles.length; i += 1) {
      parts.push('<sheet name="', xmlEscape(sheetFiles[i].name), '" sheetId="', String(i + 1),
        '" r:id="rId', String(i + 1), '"/>');
    }
    parts.push('</sheets>');
    parts.push('</workbook>');
    return parts.join('');
  }

  function workbookRelsXml(sheetCount) {
    var parts = [XML_DECL, '<Relationships xmlns="', NS_PKG_REL, '">'];
    for (var i = 1; i <= sheetCount; i += 1) {
      parts.push('<Relationship Id="rId', String(i), '" Type="', NS_REL,
        '/worksheet" Target="worksheets/sheet', String(i), '.xml"/>');
    }
    parts.push('<Relationship Id="rId', String(sheetCount + 1), '" Type="', NS_REL,
      '/styles" Target="styles.xml"/>');
    parts.push('</Relationships>');
    return parts.join('');
  }

  function corePropsXml(title) {
    var now = new Date().toISOString();
    var docTitle = title && String(title).trim() ? String(title).trim() : '客户跟进台账';
    return XML_DECL + '<cp:coreProperties xmlns:cp="' + NS_CP + '" xmlns:dc="' + NS_DC +
      '" xmlns:dcterms="' + NS_DCTERMS + '" xmlns:dcmitype="' + NS_DCMITYPE +
      '" xmlns:xsi="' + NS_XSI + '">' +
      '<dc:title>' + xmlEscape(docTitle) + '</dc:title>' +
      '<dc:creator>客户通 CustomerCRM</dc:creator>' +
      '<cp:lastModifiedBy>客户通 CustomerCRM</cp:lastModifiedBy>' +
      '<dcterms:created xsi:type="dcterms:W3CDTF">' + now + '</dcterms:created>' +
      '<dcterms:modified xsi:type="dcterms:W3CDTF">' + now + '</dcterms:modified>' +
      '</cp:coreProperties>';
  }

  /**
   * 组装并打包 .xlsx。
   * @param {Array} sheets 工作表定义数组
   * @param {Object} [options] { title } 文档属性标题
   * @returns {Uint8Array}
   */
  function buildWorkbook(sheets, options) {
    var list = Array.isArray(sheets) ? sheets : [];
    if (!list.length) {
      list = [{ name: 'Sheet1', columns: [], rows: [] }];
    }
    var styles = new StyleBook();
    var usedNames = Object.create(null);
    var sheetFiles = [];

    for (var i = 0; i < list.length; i += 1) {
      var built = buildSheetXml(list[i], styles, i);
      var name = uniqueSheetName(built.name, usedNames);
      usedNames[name] = true;
      sheetFiles.push({ name: name, xml: built.xml });
    }

    var parts = [];
    parts.push({ name: '[Content_Types].xml', xml: contentTypesXml(sheetFiles.length) });
    parts.push({ name: '_rels/.rels', xml: rootRelsXml() });
    parts.push({ name: 'docProps/core.xml', xml: corePropsXml(options && options.title) });
    parts.push({ name: 'xl/workbook.xml', xml: workbookXml(sheetFiles) });
    parts.push({ name: 'xl/_rels/workbook.xml.rels', xml: workbookRelsXml(sheetFiles.length) });
    for (var s = 0; s < sheetFiles.length; s += 1) {
      parts.push({ name: 'xl/worksheets/sheet' + (s + 1) + '.xml', xml: sheetFiles[s].xml });
    }
    parts.push({ name: 'xl/styles.xml', xml: styles.toXml() });

    var zip = new ZipWriter();
    for (var p = 0; p < parts.length; p += 1) {
      zip.add(parts[p].name, utf8Encode(parts[p].xml));
    }
    return zip.finish();
  }

  /* ================================================================== */
  /* 6. 客户台账导出：两个工作表、21 + 11 列（v1.1）                       */
  /* ================================================================== */

  /* v1.1：客户总表 21 列（新增「产品型号」「主营业务」） */
  var TOTAL_COLUMNS = [
    { header: '序号', width: 8, align: 'center' },
    { header: '首次询盘时间', width: 15, align: 'center' },
    { header: '客户名称', width: 20 },
    { header: '咨询产品', width: 26 },
    { header: '产品型号', width: 16 },
    { header: '国家', width: 14, align: 'center' },
    { header: '公司名', width: 30 },
    { header: '主营业务', width: 22 },
    { header: '成立时间', width: 14, align: 'center' },
    { header: '团队规模', width: 16 },
    { header: '年营收', width: 18 },
    { header: '跟进状态', width: 14, align: 'center' },
    { header: '是否成交', width: 12, align: 'center' },
    { header: '核心原因', width: 46, wrap: true },
    { header: '距上次跟进(天)', width: 17, align: 'center', numFmt: '0' },
    { header: '最近跟进①日期', width: 16, align: 'center' },
    { header: '最近跟进①内容', width: 46, wrap: true },
    { header: '最近跟进②日期', width: 16, align: 'center' },
    { header: '最近跟进②内容', width: 46, wrap: true },
    { header: '最近跟进③日期', width: 16, align: 'center' },
    { header: '最近跟进③内容', width: 46, wrap: true }
  ];

  /* v1.1：跟进明细 11 列（新增「主营业务」） */
  var DETAIL_COLUMNS = [
    { header: '序号', width: 8, align: 'center' },
    { header: '首次询盘时间', width: 15, align: 'center' },
    { header: '客户名称', width: 20 },
    { header: '公司名', width: 30 },
    { header: '主营业务', width: 22 },
    { header: '国家', width: 14, align: 'center' },
    { header: '跟进日期', width: 15, align: 'center' },
    { header: '距今天数', width: 12, align: 'center', numFmt: '0' },
    { header: '跟进内容', width: 60, wrap: true },
    { header: '跟进状态', width: 14, align: 'center' },
    { header: '是否成交', width: 12, align: 'center' }
  ];

  function pad2Local(n) {
    n = Number(n) || 0;
    return n < 10 ? '0' + n : String(n);
  }

  function todayISO() {
    var d = new Date();
    return d.getFullYear() + '-' + pad2Local(d.getMonth() + 1) + '-' + pad2Local(d.getDate());
  }

  function nowStamp() {
    var d = new Date();
    return String(d.getFullYear()) + pad2Local(d.getMonth() + 1) + pad2Local(d.getDate()) +
      '_' + pad2Local(d.getHours()) + pad2Local(d.getMinutes());
  }

  function isValidISODate(iso) {
    if (typeof iso !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(iso)) { return false; }
    var parts = iso.split('-');
    var y = Number(parts[0]);
    var m = Number(parts[1]);
    var d = Number(parts[2]);
    var dt = new Date(Date.UTC(y, m - 1, d));
    return dt.getUTCFullYear() === y && dt.getUTCMonth() === m - 1 && dt.getUTCDate() === d;
  }

  function daysBetweenISO(fromISO, toISODate) {
    var a = Date.parse(fromISO + 'T00:00:00Z');
    var b = Date.parse(toISODate + 'T00:00:00Z');
    if (isNaN(a) || isNaN(b)) { return 0; }
    return Math.round((b - a) / 86400000);
  }

  function asText(value) {
    if (value === null || value === undefined) { return ''; }
    if (typeof value === 'number') { return String(value); }
    if (typeof value === 'boolean') { return value ? '是' : '否'; }
    return String(value);
  }

  /** 内置跟进清洗：解析日期、倒序、去重、最多 3 条（schema 不可用时的兜底）。 */
  function normalizeFollowupsFallback(list) {
    var out = [];
    if (Array.isArray(list)) {
      list.forEach(function (item) {
        if (!item || typeof item !== 'object') { return; }
        var date = asText(item.date);
        if (!isValidISODate(date)) { return; }
        var note = asText(item.note || item.text || item.content).slice(0, 300);
        out.push({ date: date, note: note });
      });
    }
    out.sort(function (a, b) { return a.date < b.date ? 1 : (a.date > b.date ? -1 : 0); });
    var seen = {};
    var dedup = [];
    out.forEach(function (item) {
      var key = item.date + '|' + item.note;
      if (seen[key]) { return; }
      seen[key] = true;
      dedup.push(item);
    });
    return dedup.slice(0, 3);
  }

  /** 内置派生（schema.derive 不可用时的兜底），字段与 schema.derive 对齐。 */
  function deriveFallback(record, today) {
    var rec = record && typeof record === 'object' ? record : {};
    var followups = normalizeFollowupsFallback(rec.followups);
    var last3 = followups.map(function (item) {
      return { date: item.date, note: item.note, daysAgo: Math.max(0, daysBetweenISO(item.date, today)) };
    });
    var last = last3.length ? last3[0] : null;
    return {
      firstInquiryAt: asText(rec.firstInquiryAt),
      customerName: asText(rec.customerName),
      product: asText(rec.product),
      productModel: asText(rec.productModel),
      country: asText(rec.country),
      companyName: asText(rec.companyName),
      mainBusiness: asText(rec.mainBusiness),
      foundedAt: asText(rec.foundedAt),
      teamSize: asText(rec.teamSize),
      annualRevenue: asText(rec.annualRevenue),
      followupStatus: asText(rec.followupStatus) || '新询盘',
      dealStatus: asText(rec.dealStatus) || '未成交',
      coreReason: asText(rec.coreReason),
      followups: followups,
      last3: last3,
      lastFollowupAt: last ? last.date : '',
      daysSince: last ? last.daysAgo : null
    };
  }

  /** 记录已带派生字段（last3 / daysSince）但没有 followups 时的兜底组装。 */
  function deriveFromPrecomputed(record, today) {
    var rec = record && typeof record === 'object' ? record : {};
    var last3 = [];
    var list = Array.isArray(rec.last3) ? rec.last3 : [];
    list.forEach(function (item) {
      if (!item || typeof item !== 'object') { return; }
      var date = asText(item.date);
      if (!isValidISODate(date)) { return; }
      var daysAgo = item.daysAgo === null || item.daysAgo === undefined
        ? Math.max(0, daysBetweenISO(date, today))
        : Math.max(0, Number(item.daysAgo) || 0);
      last3.push({ date: date, note: asText(item.note).slice(0, 300), daysAgo: daysAgo });
    });
    last3 = last3.slice(0, 3);
    var daysSince = rec.daysSince === null || rec.daysSince === undefined
      ? (last3.length ? last3[0].daysAgo : null)
      : Math.max(0, Number(rec.daysSince) || 0);
    return {
      firstInquiryAt: asText(rec.firstInquiryAt),
      customerName: asText(rec.customerName),
      product: asText(rec.product),
      productModel: asText(rec.productModel),
      country: asText(rec.country),
      companyName: asText(rec.companyName),
      mainBusiness: asText(rec.mainBusiness),
      foundedAt: asText(rec.foundedAt),
      teamSize: asText(rec.teamSize),
      annualRevenue: asText(rec.annualRevenue),
      followupStatus: asText(rec.followupStatus) || '新询盘',
      dealStatus: asText(rec.dealStatus) || '未成交',
      coreReason: asText(rec.coreReason),
      followups: last3,
      last3: last3,
      lastFollowupAt: last3.length ? last3[0].date : '',
      daysSince: daysSince
    };
  }

  /**
   * 优先使用 schema.derive（与界面口径一致）；异常时退回内置实现。
   * 兼容两种入参：原始记录（11 字段 + followups）与已带派生字段的记录（last3/daysSince）。
   */
  function deriveRecord(record, today, schema) {
    var rec = record && typeof record === 'object' ? record : {};
    /* 已派生但没有 followups 的记录：直接沿用派生好的最近跟进，不丢信息 */
    if (!Array.isArray(rec.followups) && Array.isArray(rec.last3)) {
      return deriveFromPrecomputed(rec, today);
    }
    if (schema && typeof schema.derive === 'function') {
      try {
        var derived = schema.derive(rec, { today: today });
        if (derived && Array.isArray(derived.last3)) {
          if (derived.daysSince === undefined) { derived.daysSince = null; }
          if (!Array.isArray(derived.followups)) { derived.followups = []; }
          return derived;
        }
      } catch (err) { /* 回退到内置派生 */ }
    }
    return deriveFallback(rec, today);
  }

  /**
   * 查找共享 schema：浏览器用 window.CRM_SCHEMA；Node 下依次尝试全局与 require('./schema.js')。
   * 找不到也不影响导出（使用内置派生）。
   */
  function resolveSchema() {
    if (root && root.CRM_SCHEMA && typeof root.CRM_SCHEMA.derive === 'function') {
      return root.CRM_SCHEMA;
    }
    if (typeof module === 'object' && module.exports && typeof require === 'function') {
      try {
        var loaded = require('./schema.js');
        if (loaded && typeof loaded.derive === 'function') { return loaded; }
      } catch (err) { /* 单文件环境：没有独立 schema 模块 */ }
    }
    return null;
  }

  function resolveFilename(name) {
    var fallback = '客户跟进台账_' + nowStamp() + '.xlsx';
    if (name === null || name === undefined) { return fallback; }
    var text = String(name).trim();
    if (!text) { return fallback; }
    /* 只接受文件名：去掉路径字符，避免误写目录 */
    text = text.replace(/[\\/:*?"<>|]/g, '_').replace(/\s+/g, ' ').trim();
    if (!/\.xlsx$/i.test(text)) { text += '.xlsx'; }
    return text;
  }

  function resolveToday(value) {
    if (typeof value === 'string' && isValidISODate(value)) { return value; }
    return todayISO();
  }

  /**
   * 导出客户台账。
   * @param {Array} records 客户记录数组
   * @param {Object} [opts] { filename?, today: 'YYYY-MM-DD', title? }
   * @returns {{ filename: string, bytes: Uint8Array }}
   */
  function exportCustomers(records, opts) {
    opts = opts && typeof opts === 'object' ? opts : {};
    var list = Array.isArray(records) ? records : [];
    var today = resolveToday(opts.today);
    var schema = resolveSchema();
    var derived = list.map(function (record) { return deriveRecord(record, today, schema); });

    var totalRows = [];
    var detailRows = [];
    var seq = 1;

    derived.forEach(function (item, index) {
      var last3 = Array.isArray(item.last3) ? item.last3 : [];
      var f0 = last3[0] || null;
      var f1 = last3[1] || null;
      var f2 = last3[2] || null;
      totalRows.push([
        index + 1,
        asText(item.firstInquiryAt),
        asText(item.customerName),
        asText(item.product),
        asText(item.productModel),
        asText(item.country),
        asText(item.companyName),
        asText(item.mainBusiness),
        asText(item.foundedAt),
        asText(item.teamSize),
        asText(item.annualRevenue),
        asText(item.followupStatus),
        asText(item.dealStatus),
        asText(item.coreReason),
        item.daysSince === null || item.daysSince === undefined ? '' : Number(item.daysSince),
        f0 ? f0.date : '',
        f0 ? f0.note : '',
        f1 ? f1.date : '',
        f1 ? f1.note : '',
        f2 ? f2.date : '',
        f2 ? f2.note : ''
      ]);

      last3.forEach(function (fu) {
        detailRows.push([
          seq,
          asText(item.firstInquiryAt),
          asText(item.customerName),
          asText(item.companyName),
          asText(item.mainBusiness),
          asText(item.country),
          fu.date,
          fu.daysAgo === null || fu.daysAgo === undefined ? '' : Number(fu.daysAgo),
          fu.note || '',
          asText(item.followupStatus),
          asText(item.dealStatus)
        ]);
        seq += 1;
      });
    });

    var sheets = [
      {
        name: '客户总表',
        columns: TOTAL_COLUMNS,
        rows: totalRows,
        freeze: { rows: 1 },
        autoFilter: true
      },
      {
        name: '跟进明细',
        columns: DETAIL_COLUMNS,
        rows: detailRows,
        freeze: { rows: 1 },
        autoFilter: true
      }
    ];

    var bytes = buildWorkbook(sheets, { title: opts.title || '客户跟进台账' });
    return { filename: resolveFilename(opts.filename), bytes: bytes };
  }

  /* ================================================================== */
  /* 7. 浏览器下载                                                        */
  /* ================================================================== */

  /**
   * 触发浏览器下载。Node / 无 DOM 环境下抛明确错误（不静默失败）。
   * @param {Uint8Array} bytes
   * @param {string} [filename]
   */
  function download(bytes, filename) {
    var canUseBrowser = typeof document !== 'undefined' &&
      typeof Blob !== 'undefined' &&
      typeof URL !== 'undefined' &&
      typeof URL.createObjectURL === 'function';

    if (!canUseBrowser) {
      throw new Error('CrmXlsx.download 只能在浏览器中使用；Node 环境请使用 exportCustomers() 返回的 bytes 自行写文件');
    }
    if (!bytes || !bytes.length) {
      throw new Error('CrmXlsx.download 需要非空的 .xlsx 字节数据');
    }

    var name = resolveFilename(filename);
    var blob = new Blob([bytes], { type: MIME_TYPE });
    var url = URL.createObjectURL(blob);
    var link = document.createElement('a');
    link.href = url;
    link.download = name;
    link.rel = 'noopener';
    link.style.display = 'none';
    document.body.appendChild(link);
    try {
      link.click();
    } finally {
      setTimeout(function () {
        try { document.body.removeChild(link); } catch (err) { /* 已被移除 */ }
        try { URL.revokeObjectURL(url); } catch (err) { /* 忽略 */ }
      }, 1000);
    }
    return true;
  }

  /* ================================================================== */

  return {
    VERSION: VERSION,
    MIME_TYPE: MIME_TYPE,
    buildWorkbook: buildWorkbook,
    exportCustomers: exportCustomers,
    download: download
  };
});
