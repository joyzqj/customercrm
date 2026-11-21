/*!
 * 客户通 CustomerCRM — js/parser.js
 * 文本识别引擎：把一整段中/英询盘文本识别成客户字段（纯函数，Node 可测）。
 * 设计原则：
 *   1. 只做识别，不做业务决策；识别不到就不输出该字段（不填空串）。
 *   2. 全部规则为正则 + 小函数，函数短、注释足，便于后续维护。
 *   3. 绝不抛异常；性能上避免灾难性回溯，10 万字文本仍应 < 200ms。
 * 依赖：可复用 js/schema.js 的国家表 / 宽松日期 / 字段表；缺失时降级兜底。
 */
(function (root, factory) {
  'use strict';
  var api = factory();
  if (typeof module === 'object' && module.exports) { module.exports = api; }
  if (root) { root.CrmParser = api; }
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  /* ================================================================== */
  /* 0. 依赖与常量                                                       */
  /* ================================================================== */

  var schema = null;
  try {
    if (typeof require === 'function') { schema = require('./schema.js'); }
  } catch (e) { schema = null; }
  if (!schema && typeof globalThis !== 'undefined' && globalThis.CRM_SCHEMA) { schema = globalThis.CRM_SCHEMA; }
  if (!schema && typeof window !== 'undefined' && window.CRM_SCHEMA) { schema = window.CRM_SCHEMA; }

  var VERSION = '1.1.0';

  /* 13 个字段的固定顺序（与 schema.FIELDS 保持一致；缺 schema 时用兜底表） */
  var FALLBACK_FIELD_KEYS = [
    'firstInquiryAt', 'customerName', 'product', 'productModel', 'country', 'companyName',
    'mainBusiness', 'foundedAt', 'teamSize', 'annualRevenue', 'followupStatus', 'dealStatus', 'coreReason'
  ];
  var FIELD_KEYS = FALLBACK_FIELD_KEYS.slice();
  if (schema && Array.isArray(schema.FIELDS) && schema.FIELDS.length) {
    var fromSchema = schema.FIELDS.map(function (f) { return f.key; });
    if (fromSchema.length) { FIELD_KEYS = fromSchema; }
  }

  /* 单次识别允许处理的最大字符数（超长文本截断，保证性能与稳定性） */
  var MAX_TEXT_LENGTH = 200000;

  /* 供界面「示例」按钮使用：三段原文直接取契约 §5.3 */
  var SAMPLE_TEXTS = [
    {
      label: '中文询盘',
      text: '2026年9月12日收到德国客户 Michael Braun 的询盘。他所在的 Braun Trading GmbH 成立于1998年，主营户外用品，团队规模约50-100人，年营收约1200万欧元。客户咨询的是我们2026款不锈钢保温杯 500ml，要求加印LOGO。目前已报价，客户反馈价格略高于目标价8%，仍在等终端客户确认。11月10日电话沟通确认2000pcs试单意向，11月18日邮件确认样品尺寸。'
    },
    {
      label: '英文询盘',
      text: 'Michael Braun from Braun Trading GmbH (Germany) reached out on Sep 12, 2026 regarding our 500ml stainless steel vacuum flask with custom logo printing. The company was founded in 1998, with a team of 50-100 employees and annual revenue around EUR 12 million. We sent our quotation on Nov 10; the client said our price is about 8% above their target. The deal was closed on Nov 20.'
    },
    {
      label: '展会名片速记',
      text: '广交会名片：Mr. Ahmed Hassan, Cairo, Egypt。公司 El Nile Import & Export Co., Ltd，做家居用品进口，年采购额约 300 万美元，团队 15 人，2011 年成立。对我们的陶瓷餐具套装感兴趣，已发样册，等回复。'
    }
  ];

  /* ================================================================== */
  /* 1. 基础工具                                                         */
  /* ================================================================== */

  function pad2(n) {
    n = Number(n) || 0;
    return n < 10 ? '0' + n : String(n);
  }

  var DAYS_IN_MONTH = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];

  /** 构造合法 ISO 日期；非法（2 月 30 日等）返回 ''。与 schema 口径一致。 */
  function isoOf(year, month, day) {
    var y = Number(year);
    var m = Number(month);
    var d = Number(day);
    if (!isFinite(y) || !isFinite(m) || !isFinite(d)) { return ''; }
    if (y < 100) { y = y >= 70 ? 1900 + y : 2000 + y; }
    if (y < 1900 || y > 2100) { return ''; }
    if (m < 1 || m > 12 || d < 1) { return ''; }
    var leap = (y % 4 === 0 && y % 100 !== 0) || (y % 400 === 0);
    var max = DAYS_IN_MONTH[m - 1] + (m === 2 && leap ? 1 : 0);
    if (d > max) { return ''; }
    return y + '-' + pad2(m) + '-' + pad2(d);
  }

  /** 全角字符 → 半角（1:1 替换，保证与原文下标对齐，evidence 可回切原文） */
  function toHalfWidth(text) {
    var out = '';
    for (var i = 0; i < text.length; i += 1) {
      var code = text.charCodeAt(i);
      if (code === 0x3000) { out += ' '; }
      else if (code >= 0xff01 && code <= 0xff5e) { out += String.fromCharCode(code - 0xfee0); }
      else if (code === 0x00a0) { out += ' '; }
      else { out += text.charAt(i); }
    }
    return out;
  }

  function collapseSpace(text) {
    return String(text === null || text === undefined ? '' : text).replace(/\s+/g, ' ').trim();
  }

  /** evidence.snippet：从原文截 ≤60 字；换行/制表符替换为空格（界面可直接渲染） */
  function makeSnippet(raw, index, len) {
    var s = Math.max(0, Number(index) || 0);
    var l = Math.max(0, Number(len) || 0);
    var picked;
    if (s >= raw.length) {
      picked = raw.slice(0, 60);
    } else if (l >= 60) {
      picked = raw.slice(s, s + 60);
    } else {
      var start = Math.max(0, s - 10);
      var end = Math.min(raw.length, s + l + 40);
      if (end - start > 60) { end = start + 60; }
      picked = raw.slice(start, end);
    }
    var snippet = picked.replace(/[\r\n\t]+/g, ' ').replace(/^\s+|\s+$/g, '');
    if (!snippet) { snippet = raw.slice(s, Math.min(raw.length, s + 60)).replace(/[\r\n\t]+/g, ' '); }
    return snippet;
  }

  function uniqPush(list, item) {
    if (item && list.indexOf(item) === -1) { list.push(item); }
  }

  /** 兜底年份：opts.defaultYear > opts.today 的年份 > 当前年份 */
  function resolveFallbackYear(opts) {
    var y = 0;
    if (opts && typeof opts === 'object') {
      y = Number(opts.defaultYear);
      if (!(y >= 1900 && y <= 2100)) { y = Number(String(opts.today || '').slice(0, 4)); }
    }
    if (!(y >= 1900 && y <= 2100)) { y = new Date().getFullYear(); }
    if (!(y >= 1900 && y <= 2100)) { y = 2026; }
    return y;
  }

  /* ================================================================== */
  /* 2. 日期识别（全字段共用）                                            */
  /* ================================================================== */

  var MONTHS3 = {
    jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6,
    jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12
  };

  function monthFromToken(token) {
    var t = String(token || '').replace(/\./g, '').slice(0, 3).toLowerCase();
    return MONTHS3[t] || 0;
  }

  /**
   * 日期候选模式（按优先级排列；重叠时保留更靠前/更长的匹配）。
   * make(m, fallbackYear) -> 'YYYY-MM-DD' | ''
   */
  var DATE_PATTERNS = [
    /* 2026-09-12 / 2026/9/12 / 2026.9.12 */
    {
      id: 'iso',
      re: /(\d{4})\s*[-\/.]\s*(\d{1,2})\s*[-\/.]\s*(\d{1,2})(?!\d)/g,
      make: function (m) { return isoOf(m[1], m[2], m[3]); }
    },
    /* 2026年9月12日 */
    {
      id: 'cn-full',
      re: /(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*[日号]?/g,
      make: function (m) { return isoOf(m[1], m[2], m[3]); }
    },
    /* 26年9月12日（两位年份） */
    {
      id: 'cn-2y',
      re: /(\d{2})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*[日号]?/g,
      make: function (m) { return isoOf(m[1], m[2], m[3]); }
    },
    /* 12 Sep 2026 / 12th September, 2026（先于「月在前」模式，避免 Sep 2026 误判） */
    {
      id: 'en-dmy',
      re: /\b(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]{3,10})\.?(?:\s*,?\s*(\d{4}))?\b/g,
      make: function (m, fallbackYear) {
        var mon = monthFromToken(m[2]);
        if (!mon) { return ''; }
        return isoOf(m[3] || fallbackYear, mon, m[1]);
      }
    },
    /* Sep 12, 2026 / September 12 2026 / Sep 12 */
    {
      id: 'en-mdy',
      re: /\b([A-Za-z]{3,10})\.?\s+(\d{1,2})(?:st|nd|rd|th)?(?:\s*,?\s*(\d{4}))?\b/g,
      make: function (m, fallbackYear) {
        var mon = monthFromToken(m[1]);
        if (!mon) { return ''; }
        return isoOf(m[3] || fallbackYear, mon, m[2]);
      }
    },
    /* 9月12日 / 12月20号 */
    {
      id: 'cn-md',
      re: /(\d{1,2})\s*月\s*(\d{1,2})\s*[日号]/g,
      make: function (m, fallbackYear) { return isoOf(fallbackYear, m[1], m[2]); }
    },
    /* 12/09/2026 或 25/09/2026（无 4 位年份在前，按 schema 口诀：>12 视为日在前） */
    {
      id: 'num-dmy',
      re: /\b(\d{1,2})\s*[-\/.]\s*(\d{1,2})\s*[-\/.]\s*(\d{2,4})(?!\d)/g,
      make: function (m) {
        var b = Number(m[1]);
        var c = Number(m[2]);
        return b > 12 ? isoOf(m[3], c, b) : isoOf(m[3], b, c);
      }
    }
  ];

  /**
   * 收集所有「完整日期」候选：按位置排序，重叠时保留靠前/较长的那个。
   * 返回 [{ iso, index, len }]，index/len 基于半角化后的文本（与原文 1:1 对齐）。
   */
  function collectDates(norm, fallbackYear) {
    var found = [];
    for (var p = 0; p < DATE_PATTERNS.length; p += 1) {
      var def = DATE_PATTERNS[p];
      var re = def.re;
      re.lastIndex = 0;
      var m = re.exec(norm);
      while (m !== null) {
        if (m[0].length === 0) { re.lastIndex += 1; m = re.exec(norm); continue; }
        var iso = def.make(m, fallbackYear);
        if (iso) { found.push({ iso: iso, index: m.index, len: m[0].length }); }
        m = re.exec(norm);
      }
    }
    found.sort(function (a, b) { return (a.index - b.index) || (b.len - a.len); });
    var out = [];
    var maxEnd = -1;
    for (var i = 0; i < found.length; i += 1) {
      if (found[i].index < maxEnd) { continue; }
      out.push(found[i]);
      maxEnd = found[i].index + found[i].len;
    }
    return out;
  }

  /* 判断日期附近的语义：是询盘/联系日期，还是成立日期 */
  var INQUIRY_CTX_RE = /询盘|询价|来函|来信|收到|联系|inquiry|enquiry|reached\s+out|contacted|first\s+email|emailed/i;
  var FOUNDED_CTX_RE = /成立|创立|创办|始创|建于|founded|established|\bsince\b|inception/i;

  function tagDates(norm, dates) {
    for (var i = 0; i < dates.length; i += 1) {
      var d = dates[i];
      /* 询盘语义看前后 20 字；成立语义只看前 12 字，避免相邻句子互相污染 */
      var inBefore = norm.slice(Math.max(0, d.index - 20), d.index);
      var inAfter = norm.slice(d.index + d.len, Math.min(norm.length, d.index + d.len + 20));
      var fdBefore = norm.slice(Math.max(0, d.index - 12), d.index);
      d.inquiry = INQUIRY_CTX_RE.test((inBefore + ' ' + inAfter).toLowerCase());
      d.founded = FOUNDED_CTX_RE.test((fdBefore + ' ' + inAfter.slice(0, 6)).toLowerCase());
    }
  }

  /** 首次询盘日期：优先「带询盘语义」的日期，其次全文最靠前的非成立日期。 */
  function pickInquiryDate(dates) {
    var labeled = [];
    var plain = [];
    for (var i = 0; i < dates.length; i += 1) {
      if (dates[i].founded) { continue; }
      if (dates[i].inquiry) { labeled.push(dates[i]); }
      plain.push(dates[i]);
    }
    if (labeled.length) {
      labeled[0].rule = 'inquiry.labeled';
      return labeled[0];
    }
    if (plain.length) {
      plain[0].rule = 'inquiry.earliest';
      return plain[0];
    }
    return null;
  }

  /* ================================================================== */
  /* 3. 跟进记录（日期 + 事件）                                           */
  /* ================================================================== */

  function isSentenceStop(text, i) {
    var c = text.charAt(i);
    if (c === '\n' || c === '\r' || '。；;！!?'.indexOf(c) !== -1) { return true; }
    if (c === '.') {
      var next = text.charAt(i + 1);
      if (next === '' || /\s/.test(next)) { return true; }
    }
    return false;
  }

  /** 向左右扩展出一句（限制 80 字，避免长文本开销） */
  function expandSegment(text, index, len) {
    var limit = 80;
    var i = index - 1;
    while (i >= 0 && index - i <= limit) {
      if (isSentenceStop(text, i)) { break; }
      i -= 1;
    }
    var start = Math.max(0, i + 1);
    var end = index + len;
    var j = end;
    while (j < text.length && j - end <= limit) {
      if (isSentenceStop(text, j)) { break; }
      j += 1;
    }
    end = Math.min(text.length, j);
    return { start: start, end: end };
  }

  /** 在句子内进一步收缩到逗号分隔的小分句 */
  function expandClause(text, segStart, segEnd, index, len) {
    var limit = 40;
    var a = index;
    while (a > segStart && index - a < limit) {
      var c = text.charAt(a - 1);
      if (c === '，' || c === ',' || c === '、') { break; }
      a -= 1;
    }
    var b = index + len;
    while (b < segEnd && b - (index + len) < limit) {
      var c2 = text.charAt(b);
      if (c2 === '，' || c2 === ',' || c2 === '、') { break; }
      b += 1;
    }
    return { start: a, end: b };
  }

  function joinWithout(text, start, end, cutStart, cutLen) {
    var before = text.slice(start, cutStart);
    var after = text.slice(cutStart + cutLen, end);
    return before + ' ' + after;
  }

  /** 清洗跟进备注：去日期、去首尾连接词、压缩空白、截 60 字 */
  function normalizeNote(value) {
    var t = collapseSpace(value);
    t = t.replace(/^[\s,，;；:：、\-–—.!！?？()（）]+/, '');
    t = t.replace(/[\s,，;；:、\-–—]+$/, '');
    t = t.replace(/\s+(?:on|at|in|of|for|about|with)$/i, '');
    t = t.replace(/(?:于|在|是|为)$/, '');
    t = t.replace(/^(?:on|at|in|about|for)\s+/i, '');
    t = t.replace(/^(?:于|在|是|为|对)/, '');
    t = collapseSpace(t);
    if (t.length > 60) { t = t.slice(0, 60); }
    return t;
  }

  function extractNote(norm, rawText, d) {
    var seg = expandSegment(norm, d.index, d.len);
    var clause = expandClause(norm, seg.start, seg.end, d.index, d.len);
    var note = normalizeNote(joinWithout(rawText, clause.start, clause.end, d.index, d.len));
    if (!note) {
      note = normalizeNote(joinWithout(rawText, seg.start, seg.end, d.index, d.len));
    }
    return note;
  }

  /**
   * 生成 followups：文中最新的「日期 + 事件」，倒序最多 3 条。
   * 成立日期不进入跟进记录；若还有其他日期，首次询盘日期也不重复放进跟进。
   */
  function buildFollowups(norm, rawText, dates, inquiryISO) {
    var pool = [];
    for (var i = 0; i < dates.length; i += 1) {
      if (!dates[i].founded) { pool.push(dates[i]); }
    }
    if (!pool.length) { return []; }
    var others = pool.filter(function (d) { return d.iso !== inquiryISO; });
    if (others.length) { pool = others; }
    var seen = {};
    var items = [];
    for (var k = 0; k < pool.length; k += 1) {
      var d = pool[k];
      if (seen[d.iso]) { continue; }
      seen[d.iso] = 1;
      items.push({ date: d.iso, note: extractNote(norm, rawText, d) });
    }
    items.sort(function (a, b) { return a.date < b.date ? 1 : (a.date > b.date ? -1 : 0); });
    return items.slice(0, 3);
  }

  /* ================================================================== */
  /* 4. 国家 / 客户名 / 公司名                                            */
  /* ================================================================== */

  /** 国家：schema.findCountry 优先，补充「(DE)」这类 ISO 代码写法 */
  function extractCountry(norm) {
    if (schema && typeof schema.findCountry === 'function') {
      var hit = schema.findCountry(norm);
      if (hit) {
        return { name: hit.name, index: hit.index, len: String(hit.term || '').length };
      }
      var m = /[(（]\s*([A-Za-z]{2})\s*[)）]/.exec(norm);
      if (m && Array.isArray(schema.COUNTRIES)) {
        var code = m[1].toUpperCase();
        for (var i = 0; i < schema.COUNTRIES.length; i += 1) {
          if (schema.COUNTRIES[i].iso === code) {
            return { name: schema.COUNTRIES[i].name, index: m.index, len: m[0].length };
          }
        }
      }
    }
    return null;
  }

  /* 英文人名 token：Michael / O'Brien / Jean-Luc */
  var EN_PERSON_TOKEN = "[A-Z][A-Za-z'\\u2019\\-]{1,20}";
  var EN_PERSON = EN_PERSON_TOKEN + "(?:\\s+" + EN_PERSON_TOKEN + "){0,2}";

  var PERSON_BAD_WORDS = /^(?:the|our|your|their|company|limited|ltd|llc|inc|corp|group|trading|import|export|industrial|technology|contact|team|sales|manager|director|purchasing|procurement|mr|mrs|ms|miss|dr|prof|germany|china|usa|uk|egypt|india|japan|korea|vietnam|thailand|brazil|mexico|canada|australia|january|february|march|april|may|june|july|august|september|october|november|december)$/i;
  var PERSON_BAD_CONTAIN = /(?:gmbh|ltd|llc|inc\.|corp\.|company|trading|holding|import|export|industrial|technology|\d)/i;

  function cleanEnPerson(name) {
    var s = collapseSpace(name).replace(/[.,;:]+$/, '');
    if (!s || s.length < 2 || s.length > 40) { return ''; }
    if (PERSON_BAD_CONTAIN.test(s)) { return ''; }
    var tokens = s.split(' ');
    if (tokens.length < 1 || tokens.length > 3) { return ''; }
    for (var i = 0; i < tokens.length; i += 1) {
      if (!/^[A-Z][A-Za-z'\u2019\-]+$/.test(tokens[i])) { return ''; }
      if (PERSON_BAD_WORDS.test(tokens[i])) { return ''; }
    }
    if (schema && typeof schema.findCountry === 'function' && schema.findCountry(s)) { return ''; }
    return s;
  }

  /* 常见姓氏：中文姓名必须首字是姓氏，显著降低「客户确认」这类误判 */
  var CN_SURNAMES = '王李张刘陈杨黄赵吴周徐孙马朱胡郭何高林罗郑梁谢宋唐许韩冯邓曹彭曾肖田董袁潘于蒋蔡余杜叶程苏魏吕丁任沈姚卢姜崔钟谭陆汪范金石廖贾夏韦付方白邹孟熊秦邱江尹薛闫段雷侯龙史陶黎贺顾毛郝龚邵万钱严覃武戴莫孔向汤常温康施文牛樊葛邢安齐易乔伍庞颜倪庄聂章鲁岳翟殷詹申欧耿关兰焦俞左柳甘祝包宁尚符舒阮柯纪梅童凌毕单季裴霍涂成苗谷盛曲翁冉骆蓝路游辛靳管柴蒙鲍华喻祁蒲房滕屈饶解牟艾尤阳时穆农司卓古吉缪简车项连芦麦褚娄窦戚岑景党宫费卜冷晏席卫米柏宗瞿桂全佟应臧闵苟邬边卞姬师和仇栾隋商刁沙荣巫寇桑郎甄丛仲虞敖巩明佘池查麻苑迟邝';
  var CN_NAME_BAD = /^(?:来自|位于|是|为|在|对|要求|希望|需要|反馈|表示|说|想|要|问|询|咨询|采购|购买|做|已|还|将|把|的|和|与|及|等|他们|我们|客户|客人|公司|企业|工厂|老板|经理|负责|联系|电话|邮件|地址|国家|产品|价格|订单|样品|报价|询盘|询价|确认|同意|接受|拒绝|回复|等待|考虑|比较|觉得|认为|沟通|询问|明天|后天|下周|本周|目前|现在)/;

  function cleanCnPerson(name) {
    var s = collapseSpace(name);
    if (!/^[\u4e00-\u9fa5]{2,4}$/.test(s)) { return ''; }
    if (CN_NAME_BAD.test(s)) { return ''; }
    if (CN_SURNAMES.indexOf(s.charAt(0)) === -1) { return ''; }
    s = s.replace(/(?:先生|女士|小姐|经理|总监|老师|总)$/, '');
    if (s.length < 2) { return ''; }
    return s;
  }

  /* 客户名识别规则（按优先级从上到下） */
  var CUSTOMER_RULES = [
    {
      id: 'customer.title',
      re: new RegExp('\\b(?:Mr|Mrs|Ms|Miss|Dr|Prof|Mister)\\.?\\s+(' + EN_PERSON + ')', 'g'),
      pick: function (m) {
        var value = cleanEnPerson(m[1]);
        if (!value) { return null; }
        return { value: value, index: m.index + m[0].indexOf(m[1]), len: m[1].length };
      }
    },
    {
      /* 自我介绍：this is Sarah Johnson / I am … / my name is … */
      id: 'customer.intro',
      re: new RegExp('(?:this\\s+is|i\\s+am|i\\u2019m|i\'m|my\\s+name\\s+is|myself)\\s+(' + EN_PERSON + ')', 'gi'),
      pick: function (m) {
        var value = cleanEnPerson(m[1]);
        if (!value) { return null; }
        return { value: value, index: m.index + m[0].indexOf(m[1]), len: m[1].length };
      }
    },
    {
      id: 'customer.label-en',
      re: new RegExp('(?:客户名称|客户|联系人|姓名|负责人|对接人|采购负责人|采购经理|客人|老板)\\s*(?:是|叫|名为|名字是|姓名是)?\\s*[：:]?\\s*(' + EN_PERSON + ')', 'g'),
      pick: function (m) {
        var value = cleanEnPerson(m[1]);
        if (!value) { return null; }
        return { value: value, index: m.index + m[0].indexOf(m[1]), len: m[1].length };
      }
    },
    {
      id: 'customer.label-cn',
      re: /(?:客户名称|联系人|姓名|负责人|对接人|采购经理|采购负责人|老板|客户|客人)\s*(?:是|叫|名为|名字是|姓名是|：|:)?\s*([\u4e00-\u9fa5]{2,4})(?![\u4e00-\u9fa5])/g,
      pick: function (m) {
        var value = cleanCnPerson(m[1]);
        if (!value) { return null; }
        return { value: value, index: m.index + m[0].indexOf(m[1]), len: m[1].length };
      }
    },
    {
      id: 'customer.label-en2',
      re: new RegExp('(?:contact\\s*(?:person|name)?|attn|attention)\\s*[:.：]?\\s*(' + EN_PERSON + ')', 'gi'),
      pick: function (m) {
        var value = cleanEnPerson(m[1]);
        if (!value) { return null; }
        return { value: value, index: m.index + m[0].indexOf(m[1]), len: m[1].length };
      }
    },
    {
      id: 'customer.name-from',
      re: new RegExp('(' + EN_PERSON + ')\\s+(?:from|of|at)\\s+[A-Z]', 'g'),
      pick: function (m) {
        var value = cleanEnPerson(m[1]);
        if (!value) { return null; }
        return { value: value, index: m.index, len: m[1].length };
      }
    }
  ];

  /* 英文公司后缀（含常见欧洲/美洲写法） */
  var EN_COMPANY_SUFFIX = "(?:GmbH|Co\\.?,?\\s*Ltd\\.?|Ltd\\.?|LLC|L\\.L\\.C\\.|Inc\\.?|Incorporated|Corporation|Corp\\.?|Limited|B\\.V\\.?|N\\.V\\.?|S\\.A\\.?|S\\.p\\.A\\.?|Pty\\.?\\s*Ltd\\.?|LLP|PLC|Group|Holdings?|Trading(?:\\s+Co\\.?)?|Industrial|Import\\s*&\\s*Export(?:\\s+Co\\.?,?\\s*Ltd\\.?)?|Imports?|Exports?|Srl|S\\.r\\.l\\.?|SAS|SARL|AG|AB|ApS|Oy|KG)";

  var EN_COMPANY_RE = new RegExp(
    "\\b([A-Z][A-Za-z&'\\u2019\\-]*(?:\\s+(?:[A-Z][A-Za-z&'\\u2019\\-]*|of|and|&|de|van|der|the)){0,6}\\s+" +
    EN_COMPANY_SUFFIX + ")(?![A-Za-z0-9])", 'g'
  );

  function cleanCompany(value) {
    var t = collapseSpace(value);
    t = t.replace(/^[（(【\[]+/, '').replace(/[）)】\]]+$/, '');
    t = t.replace(/^(?:the|company|our)\s+/i, '');
    t = t.replace(/[，。；;,、：:]+$/, '');
    if (!t || t.length < 3) { return ''; }
    if (t.length > 80) { t = t.slice(0, 80); }
    return t;
  }

  /** 从「公司名：XXX，位于…」里截出公司名部分 */
  function cutCompanyTail(value) {
    var t = String(value || '');
    var idx = t.search(/(?:位于|成立|主营|主要|从事|地址|电话|传真|邮箱|规模|员工|团队|年营|做(?:家居|外贸|进口|出口))/);
    if (idx >= 0) { t = t.slice(0, idx); }
    return t;
  }

  function looksLikeCompany(value) {
    return /(?:有限公司|有限责任公司|股份有限公司|集团|公司|贸易|实业|工厂|厂|Co\.?|Ltd|LLC|Inc|GmbH|Corp|Corporation|Limited|Group|Trading|Industrial|Import|Export)/i.test(value);
  }

  /* 公司名识别规则（按优先级从上到下） */
  var COMPANY_RULES = [
    {
      id: 'company.suffix-en',
      re: EN_COMPANY_RE,
      pick: function (m) {
        var value = cleanCompany(m[1]);
        if (!value) { return null; }
        return { value: value, index: m.index + m[0].indexOf(m[1]), len: m[1].length };
      }
    },
    {
      id: 'company.label-cn',
      re: /(?:公司名称|公司全称|公司名|公司是|公司为|该公司|他们公司|客户公司|企业名称|企业是|企业为)\s*[：:]?\s*([^，。；;,\n、]{2,50})/g,
      pick: function (m) {
        var value = cleanCompany(cutCompanyTail(m[1]));
        if (!value) { return null; }
        return { value: value, index: m.index + m[0].indexOf(m[1]), len: m[1].length };
      }
    },
    {
      id: 'company.label-cn2',
      re: /(?:^|[，,。；;、\s])公司\s*[：:]\s*([^，,。；;\n]{2,50})/g,
      pick: function (m) {
        var value = cleanCompany(cutCompanyTail(m[1]));
        if (!value) { return null; }
        return { value: value, index: m.index + m[0].indexOf(m[1]), len: m[1].length };
      }
    },
    {
      id: 'company.suffix-cn',
      re: /(?:^|[，。；;、\s：:（(])([A-Za-z0-9\u4e00-\u9fa5（）()·&\-]{2,30}(?:有限公司|有限责任公司|股份有限公司|集团有限公司|集团|贸易公司|实业公司|公司|贸易|实业|工厂|厂))(?=[，。；;、\s）)的]|$)/g,
      pick: function (m) {
        var value = cleanCompany(m[1]);
        if (!value || !looksLikeCompany(value)) { return null; }
        return { value: value, index: m.index + m[0].indexOf(m[1]), len: m[1].length };
      }
    }
  ];

  /* ================================================================== */
  /* 5. 通用「规则优先」执行器                                            */
  /* ================================================================== */

  /**
   * 依次尝试 rules，返回第一个通过 pick 校验的匹配结果。
   * rule: { id, re(global), pick(m, text) -> {value,index,len}|null }
   * 个别规则（如主营业务）需要看匹配位置前后的上下文，故把全文 text 传给 pick。
   */
  function pickFirstRule(rules, text) {
    for (var i = 0; i < rules.length; i += 1) {
      var rule = rules[i];
      var re = rule.re;
      re.lastIndex = 0;
      var m = re.exec(text);
      while (m !== null) {
        if (m[0].length === 0) { re.lastIndex += 1; m = re.exec(text); continue; }
        var hit = rule.pick(m, text);
        if (hit) {
          hit.rule = rule.id;
          return hit;
        }
        m = re.exec(text);
      }
    }
    return null;
  }

  /* ================================================================== */
  /* 6. 成立时间 / 团队规模 / 年营收                                      */
  /* ================================================================== */

  var FOUNDED_RULES = [
    {
      /* 成立于1998年 / 成立时间：2010 / 创立于 2005 年 3 月 */
      id: 'founded.cn-label',
      re: /(?:成立时间|成立年份|成立于|成立|创立于|创办于|始创于|创立|创办|始创|建于|建立于)\s*(?:于|时间|年份|是|为|：|:)?\s*(?:在)?\s*(\d{4})\s*年?\s*(?:(\d{1,2})\s*月)?/g,
      pick: function (m) {
        var year = Number(m[1]);
        var month = Number(m[2]);
        var value = String(year);
        if (month >= 1 && month <= 12) { value = year + '-' + pad2(month); }
        return { value: value, index: m.index, len: m[0].length };
      }
    },
    {
      /* 1998年成立 / 2011 年成立 / 1998年3月创立 */
      id: 'founded.cn-reverse',
      re: /(\d{4})\s*年\s*(?:(\d{1,2})\s*月\s*)?(?:成立|创立|创办|始创|建成)/g,
      pick: function (m) {
        var year = Number(m[1]);
        var month = Number(m[2]);
        var value = String(year);
        if (month >= 1 && month <= 12) { value = year + '-' + pad2(month); }
        return { value: value, index: m.index, len: m[0].length };
      }
    },
    {
      /* founded in 1998 / established in 2010 / since 1994 */
      id: 'founded.en',
      re: /(?:founded|established|incorporated|in\s+business\s+since|operating\s+since)\s*(?:in\s+|since\s+|back\s+in\s+|as\s+early\s+as\s+)?(\d{4})/gi,
      pick: function (m) {
        return { value: String(Number(m[1])), index: m.index, len: m[0].length };
      }
    }
  ];

  /** 「有 20 年历史 / 30 years in business」→ 反推成立年份（排除「我方」口吻） */
  function extractFoundedByAge(norm, fallbackYear) {
    var rules = [
      { re: /(\d{1,2})\s*年(?:的)?(?:历史|经营历史)/g, years: function (m) { return Number(m[1]); } },
      { re: /(\d{1,2})\s+years?\s+(?:in\s+business|of\s+experience)/gi, years: function (m) { return Number(m[1]); } }
    ];
    for (var r = 0; r < rules.length; r += 1) {
      var re = rules[r].re;
      re.lastIndex = 0;
      var m = re.exec(norm);
      while (m !== null) {
        var before = norm.slice(Math.max(0, m.index - 16), m.index);
        var mine = /我们|我司|我公司|本公司|we\s|our\s|\bour$/i.test(before);
        var years = rules[r].years(m);
        var year = fallbackYear - years;
        if (!mine && years > 0 && years < 120 && year >= 1900 && year <= 2100) {
          return { value: String(year), index: m.index, len: m[0].length, rule: 'founded.age' };
        }
        m = re.exec(norm);
      }
    }
    return null;
  }

  function extractFoundedAt(norm, fallbackYear) {
    var hit = pickFirstRule(FOUNDED_RULES, norm);
    if (hit) { return hit; }
    return extractFoundedByAge(norm, fallbackYear);
  }

  var TEAM_RULES = [
    {
      id: 'team.cn',
      re: /(?:团队规模|团队人数|员工人数|员工总数|人员规模|团队|员工|职工|雇员|规模)\s*(?:规模)?\s*(?:约|大约|近|超过|超|有|为|是|达到|在|：|:)?\s*(\d[\d,]*(?:\s*[-~～至到–—]\s*\d[\d,]*)?)\s*(?:多|余|左右)?\s*(?:人|名员工|名|位|个)/g,
      pick: function (m) {
        return { value: formatTeamRange(m[1]), index: m.index, len: m[0].length };
      }
    },
    {
      id: 'team.en',
      re: /team\s+(?:of\s+|size\s+of\s+|of\s+)?(?:around\s+|about\s+|approximately\s+)?(\d[\d,]*(?:\s*[-–~]\s*\d[\d,]*)?)\s*(?:employees|staff|people|persons|members)?/gi,
      pick: function (m) {
        return { value: formatTeamRange(m[1]), index: m.index, len: m[0].length };
      }
    },
    {
      id: 'team.en-plain',
      re: /(\d[\d,]*(?:\s*[-–~]\s*\d[\d,]*)?)\s+(?:employees|staff|people|persons)/gi,
      pick: function (m) {
        return { value: formatTeamRange(m[1]), index: m.index, len: m[0].length };
      }
    }
  ];

  function formatTeamRange(raw) {
    var n = String(raw || '').replace(/,/g, '').replace(/\s*[-~～至到–—]\s*/g, '-').replace(/\s+/g, '');
    if (!/\d/.test(n)) { return ''; }
    return n + ' 人';
  }

  var CURRENCY_MAP = {
    eur: '欧元', usd: '美元', gbp: '英镑', rmb: '人民币', cny: '人民币',
    jpy: '日元', hkd: '港元', aud: '澳元', cad: '加元', chf: '瑞士法郎',
    '€': '欧元', '£': '英镑', '$': '美元',
    '美元': '美元', '欧元': '欧元', '英镑': '英镑', '人民币': '人民币', '日元': '日元',
    '港元': '港元', '澳元': '澳元', '加元': '加元', '元': '元'
  };

  function trimNumber(n) {
    var v = Math.round(Number(n) * 100) / 100;
    return String(v);
  }

  /** 金额归一化：EUR 12 million → 1200 万欧元；300 万美元 → 300 万美元 */
  function formatMoney(amountText, scaleText, currencyText) {
    var amount = Number(String(amountText || '').replace(/,/g, ''));
    if (!isFinite(amount) || amount <= 0) { return ''; }
    var cur = CURRENCY_MAP[String(currencyText || '').trim().toLowerCase()] || '';
    var scale = String(scaleText || '').trim().toLowerCase();
    var wan = null;
    if (scale === 'million' || scale === 'm') { wan = amount * 100; }
    else if (scale === 'billion' || scale === 'bn' || scale === 'b') { wan = amount * 100000; }
    else if (scale === 'k') { wan = amount / 10; }
    else if (scale === '万') { wan = amount; }
    else if (scale === '亿') { wan = amount * 10000; }
    if (wan !== null && isFinite(wan) && wan > 0) {
      var text;
      if (wan >= 10000 && Math.abs(wan / 10000 - Math.round(wan / 10000)) < 1e-9) {
        text = trimNumber(wan / 10000) + ' 亿';
      } else {
        text = trimNumber(wan) + ' 万';
      }
      return (text + cur).trim();
    }
    return (trimNumber(amount) + (cur ? ' ' + cur : '')).trim();
  }

  var REVENUE_RULES = [
    {
      id: 'revenue.cn',
      re: /(?:年营收|年营业额|年销售额|年采购额|年收入|营业额|营收规模|销售规模|营收|销售额|采购额)\s*(?:约为|约|大约|近|超过|超|达|达到|有|为|是|：|:)?\s*(\d[\d.,]*)\s*(万|亿)?\s*(美元|欧元|英镑|人民币|日元|港元|澳元|加元|元)?/g,
      pick: function (m) {
        var value = formatMoney(m[1], m[2], m[3]);
        if (!value) { return null; }
        return { value: value, index: m.index, len: m[0].length };
      }
    },
    {
      id: 'revenue.en',
      re: /(?:annual\s+(?:revenue|turnover|sales)|(?:yearly|annual)\s+revenue|turnover|revenue|sales)\s*(?:of|around|about|approximately|approx\.?|is|was|~)?\s*(?:(EUR|USD|GBP|RMB|CNY|JPY|HKD|AUD|CAD|CHF|€|£|\$)\s*)?(\d[\d,.]*)\s*(million|billion|bn|m|b|k)?\s*(?:(EUR|USD|GBP|RMB|CNY|JPY|HKD|AUD|CAD|CHF|€|£|\$))?/gi,
      pick: function (m) {
        var value = formatMoney(m[2], m[3], m[4] || m[1]);
        if (!value) { return null; }
        return { value: value, index: m.index, len: m[0].length };
      }
    }
  ];

  /* ================================================================== */
  /* 7. 咨询产品                                                         */
  /* ================================================================== */

  function cleanProduct(value) {
    var t = collapseSpace(value);
    t = t.replace(/^(?:我们的|我司的|我公司的|我们|我司|本公司|这款|该款|一款|一批|一个|一套|the|our|a|an)\s*/i, '');
    t = t.replace(/^(?:采购|购买|订购|询购|求购|需要)\s*/, '');
    t = t.replace(/\s*(?:的)?(?:价格|报价单|报价|价钱|quotation|price|quote)\s*$/i, '');
    t = t.replace(/[.，。；;,、：:！!？?]+$/g, '').trim();
    if (!t) { return ''; }
    if (/(?:有限公司|有限责任公司|集团|GmbH|Co\.,?\s*Ltd|LLC|Inc\.?|Corp\.?)/i.test(t)) { return ''; }
    if (t.length > 60) { t = t.slice(0, 60); }
    return t;
  }

  var PRODUCT_RULES = [
    {
      /* 客户咨询的是… / 询价产品：… / 询问… （采购不放这里，避免「年采购额」误判） */
      id: 'product.cn-label',
      re: /(?:咨询|询价|求购|询问)(?:的)?(?:产品|商品)?\s*(?:是|为|：|:)?\s*([^，,。；;、\n！!？?]{2,60})/g,
      pick: function (m) {
        var value = cleanProduct(m[1]);
        if (!value) { return null; }
        return { value: value, index: m.index + m[0].indexOf(m[1]), len: m[1].length };
      }
    },
    {
      /* 对我们的…感兴趣 */
      id: 'product.cn-interest',
      re: /对\s*(?:我们的|我司的|贵司的|这款|该款)?\s*([^，,。；;、\n！!？?]{2,40}?)\s*(?:非常|很|比较|十分)?(?:感)?兴趣/g,
      pick: function (m) {
        var value = cleanProduct(m[1]);
        if (!value) { return null; }
        return { value: value, index: m.index + m[0].indexOf(m[1]), len: m[1].length };
      }
    },
    {
      /* 需要 / 采购 / 想买 … */
      id: 'product.cn-need',
      re: /(?:需要|采购(?!额|商|部|量|金额|计划)|购买|想买|要买|订购|询购)\s*(?:的)?\s*([^，,。；;、\n！!？?]{2,60})/g,
      pick: function (m) {
        var value = cleanProduct(m[1]);
        if (!value) { return null; }
        return { value: value, index: m.index + m[0].indexOf(m[1]), len: m[1].length };
      }
    },
    {
      /* 产品：… / 询价产品：… */
      id: 'product.cn-colon',
      re: /(?:询价产品|咨询产品|产品名称|产品)\s*[：:]\s*([^，,。；;、\n！!？?]{2,60})/g,
      pick: function (m) {
        var value = cleanProduct(m[1]);
        if (!value) { return null; }
        return { value: value, index: m.index + m[0].indexOf(m[1]), len: m[1].length };
      }
    },
    {
      /* 英文：regarding our … / inquiry about … / interested in … / please quote … */
      id: 'product.en',
      re: /(?:regarding|inquir(?:y|ies)\s+(?:about|for|on)|enquir(?:y|ies)\s+(?:about|for|on)|interested\s+in|asking\s+about|asking\s+for|asked\s+about|asked\s+for|quotation\s+for|quote\s+for|please\s+quote|looking\s+for|wrote\s+about|emailed\s+(?:us\s+)?about)\s+(?:our|the|a|an|your)?\s*((?:[^。;,!\n?.]|\.(?!\s|$)){2,80})/gi,
      pick: function (m) {
        var value = cleanProduct(m[1]);
        if (!value) { return null; }
        return { value: value, index: m.index + m[0].indexOf(m[1]), len: m[1].length };
      }
    }
  ];

  /* ================================================================== */
  /* 7b. 主营业务 / 产品型号（v1.1 新增字段）                              */
  /* ================================================================== */

  /* ---------- 7b.1 主营业务 mainBusiness ---------- */

  /* 主语是「我们/同行/竞争对手」时，说的不是客户的主营业务，直接跳过 */
  var BUSINESS_NOISE_RE = /(?:我们|我司|我公司|本公司|我方|我厂|竞争对手|同行|另一家|其他供应商|竞品|对手)/;

  /** 找到匹配位置所在「分句」的起点（用于判断分句主语） */
  function clauseStartIndex(text, index) {
    var i = index - 1;
    while (i >= 0) {
      var c = text.charAt(i);
      if ('。；;，,\n\r！!？?'.indexOf(c) !== -1) { break; }
      i -= 1;
    }
    return i + 1;
  }

  function businessNoise(text, index, matchText) {
    var start = clauseStartIndex(text, index);
    if (BUSINESS_NOISE_RE.test(text.slice(start, index))) { return true; }
    /* 「主营产品/主营商品」多指我方产品，不是客户业务 */
    if (/(?:主营|主要经营|主要业务|主要从事于|主要从事|业务范围|经营范围|专营)\s*(?:产品|商品|品类)/.test(matchText)) { return true; }
    if (/^(?:做|从事|经营)(?:好|得|完|成)/.test(matchText)) { return true; }
    return false;
  }

  /** 清洗主营业务：去连接词/主语前缀、去尾缀，≤30 字 */
  function cleanBusiness(value) {
    var t = collapseSpace(value);
    var lead = /^(?:[\s:：,，。;；、\-–—]+|即|就是|是|为|一家|一个|的|我们|我司|我公司|本公司|客户|他们|其|该公司|公司)+/;
    for (var i = 0; i < 4 && lead.test(t); i += 1) { t = t.replace(lead, ''); }
    t = t.replace(/[\s:：,，。;；、]+$/, '');
    t = t.replace(/(?:的公司|的企业|等业务|业务|的公司)$/, '');
    if (!t) { return ''; }
    if (t.length > 30) {
      t = t.slice(0, 30);
      if (/^[A-Za-z]/.test(t)) {
        var space = t.lastIndexOf(' ');
        if (space > 12) { t = t.slice(0, space); }
      }
    }
    /* 截断后可能留下悬空的连接词，去掉 */
    t = t.replace(/\s+(?:and|or|of|for|with|the|a|an|to|in|&)$/i, '');
    t = collapseSpace(t);
    return t.length >= 2 ? t : '';
  }

  var BUSINESS_ROLE_MAP = {
    '进口商': '进口', '出口商': '出口', '批发商': '批发', '分销商': '分销',
    '零售商': '零售', '代理商': '代理', '经销商': '经销'
  };

  var BUSINESS_RULES = [
    {
      /* 主营户外用品 / 主要从事家居用品 / 业务范围：厨房用品 */
      id: 'business.cn-label',
      re: /(?:主营业务|主营|主要经营|主要业务|主要从事于|主要从事|业务范围|经营范围|专营)\s*(?:业务|范围|是|为|有|：|:)?\s*([^，,。；;、\n！!？?]{2,40})/g,
      pick: function (m, text) {
        if (businessNoise(text, m.index, m[0])) { return null; }
        var value = cleanBusiness(m[1]);
        if (!value) { return null; }
        return { value: value, index: m.index, len: m[0].length };
      }
    },
    {
      /* 做家居用品进口 / 从事户外用品批发 / 经营五金分销 */
      id: 'business.cn-trade',
      re: /(?:做|从事|经营|主营)(?!状况|情况|困难)(?:的是|着)?\s*([^，,。；;、\n！!？?]{2,30}?)\s*(进口|出口|批发|分销|零售|贸易|代理|经销)/g,
      pick: function (m, text) {
        if (businessNoise(text, m.index, m[0])) { return null; }
        var value = cleanBusiness(m[1] + m[2]);
        if (!value || value.length < 3) { return null; }
        return { value: value, index: m.index, len: m[0].length };
      }
    },
    {
      /* 是一家家居用品进口商 / 户外用品分销商 */
      id: 'business.cn-role',
      re: /([\u4e00-\u9fa5A-Za-z]{2,20})(?:的)?(进口商|出口商|批发商|分销商|零售商|代理商|经销商)/g,
      pick: function (m, text) {
        if (businessNoise(text, m.index, m[0])) { return null; }
        var value = cleanBusiness(m[1] + (BUSINESS_ROLE_MAP[m[2]] || ''));
        if (!value || value.length < 3) { return null; }
        return { value: value, index: m.index, len: m[0].length };
      }
    },
    {
      /* English: specialized in / mainly deals in / distributor of / importer of / wholesaler of */
      id: 'business.en',
      re: /(?:speciali[sz](?:e|es|ed|ing)\s+in|mainly\s+(?:deals?|trades?)\s+in|deals?\s+mainly\s+in|distributor\s+of|importer\s+of|wholesaler\s+of|retailer\s+of|trading\s+company\s+(?:of|speciali[sz]ing\s+in)|imports?\s+and\s+distributes|focus(?:es|ed)?\s+on)\s+(?:the\s+)?([^.;,\n]{2,60})/gi,
      pick: function (m) {
        var value = cleanBusiness(m[1]);
        if (!value) { return null; }
        return { value: value, index: m.index, len: m[0].length };
      }
    }
  ];

  function extractMainBusiness(norm) {
    return pickFirstRule(BUSINESS_RULES, norm);
  }

  /* ---------- 7b.2 产品型号 productModel ---------- */

  /* 规格模式：容量 / 件套 / pcs / 尺寸 / 套装 */
  var MODEL_PATTERNS = [
    { id: 'model.capacity', re: /(\d+(?:\.\d+)?)\s*(ml|毫升|l|升)(?![\w])/gi },
    { id: 'model.count', re: /(\d+)\s*(件套|件套装|件装|pcs\.?|pieces?|pc|套)(?![\w])/gi },
    { id: 'model.set', re: /(\d+)\s*-\s*(?:pieces?|pcs?|pc)\s+set/gi },
    { id: 'model.size', re: /(\d+(?:\.\d+)?)\s*[×xX*]\s*(\d+(?:\.\d+)?)\s*(cm|mm|厘米|毫米|inch|英寸)?/g }
  ];

  /* 显式标注的型号（型号/规格/model/SKU…），产品短语里没规格时兜底 */
  var MODEL_LABEL_RE = /(?:产品型号|型号|规格|model(?:\s*(?:no\.?|number))?|item\s*(?:no\.?|code)|sku|货号|款号)\s*[：:#]?\s*([A-Za-z0-9][A-Za-z0-9\-\/\.]{0,24})/gi;

  function normalizeModelToken(m, kind) {
    if (kind === 'model.capacity') {
      var unit = String(m[2]).toLowerCase();
      return String(m[1]) + (unit === 'ml' || unit === '毫升' ? 'ml' : 'L');
    }
    if (kind === 'model.count') {
      var u = String(m[2]).toLowerCase().replace(/\.$/, '');
      var n = String(m[1]);
      if (u === 'pcs' || u === 'pc' || u === 'piece' || u === 'pieces') { return n + ' pcs'; }
      if (u === '件套' || u === '件套装') { return n + ' 件套'; }
      if (u === '件装') { return n + ' 件装'; }
      return n + ' 套';
    }
    if (kind === 'model.set') { return String(m[1]) + '-piece set'; }
    if (kind === 'model.size') {
      var dimUnit = m[3] ? String(m[3]).toLowerCase() : '';
      if (dimUnit === '英寸') { dimUnit = 'inch'; }
      return String(m[1]) + '×' + String(m[2]) + dimUnit;
    }
    return '';
  }

  /** 从一段产品短语里收集规格 token（按出现顺序、去重） */
  function collectModelTokens(text) {
    if (!text) { return []; }
    var found = [];
    for (var p = 0; p < MODEL_PATTERNS.length; p += 1) {
      var def = MODEL_PATTERNS[p];
      def.re.lastIndex = 0;
      var m = def.re.exec(text);
      while (m !== null) {
        if (m[0].length === 0) { def.re.lastIndex += 1; m = def.re.exec(text); continue; }
        var token = normalizeModelToken(m, def.id);
        if (token) { found.push({ token: token, index: m.index }); }
        m = def.re.exec(text);
      }
    }
    found.sort(function (a, b) { return a.index - b.index; });
    var out = [];
    var seen = {};
    for (var i = 0; i < found.length; i += 1) {
      if (seen[found[i].token]) { continue; }
      seen[found[i].token] = 1;
      out.push(found[i]);
    }
    return out;
  }

  /** 在原文中定位型号的位置（evidence 用；空格/大小写做宽松匹配） */
  function locateModel(norm, token) {
    var idx = norm.indexOf(token);
    if (idx !== -1) { return { index: idx, len: token.length }; }
    var escaped = token.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s*');
    var m = new RegExp(escaped, 'i').exec(norm);
    if (m) { return { index: m.index, len: m[0].length }; }
    return null;
  }

  /**
   * 产品型号：优先从已识别的产品短语中抽规格；
   * 抽不到时再看全文里「型号：XXX」这类显式标注；都没有则不输出。
   * productHit 为产品字段的匹配位置，用于 evidence 兜底。
   */
  function extractProductModel(norm, productText, productHit) {
    var tokens = collectModelTokens(productText);
    if (tokens.length) {
      var values = [];
      var total = 0;
      for (var i = 0; i < tokens.length && values.length < 3; i += 1) {
        var token = tokens[i].token;
        if (total + token.length + (values.length ? 1 : 0) > 30) { break; }
        values.push(token);
        total += token.length + 1;
      }
      if (values.length) {
        var found = locateModel(norm, values[0]);
        if (!found && productHit) { found = { index: productHit.index, len: productHit.len }; }
        if (found) {
          return { value: values.join(' '), index: found.index, len: found.len, rule: 'model.product' };
        }
      }
    }
    MODEL_LABEL_RE.lastIndex = 0;
    var m = MODEL_LABEL_RE.exec(norm);
    while (m !== null) {
      if (m[0].length === 0) { MODEL_LABEL_RE.lastIndex += 1; m = MODEL_LABEL_RE.exec(norm); continue; }
      var label = String(m[1] || '').replace(/[.\-_/]+$/, '');
      if (label.length >= 2 && /\d/.test(label)) {
        return { value: label, index: m.index + m[0].indexOf(m[1]), len: m[1].length, rule: 'model.label' };
      }
      m = MODEL_LABEL_RE.exec(norm);
    }
    return null;
  }

  /* ================================================================== */
  /* 8. 跟进状态 / 成交状态                                               */
  /* ================================================================== */

  var STATUS_RULES = [
    {
      /* 已成交 / 下单 / 签约 / PO → 已成交（最高优先级） */
      id: 'status.deal',
      re: /成交|下单|签约|签合同|签订|已付定金|已付款|打款|placed\s+(?:an?\s+)?order|order\s+(?:placed|confirmed)|deal\s+(?:was\s+|is\s+)?(?:closed|won)|closed\s+(?:the\s+)?deal|signed\s+(?:the\s+)?(?:contract|order|p\.?o\.?)|won\s+the\s+deal|\bP\.?O\.?\b/gi,
      pick: function (m) { return { value: '已成交', index: m.index, len: m[0].length }; }
    },
    {
      /* 报价 / quotation → 已报价 */
      id: 'status.quote',
      re: /已报价|报价|报盘|发盘|quotation|quoted|\bquote[sd]?\b|price\s+(?:list|sheet)|报价单/gi,
      pick: function (m) { return { value: '已报价', index: m.index, len: m[0].length }; }
    },
    {
      /* 样品 / 样册 → 寄样中 */
      id: 'status.sample',
      re: /寄样|样品|样册|样板|打样|寄出样|\bsamples?\b|sent\s+(?:the\s+)?samples?|catalog(?:ue)?/gi,
      pick: function (m) { return { value: '寄样中', index: m.index, len: m[0].length }; }
    },
    {
      /* 谈判 / 合同条款 → 谈判中 */
      id: 'status.negotiate',
      re: /谈判|磋商|协商|合同条款|合同细节|contract\s+terms|negotiat/gi,
      pick: function (m) { return { value: '谈判中', index: m.index, len: m[0].length }; }
    },
    {
      /* 流失 / 放弃 → 已流失 */
      id: 'status.lost',
      re: /流失|放弃|丢单|不再联系|失去联系|lost\s+(?:the\s+)?(?:deal|customer)|\bnot\s+interested\b|no\s+longer\s+interested/gi,
      pick: function (m) { return { value: '已流失', index: m.index, len: m[0].length }; }
    },
    {
      /* 暂停 / on hold → 暂停跟进 */
      id: 'status.pause',
      re: /暂停|搁置|暂缓|暂不|on\s+hold|hold\s+off|put\s+on\s+hold/gi,
      pick: function (m) { return { value: '暂停跟进', index: m.index, len: m[0].length }; }
    },
    {
      /* 已回复 / replied → 已回复（「等回复」不算已回复） */
      id: 'status.replied',
      re: /已回复|已回信|已回|回复了|回复邮件|回信|replied|we\s+sent\s+(?:an?\s+)?(?:email|reply)|email\s+sent|responded/gi,
      pick: function (m) { return { value: '已回复', index: m.index, len: m[0].length }; }
    }
  ];

  /* 成交状态判断用非全局正则（只需第一次匹配，避免 lastIndex 跨调用污染） */
  var DEAL_NEG_RE = /未成交|尚未成交|还没成交|没有成交|未下单|还没下单|没有下单|未签约|未达成|not\s+(?:yet\s+)?closed|hasn'?t\s+(?:been\s+)?closed|haven'?t\s+(?:been\s+)?closed|isn'?t\s+closed|no\s+deal|not\s+a\s+(?:closed\s+)?deal|no\s+order|didn'?t\s+order|not\s+ordered\s+yet|待成交/i;
  var DEAL_POS_RE = /成交|下单|签约|签订|已付定金|已付款|placed\s+(?:an?\s+)?order|order\s+(?:placed|confirmed)|deal\s+(?:was\s+|is\s+)?(?:closed|won)|closed\s+(?:the\s+)?deal|signed\s+(?:the\s+)?(?:contract|order|po)|won\s+the\s+deal|\bpo\s+(?:issued|received)\b/i;
  var DEAL_PENDING_RE = /仍在等|还在等|等待|待确认|等确认|pending|awaiting|waiting\s+for|to\s+be\s+confirmed|comparing|比较供应商/i;

  /**
   * 成交状态：明确成交 → 已成交；明确未成交/待确认 → 未成交；
   * 已经有报价/寄样/谈判等进度但无成交词时视为未成交（契约样例 A 即此情况）。
   */
  function extractDealStatus(norm, statusValue) {
    var m = DEAL_NEG_RE.exec(norm);
    if (m) { return { value: '未成交', index: m.index, len: m[0].length, rule: 'deal.negative' }; }
    m = DEAL_POS_RE.exec(norm);
    if (m) { return { value: '已成交', index: m.index, len: m[0].length, rule: 'deal.positive' }; }
    m = DEAL_PENDING_RE.exec(norm);
    if (m) { return { value: '未成交', index: m.index, len: m[0].length, rule: 'deal.pending' }; }
    /* 其余情况不输出，界面按默认「未成交」处理（Lead 联调约定） */
    return null;
  }

  /* ================================================================== */
  /* 9. 核心原因                                                         */
  /* ================================================================== */

  var REASON_SUBJECTS = [
    { re: /价格|价位|单价|报价|price|pricing/i, w: 2 },
    { re: /目标价|target\s*price|budget|\btarget\b|心理价/i, w: 1 },
    { re: /交期|交货|货期|lead\s*time|delivery/i, w: 2 },
    { re: /付款|支付|账期|定金|deposit|payment|prepay/i, w: 2 },
    { re: /质量|品质|quality/i, w: 2 },
    { re: /认证|证书|certif/i, w: 2 },
    { re: /运费|物流|freight|shipping|logistics/i, w: 2 },
    { re: /竞争|对手|其他供应商|competitor|alternative|comparing/i, w: 2 },
    { re: /起订量|moq|数量|quantity/i, w: 1 },
    { re: /折扣|discount/i, w: 1 },
    { re: /成本|cost|expensive|too\s+high/i, w: 1 }
  ];

  var REASON_CUES = [
    { re: /原因|因为|由于|主要是|卡在|顾虑|担心|反馈|问题在于|争议|原因在于/i, w: 1 },
    { re: /太长|太慢|太高|太低|太贵|太远|不满意|不接受|再考虑|考虑一下|观望|比较过|comparing/i, w: 1 },
    { re: /because|due\s+to|reason|concern|issue|problem|worried|however|said|mentioned/i, w: 1 }
  ];

  /**
   * 原因句打分：至少要「命中 2 类关键词」或「有明确原因语气词且总分 ≥3」才算数，
   * 避免把「发送报价」这类普通跟进事件误当核心原因。
   */
  function scoreReason(segment) {
    var subjectScore = 0;
    var kinds = 0;
    var cueScore = 0;
    var i;
    for (i = 0; i < REASON_SUBJECTS.length; i += 1) {
      if (REASON_SUBJECTS[i].re.test(segment)) {
        subjectScore += REASON_SUBJECTS[i].w;
        kinds += 1;
      }
    }
    for (i = 0; i < REASON_CUES.length; i += 1) {
      if (REASON_CUES[i].re.test(segment)) { cueScore += REASON_CUES[i].w; }
    }
    var score = subjectScore + cueScore;
    var accept = (kinds >= 2) || (cueScore > 0 && score >= 3);
    return { score: score, accept: accept, kinds: kinds, cues: cueScore };
  }

  /** 把全文切成句子片段（保留原文下标，供 evidence 使用） */
  function listSegments(text) {
    var list = [];
    var start = 0;
    var i = 0;
    while (i < text.length) {
      if (isSentenceStop(text, i)) {
        if (i > start) { list.push({ start: start, end: i }); }
        start = i + 1;
      }
      i += 1;
    }
    if (start < text.length) { list.push({ start: start, end: text.length }); }
    return list;
  }

  /** 把句子按逗号切成小分句（原因通常落在某个分句里） */
  function listClauses(text, start, end) {
    var list = [];
    var s = start;
    for (var i = start; i < end; i += 1) {
      var c = text.charAt(i);
      if (c === '，' || c === ',') {
        if (i > s) { list.push({ start: s, end: i }); }
        s = i + 1;
      }
    }
    if (s < end) { list.push({ start: s, end: end }); }
    return list;
  }

  function cleanReason(value) {
    var t = collapseSpace(value);
    t = t.replace(/^[\s,，。；;:：、\-–—]+/, '');
    t = t.replace(/[\s,，。；;:：、]+$/, '');
    if (t.length > 80) { t = t.slice(0, 80); }
    return t;
  }

  /** 从候选片段里挑分数最高且通过 accept 的一个（同分取更靠前的） */
  function pickReasonCandidate(norm, ranges) {
    var best = null;
    for (var i = 0; i < ranges.length; i += 1) {
      var r = ranges[i];
      var len = r.end - r.start;
      if (len < 4 || len > 260) { continue; }
      var text = norm.slice(r.start, r.end);
      var result = scoreReason(text);
      if (result.accept && (!best || result.score > best.score)) {
        best = { score: result.score, start: r.start, len: len, value: cleanReason(text) };
      }
    }
    return best;
  }

  /**
   * 核心原因：优先带价格/交期/付款等原因关键词的「小分句」，
   * 分句都没命中时再退化到整句。
   */
  function extractCoreReason(norm) {
    var segs = listSegments(norm);
    var clauses = [];
    for (var i = 0; i < segs.length; i += 1) {
      clauses = clauses.concat(listClauses(norm, segs[i].start, segs[i].end));
    }
    var best = pickReasonCandidate(norm, clauses);
    if (!best) { best = pickReasonCandidate(norm, segs); }
    if (!best || !best.value) { return null; }
    return { value: best.value, index: best.start, len: best.len, rule: 'reason.clause' };
  }

  /* ================================================================== */
  /* 10. 结果装配                                                         */
  /* ================================================================== */

  function setField(ctx, key, value, index, len, rule) {
    if (value === null || value === undefined) { return false; }
    var text = String(value).trim();
    if (!text) { return false; }
    ctx.fields[key] = text;
    if (typeof index === 'number' && index >= 0) {
      ctx.evidence[key] = { value: text, snippet: makeSnippet(ctx.raw, index, len) };
    }
    if (rule) { uniqPush(ctx.rules, rule); }
    return true;
  }

  function emptyResult() {
    return {
      ok: false,
      fields: {},
      followups: [],
      evidence: {},
      warnings: [],
      filled: 0,
      missing: FIELD_KEYS.slice(),
      raw: { textLength: 0, countryHit: '', matchedRules: [] }
    };
  }

  /**
   * 识别入口。
   * @param {string} text 询盘原文
   * @param {{today?:string, defaultYear?:number}} [opts]
   * @returns 识别结果（见契约 §5.1）
   */
  function recognize(text, opts) {
    var out = emptyResult();
    try {
      var input = (text === null || text === undefined) ? '' : String(text);
      out.raw.textLength = input.length;
      if (!input.trim()) {
        out.warnings.push({ code: 'empty', message: '请粘贴一段询盘文本' });
        return out;
      }

      var truncated = false;
      var body = input;
      if (body.length > MAX_TEXT_LENGTH) {
        body = body.slice(0, MAX_TEXT_LENGTH);
        truncated = true;
      }

      var norm = toHalfWidth(body);
      var fallbackYear = resolveFallbackYear(opts);
      var ctx = { raw: body, fields: {}, evidence: {}, rules: [] };

      /* 1) 日期候选与首次询盘时间 */
      var dates = collectDates(norm, fallbackYear);
      tagDates(norm, dates);
      var inquiry = pickInquiryDate(dates);
      if (inquiry) {
        setField(ctx, 'firstInquiryAt', inquiry.iso, inquiry.index, inquiry.len, inquiry.rule);
      }

      /* 2) 国家 */
      var country = extractCountry(norm);
      if (country) {
        setField(ctx, 'country', country.name, country.index, country.len, 'country.schema');
        out.raw.countryHit = country.name;
      }

      /* 3) 客户名（联系人） */
      var person = pickFirstRule(CUSTOMER_RULES, norm);
      if (person) { setField(ctx, 'customerName', person.value, person.index, person.len, person.rule); }

      /* 4) 公司名 */
      var company = pickFirstRule(COMPANY_RULES, norm);
      if (company) { setField(ctx, 'companyName', company.value, company.index, company.len, company.rule); }

      /* 5) 成立时间 */
      var founded = extractFoundedAt(norm, fallbackYear);
      if (founded) { setField(ctx, 'foundedAt', founded.value, founded.index, founded.len, founded.rule); }

      /* 6) 团队规模 */
      var team = pickFirstRule(TEAM_RULES, norm);
      if (team) { setField(ctx, 'teamSize', team.value, team.index, team.len, team.rule); }

      /* 7) 年营收 */
      var revenue = pickFirstRule(REVENUE_RULES, norm);
      if (revenue) { setField(ctx, 'annualRevenue', revenue.value, revenue.index, revenue.len, revenue.rule); }

      /* 8) 咨询产品 */
      var product = pickFirstRule(PRODUCT_RULES, norm);
      if (product) { setField(ctx, 'product', product.value, product.index, product.len, product.rule); }

      /* 8b) 产品型号（v1.1：从产品短语抽规格，无则不出） */
      var model = extractProductModel(norm, product ? product.value : '', product);
      if (model) {
        setField(ctx, 'productModel', model.value, model.index, model.len, model.rule);
      }

      /* 8c) 主营业务（v1.1） */
      var business = extractMainBusiness(norm);
      if (business) {
        setField(ctx, 'mainBusiness', business.value, business.index, business.len, business.rule);
      }

      /* 9) 跟进状态：识别不到则按契约给默认「新询盘」 */
      var status = pickFirstRule(STATUS_RULES, norm);
      var statusValue = status ? status.value : '新询盘';
      if (status) {
        setField(ctx, 'followupStatus', statusValue, status.index, status.len, status.rule);
      } else {
        ctx.fields.followupStatus = statusValue;
        /* 默认值没有命中规则，用开头原文作为对照依据（不编造内容） */
        ctx.evidence.followupStatus = { value: statusValue, snippet: makeSnippet(body, 0, 0) };
      }

      /* 10) 成交状态 */
      var deal = extractDealStatus(norm, statusValue);
      if (deal) { setField(ctx, 'dealStatus', deal.value, deal.index, deal.len, deal.rule); }

      /* 11) 核心原因 */
      var reason = extractCoreReason(norm);
      if (reason) { setField(ctx, 'coreReason', reason.value, reason.index, reason.len, reason.rule); }

      /* 12) 跟进记录（最新在前，最多 3 条） */
      out.followups = buildFollowups(norm, body, dates, inquiry ? inquiry.iso : '');

      /* 13) 汇总 */
      out.fields = ctx.fields;
      out.evidence = ctx.evidence;
      out.filled = Object.keys(ctx.fields).length;
      out.missing = FIELD_KEYS.filter(function (k) { return !ctx.fields[k]; });
      out.ok = out.filled > 0;
      out.raw.matchedRules = ctx.rules.slice();
      if (truncated) {
        out.warnings.push({ code: 'truncated', message: '文本过长，仅识别前 20 万字' });
      }
      return out;
    } catch (err) {
      /* 契约要求绝不抛异常：兜底返回可用的空结果 + 警告 */
      out.warnings.push({ code: 'internal', message: '识别过程出现异常，请手动核对结果' });
      return out;
    }
  }

  /**
   * 单字段日期识别（便于单测）：返回文本中最靠前的一个完整日期。
   */
  function parseDate(text, opts) {
    try {
      var input = (text === null || text === undefined) ? '' : String(text);
      if (!input.trim()) { return ''; }
      var norm = toHalfWidth(input);
      var dates = collectDates(norm, resolveFallbackYear(opts));
      return dates.length ? dates[0].iso : '';
    } catch (err) {
      return '';
    }
  }

  /* ================================================================== */
  /* 11. 导出                                                             */
  /* ================================================================== */

  return {
    VERSION: VERSION,
    recognize: recognize,
    parseDate: parseDate,
    SAMPLE_TEXTS: SAMPLE_TEXTS
  };
});
