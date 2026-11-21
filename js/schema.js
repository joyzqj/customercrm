/*!
 * 客户通 CustomerCRM — js/schema.js
 * 共享数据契约：字段定义、枚举、国家表、派生计算、宽松日期解析。
 * 不依赖任何其他文件；浏览器挂到 window.CRM_SCHEMA，Node 下 module.exports。
 */
(function (root, factory) {
  'use strict';
  var api = factory();
  if (typeof module === 'object' && module.exports) { module.exports = api; }
  if (root) { root.CRM_SCHEMA = api; }
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  var VERSION = '1.2.0';
  var STORAGE_KEY = 'customer-crm.v1';

  /* ------------------------------------------------------------------ */
  /* 1. 枚举                                                             */
  /* ------------------------------------------------------------------ */

  var FOLLOWUP_STATUS = ['新询盘', '已回复', '已报价', '寄样中', '谈判中', '已成交', '暂停跟进', '已流失'];
  var DEAL_STATUS = ['已成交', '未成交'];

  var STATUS_TONE = {
    '新询盘': 'info',
    '已回复': 'neutral',
    '已报价': 'brand',
    '寄样中': 'warn',
    '谈判中': 'warn',
    '已成交': 'ok',
    '暂停跟进': 'neutral',
    '已流失': 'danger'
  };
  var DEAL_TONE = { '已成交': 'ok', '未成交': 'neutral' };

  /* ------------------------------------------------------------------ */
  /* 2. 国家表（中文名 + ISO2 + 英文名 + 别名，供识别与旗帜使用）           */
  /* ------------------------------------------------------------------ */

  var COUNTRIES = [
    /* 亚洲 */
    { name: '中国', iso: 'CN', en: 'China', aliases: ['中国大陆', 'mainland china', 'prc'] },
    { name: '中国香港', iso: 'HK', en: 'Hong Kong', aliases: ['hongkong', 'hk'] },
    { name: '中国台湾', iso: 'TW', en: 'Taiwan', aliases: ['taiwanese'] },
    { name: '中国澳门', iso: 'MO', en: 'Macau', aliases: ['macao'] },
    { name: '日本', iso: 'JP', en: 'Japan', aliases: ['japanese', 'nippon'] },
    { name: '韩国', iso: 'KR', en: 'South Korea', aliases: ['korea', 'korean', 'republic of korea'] },
    { name: '印度', iso: 'IN', en: 'India', aliases: ['indian'] },
    { name: '越南', iso: 'VN', en: 'Vietnam', aliases: ['vietnamese', 'viet nam'] },
    { name: '泰国', iso: 'TH', en: 'Thailand', aliases: ['thai'] },
    { name: '马来西亚', iso: 'MY', en: 'Malaysia', aliases: ['malaysian'] },
    { name: '新加坡', iso: 'SG', en: 'Singapore', aliases: ['singaporean'] },
    { name: '印度尼西亚', iso: 'ID', en: 'Indonesia', aliases: ['indonesian', '印尼'] },
    { name: '菲律宾', iso: 'PH', en: 'Philippines', aliases: ['philippine', 'filipino'] },
    { name: '巴基斯坦', iso: 'PK', en: 'Pakistan', aliases: ['pakistani'] },
    { name: '孟加拉国', iso: 'BD', en: 'Bangladesh', aliases: ['bangladeshi', '孟加拉'] },
    { name: '斯里兰卡', iso: 'LK', en: 'Sri Lanka', aliases: ['srilanka'] },
    { name: '尼泊尔', iso: 'NP', en: 'Nepal' },
    { name: '缅甸', iso: 'MM', en: 'Myanmar', aliases: ['burma'] },
    { name: '柬埔寨', iso: 'KH', en: 'Cambodia' },
    { name: '老挝', iso: 'LA', en: 'Laos' },
    { name: '文莱', iso: 'BN', en: 'Brunei' },
    { name: '蒙古', iso: 'MN', en: 'Mongolia' },
    { name: '哈萨克斯坦', iso: 'KZ', en: 'Kazakhstan' },
    { name: '乌兹别克斯坦', iso: 'UZ', en: 'Uzbekistan' },
    { name: '阿联酋', iso: 'AE', en: 'United Arab Emirates', aliases: ['uae', 'emirates', 'dubai', 'abu dhabi', '阿拉伯联合酋长国', '迪拜'] },
    { name: '沙特阿拉伯', iso: 'SA', en: 'Saudi Arabia', aliases: ['saudi', 'ksa', '沙特'] },
    { name: '卡塔尔', iso: 'QA', en: 'Qatar' },
    { name: '科威特', iso: 'KW', en: 'Kuwait' },
    { name: '阿曼', iso: 'OM', en: 'Oman' },
    { name: '巴林', iso: 'BH', en: 'Bahrain' },
    { name: '约旦', iso: 'JO', en: 'Jordan' },
    { name: '以色列', iso: 'IL', en: 'Israel', aliases: ['israeli'] },
    { name: '黎巴嫩', iso: 'LB', en: 'Lebanon' },
    { name: '伊拉克', iso: 'IQ', en: 'Iraq' },
    { name: '伊朗', iso: 'IR', en: 'Iran', aliases: ['iranian'] },
    { name: '土耳其', iso: 'TR', en: 'Turkey', aliases: ['turkish', 'turkiye', 'türkiye'] },
    { name: '塞浦路斯', iso: 'CY', en: 'Cyprus' },
    { name: '叙利亚', iso: 'SY', en: 'Syria' },
    { name: '也门', iso: 'YE', en: 'Yemen' },
    { name: '阿富汗', iso: 'AF', en: 'Afghanistan' },
    { name: '马尔代夫', iso: 'MV', en: 'Maldives' },
    /* 欧洲 */
    { name: '德国', iso: 'DE', en: 'Germany', aliases: ['german', 'deutschland'] },
    { name: '法国', iso: 'FR', en: 'France', aliases: ['french'] },
    { name: '英国', iso: 'GB', en: 'United Kingdom', aliases: ['uk', 'britain', 'british', 'england', 'english', 'great britain', 'scotland', 'wales'] },
    { name: '意大利', iso: 'IT', en: 'Italy', aliases: ['italian'] },
    { name: '西班牙', iso: 'ES', en: 'Spain', aliases: ['spanish'] },
    { name: '葡萄牙', iso: 'PT', en: 'Portugal', aliases: ['portuguese'] },
    { name: '荷兰', iso: 'NL', en: 'Netherlands', aliases: ['dutch', 'holland', 'the netherlands'] },
    { name: '比利时', iso: 'BE', en: 'Belgium', aliases: ['belgian'] },
    { name: '卢森堡', iso: 'LU', en: 'Luxembourg' },
    { name: '瑞士', iso: 'CH', en: 'Switzerland', aliases: ['swiss'] },
    { name: '奥地利', iso: 'AT', en: 'Austria', aliases: ['austrian'] },
    { name: '瑞典', iso: 'SE', en: 'Sweden', aliases: ['swedish'] },
    { name: '挪威', iso: 'NO', en: 'Norway', aliases: ['norwegian'] },
    { name: '丹麦', iso: 'DK', en: 'Denmark', aliases: ['danish'] },
    { name: '芬兰', iso: 'FI', en: 'Finland', aliases: ['finnish'] },
    { name: '冰岛', iso: 'IS', en: 'Iceland' },
    { name: '爱尔兰', iso: 'IE', en: 'Ireland', aliases: ['irish'] },
    { name: '波兰', iso: 'PL', en: 'Poland', aliases: ['polish'] },
    { name: '捷克', iso: 'CZ', en: 'Czech Republic', aliases: ['czechia', 'czech', '捷克共和国'] },
    { name: '斯洛伐克', iso: 'SK', en: 'Slovakia', aliases: ['slovak'] },
    { name: '匈牙利', iso: 'HU', en: 'Hungary', aliases: ['hungarian'] },
    { name: '罗马尼亚', iso: 'RO', en: 'Romania', aliases: ['romanian'] },
    { name: '保加利亚', iso: 'BG', en: 'Bulgaria', aliases: ['bulgarian'] },
    { name: '希腊', iso: 'GR', en: 'Greece', aliases: ['greek'] },
    { name: '塞尔维亚', iso: 'RS', en: 'Serbia', aliases: ['serbian'] },
    { name: '克罗地亚', iso: 'HR', en: 'Croatia', aliases: ['croatian'] },
    { name: '斯洛文尼亚', iso: 'SI', en: 'Slovenia' },
    { name: '波黑', iso: 'BA', en: 'Bosnia and Herzegovina', aliases: ['bosnia'] },
    { name: '黑山', iso: 'ME', en: 'Montenegro' },
    { name: '北马其顿', iso: 'MK', en: 'North Macedonia', aliases: ['macedonia'] },
    { name: '阿尔巴尼亚', iso: 'AL', en: 'Albania' },
    { name: '乌克兰', iso: 'UA', en: 'Ukraine', aliases: ['ukrainian'] },
    { name: '俄罗斯', iso: 'RU', en: 'Russia', aliases: ['russian', '俄联邦', '俄国'] },
    { name: '白俄罗斯', iso: 'BY', en: 'Belarus' },
    { name: '摩尔多瓦', iso: 'MD', en: 'Moldova' },
    { name: '立陶宛', iso: 'LT', en: 'Lithuania' },
    { name: '拉脱维亚', iso: 'LV', en: 'Latvia' },
    { name: '爱沙尼亚', iso: 'EE', en: 'Estonia' },
    /* 北美 */
    { name: '美国', iso: 'US', en: 'United States', aliases: ['usa', 'u.s.a', 'america', 'american', 'the states', 'united states of america', '美利坚'] },
    { name: '加拿大', iso: 'CA', en: 'Canada', aliases: ['canadian'] },
    { name: '墨西哥', iso: 'MX', en: 'Mexico', aliases: ['mexican'] },
    /* 中南美 */
    { name: '巴西', iso: 'BR', en: 'Brazil', aliases: ['brazilian'] },
    { name: '阿根廷', iso: 'AR', en: 'Argentina', aliases: ['argentinian', 'argentine'] },
    { name: '智利', iso: 'CL', en: 'Chile', aliases: ['chilean'] },
    { name: '秘鲁', iso: 'PE', en: 'Peru', aliases: ['peruvian'] },
    { name: '哥伦比亚', iso: 'CO', en: 'Colombia', aliases: ['colombian'] },
    { name: '厄瓜多尔', iso: 'EC', en: 'Ecuador' },
    { name: '委内瑞拉', iso: 'VE', en: 'Venezuela' },
    { name: '乌拉圭', iso: 'UY', en: 'Uruguay' },
    { name: '巴拉圭', iso: 'PY', en: 'Paraguay' },
    { name: '玻利维亚', iso: 'BO', en: 'Bolivia' },
    { name: '圭亚那', iso: 'GY', en: 'Guyana' },
    { name: '苏里南', iso: 'SR', en: 'Suriname' },
    { name: '危地马拉', iso: 'GT', en: 'Guatemala' },
    { name: '洪都拉斯', iso: 'HN', en: 'Honduras' },
    { name: '萨尔瓦多', iso: 'SV', en: 'El Salvador' },
    { name: '尼加拉瓜', iso: 'NI', en: 'Nicaragua' },
    { name: '哥斯达黎加', iso: 'CR', en: 'Costa Rica' },
    { name: '巴拿马', iso: 'PA', en: 'Panama' },
    { name: '古巴', iso: 'CU', en: 'Cuba' },
    { name: '多米尼加', iso: 'DO', en: 'Dominican Republic' },
    { name: '海地', iso: 'HT', en: 'Haiti' },
    { name: '牙买加', iso: 'JM', en: 'Jamaica' },
    { name: '特立尼达和多巴哥', iso: 'TT', en: 'Trinidad and Tobago' },
    { name: '巴哈马', iso: 'BS', en: 'Bahamas' },
    { name: '巴巴多斯', iso: 'BB', en: 'Barbados' },
    /* 大洋洲 */
    { name: '澳大利亚', iso: 'AU', en: 'Australia', aliases: ['australian', 'aussie', '澳洲'] },
    { name: '新西兰', iso: 'NZ', en: 'New Zealand', aliases: ['kiwi'] },
    { name: '斐济', iso: 'FJ', en: 'Fiji' },
    { name: '巴布亚新几内亚', iso: 'PG', en: 'Papua New Guinea' },
    /* 非洲 */
    { name: '南非', iso: 'ZA', en: 'South Africa', aliases: ['south african'] },
    { name: '埃及', iso: 'EG', en: 'Egypt', aliases: ['egyptian'] },
    { name: '尼日利亚', iso: 'NG', en: 'Nigeria', aliases: ['nigerian'] },
    { name: '肯尼亚', iso: 'KE', en: 'Kenya', aliases: ['kenyan'] },
    { name: '摩洛哥', iso: 'MA', en: 'Morocco', aliases: ['moroccan'] },
    { name: '阿尔及利亚', iso: 'DZ', en: 'Algeria', aliases: ['algerian'] },
    { name: '突尼斯', iso: 'TN', en: 'Tunisia', aliases: ['tunisian'] },
    { name: '利比亚', iso: 'LY', en: 'Libya' },
    { name: '苏丹', iso: 'SD', en: 'Sudan' },
    { name: '埃塞俄比亚', iso: 'ET', en: 'Ethiopia' },
    { name: '加纳', iso: 'GH', en: 'Ghana' },
    { name: '科特迪瓦', iso: 'CI', en: "Cote d'Ivoire", aliases: ['ivory coast'] },
    { name: '塞内加尔', iso: 'SN', en: 'Senegal' },
    { name: '喀麦隆', iso: 'CM', en: 'Cameroon' },
    { name: '坦桑尼亚', iso: 'TZ', en: 'Tanzania' },
    { name: '乌干达', iso: 'UG', en: 'Uganda' },
    { name: '安哥拉', iso: 'AO', en: 'Angola' },
    { name: '莫桑比克', iso: 'MZ', en: 'Mozambique' },
    { name: '津巴布韦', iso: 'ZW', en: 'Zimbabwe' },
    { name: '赞比亚', iso: 'ZM', en: 'Zambia' },
    { name: '博茨瓦纳', iso: 'BW', en: 'Botswana' },
    { name: '纳米比亚', iso: 'NA', en: 'Namibia' },
    { name: '毛里求斯', iso: 'MU', en: 'Mauritius' },
    { name: '马达加斯加', iso: 'MG', en: 'Madagascar' }
  ];

  /* ------------------------------------------------------------------ */
  /* 3. 字段定义                                                          */
  /* ------------------------------------------------------------------ */

  var FIELDS = [
    {
      key: 'firstInquiryAt', label: '首次询盘时间', type: 'date', required: true,
      table: true, excel: '首次询盘时间', width: 14, hint: '客户第一次询盘的日期',
      placeholder: '2026-09-12'
    },
    {
      key: 'customerName', label: '客户名称', type: 'text', required: true,
      table: true, excel: '客户名称', width: 20, hint: '联系人姓名，如 Michael Braun',
      placeholder: 'Michael Braun'
    },
    {
      key: 'email', label: '邮箱', type: 'email', required: false,
      table: true, excel: '邮箱', width: 26, hint: '客户联系邮箱，用于发送报价 / 跟进邮件',
      placeholder: 'michael.braun@braun-trading.de'
    },
    {
      key: 'phone', label: '电话', type: 'tel', required: false,
      table: true, excel: '电话', width: 20, hint: '客户电话 / 手机 / WhatsApp，建议带国际区号',
      placeholder: '+49 171 1234567'
    },
    {
      key: 'address', label: '地址', type: 'text', required: false, full: true,
      table: true, excel: '地址', width: 34, hint: '客户公司地址 / 收货地址（中英文均可）',
      placeholder: 'Industriestr. 12, 70565 Stuttgart, Germany'
    },
    {
      key: 'product', label: '咨询产品', type: 'text', required: false,
      table: true, excel: '咨询产品', width: 26, hint: '客户询价的产品名称，可从「产品库」下拉选择',
      placeholder: '不锈钢保温杯'
    },
    {
      key: 'productModel', label: '产品型号', type: 'text', required: false,
      table: false, excel: '产品型号', width: 16, hint: '与产品名称对应的型号 / 规格，如 500ml、16 件套',
      placeholder: '500ml'
    },
    {
      key: 'country', label: '国家', type: 'select', required: false,
      table: true, excel: '国家', width: 14, hint: '客户所在国家 / 地区', options: 'countries',
      placeholder: '德国'
    },
    {
      key: 'companyName', label: '公司名', type: 'text', required: false,
      table: true, excel: '公司名', width: 30, hint: '客户公司全称',
      placeholder: 'Braun Trading GmbH'
    },
    {
      key: 'mainBusiness', label: '主营业务', type: 'text', required: false,
      table: true, excel: '主营业务', width: 22, hint: '客户公司的主营业务 / 主要进口品类，如 户外用品进口分销',
      placeholder: '户外用品进口分销'
    },
    {
      key: 'foundedAt', label: '成立时间', type: 'text', required: false,
      table: true, excel: '成立时间', width: 14, hint: '公司成立年份或时间，如 1998 / 1998年3月',
      placeholder: '1998'
    },
    {
      key: 'teamSize', label: '团队规模', type: 'text', required: false,
      table: true, excel: '团队规模', width: 16, hint: '员工人数或团队规模，如 50-100 人',
      placeholder: '50-100 人'
    },
    {
      key: 'annualRevenue', label: '年营收', type: 'text', required: false,
      table: true, excel: '年营收', width: 18, hint: '年营业额 / 采购额，如 1200 万欧元',
      placeholder: '1200 万欧元'
    },
    {
      key: 'followupStatus', label: '跟进状态', type: 'select', required: false,
      table: true, excel: '跟进状态', width: 14, hint: '当前跟进阶段', options: 'status'
    },
    {
      key: 'dealStatus', label: '是否成交', type: 'select', required: false,
      table: true, excel: '是否成交', width: 12, hint: '是否已成交下单', options: 'deal'
    },
    {
      key: 'coreReason', label: '核心原因', type: 'textarea', required: false,
      table: true, excel: '核心原因', width: 44, wrap: true, hint: '成交 / 未成交的关键原因',
      placeholder: '价格高出目标价 8%，等待终端客户确认'
    }
  ];

  var FIELD_GROUPS = [
    {
      title: '客户与公司信息',
      hint: '询盘来源与客户背景，用于判断客户质量与采购能力',
      keys: ['firstInquiryAt', 'customerName', 'email', 'phone', 'address', 'product', 'productModel', 'country', 'companyName', 'mainBusiness', 'foundedAt', 'teamSize', 'annualRevenue']
    },
    {
      title: '跟进与成交',
      hint: '当前进展、是否成交，以及成败的关键原因',
      keys: ['followupStatus', 'dealStatus', 'coreReason']
    }
  ];

  /* ------------------------------------------------------------------ */
  /* 4. 日期与 id 工具                                                    */
  /* ------------------------------------------------------------------ */

  function pad2(n) { n = Number(n) || 0; return n < 10 ? '0' + n : String(n); }

  function todayISO(date) {
    var d = date instanceof Date ? date : new Date();
    return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate());
  }

  function nowISO(date) {
    var d = date instanceof Date ? date : new Date();
    return d.toISOString();
  }

  function isValidISO(iso) {
    if (typeof iso !== 'string') return false;
    if (!/^\d{4}-\d{2}-\d{2}$/.test(iso)) return false;
    return roundTrip(iso) === iso;
  }

  function roundTrip(iso) {
    var parts = iso.split('-');
    var y = Number(parts[0]); var m = Number(parts[1]); var d = Number(parts[2]);
    var dt = new Date(Date.UTC(y, m - 1, d));
    if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== m - 1 || dt.getUTCDate() !== d) return '';
    return dt.getUTCFullYear() + '-' + pad2(dt.getUTCMonth() + 1) + '-' + pad2(dt.getUTCDate());
  }

  function buildISO(y, m, d) {
    y = Number(y); m = Number(m); d = Number(d);
    if (!isFinite(y) || !isFinite(m) || !isFinite(d)) return '';
    if (y < 100) { y += y >= 70 ? 1900 : 2000; }
    if (y < 1900 || y > 2100) return '';
    if (m < 1 || m > 12) return '';
    if (d < 1 || d > 31) return '';
    return roundTrip(y + '-' + pad2(m) + '-' + pad2(d));
  }

  function addDays(iso, n) {
    if (!isValidISO(iso)) return '';
    n = Number(n) || 0;
    var parts = iso.split('-');
    var dt = new Date(Date.UTC(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]) + n));
    return dt.getUTCFullYear() + '-' + pad2(dt.getUTCMonth() + 1) + '-' + pad2(dt.getUTCDate());
  }

  function daysBetween(fromISO, toISOValue) {
    if (!isValidISO(fromISO) || !isValidISO(toISOValue)) return 0;
    var a = Date.parse(fromISO + 'T00:00:00Z');
    var b = Date.parse(toISOValue + 'T00:00:00Z');
    return Math.round((b - a) / 86400000);
  }

  var MONTH_NAMES = {
    jan: 1, january: 1, feb: 2, february: 2, mar: 3, march: 3, apr: 4, april: 4,
    may: 5, jun: 6, june: 6, jul: 7, july: 7, aug: 8, august: 8,
    sep: 9, sept: 9, september: 9, oct: 10, october: 10, nov: 11, november: 11,
    dec: 12, december: 12
  };
  var CN_MONTHS = { '一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9, '十': 10, '十一': 11, '十二': 12 };

  /**
   * 宽松日期解析：'2026-09-12' / '2026/9/12' / '2026年9月12日' / '9月12日'
   * / 'Sep 12, 2026' / '12 September 2026' / '12/09/2026'（按 DD/MM 处理）。
   * 返回 'YYYY-MM-DD'，失败返回 ''。
   */
  function parseLooseDate(text, opts) {
    if (text === null || text === undefined) return '';
    var s = String(text).trim();
    if (!s) return '';
    opts = opts || {};
    var fallbackYear = Number(opts.fallbackYear) || new Date().getFullYear();
    var m;

    if (isValidISO(s)) return s;

    /* 2026-09-12 / 2026/9/12 / 2026.9.12 / 2026年9月12日 */
    m = s.match(/(\d{4})\s*[-/.年]\s*(\d{1,2})\s*[-/.月]\s*(\d{1,2})\s*[日号]?/);
    if (m) { var iso1 = buildISO(m[1], m[2], m[3]); if (iso1) return iso1; }

    /* 2026年9月12日（无分隔符变体） */
    m = s.match(/(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*[日号]?/);
    if (m) { var iso2 = buildISO(m[1], m[2], m[3]); if (iso2) return iso2; }

    /* 26年9月12日 */
    m = s.match(/(\d{2})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*[日号]?/);
    if (m) { var iso3 = buildISO(m[1], m[2], m[3]); if (iso3) return iso3; }

    /* 每年 / YYYY-MM / YYYY年M月（不构成完整日期，交给调用方处理年份字段） */
    if (/^\d{4}\s*[-/.年]\s*\d{1,2}\s*月?$/.test(s)) return '';

    /* 9月12日 / 9月12号 */
    m = s.match(/(\d{1,2})\s*月\s*(\d{1,2})\s*[日号]?/);
    if (m) { var iso4 = buildISO(fallbackYear, m[1], m[2]); if (iso4) return iso4; }

    /* 九月十二日 */
    m = s.match(/([一二三四五六七八九十]{1,2})\s*月\s*([一二三四五六七八九十]{1,3})\s*[日号]/);
    if (m && CN_MONTHS[m[1]]) { var cnDay = cnNumber(m[2]); if (cnDay) { var iso5 = buildISO(fallbackYear, CN_MONTHS[m[1]], cnDay); if (iso5) return iso5; } }

    /* Sep 12, 2026 / September 12 2026 / Sep 12 */
    m = s.match(/\b([A-Za-z]{3,9})\.?\s+(\d{1,2})(?:st|nd|rd|th)?(?:\s*,?\s*(\d{4}))?\b/);
    if (m) {
      var mon = MONTH_NAMES[m[1].toLowerCase()];
      if (mon) { var iso6 = buildISO(m[3] || fallbackYear, mon, m[2]); if (iso6) return iso6; }
    }

    /* 12 Sep 2026 / 12th September, 2026 */
    m = s.match(/\b(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]{3,9})\.?(?:\s*,?\s*(\d{4}))?\b/);
    if (m) {
      var mon2 = MONTH_NAMES[m[2].toLowerCase()];
      if (mon2) { var iso7 = buildISO(m[3] || fallbackYear, mon2, m[1]); if (iso7) return iso7; }
    }

    /* 12/09/2026：两位都不超过 12 时按 DD/MM 处理（外贸场景惯例），否则按较大值判定 */
    m = s.match(/^(\d{1,2})\s*[/\-.]\s*(\d{1,2})\s*[/\-.]\s*(\d{2,4})$/);
    if (m) {
      var b = Number(m[1]); var c = Number(m[2]);
      var iso8 = b > 12 ? buildISO(m[3], c, b) : c > 12 ? buildISO(m[3], b, c) : buildISO(m[3], c, b);
      if (iso8) return iso8;
    }

    return '';
  }

  function cnNumber(text) {
    var map = { '一': 1, '二': 2, '两': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9, '十': 10 };
    if (text.length === 1) return map[text] || 0;
    if (text === '十一') return 11; if (text === '十二') return 12;
    if (text === '二十') return 20;
    if (text.length === 2 && text.charAt(0) === '十') return 10 + (map[text.charAt(1)] || 0);
    return 0;
  }

  function formatShort(iso, withYear) {
    if (!isValidISO(iso)) return iso || '';
    return withYear ? iso : iso.slice(5);
  }

  function formatDateTime(value) {
    if (!value) return '';
    var d = value instanceof Date ? value : new Date(value);
    if (isNaN(d.getTime())) return String(value);
    return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate()) +
      ' ' + pad2(d.getHours()) + ':' + pad2(d.getMinutes());
  }

  function uid(prefix) {
    var stamp = Date.now().toString(36);
    var rand = Math.random().toString(36).slice(2, 8);
    return (prefix || 'c_') + stamp + rand;
  }

  /* ------------------------------------------------------------------ */
  /* 5. 国家工具                                                          */
  /* ------------------------------------------------------------------ */

  function asciiBoundaryOk(text, index, termLen) {
    var before = index > 0 ? text.charAt(index - 1) : '';
    var after = text.charAt(index + termLen);
    var word = /[A-Za-z0-9]/;
    if (before && word.test(before)) return false;
    if (after && word.test(after)) return false;
    return true;
  }

  var COUNTRY_TERMS = null;
  function buildCountryTerms() {
    if (COUNTRY_TERMS) return COUNTRY_TERMS;
    var terms = [];
    COUNTRIES.forEach(function (c) {
      var list = [c.name, c.en].concat(c.aliases || []);
      list.forEach(function (term) {
        if (!term) return;
        terms.push({ term: String(term), entry: c, isAscii: /^[\x20-\x7E]+$/.test(String(term)) });
      });
    });
    terms.sort(function (a, b) { return b.term.length - a.term.length; });
    COUNTRY_TERMS = terms;
    return terms;
  }

  /** 在文本中查找国家（最长匹配优先），返回国家条目或 null。 */
  function findCountry(text) {
    if (!text) return null;
    var s = String(text);
    var lower = s.toLowerCase();
    var terms = buildCountryTerms();
    for (var i = 0; i < terms.length; i += 1) {
      var t = terms[i];
      if (t.isAscii) {
        var needle = t.term.toLowerCase();
        var from = 0;
        var idx = lower.indexOf(needle, from);
        while (idx !== -1) {
          if (asciiBoundaryOk(lower, idx, needle.length)) {
            return { name: t.entry.name, iso: t.entry.iso, en: t.entry.en, term: t.term, index: idx };
          }
          from = idx + 1;
          idx = lower.indexOf(needle, from);
        }
      } else {
        var pos = s.indexOf(t.term);
        if (pos !== -1) {
          return { name: t.entry.name, iso: t.entry.iso, en: t.entry.en, term: t.term, index: pos };
        }
      }
    }
    return null;
  }

  function countryByName(name) {
    if (!name) return null;
    var s = String(name).trim();
    for (var i = 0; i < COUNTRIES.length; i += 1) {
      if (COUNTRIES[i].name === s) return COUNTRIES[i];
    }
    return findCountry(s);
  }

  function flagFromISO(iso) {
    if (!iso || String(iso).length !== 2) return '';
    var code = String(iso).toUpperCase();
    if (!/^[A-Z]{2}$/.test(code)) return '';
    var base = 0x1F1E6;
    return String.fromCodePoint(base + code.charCodeAt(0) - 65, base + code.charCodeAt(1) - 65);
  }

  function countryFlag(name) {
    var c = countryByName(name);
    return c ? flagFromISO(c.iso) : '';
  }

  function countryOptions() {
    return COUNTRIES.map(function (c) { return c.name; });
  }

  /* ------------------------------------------------------------------ */
  /* 6. 记录构造、清洗、派生                                              */
  /* ------------------------------------------------------------------ */

  var DEFAULT_RECORD = {
    id: '',
    firstInquiryAt: '',
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
    createdAt: '',
    updatedAt: ''
  };

  /* 产品库预置数据：首次使用 / 清空后恢复时写入，可在设置页增删 */
  var DEFAULT_PRODUCTS = [
    { name: '不锈钢保温杯', model: '500ml' },
    { name: '不锈钢保温杯', model: '750ml' },
    { name: '不锈钢保温杯', model: '1L' },
    { name: '不锈钢运动水壶', model: '1L' },
    { name: '陶瓷餐具套装', model: '16 件套' },
    { name: '硅胶厨具套装', model: '12 件套' },
    { name: '户外折叠椅', model: '标准款' },
    { name: '保温饭盒', model: '1.2L' },
    { name: '玻璃保鲜盒', model: '5 件套' },
    { name: '硅胶折叠水杯', model: '500ml' },
    { name: '密胺餐具', model: '10 件套' },
    { name: '不粘锅套装', model: '3 件套' }
  ];

  var FIELD_KEYS = FIELDS.map(function (f) { return f.key; });

  function emptyRecord(today) {
    var rec = JSON.parse(JSON.stringify(DEFAULT_RECORD));
    rec.id = uid('c_');
    rec.firstInquiryAt = today || todayISO();
    rec.createdAt = nowISO();
    rec.updatedAt = rec.createdAt;
    return rec;
  }

  function plainText(value) {
    if (value === null || value === undefined) return '';
    if (typeof value === 'number') return String(value);
    if (typeof value === 'boolean') return value ? '是' : '否';
    return String(value).replace(/\r\n/g, '\n').replace(/\u00a0/g, ' ').trim();
  }

  function normalizeStatus(value) {
    var s = plainText(value);
    if (!s) return '新询盘';
    if (FOLLOWUP_STATUS.indexOf(s) !== -1) return s;
    if (/成交|下单|签约|closed|won|deal/i.test(s)) return '已成交';
    if (/报价|quote|quotation/i.test(s)) return '已报价';
    if (/样品|寄样|sample/i.test(s)) return '寄样中';
    if (/谈判|合同|negoti/i.test(s)) return '谈判中';
    if (/流失|放弃|丢了|lost/i.test(s)) return '已流失';
    if (/暂停|搁置|hold|pause/i.test(s)) return '暂停跟进';
    if (/回复|已回|reply|replied/i.test(s)) return '已回复';
    return '新询盘';
  }

  function normalizeDeal(value) {
    var s = plainText(value);
    if (DEAL_STATUS.indexOf(s) !== -1) return s;
    if (/^(是|y|yes|true|1|已成交|成交)$/i.test(s)) return '已成交';
    if (/成交|下单|签约|closed|won/i.test(s) && !/未|没|not|no/i.test(s)) return '已成交';
    return '未成交';
  }

  function normalizeFollowups(list) {
    var out = [];
    if (Array.isArray(list)) {
      list.forEach(function (item) {
        var date = '';
        var note = '';
        if (item && typeof item === 'object') {
          date = parseLooseDate(item.date);
          note = plainText(item.note || item.text || item.content);
        } else if (typeof item === 'string') {
          var s = plainText(item);
          date = parseLooseDate(s);
          note = s.replace(/^\s*\d{4}\s*[-/.年]\s*\d{1,2}\s*[-/.月]\s*\d{1,2}\s*[日号]?[：:，,\s-]*/, '').trim();
        }
        if (!date) return;
        out.push({ date: date, note: note.slice(0, 300) });
      });
    }
    out.sort(function (a, b) { return a.date < b.date ? 1 : a.date > b.date ? -1 : 0; });
    var seen = {};
    var dedup = [];
    out.forEach(function (f) {
      var key = f.date + '|' + f.note;
      if (seen[key]) return;
      seen[key] = 1;
      dedup.push(f);
    });
    return dedup.slice(0, 3);
  }

  /**
   * 把任意来源（表单 / 识别结果 / JSON 导入）清洗成合法记录。
   * 保证：11 个字段均为字符串、枚举合法、followups 合法且最多 3 条、id/时间戳存在。
   */
  function normalize(raw, today) {
    var rec = {};
    var src = raw && typeof raw === 'object' ? raw : {};
    rec.id = plainText(src.id) || uid('c_');
    FIELD_KEYS.forEach(function (key) {
      var field = fieldByKey(key);
      var value = plainText(src[key]);
      if (key === 'firstInquiryAt') {
        rec[key] = parseLooseDate(value) || (isValidISO(value) ? value : '');
      } else if (key === 'country') {
        var hit = value ? countryByName(value) : null;
        rec[key] = hit ? hit.name : value;
      } else if (key === 'followupStatus') {
        rec[key] = normalizeStatus(value);
      } else if (key === 'dealStatus') {
        rec[key] = value ? normalizeDeal(value) : '未成交';
      } else if (field && field.type === 'text') {
        rec[key] = value.slice(0, 500);
      } else {
        rec[key] = value.slice(0, 2000);
      }
    });
    if (!rec.firstInquiryAt) rec.firstInquiryAt = today || todayISO();
    rec.followups = normalizeFollowups(src.followups);
    rec.createdAt = plainText(src.createdAt) || nowISO();
    rec.updatedAt = plainText(src.updatedAt) || rec.createdAt;
    return rec;
  }

  /**
   * 派生视图：距上次跟进、新鲜度、状态色调等。不修改入参。
   * opts: { today, warnDays=7, staleDays=14 }
   */
  function derive(record, opts) {
    opts = opts || {};
    var rec = record || {};
    var today = isValidISO(opts.today) ? opts.today : todayISO();
    var warnDays = opts.warnDays === undefined ? 7 : Number(opts.warnDays) || 0;
    var staleDays = opts.staleDays === undefined ? 14 : Number(opts.staleDays) || 0;

    var followups = normalizeFollowups(rec.followups);
    var last = followups.length ? followups[0] : null;
    var lastFollowupAt = last ? last.date : '';
    var daysSince = last ? Math.max(0, daysBetween(lastFollowupAt, today)) : null;
    var inquiryDays = isValidISO(rec.firstInquiryAt) ? Math.max(0, daysBetween(rec.firstInquiryAt, today)) : null;

    var isDeal = rec.dealStatus === '已成交';
    var isOpen = !isDeal && rec.followupStatus !== '已流失';
    var isOverdue = isOpen && daysSince !== null && daysSince >= staleDays;
    var isDue = isOpen && daysSince !== null && daysSince >= warnDays;
    var tone = ageTone(daysSince, { warnDays: warnDays, staleDays: staleDays });

    return {
      id: rec.id,
      firstInquiryAt: rec.firstInquiryAt,
      customerName: rec.customerName || '',
      email: rec.email || '',
      phone: rec.phone || '',
      address: rec.address || '',
      product: rec.product || '',
      productModel: rec.productModel || '',
      country: rec.country || '',
      companyName: rec.companyName || '',
      mainBusiness: rec.mainBusiness || '',
      foundedAt: rec.foundedAt || '',
      teamSize: rec.teamSize || '',
      annualRevenue: rec.annualRevenue || '',
      followupStatus: rec.followupStatus || '新询盘',
      dealStatus: rec.dealStatus || '未成交',
      coreReason: rec.coreReason || '',
      followups: followups,
      createdAt: rec.createdAt || '',
      updatedAt: rec.updatedAt || '',
      lastFollowupAt: lastFollowupAt,
      daysSince: daysSince,
      inquiryDays: inquiryDays,
      isDeal: isDeal,
      isOpen: isOpen,
      isDue: isDue,
      isOverdue: isOverdue,
      ageTone: tone,
      statusTone: statusTone(rec.followupStatus),
      dealTone: dealTone(rec.dealStatus),
      last3: followups.map(function (f) {
        return { date: f.date, note: f.note, daysAgo: Math.max(0, daysBetween(f.date, today)) };
      })
    };
  }

  function statusTone(status) { return STATUS_TONE[status] || 'neutral'; }
  function dealTone(status) { return DEAL_TONE[status] || 'neutral'; }

  function ageTone(days, opts) {
    opts = opts || {};
    var warnDays = opts.warnDays === undefined ? 7 : Number(opts.warnDays) || 0;
    var staleDays = opts.staleDays === undefined ? 14 : Number(opts.staleDays) || 0;
    if (days === null || days === undefined || !isFinite(days)) return 'none';
    if (days >= staleDays) return 'danger';
    if (days >= warnDays) return 'warn';
    return 'ok';
  }

  function fieldByKey(key) {
    for (var i = 0; i < FIELDS.length; i += 1) { if (FIELDS[i].key === key) return FIELDS[i]; }
    return null;
  }

  function fieldLabel(key) {
    var f = fieldByKey(key);
    return f ? f.label : key;
  }

  /** 邮箱格式校验（宽松：xx@yy.zz 即可，允许空值） */
  function isValidEmail(value) {
    var text = plainText(value);
    if (!text) return true;
    return /^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$/.test(text);
  }

  /** 表单校验：返回 { ok, errors: { key: message } } */
  function validate(record) {
    var errors = {};
    var rec = record || {};
    if (!isValidISO(rec.firstInquiryAt)) errors.firstInquiryAt = '请选择首次询盘日期';
    if (!plainText(rec.customerName)) errors.customerName = '请填写客户名称';
    if (plainText(rec.customerName).length > 120) errors.customerName = '客户名称过长';
    if (!isValidEmail(rec.email)) errors.email = '邮箱格式不正确，例如 name@example.com';
    if (plainText(rec.phone).length > 60) errors.phone = '电话过长';
    if (plainText(rec.address).length > 500) errors.address = '地址过长';
    if (rec.followupStatus && FOLLOWUP_STATUS.indexOf(rec.followupStatus) === -1) errors.followupStatus = '跟进状态不合法';
    if (rec.dealStatus && DEAL_STATUS.indexOf(rec.dealStatus) === -1) errors.dealStatus = '成交状态不合法';
    return { ok: Object.keys(errors).length === 0, errors: errors };
  }

  /* ------------------------------------------------------------------ */
  /* 7. 导出                                                             */
  /* ------------------------------------------------------------------ */

  return {
    VERSION: VERSION,
    STORAGE_KEY: STORAGE_KEY,
    FOLLOWUP_STATUS: FOLLOWUP_STATUS,
    DEAL_STATUS: DEAL_STATUS,
    STATUS_TONE: STATUS_TONE,
    DEAL_TONE: DEAL_TONE,
    COUNTRIES: COUNTRIES,
    FIELDS: FIELDS,
    FIELD_GROUPS: FIELD_GROUPS,
    FIELD_KEYS: FIELD_KEYS,
    DEFAULT_RECORD: DEFAULT_RECORD,
    DEFAULT_PRODUCTS: DEFAULT_PRODUCTS,

    emptyRecord: emptyRecord,
    normalize: normalize,
    derive: derive,
    validate: validate,
    isValidEmail: isValidEmail,
    fieldByKey: fieldByKey,
    fieldLabel: fieldLabel,

    todayISO: todayISO,
    nowISO: nowISO,
    isValidISO: isValidISO,
    addDays: addDays,
    daysBetween: daysBetween,
    parseLooseDate: parseLooseDate,
    formatShort: formatShort,
    formatDateTime: formatDateTime,
    uid: uid,

    findCountry: findCountry,
    countryByName: countryByName,
    countryFlag: countryFlag,
    flagFromISO: flagFromISO,
    countryOptions: countryOptions,

    statusTone: statusTone,
    dealTone: dealTone,
    ageTone: ageTone
  };
});
