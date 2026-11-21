/**
 * 客户通 CustomerCRM — js/schema.js 单元测试
 * 运行：node tests/test-schema.mjs
 */
import { createRequire } from 'node:module';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const require = createRequire(import.meta.url);
const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const schema = require(path.join(ROOT, 'js', 'schema.js'));

let pass = 0;
const failures = [];

function ok(name, condition, detail) {
  if (condition) { pass += 1; }
  else { failures.push({ name, detail: detail === undefined ? '' : String(detail) }); }
}

function eq(name, actual, expected) {
  const good = JSON.stringify(actual) === JSON.stringify(expected);
  ok(name, good, good ? '' : `expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`);
}

/* ---------------- 字段与枚举 ---------------- */

eq('字段数量为 16', schema.FIELDS.length, 16);
eq('字段顺序', schema.FIELDS.map((f) => f.key), [
  'firstInquiryAt', 'customerName', 'email', 'phone', 'address',
  'product', 'productModel', 'country', 'companyName', 'mainBusiness',
  'foundedAt', 'teamSize', 'annualRevenue', 'followupStatus', 'dealStatus', 'coreReason'
]);
ok('每个字段都有 excel 表头与宽度', schema.FIELDS.every((f) => f.excel && f.width > 0));
ok('产品库预置 ≥ 10 条', schema.DEFAULT_PRODUCTS.length >= 10, schema.DEFAULT_PRODUCTS.length);
ok('预置产品都有名称与型号', schema.DEFAULT_PRODUCTS.every((p) => p.name && p.model));
ok('必填字段只有两个', schema.FIELDS.filter((f) => f.required).length === 2);
ok('跟进状态 8 项', schema.FOLLOWUP_STATUS.length === 8, schema.FOLLOWUP_STATUS.join(','));
ok('成交状态 2 项', schema.DEAL_STATUS.length === 2);
ok('每个状态都有色调', schema.FOLLOWUP_STATUS.every((s) => schema.statusTone(s) !== 'neutral' || ['已回复', '暂停跟进'].includes(s)));
ok('国家表 > 100 条', schema.COUNTRIES.length > 100, schema.COUNTRIES.length);
const countryNames = schema.COUNTRIES.map((c) => c.name);
ok('国家名不重复', new Set(countryNames).size === countryNames.length);
const isoCodes = schema.COUNTRIES.map((c) => c.iso);
ok('ISO 码不重复', new Set(isoCodes).size === isoCodes.length);
ok('国家表都有 en 名', schema.COUNTRIES.every((c) => c.en && c.en.length > 1));
ok('FIELD_GROUPS 覆盖全部字段', schema.FIELD_GROUPS.flatMap((g) => g.keys).length === 16);

/* ---------------- 日期工具 ---------------- */

const today = schema.todayISO();
ok('todayISO 形状', /^\d{4}-\d{2}-\d{2}$/.test(today), today);
eq('addDays +7', schema.addDays('2026-11-18', 7), '2026-11-25');
eq('addDays -1', schema.addDays('2026-01-01', -1), '2025-12-31');
eq('daysBetween', schema.daysBetween('2026-09-12', '2026-11-21'), 70);
eq('daysBetween 同一天', schema.daysBetween('2026-09-12', '2026-09-12'), 0);
eq('isValidISO 闰日', schema.isValidISO('2024-02-29'), true);
eq('isValidISO 非法日', schema.isValidISO('2026-02-30'), false);
eq('isValidISO 格式错', schema.isValidISO('2026/9/1'), false);

const dateCases = [
  ['2026-09-12', '2026-09-12'],
  ['2026/9/12', '2026-09-12'],
  ['2026.9.12', '2026-09-12'],
  ['2026年9月12日', '2026-09-12'],
  ['26年9月12日', '2026-09-12'],
  ['Sep 12, 2026', '2026-09-12'],
  ['September 12 2026', '2026-09-12'],
  ['12 Sep 2026', '2026-09-12'],
  ['12th September, 2026', '2026-09-12'],
  ['12/09/2026', '2026-09-12'],
  ['9月12日', '2026-09-12'],
  ['九月十二日', '2026-09-12'],
  ['not a date', ''],
  ['2026-02-30', '']
];
dateCases.forEach(([input, expected]) => {
  const actual = schema.parseLooseDate(input, { fallbackYear: 2026 });
  eq(`parseLooseDate(${input})`, actual, expected);
});
eq('parseLooseDate 空值', schema.parseLooseDate(''), '');
eq('parseLooseDate null', schema.parseLooseDate(null), '');

/* ---------------- 国家识别 ---------------- */

const countryCases = [
  ['德国客户 Michael', '德国'],
  ['German customer', '德国'],
  ['Braun Trading GmbH, Germany', '德国'],
  ['客户来自越南胡志明', '越南'],
  ['based in Dubai', '阿联酋'],
  ['中国香港客户', '中国香港'],
  ['美国采购商', '美国'],
  ['São Paulo, Brazil', '巴西'],
  ['Japanese buyer', '日本'],
  ['El Nile Import, Cairo, Egypt', '埃及'],
  ['no country here', null]
];
countryCases.forEach(([text, expected]) => {
  const hit = schema.findCountry(text);
  eq(`findCountry(${text})`, hit ? hit.name : null, expected);
});
eq('countryFlag(德国)', schema.countryFlag('德国'), '🇩🇪');
eq('countryFlag(美国)', schema.countryFlag('美国'), '🇺🇸');
eq('countryFlag 未匹配', schema.countryFlag('火星'), '');

/* ---------------- 记录清洗 ---------------- */

const raw = schema.normalize({
  firstInquiryAt: '2026年9月12日',
  customerName: '  Michael Braun  ',
  product: '保温杯 500ml',
  country: 'germany',
  companyName: 'Braun Trading GmbH',
  foundedAt: '1998',
  teamSize: '50-100 人',
  annualRevenue: '1200 万欧元',
  followupStatus: '已报价',
  dealStatus: '未成交',
  coreReason: '价格高出 8%',
  followups: [
    { date: '2026-10-01', note: '旧记录' },
    { date: '2026-11-18', note: '最新记录' },
    { date: '2026-11-10', note: '中间记录' },
    { date: '2026-05-01', note: '第四条应被截断' }
  ]
});
eq('normalize 日期', raw.firstInquiryAt, '2026-09-12');
eq('normalize 去空格', raw.customerName, 'Michael Braun');
eq('normalize 国家规范化', raw.country, '德国');
eq('normalize 跟进排序', raw.followups.map((f) => f.date), ['2026-11-18', '2026-11-10', '2026-10-01']);
eq('normalize 跟进最多 3 条', raw.followups.length, 3);
ok('normalize 有 id', /^c_/.test(raw.id), raw.id);
ok('normalize 有创建时间', !!raw.createdAt);

const messy = schema.normalize({
  customerName: 'A',
  firstInquiryAt: '2026/9/1',
  followupStatus: 'price quoted',
  dealStatus: 'invalid',
  country: '',
  followups: ['2026-11-01 寄出样品', 'garbage', { note: '无日期' }]
});
eq('状态宽松归一', messy.followupStatus, '已报价');
eq('成交宽松归一', messy.dealStatus, '未成交');
eq('字符串跟进解析', messy.followups.length, 1);
eq('字符串跟进日期', messy.followups[0] && messy.followups[0].date, '2026-11-01');

eq('已成交识别', schema.normalize({ customerName: 'B', firstInquiryAt: '2026-01-01', dealStatus: '已成交' }).dealStatus, '已成交');
eq('中文成交词归一', schema.normalize({ customerName: 'B', firstInquiryAt: '2026-01-01', followupStatus: '已成交' }).followupStatus, '已成交');

/* ---------------- 派生与色调 ---------------- */

const derived = schema.derive({
  id: 'c_test',
  firstInquiryAt: schema.addDays(today, -71),
  customerName: 'T',
  followupStatus: '已报价',
  dealStatus: '未成交',
  followups: [
    { date: schema.addDays(today, -3), note: 'a' },
    { date: schema.addDays(today, -11), note: 'b' }
  ]
}, { today, warnDays: 7, staleDays: 14 });
eq('derive lastFollowupAt', derived.lastFollowupAt, schema.addDays(today, -3));
eq('derive daysSince', derived.daysSince, 3);
eq('derive inquiryDays', derived.inquiryDays, 71);
eq('derive ageTone=ok', derived.ageTone, 'ok');
eq('derive isOpen', derived.isOpen, true);
eq('derive isOverdue', derived.isOverdue, false);
eq('derive last3 长度', derived.last3.length, 2);
eq('derive statusTone', derived.statusTone, 'brand');

const stale = schema.derive({ id: 'x', firstInquiryAt: schema.addDays(today, -90), customerName: 'S', followupStatus: '谈判中', dealStatus: '未成交', followups: [{ date: schema.addDays(today, -20), note: 'c' }] }, { today, warnDays: 7, staleDays: 14 });
eq('derive 超期', stale.isOverdue, true);
eq('derive danger', stale.ageTone, 'danger');

const noFollow = schema.derive({ id: 'y', firstInquiryAt: schema.addDays(today, -30), customerName: 'N', followupStatus: '新询盘', dealStatus: '未成交', followups: [] }, { today });
eq('无跟进 daysSince', noFollow.daysSince, null);
eq('无跟进 tone', noFollow.ageTone, 'none');
eq('无跟进 lastFollowupAt', noFollow.lastFollowupAt, '');

const dealRec = schema.derive({ id: 'z', firstInquiryAt: schema.addDays(today, -10), customerName: 'D', followupStatus: '已成交', dealStatus: '已成交', followups: [] }, { today });
eq('已成交不是 open', dealRec.isOpen, false);
eq('已成交 tone', dealRec.dealTone, 'ok');

eq('ageTone(null)', schema.ageTone(null), 'none');
eq('ageTone(0)', schema.ageTone(0), 'ok');
eq('ageTone(7)', schema.ageTone(7), 'warn');
eq('ageTone(14)', schema.ageTone(14), 'danger');

/* ---------------- 校验 ---------------- */

const bad = schema.validate({ firstInquiryAt: '', customerName: '' });
eq('校验失败', bad.ok, false);
ok('校验错误字段', !!bad.errors.firstInquiryAt && !!bad.errors.customerName);
const good = schema.validate({ firstInquiryAt: '2026-09-12', customerName: 'OK', followupStatus: '新询盘', dealStatus: '未成交' });
eq('校验通过', good.ok, true);

/* ---------------- 其他 ---------------- */

const empty1 = schema.emptyRecord();
const empty2 = schema.emptyRecord();
eq('emptyRecord 默认状态', empty1.followupStatus, '新询盘');
eq('emptyRecord 默认成交', empty1.dealStatus, '未成交');
ok('emptyRecord id 唯一', empty1.id !== empty2.id);
eq('fieldByKey', schema.fieldByKey('country').label, '国家');
eq('fieldLabel', schema.fieldLabel('customerName'), '客户名称');
eq('uid 前缀', schema.uid('t_').startsWith('t_'), true);
eq('formatShort', schema.formatShort('2026-11-18'), '11-18');
eq('formatShort 带年', schema.formatShort('2026-11-18', true), '2026-11-18');
ok('formatDateTime', /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/.test(schema.formatDateTime('2026-11-18T09:30:00.000Z')), schema.formatDateTime('2026-11-18T09:30:00.000Z'));
eq('STORAGE_KEY', schema.STORAGE_KEY, 'customer-crm.v1');

/* ---------------- v1.1 新增字段：产品型号 / 主营业务 ---------------- */

const v11 = schema.normalize({
  firstInquiryAt: '2026-10-01',
  customerName: 'V11',
  product: '不锈钢保温杯',
  productModel: '500ml',
  mainBusiness: '户外用品进口分销'
});
eq('normalize productModel', v11.productModel, '500ml');
eq('normalize mainBusiness', v11.mainBusiness, '户外用品进口分销');
const v11d = schema.derive(v11, { today });
eq('derive productModel', v11d.productModel, '500ml');
eq('derive mainBusiness', v11d.mainBusiness, '户外用品进口分销');
eq('emptyRecord productModel 空', schema.emptyRecord().productModel, '');
eq('emptyRecord mainBusiness 空', schema.emptyRecord().mainBusiness, '');
ok('产品库存在同名多型号', schema.DEFAULT_PRODUCTS.filter((p) => p.name === '不锈钢保温杯').length >= 2);

/* ---------------- v1.2 新增字段：邮箱 / 电话 / 地址 ---------------- */

const v12 = schema.normalize({
  firstInquiryAt: '2026-10-01',
  customerName: 'V12',
  email: 'm.braun@braun-trading.de',
  phone: '+49 711 4567 890',
  address: 'Industriestr. 12, 70565 Stuttgart, Germany'
});
eq('normalize email', v12.email, 'm.braun@braun-trading.de');
eq('normalize phone', v12.phone, '+49 711 4567 890');
eq('normalize address', v12.address, 'Industriestr. 12, 70565 Stuttgart, Germany');
const v12d = schema.derive(v12, { today });
eq('derive email', v12d.email, 'm.braun@braun-trading.de');
eq('derive phone', v12d.phone, '+49 711 4567 890');
eq('derive address', v12d.address, 'Industriestr. 12, 70565 Stuttgart, Germany');
eq('emptyRecord email 空', schema.emptyRecord().email, '');
eq('emptyRecord phone 空', schema.emptyRecord().phone, '');
eq('emptyRecord address 空', schema.emptyRecord().address, '');

ok('邮箱合法性：正常', schema.isValidEmail('a@b.co') === true);
ok('邮箱合法性：带点号/加号', schema.isValidEmail('first.last+crm@sub.example.com') === true);
ok('邮箱合法性：空值允许', schema.isValidEmail('') === true);
ok('邮箱合法性：缺 @', schema.isValidEmail('abc.example.com') === false);
ok('邮箱合法性：缺域名后缀', schema.isValidEmail('a@b') === false);
ok('邮箱合法性：含空格', schema.isValidEmail('a b@c.com') === false);
const emailBad = schema.validate({ firstInquiryAt: '2026-09-12', customerName: 'X', email: 'bad-email' });
eq('validate 拒绝非法邮箱', emailBad.ok, false);
ok('validate 邮箱错误字段', !!emailBad.errors.email);
const emailOk = schema.validate({ firstInquiryAt: '2026-09-12', customerName: 'X', email: 'x@y.com', phone: '+1 555 0100', address: 'Somewhere 1' });
eq('validate 通过合法联系方式', emailOk.ok, true);

/* ---------------- 汇总 ---------------- */

const total = pass + failures.length;
if (failures.length) {
  console.log('\n失败用例：');
  failures.forEach((f) => console.log(`  ✗ ${f.name}${f.detail ? ' — ' + f.detail : ''}`));
  console.log(`\nschema 测试：${pass}/${total} 通过，${failures.length} 失败`);
  process.exit(1);
} else {
  console.log(`schema 测试：${pass}/${total} 全部通过`);
}
