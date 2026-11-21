/*!
 * 客户通 CustomerCRM — tests/test-parser.mjs
 * 用法：node tests/test-parser.mjs   （在 customer-crm 目录下执行）
 * 仅使用 Node 内置模块；全部用例 PASS 时退出码 0，否则退出码 1。
 */
import assert from 'node:assert';
import { createRequire } from 'node:module';
import vm from 'node:vm';
import fs from 'node:fs';

const require = createRequire(import.meta.url);
const parser = require('../js/parser.js');
const schema = require('../js/schema.js');

const TODAY = '2026-11-21';
const OPTS = { today: TODAY };
const FIELD_KEYS = [
  'firstInquiryAt', 'customerName', 'email', 'phone', 'address',
  'product', 'productModel', 'country', 'companyName', 'mainBusiness',
  'foundedAt', 'teamSize', 'annualRevenue', 'followupStatus', 'dealStatus', 'coreReason'
];

/* 三段示例原文（与 parser.SAMPLE_TEXTS 必须完全一致；v1.2 追加联系方式） */
const TEXT_A = '2026年9月12日收到德国客户 Michael Braun 的询盘。他所在的 Braun Trading GmbH 成立于1998年，主营户外用品，团队规模约50-100人，年营收约1200万欧元。客户咨询的是我们2026款不锈钢保温杯 500ml，要求加印LOGO。目前已报价，客户反馈价格略高于目标价8%，仍在等终端客户确认。11月10日电话沟通确认2000pcs试单意向，11月18日邮件确认样品尺寸。邮箱 m.braun@braun-trading.de，电话 +49 711 4567 890，地址 Industriestr. 12, 70565 Stuttgart, Germany。';
const TEXT_B = 'Michael Braun from Braun Trading GmbH (Germany) reached out on Sep 12, 2026 regarding our 500ml stainless steel vacuum flask with custom logo printing. The company was founded in 1998, with a team of 50-100 employees and annual revenue around EUR 12 million. We sent our quotation on Nov 10; the client said our price is about 8% above their target. The deal was closed on Nov 20. Email: michael.braun@braun-trading.de, Tel: +49 711 4567 890, Address: Industriestr. 12, 70565 Stuttgart, Germany.';
const TEXT_C = '广交会名片：Mr. Ahmed Hassan, Cairo, Egypt。公司 El Nile Import & Export Co., Ltd，做家居用品进口，年采购额约 300 万美元，团队 15 人，2011 年成立。对我们的陶瓷餐具套装感兴趣，已发样册，等回复。邮箱 a.hassan@elnile-import.com，电话 +20 100 234 5678，地址 25 Nile Corniche, Cairo 11511, Egypt。';

/* ------------------------------ 迷你测试框架 ------------------------------ */

let passed = 0;
const failures = [];

function test(name, fn) {
  try {
    fn();
    passed += 1;
    console.log('PASS  ' + name);
  } catch (err) {
    failures.push({ name, err });
    console.log('FAIL  ' + name);
    const detail = String((err && err.message) || err).split('\n').join('\n      ');
    console.log('      ' + detail);
  }
}

function deepEqual(actual, expected, message) {
  assert.deepStrictEqual(actual, expected, message);
}

/* ------------------------------ 预跑结果 ------------------------------ */

const resultA = parser.recognize(parser.SAMPLE_TEXTS[0].text, OPTS);
const resultB = parser.recognize(parser.SAMPLE_TEXTS[1].text, OPTS);
const resultC = parser.recognize(parser.SAMPLE_TEXTS[2].text, OPTS);

/* ==================================================================== */
/* 1. 接口与样例文本                                                     */
/* ==================================================================== */

test('接口：VERSION / recognize / parseDate / SAMPLE_TEXTS 齐备', () => {
  assert.equal(parser.VERSION, '1.2.0');
  assert.equal(typeof parser.recognize, 'function');
  assert.equal(typeof parser.parseDate, 'function');
  assert.ok(Array.isArray(parser.SAMPLE_TEXTS));
  assert.ok(parser.SAMPLE_TEXTS.length >= 3);
});

test('样例：标签依次为 中文询盘 / 英文询盘 / 展会名片速记，原文与契约一致', () => {
  deepEqual(parser.SAMPLE_TEXTS.map((s) => s.label), ['中文询盘', '英文询盘', '展会名片速记']);
  assert.equal(parser.SAMPLE_TEXTS[0].text, TEXT_A);
  assert.equal(parser.SAMPLE_TEXTS[1].text, TEXT_B);
  assert.equal(parser.SAMPLE_TEXTS[2].text, TEXT_C);
});

test('结果结构：16 个字段 key 与 CRM_SCHEMA.FIELD_KEYS 完全一致', () => {
  deepEqual(FIELD_KEYS, schema.FIELD_KEYS);
  assert.equal(schema.FIELD_KEYS.length, 16);
  [resultA, resultB, resultC].forEach((r) => {
    Object.keys(r.fields).forEach((key) => {
      assert.ok(schema.FIELD_KEYS.indexOf(key) !== -1, '未知字段 key: ' + key);
    });
  });
});

/* ==================================================================== */
/* 2. 契约样例 A（中文）                                                 */
/* ==================================================================== */

test('A：首次询盘时间 = 2026-09-12', () => {
  assert.equal(resultA.fields.firstInquiryAt, '2026-09-12');
});

test('A：客户名称 = Michael Braun', () => {
  assert.equal(resultA.fields.customerName, 'Michael Braun');
});

test('A：公司名 = Braun Trading GmbH', () => {
  assert.equal(resultA.fields.companyName, 'Braun Trading GmbH');
});

test('A：国家 = 德国', () => {
  assert.equal(resultA.fields.country, '德国');
});

test('A：成立时间 ≈ 1998', () => {
  assert.ok(/^1998/.test(resultA.fields.foundedAt), 'foundedAt=' + resultA.fields.foundedAt);
});

test('A：团队规模 ≈ 50-100 人', () => {
  assert.ok(/50\s*-\s*100/.test(resultA.fields.teamSize), 'teamSize=' + resultA.fields.teamSize);
  assert.ok(/人/.test(resultA.fields.teamSize));
});

test('A：年营收 ≈ 1200 万欧元', () => {
  const v = resultA.fields.annualRevenue || '';
  assert.ok(/1200/.test(v) && /万/.test(v) && /欧元/.test(v), 'annualRevenue=' + v);
});

test('A：咨询产品含「保温杯」', () => {
  assert.ok((resultA.fields.product || '').indexOf('保温杯') !== -1, 'product=' + resultA.fields.product);
});

test('A：产品型号含 500ml', () => {
  assert.ok((resultA.fields.productModel || '').indexOf('500ml') !== -1, 'productModel=' + resultA.fields.productModel);
});

test('A：主营业务含「户外用品」且 ≤30 字', () => {
  const v = resultA.fields.mainBusiness || '';
  assert.ok(v.indexOf('户外用品') !== -1, 'mainBusiness=' + v);
  assert.ok(v.length <= 30, 'mainBusiness 超过 30 字：' + v.length);
});

test('A：跟进状态 = 已报价，成交状态 = 未成交', () => {
  assert.equal(resultA.fields.followupStatus, '已报价');
  assert.equal(resultA.fields.dealStatus, '未成交');
});

test('A：核心原因含「价格」且 ≤80 字', () => {
  const v = resultA.fields.coreReason || '';
  assert.ok(v.indexOf('价格') !== -1, 'coreReason=' + v);
  assert.ok(v.length <= 80, 'coreReason 超过 80 字：' + v.length);
});

test('A：跟进记录首条 2026-11-18，至少 2 条', () => {
  assert.ok(resultA.followups.length >= 2, 'followups=' + JSON.stringify(resultA.followups));
  assert.equal(resultA.followups[0].date, '2026-11-18');
});

test('A：联系方式全命中（email / phone / address），filled=16/16', () => {
  assert.equal(resultA.ok, true);
  assert.equal(resultA.fields.email, 'm.braun@braun-trading.de');
  assert.ok(/\+49/.test(resultA.fields.phone || ''), 'phone=' + resultA.fields.phone);
  assert.ok((resultA.fields.address || '').indexOf('Stuttgart') !== -1, 'address=' + resultA.fields.address);
  assert.equal(resultA.filled, 16);
  deepEqual(resultA.missing, []);
});

/* ==================================================================== */
/* 3. 契约样例 B（英文）                                                 */
/* ==================================================================== */

test('B：国家 = 德国，首次询盘 = 2026-09-12', () => {
  assert.equal(resultB.fields.country, '德国');
  assert.equal(resultB.fields.firstInquiryAt, '2026-09-12');
});

test('B：成立时间 = 1998', () => {
  assert.equal(resultB.fields.foundedAt, '1998');
});

test('B：团队规模含 50-100', () => {
  assert.ok((resultB.fields.teamSize || '').indexOf('50-100') !== -1, 'teamSize=' + resultB.fields.teamSize);
});

test('B：年营收含 12', () => {
  assert.ok((resultB.fields.annualRevenue || '').indexOf('12') !== -1, 'annualRevenue=' + resultB.fields.annualRevenue);
});

test('B：跟进状态 = 已成交，成交状态 = 已成交', () => {
  assert.equal(resultB.fields.followupStatus, '已成交');
  assert.equal(resultB.fields.dealStatus, '已成交');
});

test('B：跟进记录含 2026-11-20', () => {
  const dates = resultB.followups.map((f) => f.date);
  assert.ok(dates.indexOf('2026-11-20') !== -1, 'followups=' + JSON.stringify(dates));
});

test('B：产品型号含 500ml；文本无主营描述时不输出 mainBusiness', () => {
  assert.ok((resultB.fields.productModel || '').indexOf('500ml') !== -1, 'productModel=' + resultB.fields.productModel);
  assert.equal(resultB.fields.mainBusiness, undefined);
  assert.ok(resultB.missing.indexOf('mainBusiness') !== -1);
});

test('B：联系方式全命中（Email / Tel / Address），filled=15/16', () => {
  assert.equal(resultB.fields.email, 'michael.braun@braun-trading.de');
  assert.ok(/\+49/.test(resultB.fields.phone || ''), 'phone=' + resultB.fields.phone);
  assert.ok((resultB.fields.address || '').indexOf('Stuttgart') !== -1, 'address=' + resultB.fields.address);
  assert.equal(resultB.filled, 15);
  deepEqual(resultB.missing, ['mainBusiness']);
});

/* ==================================================================== */
/* 4. 契约样例 C（展会名片速记）                                         */
/* ==================================================================== */

test('C：国家 = 埃及', () => {
  assert.equal(resultC.fields.country, '埃及');
});

test('C：客户名含 Ahmed', () => {
  assert.ok((resultC.fields.customerName || '').indexOf('Ahmed') !== -1, 'customerName=' + resultC.fields.customerName);
});

test('C：公司名含 El Nile', () => {
  assert.ok((resultC.fields.companyName || '').indexOf('El Nile') !== -1, 'companyName=' + resultC.fields.companyName);
});

test('C：成立时间 = 2011，团队规模含 15，年营收含 300', () => {
  assert.ok(/^2011/.test(resultC.fields.foundedAt), 'foundedAt=' + resultC.fields.foundedAt);
  assert.ok((resultC.fields.teamSize || '').indexOf('15') !== -1, 'teamSize=' + resultC.fields.teamSize);
  assert.ok((resultC.fields.annualRevenue || '').indexOf('300') !== -1, 'annualRevenue=' + resultC.fields.annualRevenue);
});

test('C：产品含「陶瓷餐具」，跟进状态 ∈ {已回复, 寄样中}', () => {
  assert.ok((resultC.fields.product || '').indexOf('陶瓷餐具') !== -1, 'product=' + resultC.fields.product);
  assert.ok(['已回复', '寄样中'].indexOf(resultC.fields.followupStatus) !== -1,
    'followupStatus=' + resultC.fields.followupStatus);
});

test('C：主营业务含「家居用品」；产品无规格时不输出 productModel', () => {
  assert.ok((resultC.fields.mainBusiness || '').indexOf('家居用品') !== -1, 'mainBusiness=' + resultC.fields.mainBusiness);
  assert.equal(resultC.fields.productModel, undefined);
  assert.ok(resultC.missing.indexOf('productModel') !== -1);
});

test('C：联系方式全命中（邮箱/电话/地址），filled=12/16', () => {
  assert.equal(resultC.fields.email, 'a.hassan@elnile-import.com');
  assert.equal(resultC.fields.phone, '+20 100 234 5678');
  assert.ok((resultC.fields.address || '').indexOf('Nile Corniche') !== -1, 'address=' + resultC.fields.address);
  assert.equal(resultC.filled, 12);
  deepEqual(resultC.missing, ['firstInquiryAt', 'productModel', 'dealStatus', 'coreReason']);
});

test('C：没有明确成交线索时不输出 dealStatus', () => {
  assert.equal(resultC.fields.dealStatus, undefined);
});

/* ==================================================================== */
/* 5. 边界与稳健性                                                       */
/* ==================================================================== */

test('空文本：ok=false、warning code=empty、标准提示文案', () => {
  const r = parser.recognize('', OPTS);
  assert.equal(r.ok, false);
  deepEqual(r.fields, {});
  deepEqual(r.followups, []);
  assert.equal(r.warnings.length, 1);
  assert.equal(r.warnings[0].code, 'empty');
  assert.equal(r.warnings[0].message, '请粘贴一段询盘文本');
  assert.equal(r.filled, 0);
  assert.equal(r.missing.length, 16);
});

test('空文本：空白/换行/null/undefined 均不抛异常', () => {
  const inputs = ['   ', '\r\n\t', null, undefined];
  inputs.forEach((input) => {
    const r = parser.recognize(input, OPTS);
    assert.equal(r.ok, false);
    assert.equal(r.warnings[0].code, 'empty');
  });
});

test('全角标点与全角数字：日期/国家/产品可识别', () => {
  const text = '全角标点：２０２６／９／１２ 收到德国客户 Michael Braun 的询盘，咨询产品：不锈钢保温杯 500ml！';
  const r = parser.recognize(text, OPTS);
  assert.equal(r.fields.firstInquiryAt, '2026-09-12');
  assert.equal(r.fields.country, '德国');
  assert.ok((r.fields.product || '').indexOf('保温杯') !== -1, 'product=' + r.fields.product);
});

test("parseDate('2026/9/12') = 2026-09-12", () => {
  assert.equal(parser.parseDate('2026/9/12', OPTS), '2026-09-12');
});

test("parseDate('12 Sep 2026') = 2026-09-12", () => {
  assert.equal(parser.parseDate('12 Sep 2026', OPTS), '2026-09-12');
  assert.equal(parser.parseDate('Sep 12, 2026', OPTS), '2026-09-12');
});

test("parseDate('') = ''，无日期文本返回空串", () => {
  assert.equal(parser.parseDate('', OPTS), '');
  assert.equal(parser.parseDate('没有日期的文本', OPTS), '');
});

test("'not closed yet' → dealStatus=未成交", () => {
  const text = 'John Smith from ABC Trading LLC emailed us about LED desk lamps. Quotation sent on Oct 2. The deal is not closed yet; client is comparing suppliers.';
  const r = parser.recognize(text, OPTS);
  assert.equal(r.fields.dealStatus, '未成交');
  assert.notEqual(r.fields.followupStatus, '已成交');
});

test('多日期：倒序取最近 3 条跟进，且询盘日期不重复计入', () => {
  const text = '2026年1月5日收到美国客户询盘，1月8日已回复，2月1日发送报价，3月2日寄出样品，4月20日客户下单。';
  const r = parser.recognize(text, OPTS);
  assert.equal(r.fields.firstInquiryAt, '2026-01-05');
  deepEqual(r.followups.map((f) => f.date), ['2026-04-20', '2026-03-02', '2026-02-01']);
});

test('无国家文本：不输出 country，并计入 missing', () => {
  const text = 'Michael Braun from Braun Trading GmbH sent an inquiry about 500ml stainless steel vacuum flask.';
  const r = parser.recognize(text, OPTS);
  assert.equal(r.fields.country, undefined);
  assert.ok(r.missing.indexOf('country') !== -1);
});

test('仅公司无联系人：识别公司/国家/成立，不编造客户名', () => {
  const text = '公司名称：Shenzhen ABC Technology Co., Ltd，位于中国深圳，成立于2015年。';
  const r = parser.recognize(text, OPTS);
  assert.ok((r.fields.companyName || '').indexOf('Shenzhen') !== -1, 'companyName=' + r.fields.companyName);
  assert.equal(r.fields.country, '中国');
  assert.ok(/^2015/.test(r.fields.foundedAt), 'foundedAt=' + r.fields.foundedAt);
  assert.equal(r.fields.customerName, undefined);
});

test('中文联系人：联系人：王小明 → customerName=王小明', () => {
  const text = '联系人：王小明，来自宁波，需要采购LED灯具 5000pcs。';
  const r = parser.recognize(text, OPTS);
  assert.equal(r.fields.customerName, '王小明');
  assert.ok((r.fields.product || '').indexOf('LED') !== -1, 'product=' + r.fields.product);
});

test('成立日期是完整日期时，不抢走首次询盘时间', () => {
  const text = '该公司成立于2010年5月20日，2026年9月12日收到客户询盘，客户 Michael Braun 咨询不锈钢保温杯。';
  const r = parser.recognize(text, OPTS);
  assert.equal(r.fields.firstInquiryAt, '2026-09-12');
  assert.ok(/^2010/.test(r.fields.foundedAt), 'foundedAt=' + r.fields.foundedAt);
  assert.ok(r.followups.every((f) => f.date !== '2010-05-20'));
});

test('无年份日期：用 opts.today 的年份兜底', () => {
  const text = '客户 11月10日来函询问不锈钢保温杯的报价。';
  const r = parser.recognize(text, OPTS);
  assert.equal(r.fields.firstInquiryAt, '2026-11-10');
  assert.equal(r.followups[0].date, '2026-11-10');
});

test('全英文新询盘：默认跟进状态 = 新询盘，不输出 dealStatus', () => {
  const text = 'Michael Braun from Braun Trading GmbH sent an inquiry about 500ml stainless steel vacuum flask, no reply yet.';
  const r = parser.recognize(text, OPTS);
  assert.equal(r.fields.followupStatus, '新询盘');
  assert.equal(r.fields.dealStatus, undefined);
});

/* ==================================================================== */
/* 5b. v1.1 新增字段：主营业务 mainBusiness / 产品型号 productModel       */
/* ==================================================================== */

test('mainBusiness：主营户外用品（中文标签规则）', () => {
  const r = parser.recognize('他所在的 Braun Trading GmbH 成立于 1998 年，主营户外用品，团队 30 人。', OPTS);
  assert.ok((r.fields.mainBusiness || '').indexOf('户外用品') !== -1, 'mainBusiness=' + r.fields.mainBusiness);
});

test('mainBusiness：做家居用品进口（贸易动词规则）', () => {
  const r = parser.recognize('客户公司做家居用品进口，年采购额约 200 万美元。', OPTS);
  assert.ok((r.fields.mainBusiness || '').indexOf('家居用品') !== -1, 'mainBusiness=' + r.fields.mainBusiness);
  assert.ok((r.fields.mainBusiness || '').indexOf('进口') !== -1, 'mainBusiness=' + r.fields.mainBusiness);
});

test('mainBusiness：是一家…进口商（身份名词规则）', () => {
  const r = parser.recognize('客户是一家户外用品进口商，主要在德国市场销售。', OPTS);
  assert.ok((r.fields.mainBusiness || '').indexOf('户外用品') !== -1, 'mainBusiness=' + r.fields.mainBusiness);
});

test('mainBusiness：英文 distributor of / specialized in', () => {
  const r1 = parser.recognize('We are a distributor of outdoor products and camping gear. Please quote for 500ml vacuum flasks.', OPTS);
  assert.ok((r1.fields.mainBusiness || '').indexOf('outdoor products') !== -1, 'mainBusiness=' + r1.fields.mainBusiness);
  assert.ok((r1.fields.mainBusiness || '').length <= 30, 'mainBusiness 超过 30 字：' + r1.fields.mainBusiness);
  const r2 = parser.recognize('Nordic Living AB is specialized in home decorations and kitchenware, please quote 500ml glass jars.', OPTS);
  assert.ok((r2.fields.mainBusiness || '').indexOf('home decorations') !== -1, 'mainBusiness=' + r2.fields.mainBusiness);
  assert.equal(r2.fields.mainBusiness, 'home decorations', '30 字截断后不应留下悬空 and：' + r2.fields.mainBusiness);
});

test('mainBusiness：输出 ≤30 字', () => {
  const r = parser.recognize('客户公司主要从事家居用品、厨房用品、户外用品、五金工具、电子产品、玩具文具的批发与分销业务。', OPTS);
  const v = r.fields.mainBusiness || '';
  assert.ok(v.length > 0 && v.length <= 30, 'mainBusiness 长度 ' + v.length + '：' + v);
});

test('mainBusiness：无主营描述时不输出', () => {
  const r = parser.recognize('Michael Braun 询问不锈钢保温杯价格，需要 500ml 规格。', OPTS);
  assert.equal(r.fields.mainBusiness, undefined);
  assert.ok(r.missing.indexOf('mainBusiness') !== -1);
});

test('mainBusiness：干扰句（我司/同行/竞争对手）不误判', () => {
  const r1 = parser.recognize('我们公司主营咖啡机出口。客户是家居用品分销商，询问陶瓷餐具价格。', OPTS);
  assert.ok((r1.fields.mainBusiness || '').indexOf('家居用品') !== -1, 'mainBusiness=' + r1.fields.mainBusiness);
  assert.ok((r1.fields.mainBusiness || '').indexOf('咖啡机') === -1, '不应采纳我司业务：' + r1.fields.mainBusiness);
  const r2 = parser.recognize('竞争对手主营同类产品，客户主营户外用品，咨询保温杯。', OPTS);
  assert.ok((r2.fields.mainBusiness || '').indexOf('户外用品') !== -1, 'mainBusiness=' + r2.fields.mainBusiness);
});

test('mainBusiness 边界：不吞公司名与咨询产品', () => {
  const r = parser.recognize('Braun Trading GmbH 主营户外用品，客户咨询的是保温杯。', OPTS);
  assert.equal(r.fields.mainBusiness, '户外用品');
  assert.equal(r.fields.companyName, 'Braun Trading GmbH');
  assert.equal(r.fields.product, '保温杯');
});

test('productModel：16 件套', () => {
  const r = parser.recognize('客户咨询 16 件套陶瓷餐具套装，需要报价。', OPTS);
  assert.equal(r.fields.productModel, '16 件套');
});

test('productModel：500ml 与 12 pcs 并存', () => {
  const r = parser.recognize('客户需要采购 500ml 不锈钢保温杯 12 pcs，请报价。', OPTS);
  const v = r.fields.productModel || '';
  assert.ok(v.indexOf('500ml') !== -1, 'productModel=' + v);
  assert.ok(v.indexOf('12 pcs') !== -1, 'productModel=' + v);
});

test('productModel：750ml / 1.2L 容量规格', () => {
  const r1 = parser.recognize('客户需要 750ml 不锈钢保温杯，请报价。', OPTS);
  assert.equal(r1.fields.productModel, '750ml');
  const r2 = parser.recognize('客户需要 1.2L 大容量保温壶，请报价。', OPTS);
  assert.equal(r2.fields.productModel, '1.2L');
  const r3 = parser.recognize('Please quote 1.2L glass jars.', OPTS);
  assert.equal(r3.fields.productModel, '1.2L');
});

test('productModel：3-piece set', () => {
  const r = parser.recognize('客户对 3-piece set 陶瓷餐具感兴趣。', OPTS);
  assert.equal(r.fields.productModel, '3-piece set');
});

test('productModel：无型号时不输出', () => {
  const r = parser.recognize('客户对我们的陶瓷餐具套装感兴趣。', OPTS);
  assert.equal(r.fields.productModel, undefined);
  assert.ok(r.missing.indexOf('productModel') !== -1);
});

test('productModel：显式「型号：XXX」标注兜底', () => {
  const r = parser.recognize('产品：陶瓷餐具，型号：TC-500，请报价。', OPTS);
  assert.equal(r.fields.productModel, 'TC-500');
});

test('v1.2：filled + missing.length = 16，各版本新字段带 evidence', () => {
  [resultA, resultB, resultC].forEach((r) => {
    assert.equal(r.filled + r.missing.length, 16);
  });
  assert.ok(resultA.evidence.productModel && resultA.evidence.mainBusiness);
  assert.ok(resultB.evidence.productModel);
  assert.ok(resultC.evidence.mainBusiness);
});

/* ==================================================================== */
/* 5c. v1.2 新增字段：联系方式 email / phone / address                   */
/* ==================================================================== */

test('email：中文标签「邮箱：」并带 evidence', () => {
  const r = parser.recognize('联系人 Michael Braun，邮箱：michael.braun@braun-trading.de，请报价。', OPTS);
  assert.equal(r.fields.email, 'michael.braun@braun-trading.de');
  assert.ok(r.evidence.email && r.evidence.email.snippet.indexOf('michael.braun') !== -1,
    'evidence=' + JSON.stringify(r.evidence.email));
});

test('email：英文标签 Email: 且允许 + 号', () => {
  const r = parser.recognize('Please email Sarah at Email: sarah.johnson+nordic@gmail.com for details.', OPTS);
  assert.equal(r.fields.email, 'sarah.johnson+nordic@gmail.com');
});

test('email：裸邮箱可识别；多个邮箱只取第一个', () => {
  const r1 = parser.recognize('请把报价发到 info@abc-trading.com 谢谢。', OPTS);
  assert.equal(r1.fields.email, 'info@abc-trading.com');
  const r2 = parser.recognize('邮箱：first@x.com，备用 second@y.com。', OPTS);
  assert.equal(r2.fields.email, 'first@x.com');
});

test('phone：国际格式（Tel: +49 711 4567 890）保留原始格式', () => {
  const r = parser.recognize('Tel: +49 711 4567 890，请联系我们。', OPTS);
  assert.equal(r.fields.phone, '+49 711 4567 890');
  assert.ok(r.evidence.phone && r.evidence.phone.snippet.indexOf('+49') !== -1,
    'evidence=' + JSON.stringify(r.evidence.phone));
});

test('phone：(020) 7946 0123 与 0086-13800138000', () => {
  const r1 = parser.recognize('电话：(020) 7946 0123。', OPTS);
  assert.equal(r1.fields.phone, '(020) 7946 0123');
  const r2 = parser.recognize('联系方式 0086-13800138000。', OPTS);
  assert.equal(r2.fields.phone, '0086-13800138000');
});

test('phone：中文手机号 13800138000', () => {
  const r = parser.recognize('手机：13800138000。', OPTS);
  assert.equal(r.fields.phone, '13800138000');
});

test('phone：WhatsApp: +1 503 555 0142', () => {
  const r = parser.recognize('WhatsApp: +1 503 555 0142', OPTS);
  assert.equal(r.fields.phone, '+1 503 555 0142');
});

test('phone 防误判：日期 / 金额 / 订单号都不是电话', () => {
  const cases = [
    '2026-09-12 收到询盘，日期 2026年9月12日。',
    '年营收约1200万欧元，采购额 12 million USD。',
    '订单号 PO 20261122 已确认。'
  ];
  cases.forEach((text) => {
    const r = parser.recognize(text, OPTS);
    assert.equal(r.fields.phone, undefined, '误判电话：' + text + ' => ' + r.fields.phone);
  });
});

test('address：中文公司地址标签（去标签词）', () => {
  const r = parser.recognize('公司地址：浙江省宁波市鄞州区天童南路88号，请寄样品。', OPTS);
  assert.equal(r.fields.address, '浙江省宁波市鄞州区天童南路88号');
});

test('address：位于 + 省市路号', () => {
  const r = parser.recognize('公司位于广东省深圳市南山区科技园路1号。', OPTS);
  assert.equal(r.fields.address, '广东省深圳市南山区科技园路1号');
});

test('address：英文 Address / Office 标签', () => {
  const r1 = parser.recognize('Our Address: Industriestr. 12, 70565 Stuttgart, Germany', OPTS);
  assert.ok((r1.fields.address || '').indexOf('Stuttgart') !== -1, 'address=' + r1.fields.address);
  const r2 = parser.recognize('Office: 12 Baker Street, London, UK', OPTS);
  assert.ok((r2.fields.address || '').indexOf('Baker Street') !== -1, 'address=' + r2.fields.address);
});

test('address：邮箱地址里的「地址」不算通讯地址', () => {
  const r = parser.recognize('邮箱地址：a@b.com', OPTS);
  assert.equal(r.fields.email, 'a@b.com');
  assert.equal(r.fields.address, undefined);
});

test('phone 防误判：邮箱里的数字不算电话', () => {
  const r = parser.recognize('邮箱：13800138000@qq.com，请报价。', OPTS);
  assert.equal(r.fields.email, '13800138000@qq.com');
  assert.equal(r.fields.phone, undefined);
});

test('address：中文标签 + 英文地址（含逗号）', () => {
  const r = parser.recognize('公司地址：Industriestr. 12, 70565 Stuttgart, Germany', OPTS);
  assert.ok((r.fields.address || '').indexOf('Stuttgart') !== -1, 'address=' + r.fields.address);
});

test('address：≤200 字且不含标签词', () => {
  const r = parser.recognize('公司地址：' + '浙'.repeat(300), OPTS);
  const v = r.fields.address || '';
  assert.ok(v.length > 0 && v.length <= 200, 'address 长度 ' + v.length);
  assert.ok(v.indexOf('公司地址') === -1, 'address 不应含标签词：' + v);
});

test('无联系方式：email/phone/address 均不输出并计入 missing', () => {
  const r = parser.recognize('Michael Braun 询问不锈钢保温杯价格。', OPTS);
  assert.equal(r.fields.email, undefined);
  assert.equal(r.fields.phone, undefined);
  assert.equal(r.fields.address, undefined);
  ['email', 'phone', 'address'].forEach((k) => {
    assert.ok(r.missing.indexOf(k) !== -1, k + ' 未计入 missing');
  });
});

test('v1.2：三种联系方式同时识别，filled+missing=16 且都有 evidence', () => {
  const r = parser.recognize('邮箱：a@b.com，电话：13800138000，地址：浙江省宁波市鄞州区天童南路88号。', OPTS);
  assert.equal(r.fields.email, 'a@b.com');
  assert.equal(r.fields.phone, '13800138000');
  assert.equal(r.fields.address, '浙江省宁波市鄞州区天童南路88号');
  assert.ok(r.evidence.email && r.evidence.phone && r.evidence.address);
  assert.equal(r.filled + r.missing.length, 16);
});

/* ==================================================================== */
/* 6. 结果质量约束                                                       */
/* ==================================================================== */

test('只输出非空字段；filled/missing 与 fields 一致', () => {
  const all = [
    parser.recognize(TEXT_A, OPTS),
    parser.recognize(TEXT_B, OPTS),
    parser.recognize(TEXT_C, OPTS),
    parser.recognize('随便一段没有明确字段的文本。', OPTS)
  ];
  all.forEach((r) => {
    const keys = Object.keys(r.fields);
    keys.forEach((k) => {
      assert.equal(typeof r.fields[k], 'string');
      assert.ok(r.fields[k].trim().length > 0, '字段 ' + k + ' 为空串');
    });
    assert.equal(r.filled, keys.length);
    assert.equal(r.missing.length, FIELD_KEYS.length - keys.length);
    assert.equal(r.filled + r.missing.length, FIELD_KEYS.length);
    keys.forEach((k) => { assert.ok(r.missing.indexOf(k) === -1); });
    r.missing.forEach((k) => { assert.ok(FIELD_KEYS.indexOf(k) !== -1); });
  });
});

test('evidence：snippet ≤60 字、不含换行、来自原文，value 与字段一致', () => {
  [['A', TEXT_A, resultA], ['B', TEXT_B, resultB], ['C', TEXT_C, resultC]].forEach((item) => {
    const name = item[0];
    const raw = item[1];
    const r = item[2];
    const flat = raw.replace(/[\r\n\t]+/g, ' ');
    Object.keys(r.evidence).forEach((key) => {
      const ev = r.evidence[key];
      assert.equal(ev.value, r.fields[key], name + '.' + key + ' evidence.value 不一致');
      assert.equal(typeof ev.snippet, 'string');
      assert.ok(ev.snippet.length > 0 && ev.snippet.length <= 60, name + '.' + key + ' snippet 长度 ' + ev.snippet.length);
      assert.ok(!/[\r\n]/.test(ev.snippet), name + '.' + key + ' snippet 含换行');
      assert.ok(flat.indexOf(ev.snippet) !== -1, name + '.' + key + ' snippet 不在原文中：' + ev.snippet);
    });
    Object.keys(r.fields).forEach((key) => {
      if (key === 'followupStatus' && r.fields[key] === '新询盘') { return; }
      assert.ok(r.evidence[key], name + ' 缺少 evidence: ' + key);
    });
  });
});

test('followups：≤3 条、严格倒序、日期合法、note 为字符串且不带日期前缀', () => {
  const cases = [
    parser.recognize(TEXT_A, OPTS),
    parser.recognize(TEXT_B, OPTS),
    parser.recognize('2026年1月5日收到询盘，1月8日已回复，2月1日发送报价，3月2日寄出样品，4月20日客户下单。', OPTS)
  ];
  cases.forEach((r) => {
    assert.ok(r.followups.length <= 3);
    for (let i = 0; i < r.followups.length; i += 1) {
      const f = r.followups[i];
      assert.ok(/^\d{4}-\d{2}-\d{2}$/.test(f.date), '非法日期 ' + f.date);
      assert.ok(schema.isValidISO(f.date), '日期不合法 ' + f.date);
      assert.equal(typeof f.note, 'string');
      assert.ok(!/^\d{4}\s*[-/.年]/.test(f.note), 'note 不应以日期开头：' + f.note);
      if (i > 0) { assert.ok(r.followups[i - 1].date > f.date, 'followups 未倒序'); }
    }
  });
});

test('无年份日期：opts.defaultYear 优先于点 today 的年份', () => {
  const text = '客户 11月10日来函询问不锈钢保温杯的报价。';
  const r = parser.recognize(text, { today: TODAY, defaultYear: 2024 });
  assert.equal(r.fields.firstInquiryAt, '2024-11-10');
});

test('recognize 是纯函数：重复调用结果完全一致', () => {
  [TEXT_A, TEXT_B, TEXT_C].forEach((text) => {
    const first = JSON.stringify(parser.recognize(text, OPTS));
    const second = JSON.stringify(parser.recognize(text, OPTS));
    assert.equal(first, second);
  });
});

test('10 万字文本：识别耗时 < 200ms', () => {
  let big = '';
  while (big.length < 100000) { big += TEXT_A + '\n'; }
  big = big.slice(0, 100000);
  const start = process.hrtime.bigint();
  const r = parser.recognize(big, OPTS);
  const elapsed = Number(process.hrtime.bigint() - start) / 1e6;
  assert.ok(elapsed < 200, '10 万字耗时 ' + elapsed.toFixed(1) + 'ms，超过 200ms');
  assert.ok(r.filled >= 8, '超长文本识别结果异常：filled=' + r.filled);
});

test('recognize 不产生 console.log 噪音', () => {
  const originalLog = console.log;
  let logs = 0;
  console.log = () => { logs += 1; };
  try {
    parser.SAMPLE_TEXTS.forEach((s) => parser.recognize(s.text, OPTS));
    parser.recognize('', OPTS);
    parser.parseDate('2026/9/12', OPTS);
  } finally {
    console.log = originalLog;
  }
  assert.equal(logs, 0);
});

test('UMD：浏览器式加载（无 require、有 window.CRM_SCHEMA）挂到 window.CrmParser', () => {
  const sandbox = { console };
  sandbox.window = sandbox;
  sandbox.self = sandbox;
  vm.createContext(sandbox);
  const readSource = (relative) => fs.readFileSync(new URL(relative, import.meta.url), 'utf8');
  vm.runInContext(readSource('../js/schema.js'), sandbox, { filename: 'schema.js' });
  vm.runInContext(readSource('../js/parser.js'), sandbox, { filename: 'parser.js' });
  assert.equal(typeof sandbox.CrmParser, 'object');
  assert.equal(typeof sandbox.CrmParser.recognize, 'function');
  assert.ok(sandbox.CrmParser.SAMPLE_TEXTS.length >= 3);
  const r = sandbox.CrmParser.recognize(TEXT_A, OPTS);
  assert.equal(r.filled, 16);
  assert.equal(r.fields.country, '德国');
});

/* ------------------------------ 汇总 ------------------------------ */

const total = passed + failures.length;
console.log('');
console.log('SUMMARY: ' + passed + '/' + total + ' passed, ' + failures.length + ' failed');
if (failures.length) {
  console.log('FAILED:');
  failures.forEach((f) => console.log('  - ' + f.name));
  process.exitCode = 1;
}
