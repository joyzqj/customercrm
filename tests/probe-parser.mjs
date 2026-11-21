import { createRequire } from 'node:module';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const require = createRequire(import.meta.url);
const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const schema = require(path.join(ROOT, 'js', 'schema.js'));
globalThis.CRM_SCHEMA = schema;
const parser = require(path.join(ROOT, 'js', 'parser.js'));

const sampleA = '2026年9月12日收到德国客户 Michael Braun 的询盘。他所在的 Braun Trading GmbH 成立于1998年，主营户外用品，团队规模约50-100人，年营收约1200万欧元。客户咨询的是我们2026款不锈钢保温杯 500ml，要求加印LOGO。目前已报价，客户反馈价格略高于目标价8%，仍在等终端客户确认。11月10日电话沟通确认2000pcs试单意向，11月18日邮件确认样品尺寸。';
const r = parser.recognize(sampleA, { today: '2026-11-22' });
console.log('ok:', r.ok, 'filled:', r.filled, 'missing:', r.missing);
console.log(JSON.stringify(r.fields, null, 2));
console.log('followups:', JSON.stringify(r.followups, null, 2));
console.log('warnings:', JSON.stringify(r.warnings));
if (parser.SAMPLE_TEXTS) console.log('samples:', parser.SAMPLE_TEXTS.map((s) => s.label));
