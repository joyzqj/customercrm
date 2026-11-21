# 客户通 CustomerCRM — 工程契约 v1.0

> 本文件是团队唯一的接口真源（source of truth）。任何接口变更必须先改本文件，并在任务板留痕。

## 0. 项目目标

一套**免部署、零后端、数据留在本机**的外贸客户管理（CRM）Web 系统，视觉与交互对标同目录下的 `trade-docs`（单证通 TradeDocs），最终部署到 GitHub Pages。

核心能力：

1. 客户台账：首次询盘时间、客户名称、咨询产品、国家、公司名、成立时间、团队规模、年营收、跟进状态、是否成交、核心原因、距上次跟进时间（记录最近三次跟进）。
2. 一键导出 Excel（`.xlsx`，真实可打开、带表头样式与筛选）。
3. 一键识别一整段中文/英文询盘文本，自动填入对应表格字段。
4. 仪表盘统计、搜索筛选排序、编辑/删除、跟进记录增删、JSON 备份与恢复、示例数据。

## 1. 目录与文件所有权（写作用域，互不重叠）

项目根目录：`C:\Users\Administrator\Desktop\Global Country\customer-crm\`

| 文件 / 目录 | 负责人 | 说明 |
| --- | --- | --- |
| `index.html` | Lead | 应用外壳（静态骨架 + script 顺序） |
| `js/core.js` | Lead | DOM/日期/转义/下载/Toast 等基础工具 |
| `js/schema.js` | Lead | **共享数据契约**：字段、枚举、国家表、派生计算 |
| `js/store.js` | Lead | localStorage 仓储、设置、示例数据、导入导出 |
| `js/app.js` | Lead | 路由、页面渲染、表单、表格、识别页、仪表盘、设置 |
| `js/parser.js` | parser-dev | 文本识别引擎（纯函数，可在 Node 中测试） |
| `js/xlsx.js` | export-dev | 自写 XLSX 生成器 + 导出/下载（可在 Node 中测试） |
| `css/app.css` | ui-dev | 全站样式（含响应式与 `@media print`） |
| `README.md`、`.github/**`、`.gitignore`、`.nojekyll`、`打开客户通.cmd` | Lead | 文档与部署 |
| `tests/cdp.mjs`、`tests/test-schema.mjs`、`tests/test-parser.mjs`(parser-dev)、`tests/test-xlsx.mjs`(export-dev)、`tests/e2e.mjs`(qa-dev) | 见括号 | 测试 |
| `screenshots/**`、`tests/QA_REPORT.md` | qa-dev | 验收产物 |
| `dist/**` | Lead | 单文件构建产物（可选） |

**禁止**修改他人作用域内的文件。需要别人改动时，发消息说明「文件 + 需要的行为 + 理由」。

## 2. 硬性技术约束

1. **零构建、零依赖、零后端**：纯静态文件，浏览器直接打开（`file://` 双击 `index.html` 也要能用），也能被 GitHub Pages 托管。
2. **禁止任何外链资源**（CDN、Google Fonts、远程图标一律不要）。图标用内联 SVG，字体用系统字体。
3. 所有文件 **UTF-8 无 BOM**，`<meta charset="utf-8">`。中文标点正常显示。
4. 数据存 `localStorage`，键名 `customer-crm.v1`；首次访问自动写入示例数据（可清空）。
5. 脚本使用传统 `<script src>` 顺序加载，**不要 ESM**；每个模块挂到 `window` 上（`window.CRM_SCHEMA` / `window.CrmParser` / `window.CrmXlsx` / `window.CrmStore` …）。同时提供 Node 导出以便测试：

```js
(function (root, factory) {
  var api = factory();
  if (typeof module === 'object' && module.exports) { module.exports = api; }
  root.CrmParser = api;            // 浏览器
})(typeof globalThis !== 'undefined' ? globalThis : this, function () { /* ... */ return api; });
```

6. 语法：ES2019 即可（`const/let`、箭头函数、模板串、`Array.prototype.flat` 可用）。不需要兼容 IE。不要用可选链之外的新特性也无妨，但 `?.` 与 `??` 可用。
7. 无 `alert()/confirm()` 作为主要交互（用页面内 Toast/Modal；`confirm` 只允许用于删除二次确认）。
8. 可访问性：交互元素有可读 `aria-label` 或文本；`<table>` 有 `<thead>`；Modal 有 `role="dialog" aria-modal="true"`。
9. 行数不是问题，**正确性和可验证性优先**。不要留 `console.log` 调试噪音（可保留 `console.warn` 用于可恢复的异常）。
10. 环境：Windows + PowerShell；`node` v24 可用，**没有 npm**；git 不在 PATH（`C:\Program Files\Git\cmd\git.exe`，只有 Lead 用）；Chrome 在 `C:\Program Files\Google\Chrome\Application\chrome.exe`；HTTP 代理 `http://127.0.0.1:10090`（本项目不需要联网）。Python 路径通过 `load_workspace_dependencies` 工具获取（含 openpyxl 3.1.5，可用于校验 xlsx）。

## 3. 数据模型

### 3.1 客户记录 CustomerRecord

```js
{
  id: 'c_lq2x8f3a',                 // 唯一 id，前缀 c_
  firstInquiryAt: '2026-09-12',     // 首次询盘时间 YYYY-MM-DD
  customerName: 'Michael Braun',    // 客户名称（联系人）
  product: '不锈钢保温杯 500ml',      // 咨询产品
  country: '德国',                   // 国家（中文名，来自 schema 国家表或自由输入）
  companyName: 'Braun Trading GmbH',// 公司名
  foundedAt: '1998',                // 成立时间（自由文本：1998 / 1998年3月 / 1998-03）
  teamSize: '50-100 人',            // 团队规模（自由文本）
  annualRevenue: '1200 万欧元',      // 年营收（自由文本）
  followupStatus: '已报价',          // 跟进状态（枚举，见 3.2）
  dealStatus: '未成交',              // 是否成交（枚举：已成交 / 未成交）
  coreReason: '价格高出目标价 8%，等待终端客户确认', // 核心原因（自由文本，可空）
  followups: [                      // 最近跟进记录（最新在前，最多 3 条）
    { date: '2026-11-18', note: '邮件确认样品尺寸，客户要求加印 LOGO' },
    { date: '2026-11-10', note: '电话沟通，确认 2000pcs 试单意向' },
    { date: '2026-10-28', note: '首次邮件回复，发送价目表与样册' }
  ],
  createdAt: '2026-09-12T02:10:00.000Z',
  updatedAt: '2026-11-18T09:30:00.000Z'
}
```

「距上次跟进的时间」不落库，由 `CRM_SCHEMA.derive()` 从 `followups[0].date` 实时计算；无跟进记录时视为「尚未跟进」。

### 3.2 枚举

```js
FOLLOWUP_STATUS = ['新询盘', '已回复', '已报价', '寄样中', '谈判中', '已成交', '暂停跟进', '已流失'];
DEAL_STATUS     = ['已成交', '未成交'];
```

状态色调（徽章 class 后缀）：新询盘→`info`、已回复→`neutral`、已报价→`brand`、寄样中→`warn`、谈判中→`warn`、已成交→`ok`、暂停跟进→`neutral`、已流失→`danger`；已成交→`ok`、未成交→`neutral`（由 `CRM_SCHEMA.statusTone/dealTone` 提供）。

### 3.3 跟进新鲜度

- `daysSince <= 6` → `ok`（绿）；`7 ~ 13` → `warn`（琥珀）；`>= 14` → `danger`（红）；无跟进 → `none`（灰）。
- 阈值来自设置 `settings.warnDays = 7`、`settings.staleDays = 14`。

## 4. `js/schema.js` 公共 API（Lead 提供，勿改）

```js
window.CRM_SCHEMA = {
  VERSION: '1.0.0',
  STORAGE_KEY: 'customer-crm.v1',

  FIELDS: [{ key, label, type, required, table, excel, width, hint, options? }],  // 11 个字段，见下
  FIELD_GROUPS: [{ title, hint, keys: [...] }],   // 表单分组
  FOLLOWUP_STATUS, DEAL_STATUS, STATUS_TONE, DEAL_TONE,
  COUNTRIES: [{ name:'德国', iso:'DE', en:'Germany', aliases:['german','deutschland'] }],

  emptyRecord(todayISO?) -> CustomerRecord,          // 空白记录（含默认值）
  normalize(raw) -> CustomerRecord,                  // 清洗/补全/截断 followups/排序
  derive(record, opts?) -> record & {
    lastFollowupAt, daysSince, inquiryDays, isDeal, isOpen, isOverdue,
    last3: [{ date, note, daysAgo }], ageTone, statusTone, dealTone
  },
  fieldByKey(key) -> field | null,
  fieldLabel(key) -> string,

  todayISO() -> 'YYYY-MM-DD',
  addDays(iso, n) -> 'YYYY-MM-DD',
  daysBetween(fromISO, toISO) -> number,
  parseLooseDate(text, opts?) -> 'YYYY-MM-DD' | '',   // 中英文各种写法
  formatShort(iso) -> '11-18',                        // 无年份时补 year 参数见实现
  formatDateTime(v) -> '2026-11-18 09:30',
  statusTone(s), dealTone(s), ageTone(days),
  findCountry(text) -> countryEntry | null,
  countryFlag(name) -> '🇩🇪' | '',
  flagFromISO(iso) -> '🇩🇪' | '',
  uid(prefix?) -> string
};
```

11 个字段（**顺序即界面、表单、Excel 列顺序**）：

| key | label | type | 说明 |
| --- | --- | --- | --- |
| `firstInquiryAt` | 首次询盘时间 | `date` | 必填，`YYYY-MM-DD` |
| `customerName` | 客户名称 | `text` | 必填，联系人名 |
| `product` | 咨询产品 | `text` | 产品/型号/规格 |
| `country` | 国家 | `select`（可自由输入） | 中文国名 |
| `companyName` | 公司名 | `text` | 客户公司全称 |
| `foundedAt` | 成立时间 | `text` | 1998 / 1998年3月 |
| `teamSize` | 团队规模 | `text` | 50-100 人 |
| `annualRevenue` | 年营收 | `text` | 1200 万欧元 |
| `followupStatus` | 跟进状态 | `select` | 枚举 |
| `dealStatus` | 是否成交 | `select` | 枚举 |
| `coreReason` | 核心原因 | `textarea` | 成败关键原因 |

## 5. `js/parser.js` 契约（parser-dev）

### 5.1 接口

```js
window.CrmParser = {
  VERSION: '1.0.0',
  recognize(text, opts) -> {
    ok: boolean,
    fields: { [fieldKey]: string },          // 只放识别到的字段（非空），值已清洗
    followups: [{ date: 'YYYY-MM-DD', note: string }],  // 最新在前，最多 3 条
    evidence: { [fieldKey]: { value, snippet } },       // snippet ≤ 60 字，用于界面对照
    warnings: [{ code, message }],
    filled: number,                          // fields 非空数量（不含 followups）
    missing: string[],                       // 11 个字段里没识别到的 key
    raw?: { textLength, countryHit, matchedRules: string[] }   // 调试信息，可省略字段内容
  },
  // 供界面「示例」按钮使用，至少 3 段，中英混合
  SAMPLE_TEXTS: [{ label: '中文询盘', text: '...' }, { label: '英文询盘', text: '...' }, { label: '展会名片速记', text: '...' }],
  // 可选：单字段识别，便于单测
  parseDate(text, opts) -> 'YYYY-MM-DD'|''
};
```

`opts`：`{ today: 'YYYY-MM-DD', defaultYear: number }`。`recognize` 必须是**纯函数**（同样输入 → 同样输出，除 `today` 依赖外）。

### 5.2 字段识别规则（最低要求）

| 字段 | 识别线索（中/英） |
| --- | --- |
| `firstInquiryAt` | 「询盘时间/首次询盘/收到询盘/来函/第一次联系/询价日期」+ 日期；`inquiry on/inquiry dated/reached out on/contacted us on/first email on`；无标签时取文中最靠前的日期 |
| `customerName` | 「客户/联系人/姓名/负责人」+ 名字；`contact person/Mr./Ms./Mrs./Dr./from ...`、英文 `First Last` 人名模式；中文姓名 2-4 字 |
| `product` | 「咨询产品/产品/询价产品/需要/采购/感兴趣的是」+ 名词短语；`inquiry about/regarding/interested in/product/asked for/quotation for`；⚠️ 不要吞掉时间/公司描述 |
| `country` | schema 国家表匹配（中文名、英文名、形容词如 German/Italian、`(DE)` 等）；`from Germany/based in Vietnam/客户来自...` |
| `companyName` | `XX有限公司/XX有限责任公司/XX集团/XX GmbH/XX Co., Ltd/XX LLC/XX Inc/XX Ltd/XX Trading/XX Industrial` 等后缀；「公司名/公司是/他们公司」+ 名称；`from/at + Company` |
| `foundedAt` | 「成立于 1998 年/成立时间 2010/创立于/有 20 年历史」；`founded in 1998/established in 2010/since 1994`；输出年份或 `YYYY-MM` |
| `teamSize` | 「团队 50 人/员工 100 人/规模 50-100 人/20 名员工」；`team of 50/50 employees/50-100 staff/workforce of 30`；保留量级与单位，输出如 `50-100 人` |
| `annualRevenue` | 「年营收/年营业额/销售额 1200 万欧元/营收约 5 亿美元」；`annual revenue/turnover/sales of EUR 12 million/USD 5M`；输出如 `1200 万欧元`、`500 万美元` |
| `followupStatus` | 已成交/下单/签约/PO → 已成交；报价/quotation → 已报价；样品/sample → 寄样中；谈判/contract terms/negotiat → 谈判中；回复/replied/email sent → 已回复；流失/放弃/lost → 已流失；暂停/on hold → 暂停跟进；否则 `新询盘` |
| `dealStatus` | 明确成交词才给 `已成交`；明确 `未成交/尚未成交/not closed yet/no deal/未下单` 给 `未成交`；两者都没有则**不输出**该字段（界面按默认 未成交） |
| `coreReason` | 「原因/因为/由于/主要是/卡在/顾虑/担心」后的句子；`because/due to/main reason/concern/issue/price is ... above`；优先含 价格/交期/付款/质量/认证/运费/competitor 等关键词的分句；≤ 80 字 |
| `followups` | 文出现的所有「日期 + 事件」组合，按时间倒序取 3 条；note 取该日期附近 10~60 字事件描述（去掉日期本身）；只有 1 个日期时也可作为 1 条跟进 |

通用要求：

- 全角/半角标点、多余空格、换行、`\r\n`、制表符都要健壮；中英文混排文本要正常。
- 不要因为「德国客户」把 `country` 识别成「客户」之类；所有正则加边界，中文按子串、英文按 `\b`。
- 识别不到就**不输出**该 key（不要填空串或「未知」），并计入 `missing`。
- `evidence.snippet` 从原文截取，帮助用户核对；不要编造。
- 不允许抛异常：空文本返回 `{ ok:false, fields:{}, followups:[], warnings:[{code:'empty', message:'请粘贴一段询盘文本'}] }`。
- 性能：10 万字以内文本 < 200ms（正则避免灾难性回溯）。

### 5.3 必须通过的样例（写进 `tests/test-parser.mjs`）

**样例 A（中文）**

```
2026年9月12日收到德国客户 Michael Braun 的询盘。他所在的 Braun Trading GmbH 成立于1998年，主营户外用品，团队规模约50-100人，年营收约1200万欧元。客户咨询的是我们2026款不锈钢保温杯 500ml，要求加印LOGO。目前已报价，客户反馈价格略高于目标价8%，仍在等终端客户确认。11月10日电话沟通确认2000pcs试单意向，11月18日邮件确认样品尺寸。
```

断言：`firstInquiryAt=2026-09-12`、`customerName≈Michael Braun`、`companyName≈Braun Trading GmbH`、`country=德国`、`foundedAt≈1998`、`teamSize≈50-100人`、`annualRevenue≈1200万欧元`、`product` 含 `保温杯`、`followupStatus=已报价`、`dealStatus=未成交`、`coreReason` 含 `价格`、`followups.length>=2` 且首条日期 `2026-11-18`。

**样例 B（英文）**

```
Michael Braun from Braun Trading GmbH (Germany) reached out on Sep 12, 2026 regarding our 500ml stainless steel vacuum flask with custom logo printing. The company was founded in 1998, with a team of 50-100 employees and annual revenue around EUR 12 million. We sent our quotation on Nov 10; the client said our price is about 8% above their target. The deal was closed on Nov 20.
```

断言：`country=德国`、`firstInquiryAt=2026-09-12`、`foundedAt=1998`、`teamSize` 含 `50-100`、`annualRevenue` 含 `12`、`followupStatus=已成交`、`dealStatus=已成交`、`followups` 含 `2026-11-20`。

**样例 C（展会名片速记，字段少）**

```
广交会名片：Mr. Ahmed Hassan, Cairo, Egypt。公司 El Nile Import & Export Co., Ltd，做家居用品进口，年采购额约 300 万美元，团队 15 人，2011 年成立。对我们的陶瓷餐具套装感兴趣，已发样册，等回复。
```

断言：`country=埃及`、`customerName` 含 `Ahmed`、`companyName` 含 `El Nile`、`foundedAt=2011`、`teamSize` 含 `15`、`annualRevenue` 含 `300`、`product` 含 `陶瓷餐具`、`followupStatus` ∈ {已回复, 寄样中}。

另需覆盖：空文本、仅公司无联系人、全角标点、`2026/9/12` 与 `12 Sep 2026` 日期、`not closed yet`、超长文本不超时。

## 6. `js/xlsx.js` 契约（export-dev）

### 6.1 接口

```js
window.CrmXlsx = {
  VERSION: '1.0.0',
  buildWorkbook(sheets) -> Uint8Array,
  // sheets: [{
  //   name: '客户总表',
  //   columns: [{ header: '客户名称', width: 18, numFmt?: '0', align?: 'left'|'center'|'right', wrap?: bool }],
  //   rows: [[cell, ...], ...],       // 字符串或数字；null/undefined → 空单元格
  //   freeze: { rows: 1 },            // 冻结表头
  //   autoFilter: true,
  //   headerStyle?: 'default'
  // }]
  exportCustomers(records, opts) -> { filename, bytes },   // 组装「客户总表 + 跟进明细」两个工作表
  download(bytes, filename) -> void                        // 触发浏览器下载
};
```

`exportCustomers` 的列（顺序固定）：

**Sheet 1 `客户总表`**：序号 | 首次询盘时间 | 客户名称 | 咨询产品 | 国家 | 公司名 | 成立时间 | 团队规模 | 年营收 | 跟进状态 | 是否成交 | 核心原因 | 距上次跟进(天) | 最近跟进①日期 | 最近跟进①内容 | 最近跟进②日期 | 最近跟进②内容 | 最近跟进③日期 | 最近跟进③内容

**Sheet 2 `跟进明细`**：序号 | 首次询盘时间 | 客户名称 | 公司名 | 国家 | 跟进日期 | 距今天数 | 跟进内容 | 跟进状态 | 是否成交（每个客户每条跟进一行，倒序）。

`opts`：`{ filename?: string, today: 'YYYY-MM-DD', title?: string }`。默认文件名 `客户跟进台账_YYYYMMDD_HHmm.xlsx`（本地时间）。

### 6.2 XLSX 实现要求

- 自写 ZIP（store 方式即可，不压缩也能被 Excel 打开）+ CRC32；必须包含 `[Content_Types].xml`、`_rels/.rels`、`xl/workbook.xml`、`xl/_rels/workbook.xml.rels`、`xl/worksheets/sheetN.xml`、`xl/styles.xml`；字符串用 `inlineStr`（`<c t="inlineStr"><is><t xml:space="preserve">…</t></is></c>`）即可，不必 sharedStrings。
- XML 转义 `& < > " '`；剔除 XML 非法控制字符；空单元格不写；数字单元格不写 `t`；日期一律**文本** `YYYY-MM-DD`（避免 Excel 日期序列/本地化坑）。
- 表头：加粗白字、深蓝底 `#0B2239`、居中；表头行高 24；首行冻结 + 自动筛选（`autoFilter` 覆盖整个数据区）；列宽按 `columns[].width`，`核心原因`/跟进内容 `wrapText`。
- 中文文件名与 UTF-8 内容必须正确；不得出现乱码或「文件已损坏」。
- 大数据量：500 行 × 19 列 < 300ms。
- 浏览器与 Node 双端可用：Node 下 `module.exports = api`（用 `TextEncoder` / `Uint8Array`，不要用 `document`/`Blob` 在 Node 分支里）。

### 6.3 验收（`tests/test-xlsx.mjs`）

1. 生成 8 条示例客户的数据 → 写出 `%TEMP%/crm-export-test.xlsx`。
2. 用 Python openpyxl 校验（`load_worksheet` 读取）：工作表名、表头 19 列文本、行数、中文值、冻结窗格、列宽、`autoFilter` 引用、`zipfile.testzip()` 返回 `None`。
3. 断言 `bytes.length > 3000` 且以 `PK\x03\x04` 开头。
4. 至少 3 个单测：空记录、超长中文、特殊字符 `<&"'>` 与换行。

> Python 路径用 `load_workspace_dependencies` 工具获取（`python.exe` + site-packages 里的 openpyxl）。测试脚本用 `node tests/test-xlsx.mjs` 运行，内部用 `child_process.execFileSync(python, [script])` 调 Python 校验。

## 7. UI 契约（ui-dev）

### 7.1 视觉基调

- 与 `../trade-docs` 同款家族视觉：深海军蓝侧栏（#0b2239 → #0c2743 渐变 + 青绿光晕）、青绿主色 `#0f766e / #0d9488`、琥珀强调 `#f59e0b`、浅灰画布 `#f4f6f8`、卡片白底 14px 圆角 + 轻投影。
- **复用 `trade-docs/css/app.css` 的 design tokens 命名**（`--bg --surface --ink --muted --line --brand --brand-600 --brand-700 --navy --accent --danger --ok --radius --shadow --sidebar-w --font-ui --font-num` 等），但样式要为本应用重写，不要把单证 CSS 整体搬过来。
- 信息密度高、专业 B2B 后台感；表格为主角；数字用 `--font-num`。
- 中文界面。无外部字体/图标；图标以内联 SVG（`stroke="currentColor"`）。

### 7.2 必须覆盖的 class（app.js 会实际输出，全部都要有可见效果）

**外壳**：`.app-shell .sidebar .brand .brand-mark .brand-text .nav .nav-section .nav-item .is-active .sidebar-foot .sidebar-note .sidebar-version .app-frame .topbar .topbar-title .crumbs .topbar-heading .topbar-actions .icon-btn .menu-toggle .save-state .save-state-dot .btn .btn--primary .btn--ghost .btn--danger .btn--sm .btn-label .sidebar-backdrop .view .view-inner .boot-state .boot-spinner .toast-stack .toast .toast--ok .toast--error .toast--info .toast-icon .toast-text .toast-close .modal-backdrop .modal .modal--lg .modal--sm .modal-head .modal-title .modal-sub .modal-close .modal-body .modal-foot .sr-only .muted .mono .no-wrap .nowrap`

**页面**：`.page .page-head .page-head-main .page-title .page-sub .page-head-actions .page-grid .page-grid--2 .section-title .hint`

**仪表盘**：`.stats-grid .stat-card .stat-card--ok .stat-card--warn .stat-card--danger .stat-icon .stat-body .stat-label .stat-value .stat-hint .card .card-head .card-title .card-sub .card-actions .card-body .bar-list .bar-row .bar-meta .bar-track .bar-fill .bar-fill--ok .bar-fill--warn .bar-fill--danger .timeline .timeline-item .timeline-dot .timeline-body .timeline-date .timeline-text`

**客户台账**：`.toolbar .toolbar-left .toolbar-right .search-box .search-input .filter-select .select--sm .count-pill .table-wrap .data-table .data-table--customers .row .row--expanded .row-detail .row-toggle .col-expand .cell-strong .cell-muted .cell-num .cell-reason .cell-country .flag .badge .badge--ok .badge--warn .badge--danger .badge--info .badge--brand .badge--neutral .age .age--ok .age--warn .age--danger .age--none .followup-strip .followup-chip .followup-date .followup-note .row-actions .icon-btn--sm .icon-btn--danger .sort-btn .is-active .empty .empty-art .empty-title .empty-text .empty-actions .table-footer .table-count .data-table th[data-col] .data-table td[data-col]`

**表单**：`.form-grid .field .field--full .field-label .req .field-hint .field-error .input .select .textarea .textarea--lg .form-section .form-section-title .followup-editor .followup-row .followup-empty .chip-row .chip .chip.is-active .input-group`

**智能识别**：`.recog-layout .recog-panel .recog-panel--source .recog-panel--result .panel-head .panel-title .panel-hint .recog-summary .recog-count .recog-list .recog-item .recog-item--filled .recog-item--missing .recog-key .recog-value .recog-evidence .recog-actions .sample-pill`

**设置/帮助**：`.settings-list .setting-row .setting-info .setting-title .setting-desc .setting-actions .danger-zone .data-actions .file-btn .help-grid .help-card .help-list .kbd .help-note .tag`

**状态修饰**：`body.is-modal-open`（锁滚动）、`.sidebar.is-open`、`.sidebar-backdrop.is-open`、`.view-inner` 最大宽 `--content-max` 居中、表格横向滚动条外观。

### 7.3 响应式与打印

- `>= 1280px`：完整 13 列 + 操作列。
- `1024 ~ 1279px`：隐藏 `[data-col="foundedAt"]`、`[data-col="teamSize"]`、`[data-col="annualRevenue"]`。
- `< 1024px`：侧栏默认隐藏（`transform: translateX(-100%)`，`.is-open` 滑入 + 背景遮罩）。
- `< 768px`：隐藏 `[data-col="product"]`、`[data-col="companyName"]`、`[data-col="coreReason"]`；`.toolbar` 换行；表格横向滚动，禁止整页横向溢出（`body` 不能出现 x 轴滚动条）。
- `@media print`：隐藏侧栏/顶栏/工具栏/操作列/Toast/Modal，表格全宽、表头重复（`thead { display: table-header-group }`），打印字号 11pt、去阴影。
- 所有断点都要保证：`document.documentElement.scrollWidth <= innerWidth + 1`（qa-dev 会断言）。

### 7.4 截图验收工况（ui-dev 自查）

用 `tests/cdp.mjs`（Lead 会先放好）在 1440×900 与 390×844 下截图：
`#/dashboard`、`#/customers`（含展开一行）、`#/recognize`（识别结果态）、新建客户 Modal、`#/settings`。存 `screenshots/`（qa-dev 会用同名覆盖，ui-dev 先放草稿到 `screenshots/ui/`）。

## 8. `js/app.js` 公共测试接口（qa-dev 依赖）

```js
window.CustomerCRM = {
  version: '1.0.0',
  app: {
    navigate(hash),                  // '#/customers' | 'customers' | '#/customer/xxx' 均可
    getRoute() -> 'dashboard'|'customers'|'recognize'|'settings'|'help',
    refresh(),
    newCustomer(),                   // 打开新建 Modal
    editCustomer(id),                // 打开编辑 Modal
    saveCustomerForm(),              // 用当前 Modal 表单值保存（新增或更新），返回保存后的记录或 null
    deleteCustomer(id),
    addFollowup(id, { date, note }), // 追加跟进（自动截断到 3 条、自动重排）
    recognize(text),                 // = CrmParser.recognize，便于测试直接调用
    applyRecognition(result, mode),  // 'fill'（打开预填表单）| 'add'（直接新增一行）
    exportExcel(),                   // 触发下载，返回 { filename, bytes: number }
    stats(),                         // { total, deal, open, overdue, avgDaysSince }
    toast(message, type)
  },
  store,                             // window.CrmStore
  schema: window.CRM_SCHEMA,
  parser: window.CrmParser,
  xlsx: window.CrmXlsx
};
```

**约定**：

- 表单控件 `name` 属性 = 字段 key（qa 用 `page.fill('customerName', 'X')` 直接填充）；表单 `<form id="customer-form" data-testid="form-customer">`。
- 稳定 `data-testid`：
  - `new-customer`、`btn-export-excel`、`btn-recognize`、`btn-seed-demo`、`btn-clear-all`
  - `customer-row`（每个 `<tr.row>`）、`row-toggle`、`btn-edit-customer`、`btn-delete-customer`、`btn-add-followup`
  - `customer-form`、`btn-save-customer`、`btn-cancel-customer`、`form-error`
  - `customer-search`、`filter-country`、`filter-status`、`filter-deal`、`filter-reset`
  - `recognize-input`、`btn-run-recognize`、`recognize-result`、`btn-fill-customer`、`btn-add-recognized`
  - `stat-total`、`stat-deal`、`stat-open`、`stat-overdue`
  - `toast-stack`、`modal-root`
- 路由 `hashchange` 驱动；刷新后路由保持。
- 所有写操作后 `#save-state` 显示「已保存」（`data-state="saved"`）。
- 空表时有 `.empty` 引导（含「新建客户」「智能识别」「载入示例数据」）。

## 9. 测试与验收

- 单元测试：`node tests/test-schema.mjs`（Lead）、`node tests/test-parser.mjs`（parser-dev）、`node tests/test-xlsx.mjs`（export-dev），全部必须 PASS 且退出码 0。
- E2E：`node tests/e2e.mjs`（qa-dev）用 `tests/cdp.mjs` 启动本地静态服务 + headless Chrome，覆盖：
  1. 首屏加载无 console error / pageerror；
  2. 仪表盘 4 个统计卡数值 = store 计算值；
  3. 新建客户（表单校验：名称/日期必填）→ 表格出现 → 刷新仍在；
  4. 编辑、删除、添加跟进（最多 3 条、最新在前、距上次跟进数字正确）；
  5. 搜索/国家/状态/成交筛选 + 排序；
  6. 智能识别：粘贴 `CrmParser.SAMPLE_TEXTS[0].text` → 结果面板 ≥ 8 项 → 填入表单 → 保存 → 表格字段正确；
  7. 一键导出：拦截下载，文件存在且 `PK` 头、openpyxl 可读、含两个工作表、行数正确；
  8. 响应式：1440/1024/768/390 四档无横向溢出、控制台无错；
  9. 截图：`screenshots/dashboard.png` 等 6 张。
- 验收报告：`tests/QA_REPORT.md`（用例表 + 结果 + 已知问题 + 截图索引）。

## 10. 协作规则

1. 开工前 `team_task_list` 找到自己的任务，`claim` 后再动手；完成后 `complete`。
2. 每 30~45 分钟用 `send_message` 给 `lead` 一条进展（文件、测试输出、阻塞点）。
3. 接口有疑问先按本契约实现，把疑问和你的取舍写进消息；不要擅自改契约。
4. 完成后消息格式：`完成 <task id>：<文件清单>；测试：<命令与结果>；截图/产物：<路径>；备注：<偏差>`。
5. 不启动常驻服务器占端口（测试用的随机端口除外）；不执行 `git` 命令（Lead 统一处理）；不改 `trade-docs/`。
