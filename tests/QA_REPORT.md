# 客户通 CustomerCRM — 端到端验收报告（QA Report）

| 项目 | 内容 |
| --- | --- |
| 验收对象 | 客户通 CustomerCRM **v1.2.0**（静态单页应用，`customer-crm/`） |
| 验收依据 | `docs/CONTRACT.md` **v1.2** §8 公共测试接口 + §9 E2E 要求；task-12 验收清单 |
| 验收人 | qa-dev |
| 验收日期 | 2026-11-22（Asia/Shanghai） |
| 自动化脚本 | `tests/e2e.mjs`（零 npm 依赖，`tests/cdp.mjs` + headless Chrome 157） |
| 最终结果 | **87 / 87 用例 PASS，退出码 0**（2026-11-22 03:01:23–03:01:46 最终验收运行，耗时 23s）；**无遗留产品缺陷（P0/P1/P2 = 0）** |
| 发布结论 | **可以发布（PASS）** |

---

## 1. 概要

### 1.1 v1.2 变更内容（相对 v1.1）

| # | 变更 | E2E 覆盖 |
| --- | --- | --- |
| 1 | **字段 13 → 16**：新增「邮箱 email」「电话 phone」「地址 address」（schema v1.2.0，`isValidEmail`） | A3、C1/C4/C5/C6、E1/E2/E5、F3、H1、I2 |
| 2 | **台账客户名称单元格双行**：`.cell-main`（姓名）+ `.cell-contact`（邮箱 · 电话） | B1、C4、H1 |
| 3 | **新增 `[data-col="address"]` 地址列**：DOM 共 16 列；`<=1599px` 隐藏（改由展开行 meta 展示），`>=1600px` 显示，表格 `min-width 1313` | B1、H1、H1c、H2/H3/H4、H6 |
| 4 | **展开行联系方式**：`.row-detail-meta` 含邮箱/电话/地址，并渲染 `mailto:` / `tel:` 链接 `.contact-link` | B12、C6、H1、H6 |
| 5 | **表单**：邮箱（`type=email`）/电话（`type=tel`）并排、地址整行；非法邮箱保存被拦（`[data-error-for="email"]` + `#form-error`） | C1、C2、C3b、C4、C6 |
| 6 | **搜索**支持邮箱/电话/地址（matchRecord haystack 扩展） | B18 |
| 7 | **Excel**：客户总表 **24 列**（D=邮箱 / E=电话 / F=地址，autoFilter A1:X）、跟进明细 **13 列**（autoFilter A1:M） | F2、F3、F4 |
| 8 | **智能识别 16 字段**：新增邮箱/电话/地址抽取，示例 A/B/C 均带联系方式（样例 A filled=16/16） | E1、E2、E3、E5 |
| 9 | 表格 16 列（expand + 14 + actions）；1440 内容宽 1152、`min-width 1313`；`>=1600` 全列 `min-width 1722` | B1、H1、H1b、H1c |

### 1.2 最终验收运行（发布门禁）

| 项 | 值 |
| --- | --- |
| 运行命令 | `node tests/e2e.mjs`（独占运行，期间无其他人跑 e2e/写 `screenshots/` 根目录） |
| 运行时间 | 2026-11-22 03:01:23 – 03:01:46（Asia/Shanghai，耗时 23s） |
| 结果 | **87 / 87 PASS，FAIL 0，SKIP 0，退出码 0** |
| 分组 | A 10/10、B 18/18、C 12/12、P 6/6、D 4/4、E 7/7、F 5/5、G 10/10、H 11/11、I 2/2、S 2/2 |
| 截图 | 8 张全部由本次运行重新生成（03:01:30–03:01:44），尺寸/内容校验通过 |
| 被测版本 | 前端文件指纹见附录 B；`css/app.css` 为 ui-dev 冻结版（90123B / `9D7132A103C0CD83`，2026-11-22 02:52:51） |
| 契约 | `docs/CONTRACT.md` **v1.2**（含 §7.3 地址列断点口径：`>=1600` 显示、`<=1599` 折叠到行展开 meta） |
| 辅助测试 | schema 122/122、parser 87/87、xlsx 165 断言、smoke 35/35、v11-check 全通过（均 exit 0） |

### 1.3 版本演进对照（v1.0 → v1.1 → v1.2）

| 版本 | 最终运行 | 用例 | 字段/表格/Excel | 关键能力 |
| --- | --- | --- | --- | --- |
| v1.0 | 2026-11-22 00:23:41–00:24:00（19s） | **72/72 PASS** | 11 字段 / 表格 14 列 / Excel 19+10 列 | 台账、CRUD、跟进、识别 11 字段、导出、响应式/打印、损坏兜底 |
| v1.1 | 2026-11-22 01:18:20–01:18:53（33s） | **84/84 PASS** | 13 字段 / 表格 15 列 / Excel 21+11 列 | 新增产品型号/主营业务、产品库与联动下拉、扇形图、宽版台账 |
| v1.2 | 2026-11-22 03:01:23–03:01:46（23s） | **87/87 PASS** | **16 字段 / 表格 16 列 / Excel 24+13 列** | 新增邮箱/电话/地址、名称双行联系信息、地址列断点、mailto/tel 链接、邮箱校验、联系方式搜索 |

> 三个版本的用例全部保留并逐版更新断言（旧断言按新字段/列数升级）；本报告只记录 v1.2 的用例明细，v1.0/v1.1 结论作为演进存档。

### 1.4 v1.0 / v1.1 历史结论（存档，已被 v1.2 取代）

- v1.0：72/72 PASS，8 图齐备，结论「可以发布」；验收期间发现 QA-001（P1，btn-add-followup 缺失）、QA-002（P2，row-toggle/toast-stack/modal-root testid 缺失），均由 Lead 修复并回归通过。
- v1.1：84/84 PASS，8 图齐备，结论「可以发布」；v1.1 验收未发现新增产品缺陷。
- v1.2 保留上述全部回归用例，并在 16 字段/16 列/24+13 列新契约下重新验证。

---

## 2. 验收环境与工具链

| 项 | 值 |
| --- | --- |
| OS | Windows 10 Pro 22H2（`10.0.19045`） |
| 浏览器 | Google Chrome `157.0.8103.79`（headless=new，CDP） |
| Node.js | `v24.17.0`（无 npm，零第三方依赖） |
| Python / openpyxl | Python `3.12.14` + openpyxl `3.1.5`（XLSX 读回校验） |
| 运行命令 | `cd customer-crm && node tests/e2e.mjs` |
| 静态服务 | `tests/cdp.mjs` 在 `127.0.0.1` 随机端口起静态服务（不占固定端口） |
| 数据 | 浏览器 `localStorage['customer-crm.v1']`：12 条示例客户（含联系方式）+ 12 条预置产品，每轮清空重建 |
| 截图 | `screenshots/{dashboard,customers,customers-expanded,recognize,form,settings,mobile-dashboard,mobile-customers}.png` |

辅助测试（验收时复跑，作为发布参考；非本任务交付物）：

| 测试 | 结果 |
| --- | --- |
| `node tests/test-schema.mjs` | 122/122 PASS，exit 0 |
| `node tests/test-parser.mjs` | 87/87 PASS，exit 0（样例 A filled=16/16） |
| `node tests/test-xlsx.mjs` | 165 项断言 PASS（openpyxl 读回 10 次），exit 0 |
| `node tests/smoke.mjs` | 35/35 PASS，exit 0 |
| `node tests/v11-check.mjs` | v1.1 自检全部通过，exit 0 |

---

## 3. 验收方法

1. **真实交互**：`PageSession.fill()` 用原生 value setter + `input/change/blur` 事件填表；下拉联动通过真实 `change` 事件驱动；点击全部走真实 DOM `click()`；键盘用例走 CDP `Input.dispatchKeyEvent`（`/`、`N`、`Esc`）。
2. **DOM ↔ store 交叉断言**：表格行数/列值/顺序、名称双行（`.cell-main/.cell-contact`）、统计卡数字、age 色调分布、行展开 meta（含邮箱/电话/地址）与跟进内容、产品库使用次数、扇形图切片/图例/百分比，均与 `CrmStore` / `app.filtered()` / `app.stats()` 实时比对。
3. **v1.2 新增断言**：
   - 联系方式：名称副行 `.cell-contact` 含邮箱/电话；展开行 `.contact-link` 的 `mailto:` / `tel:` href 正确（tel 保留 `+` 区号）；
   - 地址列断点：1440（`<=1599`）`td[data-col=address]` display:none 且仍存在 16 个 td；1680（`>=1600`）地址列显示、表格 min-width 1722；
   - 邮箱校验：填 `not-an-email` 保存被拦（字段级错误 + `#form-error` + store 不新增），改回合法邮箱后可保存；
   - 联系方式搜索：邮箱片段、电话片段、地址片段各搜一次，DOM 顺序 = `app.filtered()`，命中行文本包含关键词；
   - 编辑清空地址：store.address 为空串，展开行地址显示「—」；
   - Excel：openpyxl 读回 24/13 列表头、D/E/F 列值 = 导出首条记录的邮箱/电话/地址、autoFilter 范围 A1:X / A1:M；
   - 识别：16 项、filled=16/16，email=m.braun@braun-trading.de、phone 含 +49、address 含 Stuttgart，填入表单后三字段预填。
4. **真实下载**：每次导出前清空临时下载目录并 `Page.setDownloadBehavior`，轮询等待文件落盘；先做 Node ZIP 结构断言，再由 openpyxl 做权威读回。
5. **打印媒体**：`Emulation.setEmulatedMedia({media:'print'})` 后读 `getComputedStyle`：台账恢复全部业务列（含地址）、侧栏/工具栏/操作列隐藏、`thead=table-header-group`、`.contact-link` 可读；仪表盘扇形图保留可见。
6. **截图是断言的一部分**：每张截图校验 PNG 尺寸（1440×900 / 390×844）与体积下限；`form.png` 断言截图前邮箱/电话/地址/产品名称/型号已填入。

---

## 4. 用例与结果

> 结果均来自最终验收运行 `node tests/e2e.mjs`。证据列中的编号为脚本运行日志中的用例编号。

### A. 首屏启动与仪表盘（10/10 PASS）

| 编号 | 场景 | 步骤 | 期望 | 结果 | 证据 |
| --- | --- | --- | --- | --- | --- |
| A1 | 首屏启动 | 打开 `index.html`，等待 app | hash=`#/dashboard`、导航高亮、标题含「仪表盘」、统计卡渲染 | PASS | 运行日志 A1 |
| A2 | 首屏无错误 | 监听 console/pageerror/log.error | 0 条错误级输出 | PASS | 运行日志 A2 |
| A3 | 示例数据 | 首次访问未预置 localStorage | 版本 1.2.0、16 字段、12 客户、12 产品；每条示例含邮箱/电话/地址；Michael 邮箱/电话/地址正确 | PASS | 运行日志 A3 |
| A4 | 统计卡一致性 | 读 4 卡与 `app.stats()` | total/deal/open/overdue 完全相等，total=12 | PASS | 运行日志 A4 |
| A5 | 仪表盘区块 | 统计卡/状态分布/动态/优先跟进 | 4 卡、状态条 ≥4、动态 8 条、优先条数 = `stats.overdueList.length` | PASS | 运行日志 A5 |
| A6 | 保存指示器 | 读 `#save-state` | `data-state="saved"` 且文本非空 | PASS | 运行日志 A6 |
| A7 | 公共接口自检 | 校验 CONTRACT v1.2 §8 方法与常量 | 版本 1.2.0、schema/parser 1.2.0、`isValidEmail`、app/store/产品库方法齐全 | PASS | 运行日志 A7 |
| A8 | 1440 无横向溢出 | doc/body scrollWidth 与 innerWidth | ≤ innerWidth + 1 | PASS | 运行日志 A8 |
| A9 | 仪表盘截图 | `Page.captureScreenshot` | PNG 1440×900、≥20KB | PASS | `screenshots/dashboard.png` |
| A10 | 国家分布扇形图 | 读 `[data-testid=country-pie]` | SVG 存在、7 切片 = 7 图例、百分比合计 100、中心数=12、切片有 path 与 `<title>`、含「其他 6 个国家」 | PASS | 运行日志 A10；dashboard.png |

### B. 客户台账（18/18 PASS）

| 编号 | 场景 | 步骤 | 期望 | 结果 | 证据 |
| --- | --- | --- | --- | --- | --- |
| B1 | 表格结构 | 打开 `#/customers` | 12 行、每行 16 td、表头 14 个业务列（含地址）、`body.route-customers`、名称双行 `.cell-main/.cell-contact`、产品 `.cell-sub`、testid 齐全 | PASS | 运行日志 B1 |
| B2 | 关键词搜索 | 搜索「德国」 | 1 行，DOM 顺序=`app.filtered()`，显示已筛选标记 | PASS | 运行日志 B2 |
| B3 | 搜索空结果 | 输入不存在关键词 | `.empty`「没有匹配的客户」；重置后 12 行 | PASS | 运行日志 B3 |
| B4 | 国家筛选 | 选「德国」 | 1 行，与 store 计算一致 | PASS | 运行日志 B4 |
| B5 | 状态筛选 | 选「已报价」 | 3 行且状态列均为已报价 | PASS | 运行日志 B5 |
| B6 | 成交筛选 | 选「已成交」 | 2 行且成交列均为已成交 | PASS | 运行日志 B6 |
| B7 | 组合筛选 | 已报价 + 德国 | 1 行（Michael Braun，名称取 `.cell-main`） | PASS | 运行日志 B7 |
| B8 | 重置筛选 | 点「重置」 | filters 全空、下拉归位、12 行 | PASS | 运行日志 B8 |
| B9 | 排序：首次询盘 | 连点列头 3 次 | desc→asc→默认；DOM 顺序=app.filtered()，日期单调 | PASS | 运行日志 B9 |
| B10 | 排序：客户名称 | 连点 2 次 | desc→asc，DOM 顺序=app.filtered()（按 `.cell-main` 比较） | PASS | 运行日志 B10 |
| B11 | 排序：距上次跟进 | 点 1 次 | daysSince-desc，DOM 顺序=派生排序 | PASS | 运行日志 B11 |
| B12 | 展开行 | 点 row-toggle | 跟进 chips=min(3,n)、首条「最近」；`.row-detail-meta` 含邮箱/电话/地址/产品型号/主营业务/成立时间/团队规模/年营收/首次询盘/最近更新及值；含 `mailto:`/`tel:` 链接 | PASS | 运行日志 B12；customers-expanded.png |
| B13 | 收起行 | 再点 row-toggle | 详情 hidden=true、无 `row--expanded` | PASS | 运行日志 B13 |
| B14 | 台账截图 | 收起态/展开态各 1 张 | PNG 1440×900 | PASS | `customers.png`、`customers-expanded.png` |
| B16 | 搜索主营业务 | 搜索「户外用品」 | 命中 2 行、DOM=app.filtered()、主营列含关键词 | PASS | 运行日志 B16 |
| B17 | 搜索产品型号 | 搜索「500ml」 | 命中 2 行、产品格含 500ml、`.cell-sub` 显示型号 | PASS | 运行日志 B17 |
| B18 | 搜索联系方式 | 邮箱片段 `braun-trading.de` / 电话片段 `+49 711` / 地址片段 `Stuttgart` 各搜一次 | 每次命中行数正确、DOM=app.filtered()、命中行文本包含关键词（名称副行/地址列） | PASS | 运行日志 B18 |
| B19 | 无控制台错误 | 全程监听 | 0 条 | PASS | 运行日志 B19 |

### C. 客户增删改与跟进（12/12 PASS）

| 编号 | 场景 | 步骤 | 期望 | 结果 | 证据 |
| --- | --- | --- | --- | --- | --- |
| C1 | 新建弹窗结构 | 点 new-customer | 16 字段均可映射、email type=email/tel type=tel 并排、地址 `field--full`、产品选择器 3 testid、2 必填、role=dialog | PASS | 运行日志 C1 |
| C2 | 表单截图 | 填客户/邮箱/电话/地址/主营 + 选产品名称/型号后截图 | 截图前 7 项已填；PNG 1440×900 | PASS | `form.png` |
| C3 | 必填校验 | 清空客户名称/首次询盘时间→保存 | `#form-error` 含「首次询盘时间、客户名称」，字段级错误，弹窗不关闭 | PASS | 运行日志 C3 |
| C3b | 邮箱格式校验 | 填 `not-an-email` → 保存；再改合法邮箱 | `[data-error-for=email]` 与 `#form-error` 提示邮箱格式；store 不新增；弹窗不关闭；改合法后可继续 | PASS | 运行日志 C3b |
| C4 | 保存新客户 | 16 字段 + 产品联动 → 保存 | store 13 条；email/phone/address/product/model/mainBusiness 正确；名称格 `.cell-main`=姓名、`.cell-contact` 含邮箱与电话；产品 `.cell-sub`=500ml；`#save-state=saved` | PASS | 运行日志 C4 |
| C5 | 刷新持久化 | reload | localStorage 13 条（16 字段含联系方式）+ 12 产品；表格仍显示 | PASS | 运行日志 C5 |
| C6 | 编辑客户 | 行内编辑：联系方式回填 → 改电话、清空地址 → 保存 → 展开行 | 邮箱/电话/地址预填正确；保存后 store 电话更新、address=''；展开行显示新电话、地址「—」、mailto/tel 链接存在 | PASS | 运行日志 C6 |
| C7 | 添加跟进 1~3 | 跟进弹窗添加 3 条不同日期 | 条数递增、最新在前、日期倒序 | PASS | 运行日志 C7 |
| C8 | 第 4 条跟进 | 再添加 1 条（今天） | 只保留 3 条、最新在前、最旧被截断、「距上次跟进=今天」且 `age--ok` | PASS | 运行日志 C8 |
| C9 | 删除客户 | 取消；再确认删除 | 取消不删（13）；确认后 12 行、toast、行消失、store 无记录 | PASS | 运行日志 C9 |
| C10 | data-testid 契约 | 检查行内按钮 | `btn-add-followup`、`row-toggle` 均存在 | PASS | 运行日志 C10 |
| C12 | 无控制台错误 | 全程监听 | 0 条 | PASS | 运行日志 C12 |

### P. 产品库与产品选择器（6/6 PASS）

| 编号 | 场景 | 步骤 | 期望 | 结果 | 证据 |
| --- | --- | --- | --- | --- | --- |
| P1 | 产品库渲染 | 打开设置页产品库卡 | 12 行、名称/型号/使用次数/删除按钮；「共 12 个产品」；`productUsage()['不锈钢保温杯\|500ml']=1` 且行文案「1 位客户使用」 | PASS | 运行日志 P1；settings.png |
| P2 | 新增产品 | 填「测试产品/X1」→ 加入产品库 | 13 行、`productNames()` 含「测试产品」、`productModels('测试产品')=['X1']`；新建客户下拉包含产品与型号 | PASS | 运行日志 P2 |
| P3 | 删除产品 | 点该行删除按钮 | 回 12 行、下拉同步移除 | PASS | 运行日志 P3 |
| P4 | 恢复预置产品 | 删除预置产品 → 恢复预置 | 先 11 行；恢复后 12 行，产品名与 500ml 恢复 | PASS | 运行日志 P4 |
| P5 | 产品联动与自定义 | 名称→型号联动；`app.selectProduct`；自定义并加入产品库 | 型号 ≥3 个且可选 500ml；API 切 750ml 生效；自定义输入显示、型号下拉隐藏；加入后产品库包含自定义产品 | PASS | 运行日志 P5 |
| P6 | 无控制台错误 | 全程监听 | 0 条 | PASS | 运行日志 P6 |

### D. 键盘快捷键（4/4 PASS）

| 编号 | 场景 | 步骤 | 期望 | 结果 | 证据 |
| --- | --- | --- | --- | --- | --- |
| D1 | `/` 聚焦搜索 | 仪表盘按 CDP 键盘 `/` | 跳转 `#/customers` 且焦点在 `#customer-search` | PASS | 运行日志 D1 |
| D2 | `N` 新建 | 按 `N` | 打开新建客户弹窗（mode=create） | PASS | 运行日志 D2 |
| D3 | `Esc` 关闭 | 按 `Esc` | 弹窗关闭、`body.is-modal-open` 移除 | PASS | 运行日志 D3 |
| D4 | 无控制台错误 | 全程监听 | 0 条 | PASS | 运行日志 D4 |

### E. 智能识别（7/7 PASS）

| 编号 | 场景 | 步骤 | 期望 | 结果 | 证据 |
| --- | --- | --- | --- | --- | --- |
| E1 | 识别示例 A | 点「中文询盘」示例/写入原文 → 开始识别 | 16 项、filled≥15（实测 16/16）、摘要分母 16、证据 ≥15；email=m.braun@braun-trading.de、phone 含 +49、address 含 Stuttgart；产品型号 500ml、主营 户外用品 | PASS | 运行日志 E1 |
| E2 | 字段一致性 | DOM 值 vs `CrmParser.recognize()` | 16 字段逐项相等；parser 样例 A filled=16/16；联系方式与既有字段值正确 | PASS | 运行日志 E2 |
| E3 | 填入表格并核对 | 点「填入表格并核对」 | 新建表单预填；email/phone/address 预填正确；产品名称下拉=不锈钢保温杯、型号=500ml；显示识别说明 | PASS | 运行日志 E3 |
| E4 | 识别页截图 | 结果态截图 | PNG 1440×900 | PASS | `recognize.png` |
| E5 | 直接新增一行 | 点「直接新增一行」 | store +1；新记录 email/phone/address/product/productModel/mainBusiness 与识别结果一致 | PASS | 运行日志 E5 |
| E6 | 空文本兜底 | 清空文本 → 开始识别 | 仅 error toast，不产生结果、不抛异常 | PASS | 运行日志 E6 |
| E7 | 无控制台错误 | 全程监听 | 0 条 | PASS | 运行日志 E7 |

### F. 一键导出 Excel（5/5 PASS）

| 编号 | 场景 | 步骤 | 期望 | 结果 | 证据 |
| --- | --- | --- | --- | --- | --- |
| F1 | 无筛选导出 | 点 `[data-testid=btn-export-excel]`，等待文件落盘 | 文件名 `客户跟进台账_YYYYMMDD_HHmm.xlsx` 且日期为当天；toast 含 12 位客户 | PASS | 运行日志 F1；临时下载目录 |
| F2 | ZIP 结构 | Node 解析 ZIP | PK 头、体积>3000B、7 个必需部件、两个工作表 XML、客户总表 12+1 行、跟进明细 25+1 行 | PASS | 运行日志 F2 |
| F3 | openpyxl 读回 | Python openpyxl 打开导出文件 | 24 列表头逐项一致（D=邮箱/E=电话/F=地址）；首行 D/E/F/G/H/K 列值正确；冻结 A2；autoFilter=A1:X；客户名顺序；明细 13 列表头逐项一致、autoFilter=A1:M | PASS | 运行日志 F3 |
| F4 | 筛选导出 | 筛选「已成交」2 条 → 导出 | 仅导出筛选结果（总表 3 行、明细行数=筛选跟进数+1）；openpyxl 读回客户名=筛选结果、列数 24/13 | PASS | 运行日志 F4 |
| F5 | 无控制台错误 | 全程监听 | 0 条 | PASS | 运行日志 F5 |

### G. 设置与数据管理（10/10 PASS）

| 编号 | 场景 | 步骤 | 期望 | 结果 | 证据 |
| --- | --- | --- | --- | --- | --- |
| G1 | 设置页渲染 | 打开 `#/settings` | 默认阈值 7/14、数据管理/产品库相关按钮齐全 | PASS | 运行日志 G1 |
| G2 | 设置页截图 | 截图 | PNG 1440×900 | PASS | `settings.png` |
| G3 | 阈值保存/校验 | 保存 3/5；再填 30/10 | 3/5 写入 store；非法阈值被拒绝且 store 不变 | PASS | 运行日志 G3 |
| G4 | 仪表盘联动 | 回仪表盘 | 4 卡 DOM=stats()；staleDays 5 时超期数 >14 时；hint 文案含 `≥3`/`≥5` | PASS | 运行日志 G4 |
| G5 | age 颜色分布 | 台账读 `.age--ok/warn/danger/none` | 与 `store.deriveAll()` 当前阈值统计完全一致，共 12 个 | PASS | 运行日志 G5 |
| G6 | JSON 备份导出 | 点「导出备份」 | 文件名规范；合法 JSON；version 1.2.0；12 客户且均含 email/phone/address 字段、含邮箱值；12 产品；含设置 | PASS | 运行日志 G6 |
| G7 | 清空数据 | 点「清空数据」→ 确认 | store 0；空状态引导；仪表盘总数=0；产品库保留 12 条 | PASS | 运行日志 G7 |
| G8 | JSON 恢复 | 选择备份文件（真实 FileReader 链路） | store 恢复 12 客户、产品库 12 条、表格 12 行、成功 toast | PASS | 运行日志 G8 |
| G9 | 载入示例数据 | 清空后点「载入示例数据」 | store 恢复 12 条、表格 12 行、toast 正确 | PASS | 运行日志 G9 |
| G10 | 无控制台错误 | 全程监听 | 0 条 | PASS | 运行日志 G10 |

### H. 响应式、宽版台账、地址断点、抽屉与打印（11/11 PASS）

> 按 CONTRACT v1.2 §7.3：`>=1600px` 显示全部 16 列（含地址），表格 `min-width 1722`；`<=1599px` 折叠「地址/成立时间/团队规模/年营收」（行展开 `.row-detail-meta` 展示），表格 `min-width 1313`；`<768px` 折叠「咨询产品/主营业务/公司名/核心原因/距上次跟进」；`<=560px` 再折叠「首次询盘时间」；打印恢复全部业务列。

| 编号 | 场景 | 步骤 | 期望 | 结果 | 证据 |
| --- | --- | --- | --- | --- | --- |
| H1 | 1440 宽版台账 | 1440 视口打开台账 | 无横向溢出；侧栏可见；`body.route-customers`；`.view-inner`≥1120（实测 1152）；`min-width`≥1313；16 td；10 个业务列可见、地址+3 列折叠；名称副行含邮箱；展开行 meta 含邮箱/电话/地址及全部低优先字段、mailto/tel 链接 | PASS | 运行日志 H1 |
| H1b | 宽版对比 | 读台账 vs 仪表盘 `.view-inner` 宽度 | 台账宽度 ≥1120 且明显大于仪表盘 | PASS | 运行日志 H1b |
| H1c | 1680 全列 | 1680 视口（≥1600）打开台账 | 地址/成立时间/团队规模/年营收恢复显示；`min-width`≥1722；16 td；地址列有文本；无整页横向溢出 | PASS | 运行日志 H1c |
| H2 | 1024×800 | 1024 视口 | 无横向溢出；侧栏收起；地址+3 列折叠；产品/主营业务/公司名/核心原因可见 | PASS | 运行日志 H2 |
| H3 | 768 边界 + 767 | 768 与 767 两档 | 768 无溢出、侧栏收起、产品/主营/距上次跟进可见、地址折叠；767 产品/主营/公司名/核心原因/距上次跟进/地址均隐藏、客户名称可见、无溢出 | PASS | 运行日志 H3 |
| H4 | 390×844 | 390 视口 | 无横向溢出；侧栏收起；菜单按钮可见；移动端折叠列（含地址）隐藏、客户名称可见 | PASS | 运行日志 H4 |
| H5 | 侧栏抽屉 | 点菜单按钮 → 点遮罩 | 打开后侧栏滑入、遮罩可见、aria-expanded=true；点遮罩后收起移出视口 | PASS | 运行日志 H5 |
| H6 | 打印媒体（台账） | 展开一行后 `Emulation.setEmulatedMedia('print')` | 侧栏/顶栏/工具栏/页头操作/操作列/Toast 隐藏；全部业务列恢复（含地址）；`thead=table-header-group`；`.contact-link` 的 mailto/tel 可读 | PASS | 运行日志 H6 |
| H6b | 打印媒体（仪表盘） | 打印媒体下打开仪表盘 | 扇形图容器保留可见、有实际尺寸、7 切片、中心数 12；侧栏隐藏 | PASS | 运行日志 H6b |
| H7 | 移动端截图 | 390×844 仪表盘/台账各 1 张 | PNG 390×844 | PASS | `mobile-dashboard.png`、`mobile-customers.png` |
| H8 | 无控制台错误 | 全程监听 | 0 条 | PASS | 运行日志 H8 |

### I. 数据损坏与旧数据升级兜底（2/2 PASS）

| 编号 | 场景 | 步骤 | 期望 | 结果 | 证据 |
| --- | --- | --- | --- | --- | --- |
| I1 | 非法 JSON 兜底 | 写入 `{bad json` → reload | 应用正常启动不白屏；重建 12 客户 + 12 产品；localStorage 重写为合法 payload；无 pageerror | PASS | 运行日志 I1 |
| I2 | v1.1 旧 payload 升级 | 写入无 products 且客户无联系方式的旧 payload → reload | 旧客户保留；补齐 12 条预置产品；客户记录补全 16 字段结构、email/phone/address 为字符串；首次写操作后持久化 products | PASS | 运行日志 I2 |

### S. 交付物校验（2/2 PASS）

| 编号 | 场景 | 步骤 | 期望 | 结果 | 证据 |
| --- | --- | --- | --- | --- | --- |
| S1 | 8 张截图齐备 | 校验存在性、PNG 合法性、尺寸、体积 | 6 张 1440×900 + 2 张 390×844，均 ≥20KB；尺寸完全匹配 | PASS | 运行日志 S1 |
| S2 | 命名与 README 一致 | 校验文件名规范 | `screenshots/*.png` 命名与 README 引用路径一致 | PASS | 运行日志 S2 |

---

## 5. 缺陷清单

### 5.1 v1.2 缺陷

**无遗留缺陷（P0 / P1 / P2 = 0）。** 本次 v1.2 验收未发现产品功能缺陷；自动化脚本 1 处断言口径问题（客户名称单元格含 `.cell-contact` 副行文本）已在收敛过程中修正，记录于 §5.4。

### 5.2 v1.0 历史缺陷（已修复，保留对照）

| 编号 | 严重级 | 现象 | 复现步骤 | 期望行为 | 归属 | 状态 |
| --- | --- | --- | --- | --- | --- | --- |
| QA-001 | P1 | 客户台账行内「记录跟进」等 3 处按钮缺少 CONTRACT §8 要求的 `data-testid="btn-add-followup"` | 打开 `#/customers` 执行 `document.querySelector('[data-testid="btn-add-followup"]')` → `null` | 三个「记录跟进」入口均带该 testid | Lead（js/app.js） | v1.0 已修复；v1.1/v1.2 回归通过 |
| QA-002 | P2 | `row-toggle` / `toast-stack` / `modal-root` 无 `data-testid` | 执行 `document.querySelector('[data-testid="row-toggle"]')` → `null` | 与契约 §8 一致补 testid | Lead（js/app.js、index.html） | v1.0 已修复；v1.1/v1.2 回归通过 |

### 5.3 已知设计取舍 / 口径说明（非缺陷，已在用例中固化）

| 编号 | 说明 |
| --- | --- |
| N-01 | `stats().open`（待跟进卡）= 「未成交且未流失」客户总数，与 `warnDays` 无关；`warnDays` 影响卡片提示文案与 age 色调，`staleDays` 影响 `overdue`（G4/G5 断言）。 |
| N-02 | 响应式列优先级（契约 v1.2 §7.3）：`>=1600px` 显示全部 16 列（含地址，`min-width 1722`）；`<=1599px` 折叠「地址/成立时间/团队规模/年营收」（行展开展示，`min-width 1313`，1440 内容宽 1152）；`<768px` 折叠「咨询产品/主营业务/公司名/核心原因/距上次跟进」；`<=560px` 再折叠「首次询盘时间」；打印媒体恢复全部业务列。H1–H4/H6 覆盖边界（768px 时产品列仍可见）。 |
| N-03 | **邮箱/电话不单独占用表格列**（与地址不同）：表格 16 列 = 展开 + 14 + 操作，邮箱/电话合并到客户名称单元格副行 `.cell-contact`；Excel 中邮箱/电话/地址各自成列（D/E/F）。 |
| N-04 | 联系方式为可选字段：空值合法；一旦填写邮箱则按 `isValidEmail` 校验格式，非法时阻止保存并定位到邮箱字段（C3b）。 |
| N-05 | 展开行 `.row-detail-meta` 提供完整联系方式与业务背景（邮箱/电话/地址/产品型号/主营业务/成立时间/团队规模/年营收/首次询盘/最近更新），并渲染 `mailto:`/`tel:` 链接；邮箱/电话列在窄屏与打印场景的完整性由此保证。 |
| N-06 | 「产品型号」不单独占用表格列，产品名称格内以 `.cell-sub` 小字展示；Excel 客户总表单独成列（H 列）；行展开 meta 也展示。 |
| N-07 | 产品库「恢复预置产品」为整体重置为 12 条预置数据（会覆盖自建产品）；产品库不随「清空客户数据」清空（G7/P4 断言）。 |
| N-08 | 产品使用次数 = 客户记录中 `product|productModel` 精确匹配的客户数（P1 断言 1 位客户使用）。 |
| N-09 | 旧 payload 迁移：启动时内存补齐 12 条预置产品与 16 字段结构，首次任意写操作后持久化（I2 断言）。 |
| N-10 | 移动端（`<768px`）Toast 固定在底部，避免遮挡页头标题与操作按钮（H4/H5 覆盖）。 |
| N-11 | 打印仪表盘保留国家分布扇形图（独立打印配色/尺寸规则），H6b 断言。 |

### 5.4 测试资产迭代记录（非产品缺陷，留档备查）

| 问题 | 处理 |
| --- | --- |
| v1.0：`tests/e2e.mjs` 曾被 PowerShell 批量改写导致 GBK 双重编码损坏 | 改用 `write`/`edit` 工具以 UTF-8 重写；`node --check` + BOM/U+FFFD 检查 |
| v1.0：openpyxl 校验脚本在 Windows 管道 stdout 默认 cp936 导致乱码 | Python 子进程 `PYTHONIOENCODING=utf-8` + `PYTHONUTF8=1`，脚本 `ensure_ascii=True` |
| v1.0：窄屏断言读到侧栏 `transform` 过渡动画中间态 | `waitUntil` 等待动画结束（`right<=1` / `left>=-1`） |
| v1.0：损坏兜底用例假定 reload 后停在 dashboard | 先等页面渲染，再显式 `goto('#/dashboard')` |
| v1.1：产品库卡副标题断言写成 `indexOf(...)===0` | 改为 `!== -1`；旧 payload 升级持久化断言改为「首次写操作后持久化」 |
| v1.2：`readCol('customerName')` 返回名称单元格的 `cell-contact` 副行文本，导致组合筛选/客户名排序断言偏差 | 读取时优先取 `.cell-main` 文本（B7/B10） |
| v1.2：调试期用 `E2E_SHOTS_DIR` 将截图写临时目录，避免冻结前覆盖正式截图；最终运行不设该变量 | 最终运行写 `screenshots/` 根目录 |

---

## 6. 截图索引

> 全部截图由最终验收运行（2026-11-22 03:01:30–03:01:44）直接生成，均为有效 PNG；v1.0/v1.1 同名截图已被覆盖。

| 文件 | 尺寸 | 字节数 | 内容 | 来源用例 |
| --- | --- | --- | --- | --- |
| `screenshots/dashboard.png` | 1440×900 | 147846 | 仪表盘首屏：4 统计卡、国家分布扇形图（7 切片 + 图例 + 中心 12）、状态分布、最近动态、优先跟进 | A9/A10 |
| `screenshots/customers.png` | 1440×900 | 238411 | 客户台账宽版：12 行 × 16 列（地址列折叠）、客户名称双行（姓名 + 邮箱/电话副行）、产品名称+型号小字、工具栏 | B14 |
| `screenshots/customers-expanded.png` | 1440×900 | 225262 | 台账展开首行：`.row-detail-meta`（邮箱/电话/地址、产品型号/主营业务/成立时间/团队规模/年营收/首次询盘/最近更新，含 mailto/tel 链接）与最近三次跟进 | B14 |
| `screenshots/recognize.png` | 1440×900 | 179811 | 智能识别结果态：中文询盘示例 + 已识别 16/16 项（含邮箱/电话/地址、产品型号 500ml、主营业务 户外用品）与原文证据 | E4 |
| `screenshots/form.png` | 1440×900 | 201725 | 新建客户弹窗：16 字段分组表单，邮箱/电话并排、地址整行、产品名称+型号联动、主营业务已填 | C2 |
| `screenshots/settings.png` | 1440×900 | 121916 | 设置页：数据管理、跟进提醒阈值、产品库、存储信息 | G2 |
| `screenshots/mobile-dashboard.png` | 390×844 | 55956 | 移动端仪表盘（单列统计卡、扇形图） | H7 |
| `screenshots/mobile-customers.png` | 390×844 | 71184 | 移动端客户台账（抽屉收起、名称双行联系方式、折叠列、表格横向滚动容器） | H7 |

---

## 7. 结论

1. 最终验收运行（2026-11-22 03:01:23–03:01:46，独占执行）：`node tests/e2e.mjs` **87/87 用例 PASS，退出码 0**，覆盖 CONTRACT v1.2 §9 全部 E2E 要求 + task-12 的 10 项必测（16 列/名称双行/地址列断点/联系方式链接/邮箱校验/联系方式搜索/16 字段识别/24+13 列导出/宽版与打印/旧数据升级），并完整回归 v1.0/v1.1 原有用例。
2. 8 张验收截图齐备，尺寸/内容校验通过；v1.2 新增的联系方式双行、地址列断点、邮箱/电话并排表单、16 字段识别均已在截图中体现。
3. 缺陷：v1.2 验收**未发现新增产品缺陷（P0/P1/P2 = 0）**；v1.0 的 2 个契约级问题（P1/P2）已修复并在 v1.1/v1.2 回归通过。
4. 辅助测试（schema 122/122、parser 87/87、xlsx 165 断言、smoke 35/35、v11-check）全部 PASS，与 E2E 结果一致。

**发布结论：可以发布（PASS）。**
建议发布前置动作（发布清单）：以本报告「附录 B」记录的文件指纹为准构建 `dist/` 单文件与 GitHub Pages 产物（`css/app.css` 冻结版 `9D7132A103C0CD83`）；如之后任何前端文件再有改动，需重跑 `node tests/e2e.mjs` 并重新生成本报告第 6 节截图、更新附录 B 指纹。

---

## 附录 A：命令与产物

```powershell
cd "C:\Users\Administrator\Desktop\Global Country\customer-crm"
node tests/e2e.mjs          # 87 用例，全绿 exit 0；失败打印编号/实际/期望并 exit 1
# 辅助：node tests/test-schema.mjs / test-parser.mjs / test-xlsx.mjs / smoke.mjs / v11-check.mjs
```

产物：`tests/e2e.mjs`（自动化脚本）、`tests/QA_REPORT.md`（本报告）、`screenshots/*.png`（8 张验收截图）。

E2E 内部自动完成：随机端口静态服务 → headless Chrome → 每轮清空 localStorage 重建 12 客户 + 12 产品 → 真实交互 / 下拉联动 / 键盘 / 下载 / 打印媒体 → Python openpyxl 读回 → 截图与断言 → 汇总结果与退出码。

## 附录 B：被测文件指纹（SHA-256 前 16 位，v1.2 最终验收运行）

| 文件 | 字节数 | SHA-256（前 16 位） |
| --- | --- | --- |
| `index.html` | 7226 | `D98592D924946B70` |
| `js/schema.js` | 35890 | `3C2B40CBD8FF21F1` |
| `js/core.js` | 20868 | `68FFC82BAA34A986` |
| `js/parser.js` | 69659 | `E596BDB6E08C0A69` |
| `js/xlsx.js` | 42913 | `241E1E6E99F191BA` |
| `js/store.js` | 29162 | `B9A5E79D9A82619D` |
| `js/app.js` | 103049 | `63CA363C7A8B933E` |
| `css/app.css` | 90123 | `9D7132A103C0CD83` |
| `tests/e2e.mjs` | 153968 | `AE4CB437FFADE2B6` |

> 说明：以上为 v1.2 最终验收运行的被测文件指纹。`css/app.css` 为 ui-dev 的冻结版（2026-11-22 02:52:51）；若发布前任一文件再改动，必须重跑 `node tests/e2e.mjs` 并重新生成本报告第 6 节截图，同时更新本表指纹。
