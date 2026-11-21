# 客户通 CustomerCRM — 端到端验收报告（QA Report）

| 项目 | 内容 |
| --- | --- |
| 验收对象 | 客户通 CustomerCRM **v1.1.0**（静态单页应用，`customer-crm/`） |
| 验收依据 | `docs/CONTRACT.md` **v1.1** §8 公共测试接口 + §9 E2E 要求；task-8 验收清单 |
| 验收人 | qa-dev |
| 验收日期 | 2026-11-22（Asia/Shanghai） |
| 自动化脚本 | `tests/e2e.mjs`（零 npm 依赖，`tests/cdp.mjs` + headless Chrome 157） |
| 最终结果 | **84 / 84 用例 PASS，退出码 0**（2026-11-22 01:18:20–01:18:53 最终验收运行，耗时 33s）；**无遗留产品缺陷（P0/P1/P2 = 0）** |
| 发布结论 | **可以发布（PASS）** |

---

## 1. 概要

### 1.1 v1.1 变更内容（相对 v1.0）

| # | 变更 | E2E 覆盖 |
| --- | --- | --- |
| 1 | **字段 11 → 13**：新增「产品型号 productModel」「主营业务 mainBusiness」（schema VERSION 1.1.0） | A3、C1/C4/C5/C6、E1/E2/E5、F3、H1、I1 |
| 2 | **客户台账 15 列**（展开 + 13 字段合并为 13 列 + 操作），产品名称+型号同格（型号用 `.cell-sub` 小字） | B1、B12、B17、C4、H1 |
| 3 | **台账宽版布局**：`body.route-customers`，1440 下 `.view-inner` 实测 1152px（≥1140），表格 `min-width: 1328px` | H1、H1b |
| 4 | **仪表盘国家分布改为 SVG 扇形图**：`[data-testid="country-pie"]`、`.pie-slice`（含 hover `<title>`）、`.pie-legend .legend-item`、`.legend-pct`、`.pie-total` | A10、H6b |
| 5 | **产品库**（v1.1 新增）：设置页 `#product-form` + 产品行 + 删除 + 恢复预置；`store.listProducts/productNames/productModels/addProduct/removeProduct/seedProducts/productUsage` | P1–P4、G6–G8 |
| 6 | **客户表单产品选择器**：名称下拉 + 型号联动 + 自定义输入 + 「加入产品库」；`app.selectProduct`；`data-testid=product-name-select / product-model-select / product-custom-input` | C1、C2、C4、C6、P2、P3、P5、E3 |
| 7 | **识别引擎升级为 13 字段**：样例 A filled=13/13（含 productModel=500ml、mainBusiness=户外用品） | E1、E2、E3、E5 |
| 8 | **Excel 21 + 11 列**：客户总表新增「产品型号」「主营业务」，跟进明细新增「主营业务」 | F2、F3、F4 |

> 本报告**保留 v1.0 历史结论作为对照**（见 §1.3、§5.2）；v1.0 全部用例在 v1.1 中继续覆盖并扩展。

### 1.2 最终验收运行（发布门禁）

| 项 | 值 |
| --- | --- |
| 运行命令 | `node tests/e2e.mjs`（独占运行，期间无其他人跑 e2e/写 `screenshots/` 根目录） |
| 运行时间 | 2026-11-22 01:18:20 – 01:18:53（Asia/Shanghai，耗时 33s） |
| 结果 | **84 / 84 PASS，FAIL 0，SKIP 0，退出码 0** |
| 分组 | A 10/10、B 17/17、C 11/11、P 6/6、D 4/4、E 7/7、F 5/5、G 10/10、H 10/10、I 2/2、S 2/2 |
| 截图 | 8 张全部由本次运行重新生成（01:18:26–01:18:48），尺寸/内容校验通过 |
| 被测版本 | 前端文件指纹见附录 B；`css/app.css` 为 ui-dev 冻结版（88363B / `5ECA76BB24EEF151`，2026-11-22 01:11:22） |
| 契约 | `docs/CONTRACT.md` **v1.1**（含 §7.3 v1.1 宽版台账/列优先级口径） |
| 辅助测试 | schema 104/104、parser 68/68、xlsx 150 断言、smoke 32/32、v11-check 全通过（均 exit 0） |

### 1.3 v1.0 历史结论对照（已被 v1.1 取代，保留存档）

- v1.0 最终验收运行：2026-11-22 00:23:41–00:24:00，`node tests/e2e.mjs` **72/72 PASS，exit 0**；8 张截图齐备；结论「可以发布（PASS）」。
- v1.0 验收期间发现并跟踪的 2 个契约级缺陷（QA-001 P1、QA-002 P2）均由 Lead 修复并回归通过（详见 §5.2）。
- v1.0 的 72 个用例在 v1.1 中全部保留（按 13 字段/15 列/21 列等新契约更新断言），并新增 12 个用例（A10、B16、B17、P1–P6、H1b、H6b、I2，扣除合并项后净增 12），总计 84 个。

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
| 数据 | 浏览器 `localStorage['customer-crm.v1']`：12 条示例客户 + 12 条预置产品，每轮清空重建 |
| 截图 | `screenshots/{dashboard,customers,customers-expanded,recognize,form,settings,mobile-dashboard,mobile-customers}.png` |

辅助测试（验收时复跑，作为发布参考；非本任务交付物）：

| 测试 | 结果 |
| --- | --- |
| `node tests/test-schema.mjs` | 104/104 PASS，exit 0 |
| `node tests/test-parser.mjs` | 68/68 PASS，exit 0（样例 A filled=13/13） |
| `node tests/test-xlsx.mjs` | 150 项断言 PASS（openpyxl 读回 10 次），exit 0 |
| `node tests/smoke.mjs` | 32/32 PASS，exit 0 |
| `node tests/v11-check.mjs` | v1.1 自检全部通过，exit 0 |

---

## 3. 验收方法

1. **真实交互**：`PageSession.fill()` 用原生 value setter + `input/change/blur` 事件填表；下拉联动通过真实 `change` 事件驱动；点击全部走真实 DOM `click()`；键盘用例走 CDP `Input.dispatchKeyEvent`（`/`、`N`、`Esc`）。
2. **DOM ↔ store 交叉断言**：表格行数/列值/顺序、统计卡数字、age 色调分布、行展开 meta 与跟进内容、产品库使用次数、扇形图切片/图例/百分比均与 `CrmStore` / `app.filtered()` / `app.stats()` 实时比对。
3. **v1.1 新增断言**：
   - 扇形图：容器/切片/图例/百分比合计/中心数/`<title>`/「其他 N 个国家」聚合；
   - 宽版台账：`body.route-customers`、`.view-inner` 实测宽度 ≥1140、表格 `min-width ≥1328`、与仪表盘宽度对比；
   - 产品库：12 行 → 新增 → 13 行 → 删除 → 恢复预置，及 `productNames/productModels/productUsage` 同步；
   - 产品选择器：名称→型号联动、`app.selectProduct`、自定义输入与「加入产品库」；
   - Excel：openpyxl 读回 21/11 列表头、`first_row` 中 E/F/H 列（咨询产品/产品型号/主营业务）有值且正确。
4. **真实下载**：每次导出前清空临时下载目录并 `Page.setDownloadBehavior`，轮询等待文件落盘；先做 Node ZIP 结构断言，再由 openpyxl 做权威读回。
5. **打印媒体**：`Emulation.setEmulatedMedia({media:'print'})` 后读 `getComputedStyle`：台账恢复全部业务列（含主营业务）、侧栏/工具栏/操作列隐藏、`thead=table-header-group`；仪表盘扇形图保留可见。
6. **截图是断言的一部分**：每张截图校验 PNG 尺寸（1440×900 / 390×844）与体积下限；`form.png` 断言截图前产品名称/型号/主营业务已填入。
7. **失败可定位**：用例按 A–S 分组编号，失败输出「实际/期望 + 等待表达式」，exit 1；全绿 exit 0。

---

## 4. 用例与结果

> 结果均来自最终验收运行 `node tests/e2e.mjs`。证据列中的 `A1…` 为脚本运行日志中的用例编号。

### A. 首屏启动与仪表盘（10/10 PASS）

| 编号 | 场景 | 步骤 | 期望 | 结果 | 证据 |
| --- | --- | --- | --- | --- | --- |
| A1 | 首屏启动 | 打开 `index.html`，等待 `CustomerCRM.app` | hash=`#/dashboard`、导航高亮、标题含「仪表盘」、统计卡渲染 | PASS | 运行日志 A1 |
| A2 | 首屏无错误 | 监听 console/pageerror/log.error | 0 条错误级输出 | PASS | 运行日志 A2 |
| A3 | 示例数据 | 首次访问未预置 localStorage | 版本 1.1.0、13 字段、12 客户、12 产品、示例含产品型号/主营业务 | PASS | 运行日志 A3 |
| A4 | 统计卡一致性 | 读 4 卡与 `app.stats()` | total/deal/open/overdue 完全相等，total=12 | PASS | 运行日志 A4 |
| A5 | 仪表盘区块 | 统计卡/状态分布/动态/优先跟进 | 4 卡、状态条 ≥4、动态 8 条、优先条数 = `stats.overdueList.length` | PASS | 运行日志 A5 |
| A6 | 保存指示器 | 读 `#save-state` | `data-state="saved"` 且文本非空 | PASS | 运行日志 A6 |
| A7 | 公共接口自检 | 校验 CONTRACT v1.1 §8 方法与常量 | app 18 方法（含 selectProduct/addProduct/products）、store 21 方法（含产品库 7 个）、版本 1.1.0、parser/xlsx 就绪 | PASS | 运行日志 A7 |
| A8 | 1440 无横向溢出 | doc/body scrollWidth 与 innerWidth | ≤ innerWidth + 1 | PASS | 运行日志 A8 |
| A9 | 仪表盘截图 | `Page.captureScreenshot` | PNG 1440×900、≥20KB | PASS | `screenshots/dashboard.png` |
| A10 | 国家分布扇形图 | 读 `[data-testid=country-pie]` | SVG 存在、7 切片 = 7 图例、百分比合计 100、中心数=12、切片有 path 与 `<title>`、含「其他 6 个国家」 | PASS | 运行日志 A10；dashboard.png |

### B. 客户台账（17/17 PASS）

| 编号 | 场景 | 步骤 | 期望 | 结果 | 证据 |
| --- | --- | --- | --- | --- | --- |
| B1 | 表格结构 | 打开 `#/customers` | 12 行、每行 15 td、表头 13 个业务列标签、`body.route-customers`、产品格含 `.cell-sub`、关键 testid 齐全 | PASS | 运行日志 B1 |
| B2 | 关键词搜索 | 搜索「德国」 | 1 行，DOM 顺序=`app.filtered()`，显示已筛选标记 | PASS | 运行日志 B2 |
| B3 | 搜索空结果 | 输入不存在关键词 | `.empty`「没有匹配的客户」；重置后 12 行 | PASS | 运行日志 B3 |
| B4 | 国家筛选 | 选「德国」 | 1 行，与 store 计算一致 | PASS | 运行日志 B4 |
| B5 | 状态筛选 | 选「已报价」 | 3 行且状态列均为已报价 | PASS | 运行日志 B5 |
| B6 | 成交筛选 | 选「已成交」 | 2 行且成交列均为已成交 | PASS | 运行日志 B6 |
| B7 | 组合筛选 | 已报价 + 德国 | 1 行（Michael Braun） | PASS | 运行日志 B7 |
| B8 | 重置筛选 | 点「重置」 | filters 全空、下拉归位、12 行 | PASS | 运行日志 B8 |
| B9 | 排序：首次询盘 | 连点列头 3 次 | desc→asc→默认；DOM 顺序=app.filtered()，日期单调 | PASS | 运行日志 B9 |
| B10 | 排序：客户名称 | 连点 2 次 | desc→asc，DOM 顺序=app.filtered() | PASS | 运行日志 B10 |
| B11 | 排序：距上次跟进 | 点 1 次 | daysSince-desc，DOM 顺序=派生排序 | PASS | 运行日志 B11 |
| B12 | 展开行 | 点 row-toggle | 跟进 chips=min(3,n)、内容与 store 一致、首条「最近」；`.row-detail-meta` 含产品型号/主营业务/成立时间/团队规模/年营收/首次询盘/最近更新及对应值 | PASS | 运行日志 B12；customers-expanded.png |
| B13 | 收起行 | 再点 row-toggle | 详情 hidden=true、无 `row--expanded` | PASS | 运行日志 B13 |
| B14 | 台账截图 | 收起态/展开态各 1 张 | PNG 1440×900 | PASS | `customers.png`、`customers-expanded.png` |
| B16 | 搜索主营业务 | 搜索「户外用品」 | 命中 2 行、DOM=app.filtered()、主营列含关键词 | PASS | 运行日志 B16 |
| B17 | 搜索产品型号 | 搜索「500ml」 | 命中 2 行、产品格含 500ml、`.cell-sub` 小字显示型号 | PASS | 运行日志 B17 |
| B18 | 无控制台错误 | 全程监听 | 0 条 | PASS | 运行日志 B18 |

### C. 客户增删改与跟进（11/11 PASS）

| 编号 | 场景 | 步骤 | 期望 | 结果 | 证据 |
| --- | --- | --- | --- | --- | --- |
| C1 | 新建弹窗结构 | 点 new-customer | 13 字段均可映射（product/productModel 由选择器承载）、产品选择器 3 个 testid、型号下拉初始仅占位、2 必填、role=dialog | PASS | 运行日志 C1 |
| C2 | 表单截图 | 填客户/主营 + 选产品名称/型号后截图 | 截图前产品名称=不锈钢保温杯、型号=500ml；PNG 1440×900 | PASS | `form.png` |
| C3 | 必填校验 | 清空客户名称/首次询盘时间→保存 | `#form-error` 含「首次询盘时间、客户名称」，字段级错误，弹窗不关闭 | PASS | 运行日志 C3 |
| C4 | 保存新客户 | 填 11 个文本字段 + 产品名称/型号下拉 → 保存 | store 13 条、product=不锈钢保温杯/productModel=500ml/mainBusiness 正确、弹窗关闭、表格产品格显示名称+`.cell-sub` 型号、`#save-state=saved` | PASS | 运行日志 C4 |
| C5 | 刷新持久化 | reload | localStorage 13 条客户（含 product/model/mainBusiness）+ 12 产品、表格仍显示 | PASS | 运行日志 C5 |
| C6 | 编辑客户 | 行内编辑 → 改名称/主营业务 → 保存 | 产品选择器预填正确（不锈钢保温杯/500ml）、store 与表格同步、toast | PASS | 运行日志 C6 |
| C7 | 添加跟进 1~3 | 跟进弹窗添加 3 条不同日期 | 条数递增、最新在前、日期倒序 | PASS | 运行日志 C7 |
| C8 | 第 4 条跟进 | 再添加 1 条（今天） | 只保留 3 条、最新在前、最旧被截断、「距上次跟进=今天」且 `age--ok` | PASS | 运行日志 C8 |
| C9 | 删除客户 | 取消；再确认删除 | 取消不删（13）；确认后 12 行、toast、行消失、store 无记录 | PASS | 运行日志 C9 |
| C10 | data-testid 契约 | 检查行内按钮 | `btn-add-followup`、`row-toggle` 均存在 | PASS | 运行日志 C10 |
| C11 | 无控制台错误 | 全程监听 | 0 条 | PASS | 运行日志 C11 |

### P. 产品库与产品选择器（v1.1 新增，6/6 PASS）

| 编号 | 场景 | 步骤 | 期望 | 结果 | 证据 |
| --- | --- | --- | --- | --- | --- |
| P1 | 产品库渲染 | 打开 `#/settings` 产品库卡 | 12 行、每行名称/型号/使用次数/删除按钮；表单与 testid 齐全；「共 12 个产品」；`productUsage()['不锈钢保温杯\|500ml']=1` 且行文案「1 位客户使用」 | PASS | 运行日志 P1；settings.png |
| P2 | 新增产品 | 填「测试产品/X1」→ 加入产品库 | 13 行、`productNames()` 含「测试产品」、`productModels('测试产品')=['X1']`；新建客户下拉包含该产品与型号 | PASS | 运行日志 P2 |
| P3 | 删除产品 | 点该行删除按钮 | 回 12 行、`productNames()` 移除、客户表单下拉同步移除 | PASS | 运行日志 P3 |
| P4 | 恢复预置产品 | 删除预置「不锈钢保温杯/500ml」→ 点恢复预置 | 先 11 行；恢复后 12 行，产品名与 500ml 型号恢复 | PASS | 运行日志 P4 |
| P5 | 产品联动与自定义 | 选名称→型号联动；`app.selectProduct`；切自定义并加入产品库 | 型号 ≥3 个且可选 500ml；API 切 750ml 生效；自定义输入显示/型号下拉隐藏；加入后 `productNames()` 包含自定义产品 | PASS | 运行日志 P5 |
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
| E1 | 识别示例 A | 粘贴 `CrmParser.SAMPLE_TEXTS[0].text` → 开始识别 | 13 项、filled ≥11、摘要 13 分母、证据 ≥11、productModel 含 500ml、mainBusiness 含 户外用品 | PASS | 运行日志 E1 |
| E2 | 字段一致性 | DOM 值 vs `CrmParser.recognize()` | 13 字段逐项相等；parser 样例 A filled=13/13；重点字段值正确 | PASS | 运行日志 E2 |
| E3 | 填入表格并核对 | 点「填入表格并核对」 | 新建表单预填；产品名称下拉=不锈钢保温杯、型号下拉=500ml、主营业务=户外用品；显示识别说明 | PASS | 运行日志 E3 |
| E4 | 识别页截图 | 结果态截图 | PNG 1440×900 | PASS | `recognize.png` |
| E5 | 直接新增一行 | 点「直接新增一行」 | store +1；新记录 product 含保温杯、productModel=500ml、mainBusiness=户外用品 | PASS | 运行日志 E5 |
| E6 | 空文本兜底 | 清空文本 → 开始识别 | 仅 error toast，不产生结果、不抛异常 | PASS | 运行日志 E6 |
| E7 | 无控制台错误 | 全程监听 | 0 条 | PASS | 运行日志 E7 |

### F. 一键导出 Excel（5/5 PASS）

| 编号 | 场景 | 步骤 | 期望 | 结果 | 证据 |
| --- | --- | --- | --- | --- | --- |
| F1 | 无筛选导出 | 点 `[data-testid=btn-export-excel]`，等待文件落盘 | 文件名 `客户跟进台账_YYYYMMDD_HHmm.xlsx` 且日期为当天；toast 含 12 位客户 | PASS | 运行日志 F1；临时下载目录 |
| F2 | ZIP 结构 | Node 解析 ZIP | PK 头、体积>3000B、7 个必需部件、两个工作表 XML、客户总表 12+1 行、跟进明细 25+1 行 | PASS | 运行日志 F2 |
| F3 | openpyxl 读回 | Python openpyxl 打开导出文件 | 工作表 `客户总表/跟进明细`；总表 21 列表头逐项一致；E/F/H 列（产品/型号/主营业务）值与导出首条记录一致；冻结 `A2`、自动筛选、客户名顺序；明细 11 列表头逐项一致 | PASS | 运行日志 F3 |
| F4 | 筛选导出 | 筛选「已成交」2 条 → 导出 | 仅导出筛选结果（总表 3 行、明细行数=筛选跟进数+1）；openpyxl 读回客户名=筛选结果、列数 21/11 | PASS | 运行日志 F4 |
| F5 | 无控制台错误 | 全程监听 | 0 条 | PASS | 运行日志 F5 |

### G. 设置与数据管理（10/10 PASS）

| 编号 | 场景 | 步骤 | 期望 | 结果 | 证据 |
| --- | --- | --- | --- | --- | --- |
| G1 | 设置页渲染 | 打开 `#/settings` | 默认阈值 7/14、数据管理/产品库相关按钮齐全 | PASS | 运行日志 G1 |
| G2 | 设置页截图 | 截图 | PNG 1440×900 | PASS | `settings.png` |
| G3 | 阈值保存/校验 | 保存 3/5；再填 30/10 | 3/5 写入 store；非法阈值被拒绝且 store 不变 | PASS | 运行日志 G3 |
| G4 | 仪表盘联动 | 回仪表盘 | 4 卡 DOM=stats()；staleDays 5 时超期数 >14 时；hint 文案含 `≥3`/`≥5` | PASS | 运行日志 G4 |
| G5 | age 颜色分布 | 台账读 `.age--ok/warn/danger/none` | 与 `store.deriveAll()` 当前阈值统计完全一致，共 12 个 | PASS | 运行日志 G5 |
| G6 | JSON 备份导出 | 点「导出备份」 | 文件名规范；合法 JSON；version 1.1.0；12 客户；**12 条产品库（含 不锈钢保温杯/500ml）**；含设置 | PASS | 运行日志 G6 |
| G7 | 清空数据 | 点「清空数据」→ 确认 | store 0；空状态引导（智能识别/新建客户/载入示例数据）；仪表盘总数=0；**产品库保留 12 条** | PASS | 运行日志 G7 |
| G8 | JSON 恢复 | 选择备份文件（真实 FileReader 链路） | store 恢复 12 客户、产品库 12 条、表格 12 行、成功 toast | PASS | 运行日志 G8 |
| G9 | 载入示例数据 | 清空后点「载入示例数据」 | store 恢复 12 条、表格 12 行、toast 正确 | PASS | 运行日志 G9 |
| G10 | 无控制台错误 | 全程监听 | 0 条 | PASS | 运行日志 G10 |

### H. 响应式、宽版台账、抽屉与打印（10/10 PASS）

> 按 CONTRACT v1.1 §7.3：`<=1599px` 隐藏「成立时间/团队规模/年营收」（行展开 `.row-detail-meta` 展示，`>=1600px` 显示全部 15 列）；`<768px` 隐藏产品/主营业务/公司名/核心原因与距上次跟进；`<=560px` 极窄屏再隐藏首次询盘时间；打印恢复全部业务列。

| 编号 | 场景 | 步骤 | 期望 | 结果 | 证据 |
| --- | --- | --- | --- | --- | --- |
| H1 | 1440 宽版台账 | 1440 视口打开台账 | 无横向溢出；侧栏可见；`body.route-customers`；`.view-inner` ≥1140；表格 `min-width ≥1328`；15 td；10 个业务列可见、3 列折叠；操作列可见；展开行 meta 含产品型号/主营业务等全部字段及值 | PASS | 运行日志 H1 |
| H1b | 宽版对比 | 读台账 vs 仪表盘 `.view-inner` 宽度 | 台账宽度 ≥1140 且明显大于仪表盘 | PASS | 运行日志 H1b |
| H2 | 1024×800 | 1024 视口 | 无横向溢出；侧栏收起；3 列折叠；产品/主营业务/公司名/核心原因可见 | PASS | 运行日志 H2 |
| H3 | 768 边界 + 767 | 768 与 767 两档 | 768 无溢出、侧栏收起、产品/主营业务/距上次跟进仍可见（媒体查询为 `<768`）；767 上述列隐藏、客户名称可见、无溢出 | PASS | 运行日志 H3 |
| H4 | 390×844 | 390 视口 | 无横向溢出；侧栏收起；菜单按钮可见；移动端折叠列隐藏、客户名称可见 | PASS | 运行日志 H4 |
| H5 | 侧栏抽屉 | 点菜单按钮 → 点遮罩 | 打开后侧栏滑入视口、遮罩可见、aria-expanded=true；点遮罩后收起并移出视口 | PASS | 运行日志 H5 |
| H6 | 打印媒体（台账） | `Emulation.setEmulatedMedia('print')` | 侧栏/顶栏/工具栏/页头操作/操作列/Toast 隐藏；全部业务列恢复（含主营业务）；展开列/操作列隐藏；`thead=table-header-group` | PASS | 运行日志 H6 |
| H6b | 打印媒体（仪表盘） | 打印媒体下打开仪表盘 | 扇形图容器保留可见、有实际尺寸、7 切片、中心数 12；侧栏隐藏 | PASS | 运行日志 H6b |
| H7 | 移动端截图 | 390×844 仪表盘/台账各 1 张 | PNG 390×844 | PASS | `mobile-dashboard.png`、`mobile-customers.png` |
| H8 | 无控制台错误 | 全程监听 | 0 条 | PASS | 运行日志 H8 |

### I. 数据损坏与旧数据升级兜底（2/2 PASS）

| 编号 | 场景 | 步骤 | 期望 | 结果 | 证据 |
| --- | --- | --- | --- | --- | --- |
| I1 | 非法 JSON 兜底 | 写入 `{bad json` → reload | 应用正常启动不白屏；重建 12 客户 + 12 产品；localStorage 重写为合法 payload；无 pageerror | PASS | 运行日志 I1 |
| I2 | v1.0 旧 payload 升级 | 写入无 `products` 的旧 payload → reload | 旧客户数据保留；内存自动补齐 12 条预置产品；首次写操作后 localStorage 持久化 products，客户数据不变 | PASS | 运行日志 I2 |

### S. 交付物校验（2/2 PASS）

| 编号 | 场景 | 步骤 | 期望 | 结果 | 证据 |
| --- | --- | --- | --- | --- | --- |
| S1 | 8 张截图齐备 | 校验存在性、PNG 合法性、尺寸、体积 | 6 张 1440×900 + 2 张 390×844，均 ≥20KB；尺寸完全匹配 | PASS | 运行日志 S1 |
| S2 | 命名与 README 一致 | 校验文件名规范 | `screenshots/*.png` 命名与 README 引用路径一致 | PASS | 运行日志 S2 |

---

## 5. 缺陷清单

### 5.1 v1.1 缺陷

**无遗留缺陷（P0 / P1 / P2 = 0）。** 本次 v1.1 验收未发现产品功能缺陷；自动化脚本自身 2 处断言口径问题已在收敛过程中修正，记录于 §5.4。

### 5.2 v1.0 历史缺陷（已修复，保留对照）

| 编号 | 严重级 | 现象 | 复现步骤 | 期望行为 | 归属 | 状态 |
| --- | --- | --- | --- | --- | --- | --- |
| QA-001 | P1 | 客户台账行内「记录跟进」等 3 处按钮缺少 CONTRACT §8 要求的 `data-testid="btn-add-followup"` | 打开 `#/customers` 执行 `document.querySelector('[data-testid="btn-add-followup"]')` → `null` | 三个「记录跟进」入口均带该 testid | Lead（js/app.js） | v1.0 已修复并回归；v1.1 C7/C8/C10/B1 继续 PASS |
| QA-002 | P2 | `row-toggle` / `toast-stack` / `modal-root` 无 `data-testid` | 执行 `document.querySelector('[data-testid="row-toggle"]')` → `null` | 与契约 §8 一致补 testid | Lead（js/app.js、index.html） | v1.0 已修复并回归；v1.1 B1/C10 继续 PASS |

### 5.3 已知设计取舍 / 口径说明（非缺陷，已在用例中固化）

| 编号 | 说明 |
| --- | --- |
| N-01 | `stats().open`（待跟进卡）= 「未成交且未流失」客户总数，与 `warnDays` 无关；`warnDays` 影响卡片提示文案与 age 色调，`staleDays` 影响 `overdue`（G4/G5 按此断言）。 |
| N-02 | 响应式列优先级（契约 v1.1 §7.3）：`>=1600px` 显示全部 15 列；`<=1599px` 折叠「成立时间/团队规模/年营收」到行展开 `.row-detail-meta`；`<768px` 折叠「咨询产品/主营业务/公司名/核心原因/距上次跟进」；`<=560px` 再折叠「首次询盘时间」；打印媒体恢复全部业务列。H1–H4/H6 已覆盖边界（768px 时产品列仍可见）。 |
| N-03 | 台账宽版：`body.route-customers` 下 `.view-inner` `max-width:1720px`，1440 实测内容宽 1152px（≥1140），表格 `min-width:1328px`，核心原因横向滚动、操作列右侧固定（H1/H1b 断言）。 |
| N-04 | 「产品型号」不单独占用表格列：表格在产品名称单元格内以 `.cell-sub` 小字展示，Excel 客户总表单独成列（E 列）；行展开 meta 也展示型号（B12/H1 断言）。 |
| N-05 | 产品库「恢复预置产品」为整体重置为 12 条预置数据（会覆盖自建产品），与删除入口并存；产品库不随「清空客户数据」清空（G7/P4 断言）。 |
| N-06 | 产品使用次数 = 客户记录中 `product|productModel` 精确匹配的客户数（P1 断言 1 位客户使用）。 |
| N-07 | v1.0 旧 payload 迁移：启动时在内存补齐 12 条预置产品，首次任意写操作后持久化到 localStorage（I2 断言）。 |
| N-08 | 移动端（`<768px`）Toast 固定在底部，避免遮挡页头标题与操作按钮（H4/H5 覆盖）。 |
| N-09 | 打印仪表盘保留国家分布扇形图（`@media print` 有独立配色/尺寸规则），H6b 断言。 |

### 5.4 测试资产迭代记录（非产品缺陷，留档备查）

| 问题 | 处理 |
| --- | --- |
| v1.0：`tests/e2e.mjs` 曾被 PowerShell 批量改写导致 GBK 双重编码损坏 | 改用 `write`/`edit` 工具以 UTF-8 重写；`node --check` + BOM/U+FFFD 检查 |
| v1.0：openpyxl 校验脚本在 Windows 管道 stdout 默认 cp936 导致乱码 | Python 子进程 `PYTHONIOENCODING=utf-8` + `PYTHONUTF8=1`，脚本 `ensure_ascii=True` |
| v1.0：窄屏断言读到侧栏 `transform` 过渡动画中间态 | `waitUntil` 等待动画结束（`right<=1` / `left>=-1`） |
| v1.0：损坏兜底用例假定 reload 后停在 dashboard | 先等页面渲染，再显式 `goto('#/dashboard')` |
| v1.1：产品库卡副标题断言写成 `indexOf(...)===0`（实际子串位置非 0） | 改为 `!== -1` |
| v1.1：旧 payload 升级断言要求 reload 后立即持久化 products | 按实现语义改为「首次写操作后持久化」（I2 走设置页保存触发） |

---

## 6. 截图索引

> 全部截图由最终验收运行（2026-11-22 01:18:26–01:18:48）直接生成，均为有效 PNG；v1.0 同名截图已被覆盖。

| 文件 | 尺寸 | 字节数 | 内容 | 来源用例 |
| --- | --- | --- | --- | --- |
| `screenshots/dashboard.png` | 1440×900 | 147822 | 仪表盘首屏：4 统计卡、**国家分布扇形图（7 切片 + 图例 + 中心 12）**、状态分布、最近动态、优先跟进 | A9/A10 |
| `screenshots/customers.png` | 1440×900 | 227502 | 客户台账宽版：12 行 × 15 列、产品名称+型号小字、主营业务列、搜索/筛选/排序工具栏 | B14 |
| `screenshots/customers-expanded.png` | 1440×900 | 220386 | 台账展开首行：`.row-detail-meta`（产品型号/主营业务/成立时间/团队规模/年营收/首次询盘/最近更新）与最近三次跟进 | B14 |
| `screenshots/recognize.png` | 1440×900 | 177071 | 智能识别结果态：示例 A 原文 + 已识别 13/13 项（含产品型号 500ml、主营业务 户外用品）与原文证据 | E4 |
| `screenshots/form.png` | 1440×900 | 203871 | 新建客户弹窗：13 字段分组表单，产品名称下拉（不锈钢保温杯）+ 型号下拉（500ml）+ 主营业务已填 | C2 |
| `screenshots/settings.png` | 1440×900 | 121880 | 设置页：数据管理、跟进提醒阈值、**产品库（v1.1）**、存储信息 | G2 |
| `screenshots/mobile-dashboard.png` | 390×844 | 50892 | 移动端仪表盘（单列统计卡、扇形图） | H7 |
| `screenshots/mobile-customers.png` | 390×844 | 63804 | 移动端客户台账（抽屉收起、折叠列、表格横向滚动容器） | H7 |

---

## 7. 结论

1. 最终验收运行（2026-11-22 01:18:20–01:18:53，独占执行）：`node tests/e2e.mjs` **84/84 用例 PASS，退出码 0**，覆盖 CONTRACT v1.1 §9 全部 E2E 要求 + task-8 的 10 项必测（15 列/扇形图/产品库/产品选择器/13 字段识别/21+11 列导出/宽版布局/响应式/打印/旧数据升级），并完整回归 v1.0 原有用例。
2. 8 张验收截图齐备，尺寸/内容校验通过；v1.1 新增的扇形图、宽版台账、产品选择器、13 字段识别均已在截图中体现。
3. 缺陷：v1.1 验收**未发现新增产品缺陷（P0/P1/P2 = 0）**；v1.0 的 2 个契约级问题（P1/P2）均已修复并在 v1.1 回归通过。
4. 辅助测试（schema 104/104、parser 68/68、xlsx 150 断言、smoke 32/32、v11-check）全部 PASS，与 E2E 结果一致。

**发布结论：可以发布（PASS）。**
建议发布前置动作（发布清单）：以本报告「附录 B」记录的文件指纹为准构建 `dist/` 单文件与 GitHub Pages 产物（`css/app.css` 冻结版 `5ECA76BB24EEF151`）；如之后任何前端文件再有改动，需重跑 `node tests/e2e.mjs` 并重新生成本报告第 6 节截图、更新附录 B 指纹。

---

## 附录 A：命令与产物

```powershell
cd "C:\Users\Administrator\Desktop\Global Country\customer-crm"
node tests/e2e.mjs          # 84 用例，全绿 exit 0；失败打印编号/实际/期望并 exit 1
# 辅助：node tests/test-schema.mjs / test-parser.mjs / test-xlsx.mjs / smoke.mjs / v11-check.mjs
```

产物：`tests/e2e.mjs`（自动化脚本）、`tests/QA_REPORT.md`（本报告）、`screenshots/*.png`（8 张验收截图）。

E2E 内部自动完成：随机端口静态服务 → headless Chrome → 每轮清空 localStorage 重建 12 客户 + 12 产品 → 真实交互 / 下拉联动 / 下载 / 打印 / 键盘模拟 → Python openpyxl 读回 → 截图与断言 → 汇总结果与退出码。

## 附录 B：被测文件指纹（SHA-256 前 16 位，v1.1 最终验收运行）

| 文件 | 字节数 | SHA-256（前 16 位） |
| --- | --- | --- |
| `index.html` | 7226 | `E855166197030E89` |
| `js/schema.js` | 34444 | `21ECABF9215A4577` |
| `js/core.js` | 20868 | `68FFC82BAA34A986` |
| `js/parser.js` | 61292 | `04972191D1A805ED` |
| `js/xlsx.js` | 42346 | `E9D885E6D6CB46E2` |
| `js/store.js` | 27613 | `FA159B516A08B6CD` |
| `js/app.js` | 100930 | `FC936874384898AF` |
| `css/app.css` | 88363 | `5ECA76BB24EEF151` |
| `tests/e2e.mjs` | 137443 | `F1CED5092B5D7614` |

> 说明：以上为 v1.1 最终验收运行的被测文件指纹。`css/app.css` 为 ui-dev 的冻结版（2026-11-22 01:11:22）；若发布前任一文件再改动，必须重跑 `node tests/e2e.mjs` 并重新生成本报告第 6 节截图，同时更新本表指纹。
