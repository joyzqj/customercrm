# 客户通 CustomerCRM — 端到端验收报告（QA Report）

| 项目 | 内容 |
| --- | --- |
| 验收对象 | 客户通 CustomerCRM v1.0.0（静态单页应用，`customer-crm/`） |
| 验收依据 | `docs/CONTRACT.md` §8 公共测试接口 + §9 E2E 要求；task-4 验收清单 |
| 验收人 | qa-dev |
| 验收日期 | 2026-11-22（Asia/Shanghai） |
| 自动化脚本 | `tests/e2e.mjs`（零 npm 依赖，`tests/cdp.mjs` + headless Chrome 157） |
| 最终结果 | **72 / 72 用例 PASS，退出码 0**（2026-11-22 00:23:41–00:24:00 最终验收运行，耗时 19s）；无遗留产品缺陷（P0/P1/P2 = 0） |
| 发布结论 | **可以发布（PASS）** |

---

## 1. 概要

- 共执行 **72 个端到端用例**，覆盖首屏、仪表盘、客户台账、增删改查与跟进、键盘快捷键、智能识别、XLSX 导出、设置与数据管理、响应式与打印、数据损坏兜底、截图交付物，全部 PASS。
- 用例不是「只调 API」的假 E2E：所有操作都通过真实 DOM 点击/输入/键盘事件驱动，断言以 `data-testid`、语义 `name`、真实 DOM 文本与 store 数据交叉比对；Excel 导出走真实浏览器下载（`Page.setDownloadBehavior`）并用 **openpyxl** 读回校验；打印媒体用 `Emulation.setEmulatedMedia` 验证 computed style。
- 每次运行/每组用例开头 `localStorage.clear()` + `reload`，自动重建 12 条示例数据，用例之间相互隔离、可重复运行（本次连续多轮运行结果一致）。
- 交付 8 张验收截图（1440×900 基准 + 390×844 移动端），尺寸与内容均已校验。
- 验收期间发现并跟踪了 2 个契约级缺陷（`btn-add-followup` / `row-toggle` 等 data-testid 缺失），均由 Lead 修复后回归通过；**当前无遗留缺陷**。

### 1.1 最终验收运行（发布门禁）

| 项 | 值 |
| --- | --- |
| 运行命令 | `node tests/e2e.mjs` |
| 运行时间 | 2026-11-22 00:23:41 – 00:24:00（Asia/Shanghai，耗时 19s，独占运行） |
| 结果 | **72 / 72 PASS，FAIL 0，SKIP 0，退出码 0** |
| 分组 | A 9/9、B 15/15、C 11/11、D 4/4、E 7/7、F 5/5、G 10/10、H 8/8、I 1/1、S 2/2 |
| 截图 | 8 张全部由本次运行重新生成（00:23:47–00:23:59），尺寸/内容校验通过 |
| 被测版本 | 前端文件指纹见附录 B；`css/app.css` 为 ui-dev 冻结版（79737B / `3BF93CD1B6065EEE`） |
| 契约 | `docs/CONTRACT.md` v1.0（含 2026-11-22 00:17 同步的 §7.3 新列优先级口径） |

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
| 数据 | 浏览器 `localStorage['customer-crm.v1']`，每轮清空重建 12 条示例数据 |
| 截图 | `screenshots/{dashboard,customers,customers-expanded,recognize,form,settings,mobile-dashboard,mobile-customers}.png` |

补充（验收时复跑，仅作发布参考，非本任务交付物）：

| 单元测试 | 结果 |
| --- | --- |
| `node tests/test-schema.mjs` | 94/94 PASS，exit 0 |
| `node tests/test-parser.mjs` | 49/49 PASS，exit 0 |
| `node tests/test-xlsx.mjs` | 132 项断言 PASS（openpyxl 读回 9 次），exit 0 |

---

## 3. 验收方法

1. **真实交互**：`PageSession.fill()` 用原生 value setter + `input/change/blur` 事件填表；点击全部走真实 DOM `click()`；键盘用例走 CDP `Input.dispatchKeyEvent`（`/`、`N`、`Esc`）。
2. **DOM ↔ store 交叉断言**：表格行数/顺序/列值、统计卡数字、age 色调分布、行展开跟进内容均与 `CustomerCRM.store` / `app.filtered()` / `app.stats()` 实时比对，避免「界面自说自话」。
3. **真实下载**：每次导出前清空临时下载目录并 `Page.setDownloadBehavior`，轮询等待文件落盘（排除 `.crdownload`）；先做 Node ZIP 结构断言（PK 头、部件、行数），再由 openpyxl 做权威读回（工作表名、19/10 列、行数、冻结、自动筛选、客户名顺序）。
4. **打印媒体**：`Emulation.setEmulatedMedia({media:'print'})` 后读 `getComputedStyle`/`offsetParent`，验证侧栏/顶栏/工具栏/操作列隐藏、3 个低优先列恢复、`thead=table-header-group`。
5. **截图是断言的一部分**：每张截图都校验 PNG 尺寸（1440×900 / 390×844）与体积下限，防止空白页/错尺寸；`form.png` 还断言截图前表单已填入可见值。
6. **失败可定位**：用例按 A–S 分组编号，失败输出「实际/期望 + 等待表达式」，脚本 exit 1；全绿 exit 0。

---

## 4. 用例与结果

> 结果均来自最终验收运行（`node tests/e2e.mjs`）。证据列中的 `A1…` 为脚本运行日志中的用例编号。

### A. 首屏启动与仪表盘（9/9 PASS）

| 编号 | 场景 | 步骤 | 期望 | 结果 | 证据 |
| --- | --- | --- | --- | --- | --- |
| A1 | 首屏启动 | 打开 `index.html`，等待 `CustomerCRM.app` | hash=`#/dashboard`、`getRoute()=dashboard`、导航高亮、标题含「仪表盘」 | PASS | 运行日志 A1 |
| A2 | 首屏无错误 | 全程监听 console/pageerror/log.error | 0 条错误级输出 | PASS | 运行日志 A2 |
| A3 | 示例数据 | 首次访问未预置 localStorage | 自动写入 12 条，11 个字段均为 string，含 Michael Braun | PASS | 运行日志 A3 |
| A4 | 统计卡一致性 | 读 4 卡 textContent 与 `app.stats()` | total/deal/open/overdue 完全相等，total=12 | PASS | 运行日志 A4；dashboard.png |
| A5 | 仪表盘区块 | 统计卡/国家分布/状态分布/最近动态/优先跟进 | 4 卡、8 条动态、优先跟进条数 = `stats.overdueList.length` | PASS | 运行日志 A5 |
| A6 | 保存指示器 | 读 `#save-state` | `data-state="saved"` 且文本非空 | PASS | 运行日志 A6 |
| A7 | 公共接口自检 | 校验 CONTRACT §8 方法/常量 | app 15 个方法、store 14 个方法、parser/xlsx/SAMPLE_TEXTS≥3 就绪 | PASS | 运行日志 A7 |
| A8 | 1440 无横向溢出 | `documentElement.scrollWidth`/`body` vs `innerWidth` | ≤ innerWidth + 1 | PASS | 运行日志 A8 |
| A9 | 仪表盘截图 | `Page.captureScreenshot` | PNG 1440×900、体积 ≥20KB | PASS | `screenshots/dashboard.png` |

### B. 客户台账（15/15 PASS）

| 编号 | 场景 | 步骤 | 期望 | 结果 | 证据 |
| --- | --- | --- | --- | --- | --- |
| B1 | 表格结构 | 打开 `#/customers` | 12 行、每行 14 个 td、表头 12 个字段标签、关键 data-testid 齐全 | PASS | 运行日志 B1 |
| B2 | 关键词搜索 | 搜索框输入「德国」 | 命中 1 行，DOM 顺序 = `app.filtered()`，显示已筛选标记 | PASS | 运行日志 B2 |
| B3 | 搜索空结果 | 输入不存在的关键词 | `.empty`「没有匹配的客户」；重置后恢复 12 行 | PASS | 运行日志 B3 |
| B4 | 国家筛选 | 国家下拉选「德国」 | 1 行，与 store 计算一致，下拉回显正确 | PASS | 运行日志 B4 |
| B5 | 状态筛选 | 跟进状态选「已报价」 | 3 行且状态列均为已报价 | PASS | 运行日志 B5 |
| B6 | 成交筛选 | 是否成交选「已成交」 | 2 行且成交列均为已成交 | PASS | 运行日志 B6 |
| B7 | 组合筛选 | 已报价 + 德国 | 1 行（Michael Braun） | PASS | 运行日志 B7 |
| B8 | 重置筛选 | 点「重置」 | filters 全空、下拉归位、12 行 | PASS | 运行日志 B8 |
| B9 | 排序：首次询盘 | 连点列头 3 次 | desc → asc → 恢复 updated-desc；DOM 顺序=app.filtered()，日期单调 | PASS | 运行日志 B9 |
| B10 | 排序：客户名称 | 连点列头 2 次 | desc → asc，DOM 顺序=app.filtered() | PASS | 运行日志 B10 |
| B11 | 排序：距上次跟进 | 点列头 1 次 | daysSince-desc，DOM 顺序 = 派生排序 | PASS | 运行日志 B11 |
| B12 | 展开行 | 点行首 row-toggle | 行详情展开、跟进 chips=min(3,n)、内容与 store 一致、首条标「最近」 | PASS | 运行日志 B12；customers-expanded.png |
| B13 | 收起行 | 再点 row-toggle | 详情 hidden=true、无 `row--expanded` | PASS | 运行日志 B13 |
| B14 | 台账截图 | 收起态/展开态各截 1 张 | PNG 1440×900 | PASS | `customers.png`、`customers-expanded.png` |
| B15 | 无控制台错误 | 全程监听 | 0 条 | PASS | 运行日志 B15 |

### C. 客户增删改与跟进（11/11 PASS）

| 编号 | 场景 | 步骤 | 期望 | 结果 | 证据 |
| --- | --- | --- | --- | --- | --- |
| C1 | 打开新建弹窗 | 点 `[data-testid=new-customer]` | 11 个 name 字段、2 个必填标记、save/cancel/form-error testid、role=dialog、日期默认今天 | PASS | 运行日志 C1 |
| C2 | 表单截图 | 填 5 个字段后截图 | 截图前表单值已写入；PNG 1440×900 | PASS | `form.png` |
| C3 | 必填校验 | 清空客户名称/首次询盘时间 → 保存 | `#form-error` 可见且含「首次询盘时间、客户名称」，字段级错误出现，弹窗不关闭 | PASS | 运行日志 C3 |
| C4 | 保存新客户 | 填 11 个字段 → 保存 | store 13 条、表格 13 行、弹窗关闭、成功 toast、`#save-state=saved`、字段落库正确 | PASS | 运行日志 C4 |
| C5 | 刷新持久化 | reload 页面 | localStorage 13 条、表格仍显示新客户 | PASS | 运行日志 C5 |
| C6 | 编辑客户 | 行内编辑 → 改名 → 保存 | 弹窗预填正确（mode=edit）、store 与表格同步更新、toast | PASS | 运行日志 C6 |
| C7 | 添加跟进 1~3 | 跟进弹窗添加 3 条不同日期 | 条数递增、最新在前、日期倒序 | PASS | 运行日志 C7 |
| C8 | 第 4 条跟进 | 再添加 1 条（今天） | 只保留 3 条、最新在前、最旧被截断、「距上次跟进=今天」且 `age--ok` | PASS | 运行日志 C8 |
| C9 | 删除客户 | 删除 → 取消；再删除 → 确认 | 取消不删（仍 13 条）；确认后 12 行、toast、行消失、store 无记录 | PASS | 运行日志 C9 |
| C10 | data-testid 契约 | 检查行内按钮 | `btn-add-followup`、`row-toggle` 均存在 | PASS | 运行日志 C10 |
| C11 | 无控制台错误 | 全程监听 | 0 条 | PASS | 运行日志 C11 |

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
| E1 | 识别示例 A | 粘贴 `CrmParser.SAMPLE_TEXTS[0].text` → 开始识别 | `.recog-item` 11 项、`.recog-item--filled` ≥6、摘要计数一致、证据 ≥6 条、跟进 chip ≥1 | PASS | 运行日志 E1 |
| E2 | 字段一致性 | 读识别输入框 value 与 `CrmParser.recognize()` | 逐字段相等；客户=Michael Braun、国家=德国、日期=2026-09-12、产品含保温杯 | PASS | 运行日志 E2 |
| E3 | 填入表格并核对 | 点「填入表格并核对」 | 新建表单按识别结果预填（名称/日期/国家/公司名），显示识别说明，取消后关闭 | PASS | 运行日志 E3 |
| E4 | 识别页截图 | 结果态截图 | PNG 1440×900 | PASS | `recognize.png` |
| E5 | 直接新增一行 | 点「直接新增一行」 | store +1，新记录名称/日期/国家/跟进与识别结果一致 | PASS | 运行日志 E5 |
| E6 | 空文本兜底 | 清空文本 → 开始识别 | 仅 error toast「请先粘贴一段询盘文本」，不产生结果、不抛异常 | PASS | 运行日志 E6 |
| E7 | 无控制台错误 | 全程监听 | 0 条 | PASS | 运行日志 E7 |

### F. 一键导出 Excel（5/5 PASS）

| 编号 | 场景 | 步骤 | 期望 | 结果 | 证据 |
| --- | --- | --- | --- | --- | --- |
| F1 | 无筛选导出 | 点 `[data-testid=btn-export-excel]`，等待文件落盘 | 文件名匹配 `客户跟进台账_YYYYMMDD_HHmm.xlsx` 且日期为当天；成功 toast 含 12 位客户 | PASS | 运行日志 F1；临时下载目录 |
| F2 | ZIP 结构 | Node 解析 ZIP | PK 头、体积>3000B、7 个必需部件齐全、两个工作表 XML、客户总表行数=12+1、跟进明细行数=25+1 | PASS | 运行日志 F2 |
| F3 | openpyxl 读回 | Python openpyxl 打开导出文件 | 工作表名 `客户总表/跟进明细`、19 列表头文本逐项一致、行数正确、冻结 `A2`、自动筛选已设置、客户名顺序=导出记录顺序、明细 10 列 | PASS | 运行日志 F3 |
| F4 | 筛选导出 | 先筛选「已成交」2 条 → 导出 | 仅导出筛选结果（总表 3 行、明细行数=筛选跟进数+1），openpyxl 读回客户名=当前筛选结果 | PASS | 运行日志 F4 |
| F5 | 无控制台错误 | 全程监听 | 0 条 | PASS | 运行日志 F5 |

### G. 设置与数据管理（10/10 PASS）

| 编号 | 场景 | 步骤 | 期望 | 结果 | 证据 |
| --- | --- | --- | --- | --- | --- |
| G1 | 设置页渲染 | 打开 `#/settings` | 默认阈值 7/14、seed/clear/export-json/import-json 按钮齐全 | PASS | 运行日志 G1 |
| G2 | 设置页截图 | 截图 | PNG 1440×900 | PASS | `settings.png` |
| G3 | 阈值保存/校验 | 保存 3/5；再填 30/10 保存 | 3/5 写入 store 并 toast；非法阈值被拒绝且 store 不变 | PASS | 运行日志 G3 |
| G4 | 仪表盘联动 | 回仪表盘读 4 卡 + hints + stats | 4 卡 DOM=stats()；staleDays 5 时超期数 >14 时；待跟进 hint 含 `≥3`、超期 hint 含 `≥5` | PASS | 运行日志 G4 |
| G5 | age 颜色分布 | 台账读 `.age--ok/warn/danger/none` 计数 | 与 `store.deriveAll()` 按当前阈值的统计完全一致，共 12 个 | PASS | 运行日志 G5 |
| G6 | JSON 备份导出 | 点「导出备份」 | 文件名 `客户通备份_YYYYMMDD_HHmm.json`；合法 JSON；app/version；12 位客户；含设置 | PASS | 运行日志 G6；临时下载目录 |
| G7 | 清空数据 | 点「清空数据」→ 确认 | store 0；台账空状态引导含「智能识别/新建客户/载入示例数据」；仪表盘总数=0 | PASS | 运行日志 G7 |
| G8 | JSON 恢复 | 选择备份文件（FileReader 真实链路） | store 恢复 12 条、表格 12 行、成功 toast、含示例客户 | PASS | 运行日志 G8 |
| G9 | 载入示例数据 | 清空后在空状态点「载入示例数据」 | store 恢复 12 条、表格 12 行、toast 正确 | PASS | 运行日志 G9 |
| G10 | 无控制台错误 | 全程监听 | 0 条 | PASS | 运行日志 G10 |

### H. 响应式、侧栏抽屉与打印（8/8 PASS）

> 按 CONTRACT §7.3 于 2026-11-22 00:17 定稿的口径验收：`<=1599px` 隐藏「成立时间/团队规模/年营收」三列，改由行展开 `.row-detail-meta` 展示；列顺序为「展开｜首次询盘时间｜距上次跟进｜客户名称｜咨询产品｜国家｜公司名｜成立时间｜团队规模｜年营收｜跟进状态｜是否成交｜核心原因｜操作」。

| 编号 | 场景 | 步骤 | 期望 | 结果 | 证据 |
| --- | --- | --- | --- | --- | --- |
| H1 | 1440×900 | 1440 视口打开台账 | 无横向溢出；侧栏可见；14 个 td 都在 DOM；9 个业务列可见；3 个低优先列 display:none；操作列可见；展开行 `.row-detail-meta` 含成立时间/团队规模/年营收/首次询盘/最近更新 | PASS | 运行日志 H1 |
| H2 | 1024×800 | 1024 视口 | 无横向溢出；侧栏收起（translateX(-100%) 动画结束后 right≤1）；3 个低优先列隐藏；product/companyName/coreReason 可见 | PASS | 运行日志 H2 |
| H3 | 768 边界 + 767 | 768 与 767 两档 | 768 无溢出、侧栏收起、product 仍可见（媒体查询为 `<768`）；767 时 product/companyName/coreReason 隐藏、客户名称可见、无溢出 | PASS | 运行日志 H3 |
| H4 | 390×844 | 390 视口 | 无横向溢出；侧栏收起；菜单按钮可见 | PASS | 运行日志 H4 |
| H5 | 侧栏抽屉 | 点菜单按钮 → 点遮罩 | 打开后侧栏滑入视口、遮罩可见、`aria-expanded=true`；点遮罩后收起且侧栏移出视口 | PASS | 运行日志 H5 |
| H6 | 打印媒体 | `Emulation.setEmulatedMedia('print')` | 侧栏/顶栏/工具栏/页头操作/操作列/Toast 均 display:none；3 个低优先列恢复 table-cell；展开列隐藏；`thead=table-header-group`；内容仍可见 | PASS | 运行日志 H6 |
| H7 | 移动端截图 | 390×844 仪表盘/台账各截 1 张 | PNG 390×844 | PASS | `mobile-dashboard.png`、`mobile-customers.png` |
| H8 | 无控制台错误 | 全程监听 | 0 条 | PASS | 运行日志 H8 |

### I. 数据损坏兜底（1/1 PASS）

| 编号 | 场景 | 步骤 | 期望 | 结果 | 证据 |
| --- | --- | --- | --- | --- | --- |
| I1 | 非法 JSON 兜底 | `localStorage.setItem('customer-crm.v1','{bad json')` → reload | 应用正常启动不白屏；自动重建 12 条示例数据；`localStorage` 被重写为合法 payload；无 pageerror | PASS | 运行日志 I1 |

### S. 交付物校验（2/2 PASS）

| 编号 | 场景 | 步骤 | 期望 | 结果 | 证据 |
| --- | --- | --- | --- | --- | --- |
| S1 | 8 张截图齐备 | 校验存在性、PNG 合法性、尺寸、体积 | 6 张 1440×900 + 2 张 390×844，均 ≥20KB；尺寸完全匹配 | PASS | 运行日志 S1 |
| S2 | 命名与 README 一致 | 校验文件名规范 | `screenshots/*.png` 命名与 README 引用路径一致 | PASS | 运行日志 S2 |

---

## 5. 缺陷清单

### 5.1 已修复（验收期间发现，修复后回归通过）

| 编号 | 严重级 | 现象 | 复现步骤 | 期望行为 | 归属 | 状态 |
| --- | --- | --- | --- | --- | --- | --- |
| QA-001 | P1 | 客户台账行内「记录跟进」、行展开「记录一次跟进」、仪表盘「优先跟进·记跟进」按钮缺少 CONTRACT §8 要求的 `data-testid="btn-add-followup"`（仅有 `data-action`），E2E 与外部自动化无法稳定定位 | 打开 `#/customers`，执行 `document.querySelector('[data-testid="btn-add-followup"]')` → `null` | 三个「记录跟进」入口均带 `data-testid="btn-add-followup"` | Lead（js/app.js） | 已修复；C7/C8/C10/B1 回归 PASS |
| QA-002 | P2 | CONTRACT §8 列出的 `row-toggle` / `toast-stack` / `modal-root` 在实现中分别是 class / id，没有 `data-testid` | 执行 `document.querySelector('[data-testid="row-toggle"]')` → `null` | 与契约一致：`.row-toggle` 带 `data-testid="row-toggle"`；`#toast-stack`/`#modal-root` 带同名 testid | Lead（js/app.js、index.html） | 已修复；B1/C10 回归 PASS |

### 5.2 遗留缺陷

**无**（P0 / P1 / P2 均为 0）。

### 5.3 说明项（非缺陷，已在用例中固化口径）

| 编号 | 说明 |
| --- | --- |
| N-01 | `stats().open`（待跟进卡）语义 = 「未成交且未流失」的客户总数，与 `warnDays` 无关；`warnDays` 只影响卡片提示文案与 age 色调，`staleDays` 影响 `overdue`。E2E 按此语义断言（G4/G5：DOM == stats()、超期数随 staleDays 变化、hint 文案随阈值变化）。 |
| N-02 | 响应式断点语义（契约 v1.0 §7.3）：`<=1599px` 隐藏「成立时间/团队规模/年营收」三列，改由行展开 `.row-detail-meta` 展示，`>=1600px` 显示完整 14 列；`<768px` 隐藏「咨询产品/公司名/核心原因」与「距上次跟进」；`<=560px` 极窄屏再隐藏「首次询盘时间」（可在行展开详情查看）；恰好 768px 时 product 仍可见（媒体查询为 `max-width: 767px`）。H1/H2/H3 已覆盖该边界。 |
| N-03 | 表格列顺序为「展开｜首次询盘时间｜距上次跟进｜客户名称｜咨询产品｜国家｜公司名｜成立时间｜团队规模｜年营收｜跟进状态｜是否成交｜核心原因｜操作」。E2E 全部使用 `data-col` 选择器断言，不依赖列索引；1440 下 3 列虽 `display:none`，14 个 td 仍在 DOM。 |
| N-04 | 打印媒体恢复被响应式隐藏的业务列（成立时间/团队规模/年营收/距上次跟进），隐藏展开列与操作列，`thead` 重复表头，符合「打印全量台账」的用途（H6 覆盖）。 |
| N-05 | 移动端（`<768px`）Toast 固定在底部，避免遮挡页头标题与操作按钮（H4/H5 覆盖移动端布局与抽屉）。 |
| N-06 | 行展开详情额外提供「首次询盘/最近更新/已跟进天数」等摘要字段，弥补桌面端折叠列与移动端隐藏列；B12/H1 已断言 `.row-detail-meta` 与跟进 chip 内容。 |

### 5.4 测试资产迭代记录（非产品缺陷，留档备查）

| 问题 | 处理 |
| --- | --- |
| `tests/e2e.mjs` 曾被 PowerShell `Get-Content/Set-Content` 批量改写导致 GBK 双重编码损坏（SyntaxError） | 改用 `write`/`edit` 工具以 UTF-8 重写；`node --check` + `node -e` 检查无 BOM、无 U+FFFD |
| openpyxl 校验脚本在 Windows 管道 stdout 默认 cp936，Node 按 UTF-8 解码导致工作表名乱码 | 给 Python 子进程加 `PYTHONIOENCODING=utf-8` + `PYTHONUTF8=1`，脚本输出 `ensure_ascii=True` |
| 窄屏断言读到侧栏 `transform` 过渡动画中间态 | 改为 `waitUntil` 等待动画结束（`right<=1` / `left>=-1`）再断言 |
| 损坏兜底用例假定 reload 后一定在 dashboard，实际 hash 会保留在 `#/customers` | 先等页面渲染，再显式 `goto('#/dashboard')` |
| `form.png` 截图曾出现「DOM 已填值但截图为空表单」的绘制时序问题 | 截图前 blur + 断言表单 5 个字段值 + 等待 200ms，截图稳定为填值态 |

---

## 6. 截图索引

> 全部截图由最终验收运行（2026-11-22 00:23:47–00:23:59）直接生成，均为有效 PNG。

| 文件 | 尺寸 | 字节数 | 内容 | 来源用例 |
| --- | --- | --- | --- | --- |
| `screenshots/dashboard.png` | 1440×900 | 146196 | 仪表盘首屏：4 统计卡、国家/状态分布、最近跟进动态、优先跟进 | A9 |
| `screenshots/customers.png` | 1440×900 | 226534 | 客户台账：12 行、搜索/筛选/排序工具栏、13 个业务列（3 列折叠） | B14 |
| `screenshots/customers-expanded.png` | 1440×900 | 213046 | 客户台账展开首行：`.row-detail-meta`（成立时间/团队规模/年营收）与最近三次跟进 | B14 |
| `screenshots/recognize.png` | 1440×900 | 175459 | 智能识别结果态：示例 A 原文 + 已识别 11/11 项字段与原文证据 | E4 |
| `screenshots/form.png` | 1440×900 | 198596 | 新建客户弹窗（已填值态）：11 字段分组表单 + 保存/取消 | C2 |
| `screenshots/settings.png` | 1440×900 | 123385 | 设置页：数据管理、跟进提醒阈值、存储信息 | G2 |
| `screenshots/mobile-dashboard.png` | 390×844 | 51770 | 移动端仪表盘（单列统计卡） | H7 |
| `screenshots/mobile-customers.png` | 390×844 | 63787 | 移动端客户台账（抽屉收起、表格横向滚动容器） | H7 |

---

## 7. 结论

1. 最终验收运行（2026-11-22 00:23:41–00:24:00，独占执行）：`node tests/e2e.mjs` **72/72 用例 PASS，退出码 0**，覆盖 CONTRACT v1.0 §9 要求的全部 E2E 场景 + task-4 补充项（键盘快捷键、打印媒体、导出文件名/筛选导出、数据损坏兜底）。
2. 8 张验收截图齐备，尺寸/内容校验通过；关键页面（仪表盘、台账、展开行、识别、表单、设置）与移动端均有截图。
3. 缺陷：验收期间 2 个契约级问题（P1/P2，data-testid 缺失）已由 Lead 修复并回归通过；**当前 0 个遗留 P0/P1/P2 缺陷**。
4. 单元测试（schema/parser/xlsx）验收时复跑全部 PASS，与 E2E 结果一致。

**发布结论：可以发布（PASS）。**
建议发布前置动作（发布清单）：以本报告「附录 B」记录的文件指纹为准构建 `dist/` 单文件与 GitHub Pages 产物；如之后 UI/JS 再有改动，需重跑 `node tests/e2e.mjs` 并重新生成本报告第 6 节截图。

---

## 附录 A：命令与产物

```powershell
cd "C:\Users\Administrator\Desktop\Global Country\customer-crm"
node tests/e2e.mjs          # 72 用例，全绿 exit 0；失败打印编号/实际/期望并 exit 1
```

产物：`tests/e2e.mjs`（自动化脚本）、`tests/QA_REPORT.md`（本报告）、`screenshots/*.png`（8 张验收截图）。

E2E 内部自动完成：随机端口静态服务 → headless Chrome → 每轮清空 localStorage 重建示例数据 → 真实交互与下载/打印/键盘模拟 → Python openpyxl 读回 → 截图与断言 → 汇总结果与退出码。

## 附录 B：被测文件指纹（SHA-256 前 16 位，最终验收运行）

| 文件 | 字节数 | SHA-256（前 16 位） |
| --- | --- | --- |
| `index.html` | 7226 | `CD2456F819908CFD` |
| `js/schema.js` | 32943 | `625BF43DEF62A3D8` |
| `js/core.js` | 20868 | `68FFC82BAA34A986` |
| `js/parser.js` | 51104 | `3D887D68FB9783BB` |
| `js/xlsx.js` | 41775 | `5C865AFD49DEE83E` |
| `js/store.js` | 21385 | `0AF5CEEBD41FAB86` |
| `js/app.js` | 79921 | `95E3ADB466A45FE6` |
| `css/app.css` | 79737 | `3BF93CD1B6065EEE` |
| `tests/e2e.mjs` | 101196 | `F037A22877F6D7F1` |

> 说明：以上为最终验收运行的被测文件指纹。`css/app.css` 为 ui-dev 的冻结版（2026-11-22 00:21:20）；若发布前任一文件再改动，必须重跑 `node tests/e2e.mjs` 并重新生成本报告第 6 节截图，同时更新本表指纹。
