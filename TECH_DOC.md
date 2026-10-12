# 羽毛球局收支管理 — 技术文档

> **本文档的用途**：项目所有者会把「源码 + 本文档 + 网页导出的数据 Excel」一起发给 AI 助手进行后续开发。请在修改代码后**同步更新本文档对应章节**（见 §10 维护约定），保持文档与代码一致。
>
> 最后更新：2026-10-13

---

## 1. 项目概述

羽毛球群主使用的**单文件 Web 记账工具**：记录每场活动的收支、羽毛球库存采购、其他资金支出，自动计算盈亏，生成可发微信群的公示账单。数据本地优先（localStorage），支持 Excel 导入导出和 GitHub 私有仓库同步备份。

- 使用者：群主本人（组织者+参与者双重身份），每周 2-3 场，单场最多 48 人
- 核心价值：盈亏可追溯、账目可公示、结算提效
- 运行方式：本地双击 `index.html` 即用；线上部署于 Vercel（https://badminton-app-qiu18.vercel.app）

## 2. 文件结构

| 文件 | 作用 |
|------|------|
| `index.html` | **全部源码**（可读版）：HTML + CSS + 原生 JS 都在这一个文件里，约 980 行 |
| `index.min.html` | 压缩版（`python compress.py` 生成），**部署时改名为 index.html 上传** |
| `compress.py` | 压缩脚本：去注释/空白，不混淆变量名。每次改完 index.html 必须重新运行 |
| `sw.js` | PWA Service Worker：缓存 CDN 资源实现离线可用 |
| `manifest.webmanifest` | PWA 清单（添加到主屏） |
| `TECH_DOC.md` | 本文档 |

**技术栈**：零构建、零框架、原生 JS。CDN 依赖三个：Chart.js 4.4.1（图表）、SheetJS(xlsx) 0.18.5（Excel）、html2canvas 1.4.1（账单导出图片）。CDN 不可达时图表/Excel/图片导出会**降级提示而不报错**（各渲染函数开头有 `typeof Chart==='undefined'` 守卫）。

**部署**：仓库 https://github.com/StudyingCoder1/badminton-app（main 分支），Vercel 自动部署。上传时需 3 个文件：`index.min.html`（改名 index.html）、`sw.js`、`manifest.webmanifest`。`sw.js` 缓存逻辑变更时记得把内部 `CACHE='bm-helper-v1'` 版本号 +1。

## 3. 页面结构（4 个 Tab）

1. **📊 单场收支**：汇总卡片（累计收入/成本/盈亏/旧球回收/场次）→ 录入表单（多场地分组）→ 明细表（可展开详情、月份筛选）→ 盈亏柱状图
2. **🏸 库存台账**：球采购记录（品牌/桶数/单价/每桶球数），自动计算各品牌**加权平均成本**
3. **💳 资金支出**：备用资金购物记录（水/装备等）
4. **📈 全局汇总**：启动资金、备用资金、趋势图/成本结构图/月度图

顶部按钮：📥 导入 / 📤 导出（Excel）/ ☁️ 同步 / ⬇️ 拉取（GitHub）/ ❓ 公式 / 清空。数据变化超 3 天未备份时顶部出现黄色备份提醒条。

## 4. 数据模型（localStorage）

| 键 | 内容 |
|------|------|
| `bm_games` | 单场记录数组（核心数据，见下方 schema） |
| `bm_inventory` | 库存采购数组 |
| `bm_expenses` | 资金支出数组 |
| `bm_init_fund` | 启动资金（数字） |
| `bm_github_config` | `{owner,repo,branch,path,token}` 同步配置，token 明文本地存储 |
| `bm_court_types` | 场地类型列表（预置：萌新养生场/微微对抗场/3+自由场） |
| `bm_ball_prices` | `{品牌: 定价}`（预置：红超18/橙超8/华美F6 8） |
| `bm_last_backup` / `bm_backup_snooze` / `bm_first_use` | 备份提醒用的时间戳 |
| `bm_venues` | ⚠️ 旧版键，读取时自动迁移到 bm_court_types |

**单场记录 games[i] schema**：
```
{ id, date:'YYYY-MM-DD', platform, duration,
  groups:[分组...],                    // 一场可含多个场地分组
  recycle, recycleUnit, recycleIncome, // 旧球回收
  totalCost, totalIncome, totalPeople, totalOrg, profit }  // 整场汇总
```

**分组 groups[j] schema**（`calcGroupCore()` 的输出，字段含义）：
| 字段 | 含义 | 来源 |
|------|------|------|
| name | 场地类型 | 下拉选择（bm_court_types + 历史记录聚合，可新增） |
| people / attend / orgCount | 报名 / 实到 / 组织者参与人数 | 手填 |
| venueCost | 场地费（付给球馆的成本） | 手填 |
| ballUsed | 球消耗（个，真实消耗） | 手填，可用 ➕1筒 快捷键 |
| extraBalls | 额外加球（个，**虚拟球**：收钱但不耗库存不计成本） | 手填 |
| ballType | 球种类（品牌） | 下拉（来自库存采购的品牌） |
| ballUnit | **球成本单价** = 该品牌库存加权平均采购价 | 自动 |
| ballPrice | **球定价** = 实际向成员收取的单价 | 自动带出（bm_ball_prices），可改，改后按品牌记住 |
| price | 人均报名费 | 手填 |
| ballCost / ballRevenue / ballProfit / ballPricePerPerson / orgCost / groupCost / groupIncome / groupProfit / perPerson | 派生值 | calcGroupCore 计算 |

**库存 invs[i]**：`{id,date,channel,brand,buckets,pricePerBucket,ballsPerBucket,totalAmount,totalBalls}`
**支出 expenses[i]**：`{id,date,channel,item,quantity,unitPrice,totalAmount}`

## 5. 核心计费模型（最重要的业务逻辑）

> 计费口径（2026-10-13 与所有者确认）：**补收球费按"定价"收，成本按"库存加权采购价"算，差价即利润；额外加球按定价计收入。**
> 示例：红超定价 18 元/个、平均采购成本 15 元 → 每消耗 1 个赚 3 元差价；多加 1 个红超 = 额外收入 18 元。

`calcGroupCore(g)` 是**唯一计算入口**（表单实时计算 calcGame、保存 saveGame→readGroups 都走它）：

```
球成本   ballCost    = ballUsed × ballUnit
球费收入 ballRevenue = (ballUsed + extraBalls) × ballPrice
球费利润 ballProfit  = ballUsed×(ballPrice−ballUnit) + extraBalls×ballPrice
人均球费补收 bpp     = ballRevenue ÷ attend
组织者成本 orgCost   = orgCount × (price + bpp)     // 组织者打球免收，免收部分计入成本
分组成本 groupCost   = ballCost + venueCost + orgCost
分组收入 groupIncome = price×people + ballRevenue   // 报名费 + 球费收入
人均应付 perPerson   = price + bpp
```

整场：`totalIncome = ΣgroupIncome + recycleIncome`；`totalCost = ΣgroupCost`；`profit = 差`。
全局：`备用资金 = initFund + Σprofit − Σ支出 − Σ库存采购额`。

**展示层**（renderGame 的展开详情、billGroupLines/buildBillHtml/buildBillText 账单）会基于存储字段**重算展示值**——修改计费公式时，除 calcGroupCore 外需同步检查这两处展示逻辑。历史记录的汇总值以存储为准（不重算），编辑旧记录后保存才会按新公式重算。

## 6. 关键函数索引（index.html 内）

| 函数 | 职责 |
|------|------|
| `calcGroupCore(g)` | **计费公式唯一入口**（见 §5） |
| `readRow(row)` / `readGroups()` | 从 DOM 分组行读取输入 → calcGroupCore |
| `calcGame()` | 表单实时计算并回填自动字段 |
| `saveGame/saveInv/saveExp` | 保存（含编辑替换：editingXxxId 过滤旧记录） |
| `normGroups(g)` | 旧数据兼容归一化（无 groups 的 v1 记录、缺失字段兜底） |
| `renderGame/renderInv/renderExp/renderAll` | 各 Tab 渲染（renderGame 含月份筛选、明细展开） |
| `getBallTypes()` | 库存按品牌聚合 → 加权平均成本 |
| `getBallPrice(brand)` / `onUnitPriceInput` | 品牌定价读取与持久化（bm_ball_prices） |
| `getCourtTypeNames/Options` / `onCourtTypeChange` / 弹窗 `courtTypeModal` | 场地类型下拉+新增 |
| `copyLastGame()` | 复制上场配置（结构保留、数值清零） |
| `addBallTube(btn)` | ➕1筒 快捷键（按该品牌 ballsPerBucket 累加） |
| `showToast(msg,undoFn)` / `runUndo()` | 底部 toast；删除撤销（6 秒） |
| `checkBackupReminder/markBackedUp/snoozeBackup` | 备份提醒条 |
| `exportExcel/importExcel` | Excel 导入导出（格式见 §7） |
| `githubSync/githubPull/doSync/doPull` | GitHub Contents API 推拉 JSON |
| `buildBillHtml/buildBillText/exportBillImage` | 公示账单（弹窗/微信文本格式/图片） |
| `esc()` | HTML 转义——**所有用户输入进 innerHTML 前必须过它**（防导入恶意 Excel 窃取 token） |

## 7. Excel 导入导出格式

导出 4 个 Sheet：**单场打球收支表 / 羽毛球库存台账 / 资金支出账单 / 全局资金汇总表**。

单场表列（26 列，多分组一场多行，仅首行带场次级字段）：
```
1 场次日期 | 2 订场平台 | 3 场地时长 | 4 场地类型 | 5 报名人数 | 6 组织者参与 | 7 实到人数
| 8 人均报名费 | 9 人均球费补收 | 10 场地费 | 11 球消耗 | 12 额外加球(个) | 13 球种类
| 14 球定价(元/个) | 15 球成本 | 16 球费补收总额 | 17 分组成本 | 18 分组收入 | 19 分组盈亏
| 20 旧球回收个数 | 21 回收单价 | 22 旧球回收收入 | 23 整场总成本 | 24 整场总收入 | 25 整场盈亏
| 26 球成本单价(元)
```
末行为「累计汇总」。导入时：可选拷新增列（第 26 列），**前 25 列位置不可改动**（importExcel 按索引解析）；导入支持覆盖/追加两模式，追加自动跳过重复（date+totals+分组数匹配）；导入的 date 兼容 Excel 日期序列号（dateNorm）。

**给 AI 解析导出数据**：用 openpyxl/pandas 读上面 4 个 Sheet 即可，列含义如上表；金额单位为元。

## 8. 已知设计决策与约束（历史原因）

- 项目有一份很重的 PRD（`deliverables/product-strategy/`，Vue3+sql.js 方案）——**已被所有者否决**，坚持单文件简单工具，勿按 PRD 技术栈重构
- 不用 `prompt()/confirm()`：部分内嵌浏览器不支持（改用页面内弹窗 / toast 撤销）
- 记录 id 必须是**数字**（`uid()`），表格按钮用内联 `onclick="editGame(${id})"`
- 金额用浮点 + `r2()`（两位舍入）展示，未做整数分存储（个人工具，误差 <1 分可接受）
- Windows 上 `python http.server` 会把 .js 发成 text/plain 导致 SW 注册失败，本地测试需自定义 MIME 的服务器
- 测试环境浏览器若点击超时，用 `element.click()`（原生）替代 Playwright click
- 旧版字段语义变迁：`extraBalls` 曾是"额外收入金额"，现为"额外加球个数"；`ballUnit` 曾同时充当收费价与成本价，现仅是成本价（收费价是 ballPrice）

## 9. 待办 / 可选优化方向（backlog）

- 场地类型列表的删除/管理界面（目前只能增不能删）
- 按 PRD 的成员维度：收款清单（谁已付/未付）
- 现场快捷"+1球"（按个而非按筒）
- 移动端表格列进一步精简的开关

## 10. 维护约定（改代码时必须遵守）

1. **改计费公式**：只改 `calcGroupCore()`，并同步检查 `renderGame` 展开详情与 `billGroupLines`（展示层重算）
2. **改导出列**：新列只能**追加在末尾**，同时同步 `importExcel.parseGRow` 的索引，保持旧导出文件可导入
3. **改完必须**：`node --check` 语法校验 → `python compress.py` 重新生成 index.min.html → 部署 3 文件（见 §2）
4. 所有用户输入渲染前过 `esc()`；金额计算用 `num()/r2()`
5. **更新本文档**：改动涉及数据模型（§4）、计费公式（§5）、函数职责（§6）、导出格式（§7）时同步修改对应章节，并在文首更新日期
6. 历史数据不可破坏：`normGroups` 的兜底逻辑要保留；旧 localStorage 键（如 bm_venues）读取时做迁移

## 11. 本地开发与测试

- 直接双击 `index.html`（file:// 下 SW 不注册，属正常）
- 起本地服务器测试：`python -m http.server` 需注意 §8 的 MIME 问题；图表/CDN 需联网
- 验证清单参考：保存→表格出现、编辑→回填正确且不重复、删除→可撤销、账单文本格式、导出再导入往返一致、刷新后 localStorage 持久
