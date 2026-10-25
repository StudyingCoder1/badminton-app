# 羽毛球局收支管理 — 技术文档

> **本文档的用途**：项目所有者会把「源码 + 本文档 + 网页导出的数据 Excel」一起发给 AI 助手进行后续开发。请在修改代码后**同步更新本文档对应章节**（见 §10 维护约定），保持文档与代码一致。
>
> 最后更新：2026-10-25（v2 大版本：多球种分组 / 旧球回收独立 / 回档系统 / Excel v2 格式）

---

## 1. 项目概述

羽毛球群主使用的**单文件 Web 记账工具**：记录每场活动的收支、羽毛球库存采购、其他资金支出、旧球回收，自动计算盈亏，生成可发微信群的公示账单。数据本地优先（localStorage），支持 Excel 导入导出和 GitHub 私有仓库同步备份。

- 使用者：群主本人（组织者+参与者双重身份），每周 2-3 场，单场最多 49 人
- 核心价值：盈亏可追溯、账目可公示、结算提效、数据可回档
- 运行方式：本地双击 `index.html` 即用；线上部署于 Vercel（https://badminton-app-qiu18.vercel.app）

## 2. 文件结构

| 文件 | 作用 |
|------|------|
| `index.html` | **全部源码**（可读版）：HTML + CSS + 原生 JS 都在这一个文件里，约 1150 行 |
| `index.min.html` | 压缩版（`python compress.py` 生成），**部署时改名为 index.html 上传** |
| `compress.py` | 压缩脚本：去注释/空白，不混淆变量名。每次改完 index.html 必须重新运行 |
| `sw.js` | PWA Service Worker：缓存 CDN 资源实现离线可用 |
| `manifest.webmanifest` | PWA 清单（添加到主屏） |
| `TECH_DOC.md` | 本文档 |

**技术栈**：零构建、零框架、原生 JS。CDN 依赖三个：Chart.js 4.4.1（图表）、SheetJS(xlsx) 0.18.5（Excel）、html2canvas 1.4.1（账单导出图片）。CDN 不可达时图表/Excel/图片导出会**降级提示而不报错**（各渲染函数开头有 `typeof Chart==='undefined'` 守卫）。

**部署**：仓库 https://github.com/StudyingCoder1/badminton-app（main 分支），Vercel 自动部署。上传时需 3 个文件：`index.min.html`（改名 index.html）、`sw.js`、`manifest.webmanifest`。`sw.js` 缓存逻辑变更时记得把内部 `CACHE='bm-helper-v1'` 版本号 +1。

## 3. 页面结构（4 个 Tab）

1. **📊 单场收支**：汇总卡片 → 录入表单（多场地分组，**每组含多条球种行**）→ 明细表（月份筛选、点击展开详情、行内 📢账单/✏️编辑/🗑删除）→ 盈亏柱状图。表单下方有「📋 复制上场」「🕘 回档」按钮
2. **🏸 库存台账**：球采购记录（渠道/品牌均为可新增下拉），自动计算各品牌**加权平均成本**与当前库存（入库−消耗）
3. **💳 资金支出 + ♻️旧球回收**：支出表单/明细 + **独立回收表单/明细**（回收是多场积累后一波统计，不再挂在单场上）
4. **📈 全局汇总**：启动资金、备用资金、趋势图/成本结构图/月度图

顶部按钮：📥 导入 / 📤 导出 / ☁️ 同步 / ⬇️ 拉取（GitHub）/ ❓ 公式 / 清空。数据变化超 3 天未备份时顶部出现黄色备份提醒条。

## 4. 数据模型（localStorage）

| 键 | 内容 |
|------|------|
| `bm_games` | 单场记录数组（核心数据，见下方 schema） |
| `bm_inventory` | 库存采购数组 |
| `bm_expenses` | 资金支出数组 |
| `bm_recycles` | **独立旧球回收记录数组**：`{id,date,count,unitPrice,income,note}` |
| `bm_backups` | **回档备份**（最多 8 份，总量超 1.5MB 淘汰最旧）：`{id,time,label,data:{games,invs,expenses,recycles,initFund}}` |
| `bm_init_fund` | 启动资金（数字） |
| `bm_github_config` | `{owner,repo,branch,path,token}` 同步配置，token 明文本地存储 |
| `bm_court_types` | 场地类型列表（预置：萌新养生场/微微对抗场/3+自由场） |
| `bm_ball_prices` | `{品牌: 定价}`（预置：红超18/橙超8/华美F6 8） |
| `bm_channels` | 采购渠道列表（预置：拼多多/淘宝/京东/阿里巴巴，支持下拉新增） |
| `bm_brands` | 球品牌列表（预置：橙超/红超/金红超，支持下拉新增） |
| `bm_last_backup` / `bm_backup_snooze` / `bm_first_use` | 备份提醒用的时间戳 |
| `bm_venues` | ⚠️ 旧版键，读取时自动迁移到 bm_court_types |

**单场记录 games[i] schema**：
```
{ id, date:'YYYY-MM-DD', platform, duration,
  groups:[分组...],                    // 一场可含多个场地分组（同名分组可以是不同物理场地，互不合并）
  recycle/recycleUnit/recycleIncome,   // ⚠️ 旧版遗留字段：仅旧数据有值；编辑旧场次保存时自动迁移为独立回收记录并置零
  totalCost, totalIncome, totalPeople, totalOrg, profit }
```

**分组 groups[j] schema**（`calcGroupCore()` 的输出）：
| 字段 | 含义 | 来源 |
|------|------|------|
| name | 场地类型 | 下拉选择（预置∪历史，可新增） |
| people / attend / orgCount | 报名 / 实到 / 组织者参与人数 | 手填 |
| venueCost | 场地费（付给球馆的成本） | 手填 |
| price | 人均报名费 | 手填 |
| **balls[]** | **球种行数组**（v2 核心变化）：每项 `{ballType, ballUsed(消耗), extraBalls(额外加球个), ballPrice(定价), ballUnit(成本)}` | 球种类下拉+定价可改（按品牌记住）、消耗/加球手填、成本自动 |
| ballCost / ballRevenue / ballProfit / ballPricePerPerson / orgCost / groupCost / groupIncome / groupProfit / perPerson | 派生值 | calcGroupCore 计算 |

旧格式分组（顶层 ballType/ballUsed/extraBalls/ballPrice/ballUnit）由 `normGroups()` 自动转为 balls[] 兼容读取。

**库存 invs[i]**：`{id,date,channel,brand,buckets,pricePerBucket,ballsPerBucket,totalAmount,totalBalls}`
**支出 expenses[i]**：`{id,date,channel,item,quantity,unitPrice,totalAmount}`

## 5. 核心计费模型（最重要的业务逻辑）

> 计费口径（所有者确认）：**补收球费按"定价"收，成本按"库存加权采购价"算，差价即利润；额外加球按定价计收入（不耗库存不计成本）。**
> 示例：红超定价 18、平均采购成本 15 → 每消耗 1 个赚 3 元；多加 1 个红超 = 额外收入 18 元。

`calcGroupCore(g)` 是**唯一计算入口**（表单实时计算 calcGame、保存 saveGame→readGroups、账单 billGroupLines 都走它）：

```
球成本   ballCost    = Σ balls[].ballUsed × ballUnit
球费收入 ballRevenue = Σ (ballUsed + extraBalls) × ballPrice
球费利润 ballProfit  = Σ [ballUsed×(ballPrice−ballUnit) + extraBalls×ballPrice]
人均球费补收 bpp     = ballRevenue ÷ attend
组织者成本 orgCost   = orgCount × (price + bpp)     // 组织者打球免收，免收部分计入成本
分组成本 groupCost   = ballCost + venueCost + orgCost
分组收入 groupIncome = price×people + ballRevenue
人均应付 perPerson   = price + bpp
```

整场：`totalIncome = ΣgroupIncome`（v2 不含回收）；`profit = totalIncome − totalCost`。
**回收**：独立记录，`recycleIncomeTotal() = Σbm_recycles.income`；`legacyRecycleIncome() = Σ旧场次遗留recycleIncome`；`allRecycleIncome() = 两者之和`（Tab0/Tab3 回收卡显示值）。
全局：`备用资金 = initFund + Σprofit − Σ支出 − Σ库存采购额 + recycleIncomeTotal()`。

**展示层**：renderGame 展开详情、buildBillHtml/Text 均基于 balls[] 重算展示。修改计费公式时只改 calcGroupCore，展示层自动跟随（它们都调用它）。

## 6. 关键函数索引（index.html 内）

| 函数 | 职责 |
|------|------|
| `calcGroupCore(g)` | **计费公式唯一入口**，含 balls[] 多球种与旧格式兜底（见 §5） |
| `readRow(row)` / `readGroups()` | 从 DOM 分组行（含 .ball-line 球种行）读取输入 |
| `calcGame()` | 表单实时计算并回填自动字段 |
| `addBallLine / removeBallLine / onBallLineTypeChange / onBallLinePriceInput` | 球种行的增删改（定价按品牌记住，仅 p>0 时持久化） |
| `saveGame/saveInv/saveExp/saveRecycle` | 保存（编辑替换；saveGame 自动迁移旧回收字段→独立回收记录；均带 250ms 防双击 saveDup） |
| `normGroups(g)` | 旧数据兼容归一化（v1 无 groups、旧单球字段→balls[]） |
| `renderGame/renderInv/renderExp/renderAll` | 各 Tab 渲染；renderInv 的库存消耗按 balls[] 求和；renderExp 含回收明细表 |
| `getBallTypes()` | 库存按品牌聚合 → 加权平均成本 |
| `getBallPrice(brand)` | 品牌定价（bm_ball_prices，未配置回退成本） |
| `getCourtTypeNames` / `getChannelNames` / `getBallBrandNames` | 各下拉选项 = 预置 ∪ 历史记录（brands 含场次用过的球种） |
| `onCourtTypeChange` / `onListChange` + `courtTypeModal` / `listModal` | 场地类型/渠道/品牌的新增弹窗 |
| `copyLastGame()` | 复制上场配置（结构含球种行保留、数值清零） |
| `makeBackup / restoreBackup / openBackupModal` | 回档系统：拉取/导入/清空前自动备份；恢复需二次点击确认，回档前也自动备份 |
| `saveRecycle / editRecycle / delRecycle` | 独立回收记录管理 |
| `showToast(msg,undoFn)` / `runUndo()` | 底部 toast；删除撤销（6 秒） |
| `exportExcel / importWorkbookData / doImportWorkbook / applyImportMode` | Excel 导入导出；有本地数据时弹**页面内**模式选择弹窗（importModeModal） |
| `githubSync/githubPull/doSync/doPull` | GitHub Contents API 推拉；payload 含全部数据+品牌定价等配置；拉取前自动备份 |
| `buildBillHtml/buildBillText/exportBillImage` | 公示账单（多球种逐行展示/微信文本格式/图片） |
| `safeParseArray(key)` | localStorage 数组安全读取（损坏时重置并警告，防整页挂死） |
| `esc()` | HTML 转义——**所有用户输入进 innerHTML 前必须过它** |

## 7. Excel 导入导出格式（v2）

导出 5 个 Sheet：**单场打球收支表 / 羽毛球库存台账 / 资金支出账单 / 旧球回收记录 / 全局资金汇总表**。

单场表 27 列，**每个球种一行**；分组级字段（4-10、17-19 列）仅每组首行填，场次级字段（1-3、20-25 列）仅每场首行填，空值写空字符串：
```
1 场次日期 | 2 订场平台 | 3 场地时长 | 4 场地类型 | 5 报名人数 | 6 组织者参与 | 7 实到人数
| 8 人均报名费 | 9 人均球费补收 | 10 场地费 | 11 球消耗 | 12 额外加球(个) | 13 球种类
| 14 球定价(元/个) | 15 球成本 | 16 球费小计 | 17 分组成本 | 18 分组收入 | 19 分组盈亏
| 20 旧球回收个数 | 21 回收单价 | 22 旧球回收收入 | 23 整场总成本 | 24 整场总收入 | 25 整场盈亏
| 26 球成本单价(元) | 27 记录ID
```
其他表末列均为「记录ID」。导入规则：
- **分组识别**：第 4 列（场地类型）非空 = 新分组开始；第 4 列空但 13/11/12 列有值 = 上一分组的追加球种行
- **去重**（追加模式）：文件带记录ID（第 27/8/7/6 列）时**只按 ID 去重**（同日同配置不同 ID 的合法记录不会被误吞）；无 ID（旧文件）退回 启发式（日期+金额+分组数）
- 导入前自动备份（失败则中止）；品牌定价仅在本地无记忆时写入；启动资金：覆盖模式采用文件值，追加模式仅本地为 0 时采用
- 结果弹窗列出跳过明细（日期+金额）

库存表 8 列（+记录ID）、支出表 7 列（+记录ID）、回收表 `[日期,回收个数,回收单价,收入,备注,记录ID]`。旧格式文件（26 列、每行一组一球、无 ID）完全兼容导入。

**给 AI 解析导出数据**：openpyxl/pandas 读 5 个 Sheet，列含义如上；金额单位为元；空字符串表示"沿用上行/不适用"。

## 8. 已知设计决策与约束（历史原因）

- 项目有一份很重的 PRD（`deliverables/product-strategy/`，Vue3+sql.js 方案）——**已被所有者否决**，坚持单文件简单工具，勿按 PRD 技术栈重构
- 不用 `prompt()`；页面内弹窗替代（部分内嵌浏览器不支持）；原生 confirm 仅用于清空二次确认（可接受）
- 记录 id 必须是**数字**（`uid()`），表格按钮用内联 `onclick="editGame(${id})"`
- 金额用浮点 + `r2()`（两位舍入）展示，未做整数分存储
- Windows 上 `python http.server` 会把 .js 发成 text/plain 导致 SW 注册失败，本地测试需自定义 MIME 的服务器
- 测试环境浏览器若点击超时，用 `element.click()`（原生）替代 Playwright click；alert 会阻塞 evaluate，测试时临时 `window.alert=()=>{}`
- 字段语义变迁史：`extraBalls` 曾是"额外收入金额"后为"额外加球个数"；`ballUnit` 曾兼任收费价与成本价，现仅是成本价；v2 起分组单球字段 → balls[] 数组

## 9. 待办 / 可选优化方向（backlog）

- 回档备份导出为 .json 文件（异地容灾）
- 多球种行在移动端的折叠摘要布局
- 球种为空但填了消耗时的行内黄色警示
- 成员维度：收款清单（谁已付/未付）
- doPull 拉到旧版远端（无 recycles 键）时的显式提示

## 10. 维护约定（改代码时必须遵守）

1. **改计费公式**：只改 `calcGroupCore()`（展示层/账单均调用它）
2. **改导出列**：新列只能**追加在末尾**，同步 `doImportWorkbook` 的列索引，保持旧文件可导入（无 ID 列时启发式去重兜底）
3. **改完必须**：`node --check` 语法校验 → `python compress.py` 重新生成 index.min.html → 部署 3 文件（见 §2）
4. 所有用户输入渲染前过 `esc()`；金额计算用 `num()/r2()`；localStorage 读取用 `safeParseArray`
5. **更新本文档**：改动涉及 §4/§5/§6/§7 时同步修改，并在文首更新日期
6. 历史数据不可破坏：`normGroups` 兜底、旧回收字段迁移逻辑、`bm_venues` 迁移都要保留
7. 破坏性操作（拉取/导入/清空/回档）前必须 `makeBackup` 且失败即中止

## 11. 本地开发与测试

- 直接双击 `index.html`（file:// 下 SW 不注册，属正常）
- 起本地服务器测试：注意 §8 的 MIME 问题；图表/CDN 需联网
- 验证清单：多球种保存→库存消耗正确；编辑旧含回收场次→备用资金不变且回收自动迁移；同文件追加导入→按 ID 全跳过；导出再导入往返一致；备份失败→操作被拦截；min 版冒烟
