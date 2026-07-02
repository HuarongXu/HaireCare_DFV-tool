# 一键生成周报邮件（Outlook 草稿）规范（Spec）

> 由 `writing-spec` / `brainstorming` 技能产出。这是该功能「系统当前应有样子」的单一真相。
> 改本规范走 `changes/` 提案流程，不直接改。

**状态：** 草稿
**日期：** 2026-07-02
**关联计划：** plans/2026-07-02-weekly-report-email.md（待写）

## 1. 目标与背景

每周 DFV report 出来后，报告负责人要手动写一封"GC HC DFV Weekly Result"邮件：固定分发列表、开头一句结果总结、正文一张与 Dashboard 完全一致的明细表，逐条列出待关闭项。手工整理耗时且易错。

本功能在现有 Flask Web 应用上加一个按钮：**一键根据当前选中周的数据生成一封 Outlook 草稿邮件**（含收件人、主题、开头总结、Executive Summary 洞察块、明细表格），负责人在 Outlook 里核对/微调后手动发送。

痛点：现在每周要人工从 Dashboard 复制表格、手写总结数字、逐个填收件人。

## 2. 范围

- **做什么：**
  - Dashboard 顶部（"Copy Table" 旁）新增按钮 **📧 生成周报邮件**。
  - 新增后端端点 `POST /api/email/weekly`，接收周标识（run_id 或周索引），构建邮件并通过 Outlook COM 弹出**草稿**（`.Display()`）。
  - 邮件正文（HTML，内联样式，Outlook 兼容）包含：
    - 主题：`GC HC DFV Weekly Result-YYYYMMDD`（YYYYMMDD = 该周 run 日期）。
    - 收件人 To/CC：从本地配置文件读入（参考历史邮件的固定分发列表），Outlook 里仍可改。
    - 开头：`Dear all` + 自动生成的总结句（本周 result X%、还有 N 项待关闭、其中 M 项超 4 周）。
    - Executive Summary 洞察块（表格上方）：① 环比上周、② Aging 分层、③ 按 Owner 汇总（owner 名字加粗）。
    - 明细表格：当前选中周的**全部 Action Items**（与 Dashboard 表格同源同列，含 Owner、Action Plan）。
  - 收件人配置：真实收件人存 `dfv_tool/email_config.json`（**gitignore，不入库**）；仓库仅提交 `dfv_tool/email_config.example.json` 模板。
  - 记录一条日志（触发时间 + 触发者 IP + 生成的 run 标识），便于审计。

- **明确不做什么（YAGNI 边界）：**
  - **绝不自动发送**：只调 `.Display()`，永不调 `.Send()`。
  - 不做网页内邮件预览/富文本编辑；一切核对与微调在 Outlook 里完成。
  - 不做真正的 Outlook @mention（owner 用加粗名字代替）。
  - 不做 SMTP 直发、不做 .eml 下载（本次定为 Outlook COM，服务运行在发信人机器上）。
  - 不做登录/鉴权（延续现有"内网可信"模型）。
  - 不改动现有 Dashboard 数据、pipeline、编辑功能的行为。

## 3. 需求

可验证条目：

- **R1**（按钮）：Dashboard 顶部渲染一个"生成周报邮件"按钮；点击时以当前选中周的 `run_id` 调用 `POST /api/email/weekly`。
- **R2**（端点）：`POST /api/email/weekly`，body `{"run_id": <int>}`。命中已存在的 run 时构建并弹出 Outlook 草稿，返回 200 + JSON `{"ok": true}`；run 不存在返回 404；缺 `run_id` 或类型非法返回 400。
- **R3**（主题）：邮件主题为 `GC HC DFV Weekly Result-YYYYMMDD`，日期取该 run 的 `run_date`（格式 `YYYYMMDD`）。
- **R4**（收件人）：To/CC 从 `email_config.json` 读取（字段 `to`、`cc`，均为字符串数组）。配置缺失或无法解析时，收件人留空但仍能生成草稿（不因此报错），并记一条 warning 日志。
- **R5**（开头总结）：正文开头为 `Dear all` + 一段自动总结，含：本周 result（`diff_pct`，保留 2 位小数）、待关闭项数（当前周 Action Items 行数）、其中 `duration >= 4`（周）的项数。
- **R6**（Executive Summary — 环比上周）：若存在上一周 run，展示 result % 与待办数相对上周的增减（含方向，如 `↓0.12pp`、`-5`）；无上一周则省略该行。
- **R7**（Executive Summary — Aging 分层）：展示 `duration >= 4` 周与 `>= 8` 周的项数。
- **R8**（Executive Summary — 按 Owner 汇总）：按 owner 分组统计待办项数，owner 名字**加粗**；空 owner 归为"(未分配)"。
- **R9**（明细表格）：正文表格的行 = 当前选中周的全部 Action Items（与 `history.get_all_data()` 该周 `errors` 一致，即 `is_hktw = 0`）；列顺序为 Product、Description、Brand、Location、Error、Forecast、Reason、Action、First Time、Duration、Priority、Owner、Action Plan。
- **R10**（草稿而非发送）：通过 Outlook COM 生成的邮件用 `.Display()` 打开为草稿；实现中**不得**出现 `.Send()`。
- **R11**（构建/发送分离）：邮件 HTML 由纯函数 `build_weekly_email(run, prev_run=None)`（不依赖 COM/Outlook）生成，可独立单元测试；COM 交互封装在单独函数 `open_outlook_draft(subject, to, cc, html)`。
- **R12**（HTML 安全）：正文中所有用户可控字段（owner、action_plan、action、reason、description、error_message、brand、location、product 等）在拼入 HTML 前做 HTML 转义，防止注入。
- **R13**（COM 失败降级）：Outlook 未安装/COM 调用失败时，端点返回 5xx + 明确错误消息，网页给出友好提示，进程不崩溃。
- **R14**（审计日志）：每次成功触发记录一条结构化日志，含触发时间、触发者 IP、生成的 run 标识；日志中不含收件人邮箱等 PII。

## 4. 数据与接口

- **输入数据**：复用 `history.get_all_data()` 返回的 runs（每个含 `id, run_date, week_label, diff_pct, total_errors, actionable_errors, vol_status, errors[...]`）。每条 error 含 `apo_product, description, brand, apo_location, error_message, idp_forecast, reason, action, action_plan, owner, first_time, duration, priority`。
- **配置文件** `dfv_tool/email_config.json`（gitignored）：
  ```json
  { "to": ["Name <addr@pg.com>", "..."], "cc": ["Name <addr@pg.com>"] }
  ```
  仓库提交 `dfv_tool/email_config.example.json` 作模板（含占位邮箱，无真实 PII）。
- **端点**：`POST /api/email/weekly`，JSON body `{"run_id": <int>}`。
- **响应**：成功 `200 {"ok": true, "subject": "..."}`；参数错误 `400`；run 不存在 `404`；COM/系统错误 `500 {"ok": false, "error": "..."}`。

## 5. 安全与合规

- **只 Display 不 Send**：绝不自动外发邮件（R10）。
- **PII**：真实收件人邮箱只存本地 gitignored 配置，不入库、不写日志（R4/R14）。
- **HTML 注入**：正文所有用户可控字段转义（R12）。
- **信任模型**：延续现有内网可信、无鉴权；草稿只在运行服务的机器（= 发信人机器）弹出。触发者 IP 记入日志留痕。
- **COM 稳定性**：COM 失败不影响 Web 应用其余功能（R13）。

## 6. 验收

- 单元测试覆盖 `build_weekly_email`：主题格式、开头总结数字、环比、Aging 分层、按 Owner 汇总、表格行列、HTML 转义、无 `.Send()`。
- 端点测试：正常 run 返回 200（COM 层可 monkeypatch 掉）、坏 run_id 404、缺参 400、COM 抛错时 500。
- 手动验收：真实数据点按钮 → Outlook 弹出草稿，收件人/主题/正文正确，核对后可手动发送。

## 7. 未决 / 后续（不阻塞本次）

- 若将来需要在另一台机器跑服务、由浏览器所在机器发信，可加"下载 .eml"通道（本次不做）。
- 若需要真正的 Outlook @mention，可后续单开变更提案。
