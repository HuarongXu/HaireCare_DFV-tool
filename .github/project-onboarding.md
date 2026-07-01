# 项目记忆 · DFV Tool Onboarding

> 一次性 onboarding 产物，供后续会话直接复用。描述**项目本身**（非方法论——方法论见 [`AGENTS.md`](../AGENTS.md) 与 [`constitution.md`](../constitution.md)）。
> 事实若与代码漂移，以代码为准，并顺手更新本文件。

## 一句话

**HairCare DFV Tool（Demand Flow Verification）**：一键把每周 SAP Analysis for Office (AO) 的
需求流校验从 15–30 分钟手工操作压到 ~90 秒的桌面自动化工具。纯本地运行，无服务端。

## 技术栈

- **语言**：Python 3.10+（推荐 3.13），Windows 专用（COM / pywinauto / win32）。
- **依赖**（`requirements.txt`）：`pywinauto`（AO ribbon 点击）、`pywin32`（Excel COM / 剪贴板 / win32gui）、
  `pandas`（数据管线）、`openpyxl`（写 Excel 报表）、`python-dotenv`、`databricks-sql-connector`（SKU 描述补全）。
- **存储**：SQLite（`dfv_history.db`，根目录，gitignored）。
- **输出**：`output/`（gitignored）——Excel 报表 + 单文件 HTML Dashboard + CSV。

## 运行 / 构建 / 测试命令

| 目的 | 命令 |
|------|------|
| 一键运行（Windows，推荐） | 双击 `start_dfv.bat`（自建/校验 `.venv` → 装依赖 → 跑 `run.py` → 失败回退 `pipeline.py` → 开 Dashboard） |
| 启动内网编辑网站 | 双击 `start_web.bat` 或 `python dfv_tool/app.py`（Flask+waitress，http://localhost:8000，可在网页直接改 Owner/Action Plan 存库） |
| 装依赖 | `pip install -r requirements.txt` |
| 全流程（含 SAP AO 自动化） | `python dfv_tool/run.py` |
| 仅数据管线（用已有 CSV，无需 SAP） | `python dfv_tool/pipeline.py` |
| 跑测试 | `.venv\Scripts\python.exe dfv_tool\test_first_seen_sort.py`（自定义 assert 测试，非 pytest；~24 个 test 函数）；`... dfv_tool\test_editable_actionplan.py`（网页编辑/回填功能测试） |
| 管理历史周次 | `python dfv_tool/manage_history.py`（交互删除重复周） / `manage_history.bat` |
| 回写 Owner 列 | `python dfv_tool/sync_owners.py` / `sync_owners.bat` |

> 无 pytest / linter / CI 配置——测试是手写 `assert` 脚本，`main()` 里逐个调用。

## 目录 / 模块职责

```
dfv_tool/
  run.py               主入口，7 步 E2E：开 AO 工作簿 → 等 ribbon → Refresh All → 填 Prompts(坐标+剪贴板) → 等 BW → COM 批量导出 CSV → 跑 pipeline
  config.py            全部配置：SAP 系统、BW 查询、过滤条件、列映射、错误类型、Owner 派单规则、HKTW 位置
  pipeline.py          数据管线：load → 过滤 DRP 位置 → classify_issues（分类+派单）→ KPI → 导 Excel → 写 history → 生成 Dashboard
  history.py           SQLite 读写（KPI 摘要 + 逐行错误），DB=根目录 dfv_history.db
  dashboard.py         生成单文件 DFV_Dashboard.html（KPI 卡片 + 可筛选行动表 + 趋势图 + 周选择器）
  databricks_lookup.py 对描述为代码的行，查 Databricks ps_psc_sku_master 补全 SKU 描述
  app.py               Flask 内网编辑网站：GET / 渲染 Dashboard、GET /api/errors、POST /api/errors/<id> 存 owner/action_plan（参数化+校验+审计）
  sync_owners.py       把人工在 Excel 改过的 Owner 列同步回 DB 再重生成 Dashboard
  manage_history.py    交互式列出/删除历史周次
  test_first_seen_sort.py  First Time/Duration 列与排序的对抗性测试
  .env                 Databricks 连接（gitignored，未入库）
```

根目录：`start_dfv.bat`（一键）、`*.bat`（各工具入口）、`DFV_Manual.html`（中英用户手册）、
`DFV_Presentation.html`、SAP `.xlsm` 工作簿与 `.pptx`（gitignored 大二进制）。

## 关键约定 / 铁则（改代码前必读）

- **绝不从 COM 调 SAP XLL 函数**（`SAPGetProperty` / `SAPLogon` / `SAPExecuteCommand`）——会永久破坏 AO ribbon。
  所有 SAP 交互走 pywinauto 点 ribbon 按钮。
- **Prompts 对话框用坐标 + 剪贴板输入**（WPF 控件，UIA 拿不到输入框）；执行 Step 1–4 时**不要动鼠标**。
- **显示缩放必须 100%**（坐标点击依赖）。
- **数据读取要 fallback**：SAP/Excel 导出格式多变（UTF-16 伪 .xls、locale 日期），失败要降级尝试。
- **KPI 目标**：18 个月量差 < 2%；13 周 SKU 差 = 0 错误。
- 主要过滤/派单参数集中在 `config.py`（`REGION=XA`、`CATEGORY=HAIRCARE`、`DRP_LOCATIONS`、
  `OWNER_MAPPING` / `FIXED_OWNERS` / `LOCATION_OWNERS` / `IOL_*` / `HKTW_LOCATIONS`）。

## 安全备注

- `dfv_tool/.env` 含明文 Databricks token，但已被 `.gitignore` 忽略、**未被 git 跟踪**（仅本地 dev 明文）。
  勿把它入库；如需轮换，走密钥流程，不硬编码进源码。
- 涉及 auth / 查询 / 子进程 / 路径 / 机密 / 依赖 / 网络 等安全敏感面时，先走
  [`skills/security/SKILL.md`](../skills/security/SKILL.md)。

## 现状

- `specs/` `plans/` `changes/` 仅有 README 占位——**尚无已确认规范**；DFV 工具属先于本方法论存在的 legacy。
- 约定：新功能走主线（brainstorming → spec → plan → 执行 → 验证）；老代码保持现状，
  仅当任务需要时才外科手术式改动，不为"对齐系统"大规模重写。
