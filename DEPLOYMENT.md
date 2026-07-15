# DFV 部署到另一台电脑 — 打包 & 上手指南

> 结论先行：**只拷贝代码不能直接跑。** 代码本身是可移植的（所有路径都基于
> `__file__` 相对定位，无硬编码盘符），但有几类**不进 git 的运行依赖**必须单独带过去或
> 在目标机上重建。照下面清单走一遍即可。

---

## 1. 打包方式（二选一）

| 方式 | 做法 | 备注 |
|------|------|------|
| Git（推荐） | 在目标机 `git clone` 本仓库 | 干净、可追更新。gitignored 的依赖仍需按第 3 步补齐 |
| 直接拷贝目录 | 复制整个项目文件夹 | **务必排除 `.venv/`**（虚拟环境含绝对路径，跨机不可用，必须重建） |

不要拷贝 `.venv/`、`__pycache__/`、`output/`——这些在目标机重新生成。

---

## 2. 目标机前置条件

- Windows 10/11
- **Python 3.10+**（推荐 3.13）
- **SAP Analysis for Office (AO)** 已安装并可 SSO 登录
- **显示缩放 = 100%**（run.py 靠坐标点击 SAP Prompts 面板，125%/150% 缩放会点偏）
- 运行 run.py 时屏幕须**已登录且未锁屏**（UI 自动化需要真实桌面）

---

## 3. 必须单独带过去 / 在目标机重建的东西（都不在 git 里）

| 项目 | 来源 | 处理方式 |
|------|------|----------|
| **Python 依赖** | `requirements.txt` | 目标机 `pip install -r requirements.txt`（或双击 `start_dfv.bat` 自动建 venv） |
| **AO 工作簿 .xlsm** | 本机根目录（gitignore 忽略 `*.xlsm`） | 手动拷到目标机**仓库根目录**，文件名须与 `config.py` 的 `WORKBOOK_PATH` 一致：`DRP Forecast Flow Validation AO (20251021) - KHP.xlsm` |
| **Databricks 凭据 `.env`** | `dfv_tool/.env`（gitignore 忽略 `.env`） | 复制 `dfv_tool/.env.example` 为 `dfv_tool/.env` 并填真实值（留空则跳过 SKU 描述补全，不影响主流程） |
| **邮件收件人配置** | `dfv_tool/email_config.json`（gitignore 忽略） | 复制 `dfv_tool/email_config.example.json` 为 `email_config.json` 并填收件人 + 可选 `dashboard_url` |
| **历史数据库**（可选） | `dfv_history.db`（gitignore 忽略） | 想保留历史周数据就手动拷过去；不拷则目标机首跑自动新建空库 |
| **SAP AO Prompts 变量集(Variant)** | 无文件，存在 SAP 账号侧 | 在目标机首次手动打开 Prompts → 填 Region=XA + Category=HAIRCARE → Save Variant，命名 `HAIRCARE_XA`（见 `run.py` 顶部说明） |

---

## 4. 目标机上手步骤

```bat
:: 1) 取代码
git clone <repo-url>
cd 6.DFV

:: 2) 装依赖（或直接双击 start_dfv.bat，它会自动建 .venv 并装依赖）
pip install -r requirements.txt

:: 3) 按第 3 步补齐 .xlsm / .env / email_config.json

:: 4) 跑主流程
python dfv_tool/run.py

:: 5) 起网页看板（在线编辑 Owner/Action Plan + 生成周报邮件）
python dfv_tool/app.py   ::  → http://localhost:8060  /  操作手册在 /manual
```

---

## 5. 验证部署是否 OK

```bat
:: 单元测试（不碰 SAP / Outlook，纯逻辑，能跑说明代码环境正常）
.venv\Scripts\python.exe dfv_tool\test_weekly_email.py
.venv\Scripts\python.exe dfv_tool\test_editable_actionplan.py
.venv\Scripts\python.exe dfv_tool\test_first_seen_sort.py

:: 网页冒烟：起 app.py 后浏览器打开 http://localhost:8060 与 /manual
```

三个测试全绿 + 网页能打开 = 代码环境就绪。之后 SAP 部分是否成功，取决于第 2/3 步的 AO/工作簿/Variant 是否到位。

---

## 6. 已知坑（换机器最容易踩）

1. **显示缩放必须 100%**——否则 SAP Prompts 面板输入坐标算偏，Region/Category 填不进去。
2. **锁屏 = 失败**——run.py 用物理鼠标点击和窗口切前台，锁屏或无人值守时跑不了（所以周二定时任务已取消）。
3. **`.venv` 不能拷**——虚拟环境脚本里写死了源机绝对路径，跨机必须重建。
4. **绝不从 Python COM 调 SAP XLL 函数**——会永久破坏 AO ribbon（`run.py` 全程用 pywinauto 点按钮）。
5. **工作簿文件名要对齐 `config.py`**——改了文件名就同步改 `WORKBOOK_PATH`。
