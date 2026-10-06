# AI4Word

> 一个调用 Atria（Intern AI discovery 平台）、通过 pywin32 COM **实时控制本机真实 Word** 的 AI Agent，提供**命令行**与**桌面悬浮窗 GUI** 两种形态。
> AI 生成的内容会**以真正的流式方式**一边接收一边带格式写入 Word；写入后还能**编辑已有内容**，排版过程**肉眼可见**。
> GUI 形态（`python ai4word.pyw`）是类桌宠 / 桌面插件的悬浮窗：托盘常驻、可开机自启、安装包一键安装，适合非技术用户日常使用。

## 功能概述

- **流式输出**：SSE 逐块接收生成内容，解析 Markdown 后以打字机动画写入 Word——AI 边想边写，Word 边出字。
- **块级编辑**：每个块登记为动态 Word Range，改写 / 插入 / 删除 / 全文替换都按块索引精确改稿，不动文档其余部分。
- **可视化排版**：AI 生成的排版代码逐语句执行，终端实时显示语句与结果，受影响区域在 Word 里滚动闪烁——排版不是黑盒。
- **完整 Markdown 支持**：1-6 级标题、有序 / 无序列表（嵌套缩进）、引用块、围栏代码块、表格（真 Word 表格）、行内代码、超链接、水平线；每种块占用的段落数被精确登记，删除 / 改写不留孤儿段落。
- **段落草稿态**：段落第一个字符到达就落盘（灰色实时打字），块边界确定后提交转正回填颜色——不会整段缓冲。
- **声明式编辑与撤销**：`apply_edit` / `preview_edit`（先预览差异再执行）、`undo` / `model_undo`、多步事务 `begin_txn`/`commit`/`rollback`、修订模式把 AI 编辑变成 Word 修订，块模型随文档保存、重开即恢复。
- **样式预设与打字档位**：`apply_preset("论文"|"公文"|"简历"|"博客")` 一键成套排版；打字档位在逐字精确与批量流畅间切换。
- **会话记忆与安全沙箱**：跨轮次记住排版偏好与成败历史；AI 生成的代码经 AST 静态检查 + 循环步数护栏（禁 import / while / 危险内建）后才执行。
- **长文档支持**：内联图片（跨流式分片不丢图）、自动目录、页眉 / 页脚、分页符。
- **桌面悬浮窗 GUI**：无边框圆角置顶窗（Win11 走 DWM 原生圆角），紧凑胶囊一点展开为完整面板；接入已有文档时自动读取其内容；生成中移动 Word 光标不影响写入（写入锚点防护）；托盘常驻、可开机自启；所有 COM 调用在独立 QThread，UI 不卡、可随时中断并回滚 / 保留。

## 项目结构

| 文件 | 说明 |
|------|------|
| `main.py` | 主程序（CLI）：两阶段编排（流式写作 → 交互式排版循环） |
| `ai4word.pyw` | GUI 启动入口（开发态与打包后共用；未捕获异常写 crash.log） |
| `app/` | 悬浮窗 GUI（PySide6）：窗口、头像、消息流、块地图、QThread 引擎、托盘、设置 |
| `build/` | 打包：`build.py`（PyInstaller + Inno Setup 一键）、`icon_gen.py`、`ai4word.iss` |
| `ai_client.py` | Atria 客户端：`ai_stream()`（SSE 流式）与 `ai_request()`（非流式） |
| `streaming_writer.py` | 流式 Markdown 写入器：块级缓冲 + 打字机动画 + 块登记 |
| `doc_model.py` | 块级文档模型与编辑原语（`replace_block` / `insert_after` / `delete_block`…） |
| `format_runner.py` | 逐语句可视化执行器：终端打印 + 选区闪烁 |
| `doc_watch.py` | 文档看门狗：主线程轮询，检测用户手动编辑与块索引漂移 |
| `sandbox.py` | AI 生成代码的 AST 静态检查与循环步数护栏 |
| `session.py` | 会话记忆：交互历史 / 用户偏好，拼进代码生成提示词 |
| `styles.py` | 样式预设：论文 / 公文 / 简历 / 博客 |
| `tests/` | 测试：`run_offline.py` 聚合全部离线单测；`deep_test.py` / `deep_parallel.py` 为全真机驱动深度测试（见 `DEEP_TEST_WORKFLOW.md`）|

## 前置条件

- Windows 操作系统（需安装 Microsoft Word）。
- Python 3.8+（推荐 conda 或虚拟环境）。
- 在项目根目录准备 `.env`（模板见 `.env.example`）：

```env
ATRIA_API_KEY=你的_api_key

# 可选：覆盖默认模型名（默认 Atria-Dawn-Preview）
# ATRIA_MODEL=Atria-Dawn-Preview
```

## 安装

1. 创建并激活环境（示例用 conda 在项目目录下创建）：

```powershell
conda create -p .venv python=3.11
conda activate ./.venv
```

2. 安装依赖：

```powershell
pip install -r requirements.txt
```

## 桌面悬浮窗 GUI

```powershell
python ai4word.pyw        # 或 python -m app
```

- **双形态悬浮窗**：无边框圆角不透明、可置顶；贴边磁性吸附。紧凑形态是「发光头像 + 输入框」，回车即发送指令（如「写一篇关于秋天的散文」）；点击展开为完整面板——消息流、实时进度条、工具条（打字档位 / 修订模式 / 样式预设 / 存档）、块地图侧栏（点击块定位到 Word 对应位置）。
- **托盘与开机自启**：关窗即最小化到托盘（右键菜单：显示 / 隐藏 / 开机自启 / 设置 / 退出）；设置面板随时切换自启。
- **首次使用**：启动后若未检测到密钥会自动弹出设置对话框，贴入 Atria API 密钥即可；设置面板也可改服务地址、模型名、窗口置顶。
- **接入已有文档**：连接 Word 时若活动文档没有 AI4Word 存档，现有内容会被自动读进块模型（块地图可见、AI 可按块索引编辑已有文字）；AI 写入的新内容会接在文档末尾。
- **不卡 UI、可中断**：所有 Word COM 调用都在独立 QThread 内完成（单实例守护，重复打开会提示「已经在运行了」）；生成过程中可随时中断，并选择「回滚」或「保留」已写入的内容。

## 打包与安装包

一键完成「图标生成 → PyInstaller 打包 → 依赖探针补漏 → Inno Setup 编译安装包」：

依赖探针（`fix_missing_dlls`）在打包后用 Windows 加载器实测启动所需模块，把 PyInstaller 漏收的 PATH 依赖 DLL（如 conda 系 venv 的 `ffi.dll` / `libssl-3-x64.dll`）显式补进 `_internal`；同时构建 PATH 剔除外来 `icu*.dll`，避免精简版 ICU 串扰 Qt6Core。

```powershell
.\.venv\Scripts\python.exe -u build\build.py
```

产物：

- `dist\AI4Word\AI4Word.exe`：解压即用的程序目录（约 70MB，已裁剪 QML / PDF / 软件渲染等无用组件）。
- `dist\AI4Word-Setup-<version>.exe`：安装包（约 23MB，LZMA2 压缩，简体中文向导）——开始菜单组、桌面快捷方式（可选）、开机自启任务（可选）、卸载时清理自启注册项。

要求：Windows 10 1903+ / Windows 11 x64 且已安装 Microsoft Word（程序不捆绑外来 ICU，运行时解析系统自带 ICU）；打包机需 Inno Setup 6（未检测到则跳过安装包步骤、仅输出 `.iss` 供自行编译）与 `requirements-dev.txt` 中的 PyInstaller / Pillow。图标由 `build/icon_gen.py` 用 QPainter + PIL 现场生成，无外部图片资源依赖。

## 使用方法

```powershell
python main.py
```

交互流程：

1. 程序连接（或启动）Word，新建 / 复用活动文档。
2. 输入写作需求 → AI 流式生成，内容带格式逐字写入 Word，终端同步回显原文。
3. 写入完成后进入「AI 智能排版模式」：
   - 编辑内容用块索引，例如「改写第 2 块：……」「在第 0 块后插入一段……」「删除第 3 块」；
   - 格式指令例如「将所有一级标题居中并改为蓝色」「页面设为横向、行距 1.5」；
   - 每条 AI 生成的语句都会在终端逐条显示并在 Word 里闪烁定位，输入「退出」结束。

## 开发与测试

- 离线测试（不启动 Word、不调 API）：

```powershell
python tests/run_offline.py
```

- 全真机驱动深度测试（真实窗口 / 真实点击 / 真实 Word COM / 真实 Atria，自动「测试 → 分析 → 修复 → 复测」闭环，详见 `DEEP_TEST_WORKFLOW.md`）：

```powershell
.\.venv\python.exe -u tests\deep_parallel.py          # 3 槽并行全量
.\.venv\python.exe -u tests\deep_parallel.py --only interrupt_r   # 定向重跑
```

- 真实 Atria API + 真实 Word 端到端（连接已运行的 Word、**新建空文档**，需要 `.env` 中的 `ATRIA_API_KEY`）：

```powershell
python tests/e2e_real_atria.py
```

- 真实 Word 冒烟测试（会打开一个临时文档，结束关闭不保存）：

```powershell
python tests/smoke_real_word.py
```

- GUI 引擎 × 真实 Word（后台线程内 Dispatch，主线程独立连接校验，需提权运行 Word）：

```powershell
python tests/smoke_engine_real_word.py
```

## 常见问题与排查

- 报错 `AttributeError: Property '<unknown>.Style' can not be set.`
  - 说明：不能直接对 `Selection.Style` 赋值。本项目通过 `selection.Range.Style = selection.Document.Styles(...)` 设置样式。
  - 解决：使用本仓库最新版本，并在 Word 的信任中心允许外部程序访问。

- 程序无法连接到 Word：确认 Word 已安装且可通过 COM 被脚本控制；检查 UAC 设置。

- `ai_stream` 报网络错误：检查网络与 `ATRIA_API_KEY` 是否有效。

## 开发者说明

- 架构要点（详见 `DEVELOPMENT_LOG.md`）：
  - **块级缓冲**：inline 格式（粗体/斜体）必须等一个块解析完才能确定结构，因此按空行 / 标题行切块，块完整后立即解析写入；块内延迟极小，体感仍是边收边写。
  - **动态 Range**：Word 的 Range 会随文档变动自动平移，每块登记 Range 后「改写第 N 块」即可精确定位；但「恰好在块边界处插入」时 Range 会扩展而非平移，因此每次编辑后按段落重新对齐（`rebuild_ranges`）。
  - **双轨编辑**：内容编辑优先走块原语（结构化、可定位），格式修改通过改样式定义（Style），同时保留 AI 写任意 pywin32 代码的能力。

## 许可证

未指定特定许可证。
