# AI4Word

> 一个调用 Atria（Intern AI discovery 平台）、通过 pywin32 COM **实时控制本机真实 Word** 的 AI Agent，提供**命令行**与**桌面悬浮窗 GUI** 两种形态。
> AI 生成的内容会**以真正的流式方式**一边接收一边带格式写入 Word；写入后还能**编辑已有内容**，排版过程**肉眼可见**。
> GUI 形态（`python ai4word.pyw`）是类桌宠 / 桌面插件的悬浮窗：托盘常驻、可开机自启、安装包一键安装，适合非技术用户日常使用。

## 功能概述

- **真正的流式输出**：SSE 逐块接收 Atria 的生成内容，按块解析 Markdown 后以打字机动画逐字写入活动的 Word 文档——AI 边想边写，Word 边出字。
- **编辑已写入的内容**：每个写入的块都会登记成一个动态 Word Range，提供块级编辑原语（改写 / 插入 / 删除 / 全文替换），AI 通过块索引精确改稿，不动文档其余部分。
- **可视化的排版过程**：AI 生成的排版代码被**逐语句**执行——终端实时显示每条语句与执行结果，Word 里受影响的区域会滚动进视野并黄色闪烁，排版不再是黑盒。
- **完整的 Markdown 块支持**：段落、1-6 级标题、有序 / 无序列表（真自动编号与项目符号、**支持嵌套层级缩进**）、引用块（缩进 + 斜体）、水平线（段落底边框）、围栏代码块（等宽字体、块内空行不会切断流式分块）、行内代码、超链接（真 `Hyperlinks`）、GFM 表格（写成**真正的 Word 表格**，含边框与单元格）。每种块占用的段落数被精确登记，删除 / 改写不留孤儿段落。
- **段落级草稿态打字**：段落文本第一个字符到达就开始往 Word 里写（灰色实时打字），块边界确定后才"提交"转正并回填颜色——不再整段缓冲，AI 想一个字、Word 出一个字。流式分片切断标记前缀（`#`、`-`、`` ` ``）时自动缓冲等定性，不会把结构标记当正文打出。
- **声明式编辑与撤销**：`apply_edit(spec)` / `preview_edit(spec)`（先预览前后对比再执行）、`undo(times)` 按编辑记录精确回退、`model_undo`/`model_redo` 按块模型快照整体撤销、`begin_txn`/`commit`/`rollback` 保证多步编辑原子化；**修订模式**把 AI 的编辑变成 Word 修订（接受/拒绝由用户拍板）；**块模型序列化进文档变量（`doc.Variables`）随文档保存**，重开同一文档即恢复块索引；**文档看门狗**在生成停顿期间检测用户的手动编辑并提示块索引漂移。
- **长文档布局与多媒体**：内联图片（草稿期跨分片到达、提交时替换，不丢图）、自动目录、页眉 / 页脚、分页符。
- **样式预设与打字档位**：`apply_preset("论文"|"公文"|"简历"|"博客")` 一套成套排版一次落地；打字档位在逐字精确与批量流畅间切换。
- **会话记忆与安全执行**：跨轮次记住用户排版偏好与成败历史（AI 也能主动 `remember`）；AI 生成的排版代码经 **AST 沙箱静态检查 + 循环步数护栏**（禁 import / while / 双下划线 / 危险内建）后才执行。
- 交互式排版：支持页边距、行间距、字体字号、纸张方向等任意排版需求；执行失败时 AI 会根据错误信息自我修复并重试。
- **桌面悬浮窗 GUI（V9.1）**：PySide6 无边框圆角置顶窗（不透明渲染；Win11 走 DWM 原生抗锯齿圆角，杜绝输入时背景变透明），紧凑胶囊（发光头像 + 输入框）一键展开为完整面板（消息流 / 工具条 / 块地图侧栏 / 实时进度状态条）；**接入已有文档时自动读取其内容**（块索引即刻作用于现有文字）；**生成期间即使挪动 Word 光标，写入也始终接着已生成内容**（写入锚点防护）；贴边吸附支持拖动中磁吸与松手吸附；系统托盘常驻、开机自启、设置面板内填密钥即用；所有 Word COM 调用在独立 QThread 完成，UI 不卡顿，生成中可中断并选择回滚 / 保留已写入内容；写作阶段只往文档里输出正文本身（无前言、无讨论）。

## 项目结构

| 文件 | 说明 |
|------|------|
| `main.py` | 主程序（CLI 形态）：两阶段编排（流式写作 → 交互式排版循环） |
| `ai4word.pyw` | GUI 启动入口（开发态与打包后共用；窗口化打包下未捕获异常统一写 crash.log） |
| `app/` | 桌面悬浮窗 GUI（PySide6）：双形态悬浮窗、桌宠头像、消息流 + 块地图、QThread 后台引擎、托盘、设置 |
| `build/` | 打包脚本：`build.py`（PyInstaller + Inno Setup 一键打包）、`icon_gen.py`（图标生成）、`ai4word.iss`（安装包脚本） |
| `ai_client.py` | Atria 客户端：`ai_stream()`（SSE 流式）与 `ai_request()`（非流式，用于代码生成） |
| `streaming_writer.py` | 流式 Markdown 写入器：块级缓冲 + 打字机动画 + 块登记（标题 / 列表 / 代码块 / 表格） |
| `doc_model.py` | 块级文档模型与编辑原语（`replace_block` / `insert_after` / `delete_block` 等） |
| `doc_watch.py` | 文档看门狗：主线程轮询，检测用户在生成停顿期间的手动编辑 |
| `format_runner.py` | 逐语句可视化执行器：终端打印 + 选区闪烁 |
| `doc_watch.py` | 文档看门狗：主线程轮询，检测用户在生成停顿期间的手动编辑 |
| `sandbox.py` | AI 生成代码的 AST 静态检查与循环步数护栏 |
| `session.py` | 会话记忆：交互历史 / 用户偏好，拼进代码生成提示词 |
| `styles.py` | 样式预设：论文 / 公文 / 简历 / 博客成套排版 |
| `tests/` | 测试：`test_fake_word.py`（离线 fake Word，含已有文档导入）、`test_format_runner.py`（离线）、`test_session.py`（离线会话记忆）、`test_settings.py`（离线设置）、`test_engine_offline.py`（GUI 引擎离线）、`test_ui.py`（GUI 外观与行为回归：气泡全高渲染、不透明窗口、贴边吸附、预设下拉）、`test_writer_anchor.py`（流式写入光标漂移防护，离线）、`anchor_real_word.py`（光标漂移防护，真实 Word）、`import_real_word.py`（已有文档读取，真实 Word）、`smoke_real_word.py`（真实 Word 端到端）、`smoke_engine_real_word.py`（GUI 引擎 × 真实 Word）、`e2e_real_atria.py`（真实 Atria API + 真实 Word 全流程） |

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

- **双形态悬浮窗**：无边框圆角不透明、可置顶；拖动靠近屏幕边缘磁性吸附，松手再吸附一次。紧凑形态是「发光桌宠头像 + 输入框」，回车即发送指令（如「写一篇关于秋天的散文」）；点击展开为完整面板——消息流（流式增量显示 AI 回复，多行消息完整渲染不截断）、实时进度状态条（已接收字数 / 执行阶段）、工具条（打字档位 / 修订模式 / 样式预设 / 一键存档）、块地图侧栏（点击块索引定位 Word 中对应位置）。
- **托盘与开机自启**：关窗即最小化到系统托盘（右键菜单：显示 / 隐藏 / 开机自启 / 设置 / 退出）；安装包提供「开机自动启动」选项，之后在设置面板里随时切换。
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

- `dist\AI4Word\AI4Word.exe`：解压即用的程序目录（约 70MB，已裁剪 QML / PDF / 软件渲染 / 多余翻译与插件等无用组件）。
- `dist\AI4Word-Setup-9.0.exe`：安装包（约 23MB，LZMA2 压缩，简体中文安装向导，打包机缺中文 isl 时自动降级英文）——开始菜单组、桌面快捷方式（可选）、开机自启任务（可选）、卸载时清理自启注册项。

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

- 离线测试（不启动 Word、不调 API，直接跑测试函数）：

```powershell
python tests/run_offline.py
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
