# AI4Word

> 一个调用 Atria（Intern AI discovery 平台）、通过 pywin32 COM **实时控制本机真实 Word** 的命令行 Agent。运行在终端，无 GUI。
> AI 生成的内容会**以真正的流式方式**一边接收一边带格式写入 Word；写入后还能**编辑已有内容**，排版过程**肉眼可见**。

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

## 项目结构

| 文件 | 说明 |
|------|------|
| `main.py` | 主程序：两阶段编排（流式写作 → 交互式排版循环） |
| `ai_client.py` | Atria 客户端：`ai_stream()`（SSE 流式）与 `ai_request()`（非流式，用于代码生成） |
| `streaming_writer.py` | 流式 Markdown 写入器：块级缓冲 + 打字机动画 + 块登记（标题 / 列表 / 代码块 / 表格） |
| `doc_model.py` | 块级文档模型与编辑原语（`replace_block` / `insert_after` / `delete_block` 等） |
| `doc_watch.py` | 文档看门狗：主线程轮询，检测用户在生成停顿期间的手动编辑 |
| `format_runner.py` | 逐语句可视化执行器：终端打印 + 选区闪烁 |
| `doc_watch.py` | 文档看门狗：主线程轮询，检测用户在生成停顿期间的手动编辑 |
| `sandbox.py` | AI 生成代码的 AST 静态检查与循环步数护栏 |
| `session.py` | 会话记忆：交互历史 / 用户偏好，拼进代码生成提示词 |
| `styles.py` | 样式预设：论文 / 公文 / 简历 / 博客成套排版 |
| `tests/` | 测试：`test_fake_word.py`（离线 fake Word）、`test_format_runner.py`（离线）、`test_session.py`（离线会话记忆）、`smoke_real_word.py`（真实 Word 端到端）、`e2e_real_atria.py`（真实 Atria API + 真实 Word 全流程） |

## 前置条件

- Windows 操作系统（需安装 Microsoft Word）。
- Python 3.8+（推荐 conda 或虚拟环境）。
- 在项目根目录准备 `.env`（模板见 `.env.example`）：

```env
ATRIA_API_KEY=你的_api_key

# 可选：覆盖默认模型名（默认 atria）
# ATRIA_MODEL=atria
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
python tests/smoke_real_word.py\n`\n\n真实 Word 冒烟测试（会打开一个临时文档，结束关闭不保存）：\n\n`powershell
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
