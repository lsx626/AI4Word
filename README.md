# AI4Word

> 一个调用 Atria（Intern AI discovery 平台）、通过 pywin32 COM **实时控制本机真实 Word** 的命令行 Agent。运行在终端，无 GUI。
> AI 生成的内容会**以真正的流式方式**一边接收一边带格式写入 Word；写入后还能**编辑已有内容**，排版过程**肉眼可见**。

## 功能概述

- **真正的流式输出**：SSE 逐块接收 Atria 的生成内容，按块解析 Markdown 后以打字机动画逐字写入活动的 Word 文档——AI 边想边写，Word 边出字。
- **编辑已写入的内容**：每个写入的块都会登记成一个动态 Word Range，提供块级编辑原语（改写 / 插入 / 删除 / 全文替换），AI 通过块索引精确改稿，不动文档其余部分。
- **可视化的排版过程**：AI 生成的排版代码被**逐语句**执行——终端实时显示每条语句与执行结果，Word 里受影响的区域会滚动进视野并黄色闪烁，排版不再是黑盒。
- 交互式排版：支持页边距、行间距、字体字号、纸张方向等任意排版需求；执行失败时 AI 会根据错误信息自我修复并重试。

## 项目结构

| 文件 | 说明 |
|------|------|
| `main.py` | 主程序：两阶段编排（流式写作 → 交互式排版循环） |
| `ai_client.py` | Atria 客户端：`ai_stream()`（SSE 流式）与 `ai_request()`（非流式，用于代码生成） |
| `streaming_writer.py` | 流式 Markdown 写入器：块级缓冲 + 打字机动画 + 块登记 |
| `doc_model.py` | 块级文档模型与编辑原语（`replace_block` / `insert_after` / `delete_block` 等） |
| `format_runner.py` | 逐语句可视化执行器：终端打印 + 选区闪烁 |
| `tests/` | 测试：`test_fake_word.py`（离线）、`test_format_runner.py`（离线）、`smoke_real_word.py`（真实 Word 端到端）、`e2e_real_atria.py`（真实 Atria API + 真实 Word 全流程） |

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
