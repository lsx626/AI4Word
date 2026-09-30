# AI4Word - 开发日志

## 项目概述

本项目创建一个Python应用，接收自然语言指令调用AI生成内容，并在Microsoft Word中以动画效果排版，支持交互式修改。

---

## 开发阶段

### V1.0: 基础AI写作

- 使用 `python-docx` 库创建并保存Word文档
- 调用Atria API生成文本内容
- 文本直接保存为 `.docx` 格式

### V2.0: 实时打字效果

- 切换使用 `pywin32` 与已打开的Word应用交互
- 实现流式接收AI返回数据
- 逐字将内容输入到Word（打字机效果）

### V3.0: Markdown格式化

- 引入 `markdown-it-py` 解析Markdown
- 将格式（标题、粗体、斜体、列表等）正确应用到Word
- 放弃逐字输入以保证格式完整性

### V4.0: 格式化动画写入

- 结合V2和V3的方案
- 先完整解析Markdown并设置格式
- 然后以逐字动画方式写入Word

### V5.0: 交互式排版尝试

- 支持对话式排版指令
- JSON指令化方案（后期被证明局限性大）

### V6.0: AI代码执行

- AI生成可执行的 `pywin32` 代码
- 支持在沙箱环境中动态执行
- 新增自我修复循环机制

### V6.1: 问题修复与文档完善

- 修复 `selection.Style` 赋值的 `AttributeError`
- 改用 `selection.Range.Style = selection.Document.Styles(...)` 设置样式
- 新增 `README.md` 和规范化文档

### V7.0: 真正的实时控制（流式 / 可编辑 / 可视化）

目标：让软件真正实现对 Word 的实时控制——真正的流式输出、编辑已写入内容、可视化的排版调整过程。

**架构重构为 4 个模块 + 主程序编排：**

- `ai_client.py`：`ai_stream()` 走 SSE 真流式，逐块 yield 内容片段；`ai_request()` 非流式，供排版阶段生成代码。
- `streaming_writer.py`：块级缓冲流式写入。按空行 / 标题行切分块，块完整后立即用 markdown-it 解析、应用格式、打字机动画写入并登记。修复了旧代码松散/紧凑列表的段落处理 bug（markdown-it 的 list_item 内也会 emit `paragraph_open`，需跳过）。
- `doc_model.py`：块级文档模型与编辑原语。每个写入的块登记为动态 Range，提供 `block_map` / `get_block_text` / `replace_block` / `insert_after` / `insert_at_end` / `delete_block` / `replace_text` / `select_block`，全部注入 AI 生成代码的执行环境。
- `format_runner.py`：用 ast 拆分 AI 代码为单条语句逐条执行，终端打印 `[n/m] 代码 [ok]`；语句改变选区时滚动进视野并黄色闪烁，排版过程可见。

**踩坑记录（真实 Word COM 的关键事实）：**

1. **Range 边界插入会扩张而非平移**：在恰好位于块边界处插入文本时，Word 会把新文字划进相邻块的 Range。因此所有编辑原语在执行后调用 `rebuild_ranges()`，按 `doc.Paragraphs` 扫描重建每块 Range（每块记录占用的段落数 `n_paras`；段落总数与 `sum(n_paras)` 不符时跳过重建，说明用户手动改过文档）。
2. **每个段落都有标记、文档永远以 `\r` 结尾**：测试用的模拟对象初始内容必须是 `"\r"`。
3. **pywin32 动态分发下 `Find.Execute` 的关键字参数不生效**：必须按位置传全部 11 个参数，`replace_text` 中即 `f.Execute(old, False, ..., new, 2)`。
4. **块写入的分块职责**：`write_block` 不负责块之间的 `TypeParagraph`（由流式写入的 `_write_top_level` 或编辑原语自己造空段落）；`replace_block` 删除块文本后写进保留的空段落；`insert_after` 在下一块开头 `InsertBefore("\r")` 造空段落再写。

**系统提示词的两大黄金法则**：内容编辑必须用块原语（通过 `block_map()` 索引定位）；格式修改必须改样式定义（`doc.Styles(...)`）而不是遍历段落。

**测试：**

- `tests/test_fake_word.py`：8 个离线测试，用字符串 + 区间平移模拟动态 Range 语义（不需要真实 Word），覆盖流式写入、标记跨片、replace / insert / delete / 列表。
- `tests/test_format_runner.py`：4 个离线测试，覆盖逐语句执行、错误隔离、语法错误、空代码。
- 	ests/smoke_real_word.py：真实 Word 端到端冒烟（流式写入、块登记、block_map、replace / insert、逐语句执行样式修改、全文替换）。\n- 	ests/e2e_real_atria.py：真实 Atria（Atria-Dawn-Preview，discovery-api.intern-ai.org.cn/v1）+ 真实 Word 全流程：连接已运行的 Word 但强制新建空文档，真实 SSE 28 片段流式写入、AI 生成排版代码逐语句执行（标题居中 + 首行缩进）、AI 生成块编辑代码（replace_block 追加文字）全部实测通过。
- `tests/run_offline.py`：无 pytest 依赖的离线测试跑批器。
- 未测：真实 Atria API（需在 `.env` 提供 `ATRIA_API_KEY`，可在 `.env` 用 `ATRIA_MODEL` 覆盖默认模型名）；`main.py` 的完整交互循环；带真实动画延迟的视觉效果（测试中 `time.sleep` 被 monkeypatch 关闭）。

---

## 核心技术栈

- **AI服务**: Atria（Intern AI discovery 平台，OpenAI 兼容接口，SSE 流式）
- **Word交互**: `pywin32` (COM接口)
- **Markdown解析**: `markdown-it-py`
- **API通信**: `requests`
- **环境管理**: `python-dotenv`

---

## 主要文件

- `main.py` - 主程序：流式写作 + 交互式排版编排
- `ai_client.py` - Atria 客户端（流式 / 非流式）
- `streaming_writer.py` - 流式 Markdown 写入器
- `doc_model.py` - 块级文档模型与编辑原语
- `format_runner.py` - 逐语句可视化执行器
- `tests/` - 离线测试与真实 Word 冒烟测试
- `requirements.txt` - 依赖列表
- `README.md` - 项目说明文档
- `DEVELOPMENT_LOG.md` - 开发文档
