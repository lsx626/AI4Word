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

### V7.1: P0 缺陷全修复（块级写入与编辑的对齐问题）

V7.0 落地了流式写入与块级编辑，但真实 Word 探针暴露出五个 P0 级缺陷：某些块类型
**写入即丢失或损坏**，块与段落的对应关系**从未真正建立**。本轮全部修复并验证。

**修复的五个缺陷（均在真实 Word 上复现并验证）：**

1. **围栏代码块整体丢失**：旧 `write_block` 只处理 paragraph/heading/list，markdown-it
   的 `fence` 是单个 token（`content` 为代码、`info` 为语言），直接被跳过。新增
   `_write_code`：逐行 TypeParagraph + Consolas 等宽字体，`n_paras = 行数`。
2. **有序列表产生双倍空段落且无编号**：`ordered_list_open` 与 bullet 结构同构但无分支。
   新增分支统一样式逻辑（`WD_STYLE_LIST_NUMBER = -50`，真自动编号），每个
   `list_item_open` 累加 `n_paras` 并 TypeParagraph。
3. **列表 `n_paras` 永远为 1 → `rebuild_ranges` 静默放弃**：块对齐逻辑形同虚设。
   现在 `n_paras` 按项目 / 行数精确登记，且 `insert_after` 用块**最后一个**段落定位
   （旧实现取 `Paragraphs(1)`，多段落块会插进块中间）。
4. **`delete_block` 留孤儿空段落**：旧实现只删块 Range（不含末尾段落标记）。现在删除
   块占用的**全部**段落；表格块走 `_delete_table_block`（见下）。
5. **表格变成 pipe 原文**：默认 preset 不含 table 规则（须 `MarkdownIt().enable("table")`），
   且旧实现无 table 分支。新增 `_collect_table_rows` + `_write_table`：写成真 Word 表格
   （Borders.Enable），块 Range 覆盖表格与容器段落，`n_paras = 表格段落数 + 1`；
   表格创建失败时退化为逐行文本，保证内容不丢。

**真实 Word COM 的新坑（本轮实测确认）：**

- **`Range.Delete` 删不掉表格**：对覆盖表格的 Range 调 Delete 只删文字，留下一堆
  单元格标记（`\x07`）的孤儿结构，`Tables.Count` 不变。表格块必须先
  `range.Tables(1).Delete()` 删结构，再删除残留的容器段落。
- **`Content.End` 处无法构造折叠 Range**：`doc.Range(pos, pos)` 在 pos 等于文档末端时
  报"数值超出范围"。`insert_after` 改用段落的 `InsertAfter("\r")` 造空段落（先把文档
  撑长，pos 随即合法）。
- **表格块的 Range 不含容器段落**：`insert_after` 取块内最后一个段落会取到表格的
  行尾段落，在那里 InsertAfter 报"此操作对行结尾无效"。表格块要改用表格之后的
  容器段落作为插入锚点。

**离线测试假阳性的教训：**

旧的 fake Word 没有 `FakeDoc.Paragraphs` 属性，`rebuild_ranges` 里的 `doc.Paragraphs`
抛 AttributeError 被 `except: return` 吞掉——**块对齐逻辑在离线从未执行过**，fake
还编码了实现的错误假设（如"列表只占一个段落"）。现在 fake 带完整段落模型
（`Paragraphs.Count` / `Paragraphs(n).Range` 按区间重叠过滤、折叠区间取包含段落），
并新增**段落数对齐断言**（`sum(n_paras) == doc.Paragraphs.Count`，各编辑操作后检查），
让"块与段落失去对应"在离线直接失败。表格不模拟（Word 表格语义太复杂，半吊子模拟
等于重蹈假阳性），真路径由 `smoke_real_word.py` 覆盖。

**测试覆盖（本轮新增）：**

- `tests/test_fake_word.py`：8 → 13 个：新增有序列表（编号样式 -50、无双倍空段落）、
  围栏代码块跨分片（空行不断切）、删除列表块无孤儿、列表块后插入定位、表格降级路径、
  段落数对齐断言。
- `tests/smoke_real_word.py`：新增有序列表、围栏代码块（Consolas 断言）、表格
  （行数 / 列数 / 边框 / 段落计数）、删除列表块与表格块、替换表格块、表格块后插入。
- 真实 Word 全流程实测通过（新建临时文档、结束不保存）。
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
