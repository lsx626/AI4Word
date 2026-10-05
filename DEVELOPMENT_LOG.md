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
### V7.2: 草稿态实时打字 + 全块型支持 + 声明式编辑与文档看门狗

V7.1 解决了"块写错 / 对不齐"，但流式体验仍是块级的：一个块要完整到达才开始
往 Word 里写，段落体感上有一段延迟。V7.2 的目标是**让"AI 边想边写、Word 边出字"
在段落级别也成立**，同时补齐块型覆盖与编辑能力。

**1. 草稿态（段落级实时打字）**

段落型文本不再整块缓冲：第一个非结构字符到达即进入草稿态，以灰色（Gray25）
逐字实时打入文档，同时跑一个增量标记状态机（`*/`` 与 bold/italic/code 字体切换），
解析出块边界（空行 / 下一行是结构行）后再"提交"：回填 `Font.ColorIndex` 为自动色、
用完整解析（`_normalize_inline`）校验字体属性、登记为 paragraph 块，**不打第二遍字**。
文本长度与解析不一致时放弃规范化（保留状态机结果），避免错位。

**2. 分片边界的歧义前缀处理（本轮最难调的 bug）**

流式分片会切断任何前缀：`#`、`-`、`` ` `` 单独到达时，无法判断它是结构标记还是
正文。旧逻辑会把 `# 标题一` 的 `#`、围栏 ```` ``` ```` 的前两个反引号按正文打出。
现在：

- `_starts_paragraph` 增加 `_AMBIGUOUS_RE`（`#` / `-` / `>` / `` ` `` / 有序前缀等单
  独或残缺形态）：歧义时先缓冲，不进草稿态也不当结构块；
- `_find_boundary` 只在下一行**确定**是结构行时才切分；
- 草稿态 `_draft_advance` 新增尾部保持逻辑：换行后若是歧义前缀，把换行与前缀整体
  攒住（只攒行内标记字符的旧逻辑漏掉了 `-` 这类列表前缀，曾是"前言一句-"泄漏的根因）；
- `_STRUCT_LINE_RE` / `_HEADING_RE` 支持 `#{1,6}` 全部级别。

**3. 块型补齐：嵌套列表 / 引用 / 水平线 / 超链接 / 行内代码**

- 嵌套列表按层级 `LeftIndent = 21 * (depth - 1)`（`ListLevelNumber` 在 pywin32 下
  OLE 失败，已用探针确认）。
- 引用块：每段一段、LeftIndent 缩进、斜体，支持多段与 `>` 空行分割。
- 水平线：优先段落底边框（`wdBorderBottom = -3`），无边框支持时降级为字符横线。
- 超链接：`Hyperlinks.Add` 真链接；无支持时退化为纯文本（文字不丢）。行内代码
  Consolas 等宽字体，打完恢复原字体。
- 块函数的 md 参数与 `md_to_text` 同步支持上述全部语法。

**4. 声明式编辑 + undo + 文档看门狗**

- `apply_edit(spec)` / `preview_edit(spec)`：`{"op": "replace|insert_after|
  insert_at_end|delete", "index": i, "md": ...}`，非法 spec 抛 ValueError 不执行；
  预览只返回文本级前后对比、不落盘。已注入 `exec_globals` 并写入 CODEGEN_SYSTEM
  （"三大黄金法则"，声明式 spec 先 preview 后 apply）。
- `undo(times=1)`：封装 Word 编辑栈，随后 `rebuild_ranges()` 重建块定位。
- `doc_watch.py`（新模块）：主线程轮询文档状态（段落数 / 选区），**不用后台线程**
  （COM apartment 安全）。写操作期间 busy 屏蔽 + 基线重置，AI 自身写入不误报；
  停顿期间的用户手动编辑入队，`drain_events()` 消费。阶段二循环在空闲点轮询，
  检测到漂移时把警告带给代码生成模型（提示其重新 `block_map()`）。

**5. 真实 Word COM 本轮新踩的坑**

- **`TypeText` / `TypeParagraph` 会把列表段落的直接缩进归一化回 List 样式内置值**
  （List Bullet 为 22 磅）：先设缩进再打字、以及逐段赋值，都会被下一段的 TypeParagraph
  冲掉。**只有整块列表打完后统一按段落重设一次才稳定**（实测后续无关块也不影响）。
  Normal 样式（引用块）无此问题，可以先设后打。
- **`wdBorderBottom` 的枚举值是 -3** 不是 3（WdBorderType 全部为负值）。
- **Word 忙时 COM 调用被拒绝（RPC_E_CALL_REJECTED, -2147418111）**：重分页 /
  界面刷新进行时取 Range 会失败。`doc_model._com_retry` 对该错误重试 6 次、
  逐次退避。
- **Word 规范化超链接地址**：无路径 URL 会补结尾斜杠（`https://example.com/`）。
- **上轮遗留的假阳性**：fake 的 `FakeSelectionFont` 缺少大写 `Name` 属性桥接
  （Python 属性大小写敏感，COM 不敏感），导致行内代码字体在离线从未真正断言到。

**测试覆盖（本轮）：**

- `tests/test_fake_word.py`：13 → 22 个：新增字符属性段模型（`_Seg`：TypeText 记录、
  Range.Font 赋值切分覆盖）、LeftIndent 记录、FakeDoc.Undo、草稿态颜色回填与
  粗体段、引用块、超链接降级 + 行内代码字体、围栏代码块跨分片、草稿被结构行打断、
  看门狗、apply/preview、嵌套列表缩进等。
- `tests/smoke_real_word.py`：新增嵌套列表缩进（0/21/42/21 断言）、引用块（缩进 +
  斜体）、水平线、超链接（`Hyperlinks.Count` 与地址）、apply/preview、undo。
- 真实 Word 全流程实测通过（新建临时文档、结束不保存；冒烟启动时清理上次失败
  遗留的文档，避免新文档与旧选区串台）。

---

### V7.3: 修订模式 + 持久化 + 快照撤销/事务 + 沙箱 + 预设 + 图片 + 长文档布局 + 会话记忆

**1. 修订模式**

- `review_on()` 打开 TrackRevisions，AI 的编辑以 Word 修订形式落地，用户逐条拍板；
  `accept_block_revisions(i)` / `reject_block_revisions(i)` 倒序处理块内修订
  （拒绝插入块即回滚新增内容），随后 `rebuild_ranges()` 重建定位。

**2. 块级快照撤销与事务**

- `model_undo()` / `model_redo()`：整篇快照级撤销/重做，与 Word 编辑栈的
  `undo(times)`（按编辑记录回退）互补。
- `begin_txn()` / `commit_txn()` / `rollback_txn()`：多步编辑原子化，失败整体回滚。

**3. 块模型持久化：CustomDocumentProperties → doc.Variables**

- 原存档走 `CustomDocumentProperties.Add`，真实 Word 16.0 下**所有参数形态恒返回
  E_INVALIDARG**（晚期绑定、无类型信息，gen_py 亦无包装可用）。改用 `doc.Variables`：
  探针实测 `Variables.Add(name, value)` / `Variables(name).Value` / `.Delete()` 全部
  可用、保存重开后仍在；单值容量 ≥50000 字符（200000 失败），分片 `_PROP_CHUNK=8000`
  留余量；同名 Add 抛"Variable 名称已经存在"，故 `_set_doc_variable` 先赋值、
  失败再 Add，必要时删除重建。`load_blocks()` 重开同一文档时按 n_paras 重新对齐
  段落、重建 Range，块索引语义跨会话保留。

**4. 内联图片与"草稿期图片被删"根因**

- 结构块内 `InlineShapes.AddPicture`（markdown-it 把路径中的 `\` 编码为 `%5C`，
  src 需 `unquote` 后再解析）；草稿期跨分片到达的图片在提交时替换草稿文本。
- 被删根因：`_flash(rng)` 清高亮后选区覆盖整个块（含内联图片），随后的
  `TypeParagraph` 会"替换选区"、连删图片。两处修复：`_collapse_selection()` 在
  `_ensure_new_paragraph` 打段前把选区折叠到末尾（`_start_draft` 与
  `_write_top_level` 两路都覆盖）；`_flash` 清高亮后把光标恢复为块尾折叠点。

**5. 长文档布局**

- `insert_toc()` 自动目录、`set_header(text)` / `set_footer(text)`、
  `insert_page_break()` 分页符。

**6. 安全沙箱 sandbox.py（新模块）**

- `run_code` 直接执行 LLM 生成的代码，"逐语句可视化看见"挡不住真正的危险动作。
  exec 前静态检查：禁 import / while / with / 双下划线属性 / 危险内建（open、exec、
  eval、getattr 等）；循环护栏给 for 体头部注入步数计数，超限抛 SandboxError，
  把死循环掐死在可承受范围内。

**7. 样式预设 styles.py（新模块）**

- `apply_preset("论文"|"公文"|"简历"|"博客")`：正文（字体/字号/首行缩进/行距）与
  1-3 级标题成套经修改样式定义一次落地，替代逐属性硬改；`preset_names()` 列清单。

**8. 会话记忆 session.py（新模块）**

- 最近 8 轮交互（指令 + 成败标记）与 10 条用户偏好（成功指令里的排版关键词
  自动升格，AI 亦可 `remember(note)` 主动记录），`memory_prompt()` 拼进 gen_code
  提示词；无记忆时返回空串。

**9. 打字档位**

- `_gear()` 在低速逐字（精确）与高速批量（流畅）间切换，兼顾"看得见"与"写得快"。

**10. COM 重试加固**

- streaming_writer 的 8 处样式赋值、12 处 `TypeParagraph` 与 `TypeText` 全部包进
  `doc_model._com_retry`（Word 重分页/界面刷新时以 RPC_E_CALL_REJECTED 拒绝调用，
  重试 6 次、逐次退避）。
- `_com_retry` 的退避 sleep 在模块导入时捕获真实 `time.sleep`：测试为加速而
  monkeypatch `time.sleep` 时，重试退避仍真实生效——否则 6 次重试零间隔空转必败。

**11. 本轮新踩的坑**

- **僵尸 Word 进程污染 Dispatch**：冒烟崩溃遗留的 Word 进程会被
  `Dispatch("Word.Application")` 复用，其实例状态会持续拒绝 COM 调用
  （表现：`Styles()` 取值被拒、重试 6 次仍败，与业务代码无关）。冒烟后须确认
  无遗留 WINWORD 进程；冒烟启动时也会先清理上次失败遗留的文档。
- V7.2 的 `_com_retry` 只护住 doc_model 两处 Range 调用，流式主路径未覆盖（本轮补）。

### V8.0: 桌面悬浮窗 GUI + 一键安装包（发布级）

目标：把 CLI 形态的 Agent 做成「桌面插件 / 桌宠 / 悬浮窗」——安装包一键
安装、托盘常驻、美观的 GUI，非技术用户也能随开随用。

**1. GUI 层（`app/` 包，PySide6）**

- `main_window.py`：无边框 + 半透明 + 置顶主窗口，紧凑胶囊 <-> 完整面板双形态，
  贴边吸附（snap_to_edge）与展开/收起动画；关窗即最小化到托盘。
- `engine.py`：`AgentWorker(QThread)`——**所有 pywin32 COM 调用都在该线程内**
  （run() 起 CoInitialize、结束 CoUninitialize），GUI 线程只通过命令队列与
  Qt 信号交互，UI 永不卡死；支持中断（`interrupt()`）与「回滚 / 保留」选择，
  QSharedMemory 单实例守护。
- `messages.py`：聊天气泡（流式增量追加）+ 可点击的块地图侧栏。
- `avatar.py`：桌宠头像——发光琥珀球 + 呼吸 + 旋转光环（idle / working 两态），
  全部 QPainter 现绘，任意分辨率清晰。
- `theme.py` / `icons.py`：墨黑 + 琥珀深色主题 QSS（刻意避开紫色主导方案）与
  全套矢量图标，无外部资源依赖。
- `settings.py` / `settings_dialog.py` / `auto_start.py` / `tray.py`：设置持久化
  （`%APPDATA%\AI4Word\settings.json`，原子保存、损坏降级默认值）、首次启动
  引导填密钥、HKCU Run 键自启、托盘菜单。
- `agent.py`：从 main.py 抽出的代码生成提示词、`gen_code` / `fix_code` /
  `build_exec_globals` / `get_word`，CLI 与 GUI 共用同一套生成与执行逻辑。

**2. 核心层最小改动复用**

- `format_runner.run_code(code, exec_globals, sink=print)`：print 改 sink，
  GUI 把每条语句与执行结果送进消息流，排版过程在 GUI 里同样逐条可见。
- `ai_client.set_base_url(url)`：设置面板可改服务地址（空串恢复默认）。
- `main.py` 保持 CLI 入口不变，被 agent.py 接管的函数直接 import 复用
  （`_patch_core.py` 为一次性迁移脚本，迁移完成后删除）。

**3. 线程模型与 COM 的关键事实**

- **COM 代理不可跨线程共享**：worker 线程内 Dispatch 的 Word 对象，主线程
  校验时必须**独立第二次 Dispatch**，否则报 `RPC_E_WRONG_THREAD
  (0x8001010E / -2147417842)`。`tests/smoke_engine_real_word.py` 即按此模式：
  worker 线程写入，主线程另起连接读 `ActiveDocument.Content.Text` 验证。
- **QThread + offscreen 离线测试**：`tests/test_engine_offline.py` patch
  `ai_client.ai_stream` / `ai_request` 与 FakeApp，覆盖 arrange 全流程 /
  自我修复重试 / 流式写入 / 中断回滚四例，不需要真实 Word 与网络。

**4. 打包链路（`build/`）**

- `icon_gen.py`：生成 7 帧 16-256px ICO。**必须先 `QApplication(sys.argv)`
  再绘制 QPixmap，否则进程 0xC0000409 硬崩**（QPixmap 需要 QGuiApplication）。
- `build.py`：PyInstaller `--onedir --windowed`（补 pythoncom / pywintypes /
  win32timezone 等隐藏 import）-> `dist\AI4Word\`（约 161MB）-> ISCC 编译
  `AI4Word-Setup-8.0.exe`（约 45MB，LZMA2 ultra + SolidCompression）。
  iss 模板按本机是否存在 `ChineseSimplified.isl` 决定语言（本机只有
  Default.isl，安装向导为英文，任务描述保留中文）。
- 入口用 `ai4word.pyw` 而非 `app/__main__`：窗口化打包没有控制台，未捕获
  异常统一写 `%APPDATA%\AI4Word\crash.log` 便于排障。

**5. 本轮新踩的坑**

- **exe 启动 `ImportError: DLL load failed while importing QtCore: 找不到
  指定的程序`**：根因是 PATH 中 `D:\ProgramData\anaconda3\Library\bin`
  的旧 Qt6Core.dll 与打包的 PySide6 版本不匹配串扰。`_strip_alien_qt_dirs()`
  在 frozen 环境下把 `_internal\PySide6` 注册为 DLL 目录
  （`os.add_dll_directory`）并从 PATH 剔除含 `Qt6*.dll` 的目录，修复后打包
  exe 实测存活（启动 10s 无退出）。
- **启动只见托盘不见悬浮窗**：`app/__main__.py` 漏了 `window.show()`。
- **动画与吸附互相打架**：展开动画 `finished` 信号未清空 `_anim` 引用，
  且 `snap_to_edge` 对非 Running 状态的动画也让位，导致贴边吸附永久失效。
  修复：动画结束清 `_anim`，仅在动画 Running 时让位。
- **PyInstaller 全量构建约 60-90s，超过 exec 会话前台超时**：必须
  `Start-Process -RedirectStandardOutput` 后台执行 + 轮询日志文件。
- **ISCC 写目标 exe 报 Error 32（文件占用）**：构建前需关闭残留的
  AI4Word.exe / WINWORD.EXE 进程。

**6. 测试与发布验证**

- 离线套件 `tests/run_offline.py`：**57 项全绿**（fake_word 35 +
  format_runner 8 + session 4 + settings 6 + engine 4）。
- 真机：开发态 GUI 窗口标题「AI4Word 悬浮助手」正确；
  `smoke_engine_real_word.py` 三段全绿（流式写入 / arrange 生成-执行-应用 /
  中断-回滚）；打包 exe 存活复测通过；安装包 `AI4Word-Setup-8.0.exe`
  可正常生成（45MB）。

### V9.0: 可用性大修（渲染可靠性 / 进度可读性 / 光标漂移防护）

V8.2 打包链路打通后，真机使用暴露了一批「能用但难用」的问题：输入时悬浮窗背景
变透明看不清、消息流高度被截断看不到进度、预设下拉逐字符、贴边吸附缺失、
生成期间用户挪动光标导致写入位置乱跳。本轮全部从根因修复，并补回归测试。

**1. 窗口渲染：不透明 + 圆角 mask（彻底消灭「输入时背景变透明」）**

- 根因：`WA_TranslucentBackground` 的分层窗口在部分显卡上，子控件局部更新
  （输入框光标闪烁 ~2Hz、头像 30fps `update()`、流式 `setText`）会留下未清除
  的半透明区域——合成器直接把桌面透过来，越输入越透明。
- 定位过程：用 `QWidget.grab()` 与 `QScreen.grabWindow()` 对照、像素分类
  ASCII 图逐区域分析，发现连「控件自己画的离屏渲染」也缺内容（说明并非
  纯合成器问题），再逐层剥离出「输入光标闪烁 / 流式 setText 才触发」的
  局部更新路径。
- 修复：改为**不透明窗口**（去掉 `WA_TranslucentBackground`），palette
  的 Window 角色设为主题深色（自动填充永不露灰底），整幅 paintEvent 自绘
  墨黑渐变，圆角由 `QRegion` mask 切出（经典 shaped-window，无分层合成，
  任何局部更新都绝对安全）。
- 附带修复：`reload_flags` 里 `setWindowFlags` 会重建 HWND，旧实现漏掉
  重新注册全局热键——改设置面板后 Ctrl+Alt+Space 会失效；现在先注销再让
  `showEvent` 在新 HWND 上重注册。

**2. 消息流高度截断（「实时响应进度展示不清楚」的根因）**

- 根因：`QLabel` 设了 `wordWrap` + `RichText` 后，`heightForWidth` 在
  布局里被以错误的宽度求解（尺寸链不稳定），多行消息 label 只分到一行高
  （16px），流式文本大半被裁掉。
- 修复：`messages.py` 新增 `_WrapLabel`——每次 `setText` 与 `resizeEvent`
  用 `QTextDocument` 按标签**实际宽度**确定性地算高并写回 `minimumHeight`，
  绝不依赖布局的 heightForWidth 猜测。
- 气泡改为全宽（Slack 式，颜色区分发送方）：长流式一行容纳 30+ 字，进度
  最清晰；`_scroll_bottom` 追加一次 0ms 延迟兜底（布局重算在事件循环里
  完成，同步取滚动条 maximum 可能还是旧值）。

**3. 进度状态条**

- 引擎新增 `progress(str)` 信号：生成中显示「正在生成… 已接收 N 字」、
  追加补充 / 排版代码生成 / 执行 / 自我修复各阶段均有文案；窗口在标题栏
  下新增 `statusBar` 状态条，idle 时清除。

**4. 预设下拉逐字符（「套用预设不好用」的根因）**

- 根因：`styles.preset_names()` 返回的是「、」连接的**字符串**，而
  `main_window` 用 `list(preset_names())` 拆分——字符串被逐字符拆成
  10+ 个下拉项（论/文/、/公/…），选什么都是「未知预设」报错。
- 修复：新增 `preset_list()` 返回真列表（`preset_names` 保留给 AI 提示词），
  下拉项 = 提示头 + 4 个真预设；应用后引擎回传 `applied` 明细
  （正文=宋体/12磅/缩进2字；H1=黑体/22磅…）。生成中套预设会被礼貌拦截
  （避免与写入器抢样式定义）。

**5. 贴边吸附（拖动磁吸 + 松手吸附）**

- 既有实现只在松手时吸附（`SNAP_MARGIN=26`），且吸附会与展开动画打架。
- 修复：`MainWindow.drag_to()` 拖动中靠近边缘（24px）即磁性吸附、拖回
  中间自动脱离；松手再以 40px 边距吸附一次；吸附半径与松手边距分离调参。
- 附带修复：构造期 `_apply_expanded(instant=True)` 会把默认 (0,0) 当用户
  位置 `_remember_geometry()` 写回 settings，**覆盖保存的窗口位置**——
  每次启动都回到 (0,0)。现在 instant 调用不记忆；并收紧工具条
  （修订/存档定宽、下拉按最短内容宽度）使展开窗口真正达到设计宽 492px
  （原来被布局最小宽度顶到 572px）。

**6. Word 光标漂移防护（写入锚点）**

- 根因：`StreamingWriter` 全程靠 `self.sel`（Word 选区）写入，用户在
  生成期间点一下文档别处，后半段就接在用户光标处。
- 修复：写入器维护**锚点**（一个折叠的动态 `doc.Range`，每次写操作后
  以选区当前位置重新捕获）；每次 `TypeText` / `TypeParagraph` / 字体属性 /
  样式设置前先 `_reselect_anchor()` 把选区拉回锚点。会话边界由引擎调用
  `reset_anchor()` 重置（`write_block` 作为原语入口也重置——由调用方
  定位选区），流式路径 `_write_top_level → _write_md(keep_anchor=True)`
  跨块延续锚点。
- 用 FakeWord（含动态 Range 平移语义）写了 5 个回归测试：草稿中漂移、
  结构块路径漂移、批次之间漂移、`write_block` 入口重置、会话边界重置。

**7. 其他顺带修复**

- `avatar.py`：窗口隐藏到托盘时 30fps 重绘空转（省电）。
- 展开输入框 `ScrollBarAsNeeded`：多行输入可滚动（原 AlwaysOff 超高内容
  被裁）。
- 安装包版本号：`ai4word.iss` 的 `MyAppVersion` 改由 `build.py` 从
  `app.__version__` 注入（`{VERSION}` 占位符），不再手写漂移。

**8. 测试与验证**

- 离线套件 `tests/run_offline.py`：**68 项全绿**（fake_word 35 +
  format_runner 8 + session 4 + settings 6 + engine 4 + ui 10 +
  writer_anchor 5）。
- 视觉回归：`widget.grab()` 与屏幕截图像素级对照，确认消息文字真实绘制、
  窗口 0% 透明像素、紧凑/展开/生成中/中断选择条各形态渲染正确；启动真实
  `ai4word.pyw` 进程存活复测通过。

### V9.2: 全量功能防呆审计修复

对所有已有功能做了一次「完善 + 防呆」审计（双路代码审计 + 逐条复核 + 离线/
真实 Word/真实 API 三级验证），修复真实可复现的缺陷：

**沙箱与执行（P0/P1）**

- `raise SystemExit` 逃逸沙箱：AST 检查未禁 BaseException 家族、run_code 的
  `except Exception` 兜不住——CLI 直接退进程，GUI 里终结 worker 线程（命令
  队列再无人消费，GUI 对 Word 操作整体无响应）。静态层禁这些名 + 运行层
  `except BaseException` 兜底转失败。
- 命令式原语无索引校验：`delete_block(-1)` 因 Python 负索引静默删掉最后一块、
  `replace_block(-1, md)` 把内容写进最后一块。所有原语统一 `_require_index`
  + 空 md 拒绝（留孤儿空段落并错位索引）+ 失活 Range 拒绝写入。
- 文档比块模型短时尾部块保留过时的活 Range → block_map 读错误文本、后续
  编辑在错位内容上操作：失配块 Range 置 None（与 realign_blocks 一致）。
- 事务中途失败永久卡住（`_txn_active` 不复位，之后所有编辑不入快照栈）：
  rollback 先复位标志；引擎/CLI 在 run_code 失败时主动回滚卡住的事务。
- `restore_snapshot` 逐块写入失败被吞（文档半写缺块仍报「已回滚成功」）：
  失败计数后抛错；恢复后清 alignment_warning。
- `replace_text` 不回写块 md → 事务回滚静默吞掉它之前的修改：回写受影响
  块 md（代价是丢 markdown 标记）；返回真实命中/未命中（原无条件返回 True，
  没匹配也记成功偏好）；old/new 校验。
- 删表格块两步都失败时不再弹块（否则文档表格还在、索引整体错位）；表格删除
  加 `_com_retry`（排版重分页期 RPC_CALL_REJECTED 高发）。
- 步数守卫由「每条语句各 50000」改为全程共享（原与文档承诺矛盾，两条
  49999 次循环可全过）。
- fix_code 修正请求携带当前 block_map：失败轮可能已改了一半文档，按旧结构
  臆测索引会错上加错。
- 沙箱补丁：禁字符串常量里的 `__`（`"{0.__globals__}".format(x)` 绕过 AST）；
  禁 COM 高危成员名（`word_app.Run` 执行任意 VBA 宏、`.System` 读写文件、
  `Documents.Open`）。
- session 记忆注入加固：去控制字符、单条 ≤200 字、≤10 条；偏好段措辞从
  「请直接沿用」降级为「可参考」（block_map 已把文档文本喂进提示，恶意
  文档可借 remember() 持久注入）。
- `ai_client` 的网络/HTTP 错误不再吞掉只 print（打包后无控制台，用户对失败
  一无所知）：改为抛出由 GUI/CLI 捕获提示；空回复明确提示；缺密钥即时提示。

**引擎与 GUI**

- 打字档位在生成中无效（worker 处理 write 期间不消费队列，`send('speed')`
  要等生成结束才执行）：新增 `worker.set_speed()` 直接写纯 Python 属性，
  立即对进行中的生成生效。
- 「追加补充」生成中中断按钮无效（extra 分支无 `_interrupt` 检查）：补上。
- 中断后点「追加补充…」是死路（选择条隐藏、发送被拦截、只能干等 5 分钟
  超时）：保留选择条可改主意；Esc 优先退出追加模式；中断时自动展开面板。
- 退出卡死：中断选择等待最长阻塞 300 秒，quit 要等它才被处理 →
  `worker.choose('keep')` 先解除等待，quit 随即被消费。
- 查看类命令（块地图）不再启动新 Word 进程：只附加已运行实例（新增
  `get_word_attach` + `attach_factory`），Word 没运行时提示。
- 用户切换活动文档后引擎缓存仍指向旧文档（内容写进用户不看的文档）：每条
  命令前检查 `ActiveDocument`，变化时自动重绑（真实 Word 双文档验证通过）。
- `_wait_choice` 先清残留选择（双按时序留下的陈旧选择会被下次中断误取，
  造成「自动回滚」）。
- 生成中点「修订 / 预设 / 存档」不再延迟到生成结束后才执行：UI 层直接拒绝
  并提示先中断。
- settings.json 类型级损坏（geometry 写成字符串等）启动必崩且无恢复路径：
  load 时按键校验类型非法回退默认；`save()` 返回成功状态供提示。
- 开机自启注册表写入失败不再静默成功（复选框勾着实际没生效）：对话框与
  托盘都检查返回值 + 回退勾选 + 警告。
- 隐藏窗口停在已断开显示器上召回后不可见：`summon()` 先 clamp 几何；
  隐藏态下改设置热键失联：`reload_flags` 无论可见性都注册；RegisterHotKey
  失败首次提示。
- 流式生成时用户上滚阅读不再被每个 chunk 拽回底部；消息流气泡上限 200 条。
- 连按两次 Enter 第二条文本丢失（拒收但已清空）：发送时乐观置 busy，第二条
  走「中断」且文本保留。
- 输入长度上限 20000 字；输入法组合态 Enter 不触发发送（preedit 守卫）。

**验证**

- 离线套件 **84 项全绿**（新增 13 项加固回归）。
- 真实 Word：导入、冒烟、引擎、锚点漂移、**文档切换重绑定**、真实 API
  端到端全部通过；GUI 启动存活复测通过。

### V9.1: 三个真机反馈的修正（圆角锯齿 / 已有文档读取 / 写作废话）

**1. 窗口边缘锯齿（V9.0 的 mask 盖掉了 Win11 原生圆角）**

V9.0 用 QRegion mask 切圆角，在 Win11 上反而难看：mask 会**覆盖**系统
合成器的原生圆角，留下像素级台阶。

- 定位：实验对照发现 Win11（build ≥ 22000）默认就对顶层窗口做 DWM
  抗锯齿圆角——只要不设 mask，角落就是平滑的。
- 修复：`MainWindow._apply_window_shape()`——Win11 上清掉 mask，并调
  `DwmSetWindowAttribute(DWMWA_WINDOW_CORNER_PREFERENCE, DWMWCP_ROUND)`
  显式固化圆角偏好；早于 Win11（无 DWM 圆角 API）才降级回 QRegion mask。
  描边半径改为与 DWM 一致的 8px。形状在 showEvent 里按 HWND 重建
  应用（setWindowFlags 会换 HWND）。

**2. 无法读取已接入的非空 Word 文档**

- 根因：接入 Word 后只调 `load_blocks()`——它从 `doc.Variables` 恢复
  AI4Word 自己存下的块模型。用户**新打开**的已有文档没有存档变量，
  于是块地图空、块编辑无从下手，程序对已有内容「失明」。
- 修复：`DocModel.import_document(limit=600)`——无存档时逐段落扫描
  `doc.Paragraphs` 登记为块：按 `Style.NameLocal` 匹配内建标题样式
  （-2..-7，兼容中文「标题 1」与英文 "Heading 1"）识别 heading1-6，
  其余为 paragraph；空段落也登记（保持与文档段落 1:1 对齐，
  `rebuild_ranges` 不报漂移）；块 Range 与流式登记一致地**不含**末尾
  段落标记（否则 replace_block 会吞掉段落分隔符——实测发现）。
- 接线：GUI 引擎 `_ensure_ready` 与 CLI `main()` 在 `load_blocks` 返回 0
  时调用，提示「已读取现有文档 N 个块」；块地图、replace/insert/delete
  原语随即作用于已有文字。
- 测试：fake 下 3 个用例（导入/空文档/导入后编辑）+ 真实 Word 验证
  （`tests/import_real_word.py`，含中文标题样式识别与 replace_block）。

**3. AI 的前言 / 修改方向讨论被写进文档**

- 根因：写作阶段的系统提示词只有「用 Markdown 回复」，模型把需求当
  对话，输出「好的，我来…需要确认…吗？」这类元话语——每个字都直接
  流式写进了 Word。
- 修复：`app/agent.py` 提取共享 `WRITER_SYSTEM`——明确「直接输出正文
  本身，禁止前言/问候/确认/提问/思路说明/修改方向讨论/所需条件询问/
  总结，因为每个字都会立即写入文档」；写作与追加补充（GUI 引擎两处、
  CLI 一处）与 e2e 测试统一引用。真实 API 验证：输出纯正文无前言。

**4. 验证**

- 离线套件 **71 项全绿**（fake_word 38 / format_runner 8 / session 4 /
  settings 6 / engine 4 / ui 10 / writer_anchor 5）。
- 真实 Word 冒烟、引擎 × 真实 Word、真实 API 端到端全部通过；
  真实 API 写作提示词检查无前言。


### V9.3: 隐藏调试模式 + 全功能实测

**调试模式（完全隐藏，零界面改动）**

- 触发方式（任一即可）：
  - 命令行参数：`AI4Word.exe -debug` / `python ai4word.pyw -debug` / `python -m app -debug`
  - 程序名后缀含 `-debug`：如 `AI4Word-debug.exe`、`ai4word-debug.pyw`（只看 basename，`D:\xxx-debug\x\AI4Word.exe` 不会误报）
- 日志位置：`%APPDATA%\AI4Word\debug.log`（回退链与 `settings.py` 一致：APPDATA → 用户主目录 → TEMP）
- 行格式：`2026-10-04 10:23:45.123 [MainThread] INFO 事件名 key=value …`；追加写，`threading.Lock` 保证 GUI 线程与 worker 线程交错安全；超 5MB 滚动为 `debug.log.bak`（只留 1 个备份）；每次 `init()` 写运行头（时间、版本、是否打包、argv、Python/exe 路径、Windows build、日志路径）
- 敏感字段（字段名含 api_key/key/token）自动打码为 `sk-…abcd`；流式正文只记计数与 200 字预览；用户提示词与 AI 生成的排版代码记全文；普通字段超 300 字符截断（`full=True` 不截）
- 装有 `sys.excepthook` 与 `threading.excepthook`：未捕获异常先写 debug.log 再走原钩子（与 `ai4word.pyw` 的 crash.log 兼容，不冲突）
- **非调试模式零副作用**：`is_debug()` 为 False 时 `log/warn/error/exc` 直接返回，不建文件、不格式化字符串、不装钩子（对照组实测验证：全功能操作一遍后无任何日志文件生成）
- 接线覆盖：启动/退出、GUI 全部按钮与状态流转（展开收起、头像、两处发送、中断、回滚/保留/追加、速度三档、写作/排版切换、修订开关、预设、存档、块地图、设置对话框、拖动吸附、summon、热键、Esc 隐藏）、后台引擎（写作流/中断三选择/排版 gen_code→run_code→自我修复/预设/存档/块地图/Word 连接与重绑定）、CLI/GUI 共享核心（`ai_client` / `streaming_writer` / `doc_model` / `format_runner`）、托盘、设置对话框、自启、settings 异常路径

**测试发现并修复的真实 bug**

1. **QPen 类型错误**（`app/main_window.py` 紧凑模式边框）：`QPen(AMBER, 1.3)` 直接传字符串色值（`AMBER = "#f2a93b"`），PySide6 无此重载，paintEvent 每帧抛 TypeError。改为 `border = AMBER if self._busy else INK_4` + `QPen(QColor(border), 1.3)`。
2. **中断等待期「追加补充」被拦截**：`_on_send_compact` / `_on_send_panel` 的 `if self._busy: interrupt(); return` 把中断等待选择期间的「追加补充」也一并拦掉了（用户无法追加，只能三选一）。改为 `if self._busy and not self._extra_mode:`。

**测试中发现的机制（harness / 环境层面，非 app bug）**

- 主线程直接读 Word COM 文档文本会抛 RPC_E_WRONG_THREAD；测试 harness 改为投递合成命令 `_sync_doc_text` 由 worker 线程读取，且 worker 端先写 text 再写 rsp（rsp 是释放标志，先写会读到陈旧 None）
- worker 阻塞在中断选择队列（`_wait_choice`）时不服务命令队列，此时跨线程读取会超时返回 None
- 解释器关闭时，主线程（从未 CoInitialize）释放 worker 线程持有的 COM 对象触发 0xC0000409（STATUS_STACK_BUFFER_OVERRUN）原生崩溃；退出流程必须先在 worker 线程内关闭文档/Word 并清理引用
- 排版自我修复链路走 `fix_code`（内部直接再问 AI），不会二次调用 `gen_code`
- 输入框 `setMaxLength(MAX_PROMPT_CHARS)` 会静默截断 `setText`，超长拒绝分支不可达，需直接调 `_send` 触发
- Atria API 偶发连接失败（0.02~0.04s 内失败）属网络波动，引擎按 `stream_error` 正确处理并回 idle
- offscreen 平台下全局热键注册拿不到真实 HWND，注册失败属平台假象

**验证**

- `tests/test_debug_mode.py`：15 个离线单测全绿（argv 触发、exe 名触发、无参不建文件不装钩子、目录名不误报、字段打码、截断与 full、5MB 滚动、`exc()` 写 traceback 等）
- `tests/debug_walkthrough.py`：真实 Word + 真实 Atria API 全功能矩阵实测通过——发送（紧凑/面板）、流式写作、中断→回滚/保留/追加补充三支路、慢/自/快三档、写作/排版模式、修订开关、四个预设、存档、块地图开关与点击定位、设置对话框（开/改/确定/保存失败回退）、托盘菜单（显示/自启/退出）、拖动吸附、summon、Esc 隐藏、超长输入拒绝、排版 self-repair（gen_code 失败→fix_code→再执行成功）、退出流程；debug.log 中除已知预期项外无 ERROR/异常：
  - `settings_load_failed`（首跑临时 settings 不存在，预期）
  - `stmt_failed idx=1 src="x = 1 / 0"`（harness 注入的失败，预期）
  - `run_code_sandbox_rejected`（沙箱正确拒绝 AI 代码中的 Import，预期且随后自修复成功）
  - `autostart_disable_failed`（自启注册表项不存在，预期路径）
  - `settings_save_failed`（harness 故意指向不存在目录，验证失败回退）
  - `hotkey_register ok=false`（offscreen 平台假象）
- 对照组：删除 debug.log 后不带 `-debug` 重跑全部操作，构成路径全部清空（`%APPDATA%\\AI4Word` 下仅余 9/30 遗留的 settings.json），通过

**深度检查轮（按键 / 输入法 / 关闭 / 托盘，全部通过）**

补齐上轮未覆盖的输入与窗口行为，断言逐条通过（真实 Word + 真实 API，offscreen）：

- 紧凑框与面板框的 Enter 发送；Shift+Enter 仅插换行不发送；**输入法组字期 Enter 不发送**（`QInputMethodEvent` 模拟 preedit 组字状态，组字结束后恢复）
- X 关闭按钮 → 隐藏到托盘 + 首次隐藏提示（新增 `close_to_tray` 日志点）；托盘双击召回、Context 激活不召回
- 生成中存档 / 套预设被拒（队列空、预设 combo 归零）；设置对话框密码显示切换、Cancel 不改设置、自启启用失败 → 警告且对话框仍关闭、`auto_start` 保持 False

**深度检查发现并修复的问题**

1. `settings.load()` 首次运行（settings.json 不存在）落到异常分支并记 `WARN settings_load_failed`：每次全新安装的首跑都会在调试日志里产生一条误导性告警。改为 `FileNotFoundError` 记 `INFO settings_load_first_run`（正常首跑），其余 OSError / JSON 损坏仍记 `WARN`。
2. `app/settings.py` 模块 docstring 内 `%APPDATA%\AI4Word` 的 `\A` 是非法转义（Python 3.12 SyntaxWarning，未来将报错），docstring 改为原始字符串（r 前缀三引号）。该告警此前在 pytest 每次运行时刷屏。

**本轮日志 triage 结论**

1138 行日志中 8 条 WARN/ERROR、1 条 traceback，全部为注入的失败路径或正常首跑（见上轮清单），无真实 bug；流式写作路径无重试、无漂移、每个流恰好一次 anchor_reset；排版自我修复链路（gen_code 失败 → fix_code 真实 API 9.3s 返回 → 沙箱拒绝含 Import 的修复码 → arrange_failed 正确上报）行为符合设计。对照组（无 `-debug`）再次全量通过且未生成任何日志文件。
**深度实验工作流（30+ 轮迭代）**

在调试模式之上建立了“实验 → 分析日志 → 修复 → 再实验”的闭环：
`tests/experiment_lab.py`（offscreen Qt + fake Word + 脚本化 AI，11 个场景
矩阵）单轮 exit code 0 = 干净。共完成 99 轮离线实验 + 2 轮真机走查（真实
Word + Atria API），修复 2 个真实 bug（块顺序错位导致回滚重放错误 / 流式
段落首尾空白导致回滚差 1 字符）与 4 处接线、时序问题，最终连续 31 轮
0 findings、110 项单测通过、真机走查 0 失败。完整迭代表与修复清单见
[EXPERIMENT_LOG.md](EXPERIMENT_LOG.md)。


---

## 核心技术栈

- **AI服务**: Atria（Intern AI discovery 平台，OpenAI 兼容接口，SSE 流式）
- **Word交互**: `pywin32` (COM接口)
- **Markdown解析**: `markdown-it-py`
- **API通信**: `requests`
- **环境管理**: `python-dotenv`
- **桌面GUI**: `PySide6`
- **打包分发**: `PyInstaller` + Inno Setup 6

---

## 主要文件

- `main.py` - 主程序：流式写作 + 交互式排版编排
- `ai_client.py` - Atria 客户端（流式 / 非流式）
- `streaming_writer.py` - 流式 Markdown 写入器
- `doc_model.py` - 块级文档模型与编辑原语
- `format_runner.py` - 逐语句可视化执行器
- `sandbox.py` - AI 生成代码的 AST 静态检查与循环步数护栏
- `session.py` - 会话记忆（交互历史 / 用户偏好）
- `styles.py` - 样式预设（论文 / 公文 / 简历 / 博客）
- `tests/` - 离线测试与真实 Word 冒烟测试
- `requirements.txt` - 依赖列表
- `README.md` - 项目说明文档
- `ai4word.pyw` - GUI 启动入口（开发态与打包共用，兜底异常写 crash.log）
- `app/` - 桌面悬浮窗 GUI（PySide6）：窗口 / 引擎 / 头像 / 消息流 / 设置 / 托盘
- `build/` - 打包脚本与安装包模板（PyInstaller + Inno Setup）
- `requirements-dev.txt` - 打包期依赖（PyInstaller / Pillow）
- `DEVELOPMENT_LOG.md` - 开发文档
