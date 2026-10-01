# -*- coding: utf-8 -*-
"""AI4Word 共享 Agent：代码生成提示词、自我修复、Word 执行环境装配。

CLI（main.py）与 GUI（app/engine.py）共用这一层，保证两端的代码生成
提示词、沙箱执行与可用符号表完全一致。
"""
import win32com.client

from styles import apply_preset, preset_names

# --- Word COM 常量 ---
WD_STYLE_NORMAL = -1
WD_STYLE_HEADING_1 = -2
WD_STYLE_HEADING_2 = -3
WD_STYLE_HEADING_3 = -4
WD_STYLE_LIST_BULLET = -49
WD_STYLE_LIST_NUMBER = -50
WD_ALIGN_PARAGRAPH_LEFT = 0
WD_ALIGN_PARAGRAPH_CENTER = 1
WD_ALIGN_PARAGRAPH_RIGHT = 2
WD_REPLACE_ALL = 2
WD_PRINT_VIEW = 3
WD_ORIENT_LANDSCAPE = 1
WD_LINE_SPACING_1_5 = 1


def get_word():
    """获取或创建一个 Word 应用实例（晚期绑定，兼容已运行的 Word）。"""
    try:
        app = win32com.client.GetObject(None, "Word.Application")
        print("已连接到现有的 Word 程序。")
    except Exception:
        try:
            app = win32com.client.Dispatch("Word.Application")
            print("启动了一个新的 Word 程序。")
        except Exception:
            return None
    app.Visible = True
    return app


def get_word_attach():
    """只附加到已运行的 Word；没运行返回 None（不启动新进程）。

    「块地图」这类查看类命令用：用户只想看看文档结构，贸然启动一个
    Word 窗口是打扰。
    """
    try:
        return win32com.client.GetObject(None, "Word.Application")
    except Exception:
        return None


WRITER_SYSTEM = (
    "你是 AI4Word 的写作助手，负责把用户要求的内容写进 Word 文档。"
    "直接输出文档正文本身，使用 Markdown 格式（标题 / 列表 / 粗体 / 引用 / 表格等）。"
    "你的每一个字都会立即逐字写入 Word 文档，因此禁止输出任何前言、问候、确认、"
    "提问、思路说明、修改方向讨论、所需条件询问或总结——只输出正文。"
)


CODEGEN_SYSTEM = """
你是世界一流的 Python 程序员，专精于使用 pywin32 库对 Microsoft Word 文档进行自动化排版与编辑。
用户指令是针对一个已经打开、且已经写有内容的 Word 文档。你的任务是编写一小段 Python 代码来完成需求。

**可用的全局变量与函数**:
- `word_app`: Word 应用程序对象；`doc`: 当前活动文档对象；`sel`: 当前选区。
- `model`: 文档块模型，下面的函数也都已注入，可直接调用：
  - `block_map()`: 返回文档结构地图，格式为每行 `[索引] 类型: 内容预览`。**索引从 0 开始。**
  - `get_block_text(i)`: 获取第 i 块的纯文本。
  - `replace_block(i, new_md)`: 用新 Markdown 内容改写第 i 块。
  - `insert_after(i, md)`: 在第 i 块之后插入新内容（可为多块 Markdown）。
  - `insert_at_end(md)`: 在文档末尾追加内容。
  - `delete_block(i)`: 删除第 i 块。
  - `replace_text(old, new)`: 全文查找替换。
  - `select_block(i)`: 滚动到第 i 块并高亮（演示用）。
  - `preview_edit(spec)`: 只预览不执行，返回受影响块的文本级前后对比。
  - `apply_edit(spec)`: 执行声明式编辑 spec = {"op": ..., "index": i, "md": ...}，
    op 支持 "replace" / "insert_after" / "insert_at_end" / "delete"。
  - `undo(times=1)`: 撤销最近 times 次编辑（Word 编辑栈）。
  - `drain_events()`: 取出流式停顿期间检测到的用户手动改动事件。
  - 修订协同（人机互审）：`review_on()` / `review_off()` 开关修订模式，
    `has_revisions()` 查未决修订，`accept_block_revisions(i)` / `reject_block_revisions(i)`
    按块接受 / 拒绝修订。
  - 快照与撤销：`snapshot()` 取整篇快照，`restore_snapshot(snap)` 按快照重写恢复，
    `model_undo()` / `model_redo()` 快照级撤销 / 重做，
    `begin_txn()` / `commit_txn()` / `rollback_txn()` 把多步编辑组成原子事务。
  - 持久化与重对齐：`save_blocks()` / `load_blocks()` 块模型存入文档变量并恢复，
    `realign_blocks()` 用户手改文档后按文本相似度重新对齐块。
  - 长文档布局：`insert_toc()` 在文档末尾插入目录，`set_header(t)` / `set_footer(t)`
    设置页眉 / 页脚，`insert_page_break()` 在光标处插入分页符。
  - 效率工具：`apply_preset(name)` 一套样式预设（论文 / 公文 / 简历 / 博客），
    `preset_names()` 列出可用预设，`set_speed("auto" / "slow" / "fast")` 打字速度，
    `remember(note)` 记住用户的长期排版偏好。
- 块函数的 md 参数支持 Markdown：段落、1-6 级标题、有序/无序列表（支持嵌套层级）、
  引用块、围栏代码块、行内代码、超链接、水平线、表格、图片（![alt](src)）。
- 常量: `WD_STYLE_NORMAL`(-1), `WD_STYLE_HEADING_1`(-2), `WD_STYLE_HEADING_2`(-3),
  `WD_STYLE_HEADING_3`(-4), `WD_STYLE_LIST_BULLET`(-49), `WD_STYLE_LIST_NUMBER`(-50),
  `WD_ALIGN_PARAGRAPH_LEFT`(0), `WD_ALIGN_PARAGRAPH_CENTER`(1), `WD_ALIGN_PARAGRAPH_RIGHT`(2),
  `WD_PRINT_VIEW`(3), `WD_ORIENT_LANDSCAPE`(1), `WD_LINE_SPACING_1_5(1)`

**三大黄金法则**:
1. **内容编辑必须用块原语**: 涉及增、删、改文字时，**必须**使用上面的块函数（通过 block_map() 的索引定位），
   而不是自己拼接 Range 或用 Find 去猜位置。块函数会精确定位到目标段落，并保留文档其余部分。
   多步内容编辑推荐声明式写法：先用 `preview_edit(spec)` 确认前后对比，再 `apply_edit(spec)` 执行；
   每条 apply_edit 只做一处修改，多条修改分多条语句写。
2. **格式修改直接改样式 (Style)**: 对字体、段落格式（对齐、缩进、行距）的修改，**必须**通过修改文档的样式定义完成。
   - 示例: `doc.Styles(WD_STYLE_NORMAL).Font.Name = "宋体"`
   - 反例（不要这样做）: 不要用 Find 或遍历所有段落的方式改格式。
   - 首行缩进 N 个字符 = 字号 × N 磅: `style.ParagraphFormat.FirstLineIndent = style.Font.Size * 2`
3. **成套排版用预设**: 整体风格（论文 / 公文等）优先 `apply_preset("论文")`，再在上面做微调。

**错误处理**:
- 禁止自己写 try...except，让主程序捕获错误（唯一例外：设置 `doc.PageSetup` 时可用 try...except 包裹）。

**输出要求**:
- 只输出纯 Python 代码。不要包含 import、函数定义、word_app/doc 的重新声明，也不要任何 Markdown 标记或解释文字。

**字号参考（磅）**: 初号 42, 小初 36, 一号 26, 小一 24, 二号 22, 小二 18, 三号 16, 小三 15, 四号 14, 小四 12, 五号 10.5, 小五 9
"""


def clean_code(raw):
    return (raw or "").strip().replace("```python", "").replace("```", "").strip()


def gen_code(prompt, api_key, session=None, block_map_fn=None, sink=print, model=None):
    """让 AI 生成排版/编辑代码（第一次尝试），提示词中附带文档结构地图与会话记忆。"""
    from ai_client import ai_request

    user_prompt = prompt
    try:
        structure = block_map_fn() if block_map_fn else "(空文档)"
    except Exception:
        structure = "(无法获取文档结构)"
    user_prompt += "\n\n当前文档结构（块索引从 0 开始）：\n" + structure
    if session is not None:
        memory = session.memory_prompt()
        if memory:
            user_prompt += "\n\n" + memory
    sink("正在请求 AI 生成代码（第 1 次尝试）...")
    raw = ai_request(user_prompt, api_key, CODEGEN_SYSTEM, model=model)
    if not raw:
        return None
    code = clean_code(raw)
    sink("AI 生成的代码:\n---\n" + code + "\n---")
    return code


def fix_code(prompt, failed_code, error_msg, api_key, session=None, sink=print,
             block_map_fn=None, model=None):
    """执行失败后，让 AI 根据错误信息重新生成一段不同思路的代码。

    修正代码必须看到「当前」文档结构：失败的那轮可能已经改了一半文档，
    按原始指令臆测索引会错上加错。
    """
    from ai_client import ai_request

    memory = ""
    if session is not None:
        memory = session.memory_prompt()
    memory_block = ("\n\n**会话记忆**（除非用户明确改变，请沿用，但可纠正）：\n" + memory) if memory else ""
    structure = "(无法获取文档结构)"
    if block_map_fn:
        try:
            structure = block_map_fn()
        except Exception:
            pass
    system_prompt = f"""
你是顶级的 Python 调试专家，专精于 pywin32 的 Word 自动化。
为了完成用户指令「{prompt}」，之前运行了如下代码：
--- FAILED CODE ---
{failed_code}
--- END FAILED CODE ---
但它失败了，错误信息：`{error_msg}`{memory_block}

注意：失败的代码可能已经对文档应用了一部分修改。当前文档结构是：
--- CURRENT BLOCK MAP ---
{structure}
--- END CURRENT BLOCK MAP ---
请以这个结构为准定位索引，不要按旧结构臆测。

你的任务：
1. 分析失败原因。
2. 用一种**全新的、不同的方法**完成原始指令（例如样式设置失败就改用块函数或反之）。
3. 只输出修正后的完整 Python 代码片段；内容编辑优先使用块函数
   (replace_block/insert_after/delete_block/replace_text 等，通过 block_map() 的索引定位)，
   格式修改继续通过修改样式定义完成。
4. 不要自己处理异常（PageSetup 除外）；不要输出任何解释文字。
"""
    sink("\n代码执行失败，正在请求 AI 自我修复...")
    raw = ai_request(prompt, api_key, system_prompt, model=model)
    if not raw:
        return None
    code = clean_code(raw)
    sink("AI 生成的修正代码:\n---\n" + code + "\n---")
    return code


def build_exec_globals(app, doc, sel, model, writer, session):
    """注入给 AI 生成代码的执行环境：Word 对象 + 块原语 + 常量。"""
    return {
        "word_app": app,
        "doc": doc,
        "sel": sel,
        "model": model,
        "block_map": model.block_map,
        "get_block_text": model.get_block_text,
        "replace_block": model.replace_block,
        "insert_after": model.insert_after,
        "insert_at_end": model.insert_at_end,
        "delete_block": model.delete_block,
        "replace_text": model.replace_text,
        "select_block": model.select_block,
        "preview_edit": model.preview_edit,
        "apply_edit": model.apply_edit,
        "undo": model.undo,
        "drain_events": model.drain_events,
        "review_on": model.review_on,
        "review_off": model.review_off,
        "has_revisions": model.has_revisions,
        "accept_block_revisions": model.accept_block_revisions,
        "reject_block_revisions": model.reject_block_revisions,
        "save_blocks": model.save_blocks,
        "load_blocks": model.load_blocks,
        "realign_blocks": model.realign_blocks,
        "snapshot": model.snapshot,
        "restore_snapshot": model.restore_snapshot,
        "model_undo": model.model_undo,
        "model_redo": model.model_redo,
        "begin_txn": model.begin_txn,
        "commit_txn": model.commit_txn,
        "rollback_txn": model.rollback_txn,
        "insert_toc": model.insert_toc,
        "set_header": model.set_header,
        "set_footer": model.set_footer,
        "insert_page_break": model.insert_page_break,
        "apply_preset": lambda name: apply_preset(doc, name),
        "preset_names": preset_names,
        "set_speed": writer.set_speed,
        "remember": session.remember,
        "WD_STYLE_NORMAL": WD_STYLE_NORMAL,
        "WD_STYLE_HEADING_1": WD_STYLE_HEADING_1,
        "WD_STYLE_HEADING_2": WD_STYLE_HEADING_2,
        "WD_STYLE_HEADING_3": WD_STYLE_HEADING_3,
        "WD_STYLE_LIST_BULLET": WD_STYLE_LIST_BULLET,
        "WD_STYLE_LIST_NUMBER": WD_STYLE_LIST_NUMBER,
        "WD_ALIGN_PARAGRAPH_LEFT": WD_ALIGN_PARAGRAPH_LEFT,
        "WD_ALIGN_PARAGRAPH_CENTER": WD_ALIGN_PARAGRAPH_CENTER,
        "WD_ALIGN_PARAGRAPH_RIGHT": WD_ALIGN_PARAGRAPH_RIGHT,
        "WD_REPLACE_ALL": WD_REPLACE_ALL,
        "WD_PRINT_VIEW": WD_PRINT_VIEW,
        "WD_ORIENT_LANDSCAPE": WD_ORIENT_LANDSCAPE,
        "WD_LINE_SPACING_1_5": WD_LINE_SPACING_1_5,
    }
