"""AI4Word 主程序：流式生成 + 打字机写入 + 可视化 AI 排版。

阶段一：SSE 流式接收 AI 生成的 Markdown，边收边按块解析、应用格式并以
        打字机动画写入活动的 Word 文档（真正的流式输出）。
阶段二：AI 智能排版模式。AI 生成的 pywin32 代码被逐语句可视化执行：
        终端实时显示每条语句与执行结果，Word 里受影响区域滚动进视野并闪烁。
        编辑已写入的内容请优先使用注入的块编辑原语（replace_block 等），
        它们基于每块登记的动态 Range，定位精确、不会破坏文档其余部分。
"""
import os

import win32com.client
from dotenv import load_dotenv

from ai_client import ai_request, ai_stream
from doc_model import DocModel
from format_runner import run_code
from streaming_writer import StreamingWriter

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
    """获取或创建一个 Word 应用实例。"""
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
- 块函数的 md 参数支持 Markdown：段落、1-6 级标题、有序/无序列表（支持嵌套层级）、
  引用块、围栏代码块、行内代码、超链接、水平线、表格。
- 常量: `WD_STYLE_NORMAL`(-1), `WD_STYLE_HEADING_1`(-2), `WD_STYLE_HEADING_2`(-3),
  `WD_STYLE_HEADING_3`(-4), `WD_STYLE_LIST_BULLET`(-49), `WD_STYLE_LIST_NUMBER`(-50),
  `WD_ALIGN_PARAGRAPH_LEFT`(0), `WD_ALIGN_PARAGRAPH_CENTER`(1), `WD_ALIGN_PARAGRAPH_RIGHT`(2),
  `WD_PRINT_VIEW`(3), `WD_ORIENT_LANDSCAPE`(1), `WD_LINE_SPACING_1_5`(1)

**三大黄金法则**:
1. **内容编辑必须用块原语**: 涉及增、删、改文字时，**必须**使用上面的块函数（通过 block_map() 的索引定位），
   而不是自己拼接 Range 或用 Find 去猜位置。块函数会精确定位到目标段落，并保留文档其余部分。
   多步内容编辑推荐声明式写法：先用 `preview_edit(spec)` 确认前后对比，再 `apply_edit(spec)` 执行；
   每条 apply_edit 只做一处修改，多条修改分多条语句写。
2. **格式修改直接改样式 (Style)**: 对字体、段落格式（对齐、缩进、行距）的修改，**必须**通过修改文档的样式定义完成。
   - 示例: `doc.Styles(WD_STYLE_NORMAL).Font.Name = "宋体"`
   - 反例（不要这样做）: 不要用 Find 或遍历所有段落的方式改格式。
   - 首行缩进 N 个字符 = 字号 × N 磅: `style.ParagraphFormat.FirstLineIndent = style.Font.Size * 2`

**错误处理**:
- 禁止自己写 try...except，让主程序捕获错误（唯一例外：设置 `doc.PageSetup` 时可用 try...except 包裹）。

**输出要求**:
- 只输出纯 Python 代码。不要包含 import、函数定义、word_app/doc 的重新声明，也不要任何 Markdown 标记或解释文字。

**字号参考（磅）**: 初号 42, 小初 36, 一号 26, 小一 24, 二号 22, 小二 18, 三号 16, 小三 15, 四号 14, 小四 12, 五号 10.5, 小五 9
"""


def clean_code(raw):
    return (raw or "").strip().replace("```python", "").replace("```", "").strip()


def gen_code(prompt, api_key):
    """让 AI 生成排版/编辑代码（第一次尝试），提示词中附带文档结构地图。"""
    user_prompt = prompt
    try:
        structure = model_block_map()
    except Exception:
        structure = "(无法获取文档结构)"
    user_prompt += "\n\n当前文档结构（块索引从 0 开始）：\n" + structure
    print("正在请求 AI 生成代码（第 1 次尝试）...")
    raw = ai_request(user_prompt, api_key, CODEGEN_SYSTEM)
    if not raw:
        return None
    code = clean_code(raw)
    print("AI 生成的代码:\n---\n" + code + "\n---")
    return code


def fix_code(prompt, failed_code, error_msg, api_key):
    """执行失败后，让 AI 根据错误信息重新生成一段不同思路的代码。"""
    system_prompt = f"""
你是顶级的 Python 调试专家，专精于 pywin32 的 Word 自动化。
为了完成用户指令「{prompt}」，之前运行了如下代码：
--- FAILED CODE ---
{failed_code}
--- END FAILED CODE ---
但它失败了，错误信息：`{error_msg}`

你的任务：
1. 分析失败原因。
2. 用一种**全新的、不同的方法**完成原始指令（例如样式设置失败就改用块函数或反之）。
3. 只输出修正后的完整 Python 代码片段；内容编辑优先使用块函数
   (replace_block/insert_after/delete_block/replace_text 等，通过 block_map() 的索引定位)，
   格式修改继续通过修改样式定义完成。
4. 不要自己处理异常（PageSetup 除外）；不要输出任何解释文字。
"""
    print("\n代码执行失败，正在请求 AI 自我修复...")
    raw = ai_request(prompt, api_key, system_prompt)
    if not raw:
        return None
    code = clean_code(raw)
    print("AI 生成的修正代码:\n---\n" + code + "\n---")
    return code


# 主流程用到的全局模型（gen_code 读取结构地图）
_model = None


def model_block_map():
    return _model.block_map() if _model else "(空文档)"


def main():
    global _model

    dotenv_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    load_dotenv(dotenv_path=dotenv_path)
    api_key = os.getenv("ATRIA_API_KEY")
    if not api_key:
        print("错误：未能加载有效的 ATRIA_API_KEY（请在 .env 中设置）。")
        return

    app = get_word()
    if not app:
        print("错误：无法启动或连接到 Word。")
        return
    doc = app.Documents.Add() if app.Documents.Count == 0 else app.ActiveDocument
    sel = app.Selection
    print("\nWord 已准备就绪。")

    writer = StreamingWriter(app, doc, sel, model=None, char_delay=0.01)
    model = DocModel(app, doc, sel, writer)
    writer.model = model
    _model = model

    # --- 阶段一：流式生成 + 动画写入 ---
    prompt = input("请输入您的写作需求：")
    if prompt:
        system_prompt = "你是一个乐于助人的助手，你总是使用 Markdown 格式进行回复。"
        print("\n正在流式生成并以动画效果写入 Word...")
        received = False
        for piece in ai_stream(prompt, api_key, system_prompt):
            if not piece:
                continue
            received = True
            print(piece, end="", flush=True)  # 终端同步显示流式原文
            writer.feed(piece)
        writer.flush()
        print()  # 换行结束终端原文回显
        if received:
            print("写入完成。")

    # --- 阶段二：可视化 AI 智能排版 ---
    print("\n--- 进入 AI 智能排版模式（逐语句可视化执行 + 自我修复）---")
    print("编辑内容请用块索引，例如「改写第 2 块：……」「在第 0 块后插入一段……」；")
    print("格式指令例如「将所有一级标题居中并改为蓝色」。输入「退出」结束。")

    exec_globals = {
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

    while True:
        model.watch.poll()  # 空闲点：检测用户在停顿期间的手动改动
        notices = model.drain_events()
        for ev in notices:
            print(f"  [注意] {ev}")
        fmt = input("\n您想如何调整？> ").strip()
        if not fmt:
            continue
        if fmt.lower() in ("退出", "exit", "quit"):
            print("程序已结束。")
            break

        prompt = fmt
        if notices:
            # 用户在生成期间动过文档：把漂移警告带给代码生成模型，避免按旧结构猜索引
            prompt += "\n\n注意：用户在我生成期间手动编辑了文档，块索引可能已偏移，"
            prompt += "请先重新调用 block_map() 确认当前结构。"
        code = gen_code(prompt, api_key)
        if not code:
            print("无法理解您的指令或 AI 未返回有效代码。")
            continue

        ok, err = run_code(code, exec_globals)
        if ok:
            print("代码执行完毕，已应用。")
            continue

        corrected = fix_code(fmt, code, err, api_key)
        if not corrected:
            print("AI 未能生成有效的修正代码。")
            continue
        ok2, err2 = run_code(corrected, exec_globals)
        if ok2:
            print("修正代码执行完毕，已应用。")
        else:
            print(f"自我修复仍然失败（{err2}），请尝试更简单的指令。")


if __name__ == "__main__":
    main()
