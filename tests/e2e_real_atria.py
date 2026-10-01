"""真实端到端测试：真实 Atria API + 真实 Word（新建空文档，不动活动文档）。

复用 main.py 的 gen_code / fix_code 与全部模块，验证完整闭环：
真实 SSE 流式生成 -> 打字机写入 -> 块登记 -> AI 代码排版（含块编辑、样式修改）。
需要 .env 中的 ATRIA_API_KEY；需要本机安装 Word。
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import win32com.client
from dotenv import load_dotenv

import ai_client
import main
from doc_model import DocModel
from format_runner import run_code
from streaming_writer import StreamingWriter

WD_STYLE_NORMAL = -1
WD_STYLE_HEADING_1 = -2
WD_STYLE_HEADING_2 = -3
WD_STYLE_HEADING_3 = -4
WD_STYLE_LIST_BULLET = -49
WD_ALIGN_PARAGRAPH_LEFT = 0
WD_ALIGN_PARAGRAPH_CENTER = 1
WD_ALIGN_PARAGRAPH_RIGHT = 2
WD_REPLACE_ALL = 2
WD_PRINT_VIEW = 3
WD_ORIENT_LANDSCAPE = 1
WD_LINE_SPACING_1_5 = 1


def main_flow():
    load_dotenv()
    api_key = os.getenv("ATRIA_API_KEY")
    assert api_key, "请在 .env 设置 ATRIA_API_KEY"

    # 连接到已运行的 Word，但强制新建空文档，绝不写入用户的活动文档
    try:
        app = win32com.client.GetObject(None, "Word.Application")
        print("已连接到现有的 Word 程序。")
    except Exception:
        app = win32com.client.Dispatch("Word.Application")
        print("启动了一个新的 Word 程序。")
    app.Visible = True
    doc = app.Documents.Add()
    sel = app.Selection
    print("已新建空文档用于测试（不动活动文档）。")

    writer = StreamingWriter(app, doc, sel, model=None, char_delay=0.01)
    model = DocModel(app, doc, sel, writer)
    writer.model = model
    main._model = model  # gen_code 读取结构地图用的全局模型

    # --- 阶段一：真实 SSE 流式生成 + 动画写入 ---
    from app.agent import WRITER_SYSTEM
    prompt = "写一首关于秋天的小诗：一个一级标题，下面两个段落，第二个段落里要有粗体。"
    system_prompt = WRITER_SYSTEM
    print("\n正在请求 Atria 流式生成...")
    received = 0
    for piece in ai_client.ai_stream(prompt, api_key, system_prompt):
        if not piece:
            continue
        received += 1
        print(piece, end="", flush=True)
        writer.feed(piece)
    writer.flush()
    print()
    assert received > 0, "流式接口没有收到任何内容"
    content = doc.Content.Text
    assert "#" not in content and "**" not in content, f"markdown 标记残留: {content!r}"
    kinds = [b.kind for b in model.blocks]
    assert "heading1" in kinds, f"标题未识别: {kinds}"
    print(f"ok: 真实流式写入（{received} 个片段，块类型: {kinds}）")

    print("\nblock_map:\n" + model.block_map())

    # 直接用 GUI/CLI 共享的真实执行环境（build_exec_globals），
    # 保证测试环境与生产环境一致，避免 AI 用到测试 env 里没有的符号
    from app.agent import build_exec_globals
    from session import Session
    exec_globals = build_exec_globals(app, doc, sel, model, writer, Session())

    # --- 阶段二：AI 代码排版（样式修改） ---
    fmt = "将所有一级标题居中，并把正文段落首行缩进 2 个字符"
    print(f"\n排版指令: {fmt}")
    code = main.gen_code(fmt, api_key)
    assert code, "AI 未生成排版代码"
    ok, err = run_code(code, exec_globals)
    assert ok, f"排版代码执行失败: {err}"
    print("ok: AI 排版代码（样式修改）执行完毕")

    # --- 阶段二（续）：AI 块编辑 ---
    fmt2 = "在第 0 块结尾追加一句：落叶满天飞"
    print(f"\n编辑指令: {fmt2}")
    code2 = main.gen_code(fmt2, api_key)
    assert code2, "AI 未生成编辑代码"
    ok2, err2 = run_code(code2, exec_globals)
    assert ok2, f"编辑代码执行失败: {err2}"
    assert "落叶满天飞" in doc.Content.Text, "块编辑后文档里没有追加的文字"
    print("ok: AI 块编辑（改写/插入）真实生效")

    print("\n真实端到端测试全部通过：流式 -> 写入 -> 排版 -> 块编辑。")
    print("测试文档保持打开，可手动检查后关闭（不保存）。")


if __name__ == "__main__":
    main_flow()
