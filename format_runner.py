"""逐语句可视化执行 AI 生成的排版代码。

不再黑盒 exec 整段代码：每条语句执行前在终端打印出来，执行后标注结果；
如果某条语句改变了选区位置，就把它滚动进视野并闪烁高亮。
排版过程从"执行完毕才知道结果"变成"一步步可观看"。
"""
import ast
import time

WD_NO_HIGHLIGHT = 0
WD_YELLOW = 7


def flash_range(app, rng, seconds=0.4):
    """让一个区域滚动进视野并短暂高亮。"""
    if rng is None:
        return
    try:
        rng.Select()
        rng.HighlightColorIndex = WD_YELLOW
        time.sleep(seconds)
        rng.HighlightColorIndex = WD_NO_HIGHLIGHT
    except Exception:
        pass


def flash_selection(app, seconds=0.4):
    try:
        flash_range(app, app.Selection.Range, seconds)
    except Exception:
        pass


def run_code(code, exec_globals):
    """逐语句执行 AI 生成的代码，返回 (是否成功, 错误信息)。

    成功后 exec_globals 里可能留下语句产生的副作用（样式修改等）。
    """
    code = (code or "").strip()
    if not code:
        return False, "AI 返回了空代码"
    try:
        tree = ast.parse(code)
    except SyntaxError as e:
        print(f"  [语法错误] {e}")
        return False, f"SyntaxError: {e}"

    statements = tree.body
    total = len(statements)
    app = exec_globals.get("word_app")
    for idx, node in enumerate(statements, 1):
        src = ast.unparse(node)
        print(f"[{idx}/{total}] {src}")
        before = None
        if app is not None:
            try:
                before = (app.Selection.Start, app.Selection.End)
            except Exception:
                before = None
        try:
            exec(compile(ast.Module(body=[node], type_ignores=[]), "<ai>", "exec"),
                 exec_globals)
            print("    [ok]")
        except Exception as e:
            print(f"    [fail] {e}")
            return False, f"{type(e).__name__}: {e}"
        # 可视化：该语句移动了选区 => 滚动到选区并闪烁
        if app is not None and before is not None:
            try:
                after = (app.Selection.Start, app.Selection.End)
                if after != before:
                    flash_selection(app)
            except Exception:
                pass
    return True, None
