"""逐语句可视化执行 AI 生成的排版代码。

不再黑盒 exec 整段代码：每条语句执行前在终端打印出来，执行后标注结果；
如果某条语句改变了选区位置，就把它滚动进视野并闪烁高亮。
排版过程从"执行完毕才知道结果"变成"一步步可观看"。

安全：AI 生成的代码先过沙箱静态检查（sandbox.check：禁止 import / while /
with / 双下划线属性 / 危险内建），for 循环注入步数护栏（__guard_step），
把死循环与危险调用掐死在执行之前 / 可承受范围内。
"""
import ast
import time

from sandbox import check, instrument, make_guard, SandboxError

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


def run_code(code, exec_globals, sink=print):
    """逐语句执行 AI 生成的代码，返回 (是否成功, 错误信息)。

    沙箱不过则拒绝执行；通过后每条语句执行前注入一次新的步数守卫，
    保证任何一条 for 循环都不能无限转下去。成功后 exec_globals 里可能
    留下语句产生的副作用（样式修改等）。
    """
    code = (code or "").strip()
    if not code:
        return False, "AI 返回了空代码"
    try:
        tree = ast.parse(code)
    except SyntaxError as e:
        sink(f"  [语法错误] {e}")
        return False, f"SyntaxError: {e}"
    try:
        check(tree)
    except SandboxError as e:
        sink(f"  [沙箱拦截] {e}")
        return False, f"SandboxError: {e}"
    guarded = instrument(tree)

    statements = tree.body
    total = len(statements)
    app = exec_globals.get("word_app")
    for idx, (node, gnode) in enumerate(zip(statements, guarded.body), 1):
        src = ast.unparse(node)
        sink(f"[{idx}/{total}] {src}")
        before = None
        if app is not None:
            try:
                before = (app.Selection.Start, app.Selection.End)
            except Exception:
                before = None
        try:
            exec_globals["__guard_step"] = make_guard()
            exec(compile(ast.Module(body=[gnode], type_ignores=[]), "<ai>", "exec"),
                 exec_globals)
            sink("    [ok]")
        except SandboxError as e:
            sink(f"    [沙箱拦截] {e}")
            return False, f"SandboxError: {e}"
        except Exception as e:
            sink(f"    [fail] {e}")
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
