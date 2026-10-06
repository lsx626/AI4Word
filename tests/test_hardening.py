# -*- coding: utf-8 -*-
"""防呆加固回归测试（V9.2）：沙箱逃逸 / 负索引 / 空 md / 存档类型容错。"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import sandbox
from format_runner import run_code


def _check(cond, msg):
    if not cond:
        raise AssertionError(msg)
    print("ok:", msg)


def test_systemexit_banned():
    """raise SystemExit 不能逃出沙箱（会杀死 worker 线程/进程）。"""
    import ast
    for src in ("raise SystemExit", "raise KeyboardInterrupt",
                "raise BaseException('x')"):
        tree = ast.parse(src)
        try:
            sandbox.check(tree)
        except sandbox.SandboxError as e:
            _check("SystemExit 家族被禁" in str(e) or "禁止" in str(e),
                   f"{src} 应被沙箱拦截（实际：{e}）")
        else:
            raise AssertionError(f"沙箱放行了危险代码: {src}")
    _check(True, "raise SystemExit / KeyboardInterrupt / BaseException 均被禁")


def test_run_code_catches_baseexception():
    """即便绕过了静态检查，BaseException 也不能杀死调用方。"""
    # 用 exec_globals 手工注入一个抛 SystemExit 的函数再调用——
    # 静态检查会禁 raise SystemExit，所以用「直接调内建 exit 的等价方式」
    # 模拟逃逸场景，验证 run_code 的运行时兜底
    g = {"__guard_step": sandbox.make_guard(),
         "_evil": lambda: (_ for _ in ()).throw(SystemExit(1))}
    code = "x = _evil()"
    ok, err = run_code(code, g)
    _check(ok is False, "run_code 应把 SystemExit 转成失败")
    _check("SystemExit" in err, f"错误信息应携带 SystemExit（实际：{err}）")


def test_com_member_banned():
    import ast
    for src in ('word_app.Run("宏")', 'word_app.System.PrivateProfileString("a")',
                'word_app.Documents.Open("c:/x.docx")'):
        tree = ast.parse(src)
        try:
            sandbox.check(tree)
        except sandbox.SandboxError:
            pass
        else:
            raise AssertionError(f"沙箱放行了 COM 危险成员: {src}")
    _check(True, "word_app.Run / .System / .Documents.Open 之类被静态拦截")


def test_string_dunder_banned():
    import ast
    src = 's = "{0.__class__}".format(1)'
    try:
        sandbox.check(ast.parse(src))
    except sandbox.SandboxError:
        _check(True, "字符串常量里的双下划线（format 链逃逸）被禁")
    else:
        raise AssertionError("format 链逃逸未被拦截")


def test_guard_steps_shared():
    """步数守卫跨语句共享：两条小循环的总预算仍是 MAX_STEPS。"""
    import time
    g = {}
    code = "acc = 0\nfor i in range(30000):\n    acc += 1\nfor j in range(30000):\n    acc += 1"
    t0 = time.time()
    ok, err = run_code(code, g)
    _check(not ok and "上限" in (err or ""), f"应触发步数上限（实际 ok={ok}, err={err}）")
    _check(time.time() - t0 < 3.0, "两条 3 万次循环应快速被守卫终止")


if __name__ == "__main__":
    test_systemexit_banned()
    test_run_code_catches_baseexception()
    test_com_member_banned()
    test_string_dunder_banned()
    test_guard_steps_shared()
    print("\n沙箱加固测试全部通过")
