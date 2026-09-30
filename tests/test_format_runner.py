"""离线测试 format_runner：逐语句执行、错误隔离、可视化钩子。"""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from format_runner import run_code


class FakeSel:
    def __init__(self):
        self.pos = 0

    @property
    def Start(self):
        return self.pos


class FakeApp:
    def __init__(self):
        self.Selection = FakeSel()


def test_ok():
    app = FakeApp()
    g = {"word_app": app, "log": []}
    ok, err = run_code("log.append(1)\nlog.append(2)", g)
    assert ok and err is None and g["log"] == [1, 2], (ok, err, g["log"])
    print("ok: run_code 正常逐语句执行")


def test_error_halts():
    app = FakeApp()
    g = {"word_app": app, "log": []}
    ok, err = run_code("log.append(1)\nraise ValueError('boom')\nlog.append(3)", g)
    assert not ok and "boom" in err, (ok, err)
    assert g["log"] == [1]  # 第三条没有执行
    print("ok: 出错语句之后不再执行，错误信息返回")


def test_syntax_error():
    app = FakeApp()
    g = {"word_app": app}
    ok, err = run_code("def f(:\n    pass", g)
    assert not ok and "SyntaxError" in err, (ok, err)
    print("ok: 语法错误安全捕获")


def test_empty():
    ok, err = run_code("", {})
    assert not ok and err == "AI 返回了空代码"
    print("ok: 空代码处理")


if __name__ == "__main__":
    test_ok()
    test_error_halts()
    test_syntax_error()
    test_empty()
    print("\nformat_runner 测试全部通过。")

def test_sandbox_rejects_import():
    g = {"word_app": None}
    ok, err = run_code("import os", g)
    assert not ok and "SandboxError" in err, (ok, err)
    print("ok: 沙箱拒绝 import")


def test_sandbox_rejects_while():
    g = {"word_app": None, "x": 0}
    ok, err = run_code("while x < 5:\n    x += 1", g)
    assert not ok and "SandboxError" in err, (ok, err)
    print("ok: 沙箱拒绝 while 无界循环")


def test_sandbox_rejects_dunder():
    g = {"word_app": None}
    ok, err = run_code("a = 1\nprint(a.__class__)", g)
    assert not ok and "SandboxError" in err, (ok, err)
    print("ok: 沙箱拒绝双下划线属性")


def test_sandbox_guard_limit():
    app = FakeApp()
    g = {"word_app": app, "log": []}
    ok, err = run_code("for i in range(1000000):\n    log.append(i)", g)
    assert not ok and "SandboxError" in err and "上限" in err, (ok, err)
    assert len(g["log"]) < 100000, "护栏应该在有限步内掐死循环"
    print(f"ok: 循环步数护栏（{len(g['log'])} 步后终止）")

