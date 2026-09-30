"""AI 生成代码的静态安全检查与循环步数护栏。

run_code 直接执行 LLM 生成的 Python，逐语句可视化只能"看见"它在做什么，
挡不住真正的危险动作（文件、网络、import、死循环）。exec 前在这里做两件事：

1. 静态检查 check()：禁止 import / while / with / 双下划线属性 / 危险内建
   （open、exec、eval、getattr 等）；合法的排版代码（属性赋值、块函数调用、
   有限 for 循环）不受影响。
2. 循环护栏 instrument()：给每个 for 循环体头部注入步数计数调用，
   run_code 注入的 __guard_step 会在总步数超限时抛 SandboxError，
   把死循环掐死在可承受的范围内。
"""
import ast

# 禁止直接调用的内建名（无 import 的环境下它们仍是危险面）
BANNED_NAMES = frozenset({
    "open", "exec", "eval", "compile", "__import__", "input", "breakpoint",
    "getattr", "setattr", "delattr", "globals", "locals", "vars", "help",
    "exit", "quit", "super", "memoryview", "iter", "next", "classmethod",
    "staticmethod",
})

# 禁止的语句类型（import 链、无界循环、上下文管理器、显式作用域操纵）
BANNED_STMTS = (
    ast.Import, ast.ImportFrom, ast.While, ast.AsyncFunctionDef, ast.With,
    ast.AsyncWith, ast.Global, ast.Nonlocal, ast.AsyncFor, ast.Await,
    ast.Delete, ast.TryStar,
)

MAX_STEPS = 50000  # for 循环总步数上限（排版脚本是有限的小循环）


class SandboxError(Exception):
    """沙箱拦截到的危险代码。"""


def _names_of(node):
    """收集一个节点上出现的用户定义名（变量、属性、函数名、参数名）。"""
    out = []
    if isinstance(node, ast.Name):
        out.append(node.id)
    if isinstance(node, ast.Attribute):
        out.append(node.attr)
    if isinstance(node, (ast.FunctionDef, ast.ClassDef, ast.AsyncFunctionDef)):
        out.append(node.name)
        out.extend(a.arg for a in node.args.args)
        out.extend(a.arg for a in node.args.kwonlyargs)
    if isinstance(node, ast.arg):
        out.append(node.arg)
    if isinstance(node, ast.alias):
        out.append(node.name)
    if isinstance(node, ast.keyword):
        out.append(node.arg or "")
    return out


def check(tree):
    """静态安全检查：不通过时抛 SandboxError。"""
    for node in ast.walk(tree):
        if isinstance(node, BANNED_STMTS):
            raise SandboxError(f"禁止的语句: {type(node).__name__}")
        for nm in _names_of(node):
            if not nm:
                continue
            if "__" in nm:
                raise SandboxError(f"禁止双下划线名: {nm}")
            if nm in BANNED_NAMES:
                raise SandboxError(f"禁止调用内建: {nm}")


class _GuardInstrument(ast.NodeTransformer):
    """给每个 for 循环体头部注入 __guard_step() 调用。"""

    def visit_For(self, node):
        self.generic_visit(node)
        guard = ast.Expr(
            value=ast.Call(
                func=ast.Name(id="__guard_step", ctx=ast.Load()),
                args=[], keywords=[]))
        node.body.insert(0, guard)
        ast.fix_missing_locations(node)
        return node


def instrument(tree):
    """返回注入步数护栏后的新树（不修改原树）。"""
    import copy
    return _GuardInstrument().visit(copy.deepcopy(tree))


def make_guard(limit=MAX_STEPS):
    """生成 __guard_step：累计步数超限即抛 SandboxError。"""
    state = {"steps": 0}

    def guard():
        state["steps"] += 1
        if state["steps"] > limit:
            raise SandboxError(f"循环步数超过上限 {limit}")
    return guard