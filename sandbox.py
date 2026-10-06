"""AI 生成代码的静态安全检查与循环步数护栏。

run_code 直接执行 LLM 生成的 Python，逐语句可视化只能"看见"它在做什么，
挡不住真正的危险动作（文件、网络、import、死循环）。exec 前在这里做两件事：

1. 静态检查 check()：禁止 import / while / with / 双下划线属性 / 危险内建
   （open、exec、eval、getattr 等）、BaseException 家族（raise SystemExit
   会逃出 except Exception 杀死 worker 线程）、COM 自动化面的高危成员
   （word_app.Run 执行任意 VBA 宏）；合法的排版代码（属性赋值、块函数
   调用、有限 for 循环）不受影响。
2. 循环护栏 instrument()：给每个 for 循环体头部注入步数计数调用，
   run_code 注入的 __guard_step 共享同一个计数器，总步数超限时抛
   SandboxError，把死循环掐死在可承受的范围内。
"""
import ast

# 禁止直接调用的内建名（无 import 的环境下它们仍是危险面）
BANNED_NAMES = frozenset({
    "open", "exec", "eval", "compile", "__import__", "input", "breakpoint",
    "getattr", "setattr", "delattr", "globals", "locals", "vars", "help",
    "exit", "quit", "super", "memoryview", "iter", "next", "classmethod",
    "staticmethod",
    # BaseException 家族：raise SystemExit / KeyboardInterrupt 会穿过
    # run_code 的 except Exception，直接终结进程或 worker 线程
    "BaseException", "SystemExit", "KeyboardInterrupt", "GeneratorExit",
})

# 禁止的语句类型（import 链、无界循环、上下文管理器、显式作用域操纵）
BANNED_STMTS = (
    ast.Import, ast.ImportFrom, ast.While, ast.AsyncFunctionDef, ast.With,
    ast.AsyncWith, ast.Global, ast.Nonlocal, ast.AsyncFor, ast.Await,
    ast.Delete, ast.TryStar,
)

# COM 自动化面上的高危成员名：word_app.Run 能执行任意 VBA 宏（等于完全
# 机器控制）、System.PrivateProfileString 读写任意文件、Documents.Open
# 打开任意外部文档。AI 排版代码不需要它们
BANNED_ATTRS = frozenset({
    "Run", "System", "Shell", "WScript", "Documents", "Open",
})

MAX_STEPS = 50000  # for 循环总步数上限（排版脚本是有限的小循环）

# 给代码生成 / 自我修复提示词用的禁令清单（中文）：直接从上面的实际
# 禁令集合生成，sandbox 与提示词只有这一份真相，改一处不会漂移。
# fix_code 不带这份清单时，AI 的修复代码常以 `import ...` 开头，被
# check() 判「禁止的语句: Import」，第二段代码必死，arrange 从此
# 收敛不回 arrange_applied（gen_code 的 CODEGEN_SYSTEM 有同等提示）。
_BANNED_STMT_CN = {
    ast.Import: "import",
    ast.ImportFrom: "from ... import",
    ast.While: "while 无界循环",
    ast.With: "with",
    ast.AsyncWith: "async with",
    ast.AsyncFunctionDef: "async 函数定义",
    ast.AsyncFor: "async for",
    ast.Await: "await",
    ast.Global: "global",
    ast.Nonlocal: "nonlocal",
    ast.Delete: "del",
    ast.TryStar: "try*",
}

SANDBOX_RULES_TEXT = (
    "**沙箱禁令（违反任意一条，整段代码直接作废、不会执行）**:\n"
    "- 禁止语句："
    + "、".join(sorted(_BANNED_STMT_CN.get(t, t.__name__) for t in BANNED_STMTS))
    + "\n- 禁止任何双下划线名字（__x__、obj.__class__ 等；字符串常量里含"
    "双下划线也不行）\n"
    "- 禁止调用这些内建/名字：" + "、".join(sorted(BANNED_NAMES)) + "\n"
    "- 禁止调用这些 COM 成员：" + "、".join(sorted(BANNED_ATTRS)) + "\n"
    "- for 循环总步数上限 " + str(MAX_STEPS)
    + "（排版脚本足够；死循环会被掐死）\n"
    "- 只输出纯 Python 代码：不写 import、不定义函数/类、不重新声明"
    " word_app/doc/sel，不加任何 Markdown 标记或解释文字"
)


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
            if nm in BANNED_ATTRS:
                raise SandboxError(f"禁止调用 COM 成员: {nm}")
        # 字符串常量里的双下划线藏不住 AST：「{0.__globals__}」.format(x)
        # 的迷你语言支持属性/下标链，能拿到模块 internals 的 repr
        if isinstance(node, ast.Constant) and isinstance(node.value, str):
            if "__" in node.value:
                raise SandboxError("字符串常量含双下划线（疑似 format 链逃逸）")


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
