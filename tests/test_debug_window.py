# -*- coding: utf-8 -*-
"""离线回归：调试日志毫秒分辨率与 deep_test 事件窗口的对齐。

根因（真机日志验证，见 tests/deep_runs/20261005-213146）：app.debug._emit
把时间戳截断到毫秒（地板），而 time.time() 是微秒精度；run_scenario 以
微秒级 time.time() 作窗口起点，场景第一个动作恰好落在与起点同一毫秒时，
其日志时间戳 floor 后小于起点（如 14.759 < 14.7594），被 _parse_window
误剔除——系统性吃掉每个场景的首事件：

- speed_gears         首个 speed_clicked 21:31:49.768（报 speed_clicked(2/3)）
- arrange_ok          首个 mode_clicked  21:32:10.795（报 mode_clicked(1/2)）
- review_presets_save 首个 review_toggled 21:32:14.759（报 review_toggled(1/2)）

修复回归两处：
1) tests/deep_log.py（从 deep_test.py 抽出）：_parse_window 的 start 先经
   _ms_floor 向下取整到毫秒，对齐日志分辨率；同毫秒首事件不再丢失，起点
   前一毫秒的事件仍排除。窗口解析放进纯 stdlib 模块，是因为 deep_test.py
   本身不能被离线套件 import（它导入即锁定 QT_QPA_PLATFORM=windows 并向
   sys.argv 追加 -debug）。
2) app/debug.py：_emit 的秒与毫秒取自同一次时钟读取——原来
   strftime(localtime()) 与 time.time() 是两次独立调用，跨越整秒边界时
   时间戳会倒退最多 1 秒（同样会把事件甩出窗口），这里用确定的
   「跨秒时钟」复现并锁定。

三入口均可：
    PYTHONPATH=. python -u tests/run_offline.py
    PYTHONPATH=. python tests/test_debug_window.py
    PYTHONPATH=. python -m pytest tests/test_debug_window.py
"""
import os
import sys
import time

_TESTS = os.path.dirname(os.path.abspath(__file__))
_REPO = os.path.dirname(_TESTS)
sys.path.insert(0, _TESTS)
sys.path.insert(0, _REPO)

from app import debug          # noqa: E402  -- 纯 stdlib 模块，无 Qt / COM
import deep_log                # noqa: E402  -- deep_test 的日志窗口解析


# 真机日志里的事件时间（2026-10-05 21:32:14.759 = review_presets_save 首事件）
_PREV_MS = "2026-10-05 21:32:14.758 [MainThread] INFO speed_set gear=slow"
_FIRST = "2026-10-05 21:32:14.759 [MainThread] INFO review_toggled on=true"
_SECOND = "2026-10-05 21:32:15.003 [MainThread] INFO review_toggled on=false"
_NEXT_MS = "2026-10-05 21:32:14.760 [MainThread] INFO mode_clicked mode=write"

BASE = deep_log._ts("2026-10-05 21:32:14.759")   # ...14.759 的 epoch（时区无关）
SECOND_TS = deep_log._ts("2026-10-05 21:32:15.003")


def _scratch_dir():
    """日志目录：TEMP 根下可直接写（与 test_debug_mode 一致）。"""
    base = os.environ.get("TEMP") or os.path.expanduser("~") or "."
    probe = os.path.join(base, "deep_window_test.log")
    try:
        with open(probe, "a", encoding="utf-8"):
            pass
        return base
    except OSError:
        pass
    d = os.path.join(_TESTS, "_debug_scratch")
    try:
        os.makedirs(d, exist_ok=True)
    except OSError:
        pass
    return d


class _ScratchLog:
    """把 debug.log_path() 指向写好 lines 的临时文件；退出还原 _state。

    不 monkeypatch、不 init，不碰 %APPDATA%\\AI4Word\\debug.log，用完
    原样还原（run_offline 顺序加载同进程的其它测试不受影响）。
    """

    def __init__(self, lines):
        self._lines = list(lines)

    def __enter__(self):
        self._saved = dict(debug._state)
        path = os.path.join(_scratch_dir(), "deep_window_test.log")
        with open(path, "w", encoding="utf-8") as f:
            f.write("".join(ln + "\n" for ln in self._lines))
        debug._state.update({"debug": True, "path": path})
        self.path = path
        return self

    def __exit__(self, *exc):
        debug._state.clear()
        debug._state.update(self._saved)
        return False


class _Patch:
    """临时替换对象属性（无 monkeypatch 依赖）。"""

    def __init__(self, obj, name, value):
        self._obj = obj
        self._name = name
        self._value = value

    def __enter__(self):
        self._old = getattr(self._obj, self._name)
        setattr(self._obj, self._name, self._value)
        return self

    def __exit__(self, *exc):
        setattr(self._obj, self._name, self._old)
        return False


class _RolloverClock:
    """模拟原来的两次系统时钟调用跨越整秒边界：localtime() 读到
    14.9999（秒为 14），随后的 time.time() 已进到下一秒 15.0001。"""

    def __init__(self, before, after):
        self._before = before
        self._after = after

    def localtime(self, secs=None):
        return time.localtime(self._before if secs is None else secs)

    def strftime(self, fmt, tup):
        return time.strftime(fmt, tup)

    def time(self):
        return self._after


def _names(evs):
    return [ev for (_lvl, ev, _f) in evs]


# ---------------------------------------------------------------- 窗口对齐

def test_first_event_in_same_millisecond_kept():
    """复现报错：首事件日志时间 14.759，场景起点 14.7594（同一毫秒、
    事件在起点之后）。修复前 ts < start 被剔除 -> MUST 缺失。"""
    with _ScratchLog([_FIRST]):
        evs = deep_log._parse_window(BASE + 0.0004, BASE + 5.0)
    assert _names(evs) == ["review_toggled"], _names(evs)


def test_first_event_pair_both_kept():
    """review_presets_save 的成对事件 14.759(on=true)/15.003(on=false)
    都落进窗口：review_toggled(2/2) 满足，不再报 (1/2)。"""
    with _ScratchLog([_FIRST, _SECOND]):
        evs = deep_log._parse_window(BASE + 0.0004, BASE + 5.0)
    assert deep_log._count(evs, "review_toggled") == 2


def test_previous_millisecond_still_excluded():
    """取整不能过头：起点 14.7594 -> 14.759，14.758 的旧事件仍排除。"""
    with _ScratchLog([_PREV_MS, _FIRST]):
        evs = deep_log._parse_window(BASE + 0.0004, BASE + 5.0)
    assert _names(evs) == ["review_toggled"], _names(evs)


def test_aligned_millisecond_start_unchanged():
    """起点本身是整毫秒时窗口行为不变：前一毫秒排除、本毫秒起收录。"""
    with _ScratchLog([_PREV_MS, _FIRST, _NEXT_MS]):
        evs = deep_log._parse_window(BASE, BASE + 1.0)
    assert _names(evs) == ["review_toggled", "mode_clicked"], _names(evs)


def test_end_boundary_inclusive():
    """t == end 仍在窗口内（run_scenario 的 end 本来就带 +2s 收尾尾巴）。

    end 直接用「事件时间戳本身」，排除浮点取整造成的边界漂移。"""
    with _ScratchLog([_FIRST, _SECOND]):
        evs = deep_log._parse_window(BASE, SECOND_TS)     # end 恰为 15.003
    assert _names(evs) == ["review_toggled", "review_toggled"], _names(evs)


def test_end_just_below_excludes():
    """end 的一毫秒之内的事件被排除（窗口上界不回退过头）。"""
    with _ScratchLog([_FIRST, _SECOND]):
        evs = deep_log._parse_window(BASE, SECOND_TS - 0.0005)   # 15.0025
    assert _names(evs) == ["review_toggled"], _names(evs)


def test_ms_floor_helper():
    assert deep_log._ms_floor(BASE) == BASE                      # 整毫秒不变
    assert deep_log._ms_floor(BASE + 0.0004) == BASE             # 同毫秒地板
    # 上一毫秒（浮点减法有 1 ulp 误差，按毫秒整数比较）
    assert abs(deep_log._ms_floor(BASE - 0.0004) - (BASE - 0.001)) < 1e-6
    assert deep_log._ms_floor(1.5004) == 1.5
    assert deep_log._ms_floor(1.5) == 1.5


# ---------------------------------------------------------------- 时间戳完整性

def test_emit_timestamp_no_cross_second_drift():
    """_emit 的秒与毫秒必须同源：跨秒边界时不能写出倒退 1 秒的时间戳。

    旧实现 strftime/localtime() 得到 21:32:14、time.time() 得 15 秒的毫秒
    -> "21:32:14.000"（倒退）；新实现取自同一次 -> "21:32:15.000"。"""
    with _ScratchLog([]) as scratch:
        with _Patch(debug, "time",
                    _RolloverClock(int(BASE) + 0.9999, int(BASE) + 1.0001)):
            debug.log("rollover_probe")
        with open(scratch.path, encoding="utf-8") as f:
            text = f.read()
    m = deep_log._LINE_RE.search(text)
    assert m, text
    assert m.group(1).endswith("21:32:15.000"), m.group(1)
    assert deep_log._ts(m.group(1)) == int(BASE) + 1.0


def test_emit_timestamp_within_call_bounds():
    """真实时钟下：写出的毫秒级时间戳落在 [t0, t1] 的毫秒地板区间内，
    且行能被 deep_log._LINE_RE 解析（emit 与窗口解析格式保持一致）。

    留半毫秒余量吸收浮点取整（1 ulp），真正要钉的是 1 秒级漂移。"""
    with _ScratchLog([]) as scratch:
        t0 = time.time()
        debug.log("ts_probe", stage="bounds")
        t1 = time.time()
        with open(scratch.path, encoding="utf-8") as f:
            text = f.read()
    line = [ln for ln in text.splitlines() if "ts_probe" in ln][0]
    m = deep_log._LINE_RE.match(line)
    assert m, line
    ts = deep_log._ts(m.group(1))
    lo = deep_log._ms_floor(t0) - 0.0005
    hi = deep_log._ms_floor(t1) + 0.0005
    assert lo <= ts <= hi, (ts, lo, hi)


if __name__ == "__main__":
    import inspect
    import traceback

    failed = 0
    tests = [(k, v) for k, v in sorted(globals().items())
             if k.startswith("test_") and callable(v)]
    for name, fn in tests:
        try:
            fn()
            print("ok   %s" % name)
        except Exception:
            failed += 1
            print("FAIL %s" % name)
            traceback.print_exc()
    print("%d failed, %d passed" % (failed, len(tests) - failed))
    sys.exit(1 if failed else 0)
