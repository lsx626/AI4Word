# -*- coding: utf-8 -*-
"""deep_test 的调试日志窗口解析（纯 stdlib + app.debug，无 Qt / COM 依赖）。

从 tests/deep_test.py 抽出，目的：deep_test.py 是真实驱动入口，导入即
锁定 QT_QPA_PLATFORM=windows 并向 sys.argv 追加 -debug，不能被离线测试套件
（run_offline）导入；本模块只依赖 os/re/datetime/math 与 app.debug，离线
回归测试可以直接 import，直接覆盖 _parse_window 的窗口语义（见
tests/test_debug_window.py）。

分辨率对齐（V9.6 修复的根因）：app.debug._emit 写日志时把时间戳截断到
毫秒（地板），而调用方拿到的是微秒精度的 time.time()。若窗口起点保留
亚毫秒尾巴，场景第一个动作恰好落在与起点同一毫秒时，其日志时间戳 floor
后小于起点（如 14.759 < 14.7594）会被窗口剔除——系统性吃掉每个场景的
首事件（speed_gears 的首个 speed_clicked、arrange_ok 的首个 mode_clicked、
review_presets_save 的首个 review_toggled 等 MUST 事件）。
"""
import datetime
import math
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import debug   # noqa: E402  -- 只用 log_path() 定位当前日志文件

# ---------------------------------------------------------------- 行解析

_LINE_RE = re.compile(
    r"^(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\.\d{3}) \[[^\]]*\] "
    r"(INFO|WARN|ERROR) (\S+)(?:\s+(.*))?$")


def _ts(text):
    try:
        return datetime.datetime.strptime(text, "%Y-%m-%d %H:%M:%S.%f").timestamp()
    except Exception:
        return None


def _ms_floor(t):
    """秒时间戳向下取整到毫秒，对齐 debug 日志的时间戳分辨率。

    app.debug._emit 只写毫秒（地板截断），而 time.time() 是微秒精度：
    若窗口起点保留亚毫秒尾巴，场景第一个动作恰好落在与起点同一毫秒时，
    其日志时间戳 floor 后会小于起点（如 14.759 < 14.7594）被窗口剔除，
    系统性地吃掉每个场景的首事件（speed_gears 的首个 speed_clicked、
    arrange_ok 的首个 mode_clicked、review_presets_save 的首个
    review_toggled 等 MUST 事件）。取整后只要事件真实发生在起点之后，
    其毫秒级日志时间戳就一定 >= 取整后的起点。
    """
    return math.floor(t * 1000) / 1000.0


def _read_log_lines():
    """读取当前 debug.log 全文（调试模块单线追写、每行独立 open/write，落读即所见）。"""
    path = debug.log_path()
    if not path or not os.path.exists(path):
        return []
    try:
        with open(path, encoding="utf-8", errors="replace") as f:
            return f.read().splitlines()
    except OSError:
        return []


def _parse_window(start, end):
    """返回 [(level, event, field_str)]：时间落在 [start, end] 的全部事件。

    start 先经 _ms_floor 向下取整到毫秒：日志本身只有毫秒分辨率，亚毫秒
    级的窗口起点与日志时间戳不是同一分辨率上的比较，会把与起点同一毫秒
    内发生的首事件误判到窗口之外。end 保留原值（日志 floor 只会把时间
    往前偏，尾巴方向只会多收不会少收）。
    """
    start = _ms_floor(start)
    out = []
    for ln in _read_log_lines():
        m = _LINE_RE.match(ln)
        if not m:
            continue
        t = _ts(m.group(1))
        if t is None or t < start or t > end:
            continue
        out.append((m.group(2), m.group(3), m.group(4) or ""))
    return out


def _count(events, name):
    return sum(1 for (_lvl, ev, _f) in events if ev == name)


def _field_int(events, name, field):
    """从某事件的字段里取整数（如 word_assembled 的 restored=）。"""
    for (_lvl, ev, f) in events:
        if ev != name:
            continue
        m = re.search(field + r"=(\-?\d+)", f)
        if m:
            try:
                return int(m.group(1))
            except Exception:
                pass
    return None
