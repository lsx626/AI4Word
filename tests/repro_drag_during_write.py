# -*- coding: utf-8 -*-
"""复现用户反馈：向 Word 输出进行中，拖动或大/小化窗口 -> 进程闪退。

前台跑（控制台进程，CPython/Qt 的 fatal 消息直接打到 stderr）：
    .venv/python.exe -u tests/repro_drag_during_write.py

长文写作（slow 档拉长流式）进行中，轮番轰炸各类窗口交互：
拖动 / 吸附 / 展开收起（大/小化）/ 最小化-还原 / 最大化-还原 /
resize / summon / 隐藏-显示 / 模式切换。
"""
import os
import subprocess
import sys
import traceback
import uuid
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
os.chdir(REPO)
sys.path.insert(0, str(REPO))
sys.path.insert(0, str(REPO / "tests"))

stamp = "repro-" + uuid.uuid4().hex[:8]
run_dir = REPO / "tests" / "deep_runs" / stamp
run_dir.mkdir(parents=True, exist_ok=True)
os.environ["AI4WORD_DEBUG_LOG"] = str(run_dir / "debug.log")
os.environ["AI4WORD_SHM_KEY"] = "AI4Word-SingleInstance-v8-" + stamp
os.environ["AI4WORD_TEST_SETTINGS"] = str(run_dir / "settings.json")
os.environ["AI4WORD_TEST_FORCE_NEW_WORD"] = "1"
os.environ["QT_QPA_PLATFORM"] = "windows"
sys.argv = [sys.argv[0], "-debug"]

LONG_PROMPT = ("写一篇关于秋天的长篇抒情散文，至少 1000 字，"
               "要有一级标题、多个二级标题、若干段落和一个列表。")


def p(*a):
    print(*a, flush=True)


def winword_pids():
    out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE", "/FO", "CSV", "/NH"],
                         capture_output=True, text=True).stdout
    pids = []
    for line in out.splitlines():
        parts = [x.strip('"') for x in line.split('","')]
        if len(parts) >= 2 and parts[0].upper() == "WINWORD.EXE":
            try:
                pids.append(int(parts[1]))
            except ValueError:
                pass
    return pids


before_winword = set(winword_pids())

from dotenv import load_dotenv  # noqa: E402
load_dotenv(str(REPO / ".env"))

from PySide6.QtCore import QPoint, Qt  # noqa: E402
from PySide6.QtTest import QTest  # noqa: E402

import debug_walkthrough as W  # noqa: E402
from debug_walkthrough import pump, wait_busy, wait_for  # noqa: E402


def teardown(worker):
    try:
        worker.choose("keep")
    except Exception:
        pass
    try:
        W.close_word(worker)
    except Exception:
        pass
    try:
        worker.send("quit")
        worker.wait(20000)
    except Exception:
        pass


def interact_round(window, worker, n):
    """一轮窗口交互轰炸；任一步崩溃进程即死，外层看不到下一行日志。"""
    base = window.pos()
    p("[repro] r%d 拖动 x6" % n)
    for i in range(6):
        if not worker._busy:
            break
        window.drag_to(base + QPoint(70 * (i + 1), 36 * (i + 1)), QPoint(0, 0))
        pump(0.2)
    p("[repro] r%d 吸附" % n)
    if worker._busy:
        window.snap_to_edge()
        pump(0.3)
    p("[repro] r%d 展开（大）" % n)
    if worker._busy:
        window.set_expanded(True)
        pump(0.5)
    p("[repro] r%d 收起（小）" % n)
    if worker._busy:
        window.set_expanded(False)
        pump(0.5)
    p("[repro] r%d 最小化-还原" % n)
    if worker._busy:
        window.showMinimized()
        pump(0.5)
    if not window.isVisible():
        window.showNormal()
        pump(0.5)
    p("[repro] r%d 最大化-还原" % n)
    if worker._busy:
        window.showMaximized()
        pump(0.5)
        if worker._busy:
            window.showNormal()
            pump(0.5)
    p("[repro] r%d resize" % n)
    if worker._busy:
        window.resize(window.width() + 90, window.height() + 70)
        pump(0.3)
        if worker._busy:
            window.resize(window.width() - 60, window.height() - 40)
            pump(0.3)
    p("[repro] r%d summon / 隐藏-显示" % n)
    if worker._busy:
        window.summon()
        pump(0.3)
        if worker._busy:
            window.hide()
            pump(0.4)
            if not window.isVisible():
                window.summon()
                pump(0.3)
    p("[repro] r%d 切模式" % n)
    if worker._busy:
        window._mode_group.button(1).click()
        pump(0.2)
        window._mode_group.button(0).click()
        pump(0.2)


def main():
    p("[repro] assemble app (真实 Word / 真实平台)")
    W.assemble()
    worker = W.S["worker"]
    window = W.S["window"]
    worker.start()
    pump(0.5)

    for rnd in range(2):
        p("[repro] === 写作轮 %d（slow 长文）===" % rnd)
        window.set_expanded(False)
        pump(0.2)
        window._on_speed(0)  # slow：流式拉长
        pump(0.2)
        window.input_c.setText(LONG_PROMPT)
        QTest.keyClick(window.input_c, Qt.Key_Return)
        if not wait_busy(worker, 60):
            p("[repro] 未进入 busy，环境问题，退出")
            teardown(worker)
            return 2
        # 等流式开始
        pump(3.0)
        n = 0
        while worker._busy:
            n += 1
            try:
                interact_round(window, worker, n)
            except Exception:
                p("[repro] 交互异常：")
                traceback.print_exc()
            if n >= 8:
                break
        ok = wait_for(lambda: not worker._busy, 240)
        p("[repro] 轮 %d write finished: %s（交互 %d 轮）" % (rnd, ok, n))
        pump(1.0)

    teardown(worker)
    p("[repro] === 全部存活 ===")
    return 0


if __name__ == "__main__":
    rc = main()
    for pid in set(winword_pids()) - before_winword:
        try:
            subprocess.run(["taskkill", "/PID", str(pid), "/F"],
                           capture_output=True, timeout=10)
            p("[repro] 清理 Word pid=%d" % pid)
        except Exception:
            pass
    sys.exit(rc)
