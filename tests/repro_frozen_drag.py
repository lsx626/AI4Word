# -*- coding: utf-8 -*-
"""驱动 frozen 窗口化 AI4Word.exe，复现「输出中拖动/大/小化窗口 -> 闪退」。

用法（仓库根目录）：
    .venv/python.exe -u tests/repro_frozen_drag.py [old|new|olddev|newdev]
      old     -> D:\\Program Files\\AI4Word\\AI4Word.exe（用户实际在跑的安装版 9.2）
      new     -> dist\\AI4Word\\AI4Word.exe（当前代码打包）
      olddev  -> git worktree 的 v9.2 源码，控制台 python 跑（traceback 可见）
                 需先 git worktree add D:\\Projects\\AI4Word-92 v9.2
      newdev  -> 当前仓库源码，控制台 python 跑
    不传参数默认 new

为什么这样驱动：窗口化 PyInstaller 应用的 C 层 fprintf(stderr) 在控制台看不到，
但继承的文件句柄照样能收到 -> 启动子进程时把 stdout/stderr 重定向到日志。
交互全部用真实 SendInput（Ctrl+Alt+Space 召唤 -> 输入英文写作 -> Enter ->
写作进行中按住头像拖窗口 / 点头像展开收起 / WIN+D 最小化还原），
完全按用户的真实操作路径（头像是拖动柄，点头像是展开/收起 = 大/小化）。
"""
import ctypes
import os
import subprocess
import sys
import time
import uuid
from ctypes import wintypes
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
os.chdir(REPO)

TARGETS = {
    "old": Path(r"D:\Program Files\AI4Word\AI4Word.exe"),
    "new": Path(r"D:\Projects\AI4Word\dist\AI4Word\AI4Word.exe"),
    "olddev": Path(r"D:\Projects\AI4Word-92"),
    "newdev": Path(r"D:\Projects\AI4Word"),
}
which = sys.argv[1] if len(sys.argv) > 1 else "new"
target = TARGETS[which]
exe = target
if which in ("olddev", "newdev"):
    cmd = [r"D:\Projects\AI4Word\.venv\python.exe", "-u",
           "ai4word.pyw", "-debug"]
else:
    cmd = [str(target), "-debug"]
assert exe.exists(), "target not found: %s" % exe

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

INPUT_MOUSE = 0
INPUT_KEYBOARD = 1
KEYEVENTF_UNICODE = 0x04
KEYEVENTF_KEYUP = 0x02
MOUSEEVENTF_MOVE = 0x01
MOUSEEVENTF_LEFTDOWN = 0x02
MOUSEEVENTF_LEFTUP = 0x04
MOUSEEVENTF_ABSOLUTE = 0x8000

VK_RETURN = 0x0D
VK_CONTROL = 0x11
VK_MENU = 0x12
VK_SPACE = 0x20
VK_LWIN = 0x5B
VK_D = 0x44
VK_UP = 0x26


class KEYBDINPUT(ctypes.Structure):
    _fields_ = [("wVk", wintypes.WORD), ("wScan", wintypes.WORD),
                ("dwFlags", wintypes.DWORD), ("time", wintypes.DWORD),
                ("dwExtraInfo", ctypes.c_size_t)]


class MOUSEINPUT(ctypes.Structure):
    _fields_ = [("dx", ctypes.c_long), ("dy", ctypes.c_long),
                ("mouseData", wintypes.DWORD), ("dwFlags", wintypes.DWORD),
                ("time", wintypes.DWORD), ("dwExtraInfo", ctypes.c_size_t)]


class INPUT_UNION(ctypes.Union):
    _fields_ = [("ki", KEYBDINPUT), ("mi", MOUSEINPUT)]


class INPUT(ctypes.Structure):
    _fields_ = [("type", wintypes.DWORD), ("union", INPUT_UNION)]


def send_input(*inputs):
    n = len(inputs)
    arr = (INPUT * n)(*inputs)
    user32.SendInput(n, ctypes.byref(arr), ctypes.sizeof(INPUT))


def key_vk(vk, up=False):
    i = INPUT()
    i.type = INPUT_KEYBOARD
    i.union.ki.wVk = vk
    i.union.ki.dwFlags = KEYEVENTF_KEYUP if up else 0
    return i


def type_char(ch):
    down = INPUT()
    down.type = INPUT_KEYBOARD
    down.union.ki.wScan = ord(ch)
    down.union.ki.dwFlags = KEYEVENTF_UNICODE
    up = INPUT()
    up.type = INPUT_KEYBOARD
    up.union.ki.wScan = ord(ch)
    up.union.ki.dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP
    return down, up


def do_type(text):
    seq = []
    for ch in text:
        seq.extend(type_char(ch))
    send_input(*seq)


def do_key(*vks):  # 按下后依次松开（含修饰键组合）
    down = [key_vk(v) for v in vks]
    up = [key_vk(v, up=True) for v in reversed(vks)]
    send_input(*(down + up))


def mouse_at(x, y):
    user32.SetCursorPos(x, y)
    time.sleep(0.02)


def mouse_button(down):
    i = INPUT()
    i.type = INPUT_MOUSE
    i.union.mi.dwFlags = (MOUSEEVENTF_LEFTDOWN if down else MOUSEEVENTF_LEFTUP)
    send_input(i)
    time.sleep(0.05)


def screen_size():
    return user32.GetSystemMetrics(0), user32.GetSystemMetrics(1)


def find_window(pid):
    found = {}

    def cb(hwnd, _):
        pid_box = wintypes.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid_box))
        if pid_box.value == pid and user32.IsWindowVisible(hwnd):
            found["hwnd"] = hwnd
            return False
        return True

    EnumWindowsProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    user32.EnumWindows(EnumWindowsProc(cb), 0)
    hwnd = found.get("hwnd")
    if not hwnd:
        return None
    r = wintypes.RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(r))
    return hwnd, (r.left, r.top, r.right - r.left, r.bottom - r.top)


def p(*a):
    print(*a, flush=True)


PROMPT = ("Write a poem about autumn in English: one h1 heading and two short "
          "paragraphs, at least 300 words.")


def main():
    stamp = "repro-frozen-" + uuid.uuid4().hex[:8]
    run_dir = REPO / "tests" / "deep_runs" / stamp
    run_dir.mkdir(parents=True, exist_ok=True)
    env = dict(os.environ)
    env["AI4WORD_DEBUG_LOG"] = str(run_dir / "debug.log")
    env["PYTHONUTF8"] = "1"
    # 旧版（v9.2）app 不支持 -debug，也没有 dotenv；直接把主仓库 .env 的
    # ATRIA_API_KEY 注入环境，保证旧/新版都能真链路写作
    env_path = REPO / ".env"
    if env_path.exists():
        for line in env_path.read_text(encoding="utf-8", errors="replace").splitlines():
            if line.strip().startswith("ATRIA_API_KEY") and "=" in line:
                env["ATRIA_API_KEY"] = line.split("=", 1)[1].strip().strip('"\' ')

    out_log = open(run_dir / "console.log", "wb")
    p("[repro] target=%s" % exe)
    p("[repro] cmd=%s" % " ".join(cmd))
    p("[repro] log dir: %s" % run_dir)

    # 启动前先探测单实例锁：若已有 AI4Word 在跑，新进程会弹
    # "已经在运行了" 的 QMessageBox 后 return 1，被误判成崩溃
    try:
        import ctypes as _ct
        _qshm_probe = _ct.windll.kernel32.CreateFileMappingW(
            _ct.c_void_p(-1), None, 0x04, 0, 1,
            "AI4Word-SingleInstance-v8")
        already = _qshm_probe == 0 or _ct.windll.kernel32.GetLastError() == 183
        if _qshm_probe:
            _ct.windll.kernel32.CloseHandle(_qshm_probe)
        if already:
            p("[repro] !!! AI4Word 单实例锁已被占用：请先退出正在运行的 AI4Word")
            return 5
    except Exception as e:
        p("[repro] 单实例探测异常（继续）: %s" % e)

    t0 = time.time()
    proc = subprocess.Popen(
        cmd,
        cwd=str(exe if which in ("olddev", "newdev") else exe.parent),
        env=env,
        stdin=subprocess.DEVNULL,
        stdout=out_log,
        stderr=subprocess.STDOUT,
        creationflags=subprocess.CREATE_NEW_PROCESS_GROUP,
    )
    p("[repro] pid=%d" % proc.pid)

    # 等窗口出现
    win = None
    for _ in range(60):
        time.sleep(0.5)
        if proc.poll() is not None:
            p("[repro] 进程在窗口出现前就退出了 rc=%s" % proc.returncode)
            report(run_dir, proc)
            return 3
        win = find_window(proc.pid)
        if win:
            break
    if not win:
        p("[repro] 60s 内未找到窗口，放弃")
        proc.kill()
        return 3
    hwnd, rect = win
    p("[repro] window hwnd=0x%x rect=%s (up %.1fs)" % (hwnd, rect, time.time() - t0))

    # 召唤并把输入聚焦
    do_key(VK_CONTROL, VK_MENU, VK_SPACE)
    time.sleep(0.8)
    win = find_window(proc.pid)
    if win:
        rect = win[1]

    phase = "typing"
    try:
        mouse_at(rect[0] + rect[2] // 2, rect[1] + rect[3] - 20)
        mouse_button(True)
        mouse_button(False)
        time.sleep(0.3)
        do_type(PROMPT)
        time.sleep(0.3)
        do_key(VK_RETURN)
        p("[repro] 已发送写作指令")

        cycle = 0
        end = time.time() + 75
        while time.time() < end:
            if proc.poll() is not None:
                phase = "died"
                break
            cycle += 1
            win = find_window(proc.pid)
            if not win:
                phase = "window-gone"
                p("[repro] 窗口消失（可能最小化或已死）")
                # 召唤回来
                do_key(VK_CONTROL, VK_MENU, VK_SPACE)
                time.sleep(0.8)
                win = find_window(proc.pid)
                if not win:
                    continue
            rect = win[1]
            x, y, w, h = rect

            # 头像拖动（用户拖窗口的真实路径）
            phase = "avatar-drag r%d" % cycle
            p("[repro] %s rect=%s" % (phase, rect))
            ax, ay = x + 20, y + h // 2
            mouse_at(ax, ay)
            mouse_button(True)
            for k in range(5):
                if proc.poll() is not None:
                    p("[repro] !!! 拖动中进程死亡")
                    phase = "died-in-avatar-drag r%d step %d" % (cycle, k)
                    break
                mouse_at(ax + 45 * (k + 1), ay + 22 * (k + 1))
                time.sleep(0.15)
            mouse_button(False)
            if proc.poll() is not None:
                break

            # 点头像 -> 展开（=用户的"大化"）
            phase = "avatar-click-expand r%d" % cycle
            p("[repro] %s" % phase)
            win = find_window(proc.pid)
            if win:
                x, y, w, h = win[1]
                mouse_at(x + 20, y + 20)
                mouse_button(True)
                mouse_button(False)
                time.sleep(0.45)
                # 展开态再拖一次
                win = find_window(proc.pid)
                if win:
                    x, y, w, h = win[1]
                    mouse_at(x + 20, y + 20)
                    mouse_button(True)
                    for k in range(4):
                        mouse_at(x + 20 - 35 * (k + 1), y + 20 + 15 * (k + 1))
                        time.sleep(0.12)
                    mouse_button(False)
                    time.sleep(0.3)
                # 再点回头像 -> 收起（=用户的"小化"）
                win = find_window(proc.pid)
                if win:
                    x, y, w, h = win[1]
                    mouse_at(x + 20, y + 20)
                    mouse_button(True)
                    mouse_button(False)
                    time.sleep(0.4)

            # WIN+D 最小化全部 / 还原
            phase = "win+d r%d" % cycle
            p("[repro] %s" % phase)
            do_key(VK_LWIN, VK_D)
            time.sleep(0.6)
            if proc.poll() is not None:
                p("[repro] !!! WIN+D 后进程死亡")
                phase = "died-in-win+d r%d" % cycle
                break
            do_key(VK_LWIN, VK_D)
            time.sleep(0.5)

            # WIN+UP（最大化若系统允许）
            phase = "win+up r%d" % cycle
            do_key(VK_LWIN, VK_UP)
            time.sleep(0.5)
            do_key(VK_CONTROL, VK_MENU, VK_SPACE)  # 召唤保持窗口可见
            time.sleep(0.4)
    finally:
        alive = proc.poll() is None
        p("[repro] phase=%s alive=%s" % (phase, alive))
        if alive:
            time.sleep(1.0)
            try:
                proc.kill()
            except Exception:
                pass
        report(run_dir, proc, phase)
    return 0 if proc.poll() is None else 4


def report(run_dir, proc, phase=""):
    p("[repro] ============ 结果 ============")
    p("[repro] exit code: %s | phase: %s" % (proc.returncode, phase))
    console = run_dir / "console.log"
    if console.exists():
        data = console.read_bytes()
        p("[repro] console(stdout/stderr) %d bytes:" % len(data))
        try:
            txt = data.decode("utf-8", "replace")
        except Exception:
            txt = data.decode("latin-1", "replace")
        print(txt[:4000])
        p("[repro] ---- console end ----")
    db = run_dir / "debug.log"
    if db.exists():
        lines = db.read_text(encoding="utf-8", errors="replace").splitlines()
        p("[repro] debug.log %d lines, tail:" % len(lines))
        for line in lines[-25:]:
            p("  " + line[:170])
        p("[repro] ---- debug.log end ----")


if __name__ == "__main__":
    sys.exit(main())
