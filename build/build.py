# -*- coding: utf-8 -*-
"""一键打包：PyInstaller onedir +（若装了 Inno Setup）编译安装包。

用法（项目根目录）：
    .\.venv\Scripts\python.exe -u build\build.py

产物：
    dist\AI4Word\            解压即用的程序目录
    dist\AI4Word-Setup-x.y.z.exe   安装包（需要 Inno Setup 6）
"""
import io
import os
import shutil
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
os.chdir(ROOT)

from app import __version__  # noqa: E402

ENTRY = "ai4word.pyw"
NAME = "AI4Word"
ICON = os.path.join("assets", "app_icon.ico")
WORK = os.path.join("build", "_pyi")
DIST = "dist"

HIDDEN = [
    "pythoncom", "pywintypes", "win32com", "win32com.client",
    "win32timezone",  # pywin32 常见缺失
    "markdown_it", "dotenv",
    "app", "app.agent", "app.engine", "app.main_window",
]

# 开发态导入探针：用 Windows 加载器实测启动时真正需要的 DLL。
# 构建时已把含 icu*.dll 的目录从 PATH 剔除（避免精简版 ICU 串扰 Qt6Core），
# 连带 PyInstaller 收不进那些"只靠 PATH 才找得到"的依赖（如 anaconda
# Library\bin 里的 ffi-8.dll / libssl-3-x64.dll / libcrypto-3-x64.dll），
# 运行时 _ctypes、_ssl 等扩展会报"找不到指定的模块"。这里由加载器给出
# 真实缺口，打包后显式补进 _internal。
PROBE_SRC = r"""
import json, sys, ctypes
from ctypes import wintypes

def _warn(tag, e):
    print("probe-warn %s: %s" % (tag, e), file=sys.stderr)

for mod in ("pythoncom", "pywintypes", "win32com", "win32com.client",
           "win32com.shell", "ssl", "requests", "markdown_it", "dotenv"):
    try:
        __import__(mod)
    except Exception as e:
        _warn(mod, e)
try:
    from PySide6.QtCore import Qt          # noqa: F401
    from PySide6.QtWidgets import QApplication  # noqa: F401
    from PySide6.QtGui import QIcon        # noqa: F401
except Exception as e:
    _warn("pyside6", e)
try:
    import app.__main__                    # 启动时的完整导入图
except Exception as e:
    _warn("app", e)

psapi = ctypes.WinDLL("psapi.dll", use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
kernel32.GetCurrentProcess.restype = wintypes.HANDLE
psapi.EnumProcessModules.argtypes = [wintypes.HANDLE, ctypes.POINTER(wintypes.HMODULE), wintypes.DWORD, ctypes.POINTER(wintypes.DWORD)]
psapi.EnumProcessModules.restype = wintypes.BOOL
psapi.GetModuleFileNameExW.argtypes = [wintypes.HANDLE, wintypes.HMODULE, wintypes.LPWSTR, wintypes.DWORD]
psapi.GetModuleFileNameExW.restype = wintypes.DWORD
handle = kernel32.GetCurrentProcess()
arr = (wintypes.HMODULE * 2048)()
needed = wintypes.DWORD()
out = []
if psapi.EnumProcessModules(handle, arr, ctypes.sizeof(arr), ctypes.byref(needed)):
    count = min(needed.value // ctypes.sizeof(wintypes.HMODULE), 2048)
    for i in range(count):
        buf = ctypes.create_unicode_buffer(1024)
        if psapi.GetModuleFileNameExW(handle, arr[i], buf, 1024):
            out.append(buf.value)
print("PROBE_JSON " + json.dumps(out))
"""


ISCC_CANDIDATES = [
    r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
    r"C:\Program Files\Inno Setup 6\ISCC.exe",
    os.path.expandvars(r"%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe"),
]


def ensure_icon():
    if not os.path.exists(ICON):
        print("图标缺失，先运行 build/icon_gen.py ...")
        subprocess.check_call([sys.executable, "-u", "build/icon_gen.py"])


def find_iscc():
    for p in ISCC_CANDIDATES:
        if os.path.exists(p):
            return p
    where = shutil.which("ISCC")
    return where


def build_env():
    """给 PyInstaller 子进程一个干净的 PATH。

    打包机 PATH 里的外来 icuuc.dll（各运行时自带的 poppler / anacona ICU）
    会被 PyInstaller 当成 Qt6Core.dll 的依赖收进包；这些精简版 ICU 缺少
    ucnv_* 导出，运行时又遮蔽系统 system32\icuuc.dll，导致 Qt6Core 加载
    报“找不到指定的程序”。剔除含 icu*.dll 的目录与 codex 运行时缓存即可。
    """
    import glob as _glob
    env = os.environ.copy()
    kept, dropped = [], []
    for d in env.get("PATH", "").split(os.pathsep):
        if not d:
            continue
        low = d.lower()
        if "codex-runtimes" in low or _glob.glob(os.path.join(d, "icu*.dll")):
            dropped.append(d)
            continue
        kept.append(d)
    env["PATH"] = os.pathsep.join(kept)
    if dropped:
        print("构建 PATH 剔除 %d 个外来目录: %s" % (len(dropped), "; ".join(dropped)))
    return env


def run_pyinstaller():
    if os.path.isdir(os.path.join(WORK)):
        shutil.rmtree(WORK, ignore_errors=True)
    cmd = [sys.executable, "-m", "PyInstaller",
           "--noconfirm", "--windowed", "--onedir",
           "--name", NAME,
           "--icon", ICON,
           "--workpath", WORK,
           "--distpath", DIST,
           "--clean"]
    # 不要无谓的 tkinter 等膨胀，但保留 PySide6 插件分析
    cmd += ["--exclude-module", "tkinter"]
    for h in HIDDEN:
        cmd += ["--hidden-import", h]
    cmd.append(ENTRY)
    print("PyInstaller:", " ".join(cmd))
    subprocess.check_call(cmd, env=build_env())
    app_dir = os.path.join(DIST, NAME)
    if not os.path.isdir(app_dir):
        raise SystemExit("打包失败：未找到 " + app_dir)
    exe = os.path.join(app_dir, NAME + ".exe")
    total = sum(os.path.getsize(os.path.join(dp, f))
                for dp, dn, fn in os.walk(app_dir) for f in fn)
    print(f"程序目录: {app_dir} (exe {os.path.getsize(exe)} bytes, 目录 {total // 1048576} MB)")
    return app_dir


def probe_runtime_dlls():
    """在开发态解释器里实际导入一遍启动所需模块，枚举已加载模块路径。"""
    proc = subprocess.run([sys.executable, "-c", PROBE_SRC],
                          cwd=ROOT, capture_output=True,
                          encoding="utf-8", errors="replace")
    if proc.stderr.strip():
        for line in proc.stderr.splitlines():
            print("  " + line)
    for line in proc.stdout.splitlines():
        if line.startswith("PROBE_JSON "):
            import json as _json
            return _json.loads(line[len("PROBE_JSON "):])
    raise SystemExit("依赖探针失败：未取到已加载模块列表")


def fix_missing_dlls(app_dir):
    """把探针检出、但打包含漏的 PATH 依赖 DLL 补进 _internal。

    规则：跳过已打包的、系统目录里的（交给 Windows 解析）、venv 内的
    （PyInstaller 本应收录）、codex 运行时缓存与 icu*.dll（有意不进包，
    运行时用系统 ICU）；其余（典型来自 anaconda Library\\bin）显式补入。
    """
    internal = os.path.join(app_dir, "_internal")
    system_dirs = tuple(os.path.normcase(d) for d in
                        (r"c:\windows\system32", r"c:\windows\syswow64", r"c:\windows"))
    venv_prefix = os.path.normcase(sys.prefix)
    existing = set()
    for dp, dn, fn in os.walk(app_dir):
        for f in fn:
            existing.add(f.lower())
    copied = []
    for path in probe_runtime_dlls():
        base = os.path.basename(path).lower()
        if not base.endswith(".dll") or base in existing:
            continue
        low_dir = os.path.normcase(os.path.dirname(path))
        if base.startswith("icu") or base.startswith("api-ms-"):
            continue
        if low_dir.startswith(venv_prefix):
            # conda 系 venv 的依赖目录 .venv\Library\bin 不在 PyInstaller
            # 的搜索路径里（普通 venv 无 conda-meta），里面的 DLL 会漏收；
            # 其它 venv 内目录（python311.dll / VCRUNTIME140 等）PyInstaller
            # 本身会收录，不用重复拷贝。
            libbin = os.path.join(venv_prefix, "library", "bin") + os.sep
            if not (low_dir + os.sep).startswith(libbin):
                continue
        if low_dir.startswith(system_dirs) or "codex-runtimes" in low_dir:
            continue
        shutil.copy2(path, os.path.join(internal, base))
        existing.add(base)
        copied.append((base, path))
    for base, src in copied:
        print("补依赖: %-24s <- %s" % (base, src))
    if copied:
        print(f"共补充 {len(copied)} 个 PATH 依赖 DLL")
    return len(copied)


# 纯 QWidgets 应用用不上的大块组件（按 _internal 相对路径）
PRUNE_FILES = [
    "PySide6/Qt6Quick.dll", "PySide6/Qt6Qml.dll", "PySide6/Qt6QmlModels.dll",
    "PySide6/Qt6QmlMeta.dll", "PySide6/Qt6QmlWorkerScript.dll",
    "PySide6/Qt6Pdf.dll", "PySide6/Qt6VirtualKeyboard.dll",
    "PySide6/Qt6Network.dll", "PySide6/QtNetwork.pyd",
    "PySide6/opengl32sw.dll",  # 纯 QWidgets 走 GDI/Direct2D，无需软件 GL
]
PRUNE_DIRS = [
    "Pythonwin",  # pywin32 自带的 IDE，运行时不需要
]
# (glob, 保留集合)；运行时图标全部 QPainter 现绘，只留 ico/svg 插件
PRUNE_GLOBS = [
    ("PySide6/translations/*.qm", ("qtbase_zh_CN.qm",)),
    ("PySide6/plugins/imageformats/q*.dll", ("qico.dll", "qsvg.dll")),
    ("PySide6/plugins/tls/q*.dll", ()),
    ("PySide6/plugins/networkinformation/q*.dll", ()),
    ("PySide6/plugins/platforminputcontexts/q*.dll", ()),
    # 外来精简版 ICU 缺 ucnv_* 导出，进包后会遮蔽系统 icuuc.dll
    ("icu*.dll", ()),
]


def prune(app_dir):
    """删除打包产物中纯 QWidgets 应用用不上的组件，给安装包瘦身。"""
    import glob as _glob
    saved = [0]

    def rm(path):
        if os.path.exists(path):
            saved[0] += os.path.getsize(path)
            if os.path.isdir(path):
                shutil.rmtree(path, ignore_errors=True)
            else:
                os.remove(path)

    for rel in PRUNE_FILES:
        rm(os.path.join(app_dir, "_internal", *rel.split("/")))
    for rel in PRUNE_DIRS:
        rm(os.path.join(app_dir, "_internal", *rel.split("/")))
    for pattern, keep in PRUNE_GLOBS:
        for path in _glob.glob(os.path.join(app_dir, "_internal", *pattern.split("/"))):
            if os.path.basename(path) not in keep:
                rm(path)
    if saved[0]:
        print(f"prune: 删除 {saved[0] / 1048576:.1f} MB无用组件")
    return saved[0]


def render_iss():
    """按本机 Inno的语言文件 availability 生成 iss（缺中文 isl 时降级英文）。
    版本号从 app.__version__ 注入到 {VERSION} 占位符，避免与代码版本漂移。"""
    tpl = io.open(os.path.join("build", "ai4word.iss"), encoding="utf-8").read()
    iscc = find_iscc()
    zh = os.path.join(os.path.dirname(iscc), "Languages", "ChineseSimplified.isl") if iscc else ""
    if iscc and os.path.exists(zh):
        langs = ('Name: "chinesesimp"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"\n'
                 'Name: "english"; MessagesFile: "compiler:Default.isl"')
    else:
        langs = 'Name: "english"; MessagesFile: "compiler:Default.isl"'
    out = os.path.abspath(os.path.join("build", "_ai4word_gen.iss"))
    rendered = tpl.replace("{LANGUAGES}", langs).replace('"{VERSION}"', f'"{__version__}"')
    io.open(out, "w", encoding="utf-8").write(rendered)
    return out


def make_installer(app_dir):
    iscc = find_iscc()
    iss = render_iss()
    if not iscc:
        print("\n[跳过] 未检测到 Inno Setup 6，已生成 .iss，可自行编译：")
        print(f"  {iss}")
        print(f"  ISCC {iss}")
        return None
    out = subprocess.check_output([iscc, "/Qp", iss], cwd=ROOT)
    setup = os.path.join(DIST, f"{NAME}-Setup-{__version__}.exe")
    if os.path.exists(setup):
        print(f"安装包: {setup} ({os.path.getsize(setup) // 1024 // 1024} MB)")
        return setup
    print(out.decode("utf-8", "replace"))
    raise SystemExit("ISCC 编译失败")


def main():
    ensure_icon()
    app_dir = run_pyinstaller()
    fix_missing_dlls(app_dir)
    prune(app_dir)
    make_installer(app_dir)
    print("\n打包完成。开发态运行请用: .\\.venv\\Scripts\\python.exe ai4word.pyw")


if __name__ == "__main__":
    main()
