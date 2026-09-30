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
    subprocess.check_call(cmd)
    app_dir = os.path.join(DIST, NAME)
    if not os.path.isdir(app_dir):
        raise SystemExit("打包失败：未找到 " + app_dir)
    exe = os.path.join(app_dir, NAME + ".exe")
    print("程序目录:", app_dir, f"({os.path.getsize(exe)} bytes exe)")
    return app_dir


def render_iss():
    """按本机 Inno的语言文件 availability 生成 iss（缺中文 isl 时降级英文）。"""
    tpl = io.open(os.path.join("build", "ai4word.iss"), encoding="utf-8").read()
    iscc = find_iscc()
    zh = os.path.join(os.path.dirname(iscc), "Languages", "ChineseSimplified.isl") if iscc else ""
    if iscc and os.path.exists(zh):
        langs = ('Name: "chinesesimp"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"\n'
                 'Name: "english"; MessagesFile: "compiler:Default.isl"')
    else:
        langs = 'Name: "english"; MessagesFile: "compiler:Default.isl"'
    out = os.path.abspath(os.path.join("build", "_ai4word_gen.iss"))
    io.open(out, "w", encoding="utf-8").write(tpl.replace("{LANGUAGES}", langs))
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
    make_installer(app_dir)
    print("\n打包完成。开发态运行请用: .\\.venv\\Scripts\\python.exe ai4word.pyw")


if __name__ == "__main__":
    main()
