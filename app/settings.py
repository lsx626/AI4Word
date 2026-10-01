# -*- coding: utf-8 -*-
"""设置持久化：%APPDATA%\AI4Word\settings.json，GUI 设置面板的值优先于 .env。"""
import json
import os

from ai_client import DEFAULT_MODEL, get_base_url, set_base_url

APP_NAME = "AI4Word"
DEFAULT_BASE_URL = "https://discovery-api.intern-ai.org.cn/v1"

DEFAULTS = {
    "api_key": "",
    "base_url": DEFAULT_BASE_URL,
    "model": DEFAULT_MODEL,
    "auto_start": False,
    "stay_on_top": True,
    "speed": "auto",        # auto / slow / fast
    "geometry": None,       # [x, y] 悬浮窗位置
    "expanded": False,      # 上次是否展开
}


def user_dir():
    """设置目录：%APPDATA%（无则用户主目录；再失败回退 %TEMP%）。

    回退不用 CWD：开机自启（Run 键）启动时 CWD 常常是 System32 之类
    的不可写目录，后续 save() 会静默失败。
    """
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    d = os.path.join(base, APP_NAME)
    try:
        os.makedirs(d, exist_ok=True)
    except OSError:
        d = os.path.join(os.environ.get("TEMP") or os.path.expanduser("~") or ".",
                          APP_NAME)
        try:
            os.makedirs(d, exist_ok=True)
        except OSError:
            d = os.path.abspath(".")
    return d


# 各键的合法类型校验：JSON 语法合法但值类型错误（手改 / 同步软件写错）
# 会让 QRect(str, str) / QLineEdit(int) 之类在启动时直接崩溃，
# 这里在合并时丢弃非法值，回退默认。
_VALIDATORS = {
    "api_key": lambda v: isinstance(v, str),
    "base_url": lambda v: isinstance(v, str),
    "model": lambda v: isinstance(v, str),
    "auto_start": lambda v: isinstance(v, bool),
    "stay_on_top": lambda v: isinstance(v, bool),
    "speed": lambda v: isinstance(v, str) and v in ("auto", "slow", "fast"),
    "expanded": lambda v: isinstance(v, bool),
    "geometry": lambda v: (isinstance(v, (list, tuple)) and len(v) == 2
                            and all(isinstance(x, (int, float)) and not isinstance(x, bool)
                                    for x in v)),
}


class Settings:
    """一个带默认值合并的 JSON 设置对象。"""

    def __init__(self, path=None):
        self.path = path or os.path.join(user_dir(), "settings.json")
        self._data = dict(DEFAULTS)

    def load(self):
        try:
            with open(self.path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except (OSError, ValueError):
            return self  # 首次运行 / 语法损坏：沿用默认值
        if isinstance(data, dict):
            for k, v in data.items():
                if k in _VALIDATORS:
                    if _VALIDATORS[k](v):
                        self._data[k] = v
                    # 非法类型：丢弃，用默认值（防止启动崩溃）
                elif k.startswith("_free_"):
                    self._data[k] = v
        return self

    def save(self):
        """原子保存；返回是否成功（失败时让调用方提示用户，而不是静默吞掉）。"""
        tmp = self.path + ".tmp"
        try:
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(self._data, f, ensure_ascii=False, indent=2)
            os.replace(tmp, self.path)
        except OSError:
            return False
        return True

    def get(self, key, default=None):
        return self._data.get(key, default)

    def set(self, key, value):
        self._data[key] = value

    def __getitem__(self, key):
        return self._data.get(key)

    def __setitem__(self, key, value):
        self._data[key] = value

    def apply_env(self):
        """把设置写入 os.environ 与 ai_client 的全局 base_url。"""
        key = (self.get("api_key") or "").strip()
        if key:
            os.environ["ATRIA_API_KEY"] = key
        model = (self.get("model") or "").strip()
        if model:
            os.environ["ATRIA_MODEL"] = model
        base = (self.get("base_url") or "").strip() or DEFAULT_BASE_URL
        set_base_url(base)
        return {"base_url": get_base_url(), "model": model or DEFAULT_MODEL}
