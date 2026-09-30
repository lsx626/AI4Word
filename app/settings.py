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
    """设置目录：%APPDATA%（无则用户主目录）。"""
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    d = os.path.join(base, APP_NAME)
    try:
        os.makedirs(d, exist_ok=True)
    except OSError:
        d = os.path.abspath(".")
    return d


class Settings:
    """一个带默认值合并的 JSON 设置对象。"""

    def __init__(self, path=None):
        self.path = path or os.path.join(user_dir(), "settings.json")
        self._data = dict(DEFAULTS)

    def load(self):
        try:
            with open(self.path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                for k, v in data.items():
                    if k in DEFAULTS or k.startswith("_free_"):
                        self._data[k] = v
        except (OSError, ValueError):
            pass  # 首次运行 / 损坏：沿用默认值
        return self

    def save(self):
        tmp = self.path + ".tmp"
        try:
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(self._data, f, ensure_ascii=False, indent=2)
            os.replace(tmp, self.path)
        except OSError:
            pass

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
