# -*- coding: utf-8 -*-
"""离线测试：Settings 持久化层。

覆盖：默认值合并、损坏文件降级、往返保存、原子替换（不留 .tmp 残留）、
apply_env 对 os.environ 与 ai_client 全局 base_url 的写入及空值不覆盖。
"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app.settings import Settings, DEFAULTS, DEFAULT_BASE_URL


def _tmp_settings(tmpdir, name="settings.json"):
    return Settings(os.path.join(tmpdir, name))


def test_defaults_when_missing():
    s = _tmp_settings(os.environ["TEMP"], "ai4w_missing.json")
    s.load()
    assert s.get("api_key") == ""
    assert s.get("base_url") == DEFAULT_BASE_URL
    assert s.get("stay_on_top") is True
    assert s.get("speed") == "auto"
    assert s.get("geometry") is None


def test_corrupt_file_falls_back():
    path = os.path.join(os.environ["TEMP"], "ai4w_corrupt.json")
    with open(path, "w", encoding="utf-8") as f:
        f.write("{not json at all")
    s = Settings(path)
    s.load()
    assert s.get("model") == DEFAULTS["model"]


def test_roundtrip():
    s = _tmp_settings(os.environ["TEMP"], "ai4w_rt.json")
    s.set("api_key", "sk-test-123")
    s.set("model", "Atria-Dawn-Preview")
    s.set("geometry", [120, 340])
    s.set("expanded", True)
    s.save()
    s2 = Settings(s.path)
    s2.load()
    assert s2.get("api_key") == "sk-test-123"
    assert s2.get("geometry") == [120, 340]
    assert s2.get("expanded") is True


def test_unknown_keys_ignored():
    path = os.path.join(os.environ["TEMP"], "ai4w_unknown.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump({"api_key": "k", "evil_key": "x"}, f)
    s = Settings(path)
    s.load()
    assert s.get("api_key") == "k"
    assert s.get("evil_key") is None


def test_save_is_atomic():
    s = _tmp_settings(os.environ["TEMP"], "ai4w_atomic.json")
    s.set("api_key", "atomic")
    s.save()
    assert not os.path.exists(s.path + ".tmp")
    with open(s.path, encoding="utf-8") as f:
        data = json.load(f)
    assert data["api_key"] == "atomic"


def test_apply_env_sets_and_skips_empty():
    s = _tmp_settings(os.environ["TEMP"], "ai4w_env.json")
    s.set("api_key", "sk-env-1")
    s.set("model", "Atria-Dawn-Preview")
    s.set("base_url", "https://example.test/v1")
    info = s.apply_env()
    assert os.environ.get("ATRIA_API_KEY") == "sk-env-1"
    assert os.environ.get("ATRIA_MODEL") == "Atria-Dawn-Preview"
    assert info["base_url"] == "https://example.test/v1"

    # 空值不覆盖已有环境（如 .env 已配置的情况）
    s2 = _tmp_settings(os.environ["TEMP"], "ai4w_env2.json")
    s2.apply_env()
    assert os.environ.get("ATRIA_API_KEY") == "sk-env-1"

    from ai_client import get_base_url
    assert get_base_url() == DEFAULT_BASE_URL  # 空设置回落默认


def test_slot_settings_redact_api_key():
    """AI4WORD_TEST_SETTINGS（deep_test 并行槽）下 save() 不把 api_key
    写入临时文件；内存中的值不受影响，API 调用走 os.environ
    （apply_env / main.py / engine.py 均不读盘上的 api_key）。"""
    slot_path = os.path.join(os.environ["TEMP"], "ai4w_slot_settings.json")
    os.environ["AI4WORD_TEST_SETTINGS"] = slot_path
    try:
        s = Settings(slot_path)
        s.set("api_key", "sk-secret-dont-persist")
        s.set("model", "Atria-Dawn-Preview")
        assert s.save() is True
        # 内存里仍是完整密钥
        assert s.get("api_key") == "sk-secret-dont-persist"
        # 落盘的必是空串
        with open(slot_path, encoding="utf-8") as f:
            on_disk = json.load(f)
        assert on_disk["api_key"] == ""
        assert on_disk["model"] == "Atria-Dawn-Preview"
    finally:
        del os.environ["AI4WORD_TEST_SETTINGS"]
        for p in (slot_path,):
            try:
                os.remove(p)
            except OSError:
                pass
    # 非测试模式（env 已清）：密钥照常持久化
    s2 = Settings(os.path.join(os.environ["TEMP"], "ai4w_normal.json"))
    s2.set("api_key", "sk-normal-persist")
    s2.save()
    with open(s2.path, encoding="utf-8") as f:
        assert json.load(f)["api_key"] == "sk-normal-persist"
