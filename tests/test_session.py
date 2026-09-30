# -*- coding: utf-8 -*-
"""离线测试：会话记忆（Session）的记录、偏好升格与提示词拼装。"""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from session import Session


def test_record_turn_and_preference():
    s = Session()
    s.record_turn("把标题字体改成黑体", True)
    assert len(s.history) == 1
    # 含排版关键词且成功 → 升格为偏好
    assert s.preferences == ["把标题字体改成黑体"]
    # 失败的指令进历史但不进偏好
    s.record_turn("普通指令", False)
    assert s.history[-1]["ok"] is False
    assert "普通指令" not in s.preferences
    # 空串忽略
    s.record_turn("   ", True)
    assert len(s.history) == 2
    print("ok: 记录轮次（成败标记、关键词偏好升格、空串忽略）")


def test_remember_notes():
    s = Session()
    assert s.remember("用户喜欢宋体正文") == "用户喜欢宋体正文"
    assert s.remember("用户喜欢宋体正文") == "用户喜欢宋体正文"  # 去重
    assert len(s.notes) == 1
    s.remember("")  # 空备注忽略
    assert len(s.notes) == 1
    print("ok: 长期备注（去重、空串忽略）")


def test_memory_prompt():
    s = Session()
    assert s.memory_prompt() == ""  # 无记忆返回空串
    s.record_turn("标题加粗", True)
    s.remember("正文用宋体小四")
    p = s.memory_prompt()
    assert "标题加粗" in p and "成功" in p
    assert "正文用宋体小四" in p
    assert "排版偏好" in p
    # 失败指令进提示词但不进偏好
    s2 = Session()
    s2.record_turn("失败的字号调整", False)
    p2 = s2.memory_prompt()
    assert "失败" in p2 and "字号" in p2
    assert len(s2.preferences) == 0
    print("ok: memory_prompt 拼装（历史/偏好/备注三段）")


def test_limits():
    s = Session(history_limit=8, pref_limit=10)
    for i in range(12):
        s.record_turn(f"调整字号 {i}", True)  # 每条都含"字号"关键词
    assert len(s.history) == 8, len(s.history)
    assert s.history[-1]["prompt"] == "调整字号 11"  # 保留最新
    assert len(s.preferences) == 10, len(s.preferences)
    print("ok: 记忆上限（历史 8 条、偏好 10 条，保留最新）")


if __name__ == "__main__":
    test_record_turn_and_preference()
    test_remember_notes()
    test_memory_prompt()
    test_limits()
    print("\n全部测试通过。")
