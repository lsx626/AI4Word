"""会话记忆：跨轮次记住用户的排版偏好与最近的交互结果。

gen_code 每次请求都是无状态的，"标题用黑体"这类偏好用户每轮都要重复说。
Session 收集两类记忆：

- 最近 N 轮交互（指令 + 成败），让模型知道哪些表达已经失败过、哪些路子
  用户爱用，避免重复踩坑；
- 用户已确认的偏好：命中的成功指令里含排版关键词的，原样记住；另外 AI
  可以通过 remember(note) 主动记录用户口头表达的长期习惯。

memory_prompt() 把记忆拼进 gen_code 的提示词；不含任何记忆时返回空串。
"""
import re

_PREFERENCE_RE = re.compile(
    r"字体|字号|颜色|对齐|缩进|行距|间距|样式|排版|标题|段落|加粗|斜体|页边距|页眉|页脚|预设")

class Session:
    def __init__(self, history_limit=8, pref_limit=10):
        self.history_limit = history_limit
        self.pref_limit = pref_limit
        self.history = []      # [{"prompt": str, "ok": bool}]
        self.preferences = []  # 用户已确认的排版偏好
        self.notes = []        # AI 主动记录的长期备注

    def record_turn(self, prompt, ok):
        """记录一轮排版指令及其成败；成功且含排版关键词的升格为偏好。"""
        prompt = (prompt or "").strip()
        if not prompt:
            return
        self.history.append({"prompt": prompt, "ok": ok})
        if len(self.history) > self.history_limit:
            del self.history[0]
        if ok and _PREFERENCE_RE.search(prompt) and prompt not in self.preferences:
            self.preferences.append(prompt)
            if len(self.preferences) > self.pref_limit:
                del self.preferences[0]

    def remember(self, note):
        """AI 主动记录的长期备注（用户口头表达的偏好）。"""
        note = (note or "").strip()
        if note and note not in self.notes:
            self.notes.append(note)
        return note

    def memory_prompt(self):
        """拼成可注入 gen_code 的记忆文本；无记忆返回空串。"""
        parts = []
        if self.history:
            lines = []
            for h in self.history[-self.history_limit:]:
                mark = "成功" if h["ok"] else "失败"
                lines.append(f"  [{mark}] {h['prompt']}")
            parts.append("最近的排版指令（最新在下，注意不要重复失败过的做法）：\n" + "\n".join(lines))
        if self.preferences:
            items = "\n".join(f"  - {p}" for p in self.preferences)
            parts.append("用户已确认的排版偏好（除非用户明确改变，请直接沿用）：\n" + items)
        if self.notes:
            items = "\n".join(f"  - {n}" for n in self.notes)
            parts.append("长期备注：\n" + items)
        return "\n\n".join(parts)