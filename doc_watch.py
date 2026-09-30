"""文档看门狗：主线程轮询文档状态，检测用户手动编辑与选区变化。

为什么不用后台线程？Word COM 对象绑定在创建它的线程（主线程 STA）上，
跨线程调用需要 marshaling，既慢又容易把 Word 搞坏；订阅 COM 事件则依赖
单线程消息泵，在 pywin32 脚本里同样不可靠。改为在自然的空闲点（每个流式
片段到达前、排版指令处理前）由主线程轮询：

- 写操作期间（feed/flush）由 StreamingWriter 置 busy 并在写完后重置基线，
  AI 自己的写入不会误报成"用户改动"；
- 检测到段落数变化 → 用户编辑（块索引可能已偏移）；仅选区变化 → 用户挪了
  光标。事件入队，由 drain_events() 在主循环消费并提示。
"""
import time

POLL_INTERVAL = 1.0  # 轮询间隔（秒）


class DocWatch:
    def __init__(self, app, doc):
        self._app = app
        self._doc = doc
        self.busy = False
        self._last_poll = 0.0
        self._baseline = None  # (段落数, 选区起点, 选区终点)
        self.events = []

    def _snapshot(self):
        try:
            return (self._doc.Paragraphs.Count,
                    self._app.Selection.Start,
                    self._app.Selection.End)
        except Exception:
            return None

    def begin_write(self):
        """写操作开始：屏蔽轮询。"""
        self.busy = True

    def end_write(self):
        """写操作结束：重置基线，吞掉自身变更。"""
        self.busy = False
        self._baseline = self._snapshot()

    def poll(self):
        """在空闲点调用：与基线比较，有变化则入队。"""
        if self.busy:
            return
        now = time.time()
        if now - self._last_poll < POLL_INTERVAL:
            return
        self._last_poll = now
        snap = self._snapshot()
        if snap is None:
            return
        old = self._baseline
        self._baseline = snap
        if old is None or old == snap:
            return
        if old[0] != snap[0]:
            self.events.append(
                f"注意到文档段落数 {old[0]} → {snap[0]}："
                f"你在我生成期间手动编辑了文档，块索引可能已经偏移。")
        else:
            self.events.append(
                f"注意到选区移动到 {snap[1]}：你在我生成期间调整了光标。")

    def drain_events(self):
        """取出并清空事件队列。"""
        out = self.events
        self.events = []
        return out