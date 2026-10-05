# -*- coding: utf-8 -*-
"""Streaming-vs-block text invariant.

feed()+flush() and write_block() must produce a byte-identical document:
restore_snapshot() (rollback) replays blocks through write_block, so any
divergence means a rollback does not restore the pre-write text.

The closing-emphasis case used to fail: the closing ``**``/``*`` followed by
whitespace was typed literally during streaming, while the parsed block path
applied formatting and dropped the markers.
"""
import os
import sys

_TESTS = os.path.dirname(os.path.abspath(__file__))
_REPO = os.path.dirname(_TESTS)
sys.path.insert(0, _TESTS)
sys.path.insert(0, _REPO)

import test_fake_word as fw  # noqa: E402
from streaming_writer import StreamingWriter  # noqa: E402

CASES = [
    "plain paragraph with no markers at all.",
    "**bold** word and *italic* word and ***both*** word.",
    "nested **bold and *italic* inside** word.",
    "inline `code span` mid sentence.",
    "literal star math: 3 * 4 = 12 done.",
    "unordered list:\n\n- alpha\n- beta\n- gamma",
    "> quoted line one\n> quoted line two",
    "trailing emphasis at end of line **bold**",
    "multiple **a** and *b* and **c** plus *d*.",
    "# heading one\n\n"
    "a fairly long body paragraph with **bold** and *italic* "
    "markers, deliberately over one hundred characters so the "
    "stream progress threshold is crossed more than once.\n\n"
    "- item alpha\n- item beta\n- item gamma\n\n"
    "tail paragraph to close the first document.",
]


class _Blk(object):
    def __init__(self, kind):
        self.kind = kind


class _Model(object):
    def register(self, kind, *a, **k):
        return _Blk(kind)


def _fresh():
    app = fw.FakeApp()
    doc = app.doc
    writer = StreamingWriter(app, doc, app.sel, _Model())
    return doc, writer


def test_streamed_equals_block():
    for md in CASES:
        doc_a, w_a = _fresh()
        for i in range(0, len(md), 5):   # small chunks like the engine feeds
            w_a.feed(md[i:i + 5])
        w_a.flush()
        streamed = doc_a.Content.Text

        doc_b, w_b = _fresh()
        w_b.write_block(md, animate=False)
        blocked = doc_b.Content.Text

        assert streamed == blocked, (
            "streaming and write_block diverge for %r\nstreamed=%r\nblocked=%r"
            % (md[:60], streamed, blocked))
