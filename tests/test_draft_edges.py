# -*- coding: utf-8 -*-
"""Draft-edge whitespace invariant: committed doc must equal replayed md.

A stream cut off right after "plain " leaves the typed trailing space in the
document, while the committed block md is stripped.  restore_snapshot()
(rollback) replays blocks through write_block, which also strips, so the
post-rollback document used to lose that space.  _commit_draft now trims the
typed whitespace edges at commit time, keeping document, plain text and md
byte-identical -- so feed()+flush() matches write_block() exactly and a
rollback restores the pre-write text.
"""
import os
import sys

_TESTS = os.path.dirname(os.path.abspath(__file__))
_REPO = os.path.dirname(_TESTS)
sys.path.insert(0, _TESTS)
sys.path.insert(0, _REPO)

import test_fake_word as fw  # noqa: E402
from streaming_writer import StreamingWriter  # noqa: E402
from doc_model import DocModel  # noqa: E402

EDGE_CASES = [
    "paragraph that ends with a space ",
    "two spaces at the end  ",
    "paragraph with markers and trailing space **bold** ",
    "italic tail with space *italic* ",
    "plural markers **a** and *b* at the end ",
    "last word ",
    "multi paragraph doc\n\nsecond paragraph ",
    "multi paragraph doc \n\nsecond paragraph  \n\nthird ",
    "headline\n\nbody paragraph with tail \n\n- list item ",
]


def _fresh():
    app = fw.FakeApp()
    doc = app.doc
    sel = app.sel
    writer = StreamingWriter(app, doc, sel, model=None, char_delay=0.0)
    model = DocModel(app, doc, sel, writer)
    writer.model = model
    return doc, writer, model


def _stream(writer, md, chunk=5):
    for i in range(0, len(md), chunk):
        writer.feed(md[i:i + chunk])
    writer.flush()


def test_streamed_equals_block_for_edges():
    """feed()+flush() and write_block() must give a byte-identical doc."""
    for md in EDGE_CASES:
        doc_a, w_a, _ = _fresh()
        _stream(w_a, md)
        streamed = doc_a.Content.Text

        doc_b, w_b, _ = _fresh()
        w_b.write_block(md, animate=False)
        blocked = doc_b.Content.Text

        assert streamed == blocked, (
            "streaming and write_block diverge for %r\nstreamed=%r\nblocked=%r"
            % (md[:60], streamed, blocked))


def test_rollback_restores_pre_write_text():
    """Snapshot before a write with padded edges; rollback is byte-exact."""
    doc, writer, model = _fresh()
    writer.write_block("seed paragraph", animate=False)
    before = doc.Content.Text
    snap = model.snapshot()
    assert snap and len(snap) == 1

    _stream(writer, "first paragraph ends padded \n\nsecond padded  ")

    after = doc.Content.Text
    assert after != before, "streamed text never landed in the document"

    model.restore_snapshot(snap)
    restored = doc.Content.Text
    assert restored == before, (
        "rollback diverged from pre-write text\nbefore=%r\nrestored=%r"
        % (before, restored))


def test_trailing_space_not_left_in_document():
    """A trailing space typed by the stream is removed at commit time."""
    doc, writer, model = _fresh()
    _stream(writer, "plain ")
    # The committed paragraph must not keep the typed trailing space: the
    # registered md has none, and rollback replay drops it either way.
    body = doc.Content.Text.replace("\r", "")
    assert body == "plain"
    blocks = model.blocks
    assert blocks and blocks[-1].md == "plain"
    assert blocks[-1].text == "plain"
