# -*- coding: utf-8 -*-
"""Block-model document ordering invariant.

select_block() moves the Word selection to the clicked block's end so the
user (and the next AI write) can continue writing from there. register()
used to append such blocks to the tail of model.blocks even when they were
written mid-document, so rebuild_ranges() and restore_snapshot() replayed
the paragraphs in write order rather than document order; a rollback then
produced a document different from the pre-write one.

Invariants covered here:
  * model.blocks stays sorted by document position after mid-document writes
  * snapshot() records document positions and restore_snapshot() replays in
    document order, so a rollback restores the exact pre-write text
"""
import importlib.util
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

_spec = importlib.util.spec_from_file_location(
    "fake_word_mod", os.path.join(os.path.dirname(__file__), "test_fake_word.py"))
_fake = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_fake)

from streaming_writer import StreamingWriter  # noqa: E402


def _starts(model):
    out = []
    for b in model.blocks:
        try:
            out.append(int(b.range.Start))
        except Exception:
            out.append(None)
    return out


def _write(writer, md):
    writer.reset_anchor()
    _fake.feed_in_pieces(writer, md)
    writer.flush()


def test_mid_document_write_keeps_block_order():
    app, writer, model = _fake.make()
    writer.set_speed("fast")
    _write(writer, "# heading one\n\nplain body paragraph one.")
    before = app.doc.content
    snap = model.snapshot()
    assert snap and all("pos" in item for item in snap)

    # simulate a block-map click: cursor moves to block 0's end
    model.select_block(0)
    _write(writer, "mid document insert")

    starts = _starts(model)
    assert None not in starts, starts
    assert starts == sorted(starts), f"blocks out of document order: {starts}"
    kinds = [b.kind for b in model.blocks]
    assert kinds[0] == "heading1", kinds
    _fake.assert_aligned(model, app)

    # rollback must restore the exact pre-write document
    model.restore_snapshot(snap)
    assert app.doc.content == before, (repr(app.doc.content), repr(before))
    _fake.assert_aligned(model, app)


def test_two_mid_document_writes_stay_ordered():
    app, writer, model = _fake.make()
    writer.set_speed("fast")
    _write(writer, "# alpha\n\npara one.")
    _write(writer, "# beta\n\npara two.")
    before = app.doc.content
    snap = model.snapshot()

    # click block 1 (top), write mid-document, then click block 0, write again
    model.select_block(1)
    _write(writer, "first insert")
    model.select_block(0)
    _write(writer, "second insert")

    starts = _starts(model)
    assert None not in starts, starts
    assert starts == sorted(starts), starts
    _fake.assert_aligned(model, app)

    model.restore_snapshot(snap)
    assert app.doc.content == before, repr(app.doc.content)
    _fake.assert_aligned(model, app)


def test_snapshot_survives_range_loss():
    """Blocks whose range is gone must not break the sort (None keeps order)."""
    app, writer, model = _fake.make()
    writer.set_speed("fast")
    _write(writer, "# heading\n\nbody paragraph.")
    snap = model.snapshot()
    snap[0]["pos"] = None  # range lost, e.g. older snapshot without pos
    snap[1]["pos"] = None
    before = app.doc.content
    model.restore_snapshot(snap)
    assert app.doc.content == before, repr(app.doc.content)
