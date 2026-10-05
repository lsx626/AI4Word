"""Offline tests for ai_stream connection-level retry (no real network)."""
import json

import pytest
import requests

import ai_client


class FakeResp:
    def __init__(self, lines):
        self.lines = list(lines)

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False

    def raise_for_status(self):
        pass

    def iter_lines(self, decode_unicode=True):
        for line in self.lines:
            yield line


def sse(piece):
    return "data: " + json.dumps({"choices": [{"delta": {"content": piece}}]})


DONE = "data: [DONE]"


def patch_post(sequence):
    calls = {"n": 0}

    def fake_post(url, headers=None, json=None, stream=True, timeout=None):
        calls["n"] += 1
        item = sequence[min(calls["n"] - 1, len(sequence) - 1)]
        if isinstance(item, Exception):
            raise item
        if isinstance(item, FakeResp):
            return item
        return FakeResp(item)

    orig = ai_client.requests.post
    ai_client.requests.post = fake_post
    return orig, calls


def test_retry_then_success():
    seq = [
        requests.exceptions.ConnectionError("handshake aborted"),
        [sse("hello"), DONE],
    ]
    orig, calls = patch_post(seq)
    try:
        pieces = list(ai_client.ai_stream("p", "k", "s", attempts=3))
        assert pieces == ["hello"]
        assert calls["n"] == 2
    finally:
        ai_client.requests.post = orig


def test_no_retry_after_content_yielded():
    class LinesResp(FakeResp):
        def iter_lines(self, decode_unicode=True):
            yield sse("partial")
            raise requests.exceptions.ChunkedEncodingError("mid-stream reset")

    resp = LinesResp([sse("partial"), DONE])
    orig, calls = patch_post([resp])
    try:
        with pytest.raises(requests.exceptions.ChunkedEncodingError):
            list(ai_client.ai_stream("p", "k", "s", attempts=3))
        assert calls["n"] == 1
    finally:
        ai_client.requests.post = orig


def test_non_retryable_error_not_retried():
    seq = [ValueError("http 400 style logic error")]
    orig, calls = patch_post(seq)
    try:
        with pytest.raises(ValueError):
            list(ai_client.ai_stream("p", "k", "s", attempts=3))
        assert calls["n"] == 1
    finally:
        ai_client.requests.post = orig


def test_attempts_exhausted_raises():
    seq = [requests.exceptions.ConnectionError("down")]
    orig, calls = patch_post(seq)
    try:
        with pytest.raises(requests.exceptions.ConnectionError):
            list(ai_client.ai_stream("p", "k", "s", attempts=2))
        assert calls["n"] == 2
    finally:
        ai_client.requests.post = orig
