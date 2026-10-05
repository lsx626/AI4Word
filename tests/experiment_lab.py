# -*- coding: utf-8 -*-
"""Experiment lab -- fast iterative deep experiments for AI4Word.

Fake Word + scripted AI + offscreen Qt. One process == one iteration of the
deep-experiment workflow: run a scenario -> triage debug.log -> report.
Real Word / real API are NOT needed here (they are covered separately by
tests/debug_walkthrough.py); the lab trades realism for speed so that dozens
of iterations are practical.

Usage:
    python tests/experiment_lab.py --iter N [--scenario NAME] [--seed S]

Exit codes: 0 = clean, 1 = findings, 2 = harness error.
Archives the debug log of each iteration to tests/lab_runs/.
"""
import argparse
import io
import json
import os
import random
import re
import shutil
import sys
import time
import traceback

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

_TESTS = os.path.dirname(os.path.abspath(__file__))
_REPO = os.path.dirname(_TESTS)
_RUNS = os.path.join(_TESTS, "lab_runs")
os.makedirs(_RUNS, exist_ok=True)

sys.path.insert(0, _TESTS)

# test_fake_word neutralizes time.sleep at import time; keep the real one.
_real_sleep = time.sleep
import test_fake_word as fw
time.sleep = _real_sleep

sys.path.insert(0, _REPO)

# Reuse proven helpers from the real-stack walkthrough harness.
import debug_walkthrough as W

import ai_client
import app.engine as engine_mod
from app import __version__, debug
from app.engine import AgentWorker
from app.main_window import MainWindow
from app.settings import Settings
from app.tray import Tray
from PySide6.QtWidgets import QApplication

# ---------------------------------------------------------------- fake Word

class _LabDocuments:
    """AgentWorker contract: app.Documents with Count / Add()."""
    def __init__(self, fake):
        self._fake = fake
    @property
    def Count(self):
        return 1
    def Add(self):
        return self._fake.doc


class LabApp:
    """FakeApp wrapped in the surface AgentWorker._ensure_ready uses."""
    def __init__(self):
        self._fake = fw.FakeApp()
        self.Documents = _LabDocuments(self._fake)
    @property
    def ActiveDocument(self):
        return self._fake.doc
    @property
    def Selection(self):
        return self._fake.sel
    @property
    def doc(self):
        return self._fake.doc
    @property
    def sel(self):
        return self._fake.sel


# ---------------------------------------------------------------- scripted AI

class Script:
    """Deterministic AI responses for one iteration."""
    def __init__(self):
        self.writes = [
            "# heading one\n\n"
            "a fairly long body paragraph with **bold** and *italic* "
            "markers, deliberately over one hundred characters so the "
            "stream progress threshold is crossed more than once.\n\n"
            "- item alpha\n- item beta\n- item gamma\n\n"
            "tail paragraph to close the first document.",
            "second document: a plain paragraph only, no structure at "
            "all, but still long enough to cross the progress threshold "
            "several times while streaming in.",
            "### h3 title\n\n"
            "quote section:\n\n> quoted line one\n> quoted line two\n\n"
            "final paragraph, also comfortably long.",
        ]
        self.chunk_delay = 0.012
        self.connect_error = False
        self.empty_stream = False
        self.stream_error_at = None      # chunk index where stream breaks
        self.arrange_code = "doc.Styles(-2).ParagraphFormat.Alignment = 1"
        self.bad_code = "x = 1 / 0"
        self.fix_code = "doc.Styles(-2).ParagraphFormat.Alignment = 1"
        self.gen_fail_first = False
        self.gen_raises = False

    def chunks_for(self, call):
        if self.empty_stream:
            return []
        text = self.writes[call % len(self.writes)]
        if len(text) <= 6:
            return [text]
        step = 5
        return [text[i:i + step] for i in range(0, len(text), step)]


def patch_ai(sc):
    """Replace ai_client + engine gen/fix with the scripted versions."""
    state = {"stream": 0, "gen": 0, "fix": 0, "request": 0}

    def fake_stream(prompt, api_key, system_prompt, model=None,
                    connect=10, read=600):
        call = state["stream"]
        state["stream"] += 1
        if sc.connect_error:
            raise ConnectionError("lab: simulated API connect failure")
        chunks = sc.chunks_for(call)
        for i, piece in enumerate(chunks):
            if sc.stream_error_at is not None and i >= sc.stream_error_at:
                raise ConnectionError("lab: stream broke mid-flight")
            yield piece
            time.sleep(sc.chunk_delay)

    def fake_request(prompt, api_key, system_prompt, model=None, timeout=180):
        state["request"] += 1
        return sc.arrange_code

    def fake_gen(prompt, api_key, session=None, block_map_fn=None, sink=print,
                 model=None):
        state["gen"] += 1
        if sc.gen_raises:
            raise RuntimeError("lab: gen_code failure")
        if sc.gen_fail_first and state["gen"] == 1:
            return sc.bad_code
        return sc.arrange_code

    def fake_fix(prompt, failed_code, error_msg, api_key, session=None,
                 sink=print, block_map_fn=None, model=None):
        state["fix"] += 1
        return sc.fix_code

    ai_client.ai_stream = fake_stream
    ai_client.ai_request = fake_request
    engine_mod.gen_code = fake_gen
    engine_mod.fix_code = fake_fix
    return state


# ---------------------------------------------------------------- assemble

_WORD = None


def lab_word_factory():
    return _WORD


def assemble_lab():
    global _WORD
    _WORD = LabApp()
    tmp = os.environ.get("TEMP") or _REPO
    settings_path = os.path.join(tmp, "ai4word_lab_settings.json")
    try:
        os.remove(settings_path)
    except OSError:
        pass
    settings = Settings(settings_path).load()
    settings.set("api_key", "lab-fake-key")
    settings.set("speed", "fast")

    app = QApplication(["AI4Word"])
    app.setApplicationName("AI4Word")
    app.setApplicationVersion(__version__)
    app.setOrganizationName("AI4Word")
    app.setQuitOnLastWindowClosed(False)

    worker = AgentWorker(word_factory=lab_word_factory)
    worker.set_api_key("lab-fake-key")
    W.patch_handle(worker)
    worker.streamChunk.connect(W._on_stream_chunk)
    window = MainWindow(worker, settings)
    tray = Tray(window, settings)
    window.set_tray(tray)
    return worker, window, tray, settings


def lab_quit(worker, window):
    if window._awaiting_choice:
        worker.choose("keep")
        W.wait_for(lambda: not window._awaiting_choice, 30)
    worker.send("quit")
    W.wait_for(lambda: not worker.isRunning(), 60)
    W.pump(0.2)


# ---------------------------------------------------------------- triage

WARN_RE = re.compile(r"\b(WARN|ERROR)\s+(\w+)")
EVENT_RE = re.compile(r"\b(INFO|WARN|ERROR)\s+(\w+)")

# Events that are expected by design in lab contexts.
BENIGN_ALWAYS = {
    "settings_load_first_run",   # normal first run, no settings file yet
    "hotkey_register_failed",    # offscreen platform artifact (no real HWND)
}

BENIGN_BY_SCENARIO = {
    "arrange_repair": {"stmt_failed"},
    "arrange_gen_fail": {"gen_code_failed"},
    "stream_faults": {"stream_error"},
    "settings_dialog": {"settings_save_failed", "autostart_disable_failed",
                        "autostart_enable_failed"},
    "stress_mixed": {"stmt_failed", "stream_error", "gen_code_failed",
                     "settings_save_failed", "autostart_disable_failed"},
}


def triage(scenario_name, must_events):
    """Parse debug.log -> (findings, stats)."""
    path = debug.log_path()
    if not path or not os.path.exists(path):
        return ["triage: debug.log missing after run"], {"lines": 0}
    raw = io.open(path, encoding="utf-8", errors="replace").read()
    lines = raw.splitlines()
    events = set()
    for ln in lines:
        m = EVENT_RE.search(ln)
        if m:
            events.add(m.group(2))
    findings = []
    for ev in must_events:
        if ev not in events:
            findings.append("triage: missing expected event %s" % ev)
    benign = BENIGN_ALWAYS | BENIGN_BY_SCENARIO.get(scenario_name, set())
    for ln in lines:
        m = WARN_RE.search(ln)
        if not m:
            continue
        event = m.group(2)
        if event not in benign:
            findings.append("triage: unexpected %s %s :: %s"
                            % (m.group(1), event, ln[:170]))
    # exc() writes traceback text into a field; flag any uncaught-hook fires.
    if "excepthook" in raw:
        findings.append("triage: uncaught exception hit sys.excepthook")
    stats = {"lines": len(lines), "events": len(events),
             "warn_error": sum(1 for ln in lines if WARN_RE.search(ln))}
    return findings, stats
# ---------------------------------------------------------------- scenarios

def _drain(worker):
    """Barrier between scenario steps.

    The W.<step> helpers wait for exactly ONE command, so anything still
    in flight (e.g. the three speed clicks of speed_gears) would corrupt
    the next wait_handled target and trip the "worker idle" pre-checks.
    A harmless speed command doubles as the barrier marker.
    """
    worker.send("speed", "fast")
    W.wait_handled(worker, 15)
    W.pump(0.1)


def _idle(worker):
    W.wait_for(lambda: not worker._busy, 30)
    W.pump(0.1)


def write_basic(win, worker, tray, settings, sc, rng):
    W.expand_collapse_avatar(win)
    W.speed_gears(win, worker, settings)
    _drain(worker)
    W.compact_send(win, worker)
    _drain(worker)
    W.panel_send(win, worker)
    _drain(worker)
    W.block_map(win, worker)
    _drain(worker)
    W.input_rejections(win, worker)


def interrupts(win, worker, tray, settings, sc, rng):
    for choice in ("rollback", "keep", "extra"):
        sc.chunk_delay = 0.12   # slow enough that the interrupt always lands
        W.interrupt_cycle(win, worker, choice,
                          extra_text="lab extra paragraph appended after interrupt.")
        sc.chunk_delay = 0.012
        _idle(worker)
    win._on_speed(2)
    _drain(worker)


def arrange_ok(win, worker, tray, settings, sc, rng):
    sc.gen_fail_first = False
    sc.gen_raises = False
    W.compact_send(win, worker)
    _drain(worker)
    W.arrange_flow(win, worker,
                   "center all level-one headings and bold the first one")
    _drain(worker)


def arrange_repair(win, worker, tray, settings, sc, rng):
    sc.gen_fail_first = True        # first gen_code returns broken code
    sc.bad_code = "x = 1 / 0"
    sc.fix_code = "doc.Styles(-2).ParagraphFormat.Alignment = 1"
    W.compact_send(win, worker)
    _drain(worker)
    W.arrange_flow(win, worker, "first attempt is deliberately broken")
    _drain(worker)
    W.arrange_flow(win, worker, "second pass after the self-repair")
    _drain(worker)


def arrange_gen_fail(win, worker, tray, settings, sc, rng):
    sc.gen_raises = True
    W.compact_send(win, worker)
    _drain(worker)
    win._mode_group.button(1).click()
    win.input_p.setPlainText("arrange request that raises inside gen_code")
    win._on_send_panel()
    W.check(W.wait_handled(worker, 300), "gen_code raise handled, worker idle")
    W.check(not worker._busy, "worker idle after gen_code failure")
    win._mode_group.button(0).click()
    W.pump(0.2)
    sc.gen_raises = False
    _drain(worker)


def stream_faults(win, worker, tray, settings, sc, rng):
    win.set_expanded(False)
    W.pump(0.2)
    # 1) connection failure before any chunk
    sc.connect_error = True
    win.input_c.setText("lab probe: api connect failure")
    win._on_send_compact()
    W.check(W.wait_handled(worker, 180), "connect error returns worker to idle")
    W.check(not worker._busy, "worker idle after connect error")
    sc.connect_error = False
    # 2) stream breaks mid-flight
    sc.stream_error_at = 3
    win.input_c.setText("lab probe: stream breaks after a few chunks")
    win._on_send_compact()
    W.check(W.wait_handled(worker, 180), "mid-stream error returns to idle")
    W.check(not worker._busy, "worker idle after mid-stream error")
    sc.stream_error_at = None
    # 3) empty AI reply
    sc.empty_stream = True
    win.input_c.setText("lab probe: empty AI reply")
    win._on_send_compact()
    W.check(W.wait_handled(worker, 180), "empty stream returns worker to idle")
    W.check(not worker._busy, "worker idle after empty stream")
    sc.empty_stream = False
    _drain(worker)


def review_presets_save(win, worker, tray, settings, sc, rng):
    W.compact_send(win, worker)
    _drain(worker)
    W.review_toggle(win, worker)
    _drain(worker)
    W.presets(win, worker)
    _drain(worker)
    W.save_blocks(win, worker)
    _drain(worker)


def settings_dialog(win, worker, tray, settings, sc, rng):
    W.compact_send(win, worker)
    _drain(worker)
    W.settings_dialog(win, worker, settings)


def keys_tray(win, worker, tray, settings, sc, rng):
    W.compact_send(win, worker)
    _drain(worker)
    W.keys_inputs_and_close(win, worker, tray)
    _drain(worker)
    W.esc_hide_and_hint(win, tray)
    W.drag_snap_summon_hotkey(win)
    _drain(worker)


_PRESETS = 4


def doc_persistence(win, worker, tray, settings, sc, rng):
    """Save blocks into the document, retire the worker, attach a NEW
    worker to the same document, verify load_blocks restores the block
    model, then prove the new worker is usable for real writes."""
    W.compact_send(win, worker)
    _drain(worker)
    W.save_blocks(win, worker)
    _drain(worker)
    n_before = len(worker._model.blocks)
    W.check(n_before > 0, "blocks registered before reconnect")
    worker.send("quit")
    W.check(W.wait_for(lambda: not worker.isRunning(), 60),
            "first worker exited cleanly")
    W.pump(0.3)
    w2 = AgentWorker(word_factory=lab_word_factory)
    w2.set_api_key("lab-fake-key")
    W.patch_handle(w2)
    w2.streamChunk.connect(W._on_stream_chunk)
    w2.start()
    W.check(W.wait_for(lambda: w2.isRunning(), 30), "second worker started")
    w2.send("save")   # launch=True -> _assemble -> load_blocks restore
    W.check(W.wait_handled(w2, 120), "second worker attached to same doc")
    n_after = len(w2._model.blocks) if w2._model is not None else -1
    W.check(n_after == n_before,
            "load_blocks restored the block count (%d -> %d)"
            % (n_before, n_after))
    win.worker = w2   # route UI sends to the new worker
    win.set_expanded(False)
    W.pump(0.2)
    win.input_c.setText("lab: write through the reconnected worker")
    win._on_send_compact()
    W.check(W.wait_handled(w2, 300), "write served by the reconnected worker")
    W.check(len(W.doc_text(w2) or "") > 10, "text landed after reconnect")
    _drain(w2)
    return w2


def stress_mixed(win, worker, tray, settings, sc, rng):
    pool = ("compact", "panel", "interrupt_rollback", "interrupt_keep",
            "interrupt_extra", "speed", "mode", "preset", "save",
            "blockmap", "review", "esc", "expand", "empty_send",
            "bad_arrange")
    n = rng.randint(14, 20)
    for i in range(n):
        action = "compact" if i == 0 else rng.choice(pool)
        _idle(worker)
        if action == "compact":
            sc.empty_stream = False
            sc.connect_error = False
            win.set_expanded(False)
            W.pump(0.2)
            win.input_c.setText("stress %d: compact write" % i)
            win._on_send_compact()
            W.wait_handled(worker, 300)
        elif action == "panel":
            win.set_expanded(True)
            W.pump(0.2)
            win.input_p.setPlainText("stress %d: panel write" % i)
            win._on_send_panel()
            W.wait_handled(worker, 300)
        elif action.startswith("interrupt_"):
            choice = action[len("interrupt_"):]
            sc.chunk_delay = 0.12
            W.interrupt_cycle(win, worker, choice,
                              extra_text="stress extra text number %d." % i)
            sc.chunk_delay = 0.012
        elif action == "speed":
            win._on_speed(rng.randint(0, 2))
            W.wait_handled(worker, 60)
        elif action == "mode":
            win._mode_group.button(1).click()
            win.input_p.setPlainText("stress %d: arrange pass" % i)
            win._on_send_panel()
            W.wait_handled(worker, 300)
            win._mode_group.button(0).click()
            W.pump(0.2)
        elif action == "preset":
            win._on_preset(rng.randint(1, _PRESETS))
            W.wait_handled(worker, 60)
        elif action == "save":
            win._on_save()
            W.wait_handled(worker, 60)
        elif action == "blockmap":
            W.block_map(win, worker)
        elif action == "review":
            win.btn_review.setChecked(not win.btn_review.isChecked())
            W.wait_handled(worker, 60)
        elif action == "esc":
            W.esc_hide_and_hint(win, tray)
        elif action == "expand":
            W.expand_collapse_avatar(win)
        elif action == "empty_send":
            W.input_rejections(win, worker)
        elif action == "bad_arrange":
            sc.gen_raises = True
            win._mode_group.button(1).click()
            win.input_p.setPlainText("stress %d: arrange that raises" % i)
            win._on_send_panel()
            W.wait_handled(worker, 300)
            win._mode_group.button(0).click()
            W.pump(0.2)
            sc.gen_raises = False
    win.summon()
    win._on_speed(2)
    _drain(worker)


# ---------------------------------------------------------------- driver

SCENARIOS = {
    "write_basic": write_basic,
    "interrupts": interrupts,
    "arrange_ok": arrange_ok,
    "arrange_repair": arrange_repair,
    "arrange_gen_fail": arrange_gen_fail,
    "stream_faults": stream_faults,
    "review_presets_save": review_presets_save,
    "settings_dialog": settings_dialog,
    "keys_tray": keys_tray,
    "doc_persistence": doc_persistence,
    "stress_mixed": stress_mixed,
}

# Events that must appear in the debug log per scenario -- proves the
# instrumentation covers the exercised code paths. ai_client events are
# excluded on purpose: the lab scripts the AI layer itself, so those
# log lines never run here.
MUST_EVENTS = {
    "write_basic": {
        "worker_loop_start", "queue_put", "handle_start", "word_assembled",
        "write_start", "snapshot_taken", "stream_progress",
        "write_done", "block_register", "feed", "flush",
        "speed_set", "select_block", "speed_clicked", "toggle_expand",
        "send_compact", "send_panel", "write_sent", "blockmap_toggled",
        "block_clicked", "send_rejected_too_long",
    },
    "interrupts": {
        "write_start", "interrupt_requested", "write_done", "choice_put",
        "choice_made", "rollback_ok", "extra_start", "extra_done",
        "interrupted_ui", "enter_extra_mode", "extra_submitted", "choose",
    },
    "arrange_ok": {
        "arrange_start", "gen_code_result", "gen_code", "run_code_result",
        "arrange_applied", "run_code_start", "run_code_done", "stmt_ok",
        "arrange_sent", "mode_clicked",
    },
    "arrange_repair": {
        "arrange_start", "gen_code_result", "run_code_result",
        "fix_code_result", "fix_code", "run_code2_result", "arrange_applied",
        "stmt_failed", "run_code_done",
    },
    "arrange_gen_fail": {
        "arrange_start", "gen_code_failed", "arrange_failed",
    },
    "stream_faults": {
        "word_assembled", "write_start", "snapshot_taken", "stream_error",
        "write_done",
    },
    "review_presets_save": {
        "review_set", "preset_applied", "save_blocks", "review_toggled",
        "preset_selected", "save_clicked", "block_register",
    },
    "settings_dialog": {
        "settings_open", "settings_accept", "settings_saved",
        "settings_save_failed", "write_start",
    },
    "keys_tray": {
        "send_compact", "write_start", "write_done", "hide_via_esc",
        "close_to_tray", "tray_activated", "snap_to_edge", "summon",
    },
    "doc_persistence": {
        "worker_loop_start", "worker_loop_exit", "word_assembled",
        "save_blocks", "load_blocks", "write_start", "write_done",
    },
    "stress_mixed": {
        "handle_start", "write_start", "write_done",
    },
}


def main():
    ap = argparse.ArgumentParser(description="AI4Word experiment lab")
    ap.add_argument("--iter", type=int, default=1)
    ap.add_argument("--scenario", default=None)
    ap.add_argument("--seed", type=int, default=None)
    args = ap.parse_args()

    scenario = args.scenario or "write_basic"
    fn = SCENARIOS.get(scenario)
    if fn is None:
        print("unknown scenario: %s (have: %s)" % (scenario, sorted(SCENARIOS)))
        return 2
    seed = args.seed if args.seed is not None else 1000 + args.iter * 7
    rng = random.Random(seed)

    W.FAIL.clear()
    W.PLATFORM.clear()
    W.DONE["n"] = 0
    W.CHARS["n"] = 0
    W.SENT["n"] = 0
    W.DOC_BOX.update({"req": 0, "rsp": 0, "text": None})
    os.environ.setdefault("ATRIA_API_KEY", "lab-fake-key")

    debug.init(["AI4Word", "-debug"])
    # init() resolves the real log path (APPDATA, or the TEMP fallback
    # when it is not writable) and writes the run header. Wipe whatever
    # it touched, .bak included, so this iteration starts from an empty
    # log and cannot inherit a previous run's events.
    for _p in (debug.log_path(), (debug.log_path() or "") + ".bak"):
        if _p and os.path.exists(_p):
            try:
                os.remove(_p)
            except OSError:
                pass
    debug.log("lab_iteration", iter=args.iter, scenario=scenario, seed=seed)

    sc = Script()
    patch_ai(sc)
    t0 = time.time()
    try:
        worker, window, tray, settings = assemble_lab()
        worker.start()
        window.show()
        W.pump(0.5)
        W.check(worker.isRunning(), "worker thread up")
        window._on_speed(2)
        _drain(worker)
        new_worker = fn(window, worker, tray, settings, sc, rng)
        if new_worker is not None:
            worker = new_worker
    except Exception:
        tb = traceback.format_exc()
        print(tb)
        W.FAIL.append("lab: scenario raised: " + tb[:300])
    finally:
        try:
            lab_quit(worker, window)
        except Exception:
            tb = traceback.format_exc()
            print(tb)
            W.FAIL.append("lab: quit failed: " + tb[:200])

    elapsed = time.time() - t0
    findings, stats = triage(scenario, MUST_EVENTS.get(scenario, set()))
    findings = list(W.FAIL) + findings

    log_path = debug.log_path()
    tag = "iter%02d_%s" % (args.iter, scenario)
    if log_path and os.path.exists(log_path):
        try:
            shutil.copyfile(log_path, os.path.join(_RUNS, tag + ".log"))
        except OSError:
            pass
    rec = {"iter": args.iter, "scenario": scenario, "seed": seed,
           "elapsed": round(elapsed, 1), "lines": stats.get("lines", 0),
           "events": stats.get("events", 0), "fails": len(W.FAIL),
           "findings": findings, "platform": list(W.PLATFORM)}
    try:
        with io.open(os.path.join(_RUNS, "results.jsonl"), "a",
                     encoding="utf-8") as f:
            f.write(json.dumps(rec, ensure_ascii=False) + "\n")
    except OSError:
        pass

    print("LAB_RESULT iter=%d scenario=%s seed=%d fails=%d findings=%d "
          "lines=%d elapsed=%.1fs"
          % (args.iter, scenario, seed, len(W.FAIL), len(findings),
             stats.get("lines", 0), elapsed))
    for item in findings:
        print("  FINDING " + str(item).encode("unicode_escape")
              .decode("ascii")[:240])
    for item in W.PLATFORM:
        print("  PLATFORM " + str(item).encode("unicode_escape")
              .decode("ascii")[:160])
    return 0 if not findings else 1


if __name__ == "__main__":
    sys.exit(main())

