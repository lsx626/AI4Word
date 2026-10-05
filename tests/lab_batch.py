# -*- coding: utf-8 -*-
"""Batch driver for the deep-experiment workflow.

One command runs a whole wave of lab iterations and then triages every
archived log of the wave:

    python tests/lab_batch.py --from 94 --to 104
    python tests/lab_batch.py --from 94 --to 104 --triage-only

Each iteration is a separate process (a crash or leaked Qt thread can
never poison the rest of the wave). The consolidated triage reports, per
iteration and aggregated over the wave:

    * findings recorded by the lab itself (W.FAIL + missing MUST events)
    * every WARN/ERROR event, split into whitelisted vs UNEXPECTED
    * any traceback text or sys.excepthook fire in the logs

UNEXPECTED lines are the workflow's work items: investigate, fix, then
re-run the wave (or --triage-only after a code-only change) until the
report is clean. Only then is the wave considered done.
"""
import argparse
import collections
import glob
import io
import json
import os
import random
import re
import subprocess
import sys

_TESTS = os.path.dirname(os.path.abspath(__file__))
_REPO = os.path.dirname(_TESTS)
_RUNS = os.path.join(_TESTS, "lab_runs")

sys.path.insert(0, _TESTS)
sys.path.insert(0, _REPO)

import experiment_lab as lab  # noqa: E402

EVENT_RE = re.compile(r"\b(INFO|WARN|ERROR)\s+(\w+)")
WARN_RE = re.compile(r"\b(WARN|ERROR)\s+(\w+)")

BENIGN_ALWAYS = lab.BENIGN_ALWAYS
BENIGN_BY_SCENARIO = lab.BENIGN_BY_SCENARIO


def rotate_scenarios():
    order = ["write_basic", "interrupts", "arrange_ok", "arrange_repair",
             "arrange_gen_fail", "stream_faults", "review_presets_save",
             "settings_dialog", "keys_tray", "doc_persistence",
             "stress_mixed"]
    return order


def run_one(iter_no, scenario, seed):
    cmd = [sys.executable, os.path.join(_TESTS, "experiment_lab.py"),
           "--iter", str(iter_no), "--scenario", scenario,
           "--seed", str(seed)]
    env = dict(os.environ)
    env["QT_QPA_PLATFORM"] = "offscreen"
    proc = subprocess.run(cmd, cwd=_REPO, env=env,
                          stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                          timeout=600)
    out = proc.stdout.decode("utf-8", "replace")
    return proc.returncode, out


def parse_log(path):
    """-> (warn_events, has_traceback, hook_fired, lines)"""
    raw = io.open(path, encoding="utf-8", errors="replace").read()
    lines = raw.splitlines()
    warns = collections.Counter()
    for ln in lines:
        m = WARN_RE.search(ln)
        if m:
            warns[m.group(2)] += 1
    # exc() embeds its traceback inside a quoted field on the same line
    # (handled exception -> fine). A bare dump at line start means an
    # unformatted crash, which is a real work item.
    bare_tb = any(ln.lstrip().startswith("Traceback") for ln in lines)
    return (warns, bare_tb, "excepthook" in raw, len(lines))


def load_results(iter_lo, iter_hi):
    path = os.path.join(_RUNS, "results.jsonl")
    rows = []
    if os.path.exists(path):
        with io.open(path, encoding="utf-8") as f:
            for ln in f:
                try:
                    r = json.loads(ln)
                except ValueError:
                    continue
                if iter_lo <= r.get("iter", -1) <= iter_hi:
                    rows.append(r)
    return {r["iter"]: r for r in rows}


def triage_wave(iter_lo, iter_hi):
    rows = load_results(iter_lo, iter_hi)
    agg_benign = collections.Counter()
    agg_unexpected = collections.Counter()
    unexpected_by_iter = collections.defaultdict(list)
    problems = []
    covered = []
    for it in range(iter_lo, iter_hi + 1):
        logs = glob.glob(os.path.join(_RUNS, "iter%02d_*.log" % it))
        if not logs:
            problems.append("iter %d: no archived log" % it)
            continue
        rec = rows.get(it)
        scenario = rec["scenario"] if rec else "?"
        warns, tb, hook, nlines = parse_log(logs[0])
        covered.append((it, scenario, nlines))
        benign = BENIGN_ALWAYS | BENIGN_BY_SCENARIO.get(scenario, set())
        for ev, n in warns.items():
            if ev in benign:
                agg_benign[ev] += n
            else:
                agg_unexpected[ev] += n
                unexpected_by_iter[it].append("%s x%d" % (ev, n))
        if tb:
            problems.append("iter %d (%s): bare traceback dump in log"
                            % (it, scenario))
        if hook:
            problems.append("iter %d (%s): sys.excepthook fired" % (it, scenario))
        if rec is None:
            problems.append("iter %d: no results.jsonl row" % it)
        else:
            if rec.get("fails"):
                problems.append("iter %d (%s): lab FAIL %s"
                                % (it, scenario, rec["fails"]))
            for f in rec.get("findings", []):
                problems.append("iter %d (%s): finding %s"
                                % (it, scenario, str(f)[:220]))
    return {
        "covered": covered,
        "benign": agg_benign,
        "unexpected": agg_unexpected,
        "unexpected_by_iter": unexpected_by_iter,
        "problems": problems,
    }


def print_report(rep):
    print("=" * 62)
    print("WAVE COVERAGE: %d iterations" % len(rep["covered"]))
    for it, sc, n in rep["covered"]:
        print("  iter %3d  %-20s %5d log lines" % (it, sc, n))
    print("-" * 62)
    print("WHITELISTED WARN/ERROR (platform/fault-injection artifacts):")
    if not rep["benign"]:
        print("  (none)")
    for ev, n in sorted(rep["benign"].items()):
        print("  %-34s %4d" % (ev, n))
    print("-" * 62)
    print("UNEXPECTED WARN/ERROR (work items):")
    if not rep["unexpected"]:
        print("  (none)")
    for ev, n in sorted(rep["unexpected"].items()):
        print("  %-34s %4d" % (ev, n))
    if rep["unexpected_by_iter"]:
        for it in sorted(rep["unexpected_by_iter"]):
            print("  iter %d: %s" % (it, ", ".join(rep["unexpected_by_iter"][it])))
    print("-" * 62)
    print("OTHER PROBLEMS:")
    if not rep["problems"]:
        print("  (none)")
    for p in rep["problems"]:
        print("  " + str(p)[:300])
    print("=" * 62)
    clean = not (rep["unexpected"] or rep["problems"])
    print("WAVE RESULT: %s" % ("CLEAN" if clean else "NOT CLEAN"))
    return clean


def main():
    ap = argparse.ArgumentParser(description="AI4Word lab batch driver")
    ap.add_argument("--from", dest="lo", type=int, required=True)
    ap.add_argument("--to", dest="hi", type=int, required=True)
    ap.add_argument("--scenarios", default=None,
                    help="comma list; default rotates all scenarios")
    ap.add_argument("--triage-only", action="store_true")
    args = ap.parse_args()
    if args.lo > args.hi:
        print("--from must be <= --to")
        return 2

    if not args.triage_only:
        order = (args.scenarios.split(",") if args.scenarios
                 else rotate_scenarios())
        for i, it in enumerate(range(args.lo, args.hi + 1)):
            scenario = order[i % len(order)].strip()
            rng = random.Random(9000 + it * 17 + i)
            seed = rng.randint(1, 99999)
            rc, out = run_one(it, scenario, seed)
            tail = [ln for ln in out.splitlines() if ln.startswith("LAB_RESULT")
                    or ln.startswith("  FINDING")]
            print(" ".join(tail) if tail else
                  ("iter %d: no LAB_RESULT line (rc=%d)" % (it, rc)))
            if rc == 2:
                print(out[-1200:])
                return 2
    rep = triage_wave(args.lo, args.hi)
    clean = print_report(rep)
    return 0 if clean else 1


if __name__ == "__main__":
    sys.exit(main())
