# -*- coding: utf-8 -*-
"""deep_test 并行槽编排器：把场景清单分到 K 个子进程槽并行跑，合并结果。

机制（详见 DEEP_TEST_WORKFLOW.md 第五节）：
- QTest 程序化事件发给具体控件、不经操作系统焦点，多进程可同桌面共存；
- 每槽环境变量隔离：独立日志（AI4WORD_DEBUG_LOG）、独立单实例锁
  （AI4WORD_SHM_KEY）、独立临时 settings（AI4WORD_TEST_SETTINGS）、
  各自 Dispatch 新 Word 实例（AI4WORD_TEST_FORCE_NEW_WORD，否则 GetObject
  会让所有槽写进同一篇文档）；槽内查看类命令的重连由 _attach_factory
  钉回本槽实例；
- 机器级全局状态只让槽 0 跑（全局热键、自启注册表，见 deep_test.SLOT0_ONLY）；
- 槽种子（SLOT_SEED = compact_send + review_presets_save）每槽先跑，
  保证块地图 / doc_persistence 的依赖在同槽闭合——合并报告里这两个
  场景名出现 K 次（每槽一份），属预期；
- 槽位划分由 deep_test.slot_assignment 统一计算，编排器与测试进程同源，
  所见即所分。

用法：
    .venv\\python.exe tests\\deep_parallel.py                 # 默认 3 槽
    .venv\\python.exe tests\\deep_parallel.py --slots 4       # 4 槽
    .venv\\python.exe tests\\deep_parallel.py --only interrupt
    .venv\\python.exe tests\\deep_parallel.py --seed 31337 --slots 2

合并产物在 tests\\deep_runs\\<时间戳>\\：summary.json（各槽并集）、
triage.txt（各槽拼接）、slot<N>/（各槽明细）。退出码：0 全绿 / 1 有发现 /
2 有槽崩溃（未产出 summary）。
"""
import argparse
import datetime
import json
import os
import subprocess
import sys
import time

_TESTS = os.path.dirname(os.path.abspath(__file__))
_REPO = os.path.dirname(_TESTS)
_RUNS = os.path.join(_TESTS, "deep_runs")

sys.path.insert(0, _TESTS)

# 复用分区逻辑（import deep_test 会顺带把 QT_QPA_PLATFORM 锁成 windows，
# 本进程不创建 QApplication，无副作用）。
from deep_test import SCENARIO_NAMES, SLOT_SEED, slot_assignment


def filtered_names(only, skip):
    """与 deep_test 内部一致的过滤后场景清单（保持顺序）。"""
    return [n for n in SCENARIO_NAMES
            if (not only or only in n)
            and not any(s in n for s in skip)]


def _totals_of(results):
    keys = ("fail", "notes", "missing_must", "unclassified", "environment",
            "harness_errors")
    field = {"fail": "fails", "notes": "notes", "missing_must": "missing_must",
             "unclassified": "unclassified", "environment": "environment",
             "harness_errors": "errors"}
    return {k: sum(len(r.get(field[k]) or []) for r in results)
            for k in keys}


def merge(base, per_slot, merged_results, stamp, seed, started_str):
    totals = _totals_of(merged_results)
    missing_slot = any(not rec.get("summary_ok") for rec in per_slot)
    clean = (not missing_slot
             and all(rec.get("clean") for rec in per_slot)
             and not any(totals[k] for k in ("fail", "missing_must",
                                             "unclassified", "harness_errors")))
    summary = {
        "program": "AI4Word",
        "layer": "deep_parallel (slotted parallel)",
        "started": started_str,
        "seed": seed,
        "slots": per_slot,
        "scenarios": merged_results,
        "totals": totals,
        "clean": clean,
        "exit_code": 0 if clean else (2 if missing_slot else 1),
    }
    with open(os.path.join(base, "summary.json"), "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)

    triage = ["AI4Word deep_parallel 合并分类报告（triage）"]
    triage.append("目录: %s    槽数: %d    种子: %d" % (
        base, len(per_slot), seed))
    triage.append("结果: %s    场景: %d    FAIL %d    MUST 缺失 %d    "
                  "待确认 %d    环境 %d" % (
                      "全绿" if clean else "有发现", len(merged_results),
                      totals["fail"], totals["missing_must"],
                      totals["unclassified"], totals["environment"]))
    triage.append("槽 crashed？" + (" 是（有槽未产出 summary）"
                                   if missing_slot else " 否"))
    triage.append("")
    for rec in per_slot:
        triage.append("=" * 60)
        triage.append("槽 %d   退出码 %s   归档 %s" % (
            rec["slot"], rec["exit_code"], rec["run_dir"]))
        slot_triage = os.path.join(rec["run_dir"], "triage.txt")
        if os.path.exists(slot_triage):
            try:
                with open(slot_triage, encoding="utf-8",
                          errors="replace") as f:
                    triage.append(f.read().rstrip())
            except OSError as e:
                triage.append("(读取该槽 triage 失败: %s)" % e)
        else:
            triage.append("(该槽未产出 triage —— 进程崩溃？退出码 %s)"
                          % rec["exit_code"])
            for r in merged_results:
                if r.get("name", "").endswith(" (slot%d)" % rec["slot"]):
                    for e in r.get("errors") or []:
                        triage.append("  [槽崩溃] " + e.splitlines()[0])
        triage.append("")
    with open(os.path.join(base, "triage.txt"), "w", encoding="utf-8") as f:
        f.write("\n".join(triage) + "\n")
    return summary


def main():
    ap = argparse.ArgumentParser(description="AI4Word deep_test 并行槽编排器")
    ap.add_argument("--slots", type=int, default=3, help="槽位数量（默认 3）")
    ap.add_argument("--only", default=None, help="只跑名字含该串的场景（透传）")
    ap.add_argument("--skip", default="", help="跳过场景（逗号分隔，透传）")
    ap.add_argument("--seed", type=int, default=1007, help="压力场景种子（透传）")
    ap.add_argument("--runs-name", default=None,
                    help="归档目录名（默认 tests/deep_runs/<时间戳>）")
    ap.add_argument("--timeout", type=int, default=5400,
                    help="单个槽的硬超时秒数（默认 90 分钟；超时即杀）")
    known, _unknown = ap.parse_known_args()

    only = known.only
    skip = {s.strip() for s in known.skip.split(",") if s.strip()}
    slots = max(1, known.slots)
    py = os.path.join(_REPO, ".venv", "Scripts", "python.exe")
    if not os.path.exists(py):
        py = sys.executable        # 兜底：用当前解释器（可能不是 .venv）

    t0 = time.time()
    stamp = known.runs_name or datetime.datetime.fromtimestamp(
        t0).strftime("%Y%m%d-%H%M%S")
    base = os.path.join(_RUNS, stamp)
    os.makedirs(base, exist_ok=True)

    names = filtered_names(only, skip)
    # 过滤后若只剩槽种子场景，多槽只会重复跑同一批种子——退回单槽。
    if slots > 1 and set(names) and set(names) <= set(SLOT_SEED):
        print("注意：过滤后只剩槽种子场景（%s），退回单槽执行"
              % ", ".join(names))
        slots = 1
    assigns = {i: sorted(slot_assignment(names, i, slots))
               for i in range(slots)}
    launch = {i: a for i, a in assigns.items() if a}
    print("=== deep_parallel：%d 个场景 × %d 槽（实际启动 %d 个槽）==="
          % (len(names), slots, len(launch)))
    for i in sorted(launch):
        print("  槽 %d：%s" % (i, ", ".join(launch[i]) or "(空)"))
    if not launch:
        print("[参数错误] 过滤后没有场景可跑")
        return 2

    started_str = datetime.datetime.fromtimestamp(t0).strftime(
        "%Y-%m-%d %H:%M:%S")

    # ---- 启动各槽子进程 ----
    procs = []
    for i in sorted(launch):
        sdir = os.path.join(base, "slot%d" % i)
        os.makedirs(sdir, exist_ok=True)
        env = dict(os.environ)
        env["AI4WORD_DEBUG_LOG"] = os.path.join(sdir, "debug.log")
        env["AI4WORD_SHM_KEY"] = "AI4Word-SingleInstance-v8-deep-%s-%d" % (
            stamp, i)
        env["AI4WORD_TEST_SETTINGS"] = os.path.join(sdir, "settings.json")
        # 并行槽必须每槽开自己的 Word 实例（否则 GetObject 会全部附加到
        # 同一个 Word 进程的同一个 ActiveDocument，互相写同一篇文档）
        env["AI4WORD_TEST_FORCE_NEW_WORD"] = "1"
        cmd = [py, "-u", os.path.join(_TESTS, "deep_test.py"),
               "--slot", "%d/%d" % (i, slots),
               "--runs-name", stamp, "--seed", str(known.seed)]
        if only:
            cmd += ["--only", only]
        if skip:
            cmd += ["--skip", ",".join(sorted(skip))]
        console = open(os.path.join(sdir, "console.log"), "w",
                       encoding="utf-8")
        print("\n--- 槽 %d 启动：%s" % (i, " ".join(cmd[1:])))
        procs.append((i, subprocess.Popen(cmd, cwd=_REPO, env=env,
                                          stdout=console,
                                          stderr=subprocess.STDOUT),
                      console))

    # ---- 等待 ----
    codes = {}
    for i, p, console in procs:
        try:
            codes[i] = p.wait(timeout=known.timeout)
        except subprocess.TimeoutExpired:
            p.kill()
            codes[i] = "timeout(%ds)" % known.timeout
        finally:
            console.close()
    print("\n=== 各槽退出码：%s ===" % ", ".join(
        "槽%d=%s" % (i, codes[i]) for i in sorted(codes)))

    # ---- 合并 ----
    merged_results = []
    per_slot = []
    for i in sorted(launch):
        sdir = os.path.join(base, "slot%d" % i)
        ssum = os.path.join(sdir, "summary.json")
        rec = {"slot": i, "exit_code": codes[i], "run_dir": sdir,
               "scenarios": sorted(launch[i])}
        if os.path.exists(ssum):
            try:
                with open(ssum, encoding="utf-8") as f:
                    data = json.load(f)
                rec["clean"] = bool(data.get("clean"))
                rec["totals"] = data.get("totals", {})
                rec["summary_ok"] = True
                merged_results.extend(data.get("scenarios") or [])
            except (OSError, ValueError) as e:
                rec["clean"] = False
                rec["summary_ok"] = False
                rec["error"] = "summary.json 解析失败: %s" % e
        else:
            rec["clean"] = False
            rec["summary_ok"] = False
            rec["error"] = "槽未产出 summary（进程崩溃？退出码 %s）" % codes[i]
        if not rec.get("summary_ok"):
            # 给该槽本应负责的每个场景记一条崩溃结果（保持清单可追踪）
            for n in launch[i]:
                merged_results.append({
                    "name": "%s (slot%d)" % (n, i), "ok": False,
                    "fails": [], "notes": [], "missing_must": [],
                    "unclassified": [], "environment": [],
                    "errors": [rec.get("error", "槽失败")],
                    "duration_s": 0.0})
        per_slot.append(rec)

    summary = merge(base, per_slot, merged_results, stamp, known.seed,
                    started_str)
    print("合并目录: " + base)
    print("合并 triage: " + os.path.join(base, "triage.txt"))
    print("结果: %s（退出码 %s）" % (
        "全绿" if summary["clean"] else "有发现", summary["exit_code"]))
    return summary["exit_code"]


if __name__ == "__main__":
    sys.exit(main())
