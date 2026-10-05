# AI4Word - 深度实验工作流日志（EXPERIMENT_LOG）

> 本文件记录“实验 → 检查/分析日志 → 修复 → 再实验”迭代工作流的过程与结论。
> 隐藏调试模式本身的用法见 `app/debug.py` 与 DEVELOPMENT_LOG.md 的 V9.3 节。
> 全部实验产物（每轮归档日志 + 汇总 jsonl）在 `tests/lab_runs/`。

## 一、工作流定义

每一轮迭代包含四个固定步骤：

1. **实验（Experiment）**：`tests/experiment_lab.py` 以 offscreen Qt + fake Word +
   脚本化 AI 在单进程内装配真实的 `MainWindow` + `AgentWorker` + `Settings`，
   按场景矩阵驱动全部功能 / 按钮 / 页面（11 个场景，见下表）。
2. **分析（Analyze）**：解析本轮 `debug.log`，列出全部 `WARN` / `ERROR` / traceback，
   逐条定性为三类之一：**真实 bug** / **平台假象**（offscreen 无真实 HWND 等）/
   **预期分支**（场景注入的故障路径）。同时校验 `MUST_EVENTS`：该场景必须打出的
   调试事件，用于暴露“接线缺失”（功能跑了但日志没记录）。
3. **修复（Fix）**：只修真实 bug；每处修复配永久单测（`tests/test_*.py`），
   并先验证该单测在旧行为下失败、新行为下通过。
4. **再实验（Re-experiment）**：用原失败种子重跑 + 新种子继续马拉松，
   直到该轮 0 findings、入口 exit code = 0。

**收尾验证（real-stack）**：`tests/debug_walkthrough.py` 使用真实 Word + 真实 Atria
API（`.env` 密钥）跑同样的功能矩阵；再跑对照组（不带 `-debug`），断言全程不生成
任何日志文件。

## 二、实验矩阵（11 场景）

| 场景 | 覆盖面 | 必须出现的日志事件（MUST_EVENTS 抽样） |
|------|--------|---------------------------------------|
| `write_basic` | 紧凑发送面板、流式逐字写入、块注册、存档 save_blocks | stream_start / stream_progress / write_block / flush / save_blocks |
| `interrupts` | 中断 → 回滚 / 保留 / 追加补充三支路 + 修订开关在生成中被拒 | interrupt / wait_choice / rollback / keep / append |
| `arrange_ok` | 排版链路正常路径：gen_code → run_code → 逐语句执行 | gen_code / run_code / arrange_done |
| `arrange_repair` | 排版自我修复：x = 1/0 语句失败 → fix_code → 重跑成功 | stmt_failed / fix_code / arrange_done |
| `arrange_gen_fail` | 排版生成失败：gen_code 抛错 → 失败上报 | gen_code_failed / arrange_failed |
| `stream_faults` | 流式故障注入：stream_error / gen_code_failed 各分支 | stream_error / 中断选择路径 |
| `review_presets_save` | 修订模式开关 + 四个预设 + 存档恢复 | review_toggle / preset_applied / save_blocks |
| `settings_dialog` | 设置对话框开 / 改 / 确定 / 保存失败回退 + 自启切换 | settings_open / settings_changed / 自启切换 |
| `keys_tray` | 热键注册/注销、托盘菜单（显示 / 自启 / 退出）、Esc 隐藏 | hotkey_register / tray 动作 |
| `doc_persistence` | 文档重绑定 + 块导入 import_document + 持久化 | doc_rebind / import_document / save_blocks |
| `stress_mixed` | 综合压力：以上矩阵随机组合 + 多轮回滚 byte-exact 校验 | 全部关键事件按随机矩阵抽检 |

## 三、迭代表（99 轮离线实验）

| 指标 | 数值 |
|------|------|
| 离线实验总轮数 | 99 |
| 其中失败轮数 | 19（失败即修复，见下表“备注”列） |
| 场景分布 | stress_mixed 37 轮（综合压力优先），其余 10 场景各 5-8 轮 |

| 迭代 | 场景 | 种子 | 结果 | 备注 |
|------|------|------|------|------|
| 1 | write_basic | 1007 | 失败 | harness 摇落期：AI 0 字符流 + 面板写入未落盘 + 事件接线缺失 |
| 1 | write_basic | 1007 | 失败 | harness 摇落期：AI 0 字符流 + 面板写入未落盘 + 事件接线缺失 |
| 1 | write_basic | 1007 | 通过 | harness 摇落期：AI 0 字符流 + 面板写入未落盘 + 事件接线缺失 |
| 2 | interrupts | 1014 | 通过 | triage 提醒：缺 arrange_failed 事件；注入故障噪音 |
| 2 | arrange_ok | 1014 | 通过 | triage 提醒：缺 arrange_failed 事件；注入故障噪音 |
| 2 | arrange_repair | 1014 | 通过 | triage 提醒：缺 arrange_failed 事件；注入故障噪音 |
| 2 | arrange_gen_fail | 1014 | 通过（triage 提醒） | triage 提醒：缺 arrange_failed 事件；注入故障噪音 |
| 2 | stream_faults | 1014 | 通过（triage 提醒） | triage 提醒：缺 arrange_failed 事件；注入故障噪音 |
| 3 | arrange_gen_fail | 2001 | 通过 |  |
| 3 | stream_faults | 2002 | 通过 |  |
| 4 | review_presets_save | 2011 | 通过 |  |
| 5 | settings_dialog | 2012 | 通过（triage 提醒） | triage 提醒：缺 settings_open 事件（接线缺失，已补） |
| 5 | settings_dialog | 2012 | 通过 | triage 提醒：缺 settings_open 事件（接线缺失，已补） |
| 6 | keys_tray | 2013 | 失败 | 拖动移动 / 边缘吸附失败（harness 时序与动画冲突，后修） |
| 6 | keys_tray | 2013 | 通过 | 拖动移动 / 边缘吸附失败（harness 时序与动画冲突，后修） |
| 7 | doc_persistence | 2014 | 通过 |  |
| 8 | stress_mixed | 3001 | 通过 |  |
| 8 | stress_mixed | 3001 | 通过 |  |
| 9 | write_basic | 4100 | 通过 |  |
| 10 | interrupts | 4101 | 通过 |  |
| 11 | arrange_ok | 4102 | 通过 |  |
| 12 | arrange_repair | 4103 | 通过 |  |
| 13 | arrange_gen_fail | 4104 | 通过 |  |
| 14 | stream_faults | 4105 | 通过 |  |
| 15 | review_presets_save | 4106 | 通过 |  |
| 16 | settings_dialog | 4107 | 通过 |  |
| 17 | keys_tray | 4108 | 通过 |  |
| 7 | doc_persistence | 2014 | 通过 |  |
| 20 | write_basic | 5100 | 通过 |  |
| 21 | interrupts | 5101 | 通过 |  |
| 22 | arrange_ok | 5102 | 通过 |  |
| 23 | arrange_repair | 5103 | 通过 |  |
| 24 | arrange_gen_fail | 5104 | 通过 |  |
| 25 | stream_faults | 5105 | 通过 |  |
| 26 | review_presets_save | 5106 | 通过 |  |
| 27 | settings_dialog | 5107 | 通过 |  |
| 28 | keys_tray | 5108 | 通过 |  |
| 29 | doc_persistence | 5109 | 通过 |  |
| 30 | stress_mixed | 5110 | 通过 |  |
| 31 | stress_mixed | 6100 | 失败 | 回滚未恢复写前文档 → 块顺序错位 bug |
| 32 | stress_mixed | 6113 | 失败 | 回滚未恢复写前文档 → 块顺序错位 bug |
| 33 | stress_mixed | 6126 | 失败 | 两个回滚周期失败 → 块顺序错位 bug |
| 34 | stress_mixed | 6139 | 通过 |  |
| 35 | stress_mixed | 6152 | 失败 | 回滚未恢复写前文档 → 块顺序错位 bug |
| 36 | stress_mixed | 6165 | 失败 | 回滚未恢复写前文档 → 块顺序错位 bug |
| 37 | stress_mixed | 6178 | 失败 | 回滚未恢复写前文档 → 块顺序错位 bug |
| 38 | stress_mixed | 6191 | 失败 | 回滚未恢复写前文档 → 块顺序错位 bug |
| 39 | stress_mixed | 6204 | 通过 |  |
| 40 | stress_mixed | 6217 | 失败 | 回滚未恢复写前文档 → 块顺序错位 bug |
| 31 | stress_mixed | 6100 | 通过 | 回滚未恢复写前文档 → 块顺序错位 bug |
| 99 | stress_mixed | 6113 | 失败 | 诊断重跑：复现回滚 bug 的经典种子 |
| 99 | stress_mixed | 6126 | 失败 | 诊断重跑：复现回滚 bug 的经典种子 |
| 99 | stress_mixed | 6152 | 通过 |  |
| 99 | stress_mixed | 6165 | 通过 |  |
| 99 | stress_mixed | 6178 | 通过 |  |
| 99 | stress_mixed | 6191 | 失败 | 诊断重跑：复现回滚 bug 的经典种子 |
| 51 | stress_mixed | 6113 | 失败 | 块顺序修复后复现：首尾空白 bug |
| 99 | stress_mixed | 6217 | 通过 |  |
| 52 | stress_mixed | 6113 | 失败 | 块顺序修复后复现：首尾空白 bug |
| 53 | stress_mixed | 6113 | 失败 | 块顺序修复后复现：首尾空白 bug |
| 54 | stress_mixed | 6113 | 失败 | 块顺序修复后复现：首尾空白 bug |
| 55 | stress_mixed | 6113 | 通过 |  |
| 56 | stress_mixed | 6928 | 通过 |  |
| 57 | stress_mixed | 6941 | 通过 |  |
| 58 | stress_mixed | 6954 | 通过 |  |
| 59 | stress_mixed | 6967 | 通过 |  |
| 60 | stress_mixed | 6980 | 通过 |  |
| 61 | stress_mixed | 6993 | 失败 | 首个 1 字符回滚误差 → 首尾空白 bug（iter61 日志定位根因） |
| 62 | stress_mixed | 7006 | 通过 |  |
| 63 | stress_mixed | 7019 | 通过 |  |
| 64 | stress_mixed | 7032 | 通过 |  |
| 65 | write_basic | 8105 | 通过 |  |
| 66 | interrupts | 8122 | 通过 |  |
| 67 | arrange_ok | 8139 | 通过 |  |
| 68 | arrange_repair | 8156 | 通过 |  |
| 69 | arrange_gen_fail | 8173 | 通过 |  |
| 70 | stream_faults | 8190 | 通过 |  |
| 71 | review_presets_save | 8207 | 通过 |  |
| 72 | settings_dialog | 8224 | 通过 |  |
| 73 | keys_tray | 8241 | 通过 |  |
| 74 | doc_persistence | 8258 | 通过 |  |
| 75 | stress_mixed | 8275 | 通过 |  |
| 76 | write_basic | 8292 | 通过 |  |
| 77 | interrupts | 8309 | 通过 |  |
| 78 | arrange_ok | 8326 | 通过 |  |
| 79 | arrange_repair | 8343 | 通过 |  |
| 80 | arrange_gen_fail | 8360 | 通过 |  |
| 81 | stream_faults | 8377 | 通过 |  |
| 82 | review_presets_save | 8394 | 通过 |  |
| 83 | settings_dialog | 8411 | 通过 |  |
| 84 | keys_tray | 8428 | 通过 |  |
| 85 | doc_persistence | 8445 | 通过 |  |
| 86 | stress_mixed | 8462 | 通过 |  |
| 87 | write_basic | 8479 | 通过 |  |
| 88 | interrupts | 8496 | 通过 |  |
| 89 | arrange_ok | 8513 | 通过 |  |
| 90 | arrange_repair | 8530 | 通过 |  |
| 91 | arrange_gen_fail | 8547 | 通过 |  |
| 92 | stream_faults | 8564 | 通过 |  |

## 四、发现并修复的问题（按时间顺序）

| # | 问题 | 性质 | 位置 | 修复方式 |
|---|------|------|------|----------|
| 1 | 多轮实验共用一个 TEMP `debug.log`，日志互相污染 | harness | `tests/experiment_lab.py` | 每轮开头删除候选日志路径，结束后归档到 `tests/lab_runs/iterNN_场景.log`；并规定禁止两个 lab 进程并行 |
| 2 | 排版链路缺调试事件 | 接线缺失 | `app/engine.py` `_cmd_arrange` | 补齐 gen_code / run_code / fix_code / arrange_failed 事件 |
| 3 | 设置对话框缺打开事件 | 接线缺失 | `app/settings_dialog.py`、`app/main_window.py` | 补 `settings_open` / 保存失败事件 |
| 4 | 拖动落点被输入动画吞掉 | 测试时序 / 真实缺陷 | `app/main_window.py` + harness | 拖动落点与输入动画时序修正 |
| 5 | `wait_handled` 的 SENT 计数竞态 | harness | `tests/debug_walkthrough.py` | 修正发送侧计数与完成计数的竞态 |
| 6 | 流式星号 marker 误判（闭合强调后随空白被当字面量） | 真实 bug（字节一致性） | `streaming_writer.py` `_resolve_markers` | marker 后随空白且强调已打开时按闭合处理，使流式与解析路径字节一致（`tests/test_streaming_markers.py`） |
| 7 | **块顺序错位 → 回滚重放错误** | **真实 bug** | `doc_model.py` | `register()` 改为按文档位置插入（新增 `_insert_index` / `_block_pos`）；`replace_block`/`insert_after` 由“尾切片”改为按 `write_block` 返回值身份收集；`snapshot()` 记录 `pos`，`restore_snapshot()` 稳定排序后重放（`tests/test_block_order.py`，3 用例） |
| 8 | **流式段落首尾空白 → 回滚差 1 字符** | **真实 bug** | `streaming_writer.py` `_commit_draft` | 新增 `_trim_draft_edges`：提交时删除文档中键入的空白边缘并同步 plain/md，使 feed+flush 与 write_block 字节一致、回滚 byte-exact（`tests/test_draft_edges.py`，3 用例） |

bug #7 与 #8 是同一条不变量的两面：**流式逐字写入的文档内容必须与 `write_block()`
解析路径产生的文档内容逐字节一致**，因为回滚经 `restore_snapshot()` →
`write_block()` 重放。修复前，中段写入的块在模型里落到错误位置（#7），或段落
边缘的空格留在文档却被 md 丢弃（#8），都会让回滚后的文档与写前文档不等。

## 五、良性白名单（triage 定性结论）

以下 WARN / ERROR / traceback 在日志中出现属**预期**，不计为 bug：

- `hotkey_register ok=false` / `hotkey_register_failed`：offscreen 平台无真实
  HWND，真机上正常注册。
- `settings_load_first_run`：首次运行无 settings.json，预期 INFO。
- `stmt_failed idx=1 src="x = 1 / 0"`：`arrange_repair` 场景**故意注入**的除零
  语句，用于检验排版自我修复链路。
- `gen_code_failed`：`arrange_gen_fail` / `stream_faults` 注入的生成失败。
- `stream_error`：`stream_faults` 注入的流式故障。
- `run_code_sandbox_rejected error="禁止导入: Import"`：沙箱安全设计，AI 生成
  代码含 Import 被正确拒绝。
- `autostart_disable_failed [WinError 2]` / `settings_save_failed
  no_such_dir_ai4word`：`settings_dialog` 场景**故意**指向不存在的目录以检验
  失败回退路径。

## 六、最终结果

| 验证层 | 结果 |
|--------|------|
| 离线马拉松 | iter 62-92 连续 **31 轮 0 findings**；末轮 64-92 为最后一次修复（trailing space）后连续 **29 轮**全部通过，覆盖全部 11 场景 |
| 单元测试 | `pytest` **110 passed**（含本次新增 `tests/test_draft_edges.py` 3 用例） |
| 真机走查（`-debug`，真实 Word + Atria API） | **失败 0**；平台假象 1 条（offscreen 热键注册）；1226 行日志中 7 条 WARN/ERROR 全部为白名单内的注入故障 |
| 对照组（不带 `-debug`） | **失败 0**，且全程**未生成任何日志文件**（程序内置断言 + 外部双重确认） |

## 七、复现方式

```powershell
# 单轮离线实验（exit code 0 = 干净）
$env:QT_QPA_PLATFORM="offscreen"
.\.venv\python.exe tests\experiment_lab.py --iter 64 --scenario stress_mixed --seed 7032

# 归档日志位置：tests/lab_runs/iter64_stress_mixed.log
# 汇总记录：    tests/lab_runs/results.jsonl

# 全量单测（.venv 无 pytest，用 anaconda 的解释器）
D:\ProgramDatanaconda3\python.exe -m pytest -q -p no:cacheprovider

# 真机走查（真实 Word + Atria API；需要 .env 里的 ATRIA_API_KEY）
$env:QT_QPA_PLATFORM="offscreen"
.\.venv\python.exe tests\debug_walkthrough.py -debug     # 调试模式（生成日志）
.\.venv\python.exe tests\debug_walkthrough.py             # 对照组（不得生成日志）
```

注意：禁止并行运行两个 lab / walkthrough 进程（共用同一个 `debug.log`，会互相
污染本轮日志）。对照组运行前先删除 `%APPDATA%\AI4Word\debug.log`。
