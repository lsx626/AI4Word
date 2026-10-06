# AI4Word 深度测试工作流

> **全真实驱动测试程序所有功能、按钮、机制 → 检查日志 → 修复问题 → 再次测试**，直到全绿（或连续两轮无进展）。
> 闭环的自动化形态是 Claude Code dynamic workflow（`claude-workflows/deep-test.js`，双击 `install_workflow.bat` 安装为 `/deep-test` 命令）。
> 与 `EXPERIMENT_LOG.md` 那套**离线实验台**（offscreen Qt + fake Word + 脚本化 AI）的关系：本工作流是它的**全真实收尾层**——真实窗口、真实点击、真实 Word COM、真实 Atria API。两者共用 `tests/debug_walkthrough.py` 的装配 / 等待原语，告警分类规则一脉相承。

---

## 一、四个固定步骤

1. **测试**：`tests/deep_test.py` 在真实平台上装配真实的 `MainWindow` + `AgentWorker` + `Tray` + `Settings`（临时 settings，不污染用户配置），以真实输入驱动 21 个场景（见下表）。全程开 `-debug`，每轮归档到 `tests/deep_runs/<时间戳>/`（`debug.log` + `summary.json` + `triage.txt`）。退出码：`0` 全绿 / `1` 有发现 / `2` 环境错误。
2. **分析**：`triage.txt` 把每个场景时间窗内的 `WARN`/`ERROR`/traceback 分为**真实 bug 候选 / 环境波动 / 平台假象 / 预期故障路径**（后者为场景注入的故障：`stmt_failed`、`gen_code_failed`、`stream_error` 等）；同时校验 **MUST_EVENTS**——功能跑了但日志没记 = 接线缺失 = 真实 bug。workflow 形态下由两个并行子智能体（分诊员 + 覆盖审计员）做这一步。
3. **修复**：只修真实 bug；每处修复配永久回归单测（`tests/test_*.py`，优先离线可跑）并现场跑通。项目铁律：内容编辑走块原语（`replace_block` / `insert_after` / `delete_block`），格式修改改样式定义（`doc.Styles`），所有 Word COM 调用只在 worker 线程，编辑原语执行后 `rebuild_ranges`。
4. **复测**：先 `--only <子串>` 定向重跑失败场景（公共前缀取够精确，如 `interrupt_r` / `interrupt_k` / `interrupt_e`），全部通过后全量重跑。未全绿回到第 2 步；连续两轮同一份真实问题清单 = 无进展，停下交给人工。

## 二、场景矩阵（21 个全真实场景）

| 场景 | 覆盖面 | MUST_EVENTS 抽样 |
|------|--------|------------------|
| `speed_gears` | 打字档位三按钮（慢/自/快）：下发 worker + 持久化 | `speed_clicked`×3 `speed_set`×3 |
| `compact_send` | 紧凑输入框真实回车发送 → 真实流式写作 → 首次连接 Word | `word_assembled` `write_start` `feed` `flush` `block_register` `write_done` |
| `block_map` | 块地图按钮开合、列表刷新、块点击定位 | `blockmap_toggled`×2 `block_clicked` `select_block` |
| `interrupt_rollback` | 长文生成中第二个真实回车中断 → 点真实「回滚」→ byte-exact 恢复 | `interrupted_ui` `choice_made` `rollback_ok` |
| `interrupt_keep` | 同上，点真实「保留」 | `choice_made` |
| `interrupt_extra` | 点真实「追加补充…」→ 面板输入 + 回车提交追加 | `enter_extra_mode` `extra_submitted` `extra_done` |
| `panel_send` | 展开按钮 → 面板输入框真实回车发送 | `send_panel` `write_sent` `write_done` |
| `keys` | Enter 发送 / Shift+Enter 换行 / IME 合成中 Enter / 关窗隐藏 / 托盘召回 / 忙时存档与预设被拒 | `close_to_tray` `tray_activated`×2 `save_rejected_busy` |
| `arrange_ok` | 排版模式 + 回车 → 真实 AI 生成代码 → 逐语句执行 | `gen_code` `run_code_start` `stmt_ok` `arrange_applied` |
| `arrange_repair` | 注入 `x = 1/0` 失败 → `fix_code` 自我修复 → 重跑成功 | `stmt_failed` `fix_code` `arrange_applied` |
| `arrange_gen_fail` | 注入 `gen_code` 抛错 → 失败上报 | `gen_code_failed` `arrange_failed` |
| `stream_fault` | 注入一次流式抛错 → worker 自恢复回 idle | `stream_error` `write_done` |
| `review_presets_save` | 修订按钮两次真实点击、四个预设真实选中、存档真实点击 | `review_toggled`×2 `preset_applied`×4 `save_blocks` |
| `expand_collapse` | 头像 / 展开 / 收起真实鼠标点击 | `toggle_expand` |
| `settings_dialog` | 设置对话框开 / 改 / 确定 / 保存失败回退 / 取消 / 自启失败回退 | `settings_open` `settings_closed` |
| `tray_actions` | 托盘菜单「显示」「开机自启」勾选与还原 | `autostart_toggle`×2 |
| `drag_snap_hotkey` | 真实拖动落点、贴边吸附、召唤、Ctrl+Alt+Space 热键注册（真实平台必须成功） | `snap_to_edge` `summon` `hotkey_register` |
| `esc_hide_hint` | Esc 收起 / Esc 隐藏到托盘 / 首次隐藏提示 | `hide_via_esc` |
| `input_rejections` | 空文本不生成命令、超长输入被拒 | `send_rejected_too_long` |
| `doc_persistence` | 存档真实落进文档变量 → 抹掉引擎绑定 → 下一条命令重连装配、恢复块索引 | `save_blocks` `word_assembled`×2 `load_blocks` |
| `stress_mixed` | 以上矩阵按种子随机组合 3 个（确定性可复现） | `queue_put` |

> 注入故障（`x = 1/0`、`gen_code`/`ai_stream` 抛错）模拟用户真实使用中的失败；triage 里标为「预期故障路径」不算 bug，它们验证的**自我修复 / 失败上报机制**才是被测对象。

## 三、日志检查规则

- 日志在 `-debug` 模式下写 `%APPDATA%\AI4Word\debug.log`（5MB 滚动为 `.bak`）；未捕获异常由 `sys.excepthook` / `threading.excepthook` 落盘，`uncaught_exception` / `uncaught_thread_exception` 一律视为真实 bug。
- 每轮归档：`debug.log` 是当轮拷贝；`triage.txt` 人类可读；`summary.json` 机器可读（场景、FAIL、MUST 缺失、告警计数、`exit_code`），是复测与 workflow 分析的输入。
- 分类口径：**真实 bug 候选**＝场景时间窗内的 WARN/ERROR 且非注入故障、非环境重试噪音（`api_*_retry` / `api_*_failed` / `word_connect_failed`）；**环境波动**＝网络 / API / Word 未运行；**平台假象**＝真实平台下仍被环境拦截（如组策略拦自启注册表）；**MUST 缺失**＝流程走完但事件没打。
- 真实平台下没有「offscreen 无 HWND」的借口：热键注册、拖动吸附、托盘的真实性都提高一个量级——这是本层相对离线实验台的价值。

## 四、怎么跑

**人工（只跑测试层，自己分析）**

```powershell
.\run_deep_test.bat                      # 全量（单进程）
.\run_deep_test_parallel.bat             # 全量（3 槽并行，约 2.5-3× 提速）
.\run_deep_test_parallel.bat --slots 2   # 内存吃紧用 2 槽
.\run_deep_test.bat --only interrupt     # 只跑名字含 interrupt 的场景
.\run_deep_test.bat --seed 31337 --keep-word
```

前提：`.venv` + `.env`（`ATRIA_API_KEY`）、退出正在运行的 AI4Word、测试期间不碰鼠标键盘。跑完看 `tests\deep_runs\` 最新目录的 `triage.txt` / `summary.json`。

**全自动闭环（Claude Code 工作流）**

```powershell
.\install_workflow.bat      # 安装一次 .claude\workflows\deep-test.js
claude                      # 在仓库目录启动 Claude Code
/deep-test                  # 测试 → 分析 → 修复 → 复测，最多 5 轮
/deep-test rounds=3 only=interrupt
```

workflow 脚本按 Claude Code 官方 [dynamic workflows](https://code.claude.com/docs/en/workflows) 形态编写：脚本持有循环与中间结果，子智能体只回结构化结果；测试 / 复测直接 `.venv\python.exe -u tests\deep_parallel.py`（本机 `.venv` 是 conda 布局、解释器在 `.venv\python.exe`，Claude Code 的 bash 里 `cmd /c` 会被 MSYS 参数转义破坏，所以不走 `.bat`；人工双击不受影响）；分析阶段并行跑分诊员 + 覆盖审计员；修复阶段按疑点文件分组——同文件多条发现交给同一个 agent 顺序修，不同文件并行修。

**并行机制**

- 能并行是因为 QTest 程序化事件发给具体控件、不经操作系统焦点，多个 QApplication 进程可同桌面共存。
- 每槽环境变量隔离：`AI4WORD_DEBUG_LOG`（否则互相覆盖共享日志）、`AI4WORD_SHM_KEY`（单实例锁）、`AI4WORD_TEST_SETTINGS`、`AI4WORD_TEST_FORCE_NEW_WORD`（否则 GetObject 会让所有槽附到同一个 Word 进程写同一篇文档）。
- 机器级全局状态只在槽 0 跑：`drag_snap_hotkey`（全局热键）与 `tray_actions`（写注册表）；槽种子 `compact_send` + `review_presets_save` 每槽先跑，保证 `block_map` / `doc_persistence` 的依赖在同槽闭合（合并报告里这两个场景名出现 K 次属预期）。
- 代价：每槽约 200-400MB（一个 Word 实例）；同桌面偶发竞争的假告警被 triage 归入「环境告警」兜底；任一槽崩溃（未产出 summary）合并退出码为 2。`--only` 命中少量场景时只启动需要的槽。

## 五、产物布局

```
tests/deep_runs/<时间戳>/            单进程 / 并行合并目录
├── summary.json     {clean, exit_code, totals, scenarios:[{name, ok, fails, missing_must, ...}]}
├── triage.txt       分类报告（FAIL / MUST 缺失 / 待确认 / 环境 / 平台假象）
└── slot0/ slot1/ slot2/   并行模式各槽明细（debug.log / console.log / summary.json / triage.txt / settings.json）
```

## 六、退出码

| 退出码 | 含义 |
|--------|------|
| `0` | 全绿：无 FAIL、无 MUST 缺失、无待确认告警 |
| `1` | 有发现（FAIL / MUST 缺失 / 待确认告警任一非零） |
| `2` | 环境错误（缺 `ATRIA_API_KEY`、单实例锁、harness 异常、无 `.venv` 等） |
