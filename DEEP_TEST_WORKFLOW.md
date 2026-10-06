# AI4Word 深度测试工作流（DEEP_TEST_WORKFLOW）

> 工作流定义：**全真实驱动测试程序所有功能、按钮、机制 → 检查日志 → 修复问题 → 再次测试**，
> 直到全绿（或连续两轮无进展）。本文件是工作流的说明与 SOP；
> 闭环的自动化形态是 Claude Code dynamic workflow（`claude-workflows/deep-test.js`，
> 双击 `install_workflow.bat` 安装为 `/deep-test` 命令）。
>
> 与既有 `EXPERIMENT_LOG.md` 工作流的关系：那一套是**离线实验台**（offscreen Qt +
> fake Word + 脚本化 AI，99 轮迭代用）；本工作流是它的**全真实收尾层**——真实窗口、
> 真实点击、真实 Word COM、真实 Atria API。两者共用 `tests/debug_walkthrough.py`
> 的装配 / 等待 / Word 工厂原语，告警分类规则一脉相承。

---

## 一、工作流的四个固定步骤

1. **测试（Test）**：`tests/deep_test.py` 在真实平台上装配真实的 `MainWindow` +
   `AgentWorker` + `Tray` + `Settings`（临时 settings 文件，不污染用户配置），
   以**真实输入**驱动全部功能 / 按钮 / 机制（21 个场景，见下表）。全程开
   `-debug` 隐藏调试模式，日志写入 `%APPDATA%\AI4Word\debug.log`。
   每轮归档到 `tests/deep_runs/<时间戳>/`：`debug.log` + `summary.json` +
   `triage.txt`，退出码 `0`（全绿）/ `1`（有发现）/ `2`（环境错误）。

2. **分析（Analyze）**：`triage.txt` 已把每个场景时间窗内的 `WARN`/`ERROR`/
   `traceback` 分为三类——**真实 bug 候选 / 环境波动 / 平台假象 / 预期故障路径**
   （后者是场景注入的故障：`stmt_failed`、`gen_code_failed`、`stream_error`、
   `settings_save_failed`、`autostart_*_failed`）。同时校验 **MUST_EVENTS**：
   每个场景必须打出的调试事件——功能跑了但日志没记 = 接线缺失 = 真实 bug。
   workflow 形态下有两个并行子智能体（分诊员 + 覆盖审计员）做这一步。

3. **修复（Fix）**：只修真实 bug；每处修复配永久回归单测（`tests/test_*.py`，
   优先离线可跑），并现场跑通。项目铁律：内容编辑走块原语（`replace_block` /
   `insert_after` / `delete_block`），格式修改改样式定义（`doc.Styles`），
   所有 Word COM 调用只在 worker 线程，编辑原语执行后 `rebuild_ranges`，
   fake Word 必须实现真实 COM 语义。

4. **复测（Retest）**：先用 `--only <子串>` 定向重跑失败场景（公共前缀取够
   精确的子串，如 `interrupt_r` / `interrupt_k` / `interrupt_e`），全部通过后
   **全量重跑**。未全绿则回到第 2 步；连续两轮同一份真实问题清单 = 无进展，
   停下交给人工。全量退出码 `0` = 闭环完成。

## 二、场景矩阵（21 个全真实场景）

| 场景 | 覆盖面（功能 / 按钮 / 机制） | MUST_EVENTS 抽样 |
|------|------------------------------|------------------|
| `speed_gears` | 打字档位三按钮（慢/自/快）：下发 worker + 持久化 | `speed_clicked`×3 `speed_set`×3 |
| `compact_send` | 紧凑输入框**真实回车**发送 → 真实流式写作 → 首次连接 Word | `word_assembled` `write_start` `feed` `flush` `block_register` `write_done` `api_stream_done` |
| `block_map` | 块地图按钮两次点击开合、列表刷新、块点击定位 | `blockmap_toggled`×2 `block_clicked` `select_block` |
| `interrupt_rollback` | 长文生成中**第二个真实回车**中断 → 点击真实「回滚」按钮 → byte-exact 恢复 | `interrupted_ui` `choice_made` `rollback_ok` `snapshot_taken` |
| `interrupt_keep` | 同上，点真实「保留」按钮 | `choice_made` |
| `interrupt_extra` | 点真实「追加补充…」→ 面板输入 + 真实回车提交追加 | `enter_extra_mode` `extra_submitted` `extra_start` `extra_done` |
| `panel_send` | 展开按钮 → 面板输入框**真实回车**发送 | `send_panel` `write_sent` `write_done` |
| `keys` | Enter 发送 / Shift+Enter 换行不发送 / IME 合成中 Enter 不发送 / 关窗隐藏 / 托盘双击召回 / 忙时存档与预设被拒 | `close_to_tray` `tray_activated`×2 `save_rejected_busy` `preset_rejected_busy` |
| `arrange_ok` | 排版模式按钮 + 真实回车 → 真实 AI 生成代码 → 逐语句执行 | `gen_code` `run_code_start` `stmt_ok` `run_code_done` `arrange_applied` |
| `arrange_repair` | 注入 `x = 1/0` 语句失败 → `fix_code` 自我修复 → 重跑成功 | `stmt_failed` `fix_code` `run_code2_result` `arrange_applied` |
| `arrange_gen_fail` | 注入 `gen_code` 抛错 → 失败上报 | `gen_code_failed` `arrange_failed` |
| `stream_fault` | 注入一次流式抛错 → `stream_error` → worker 自恢复回 idle | `stream_error` `write_done` |
| `review_presets_save` | 修订按钮两次真实点击、四个预设真实选中、存档按钮真实点击 | `review_toggled`×2 `preset_applied`×4 `save_clicked` `save_blocks` |
| `expand_collapse` | 头像 / 展开 / 收起真实鼠标点击 | `toggle_expand` |
| `settings_dialog` | 设置对话框开/改/确定/保存失败回退/密码显隐/取消/自启失败回退 | `settings_open` `settings_closed` |
| `tray_actions` | 托盘菜单「显示」「开机自启」勾选与还原 | `autostart_toggle`×2 `autostart_disabled` |
| `drag_snap_hotkey` | 真实拖动落点、贴边吸附、召唤、Ctrl+Alt+Space 热键注册（真实平台下必须成功） | `snap_to_edge` `summon` `hotkey_register` |
| `esc_hide_hint` | Esc 收起 / Esc 隐藏到托盘 / 首次隐藏提示 | `hide_via_esc` |
| `input_rejections` | 空文本不生成命令、超长输入被拒 | `send_rejected_too_long` |
| `doc_persistence` | 存档真实落进文档变量（worker 线程读取）→ 抹掉引擎绑定 → 下一条命令重连装配、恢复块索引 | `save_blocks` `word_assembled`×2 `load_blocks` |
| `stress_mixed` | 以上矩阵按种子随机组合 3 个（确定性可复现） | `queue_put`（至少发出一条命令） |

> 注：两处「注入故障」（`x = 1/0`、`gen_code`/`ai_stream` 抛错）模拟的是用户真实
> 使用中会遇到的失败；被注入的故障点在 triage 里标为「预期故障路径」不算 bug，
> 它们验证的**自我修复 / 失败上报机制**才是被测对象。

## 三、日志检查规则

- 日志位置：`-debug` 模式下 `%APPDATA%\AI4Word\debug.log`（回退 TEMP），单文件
  5MB 滚动为 `.bak`；未捕获异常由 `sys.excepthook` / `threading.excepthook` 落盘
  （`uncaught_exception` / `uncaught_thread_exception` 一律视为真实 bug）。
- 每轮归档：`tests/deep_runs/<时间戳>/debug.log` 是当轮拷贝；`triage.txt` 是
  人类可读分类；`summary.json` 是机器可读结果（场景、FAIL、MUST 缺失、告警计数、
  `clean`、`exit_code`），是复测与 workflow 分析阶段的输入。
- 分类口径：
  - **真实 bug 候选**：场景时间窗内的 `WARN/ERROR`/traceback，且不属于注入故障
    或环境重试噪音（`api_*_retry` / `api_*_failed` / `word_connect_failed`）。
  - **环境波动**：网络 / API / Word 未运行等外部因素，可复现重试。
  - **平台假象**：真实平台下仍可能被环境拦截的项（如注册表自启被组策略拦截）。
  - **MUST_EVENTS 缺失**：场景流程走完但事件没打 = 接线缺失 = 真实 bug。
- 真实平台下没有了「offscreen 无 HWND」的借口：`hotkey_register`、拖动吸附、
    贴边、托盘的真实性都被提高一个量级——这正是本层相对离线实验台的价值。

## 四、修复与复测策略

- 修复全部直接落地（用户授权「全部直接修」），但**只修真实 bug**，不顺手重构。
- 每处修复配回归单测，并验证：单测在旧行为下失败、新行为下通过（能验证失败
  形态时）；至少现场跑通 `tests/run_offline.py`。
- 复测顺序：定向（`--only`）→ 全量。全量退出码 `0` 才算闭环。
- 无进展保护：连续两轮分析出同一份真实问题清单即停止，避免无限空转（对齐
  Claude Code 官方 workflow 示例「keep fixing until two rounds in a row make
  no progress」的精神）。

## 五、怎么跑

**一键（只跑测试层，人工分析）**

```powershell
.\run_deep_test.bat                      # 全量（单进程）
.\run_deep_test_parallel.bat             # 全量（并行槽，默认 3 槽，约 2.5-3× 提速）
.\run_deep_test_parallel.bat --slots 2   # 槽位数（内存吃紧用 2）
.\run_deep_test.bat --only interrupt     # 只跑名字含 interrupt 的场景
.\run_deep_test_parallel.bat --only interrupt_r
.\run_deep_test.bat --seed 31337 --keep-word
.\run_deep_test.bat --skip doc_persistence,stream_fault
```

跑完看 `tests\deep_runs\` 最新目录的 `triage.txt` / `summary.json`。
**前提**：`.venv` + `.env`（`ATRIA_API_KEY`）、已退出正在运行的 AI4Word、
测试期间不使用鼠标键盘、每槽各附加或启动一个临时空 Word 文档（不保存）。

**并行模式说明**

- 能并行是因为测试输入是 `QTest` 程序化事件（发给具体控件，不经操作系统
  焦点），多个 QApplication 进程可在同一桌面共存。
- 每槽通过环境变量彻底隔离：`AI4WORD_DEBUG_LOG`（否则多进程会互相覆盖
  `%APPDATA%\AI4Word\debug.log` 的共享日志与滚动）、`AI4WORD_SHM_KEY`
  （单实例锁）、`AI4WORD_TEST_SETTINGS`（临时 settings）、
  `AI4WORD_TEST_FORCE_NEW_WORD`（每槽 `Dispatch` 自己的 Word 实例，
  否则 `GetObject` 会让所有槽附加到同一个 Word 进程的同一个
  ActiveDocument，互相写进同一篇文档）；槽内查看类命令的重连也被
  `_attach_factory` 钉回本槽实例。
- 机器级全局状态固定只在槽 0 跑：`drag_snap_hotkey`（Ctrl+Alt+Space 的
  RegisterHotKey 是全局热键）与 `tray_actions`（开机自启写注册表）。
- 槽种子 `compact_send` + `review_presets_save` 每槽都先跑，保证
  `block_map`（要块模型）与 `doc_persistence`（要已落盘的存档变量）依赖
  在同槽内闭合；种子很轻，并行跑墙钟成本≈一份，代价是合并报告里这两个
  场景名会出现 K 次（每槽一份），属预期。
- 代价：每槽多占约 200-400MB 内存（一个 Word 实例）；同桌面偶发竞争制造
  的假告警会被 triage 归入「环境告警」兜底。
- 合并产物在 `tests\deep_runs\<时间戳>\`；各槽明细在其 `slotN\` 子目录
  （含自己的 debug.log / console.log / summary.json / triage.txt）。
  任一槽崩溃（未产出 summary）合并退出码为 2。
- 定向重跑（`--only`）命中少量场景时，编排器只启动需要的槽（单个场景
  自然只剩一个槽），不浪费。

**全自动闭环（Claude Code 工作流：测试 → 分析 → 修复 → 复测）**

```powershell
.\install_workflow.bat                   # 安装一次：.claude\workflows\deep-test.js
claude                                   # 在仓库目录启动 Claude Code
/deep-test                               # 开始闭环（最多 5 轮修复）
/deep-test rounds=3 only=interrupt       # 或限定范围与轮数
```

workflow 脚本（`claude-workflows/deep-test.js`）按 Claude Code 官方
[dynamic workflows](https://code.claude.com/docs/en/workflows) 形态编写：
脚本持有循环 / 分支 / 中间结果，子智能体只回结构化结果；`phases` 与
`phase()` 逐字一致；时间戳一律经 `args` 传入（脚本内禁 `Date.now()`）。

各阶段的并行策略：

- **测试 / 复测**：直接用 `.venv\python.exe -u tests\deep_parallel.py`（默认
  3 槽并行；Claude Code 的 bash 环境里 `cmd /c` 会被 MSYS 参数转义破坏，
  且本机 `.venv` 是 conda 布局、解释器在 `.venv\python.exe`，所以不走
  `.bat` 包装；人工仍可双击 `run_deep_test*.bat`，两个布局都支持）；
  复测先逐个 `--only` 定向重跑失败场景（单场景自然单槽），再全量。
- **分析**：两个子智能体并行——告警分诊员（triage/debug.log 分类）与
  覆盖审计员（MUST_EVENTS / FAIL 核对）。
- **修复**：先按证据去重（两个分析 agent 可能报同一问题），再按疑点文件
  分组：同一文件的多条发现交给同一个 agent 顺序修（避免多个 agent 并行
  编辑同一文件互相覆盖），不同文件并行修；每个 agent 必须配回归单测并
  现场跑通。

## 六、产物布局

```
tests/deep_runs/20261005-160000/          单进程 / 并行合并目录
├── debug.log        当轮完整调试日志（单进程）
├── summary.json     {clean, exit_code, totals, scenarios:[{name,ok,fails,
│                                     missing_must, unclassified, ...}]}
├── triage.txt       分类报告（FAIL / MUST 缺失 / 待确认 / 环境 / 平台假象）
└── slot0/  slot1/  slot2/               并行模式下的各槽明细
    ├── debug.log   该槽自己的日志（环境变量 AI4WORD_DEBUG_LOG 隔离）
    ├── console.log 该槽子进程控制台输出
    ├── summary.json / triage.txt
    └── settings.json  该槽的临时 settings（AI4WORD_TEST_SETTINGS）
```

## 七、退出码

| 退出码 | 含义 |
|--------|------|
| `0` | 全绿：无 FAIL、无 MUST_EVENTS 缺失、无待确认告警 |
| `1` | 有发现（FAIL / 缺失 / 待确认告警任一非零） |
| `2` | 环境错误（缺 `ATRIA_API_KEY`、单实例锁、harness 异常、无 `.venv` 等） |
