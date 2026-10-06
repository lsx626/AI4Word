// AI4Word 深度测试工作流（Claude Code dynamic workflow）
//
// 形态遵循 Claude Code 官方文档的 workflow 规范
// （https://code.claude.com/docs/en/workflows）：
//   - 安装到 .claude/workflows/ 后成为 /deep-test 命令
//     （安装方式：双击仓库根目录的 install_workflow.bat）
//   - export const meta 打头：name + description + phases（标题与 phase() 逐字一致）
//   - body 是支持 top-level await 的纯 JS；脚本持有循环、分支与中间结果，
//     子智能体（agent / pipeline / parallel）只回结果，不进入会话上下文
//   - 脚本内禁止 Date.now() / Math.random()（保证重跑一致性），时间戳用 args 传入
//
// 工作流（四阶段闭环，直到全绿或连续两轮无进展）：
//   测试  跑 run_deep_test_parallel.bat（tests/deep_parallel.py 把 21 个场景
//        分到 K 个进程槽并行跑，每槽独立 Word / 日志 / 单实例锁；测试层
//        为 tests/deep_test.py：真实窗口 / 真实点击 / 真实 Word COM /
//        真实 Atria API），产出 tests/deep_runs/<ts>/ 的 debug.log +
//        summary.json + triage.txt（并行轮为合并目录），退出码 0/1/2
//   分析  并行两个子智能体：一个分类 triage/debug.log 里的告警
//        （真实 bug / 环境波动 / 平台假象 / 预期故障路径），一个校验
//        MUST_EVENTS 覆盖与 FAIL 列表；合并出「真实问题清单」
//   修复  按证据去重后按疑点文件分组：同文件多条发现给一个 agent 顺序修
//        （避免并行编辑同一文件），不同文件并行修；各自配回归单测并
//        现场验证
//   复测  先 --only 定向重跑失败场景，再全量重跑；未全绿则回到「分析」
//
// 用法（在仓库 D:\Projects\AI4Word 内的 Claude Code 会话）：
//   /deep-test                      全量跑，最多 5 轮修复
//   /deep-test rounds=3             最多 3 轮
//   /deep-test only=interrupt       首轮只跑名字含 interrupt 的场景
//   /deep-test ts=20261005-1600     轮次标记
//
// 返回：JSON 报告 {ok, rounds, final_exit_code, fixed_count, run_dir, ...}。

export const meta = {
  name: 'deep-test',
  description: 'AI4Word 全真实深度测试闭环：真实驱动测试 -> 分析日志 -> 修复 -> 复测，直至全绿',
  phases: ['测试', '分析', '修复', '复测'],
}

const REPO = 'D:\\Projects\\AI4Word'

// ---- args：可传 ts / rounds / only ----
const rawArgs = (typeof args !== 'undefined' && args) || {}
const opts = (typeof rawArgs === 'string')
  ? Object.fromEntries(String(rawArgs).split(/\s+/).filter(Boolean).map(tok => {
      const i = tok.indexOf('=')
      return i > 0 ? [tok.slice(0, i), tok.slice(i + 1)] : [tok, '1']
    }))
  : rawArgs
const maxRounds = Number(opts.rounds) || 5
const onlyFilter = opts.only || ''
const runTag = opts.ts || ''

// ---- 结构化输出契约 ----
const runSchema = {
  type: 'object',
  required: ['exit_code', 'ran'],
  properties: {
    exit_code: { type: 'integer' },        // 0 全绿 / 1 有发现 / 2 环境错误
    ran: { type: 'boolean' },              // 是否真正跑完了场景
    run_dir: { type: 'string' },           // tests/deep_runs/<最新目录> 绝对路径
    failed_scenarios: { type: 'array', items: { type: 'string' } },
    totals: { type: 'object' },            // summary.json 的 totals
    note: { type: 'string' },
  },
}

const classifierSchema = {
  type: 'object',
  required: ['findings'],
  properties: {
    findings: {
      type: 'array',
      items: {
        type: 'object',
        required: ['classification', 'evidence'],
        properties: {
          // real_bug | environment | platform_artifact | expected_fault
          classification: { type: 'string' },
          scenario: { type: 'string' },
          event: { type: 'string' },
          evidence: { type: 'string' },
          suspect_file: { type: 'string' },
          fix_hint: { type: 'string' },
        },
      },
    },
  },
}

const coverageSchema = {
  type: 'object',
  required: ['real_candidates'],
  properties: {
    missing_must: { type: 'array', items: { type: 'string' } },
    failed_checks: { type: 'array', items: { type: 'string' } },
    real_candidates: { type: 'array', items: { type: 'string' } },
    note: { type: 'string' },
  },
}

// ---- 阶段 1：测试 ----
async function runTest(extraArgs, label) {
  // extraArgs：透传给 tests/deep_test.py（--only / --seed / --skip）
  return await agent(
    [
      '你要在 Windows 上执行 AI4Word 的全真实驱动深度测试并如实汇报结果。',
      '仓库目录：' + REPO + '（你的工作目录应该就在这里）。',
      '执行方式（在 bash 里直接用 .venv 的 python 跑，不要套 cmd /c：',
      '本机 MSYS 会把 /c 参数转义掉、cmd 默认也不搜索当前目录，',
      'cmd /c run_deep_test_parallel.bat 在此环境根本跑不起来）：',
      '  .venv/python.exe -u tests/deep_parallel.py ' + extraArgs,
      '  或（并行不灵时退回单进程）.venv/python.exe -u tests/deep_test.py ' + extraArgs,
      '',
      '测试期间多个悬浮窗会被真实移动、真实点击：如果弹窗/对话框挡路，自行关掉；',
      '并行约 10-25 分钟（单进程 10-40 分钟），耐心等待全部子进程结束，',
      '不要中途杀进程（杀掉一个槽会让合并器把该槽记为崩溃，退出码 2）。',
      '若起不来（例如缺 ATRIA_API_KEY、单实例锁、缺少 .venv），直接按',
      '退出码 2 汇报原因，不要尝试修复任何东西。',
      '',
      '结束后：',
      '1) 找到 tests\\deep_runs\\ 下最新的目录（修改时间最新那个）。',
      '2) 读取该目录下的 summary.json。',
      '3) 按 schema 汇报：exit_code（测试进程的退出码）、ran=是否真正',
      '   跑完场景、run_dir=该目录绝对路径、failed_scenarios=ok=false 的',
      '   场景名列表、totals=summary.json 里的 totals、note=补充说明。',
      '只汇报，不修改任何文件。',
    ].join('\n'),
    { schema: runSchema, label },
  )
}

// ---- 阶段 2：分析 ----
function classifyPrompt(runDir) {
  return [
    '你是 AI4Word 项目的日志分诊员。仓库：' + REPO,
    '本轮深度测试产物在：' + runDir,
    '读取 triage.txt 与 debug.log（同目录），对每一个「待确认」告警、',
    'WARN / ERROR / traceback 行分类，分类只能取这四种之一：',
    '  real_bug           代码缺陷：崩溃、异常堆栈、状态不一致、',
    '                     日志接线缺失、断言 FAIL 对应的真实机制损坏',
    '  environment        网络波动 / API 限流 / Word 未运行等外部因素',
    '  platform_artifact  真实平台下窗口/热键/注册表被环境拦截',
    '  expected_fault     场景注入的故障路径（stmt_failed / gen_code_failed /',
    '                     stream_error / settings_save_failed / autostart_*_failed）',
    '',
    '判别依据：同一告警在 debug.log 里的上下文（前后事件、是否出现在注入',
    '故障的场景窗口内）、summary.json 里该场景的 FAIL 描述。',
    '只有 real_bug 才会进入修复阶段，environment / platform_artifact /',
    'expected_fault 必须给出理由。',
    '',
    '真实 bug 的判别要点（项目历史教训）：',
    '  - 同一 FAIL 在旧轮次修复后重新出现 = 修复不彻底，仍是 real_bug',
    '  - 日志缺事件（MUST_EVENTS 缺失）而流程跑完 = 接线缺失，是 real_bug',
    '  - 离线 fake Word 覆盖不到、只在真实 Word 上爆的问题 = 仍算 real_bug',
    '  - 主线程碰 COM（apartment 违规）= real_bug',
    '按 schema 返回 findings（每条含 classification / scenario / event /',
    'evidence / suspect_file / fix_hint）。',
  ].join('\n')
}

function coveragePrompt(runDir) {
  return [
    '你是 AI4Word 项目的覆盖审计员。仓库：' + REPO,
    '本轮深度测试产物在：' + runDir,
    '读取 summary.json 与 triage.txt，核对：',
    '  1) missing_must 列出的 MUST_EVENTS 缺失是否指向真实机制损坏',
    '     （对照 tests\\deep_test.py 里每个场景的 MUST 字典：事件名与数量）',
    '  2) 每个 FAIL 是否与场景的机制描述吻合、能否由环境解释',
    '  3) harness 异常（errors 非空）一律算真实问题',
    '按 schema 返回：missing_must / failed_checks / real_candidates',
    '（字符串数组，每条「场景: 事件/检查: 证据」），real_candidates 只放',
    '你判断为真实代码问题的条目。',
    '只分析，不修改任何文件。',
  ].join('\n')
}

// ---- 阶段 3：修复（按文件分组） ----
function fixGroupPrompt(file, findings) {
  const list = findings.map((f, i) =>
    '  [' + (i + 1) + '] 场景：' + (f.scenario || '(未知)') +
    '  事件/检查：' + (f.event || '(未知)') +
    '  证据：' + f.evidence +
    (f.fix_hint ? '  修复提示：' + f.fix_hint : '')
  ).join('\n')
  return [
    '你是 AI4Word 的修复工程师，负责一个文件分组内的全部问题。',
    '仓库：' + REPO,
    '你独占的疑点文件：' + file + '（其它 agent 在并行修其它文件，',
    '请不要越界改不属于本组问题直接需要的代码）',
    '本组的真实问题清单：',
    list,
    '',
    '要求：',
    '  1) 组内问题按顺序逐个修复；先定位真实原因（读相关代码与调用链），',
    '     不要只消灭症状。',
    '  2) 项目铁律（见 DEVELOPMENT_LOG.md）：内容编辑必须走块原语',
    '     （replace_block / insert_after / delete_block），格式修改改样式',
    '     定义（doc.Styles）而不是遍历段落；所有 Word COM 调用只能在',
    '     worker 线程；编辑原语执行后 rebuild_ranges；fake Word 必须实现',
    '     真实 COM 语义（动态 Range 平移、段落模型）。',
    '  3) 每处修复配一个永久回归单测（现有 tests\\test_*.py 之一或新建，',
    '     优先离线可跑），并现场跑通（注意：本机 .venv 无 pytest，',
    '     离线单测要用 anaconda 基础解释器，见 EXPERIMENT_LOG.md；',
    '     且 tests\\test_*.py 会 import 仓库根模块，必须 PYTHONPATH=.）：',
    '       PYTHONPATH=. python -u tests\\run_offline.py',
    '     或对单个文件：PYTHONPATH=. python tests\\test_xxx.py',
    '     （python 不在 PATH 时用 D:\\ProgramData\\anaconda3\\python.exe；',
    '     .venv\\python.exe 只留给 deep_test / deep_parallel 真实驱动用；',
    '     若单测有入口）pytest 可用时也可 PYTHONPATH=. python -m pytest。',
    '  4) 不要修改 .env / tests\\deep_runs\\ 产物；不要动与问题无关的代码。',
    '  5) 用户可见文案保持中文。',
    '  6) 若某个问题的修复必须碰本组之外文件，只改最少必要性，并在总结里',
    '     明确列出碰了哪些外部文件（供下一轮 review）。',
    '完成后用一段话总结：改了哪个文件、加了什么单测、单测结果、',
    '是否动过外部文件。',
  ].join('\n')
}

// ---- 阶段 4：复测 ----
function retestPrompt(failedScenarios, runDir) {
  const targeted = (failedScenarios && failedScenarios.length > 0)
  const lines = [
    '你要在 Windows 上重跑 AI4Word 的全真实驱动深度测试。仓库：' + REPO,
  ]
  if (targeted) {
    lines.push(
      '上一轮失败的场景：' + failedScenarios.join('、'),
      '请先逐个定向重跑（tests\\deep_test.py 的 --only 是子串过滤，一次一个子串，',
      '公共前缀要够精确，例如 interrupt_rollback / interrupt_keep / interrupt_extra',
      '分别用 --only interrupt_r / --only interrupt_k / --only interrupt_e；',
      'arrange_ok / arrange_repair / arrange_gen_fail 同理）。',
      '命令：.venv/python.exe -u tests/deep_parallel.py --only <子串>',
      '（每个 --only 只命中少量场景时并行器会自动只启动需要的槽，不必担心浪费）',
      '全部定向轮都通过（退出码 0）后，再跑一次全量：',
      '  .venv/python.exe -u tests/deep_parallel.py',
    )
  } else {
    lines.push('直接跑全量：.venv/python.exe -u tests/deep_parallel.py')
  }
  lines.push(
    '',
    '可以跑很久（定向轮几分钟，全量并行约 10-25 分钟），耐心等全部子进程',
    '结束，不要杀进程。',
    '结束后找到 tests\\deep_runs\\ 最新目录（并行轮是合并目录，里面直接',
    '放着 summary.json 与 triage.txt），读取 summary.json，按 schema 汇报：',
    'exit_code（最后一次运行的退出码）、ran、run_dir、failed_scenarios、',
    'totals、note。',
    '只汇报，不修改任何文件（修复在上一轮已经做过了）。',
  )
  return lines.join('\n')
}

// ---- 主循环 ----
let lastExit = null
let lastRunDir = null
let lastFailed = null
let prevSignature = null
let noProgressStreak = 0
let round = 0
let fixedCount = 0
let report = { ok: false, rounds: 0, final_exit_code: null, note: '未运行' }

await phase('测试')
let run = await runTest(onlyFilter ? '--only ' + onlyFilter : '',
  '初轮' + (runTag ? ' ' + runTag : ''))
if (!run) {
  log('测试子智能体异常返回（被停止或 API 不可恢复），工作流终止')
  report.note = '测试子智能体返回 null'
  report.rounds = 0
  return report
}
lastExit = run.exit_code
lastRunDir = run.run_dir
lastFailed = run.failed_scenarios || []
report.final_exit_code = lastExit
report.run_dir = lastRunDir

while (lastExit !== 0 && round < maxRounds) {
  round += 1
  log('第 ' + round + ' 轮：分析上一轮产物 ' + (lastRunDir || '(无目录)'))

  await phase('分析')
  let classified = null
  let coverage = null
  try {
    ;[classified, coverage] = await parallel([
      () => agent(classifyPrompt(lastRunDir), { schema: classifierSchema, label: '告警分诊' }),
      () => agent(coveragePrompt(lastRunDir), { schema: coverageSchema, label: '覆盖审计' }),
    ])
  } catch (e) {
    log('分析阶段失败：' + String(e))
    break
  }
  if (!classified || !classified.findings) {
    log('分诊未返回结果，终止循环')
    break
  }

  // 合并「真实问题」清单（分类为 real_bug 的 + 覆盖审计的 real_candidates）
  const findings = classified.findings
    .filter(f => f && f.classification === 'real_bug')
    .map(f => ({
      scenario: f.scenario || '',
      event: f.event || '',
      evidence: f.evidence || '',
      suspect_file: f.suspect_file || '',
      fix_hint: f.fix_hint || '',
    }))
  const extra = (coverage && coverage.real_candidates) || []
  for (const text of extra) {
    findings.push({ scenario: '', event: '', evidence: text,
                    suspect_file: '', fix_hint: '' })
  }

  // 连续两轮同一份真实问题清单 = 无进展
  const signature = findings.map(f => f.evidence).sort().join('||')
  if (signature === prevSignature) {
    noProgressStreak += 1
  } else {
    noProgressStreak = 0
    prevSignature = signature
  }
  if (findings.length === 0 || noProgressStreak >= 2) {
    log(findings.length === 0
      ? '没有发现真实问题（告警均为环境/平台/预期故障路径），闭环结束'
      : '连续两轮修复没有进展，停止循环，把问题留给人工')
    report.ok = false
    report.note = findings.length === 0
      ? '无真实问题：剩余告警为环境波动/平台假象/注入故障路径'
      : '连续两轮无进展，剩余 ' + findings.length + ' 个真实问题未解决'
    report.rounds = round
    report.final_exit_code = lastExit
    report.run_dir = lastRunDir
    report.remaining_findings = findings
    return report
  }

  // 去重（两个分析 agent 可能报同一问题），再按疑点文件分组：
  // 同一个文件的多条发现交给同一个 agent 顺序修，避免多个 agent 并行
  // 编辑同一文件互相覆盖（这是并行修复里最常见的正确性问题）。
  const seenEv = new Set()
  const uniq = []
  for (const f of findings) {
    const key = (f.evidence || '').trim()
    if (!key || seenEv.has(key)) continue
    seenEv.add(key)
    uniq.push(f)
  }
  const fileOf = (f) => (f.suspect_file && f.suspect_file.trim())
    ? f.suspect_file.trim() : '(未定位文件)'
  const groups = {}
  for (const f of uniq) {
    const k = fileOf(f)
    ;(groups[k] = groups[k] || []).push(f)
  }
  const groupList = Object.entries(groups).map(([file, fs]) => ({ file, fs }))
  log('发现 ' + uniq.length + ' 个真实问题，归并到 ' + groupList.length +
      ' 个文件分组并行修复（同文件内顺序修）')
  await phase('修复')
  await pipeline(groupList, g =>
    agent(fixGroupPrompt(g.file, g.fs),
      { label: g.file.replace(/[^0-9A-Za-z_\-.\\\/]/g, '').slice(-40) }))
  fixedCount += uniq.length
  log('修复完成，进入复测')

  await phase('复测')
  let rerun = null
  try {
    rerun = await agent(retestPrompt(lastFailed, lastRunDir),
      { schema: runSchema, label: '第 ' + round + ' 轮复测' })
  } catch (e) {
    log('复测阶段失败：' + String(e))
    break
  }
  if (!rerun) {
    log('复测子智能体异常返回，工作流终止')
    break
  }
  lastExit = rerun.exit_code
  lastRunDir = rerun.run_dir || lastRunDir
  lastFailed = rerun.failed_scenarios || []
  report.final_exit_code = lastExit
  report.run_dir = lastRunDir
  report.rounds = round
  if (lastExit === 0) {
    log('第 ' + round + ' 轮复测全绿（exit 0），闭环完成')
  }
}

report.ok = lastExit === 0
report.rounds = round || 0
report.fixed_count = fixedCount
report.final_exit_code = lastExit
report.run_dir = lastRunDir
return report
