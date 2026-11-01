/**
 * dsh-docx-format — 让 DSH 生成的所有 .docx 自动套用《通用公文格式规范》。
 *
 * 三层能力：
 *  1) docx_format 工具：模型可显式调用，apply（规范化）/ check（只体检）。
 *  2) systemPrompt 段：把格式硬规则写进提示词，模型一开始就知道该怎么排版。
 *  3) tools/result 观察者：任何工具产出 .docx 路径后，自动排队规范化（默认只处理本次会话新写的文件）。
 *
 * 请求门控（两段式）：
 *  - 触发：当前会话最后一条用户消息必须命中 triggerKeywords（默认「公文」），否则整个插件不介入；
 *    所以普通对话里生成的 .docx 不会被套上公文版式。
 *  - 排除：命中 skipRequestKeywords（默认「实验报告」）时同样不介入。
 * 不介入 = 不排队、不改文件、工具调用直接返回 skipped（显式调用传 force=true 可绕过门控）。
 *
 * 真实排版逻辑在 lib/docxfmt.py（python-docx），本文件只做进程调用与调度。
 */
import { existsSync, statSync } from 'node:fs'
import { spawn } from 'node:child_process'
import { isAbsolute, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import path from 'node:path'

export const name = 'dsh-docx-format'

const HERE = path.dirname(fileURLToPath(import.meta.url))
const ENGINE = path.join(HERE, 'lib', 'docxfmt.py')
const READER = path.join(HERE, 'lib', 'docx_read.py')
const LOADED_AT = Date.now()

const DEFAULTS = {
  autoFormat: true,
  formatExisting: false,
  debounceMs: 1200,
  timeoutMs: 120000,
  python: '',
  maxFilesPerTurn: 8,
  skipPatterns: ['*格式要求*', '*模板*', '~$*', '*/node_modules/*'],
  // 触发词：用户请求文字里必须出现其中之一，插件才介入（空数组 = 任何请求都触发）
  triggerKeywords: ['公文'],
  // 排除词：命中即整个插件不运行（如「实验报告」这类不该套公文版式的请求）
  skipRequestKeywords: ['实验报告'],
  prompt: { enabled: true, order: 3260 },
  tool: { enabled: true },
  defaults: {},
}

/* ------------------------------------------------------------------ 配置 */

function normalizeConfig(config) {
  const cfg = { ...DEFAULTS, ...(config ?? {}) }
  cfg.skipPatterns = Array.isArray(cfg.skipPatterns) ? cfg.skipPatterns.map(String) : DEFAULTS.skipPatterns
  const toKeywords = (value, fallback) => (Array.isArray(value)
    ? value.map(String).map((s) => s.trim()).filter(Boolean)
    : fallback)
  cfg.triggerKeywords = toKeywords(cfg.triggerKeywords, DEFAULTS.triggerKeywords)
  cfg.skipRequestKeywords = toKeywords(cfg.skipRequestKeywords, DEFAULTS.skipRequestKeywords)
  cfg.prompt = { ...DEFAULTS.prompt, ...(config?.prompt ?? {}) }
  cfg.tool = { ...DEFAULTS.tool, ...(config?.tool ?? {}) }
  cfg.defaults = config?.defaults && typeof config.defaults === 'object' ? config.defaults : {}
  const n = Number(cfg.debounceMs)
  cfg.debounceMs = Number.isFinite(n) && n >= 0 ? n : DEFAULTS.debounceMs
  const t = Number(cfg.timeoutMs)
  cfg.timeoutMs = Number.isFinite(t) && t > 0 ? t : DEFAULTS.timeoutMs
  const m = Number(cfg.maxFilesPerTurn)
  cfg.maxFilesPerTurn = Number.isFinite(m) && m > 0 ? Math.floor(m) : DEFAULTS.maxFilesPerTurn
  return cfg
}

/* --------------------------------------------------------------- 解释器 */

function pythonCandidates(cfg) {
  const list = []
  if (cfg.python) list.push(cfg.python)
  if (process.env.DSH_DOCX_PYTHON) list.push(process.env.DSH_DOCX_PYTHON)
  const home = process.env.USERPROFILE || process.env.HOME || ''
  if (home) {
    list.push(path.join(home, '.dsh', 'dsh-runtimes', 'dsh-primary-runtime', 'dependencies', 'python', 'python.exe'))
  }
  list.push('python', 'python3')
  return list
}

function resolvePython(cfg) {
  for (const c of pythonCandidates(cfg)) {
    if (c === 'python' || c === 'python3') return c
    if (existsSync(c)) return c
  }
  return 'python'
}

/* ----------------------------------------------------------- 调用引擎 */

function runScript(python, script, args, timeoutMs) {
  return new Promise((done) => {
    const env = { ...process.env, PYTHONIOENCODING: 'utf-8', PYTHONUTF8: '1' }
    const child = spawn(python, [script, ...args], { env, windowsHide: true })
    let out = ''
    let err = ''
    let settled = false
    const finish = (value) => {
      if (settled) return
      settled = true
      done(value)
    }
    const timer = setTimeout(() => {
      try { child.kill() } catch { /* ignore */ }
      finish({ ok: false, code: 'timeout', message: 'docxfmt 超时：' + timeoutMs + 'ms' })
    }, timeoutMs)
    child.stdout.setEncoding('utf8')
    child.stderr.setEncoding('utf8')
    child.stdout.on('data', (chunk) => { out += chunk })
    child.stderr.on('data', (chunk) => { err += chunk })
    child.on('error', (e) => {
      clearTimeout(timer)
      finish({ ok: false, code: 'spawn', message: '无法启动 python：' + (e?.message ?? String(e)) })
    })
    child.on('close', (code) => {
      clearTimeout(timer)
      finish({ ok: true, code, stdout: out, stderr: err })
    })
  })
}

function runEngine(python, args, timeoutMs) {
  return runScript(python, ENGINE, args, timeoutMs)
}

/**
 * DSH 要求工具返回值「可无损 JSON 化」：对象里带值为 undefined 的键会让
 * JSON.stringify → parse 后丢键，校验直接报 `value is not lossless JSON`。
 * docx_format 旧版把 `skipReason: skipped[0]?.skipReason` 直接放进去，
 * 非跳过场景 key 存在但值是 undefined —— 于是每次正常调用都被判 invalid。
 */
function stripUndefined(value) {
  if (Array.isArray(value)) return value.map(stripUndefined)
  if (value && typeof value === 'object') {
    const out = {}
    for (const [key, item] of Object.entries(value)) {
      if (item !== undefined) out[key] = stripUndefined(item)
    }
    return out
  }
  return value
}

function parseReport(stdout) {
  const text = String(stdout ?? '').trim()
  if (text === '') return undefined
  const start = text.indexOf('{')
  if (start < 0) return undefined
  try {
    return JSON.parse(text.slice(start))
  } catch {
    return undefined
  }
}

async function formatOne(cfg, python, target, mode, recursive) {
  const args = [target, '--mode', mode, '--json']
  if (recursive) args.push('--recursive')
  if (cfg.defaults && Object.keys(cfg.defaults).length > 0) args.push('--config', JSON.stringify(cfg.defaults))
  const run = await runEngine(python, args, cfg.timeoutMs)
  if (!run.ok) return { ok: false, path: target, mode, message: run.message, changed: 0, violations: 0 }
  const report = parseReport(run.stdout)
  if (!report) {
    const detail = (run.stderr || run.stdout || '').trim().split('\n').slice(-4).join(' | ')
    return { ok: false, path: target, mode, message: 'docxfmt 输出无法解析：' + detail, changed: 0, violations: 0 }
  }
  const files = Array.isArray(report.files) ? report.files : [report]
  const fixes = files.reduce((n, f) => n + (Number(f.violations) || 0), 0)
  const changedFiles = files.filter((f) => f.changed === true).length
  const skipped = files.filter((f) => f.skipped === true)
  return stripUndefined({
    ok: files.every((f) => f.ok !== false),
    path: report.path ?? target,
    mode,
    changed: mode === 'apply' ? fixes : 0,
    violations: mode === 'check' ? fixes : 0,
    fixes,
    changedFiles,
    skipped: skipped.length,
    skipReason: skipped[0]?.skipReason,
    files,
    warnings: files.flatMap((f) => (Array.isArray(f.warnings) ? f.warnings : [])),
  })
}

/* --------------------------------------------------------- 路径采集 */

const BT = String.fromCharCode(96)
const DOCX_RE = new RegExp(
  '[A-Za-z]:[\\\\/][^\\n\\"' + BT + '<>|，。；：、（）()\\[\\]{}]*?\\.docx'
  + '|[^\\s\\"' + BT + '<>|，。；：、（）()\\[\\]{}]+\\.docx',
  'gi',
)

function wildcardToRegExp(pattern) {
  const escaped = String(pattern).replace(/[{}.*+?^$()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.')
  return new RegExp('^' + escaped + '$', 'i')
}

function isSkipped(file, patterns) {
  const normalized = file.replace(/\\/g, '/')
  return patterns.some((p) => {
    try {
      return wildcardToRegExp(p.replace(/\\/g, '/')).test(normalized)
    } catch {
      return false
    }
  })
}

function collectStrings(value, out, depth = 0) {
  if (depth > 3 || out.length > 200) return
  if (typeof value === 'string') { out.push(value); return }
  if (Array.isArray(value)) { for (const item of value) collectStrings(item, out, depth + 1); return }
  if (value && typeof value === 'object') {
    for (const item of Object.values(value)) collectStrings(item, out, depth + 1)
  }
}

function extractDocxPaths(exec, result) {
  const strings = []
  collectStrings(exec?.arguments, strings)
  collectStrings(result?.content, strings)
  const found = new Set()
  for (const text of strings) {
    const matches = String(text).match(DOCX_RE)
    if (!matches) continue
    for (const raw of matches) {
      const candidate = raw.replace(/^[（(“"'【\[]+/, '').trim()
      if (!candidate || candidate.startsWith('http')) continue
      found.add(candidate)
    }
  }
  return [...found]
}

/* -------------------------------------------------- 请求文字白名单 */

/** 从一个 Message 里拼出纯文本（只取 text block）。 */
function messageText(message) {
  const parts = []
  const content = message?.content
  if (Array.isArray(content)) {
    for (const block of content) {
      if (block && block.type === 'text' && typeof block.text === 'string') parts.push(block.text)
    }
  }
  if (typeof message?.text === 'string') parts.push(message.text)
  return parts.join('\n').trim()
}

/** 最后一条 user 消息的文本；失败一律返回 ''。 */
function lastUserText(session) {
  try {
    const messages = session.deriveMessages?.() ?? []
    for (let i = messages.length - 1; i >= 0; i -= 1) {
      const m = messages[i]
      if (m?.role !== 'user') continue
      const text = messageText(m)
      if (text !== '') return text
    }
  } catch { /* ignore */ }
  return ''
}

function matchKeyword(text, keywords) {
  if (!text || !Array.isArray(keywords)) return undefined
  for (const kw of keywords) {
    if (kw && text.includes(kw)) return kw
  }
  return undefined
}

/* ------------------------------------------------------------ 插件主体 */

export function apply(ctx, config) {
  const cfg = normalizeConfig(config)
  const python = resolvePython(cfg)
  const queue = []
  const seen = new Map()
  const formattedByUs = new Set()
  /** sessionId -> { seq, text } 最近用户请求文字缓存（避免反复 deriveMessages） */
  const requestCache = new Map()
  let timer
  let running = false

  const log = (level, ...args) => {
    try { ctx.logger?.[level]?.(...args) } catch { /* ignore */ }
  }

  /* ---- 会话 / 请求文字 ---- */

  const sessionsService = () => {
    try {
      const svc = ctx.get('sessions')
      return svc && typeof svc.get === 'function' ? svc : undefined
    } catch {
      return undefined
    }
  }

  const sessionById = (sessionId) => {
    if (!sessionId) return undefined
    const svc = sessionsService()
    if (svc) {
      try {
        const found = svc.get(sessionId)
        if (found) return found
      } catch { /* ignore */ }
    }
    return undefined
  }

  /** 缓存失效判断：同一会话 seq 未变则复用上次预算的文本。 */
  const requestTextOf = (sessionId) => {
    if (!sessionId) return ''
    const session = sessionById(sessionId)
    if (session) {
      let seq
      try { seq = Number(session.seq ?? -1) } catch { seq = -1 }
      const cached = requestCache.get(sessionId)
      if (cached && cached.seq === seq) return cached.text
      const text = lastUserText(session)
      requestCache.set(sessionId, { seq, text })
      return text
    }
    return requestCache.get(sessionId)?.text ?? ''
  }

  /**
   * 请求门控：决定这条请求要不要套公文格式。
   *  - 先看排除词（skipRequestKeywords，默认「实验报告」）→ 命中则停用；
   *  - 再看触发词（triggerKeywords，默认「公文」）→ 非空且未命中则停用；
   *  - 两者都过 → 返回 { run: true }。
   */
  const requestGate = (sessionId) => {
    const text = requestTextOf(sessionId)
    const skip = matchKeyword(text, cfg.skipRequestKeywords)
    if (skip) return { run: false, kind: 'skip', keyword: skip, text }
    if (cfg.triggerKeywords.length > 0) {
      const trigger = matchKeyword(text, cfg.triggerKeywords)
      if (!trigger) return { run: false, kind: 'no-trigger', text }
    }
    return { run: true, text }
  }

  /* ---- 文件排队 ---- */

  const resolveCandidate = (candidate) => {
    if (isAbsolute(candidate)) return candidate
    const roots = []
    try { if (ctx.baseDir) roots.push(ctx.baseDir) } catch { /* ignore */ }
    roots.push(process.cwd())
    for (const root of roots) {
      const joined = resolve(root, candidate)
      if (existsSync(joined)) return joined
    }
    return resolve(process.cwd(), candidate)
  }

  const enqueue = (candidate, sessionId) => {
    const file = resolveCandidate(candidate)
    if (!file.toLowerCase().endsWith('.docx')) return
    if (!existsSync(file)) return
    if (isSkipped(file, cfg.skipPatterns)) return
    if (formattedByUs.has(file)) return
    let stat
    try { stat = statSync(file) } catch { return }
    if (!cfg.formatExisting && stat.mtimeMs < LOADED_AT) return
    const key = file.toLowerCase()
    if (seen.get(key) === stat.mtimeMs) return
    seen.set(key, stat.mtimeMs)
    if (queue.some((item) => item.file.toLowerCase() === key)) return
    queue.push({ file, sessionId })
    log('info', '[docx-format] queued ' + file)
  }

  const schedule = () => {
    if (timer) clearTimeout(timer)
    timer = setTimeout(() => {
      timer = undefined
      if (running) return
      void drain()
    }, cfg.debounceMs)
    try { timer.unref?.() } catch { /* ignore */ }
  }

  const drain = async () => {
    running = true
    try {
      let handled = 0
      while (queue.length > 0 && handled < cfg.maxFilesPerTurn) {
        const job = queue.shift()
        // 出队时再判一次门控：用户可能在这期间改了要求
        const gate = requestGate(job.sessionId)
        if (!gate.run) {
          log('info', '[docx-format] skip ' + job.file + ' — ' + (gate.kind === 'skip'
            ? ('请求文字含排除词「' + gate.keyword + '」')
            : '请求文字未提触发词「' + cfg.triggerKeywords.join('|') + '」'))
          continue
        }
        formattedByUs.add(job.file)
        const report = await formatOne(cfg, python, job.file, 'apply', false)
        handled += 1
        if (report.ok) {
          log('info', '[docx-format] ' + (report.skipped > 0 ? 'skipped ' : 'formatted ') + report.path
            + ' (' + report.changed + ' fixes, ' + report.violations + ' violations)')
        } else {
          log('warn', '[docx-format] failed ' + job.file + ': ' + (report.message ?? 'unknown'))
        }
      }
    } finally {
      running = false
      if (queue.length > 0) schedule()
    }
  }

  /* ------------------------------------------------- docx_format 工具 */
  const outputSchema = {
    type: 'object',
    properties: {
      ok: { type: 'boolean' },
      path: { type: 'string' },
      mode: { type: 'string' },
      changed: { type: 'integer' },
      violations: { type: 'integer' },
      fixes: { type: 'integer' },
      changedFiles: { type: 'integer' },
      skipped: { type: 'integer' },
      skipReason: { type: 'string' },
      warnings: { type: 'array', items: { type: 'string' } },
      message: { type: 'string' },
      files: { type: 'array', items: { type: 'object', additionalProperties: true } },
    },
    required: ['ok', 'path', 'mode', 'changed', 'violations'],
    additionalProperties: true,
  }

  const renderReport = (value) => {
    const files = Array.isArray(value?.files) ? value.files : []
    const lines = []
    if (value?.message) lines.push(value.message)
    if (value?.skipped > 0 && files.length <= 1) {
      lines.push('[SKIP] ' + value.path + ' — 白名单命中，未处理（' + (value.skipReason ?? '') + '）')
      return [{ type: 'text', text: lines.join('\n') }]
    }
    if (files.length <= 1) {
      const head = value.ok === false
        ? ('执行失败：' + (value.message ?? '未知错误'))
        : value.mode === 'check'
          ? (value.violations > 0 ? '格式偏差 ' + value.violations + ' 处' : '格式完全符合要求')
          : (value.changed > 0 ? '已修正 ' + value.changed + ' 处' : '无需修改')
      lines.push('[' + (value.ok ? 'OK' : 'FAIL') + '] ' + value.path + ' — ' + head)
      const only = files[0]
      if (only?.rules) {
        for (const [rule, info] of Object.entries(only.rules)) {
          const sample = info?.samples?.[0]
          lines.push('  - ' + rule + ' x' + info.count
            + (sample ? (': ' + sample.where + ' ' + (sample.current ?? '') + ' -> ' + (sample.expected ?? '')) : ''))
        }
      }
    } else {
      lines.push('[' + (value.ok ? 'OK' : 'FAIL') + '] ' + files.length + ' 个文件，修正 '
        + value.changed + ' 处，偏差 ' + value.violations + ' 处'
        + (value.skipped ? ('，跳过 ' + value.skipped + ' 个') : ''))
      for (const f of files) {
        lines.push('  - ' + (f.skipped ? 'SKIP' : (f.ok ? 'OK' : 'FAIL')) + ' ' + f.path
          + (f.skipped ? ' (' + (f.skipReason ?? '') + ')' : ' (fixes ' + f.violations + ')'))
      }
    }
    return [{ type: 'text', text: lines.join('\n') }]
  }

  const toolDefinition = {
    name: 'docx_format',
    description: [
      '把 .docx 规范化为《通用公文格式规范》：方正小标宋简体二号标题、仿宋三号正文、',
      '固定行距 29 磅、首行缩进 2 字符（w:firstLineChars）、段前段后 0 磅、页边距 3.6/3.0/2.7/2.7cm、',
      '表格单元格水平垂直居中与“— 1 —”页码。',
      'mode=apply 直接改文件（幂等，可重复调用）；mode=check 只报告偏差不改文件。',
      '仅在用户对话里提到「公文」（triggerKeywords）时插件才介入，未提到时直接返回 skipped；',
      '请求里出现「实验报告」等排除词时同样停用。确需强制格式化时传 force=true。',
    ].join(''),
    parameters: {
      type: 'object',
      properties: {
        path: { type: 'string', description: '要处理的 .docx 文件或目录（绝对路径最稳）。' },
        mode: { type: 'string', enum: ['apply', 'check'], description: 'apply=规范化并写回；check=只体检。默认 apply。' },
        recursive: { type: 'boolean', description: 'path 是目录时是否递归子目录。默认 false。' },
        force: { type: 'boolean', description: '忽略请求触发/排除词强制处理。默认 false。' },
      },
      required: ['path'],
      additionalProperties: false,
    },
    output: { schema: outputSchema, render: (_args, value) => renderReport(value) },
    async execute(args, exec) {
      const known = new Set(['path', 'mode', 'recursive', 'force'])
      for (const key of Object.keys(args ?? {})) {
        if (!known.has(key)) throw new Error('docx_format: 未知参数 ' + key)
      }
      const target = typeof args?.path === 'string' ? args.path.trim() : ''
      if (!target) throw new Error('docx_format: 缺少 path')
      const mode = args?.mode === 'check' ? 'check' : 'apply'
      const recursive = args?.recursive === true
      const file = resolveCandidate(target)
      if (!existsSync(file)) throw new Error('docx_format: 路径不存在 ' + file)
      if (args?.force !== true) {
        const gate = requestGate(exec?.agent?.id)
        if (!gate.run) {
          const reason = gate.kind === 'skip'
            ? '请求文字含排除词「' + gate.keyword + '」，本插件停用'
            : '请求文字未提到「' + cfg.triggerKeywords.join('|') + '」，本插件停用（确需格式化请传 force=true）'
          log('info', '[docx-format] tool skip ' + file + ' — ' + reason)
          return { ok: true, path: file, mode, changed: 0, violations: 0, fixes: 0,
            changedFiles: 0, skipped: 1, skipReason: reason, warnings: [], files: [] }
        }
      }
      const report = await formatOne(cfg, python, file, mode, recursive)
      if (report.ok && report.changed > 0 && mode === 'apply') formattedByUs.add(file.toLowerCase())
      return report
    },
  }

  /* ------------------------------------------------------ docx_read 工具 */

  const readSchema = {
    type: 'object',
    properties: {
      ok: { type: 'boolean' },
      path: { type: 'string' },
      total_chars: { type: 'number' },
      paragraphs: { type: 'number' },
      tables: { type: 'number' },
      offset: { type: 'number' },
      next_offset: { type: 'number' },
      truncated: { type: 'boolean' },
      text: { type: 'string' },
      error: { type: 'string' },
      notes: { type: 'array', items: { type: 'string' } },
    },
    required: ['ok', 'path'],
    additionalProperties: true,
  }

  const renderRead = (_args, value) => {
    if (!value?.ok) {
      return [{ type: 'text', text: '[FAIL] ' + (value?.path ?? '') + ' — ' + (value?.error ?? '未知错误') }]
    }
    const head = '[OK] ' + value.path + ' — ' + value.total_chars + ' 字 / '
      + value.paragraphs + ' 段 / ' + value.tables + ' 表'
      + (value.offset > 0 ? '（从第 ' + value.offset + ' 字开始）' : '')
    const parts = [head]
    if (Array.isArray(value.notes) && value.notes.length > 0) parts.push('注：' + value.notes.join('；'))
    parts.push('')
    parts.push(value.text ?? '')
    if (value.truncated && typeof value.next_offset === 'number') {
      parts.push('')
      parts.push('…（还有 ' + Math.max(0, value.total_chars - value.next_offset)
        + ' 字未读，继续读就再调一次并把 offset 设为 ' + value.next_offset + '）')
    }
    return [{ type: 'text', text: parts.join('\n') }]
  }

  const readDefinition = {
    name: 'docx_read',
    description: [
      '读取 .docx 的正文纯文本（自动解包 word/document.xml，只用 Python 标准库，',
      '不需要 python-docx、不需要联网、不会改动原文件）。段落按原顺序输出，表格按「[表] 单元格 | 单元格」输出。',
      '长文档分页：返回 next_offset，需要后续内容时再调一次并传 offset=next_offset。',
      '注意：这个工具只读文字，不含图片与脚注；PPT/Excel/PDF 不要用它。',
    ].join(''),
    parameters: {
      type: 'object',
      properties: {
        path: { type: 'string', description: '.docx 的绝对路径' },
        offset: { type: 'number', description: '从第几个字符开始读，默认 0' },
        max_chars: { type: 'number', description: '本次最多返回多少字符，默认 12000，上限 60000' },
      },
      required: ['path'],
      additionalProperties: false,
    },
    output: { schema: readSchema, render: renderRead },
    async execute(args) {
      const known = new Set(['path', 'offset', 'max_chars'])
      for (const key of Object.keys(args ?? {})) {
        if (!known.has(key)) throw new Error('docx_read: 未知参数 ' + key)
      }
      const target = typeof args?.path === 'string' ? args.path.trim() : ''
      if (!target) throw new Error('docx_read: 缺少 path')
      const file = resolveCandidate(target)
      if (!existsSync(file)) throw new Error('docx_read: 文件不存在 ' + file)
      if (!file.toLowerCase().endsWith('.docx')) throw new Error('docx_read: 只支持 .docx，收到 ' + file)
      const offset = Math.max(0, Math.floor(Number(args?.offset ?? 0)) || 0)
      const requested = Math.floor(Number(args?.max_chars ?? 12000)) || 12000
      const maxChars = Math.min(60000, Math.max(200, requested))
      const run = await runScript(python, READER,
        [file, '--offset', String(offset), '--max-chars', String(maxChars), '--json'], cfg.timeoutMs)
      if (!run.ok) return stripUndefined({ ok: false, path: file, error: run.message })
      const parsed = parseReport(run.stdout)
      if (!parsed) {
        const detail = (run.stderr || run.stdout || '').trim().split('\n').slice(-3).join(' | ')
        return stripUndefined({ ok: false, path: file, error: 'docx_read 输出无法解析：' + detail })
      }
      if (parsed.ok === false) {
        return stripUndefined({
          ok: false, path: file,
          error: parsed.error ?? '抽取失败',
          notes: Array.isArray(parsed.notes) ? parsed.notes : [],
        })
      }
      return stripUndefined({
        ok: true,
        path: parsed.path ?? file,
        total_chars: Number(parsed.total_chars) || 0,
        paragraphs: Number(parsed.paragraphs) || 0,
        tables: Number(parsed.tables) || 0,
        offset: Number(parsed.offset) || 0,
        next_offset: parsed.next_offset ?? undefined,
        truncated: parsed.truncated === true,
        text: String(parsed.text ?? ''),
        notes: Array.isArray(parsed.notes) ? parsed.notes : [],
      })
    },
  }

  /* ------------------------------------------------------ 提示词规则 */
  const promptText = () => [
    '## Word 公文格式（dsh-docx-format 已启用）',
    '',
    '本会话任何 .docx 都必须符合《通用公文格式规范》，生成的文档会被自动规范化。',
    '写代码生成 docx 时直接照下面这套参数排，不要等插件事后修。',
    '',
    '### 段落（硬规则）',
    '- 所有段落 `w:spacing w:before="0" w:after="0"`，段前段后一律 0 磅、0 行；行距固定值 29 磅（`line=580 lineRule="exact"`）。',
    '- 需要“空一行”时插入真正的空段落，且该空段落的字体/字号/对齐/缩进要与前后文**完全一致**（等于手工回车继承出来的效果）：',
    '  标题上空 2 行、标题下空 1 行 → 段落标记写成方正小标宋简体 二号(22pt)、居中、无首行缩进（`w:pPr/w:rPr` 里写 `w:rFonts` 与 `w:sz=44`）；',
    '  正文/附件/落款之间的空行 → 仿宋 三号(16pt)、两端对齐、首行缩进 2 字符（`w:firstLineChars="200"`、`w:sz=32`）。',
    '  **不要**用段前/段后间距凑高度，也不要让空行落成默认宋体小四左对齐。',
    '',
    '### 缩进（只有特殊缩进，普通缩进一律 0）',
    '- 正文/标题：`<w:ind w:firstLineChars="200"/>`（2 字符，可留 `w:firstLine="640"` 兜底），左/右/悬挂缩进一律 0。',
    '  禁止写成 `Pt(32)` 之类磅值。',
    '- 附件条目按“手工换行分段落”排，每段只设首行缩进，左/右/悬挂全 0：',
    '  「附件：1．……」第一段 `firstLineChars="200"`（2 字符）；该条目的折行续行单独成段、`firstLineChars="600"`（6 字符）；',
    '  多附件的「2．……」起始段 `firstLineChars="500"`（5 字符）、它的续行仍是 6 字符。',
    '  不要在附件行写 `leftChars/hangingChars`——Word 会把悬挂量反向加到折行上（折行跑到 8 字符）。',
    '',
    '### 字体字号',
    '- 大标题：方正小标宋简体 二号(22pt)，居中。',
    '- 正文：仿宋 三号(16pt)，两端对齐。一级标题 黑体三号、二级 楷体三号、三级 仿宋三号，',
    '  编号形如「一、」「（一）」「1．」（全角点，序号后不加空格）。',
    '- 数字与字母一律 Times New Roman；破折号「——」用 TNR；时间冒号半角。',
    '- 页边距：上 3.6cm / 下 3.0cm / 左 2.7cm / 右 2.7cm。',
    '',
    '### 表格',
    '- 表标题：方正小标宋简体 二号，居中，固定 29 磅；表标题与表格之间插一行空段（行高最小值 0、字号 5 号/10.5pt）。',
    '- 表头 黑体小四(12pt)，表内容 仿宋小四(12pt)；两字词语中间加两个半角空格（`姓  名`）。',
    '- **每一个单元格的每一段都水平居中**（含表头和内容，`w:jc w:val="center"`），并清掉单元格内的左/右/首行缩进；',
    '  单元格垂直居中（`w:vAlign w:val="center"`），段落单倍行距。只居中表头不算合规。',
    '- 表格整体居中、无文字环绕；外边框 1.5 磅、内边框 0.5 磅；行高最小值 0.7cm。',
    '',
    '### 图片',
    '- 图片单独成段（不要和正文挤在同一段）：**居中**（`w:jc w:val="center"`）、**无缩进**（首行/左/右/悬挂全 0）、**单倍行距**',
    '  （`w:line="240" w:lineRule="auto"`，不要套 29 磅固定值），段前段后仍为 0 磅。',
    '- 图片宽度不超过版心（15.6cm）；图注若需要，另起一段居中写「图 1  xxx」。',
    '',
    '### 落款（默认不写）',
    '- **用户没有明确要求署名/日期时，一律不生成落款段落**，写完正文就结束，不要自拟单位名、不要编日期。',
    '- 用户明确要求落款时才写：正文后空两行（两个空段落），右对齐；',
    '  日期写「2026年10月27日」并在末尾留两个全角空格；署名与日期右边缘对齐（需盖章时署名居中于日期）。',
    '',
    '### 其它',
    '- 页码：页脚居中「— 1 —」，宋体四号(14pt)，两侧破折号各留一个空格。',
    '- 附件：正文里「附件：1．xxx」前空一行、首行缩进 2 字符、下与落款空两行；多附件用「2．」等序号，',
    '  条目之间不再插空行；每条自成一个缩进阶梯（2 / 6 / 5 / 6 字符），靠**真正的分段**实现而不是悬挂缩进：',
    '  第 1 段 2 字符（「附件：1．」行首）、第 2 段 6 字符（上一段的折行续行）、第 3 段 5 字符（「2．」）、第 4 段 6 字符……',
    '  折行点由 Word 排版实测得到（`Range.Information(wdFirstCharacterLineNumber)`），插件会自动把长条目拆成多段；',
    '  引擎（Word/WPS）不可用时退回“首行缩进 2/5 字符、不自动拆段”，折行留给使用者手工换行。',
    '  文件末尾「附件1」黑体三号、无缩进、无冒号，其下与附件标题空一行。',
    '',
    '### 触发与白名单',
    '- 只有用户在这条要求里明确提到「公文」（如「按公文格式写一份通知」）时，才套用上面这套版式。',
    '- 用户没提公文时，不要主动套公文格式，按普通文档/报告排版。',
    '- 用户这条要求里出现「实验报告」时，本插件整体停用：不要套用公文格式，按普通报告排版（可直接用默认样式）。',
    '',
    '### 中英混排（与是否公文无关，一律适用）',
    '- 中文内容里出现的英文单词或缩写（如 AI、PPT、OA、Word）前后**不加空格**：写「使用AI工具」「OA系统」「Word文档」。',
    '- 正常的英文句子、英文段落仍按英文习惯加空格：「Please submit the form by Friday.」。',
    '- 中英混排的英文用 Times New Roman；中文用仿宋（标题按各自规定）。',
    '- 段落里不要开「自动调整中文与西文/数字的间距」（`w:autoSpaceDE` / `w:autoSpaceDN` 置 0），',
    '  否则即使没有空格字符，渲染出来仍像「使用 AI 工具」。',
    '',
    '生成 docx 之后，若不确定版面，调用 docx_format（mode=check 先体检，mode=apply 修复）。',
    '',
    '### 读已有 docx',
    '- 要读 .docx 正文，调用 `docx_read`（path=绝对路径）。不要用 read 工具（会报 binary file），',
    '  不要 pip install python-docx，不要上网搜。长文档用返回的 next_offset 续读。',
  ].join('\n')

  if (cfg.tool.enabled) {
    ctx.inject(['tools'], (scope) => {
      scope.effect(() => scope.tools.register(toolDefinition), 'docx-format: docx_format tool')
      scope.effect(() => scope.tools.register(readDefinition), 'docx-format: docx_read tool')
    })
  }

  if (cfg.prompt.enabled) {
    ctx.inject(['systemPrompt'], (scope) => {
      scope.effect(
        () => scope.systemPrompt.section({
          name: 'dsh-docx-format',
          order: cfg.prompt.order,
          interpolate: false,
          text: promptText,
        }),
        'docx-format: prompt section',
      )
    })
  }

  // 缓存用户请求文字（sessions 服务不可用时的兜底来源）
  try {
    ctx.on('session/event', (session, event) => {
      if (event?.type !== 'user/message') return
      const text = messageText(event.data)
      if (text === '') return
      let seq = -1
      try { seq = Number(session?.seq ?? -1) } catch { /* ignore */ }
      requestCache.set(session?.id ?? 'unknown', { seq, text })
    })
  } catch { /* 事件不可用就只依赖 sessions 服务 */ }

  if (cfg.autoFormat) {
    ctx.inject(['tools'], (scope) => {
      scope.effect(
        () => scope.on('tools/result', (exec, result) => {
          if (exec?.name === 'docx_format') return
          if (result?.isError === true) return
          const sessionId = exec?.agent?.id
          const gate = requestGate(sessionId)
          if (!gate.run) {
            log('info', '[docx-format] ' + (gate.kind === 'skip'
              ? ('排除词命中「' + gate.keyword + '」')
              : '本条请求未提「公文」') + ' — 不自动格式化')
            return
          }
          for (const candidate of extractDocxPaths(exec, result)) enqueue(candidate, sessionId)
          schedule()
        }),
        'docx-format: auto format on tools/result',
      )
    })
  }

  log('info', '[docx-format] ready python=' + python
    + ' autoFormat=' + cfg.autoFormat
    + ' trigger=' + JSON.stringify(cfg.triggerKeywords)
    + ' skip=' + JSON.stringify(cfg.skipRequestKeywords))
}
