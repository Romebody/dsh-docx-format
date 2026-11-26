/**
 * dsh-docx-format — 让 DSH 生成的所有 .docx 自动套用《通用公文格式规范》。
 *
 * 三层能力：
 *  1) docx_format 工具：模型可显式调用，apply（规范化）/ check（只体检）。
 *  2) systemPrompt 段：把格式硬规则写进提示词，模型一开始就知道该怎么排版。
 *  3) tools/result 观察者：任何工具产出 .docx 路径后，自动排队规范化（默认只处理本次会话新写的文件）。
 *
 * 请求门控：
 *  - 当前会话最后一条用户消息必须命中 triggerKeywords（默认「公文」），否则整个插件不介入；
 *    所以普通对话里生成的 .docx 不会被套上公文版式。
 * 不介入 = 不排队、不改文件、工具调用直接返回 skipped（显式调用传 force=true 可绕过门控）。
 *
 * 真实排版逻辑在 lib/docxfmt.py（python-docx），本文件只做进程调用与调度。
 */
import { existsSync, readFileSync, statSync } from 'node:fs'
import { spawn } from 'node:child_process'
import { isAbsolute, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import path from 'node:path'

export const name = 'dsh-docx-format'

const HERE = path.dirname(fileURLToPath(import.meta.url))
const ENGINE = path.join(HERE, 'lib', 'docxfmt.py')
const READER = path.join(HERE, 'lib', 'docx_read.py')
const LOADED_AT = Date.now()

/* 反馈入口：出问题时不硬塞表单，只把预填链接和一行诊断交给用户，点不点由他决定。 */
const REPO_URL = 'https://github.com/Romebody/dsh-docx-format'
const PKG_VERSION = (() => {
  try {
    return JSON.parse(readFileSync(path.join(HERE, 'package.json'), 'utf8')).version ?? 'unknown'
  } catch {
    return 'unknown'
  }
})()

const DEFAULTS = {
  autoFormat: true,
  formatExisting: false,
  debounceMs: 1200,
  timeoutMs: 120000,
  python: '',
  maxFilesPerTurn: 8,
  skipPatterns: ['*格式要求*', '*模板*', '~$*', '*/node_modules/*'],
  // 触发词：用户请求文字里必须出现其中之一，插件才套公文版式（空数组 = 任何请求都触发）
  triggerKeywords: ['公文'],
  // 没提触发词时是否仍做全局规范化（分页开关、中文引号、中英之间去空格）
  globalNormalize: true,
  // 字体档位的**持久真源只有一个**：插件页「包含的组件」里「严格公文格式」那一行的启用态。
  // standard = 方正小标宋简体/仿宋/黑体/楷体；strict = 方正小标宋_GBK/仿宋_GB2312/方正黑体_GBK/楷体_GB2312。
  // 请求里出现下面这些词时，只在本会话内临时切严格档：不切组件行、不写配置、重启或换会话即恢复。
  strictKeywords: ['严格公文格式', '严格公文'],
  // 字体自检：enabled=false 关掉检测；autoInstall 只认 'ask'（等用户授权）/'off'。
  fontCheck: { enabled: true, autoInstall: 'ask' },
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
  cfg.strictKeywords = toKeywords(cfg.strictKeywords, DEFAULTS.strictKeywords)
  cfg.fontCheck = { ...DEFAULTS.fontCheck, ...(config?.fontCheck ?? {}) }
  cfg.fontCheck.enabled = cfg.fontCheck.enabled !== false
  cfg.fontCheck.autoInstall = cfg.fontCheck.autoInstall === 'off' ? 'off' : 'ask'
  cfg.globalNormalize = cfg.globalNormalize !== false
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

async function formatOne(cfg, python, target, mode, recursive, light, fontProfile) {
  const args = [target, '--mode', mode, '--json']
  if (light) args.push('--light')
  if (fontProfile === 'strict') args.push('--font-profile', 'strict')
  if (cfg.fontCheck?.enabled === false) args.push('--no-font-check')
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
    light: light === true || files.some((f) => f.light === true) ? true : undefined,
    fontProfile: fontProfile === 'strict' || files.some((f) => f.fontProfile === 'strict') ? 'strict' : undefined,
    fonts: files.find((f) => f.fonts)?.fonts,
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

/* ------------------------------------------------- 严格档组件行（宿主插件页那颗开关） */

/** 「严格公文格式」是 cordis.patch.yml 里的第二条 insert 行，它的启用态就是严格档开关。 */
const STRICT_ROW_ID = 'docx-format-strict'

/**
 * 严格公文格式开关的**唯一持久写入口**：只切「严格公文格式」组件行的启用态。
 * 那一行就是 Loader 的 entry，插件页那颗开关写的就是它；所以这里只调 setPluginEnabled，
 * 不再往插件配置里写 strictFonts（旧键已删除，只在启动时做一次性迁移）。
 */
async function setStrictRow(pluginManager, value) {
  const want = value === true
  const row = await syncStrictRow(pluginManager, want)
  if (!row.synced) return { ok: false, reason: row.reason }
  return { ok: true, enabled: want, reason: row.reason }
}

/** 把「严格公文格式」组件行切到 want；只在值真的不同才写（写一次就是一次 Loader 重载）。 */
async function syncStrictRow(pluginManager, want) {
  if (!pluginManager || typeof pluginManager.setPluginEnabled !== 'function') {
    return { synced: false, reason: 'no-plugin-manager' }
  }
  let row
  try {
    row = await strictRowState(pluginManager)
  } catch (error) {
    return { synced: false, reason: error?.message ?? String(error) }
  }
  if (!row.found || !row.entryId) return { synced: false, reason: 'row-not-found' }
  if (row.enabled === want) return { synced: true, reason: 'unchanged' }
  try {
    await pluginManager.setPluginEnabled(row.entryId, want)
    return { synced: true }
  } catch (error) {
    return { synced: false, reason: error?.message ?? String(error) }
  }
}

/** 读「严格公文格式」组件行的现状：{ found, enabled, entryId }。 */
async function strictRowState(pluginManager) {
  if (!pluginManager || typeof pluginManager.listPlugins !== 'function') {
    return { found: false, enabled: false, entryId: undefined }
  }
  const list = await pluginManager.listPlugins()
  const hit = Array.isArray(list)
    ? list.find((p) => p && (p.patchId === STRICT_ROW_ID
      || p.id === STRICT_ROW_ID
      || (typeof p.id === 'string' && p.id.endsWith(':' + STRICT_ROW_ID))))
    : undefined
  if (!hit) return { found: false, enabled: false, entryId: undefined }
  return { found: true, enabled: hit.enabled === true, entryId: hit.entryId ?? hit.id }
}

/* 会话内临时严格档：请求词命中只影响本会话，不落盘、不写配置、不切组件行。
 * 上限 64 个会话，超了淘汰最早的（Map 保序）；进程重启即空。 */
const SESSION_STRICT_MAX = 64
const sessionStrict = new Map()

function markSessionStrict(sessionId) {
  if (!sessionId) return
  sessionStrict.delete(sessionId)
  sessionStrict.set(sessionId, true)
  while (sessionStrict.size > SESSION_STRICT_MAX) {
    const oldest = sessionStrict.keys().next().value
    sessionStrict.delete(oldest)
  }
}

function isSessionStrict(sessionId) {
  return sessionId ? sessionStrict.has(sessionId) : false
}

/** 有效严格档：组件行处于启用态，或本会话命中过请求词（两者取或）。 */
function strictOn(row, sessionId) {
  return row?.enabled === true || isSessionStrict(sessionId)
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

  /* ---- 「严格公文格式」组件行：缓存 + 后台刷新 ----
   * requestGate 是同步的，而 listPlugins() 是异步的，所以门控只读缓存，
   * 缓存过期（>1.5s）时点火一次后台刷新，不阻塞本条请求。 */
  const pluginManager = () => {
    try {
      const svc = ctx.get('pluginManager')
      return svc && typeof svc.listPlugins === 'function' ? svc : undefined
    } catch {
      return undefined
    }
  }
  let strictRow = { found: false, enabled: false, entryId: undefined }
  let strictRowAt = 0
  const refreshStrictRow = async () => {
    strictRowAt = Date.now()
    try {
      strictRow = await strictRowState(pluginManager())
    } catch { /* 保持旧值 */ }
    return strictRow
  }
  const strictRowNow = () => {
    if (Date.now() - strictRowAt > 1500) void refreshStrictRow()
    return strictRow
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
   * 请求门控：决定这条请求用哪种力度处理。
   *  - 触发词（triggerKeywords，默认「公文」）非空且未命中 → 'light'（只做全局规范化）；globalNormalize=false 时 → 'off'；
   *  - 命中触发词（或触发词为空数组）→ 'full'（连版式一起套）。
   * 全局规范化 = 分页开关全关、中文引号换“”、中英之间不留空格，与是否提公文无关。
   */
  const requestGate = (sessionId) => {
    const text = requestTextOf(sessionId)
    // 字体档：组件行启用态（持久）或本会话说过请求词（临时）——两者取或
    const strictHit = matchKeyword(text, cfg.strictKeywords)
    if (strictHit) markSessionStrict(sessionId)
    const profile = strictOn(strictRowNow(), sessionId) ? 'strict' : 'standard'
    if (cfg.triggerKeywords.length === 0) return { mode: 'full', text, profile, strictHit }
    const trigger = matchKeyword(text, cfg.triggerKeywords)
    if (trigger) return { mode: 'full', text, trigger, profile, strictHit }
    if (cfg.globalNormalize) return { mode: 'light', text, profile, strictHit }
    return { mode: 'off', text, profile, strictHit }
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
        if (gate.mode === 'off') {
          log('info', '[docx-format] skip ' + job.file + ' — '
            + '请求文字未提触发词「' + cfg.triggerKeywords.join('|') + '」')
          continue
        }
        formattedByUs.add(job.file)
        const report = await formatOne(cfg, python, job.file, 'apply', false, gate.mode === 'light', gate.profile)
        handled += 1
        if (report.ok) {
          log('info', '[docx-format] ' + (report.skipped > 0
            ? 'skipped '
            : (report.light ? 'normalized ' : 'formatted ')) + report.path
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
      fontProfile: { type: 'string' },
      fonts: { type: 'object', additionalProperties: true },
      consentRequired: { type: 'object', additionalProperties: true },
      installed: { type: 'array', items: { type: 'object', additionalProperties: true } },
      enabled: { type: 'boolean' },
      strictRow: { type: 'object', additionalProperties: true },
    },
    required: ['ok', 'path', 'mode', 'changed', 'violations'],
    additionalProperties: true,
  }

  /* 出问题时给一条反馈入口：只生成链接与一行诊断，不自动提交任何东西。 */
  const issueUrl = (value) => {
    const head = value?.ok === false
      ? ('执行失败 · ' + PKG_VERSION)
      : ('格式偏差 ' + (value?.violations ?? 0) + ' 处 · ' + PKG_VERSION)
    return REPO_URL + '/issues/new?template=bug_report.yml&title=' + encodeURIComponent('[bug] ' + head)
  }
  const diagnostics = (value) => {
    const rules = {}
    for (const f of Array.isArray(value?.files) ? value.files : []) {
      for (const [rule, info] of Object.entries(f?.rules ?? {})) {
        rules[rule] = (rules[rule] ?? 0) + (Number(info?.count) || 0)
      }
    }
    const top = Object.entries(rules)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 6)
      .map(([rule, n]) => rule + '×' + n)
      .join(', ') || '无偏差规则'
    return 'dsh-docx-format ' + PKG_VERSION + ' · ' + process.platform + ' · node ' + process.version
      + ' · ' + (value?.mode ?? '?') + ' · ' + (value?.fontProfile ?? 'standard') + ' · ' + top
  }

  const renderReport = (value) => {
    const files = Array.isArray(value?.files) ? value.files : []
    const lines = []
    if (value?.message) lines.push(value.message)

    /* ---- 字体 ---- */
    if (value?.mode === 'fonts' || value?.mode === 'installFonts') {
      const cr = value?.consentRequired
      if (value.mode === 'installFonts') {
        lines.push('[' + (value.ok ? 'OK' : 'FAIL') + '] 字体安装 — ' + (value.message ?? ''))
        for (const it of value.installed ?? []) {
          const verb = it.ok ? (it.action === 'exists' ? '已存在' : '已安装') : '失败'
          lines.push('  - ' + verb + ' ' + (it.family ?? it.file ?? '') + (it.dest ? (' → ' + it.dest) : ''))
        }
        return [{ type: 'text', text: lines.join('\n') }]
      }
      const info = value.fonts ?? {}
      const missing = Array.isArray(info.missing) ? info.missing : []
      if (missing.length === 0) {
        lines.push('[OK] 字体齐备（profile=' + (value.fontProfile ?? 'standard') + '）')
      } else {
        lines.push('[FAIL] 缺 ' + missing.length + ' 个字体（profile=' + (value.fontProfile ?? 'standard') + '）：'
          + missing.join('、'))
        if (cr) {
          lines.push('  来源：' + (cr.source ?? ''))
          lines.push('  体积：约 ' + (cr.mb ?? '?') + ' MB')
          lines.push('  要不要现在下载安装？同意的话我传 confirm=' + (cr.token ?? '') + ' 走 mode=installFonts。')
        }
        if (info.hint) lines.push('  ' + info.hint)
      }
      return [{ type: 'text', text: lines.join('\n') }]
    }
    /* ---- 配置开关 ---- */
    if (value?.mode === 'config') {
      lines.push('[' + (value.ok ? 'OK' : 'FAIL') + '] 配置 — 「严格公文格式」组件行'
        + (value.enabled ? '已打开' : '已关闭'))
      return [{ type: 'text', text: lines.join('\n') }]
    }
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
      lines.push('[' + (value.ok ? 'OK' : 'FAIL') + '] ' + value.path + ' — ' + head
        + (value.fontProfile === 'strict' ? '（字体档：严格公文格式）' : '')
        + (value.light === true ? '（本条请求没提「公文」，只做全局规范化：分页开关/中文引号/中英之间不留空格）' : ''))
      const only = files[0]
      if (only?.rules) {
        for (const [rule, info] of Object.entries(only.rules)) {
          const sample = info?.samples?.[0]
          lines.push('  - ' + rule + ' x' + info.count
            + (sample ? (': ' + sample.where + ' ' + (sample.current ?? '') + ' -> ' + (sample.expected ?? '')) : ''))
        }
      }
      for (const w of Array.isArray(value.warnings) ? value.warnings : []) lines.push('  ! ' + w)
    } else {
      lines.push('[' + (value.ok ? 'OK' : 'FAIL') + '] ' + files.length + ' 个文件，修正 '
        + value.changed + ' 处，偏差 ' + value.violations + ' 处'
        + (value.skipped ? ('，跳过 ' + value.skipped + ' 个') : ''))
      for (const f of files) {
        lines.push('  - ' + (f.skipped ? 'SKIP' : (f.ok ? 'OK' : 'FAIL')) + ' ' + f.path
          + (f.skipped ? ' (' + (f.skipReason ?? '') + ')' : ' (fixes ' + f.violations + ')'))
      }
    }
    if (value?.ok === false || (value?.mode === 'check' && value?.violations > 0)) {
      lines.push('')
      lines.push('  · 反馈：' + issueUrl(value))
      lines.push('  · 诊断：' + diagnostics(value))
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
      '用户请求提到「公文」（triggerKeywords）时套完整版式；没提到时仍会做全局规范化',
      '（段落分页六项全关、中文引号换“”、中文里英文前后不加空格），并只改这些；',
      '确需强制套完整版式时传 force=true。',
      'mode=fonts 只报当前字体是否齐备（缺字体时返回 consentRequired，含一次性 token）；',
      'mode=installFonts 需要把 token 原样放进 confirm——只有用户明确同意后才调用。',
    ].join(''),
    parameters: {
      type: 'object',
      properties: {
        path: { type: 'string', description: '要处理的 .docx 文件或目录（绝对路径最稳）。mode=fonts/installFonts/config 时可省略。' },
        mode: { type: 'string', enum: ['apply', 'check', 'fonts', 'installFonts', 'config'], description: 'apply=规范化并写回；check=只体检；fonts=只报告字体；installFonts=按授权安装字体；config=开关「严格公文格式」组件行。默认 apply。' },
        recursive: { type: 'boolean', description: 'path 是目录时是否递归子目录。默认 false。' },
        force: { type: 'boolean', description: '忽略请求触发词强制处理。默认 false。' },
        confirm: { type: 'string', description: 'mode=installFonts 时必填：mode=fonts 返回的 consentRequired.token，原样传回。' },
        preset: { type: 'string', description: 'mode=installFonts 可选：gongwen（普通档）/ gongwen-strict（严格档）。' },
        fromZip: { type: 'string', description: 'mode=installFonts 可选：字体包 zip 的 URL 或本地路径。' },
        enabled: { type: 'boolean', description: 'mode=config 时必填：true=把「严格公文格式」组件行打开（持久，重启仍生效），false=关掉。写的正是插件页那颗开关。' },
      },
      required: [],
      additionalProperties: false,
    },
    output: { schema: outputSchema, render: (_args, value) => renderReport(value) },
    async execute(args, exec) {
      const known = new Set(['path', 'mode', 'recursive', 'force', 'confirm', 'preset', 'fromZip', 'enabled'])
      for (const key of Object.keys(args ?? {})) {
        if (!known.has(key)) throw new Error('docx_format: 未知参数 ' + key)
      }
      const requested = typeof args?.mode === 'string' ? args.mode : 'apply'
      const mode = ['check', 'fonts', 'installFonts', 'config'].includes(requested) ? requested : 'apply'
      const gate = args?.force === true ? undefined : requestGate(exec?.agent?.id)
      const profile = gate?.profile === 'strict' ? 'strict' : undefined

      /* ---- 配置：切「严格公文格式」组件行（与插件页那颗开关同一个后端） ---- */
      if (mode === 'config') {
        if (typeof args?.enabled !== 'boolean') {
          throw new Error('docx_format: mode=config 需要布尔参数 enabled')
        }
        const result = await setStrictRow(pluginManager(), args.enabled)
        if (result.ok) await refreshStrictRow()
        log('info', '[docx-format] config enabled=' + args.enabled + ' ok=' + result.ok
          + (result.ok ? '' : ' reason=' + result.reason))
        if (!result.ok) throw new Error('docx_format: 切换「严格公文格式」组件行失败 — ' + (result.reason ?? 'unknown'))
        return stripUndefined({
          ok: true,
          path: '(配置)',
          mode,
          changed: 0,
          violations: 0,
          enabled: result.enabled,
          strictRow: strictRow,
          message: '「严格公文格式」组件行已' + (result.enabled ? '打开' : '关闭') + '，下一次格式化生效',
        })
      }

      /* ---- 字体：只报告 ---- */
      if (mode === 'fonts') {
        const argv = ['--fonts', '--json']
        if (profile === 'strict') argv.push('--font-profile', 'strict')
        const run = await runEngine(python, argv, cfg.timeoutMs)
        const info = parseReport(run.stdout)
        if (!info) {
          const detail = (run.stderr || run.stdout || '').trim().split('\n').slice(-3).join(' | ')
          throw new Error('docx_format: 字体检测失败 ' + detail)
        }
        const missing = Array.isArray(info.missing) ? info.missing : []
        log('info', '[docx-format] fonts profile=' + (profile ?? 'standard')
          + ' missing=' + JSON.stringify(missing))
        return stripUndefined({
          ok: missing.length === 0,
          path: '(字体)',
          mode,
          changed: 0,
          violations: 0,
          fontProfile: info.profile ?? profile ?? 'standard',
          fonts: info,
          consentRequired: info.consentRequired,
          message: missing.length === 0 ? '字体齐备' : info.hint,
        })
      }

      /* ---- 字体：按授权安装 ---- */
      if (mode === 'installFonts') {
        const confirm = typeof args?.confirm === 'string' ? args.confirm.trim() : ''
        if (!confirm) throw new Error('docx_format: mode=installFonts 需要 confirm 令牌（先用 mode=fonts 取 consentRequired.token）')
        const argv = ['--install-fonts', '--json', '--confirm', confirm]
        if (typeof args?.preset === 'string' && args.preset.trim()) argv.push('--preset', args.preset.trim())
        if (typeof args?.fromZip === 'string' && args.fromZip.trim()) argv.push('--from-zip', args.fromZip.trim())
        const run = await runEngine(python, argv, Math.max(cfg.timeoutMs, 900000))
        const info = parseReport(run.stdout)
        if (!info) {
          const detail = (run.stderr || run.stdout || '').trim().split('\n').slice(-3).join(' | ')
          throw new Error('docx_format: 字体安装失败 ' + detail)
        }
        const installed = Array.isArray(info.installed) ? info.installed : []
        const okCount = installed.filter((i) => i?.ok).length
        log('info', '[docx-format] fonts install confirm=ok installed=' + okCount
          + '/' + installed.length + ' dryRun=' + (info.dryRun === true))
        return stripUndefined({
          ok: okCount > 0 && info.error === undefined,
          path: '(字体)',
          mode,
          changed: 0,
          violations: 0,
          installed,
          message: info.error ?? (okCount + ' 个字体已安装，重启 Word 后生效'),
        })
      }

      const target = typeof args?.path === 'string' ? args.path.trim() : ''
      if (!target) throw new Error('docx_format: 缺少 path')
      const recursive = args?.recursive === true
      const file = resolveCandidate(target)
      if (!existsSync(file)) throw new Error('docx_format: 路径不存在 ' + file)
      let light = false
      if (gate) {
        if (gate.mode === 'off') {
          const reason = '请求文字未提到「' + cfg.triggerKeywords.join('|') + '」，本插件停用（确需格式化请传 force=true）'
          log('info', '[docx-format] tool skip ' + file + ' — ' + reason)
          return { ok: true, path: file, mode, changed: 0, violations: 0, fixes: 0,
            changedFiles: 0, skipped: 1, skipReason: reason, warnings: [], files: [] }
        }
        light = gate.mode === 'light'
      }
      const report = await formatOne(cfg, python, file, mode, recursive, light, profile)
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
    '- 副标题（标题下紧跟的那一行，常以「——」开头，如「——在XX会议上的讲话」）：楷体 三号(16pt)，居中；',
    '  主标题与副标题之间**不插空行**，标题后的空行从副标题之后算起。',
    '- 正文：仿宋 三号(16pt)，两端对齐。一级标题 黑体三号、二级 楷体三号、三级 仿宋三号，',
    '  编号形如「一、」「（一）」「1．」（全角点，序号后不加空格）。',
    '- 数字与字母一律 Times New Roman；公文破折号「——」用 TNR；时间冒号半角。',
    '  中文标点「“”‘’……」与时间区间里的单个连字符「—」（如「8:30—17:30」）必须落仿宋：',
    '  这类字符 Word 按 ascii 字体渲染，不单独拆 run 就会跑进 Times New Roman。',
    '- 页边距：上 3.6cm / 下 3.0cm / 左 2.7cm / 右 2.7cm；纸张 A4（21.0×29.7cm），不要留源文档的 Letter/B5。',
    '- 主送机关（「各学院、各学生组织：」这类称呼）**顶格书写、不缩进**，左对齐；它上面的空行同样顶格。',
    '- 正文（含落款）与附件之间、附件与附件之间用**分页符**断开（`w:br w:type="page"`），',
    '  不要用段落属性 `w:pageBreakBefore` —— 那属于「换行和分页」六项，必须保持关闭。',
    '',
    '### 表格',
    '- 表标题：方正小标宋简体 二号，居中，固定 29 磅；表标题与表格之间插一行空段（行距=最小值 0 磅、字号 5 磅）。',
    '  这个空段只能有**一个** `w:spacing`：重复写两个时 Word 只认最后一个，看起来「设了 0 磅却还是 12 磅」。',
    '- **正文与表标题之间不留空行**：规范只要求「表标题与表格之间空一行」，表格前面多敲的回车要删掉（规则 `table.captionBlankBefore`，期望 0 行）。',
    '- 表头 黑体小四(12pt)，表内容 仿宋小四(12pt)；两字词语中间加两个半角空格（`姓  名`）。',
    '- **每一个单元格的每一段都水平居中**（含表头和内容，`w:jc w:val="center"`），并清掉单元格内的左/右/首行缩进；',
    '  单元格垂直居中（`w:vAlign w:val="center"`），段落单倍行距。只居中表头不算合规。',
    '- 表格整体居中、无文字环绕（不写 `w:tblpPr`）；外边框 1.5 磅、内边框 0.5 磅；行高最小值 0.7cm。',
    '- 列宽固定：`w:tblLayout w:type="fixed"`，整表宽 = 版心 − 2×0.19cm（A4 下约 15.22cm），左缩进 0，',
    '  不允许宽过版心或整体右移；单元格自带的 `w:tcBorders` 会把外框撑成 1.5 磅，必须先删掉。',
    '',
    '### 图片',
    '- 图片单独成段（不要和正文挤在同一段）：**居中**（`w:jc w:val="center"`）、**无缩进**（首行/左/右/悬挂全 0）、**单倍行距**',
    '  （`w:line="240" w:lineRule="auto"`，不要套 29 磅固定值），段前段后仍为 0 磅。',
    '- 图片宽度不超过版心（15.6cm）；图注另起一段：「图 1  xxx」，仿宋 10.5pt、居中、无首行缩进。',
    '',
    '### 落款（默认不写）',
    '- **用户没有明确要求署名/日期时，一律不生成落款段落**，写完正文就结束，不要自拟单位名、不要编日期。',
    '- 用户明确要求落款时才写：正文后空两行（两个空段落），右对齐；日期写「2026年10月27日」，行尾补两个全角空格（真字符）。',
    '  署名行尾补的全角空格数 = 2 + (日期宽 − 署名宽) + 2（**无需盖章**，署名首字比日期首字往左偏两个全角字）；',
    '  需盖章时 = 2 + (日期宽 − 署名宽)/2（署名与日期中心轴对齐）；署名比日期还宽时空格收敛为 0。',
    '  两段都必须写 `w:wordWrap="0"`（否则右对齐会忽略行尾空格，补了也白补），并关掉 `w:autoSpaceDE/DN`（否则日期实际宽度偏大、空格数算少）。',
    '',
    '### 其它',
    '- 页码：页脚居中「— 1 —」，宋体四号(14pt)，两侧破折号各留一个空格。',
    '- 附件：正文里「附件：1．xxx」前空一行、首行缩进 2 字符、下与落款空两行；多附件用「2．」等序号，',
    '  条目之间不再插空行；每条自成一个缩进阶梯（2 / 6 / 5 / 6 字符），靠**真正的分段**实现而不是悬挂缩进：',
    '  第 1 段 2 字符（「附件：1．」行首）、第 2 段 6 字符（上一段的折行续行）、第 3 段 5 字符（「2．」）、第 4 段 6 字符……',
    '  折行点由 Word 排版实测得到（`Range.Information(wdFirstCharacterLineNumber)`），插件会自动把长条目拆成多段；',
    '  引擎（Word/WPS）不可用时退回“首行缩进 2/5 字符、不自动拆段”，折行留给使用者手工换行。',
    '  文件末尾「附件1」黑体三号、无缩进、无冒号，其下空一行后接附件标题（方正小标宋二号、居中）。',
    '',
    '### 字体',
    '- 默认用的是**普通档**：方正小标宋简体 / 仿宋 / 黑体 / 楷体——不是 GBK/GB2312 那一套。',
    '- 严格档是 方正小标宋_GBK / 仿宋_GB2312 / 方正黑体_GBK / 楷体_GB2312，两个入口：',
    '  ① 插件页「包含的组件」里「严格公文格式」那一行的启用态（持久，重启仍生效）；',
    '  ② 用户在本会话里说过「严格公文格式」——只在**当前对话**内生效，不切那行、不改配置，换会话或重启恢复普通档。',
    '- 用户要求打开/关闭严格档时，调 docx_format(mode="config", enabled=true|false)，写的就是那颗开关。',
    '- 数字与字母一律 Times New Roman，两档都一样。',
    '- 不要凭印象说「你缺某字体」：要判断就调 docx_format(mode="fonts")，它返回缺哪几个、约几 MB、从哪下。',
    '- 缺字体只做一件事：把返回的 consentRequired 讲清楚（缺几个、约几 MB、来源），',
    '  问用户一句「要不要现在下载安装」；用户明确同意后，才调 docx_format(mode="installFonts", confirm=<token>)。',
    '- 用户没点头就不要调 installFonts；也不要去别的站点找字体包。安装完提示「重启 Word 后生效」。',
    '',
    '### 触发',
    '- 只有用户在这条要求里明确提到「公文」（如「按公文格式写一份通知」）时，才套用上面这套版式。',
    '- 用户没提公文时，不要主动套公文格式，按普通文档/报告排版——但下面的「全局规则」照旧全部适用。',
    '',
    '### 全局规则（与是否提公文无关，一律适用）',
    '- 段落「换行和分页」：孤行控制、与下段同页、段中不分页、段前分页、取消行号、取消断字，**全部不启用**',
    '  （`w:widowControl="0"`、`w:keepNext="0"`、`w:keepLines="0"`、`w:pageBreakBefore="0"`、',
    '  `w:suppressLineNumbers="0"`、`w:suppressAutoHyphens="0"`）。',
    '- 中文里不要用「」『』当引号，一律用“”‘’（「」只出现在讲格式规则的说明文字里）。',
    '- 文本级细节（一律处理）：时间冒号半角（`9：30` → `9:30`）、小数点半角（`1．5` → `1.5`）、',
    '  单独成段的日期统一成 `2024年10月27日`、文末「附件1：」去掉冒号、字符间距取「标准」（不写 `w:spacing/@w:char`）。',
    '- 段落不要挂网格：`w:snapToGrid w:val="0"`，否则行距会被文档网格改写。',
    '- 一个段落只能有一个 `w:spacing`（重复的会被 Word 的后一个覆盖，也是「行距设了没生效」的常见原因）。',
    '- 中英混排：见下一节。',
    '',
    '### 中英混排（与是否公文无关，一律适用）',
    '- 中文内容里出现的英文单词或缩写（如 AI、PPT、OA、Word）前后**不加空格**：写「使用AI工具」「OA系统」「Word文档」。',    '- 正常的英文句子、英文段落仍按英文习惯加空格：「Please submit the form by Friday.」。',
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

  /* -------------------- 启动：读一次「严格公文格式」组件行 -------------------- */

  void refreshStrictRow().then((row) => {
    log('info', '[docx-format] strictRow found=' + row.found + ' enabled=' + row.enabled)
  })

  /* 一次性迁移：v1.3.1 及以前的配置项 strictFonts 已删除，值 true 时转成组件行启用态。 */
  if (config?.strictFonts === true) {
    void setStrictRow(pluginManager(), true).then((res) => {
      log('info', '[docx-format] 迁移：配置 strictFonts=true → 组件行启用态 '
        + (res.ok ? 'ok' : 'failed(' + (res.reason ?? 'unknown') + ')'))
      if (res.ok) void refreshStrictRow()
    })
  }

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
          if (gate.mode === 'off') {
            log('info', '[docx-format] 本条请求未提「公文」 — 不自动格式化')
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
    + ' strict=' + JSON.stringify(cfg.strictKeywords)
    + ' fontCheck=' + cfg.fontCheck.enabled)
}
