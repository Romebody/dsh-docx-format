# dsh-docx-format

让 DSH 生成的所有 `.docx` 自动符合《通用公文格式规范》。

三层生效方式，任一层命中就够：

| 层 | 机制 | 作用 |
| --- | --- | --- |
| 提示词 | `systemPrompt.section`（order 3260） | 模型生成 docx 前就知道字体、字号、行距、缩进、页边距，直接照规矩排 |
| 工具 | `docx_format` | 模型可显式调用：`mode=check` 体检、`mode=apply` 规范化并写回 |
| 钩子 | `tools/result` 观察者 | 任何工具（含 `ops_bash`、技能脚本）产出 .docx 路径后，自动排队规范化 |

真正的排版逻辑在 `lib/docxfmt.py`（python-docx），插件只负责调度进程。

## 本版（v2）修的四件事

1. **段前段后一律 0 磅，空行用空段落表达**
   所有段落 `w:spacing w:before="0" w:after="0"`（不再用 `beforeLines/afterLines`，也不再拿段前 58pt 冒充空行）。
   注意是**显式写 0**：以前的做法是把 `before/after` 删掉，Word 会继续继承 `docDefaults`/Normal 里的段后间距（模板里就有 `w:after="200"` = 10 磅），
   看起来“段前段后还是不是 0”。
   「空一行」现在是真的插入空段落：标题上空 2 行、标题下空 1 行、附件说明前 1 行、附件说明与落款之间 2 行、落款前 2 行、
   文件末尾 `附件1` 标记与其标题之间 1 行（标记本身上面不插空行）；附件多条目（`1．`/`2．`）视为同一段，条目间不插空行。
   空段落本身同样段前段后 0、固定 29 磅，**且字体/字号/对齐/缩进必须与前后文一致**（等价于在 Word 里手工按回车继承出来的空行）：
   标题上下的空行写成 方正小标宋简体 二号(22pt)、居中、无首行缩进；正文/附件/落款之间的空行写成 仿宋 三号(16pt)、两端对齐、`firstLineChars="200"`。
   实现写在 `_style_blank(par, cfg, kind)`（`kind` 为 `blank` / `blank_title`）里，字体落在段落标记 `w:pPr/w:rPr`，`w:sz` 写磅值×2。
   已有空段落会被复用，不够才补，多了才删；`check` 会把字体/字号/对齐/首行缩进不符的空行报成 `blank.mark`。
2. **首行缩进用「字符」而不是磅；附件按 2 / 6 / 5 / 6 字符分段**
   正文写的是 OOXML `w:ind w:firstLineChars="200"`（2 字符），同时保留 `w:firstLine="640"` 作为不认识 `*Chars` 的阅读器兜底；
   `check` 模式把 `firstLine` 有值但 `firstLineChars` 缺失的情况判为偏差。除首行缩进外，左/右/悬挂缩进一律 0。
   附件不走悬挂缩进（Word 在 `leftChars`+`hangingChars` 并存时把悬挂量加到**折行**上，与 OOXML 语义相反；
   纯磅值虽然几何正确，但缩进阶梯要靠段前段后和左缩进拼，跟“普通缩进全 0”的要求冲突），改成**真正的分段**：
   「附件：1．……」所在段落 `firstLineChars="200"`；该条目的折行续行单独成段、`firstLineChars="600"`（6 字符）；
   多附件的「2．……」段落 `firstLineChars="500"`（5 字符），它的续行同样是 6 字符。
   折行点用 Word 实测：把待测段落单独存成临时文档，交给 Word/WPS 排版后用
   `Range.Information(wdFirstCharacterLineNumber)` 读每个字符落在第几行，按实测边界拆段，
   所以拆出来的每一段在最终打印版面里都正好一行、互不重叠。拆完再体检一遍，幂等。
   探测段落只用磅值左缩进（`w:left`），不能写 `w:leftChars`：实测两者并存时 Word 会把
   “6 字符”当成 30 磅，折行点偏后 2 个字，拆出来的续行会再折一次。
   版面坐标（Word 导 PDF 实测，左边距 76.5pt）：首行 108.6pt = +2 字符、续行 / 6 字符 172.6pt、
   `2．` 5 字符 156.6pt；`tools/smoke_attach_split.py` 会重新量一遍并断言每段正好一行。
   没有可调用的 Word/WPS 时退回 `attachSplit.mode="indent-only"` 的效果：只排 2/5 字符缩进、不拆段，
   并给出 `附件自动分段不可用` 的提示，折行留给使用者手工换行。
   条目文字短于 `attachSplit.minCharsToMeasure`（默认 20 字）时不可能折行，直接跳过、不启动 Word
   （省 3~6 秒；短附件因此仍是 0.2 秒级的体检）。
3. **请求门控（只看触发词）+ 默认不写落款**
   - 触发：看**当前会话最后一条用户消息**，必须命中 `triggerKeywords`（默认 `["公文"]`）才套版式。
     没提「公文」的请求（写周报、排简历、导出任何 docx）整条停用插件，文件保持原样；
     `triggerKeywords` 置成空数组表示任何请求都触发（旧行为）。
   - 停用 = `tools/result` 钩子不排队、队列里的任务出队时再判一次、`docx_format` 直接返回 `skipped`（要强制处理传 `force=true`）。
     会话文字从 `sessions` 服务的 `session.deriveMessages()` 取，另有 `session/event` 缓存兜底。
   - 落款：提示词明确「用户没要求署名/日期就不写落款」；引擎只格式化**文档里已经存在**的落款，绝不凭空新增。
4. **表格单元格完全水平居中**
   表头 + 表内容**每一个单元格的每一段**都设 `w:jc w:val="center"`，并清掉单元格内的 `firstLine/firstLineChars/left/leftChars` 缩进，
   同时保持垂直居中、单倍行距。只居中表头不再算合规。

## 规范内容

- 页面 A4，页边距 上 3.6cm / 下 3.0cm / 左 2.7cm / 右 2.7cm。
- 正文 仿宋 三号(16pt)，两端对齐，首行缩进 2 字符（`w:firstLineChars="200"`），行距**固定值 29 磅**，段前段后 0。
- 大标题 方正小标宋简体 二号(22pt)，居中；标题上方空 2 行、下方空 1 行——用**两个 / 一个空段落**表达。
- 一级标题 黑体三号、二级标题 楷体三号、三级标题 仿宋三号；编号「一、」「（一）」「1．」（全角点，无尾随空格）。
- 数字与字母一律 Times New Roman；破折号 `——` 用 TNR；小数点半角。
- 中英混排：中文内容里的英文单词或缩写（AI、PPT、OA、Word）**前后不加空格**——写「使用AI工具」「OA系统」「Word文档」；
  正常的英文句子仍按英文习惯加空格（`Please submit the form by Friday.`）。中英混排的英文用 Times New Roman。
  引擎里由 `normalize.cjkLatinSpace`（默认 `true`）处理，正文、各级标题、表标题、表格单元格都会清，
  规则名 `text.cjkLatinSpace`；数字不算英文，`2026 年` 这种保持原样。
- 光删空格字符还不够：Word 默认开着「自动调整中文与西文/数字的间距」，没有空格也看着像有空格。
  引擎会给每个段落写 `w:autoSpaceDE="0" w:autoSpaceDN="0"`（规则名 `par.autoSpace`），
  PDF 里实测 `的AI素养`、`OA系统`、`Word文档` 之间已无空隙。
- 表格：标题 方正小标宋简体二号；表头 黑体小四、表内容 仿宋小四；**所有单元格段落水平+垂直居中**、单倍行距；外边框 1.5 磅、内边框 0.5 磅；行高最小值 0.7cm；表格居中无环绕；两字词语中间加两个半角空格（`姓  名`）；表标题与表格之间插一行 0 行高、10.5pt 的空段。
- 图片：图片单独成段，**居中 + 无缩进（首行/左/右/悬挂全 0）+ 单倍行距**（`w:line="240" w:lineRule="auto"`，
  不走 29 磅固定值，避免图片被裁或行高被拉死），段前段后仍 0；判定规则 `par.lineSingle`。
- 页码：页脚居中 `— 1 —`，宋体四号(14pt)，字段拆成 6 个 run，保证 Word / LibreOffice 都渲染成 `— N —`。
- 落款（**只在用户明确要求时才有**）：正文后空两行（两个空段落）、右对齐；日期末尾补两个全角空格；署名与日期右边缘对齐（需盖章时署名居中于日期，`seal: true`）。
- 附件：`附件：1．xxx` 前空 1 行、下与落款空 2 行；缩进阶梯是 2 / 6 / 5 / 6 **字符**，靠真实分段实现：
  第一段 `firstLineChars="200"`、续行段 `600`、`2．…` 段 `500`、它的续行又是 `600`，左右悬挂全 0；
  折行点由 Word 排版实测（`wdFirstCharacterLineNumber`）后自动拆段，拆完每段在版面里正好一行；
  文件末尾 `附件1` 标记用黑体三号、无缩进、两端对齐、无冒号，其下与附件标题空 1 行。
- 幂等：重复 `apply` 完全无副作用，`check` 稳定为 0 偏差。

## 用法

命令行（脱离 DSH 也能用，只要装了 python-docx）：

```bash
python lib/docxfmt.py 报告.docx --mode check          # 只体检，有偏差退出码 2
python lib/docxfmt.py 报告.docx --mode apply          # 规范化并原子写回
python lib/docxfmt.py ./docs --mode apply --recursive # 批量
python lib/docxfmt.py 报告.docx --mode check --json   # 机器可读报告
# 临时关掉引擎侧内容白名单
python lib/docxfmt.py 报告.docx --mode apply --config '{"skipContentKeywords":[]}'
```

模型侧：

```text
docx_format(path="C:/.../通知.docx", mode="apply")
docx_format(path="C:/.../通知.docx", mode="check")
docx_format(path="C:/.../公文目录", mode="apply", recursive=true)
docx_format(path="C:/.../通知.docx", mode="apply", force=true)   # 无视请求文字白名单
```

### 读已有 docx：`docx_read`

```text
docx_read(path="C:/.../毕业论文.docx")                 # 默认返回前 12000 字
docx_read(path="C:/.../毕业论文.docx", max_chars=60000)
docx_read(path="C:/.../毕业论文.docx", offset=12000)    # 用上次返回的 next_offset 续读
```

- 纯标准库实现（`lib/docx_read.py`），**不需要 python-docx、不需要联网、不改原文件**。
- 表格按 `单元格 | 单元格` 输出；返回 `total_chars / paragraphs / tables / next_offset / truncated / notes`。
- 只取正文文字：图片与脚注会以 `notes` 提示，不会静默吞掉。
- 给模型用的一等工具，就是为了替掉「read 工具报 binary file → pip install python-docx → 上网搜」这条死路。
  同理 `docx_format` 的返回值已统一过 `stripUndefined`，不会再触发 DSH 的 `value is not lossless JSON` 校验错误。

命令行等价物：

```bash
python lib/docx_read.py 毕业论文.docx --max-chars 2000
python lib/docx_read.py 毕业论文.docx --offset 2000 --json
```

## 安装

### 方式一：从 npm 安装（推荐）

已发布为 [`dsh-docx-format`](https://www.npmjs.com/package/dsh-docx-format)：

```bash
dsh plugin --profile <profile> add dsh-docx-format
```

收录进插件市场后，也可以直接在市场里搜 `docx` 点安装。

### 方式二：本地目录（开发调试）

1. 把 `dsh-docx-format` 整个目录放到任意位置（记为 `<PLUGIN_DIR>`）。
2. profile 的 `package.json` 里 `dependencies` 增加 `"dsh-docx-format": "link:<PLUGIN_DIR>"`；
   已发布到 npm 时直接写版本号，如 `"dsh-docx-format": "^1.0.0"`。
3. 同一个 `package.json` 的 `dsh.profile.bundles` 追加 `"dsh-docx-format"`。
4. `node_modules/dsh-docx-format` 指向该目录（junction / pnpm link / npm install 均可）。
5. 重启 dsh。

环境要求：Node ≥ 20；Python 3.9+ 与 `python-docx`（DSH 自带运行时里已有）。
装了 Word 或 WPS 时附件会自动按实测折行拆段；没有则退回 2/5 字符缩进、不拆段。

## 配置

`cordis.patch.yml` 里是默认值：

| 键 | 默认 | 说明 |
| --- | --- | --- |
| `autoFormat` | `true` | 关闭后只保留工具与提示词 |
| `formatExisting` | `false` | 只处理插件加载后新写的文件；改 `true` 会连旧文件一起规范化 |
| `debounceMs` | `1200` | 结果钩子去抖 |
| `timeoutMs` | `120000` | 单次 python 超时 |
| `python` | `''` | 显式解释器；空则用 `DSH_DOCX_PYTHON` → `~/.dsh/dsh-runtimes/.../python.exe` → PATH |
| `skipPatterns` | `*格式要求*` 等 | 通配符匹配文件路径，命中不自动改（防止改到格式规范原件） |
| `triggerKeywords` | `['公文']` | **触发词**：会话最后一条用户消息必须含其中之一才介入；空数组 = 任何请求都触发 |
| `maxFilesPerTurn` | `8` | 一轮最多自动处理几个文件 |
| `prompt.enabled / order` | `true / 3260` | 提示词段开关与排序 |
| `defaults` | `{}` | 透传给 docxfmt.py，覆盖任何 DEFAULTS（如 `linePt`、`signature.seal`、`skipContentKeywords`） |

引擎侧的附件相关键（写进 `defaults` 即可覆盖）：

| 键 | 默认 | 说明 |
| --- | --- | --- |
| `attachLineIndentChars` | `2` | 「附件：1．」首行缩进字符数 |
| `attachContIndentChars` | `5` | 多附件「2．」首行缩进字符数 |
| `attachWrapIndentChars` | `6` | 折行续行段的首行缩进字符数 |
| `attachSplit.enabled` | `true` | 关掉就完全不碰附件分段 |
| `attachSplit.mode` | `word` | `word` = 用 Word/WPS 实测折行后自动拆段；`indent-only` = 只排 2/5 缩进、不拆段 |
| `attachSplit.maxRounds` | `12` | 长条目要拆成多段，最多迭代几轮 |
| `attachSplit.minCharsToMeasure` | `20` | 短于此长度的条目直接跳过，不启动 Word |

## 目录

```text
dsh-docx-format/
├── index.js            # 插件入口：工具 + 提示词 + 结果钩子 + 请求文字白名单
├── lib/docxfmt.py      # 排版引擎（apply/check，命令行可用）
├── cordis.patch.yml    # bundle patch（insert 一行配置）
├── locale/zh.json      # 展示文案
├── locale/en.json
├── icon.svg
└── package.json
```

## 排查

- 没生效：重启 dsh；再看启动日志有没有 `[docx-format] ready python=... trigger=["公文"]`。
- 明明该格式化却没动：看日志有没有 `本条请求未提「公文」 — 不自动格式化`；那是门控在生效，用 `force=true` 绕过。
- `无法启动 python`：在配置里写死 `python` 路径。
- 旧文件不自动改：默认 `formatExisting: false`，用 `docx_format` 指定文件，或把它改成 `true`。
- 只想看问题：`docx_format(mode="check")` 打印逐条 `规则 x次数 where 当前 -> 期望`。
- 自测：`node tools/smoke_whitelist.mjs`（工作区根目录），`python tools/smoke_attach_split.py`
  （附件 2/6/5/6 阶梯 + Word 实测每段一行 + 幂等），`python tools/smoke_blank_fmt.py`
  （空段落字体/字号/对齐/缩进是否与前后文一致、幂等），`python tools/smoke_image_line.py`
  （图片段居中/无缩进/单倍行距、正文仍是 29 磅、幂等），`python tools/smoke_cjk_space.py`
  （中英混排去空格 + 英文句子保持 + 表格单元格 + 幂等），`python tools/verify_v2.py <docx> -v` 逐段核对 OOXML。
- 中英之间还看着空：Word 的「自动调整中西文间距」没关。看 `docx_format(mode="check")` 里的 `par.autoSpace`，
  `apply` 会给段落补 `w:autoSpaceDE="0" w:autoSpaceDN="0"`。
- 空行看起来“比手工回车矮/宽”：那是空段落没继承前后文字体。用 `docx_format(mode="check")` 看 `blank.mark`，
  它会给出“当前字体/字号/对齐 -> 期望值”，`apply` 会就地重写段落标记的 `w:pPr/w:rPr`。
