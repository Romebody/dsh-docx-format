# PR：把 dsh-docx-format 收录进插件目录

打开 https://github.com/awesome-dsh-plugin/awesome-dsh-plugin → Fork，
在新仓库里新建文件 `data/plugins/Romebody__dsh-docx-format.yml`，
内容直接复制 `docs/awesome-entry.yml`（脚本 `tools/publish_market.sh` 可全自动完成，含开 PR）。

- 标题：`Add Romebody/dsh-docx-format`
- 正文：见下方（可直接粘贴）
- 一个 PR 只提这一条；README 由脚本生成，不要手工改

---

## PR 正文（英文，供复制）

```markdown
Adds `dsh-docx-format`, a Docs & Rendering plugin that normalizes generated Word
documents to one Chinese official-document layout.

What it does:

- Registers a `docx_format` tool (`mode=check|apply`) plus a `systemPrompt` rule
  section, so both the model and the user drive the same rule set.
- Applies the layout at OOXML level: FangZheng XiaoBiaoSong two-size title,
  FangSong body at a fixed 29pt leading with `w:before="0" w:after="0"`,
  character-unit first-line indent (`w:firstLineChars="200"` rather than points),
  fully centered table cells, 1.5pt/0.5pt borders, footer `— 1 —` page numbers.
- Formats blank lines the same way a manual Enter does: the empty paragraph's
  paragraph mark carries the neighbouring font, size, alignment and indent.
- Splits long attachment entries into real paragraphs at the wrap point measured
  by Word (`Range.Information(wdFirstCharacterLineNumber)`), producing the
  2 / 6 / 5 / 6 character indent ladder; falls back to 2/5 with no split when
  Word or WPS is unavailable.
- Skips itself entirely when the request text matches a configured allowlist
  keyword (defaults to `实验报告` / lab report), and writes no signature block
  unless the user explicitly asks for one.

Notes for reviewers:

- Manifest: `dsh.bundle.patch` → `cordis.patch.yml`, which holds a single
  `- insert:` entry.
- Zero runtime dependencies: `index.js` imports only `node:fs`,
  `node:child_process`, `node:path`, `node:url`. The formatter is a standard-library
  Python script spawned on demand; `python-docx` is required on the host.
- Published on npm as `dsh-docx-format`.

Smoke tests in the repo cover the allowlist gate, the auto-format hook, the tool,
attachment splitting, blank-line formatting and image paragraphs.
```
