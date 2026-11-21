// 「严格公文格式」组件行。
//
// 这个包**只**做一件事：让宿主插件页的「包含的组件」里出现一行，
// 那一行自带的启用开关就是严格档（GB2312 档字体）的总开关。
//
// 为什么必须单独一个包：patch 行的 EntryOptions 只有 id/name/config/group/disabled/inject，
// 没有 meta 字段；组件行的标题与描述只能取自被引用包的 `locale/*.json` 的 `meta.*`。
// 若和主插件共用一份 locale，同一份文案会把两行写成同一个标题，分不开。
//
// 真正的排版逻辑全在主插件 dsh-docx-format 里：它读这一行的启用态
// （pluginManager.listPlugins() → patchId === 'docx-format-strict' → enabled），
// 得到 `strict = 这一行启用 || 本会话命中过 strictKeywords`，再决定用哪一档字体。

export const name = 'dsh-docx-format-strict'

export function apply() {
  // 故意留空：这一行不注册工具、不注入服务、不写提示词，只提供「存在 + 可启停」。
}
