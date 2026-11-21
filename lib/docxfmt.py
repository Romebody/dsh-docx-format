# -*- coding: utf-8 -*-
"""docxfmt.py — 通用公文格式规范化引擎 (dsh-docx-format)

用法:
  python docxfmt.py <file.docx|dir> [--mode apply|check] [--recursive] [--json] [--print] [--config '<json>']
模式:
  apply  规范化并写回（原子替换）
  check  只检查不修改，输出违规项
退出码: 0 成功；1 参数/运行错误；2 check 模式下存在违规。
"""
from __future__ import annotations

import argparse
import base64
import copy
import json
import os
import re
import shutil
import stat
import subprocess
import sys
import tempfile
import time
from collections import Counter

try:
    from docx import Document
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt, Twips
    from docx.table import Table
    from docx.text.paragraph import Paragraph
except Exception as exc:  # pragma: no cover
    print(json.dumps({"ok": False, "error": "python-docx 不可用: %s" % exc}, ensure_ascii=False))
    sys.exit(1)

LEFT = WD_ALIGN_PARAGRAPH.LEFT
CENTER = WD_ALIGN_PARAGRAPH.CENTER
RIGHT = WD_ALIGN_PARAGRAPH.RIGHT
JUSTIFY = WD_ALIGN_PARAGRAPH.JUSTIFY

DEFAULTS = {
    "fonts": {
        "title": "方正小标宋简体",
        "body": "仿宋",
        "h1": "黑体",
        "h2": "楷体",
        "h3": "仿宋",
        "latin": "Times New Roman",
        "footer": "宋体",
        "tableHead": "黑体",
        "tableBody": "仿宋",
        "tableCaption": "方正小标宋简体",
    },
    "sizes": {"title": 22, "body": 16, "table": 12, "footer": 14, "spacer": 10.5,
              "figCaption": 10.5},
    "linePt": 29,
    # 首行缩进以“字符”为单位（w:firstLineChars），不用磅值
    "indentChars": 2,
    # 附件：左/右/悬挂缩进一律 0，只用“首行缩进”这一种特殊缩进。
    #   attach_line  「附件：1．……」      首行 2 字符
    #   attach_cont  「2．……」            首行 5 字符（= 3 字“附件：”+ 2 字“1．”）
    #   attach_wrap  手工换行出来的续行    首行 6 字符
    "attachLineIndentChars": 2,
    "attachContIndentChars": 5,
    "attachWrapIndentChars": 6,
    # 附件条目超出一行时，用 Word/WPS 实测折行点，把视觉行拆成真实段落（续行 6 字符）。
    # 拿不到排版引擎就只设首行缩进，不拆段（由用户手工换行）。
    "attachSplit": {"enabled": True, "mode": "word", "maxRounds": 12,
                    "minCharsToMeasure": 20},
    # 所有段落段前/段后一律 0 磅；空行用独立空段落（整行）表达
    # 纸张：A4（11906×16838 twips = 21.0×29.7cm）；keep = 保留原文档纸张不动
    "pageSize": "A4",
    "marginsCm": {"top": 3.6, "bottom": 3.0, "left": 2.7, "right": 2.7},
    "titleBlankBeforeLines": 2,
    # 标题后的空行从“副标题”算起：主标题与副标题之间不留空行
    "titleBlankAfterLines": 1,
    # 副标题：紧跟主标题（中间最多隔一个空段）的另一段短文本，仿宋三号居中
    "subtitle": {"enabled": True, "mainMaxChars": 18, "maxChars": 40,
                 "gapTolerance": 1},
    "attachBlankBeforeLines": 1,
    "attachBlankAfterLines": 2,
    # 文件中的“附件1”标记：下与附件标题空出一行（参考要求未规定其上要空行）
    "attachTagBlankAfterLines": 1,
    "tableBordersPt": {"outer": 1.5, "inner": 0.5},
    "rowHeightCm": 0.7,
    "pageNumber": True,
    # enabled=False：完全不处理落款（不右对齐、不加空行、不补空格）。
    # 默认 True 只格式化文档里【已有】的落款，绝不新增；“默认不写落款”由提示词层
    # 控制（未明确要求署名/日期时不生成落款文字）。
    "signature": {"enabled": True, "seal": False, "datePadSpaces": 2,
                  "orgOverhangSpaces": 2, "blankLinesBefore": 2},
    # 字体档：standard 用常见简体字体，strict 用 GBK/GB2312 全字库（更接近
    # GB/T 9704-2012 的“标准字体”，缺字更少）。
    # 档位由调用方用 --font-profile 指定（插件算好 `组件行启用态 || 本会话命中请求词` 再传进来）；
    # 直接命令行跑、不传 --font-profile 时一律 standard。
    "fontProfile": "standard",
    # strict 档在 standard 之上覆盖这些位置
    "fontsStrict": {
        "title": "方正小标宋_GBK",
        "tableCaption": "方正小标宋_GBK",
        "body": "仿宋_GB2312",
        "h3": "仿宋_GB2312",
        "tableBody": "仿宋_GB2312",
        "h1": "方正黑体_GBK",
        "tableHead": "方正黑体_GBK",
        "h2": "楷体_GB2312",
        "footer": "宋体",
    },
    # 字体检测与安装。检测只读注册表/字体目录，**从不联网**；
    # 下载必须拿到 consent token（或 CLI --yes）。是否自动下载由插件层
    # （index.js 的 fontCheck.autoInstall）决定，引擎只负责「报告 + 按令牌安装」。
    "fontCheck": {
        "enabled": True,
        "onRun": True,
        "sources": [],
        "installDir": "",
    },
    "normalize": {
        "stripLeadingSpace": True,
        "fullwidthNumberDot": True,
        "twoCharCellSpace": True,
        "dashFont": True,
        "boldOff": True,
        # 中文内容里英文前后不留空格（「使用AI工具」）；纯英文语句的空格保持原样
        "cjkLatinSpace": True,
        # 中文里的引号用“”‘’，不用「」『』
        "quotePair": True,
        # 「换行和分页」里的 6 个开关（孤行控制/与下段同页/段中不分页/段前分页/
        # 取消行号/取消断字）一律不启用
        "paginationOff": True,
        # 字符间距「标准」：run 上写 w:spacing w:val="0"
        "charSpacing": True,
        # 时间冒号用半角：「15：30」→「15:30」
        "timeColon": True,
        # 小数点半角：「3．14」→「3.14」（不碰「1．」这种序号）
        "numberHalfDot": True,
        # 文件末尾的「附件1：」去掉冒号
        "attachTagColon": True,
        # 落款日期统一成「2026年10月27日」
        "dateFormat": True,
    },
}

RE_H1 = re.compile(r"^[ \u3000]*[一二三四五六七八九十百]{1,4}、")
RE_H2 = re.compile(r"^[ \u3000]*[（(][一二三四五六七八九十百]{1,4}[）)]")
RE_H3 = re.compile(r"^[ \u3000]*\d{1,2}\s*[．.、](?!\d)")
RE_ATTACH_TAG = re.compile(r"^[ \u3000]*附件\s*\d{0,2}\s*$")
RE_ATTACH_LINE = re.compile(r"^[ \u3000]*附件\s*[：:]")
RE_ATTACH_ITEM = re.compile(r"^[（(]?\d{1,2}\s*[．.、）)]")
RE_DATE = re.compile(r"\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日")
# 副标题常以破折号开头：「——在XX会议上的讲话」
RE_DASH_LEAD = re.compile(r"^[—–]{1,2}|^-{1,2}\s")
# 「表 1 xxx」「图 3 xxx」是表/图标题，不能被判成副标题
RE_CAPTION_LEAD = re.compile(r"^[表图]\s*[0-9一二三四五六七八九十]+")
# 图注：以「图 1」开头、自成一段的短句（表标题由“紧邻表格”判定，不走这条）
RE_FIG_CAPTION = re.compile(r"^[ 　]*图\s*\d{1,3}(?:[.．、\-—]|\s)")
RE_ADDRESSEE = re.compile(r"^[^。；！？!?]{1,30}[：:]$")
RE_LEAD_SPACE = re.compile(r"^[\u3000 ]+")
RE_FULLWIDTH_SPACE = "\u3000"
# 落款日期的各种写法：2024.10.27 / 2024-10-27 / 2024/10/27 / 2024年10月27日
RE_DATE_LOOSE = re.compile(r"(\d{4})\s*[.\-/年]\s*(\d{1,2})\s*[.\-/月]\s*(\d{1,2})\s*日?")
# 时间里的全角冒号：15：30:05
RE_TIME_COLON = re.compile(r"(\d{1,2})[：:](\d{2})(?=[：:]\d{2}|\s|$|[，。、；）)])")
# 完整时间（可带秒）：9：30 / 15：30：05 —— 前后的数字或冒号都不算，避免把「302会议室」当成时间
RE_TIME_FULL = re.compile(r"(?<![\d：:])(\d{1,2})[：:](\d{2})(?:[：:](\d{2}))?(?![\d：:])")
# 正文里顺手写的点分日期：2024.10.27 / 2024-10-27（整段日期走 RE_DATE_LOOSE）
RE_DATE_INLINE = re.compile(r"(?<!\d)(\d{4})\s*[.\-/]\s*(\d{1,2})\s*[.\-/]\s*(\d{1,2})(?!\d)")
# 小数里的全角点：3．14（后面还得跟数字，才不是「1．」这种序号）
RE_NUMBER_FULL_DOT = re.compile(r"(?<=\d)．(?=\d)")
# 文件末尾的附件标记：附件1：/ 附件 2:
RE_ATTACH_TAG_COLON = re.compile(r"^(附件\s*\d{1,2})\s*[：:]\s*$")

# w:spacing 之后可以出现的元素（按 CT_PPr 的 schema 顺序）
SPACING_SUCCESSORS = (
    "w:ind", "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap",
    "w:jc", "w:textDirection", "w:textAlignment", "w:textboxTightWrap",
    "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange",
)
# 合并重复 w:spacing 时，从前面的节点里挑这几个属性补进来
SPACING_KEEP_ATTRS = ("w:line", "w:lineRule", "w:before", "w:after",
                      "w:beforeLines", "w:afterLines")



# ---------------------------------------------------------------- 工具函数
def ename(value):
    if value is None:
        return None
    return getattr(value, "name", None) or str(value)


def pt_of(length):
    if length is None:
        return None
    try:
        return round(length.pt, 2)
    except Exception:
        return None


def text_width(s):
    """按全角字符数估算显示宽度（半角记 0.5）。"""
    w = 0.0
    for ch in s:
        w += 0.5 if ord(ch) < 0x2000 else 1.0
    return w


def merge_config(user):
    cfg = copy.deepcopy(DEFAULTS)
    if not user:
        return cfg
    for k, v in user.items():
        if isinstance(v, dict) and isinstance(cfg.get(k), dict):
            cfg[k].update(v)
        else:
            cfg[k] = v
    return cfg


# ---------------------------------------------------------------- 报告收集
class Report:
    def __init__(self, path, mode, fix):
        self.path = path
        self.mode = mode
        self.fix = fix
        self.rules = {}          # rule -> {"count":int,"samples":[...]}
        self.warnings = []
        self.stats = Counter()
        self.total = 0

    def want(self, rule, where, current, expected, apply_fn=None):
        """记录一处偏差；fix 模式下调用 apply_fn。返回是否命中偏差。"""
        if current == expected:
            return False
        self.total += 1
        entry = self.rules.setdefault(rule, {"count": 0, "samples": []})
        entry["count"] += 1
        if len(entry["samples"]) < 3:
            entry["samples"].append(
                {"where": where, "current": _s(current), "expected": _s(expected)}
            )
        if self.fix and apply_fn is not None:
            apply_fn()
        return True

    def warn(self, msg):
        if msg not in self.warnings:
            self.warnings.append(msg)

    def as_dict(self, changed):
        return {
            "ok": True,
            "file": self.path,
            "mode": self.mode,
            "changed": bool(changed),
            "violations": self.total,
            "rules": self.rules,
            "stats": dict(self.stats),
            "warnings": self.warnings,
        }


def _s(v):
    if v is None:
        return "无"
    if isinstance(v, bool):
        return "是" if v else "否"
    return str(v)


# ---------------------------------------------------------------- 字体/段落
def _safe_attr(obj, name, default=None):
    """读 python-docx 的属性，但吞掉“枚举映射不了”的异常。

    真实文档里 `w:highlight w:val="none"` 会让 `f.highlight_color` 抛
    `ValueError: WD_COLOR_INDEX has no XML mapping for 'none'`，
    整篇文档直接崩掉。读不到就当它是 default。
    """
    try:
        return getattr(obj, name)
    except (ValueError, KeyError, AttributeError):
        return default


def set_run(run, font, size, latin, report, where, cfg):
    if run.text is None or run.text == "":
        if run._element.findall(".//" + qn("w:drawing")):
            pass
    f = run.font
    rPr = run._element.get_or_add_rPr()

    if cfg["normalize"]["boldOff"]:
        if f.bold:
            report.want("run.bold", where, "加粗", "常规",
                        lambda: setattr(f, "bold", None))
    cur_ascii = f.name
    if cur_ascii != latin:
        report.want("run.font.latin", where, cur_ascii, latin,
                    lambda: setattr(f, "name", latin))
    if rPr.find(qn("w:rFonts")) is None:
        f.name = latin
    rFonts = rPr.find(qn("w:rFonts"))
    cur_ea = rFonts.get(qn("w:eastAsia"))
    if cur_ea != font:
        report.want("run.font.eastAsia", where, cur_ea, font,
                    lambda: rFonts.set(qn("w:eastAsia"), font))
    cur_sz = pt_of(f.size)
    if cur_sz != size:
        report.want("run.font.size", where, cur_sz, size,
                    lambda: setattr(f, "size", Pt(size)))
    hl = rPr.find(qn("w:highlight"))
    if hl is not None:
        val = hl.get(qn("w:val"))
        if val in (None, "none", "auto"):
            report.want("run.highlight", where, "高亮「%s」" % (val or "无值"), "无",
                        (lambda r=rPr, h=hl: r.remove(h)))
        else:
            report.want("run.highlight", where,
                        ename(_safe_attr(f, "highlight_color", val)), "无",
                        (lambda ff=f: setattr(ff, "highlight_color", None)))


def style_runs(par, font, size, latin, report, label, cfg):
    for idx, run in enumerate(par.runs):
        where = "%s#r%d" % (label, idx)
        if (cfg["normalize"]["dashFont"] and run.text
                and set(run.text) <= {"—", "－"} and run.text.strip()):
            # 破折号用 Times New Roman
            set_run(run, latin, size, latin, report, where, cfg)
        else:
            set_run(run, font, size, latin, report, where, cfg)


# ---------------------------------------------------------------- 缩进（字符单位）
# python-docx 没有 firstLineChars / leftChars 的 API，直接写 OOXML 属性。
# Word 优先使用 *Chars 属性（1/100 字符），*FirstLine/*Left 仅作兼容回退值。
IND_SUCCESSORS = ("w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc",
                  "w:textDirection", "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl",
                  "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange")

IND_CHAR_ATTRS = ("w:firstLineChars", "w:leftChars", "w:rightChars",
                  "w:hangingChars", "w:hanging")

# w:autoSpaceDE / autoSpaceDN 的 schema 后继元素（必须插在它们之前）
AUTOSPACE_SUCCESSORS = ("w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing",
                        "w:ind", "w:contextualSpacing", "w:mirrorIndents",
                        "w:suppressOverlap", "w:jc", "w:textDirection",
                        "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl",
                        "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange")

# 空段落有两种：正文侧（沿用仿宋三号正文格式）与标题侧（沿用方正小标宋二号居中）
BLANK_KINDS = ("blank", "blank_title", "blank_add")


def _mark_rpr(par, make=False):
    """段落标记的 w:pPr/w:rPr（空段落的“字体”就存在这里）。"""
    pPr = par._p.pPr
    if pPr is None:
        if not make:
            return None
        pPr = par._p.get_or_add_pPr()
    rPr = pPr.find(qn("w:rPr"))
    if rPr is None and make:
        rPr = OxmlElement("w:rPr")
        pPr.append(rPr)      # w:rPr 是 pPr 的最后一个子元素
    return rPr


def _mark_font(par):
    """读段落标记的 (东亚字体, 字号磅)；未显式设置则为 None（继承样式）。"""
    rPr = _mark_rpr(par)
    if rPr is None:
        return None, None
    font = None
    rf = rPr.find(qn("w:rFonts"))
    if rf is not None:
        font = rf.get(qn("w:eastAsia")) or rf.get(qn("w:ascii"))
    size = None
    sz = rPr.find(qn("w:sz"))
    if sz is not None:
        try:
            size = int(sz.get(qn("w:val"))) / 2.0
        except (TypeError, ValueError):
            size = None
    return font, size


def _write_mark_font(par, font, size_pt, latin):
    """写段落标记的字体字号：手工回车时 Word 也是把字体记在段落标记上。"""
    rPr = _mark_rpr(par, make=True)
    rf = rPr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rPr.insert(0, rf)
    rf.set(qn("w:ascii"), latin)
    rf.set(qn("w:hAnsi"), latin)
    rf.set(qn("w:eastAsia"), font)
    rf.set(qn("w:cs"), latin)
    val = str(int(round(float(size_pt) * 2)))
    for tag in ("w:sz", "w:szCs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = OxmlElement(tag)
            rPr.append(el)
        el.set(qn("w:val"), val)


def _get_ind(par, make=False):
    pPr = par._p.pPr
    if pPr is None:
        if not make:
            return None
        pPr = par._p.get_or_add_pPr()
    ind = pPr.find(qn("w:ind"))
    if ind is None and make:
        ind = OxmlElement("w:ind")
        try:
            pPr.insert_element_before(ind, *IND_SUCCESSORS)
        except Exception:
            pPr.append(ind)
    return ind


def _get_chars(par, attr):
    ind = _get_ind(par)
    if ind is None:
        return None
    raw = ind.get(qn(attr))
    if raw is None or raw == "":
        return None
    try:
        return int(round(int(raw))) / 100.0
    except ValueError:
        return None


def set_chars_indent(par, attr, chars, fallback_pt):
    """把 *Chars 缩进写入 w:ind（chars 为 None 时清除该属性）。"""
    ind = _get_ind(par, make=True)
    char_attr = qn("w:" + attr + "Chars")
    pt_attr = qn("w:" + attr)
    other = {
        "firstLine": "hanging",
        "left": "leftChars",
        "right": "rightChars",
    }.get(attr)
    if chars is None or chars == 0:
        for key in (char_attr, pt_attr):
            if ind.get(key) is not None:
                del ind.attrib[key]
    else:
        ind.set(char_attr, str(int(round(chars * 100))))
        ind.set(pt_attr, str(int(round(fallback_pt.twips))))
    if attr == "firstLine":
        # 首行缩进与悬挂互斥：写首行就清悬挂
        for key in (qn("w:hangingChars"), qn("w:hanging")):
            if ind.get(key) is not None:
                del ind.attrib[key]
    elif attr == "hanging" and chars:
        for key in (qn("w:firstLineChars"), qn("w:firstLine")):
            if ind.get(key) is not None:
                del ind.attrib[key]


def clear_side_indent(par):
    """清掉左/右/悬挂缩进（字符与磅两套属性），保留首行缩进。

    公文里唯一允许的缩进是“首行缩进”，左右缩进与悬挂缩进一律 0。
    """
    ind = _get_ind(par)
    if ind is None:
        return
    for key in ("w:leftChars", "w:left", "w:rightChars", "w:right",
                "w:hangingChars", "w:hanging"):
        if ind.get(qn(key)) is not None:
            del ind.attrib[qn(key)]


def _has_side(chars, pt):
    """侧边缩进是否非 0（字符或磅任一表达）。"""
    if chars:
        return abs(chars) > 0.01
    return pt is not None and abs(pt) > 0.5


def _side_desc(chars, pt):
    """侧边缩进的当前值描述；0 与“无”统一写成 0，避免误判偏差。"""
    if chars:
        return "%.2f 字符" % chars
    if pt is not None and abs(pt) > 0.5:
        return "%.1f 磅" % pt
    return "0"


def _get_pt(ind, attr):
    """读 w:left / w:hanging 的磅值（twips / 20）。"""
    if ind is None:
        return None
    raw = ind.get(qn(attr))
    if raw in (None, ""):
        return None
    try:
        return int(raw) / 20.0
    except ValueError:
        return None


def _style_blank(par, cfg, kind="blank"):
    """空行 = 独立空段落，但格式要与“手工回车”的结果一致：

    - 标题上/下那几行沿用标题格式：方正小标宋简体二号、居中、无首行缩进；
    - 其余空行沿用正文格式：仿宋三号、两端对齐、首行缩进 2 字符。

    段落标记（w:pPr/w:rPr）就是 Word 存“空段落字体”的地方，所以字体字号写在那里。
    """
    title_side = kind == "blank_title"
    add_side = kind == "blank_add"
    pf = par.paragraph_format
    _sole_spacing(par._p)          # 先把重复的 w:spacing 合并掉，下面 pf.* 写的就是唯一节点
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.line_spacing = Pt(cfg["linePt"])
    _write_zero_spacing(par)
    # 主送机关是顶格的，紧挨着它的空行也要顶格，否则视觉上“空行比正文更缩进”
    pf.alignment = CENTER if title_side else (LEFT if add_side else JUSTIFY)
    font = cfg["fonts"]["title"] if title_side else cfg["fonts"]["body"]
    size = cfg["sizes"]["title"] if title_side else cfg["sizes"]["body"]
    _write_mark_font(par, font, size, cfg["fonts"]["latin"])
    chars = 0.0 if (title_side or add_side) else float(cfg["indentChars"])
    set_chars_indent(par, "firstLine", chars, Pt(chars * float(cfg["sizes"]["body"])))
    set_chars_indent(par, "left", None, Pt(0))
    set_chars_indent(par, "right", None, Pt(0))


def style_par(par, kind, report, label, cfg):
    pf = par.paragraph_format
    fonts = cfg["fonts"]
    sizes = cfg["sizes"]
    line = Pt(cfg["linePt"])
    body_pt = Pt(sizes["body"])
    indent_chars = float(cfg["indentChars"])   # 首行缩进 2 字符

    if kind in BLANK_KINDS:
        report.stats["para.blank"] += 1
        title_side = kind == "blank_title"
        add_side = kind == "blank_add"
        b_font = fonts["title"] if title_side else fonts["body"]
        b_size = float(sizes["title"] if title_side else sizes["body"])
        b_align = CENTER if title_side else (LEFT if add_side else JUSTIFY)
        b_chars = 0.0 if (title_side or add_side) else indent_chars
        if pf.line_spacing_rule != WD_LINE_SPACING.EXACTLY or pt_of(pf.line_spacing) != float(cfg["linePt"]):
            report.want("blank.lineSpacing", label, pt_of(pf.line_spacing), float(cfg["linePt"]),
                        (lambda: (setattr(pf, "line_spacing_rule", WD_LINE_SPACING.EXACTLY),
                                  setattr(pf, "line_spacing", line))))
        _zero_spacing(par, pf, report, label, cfg)
        # 空段落也要像手工回车那样带着前后文的字体、字号、对齐与首行缩进
        cur_font, cur_size = _mark_font(par)
        cur_chars = _get_chars(par, "w:firstLineChars")
        drift = []
        if cur_font != b_font:
            drift.append("字体 %s" % (cur_font or "继承样式"))
        if cur_size is None or abs(cur_size - b_size) > 0.01:
            drift.append("字号 %s" % ("继承样式" if cur_size is None else "%g 磅" % cur_size))
        if pf.alignment != b_align:
            drift.append("对齐 %s" % ("居中" if pf.alignment == CENTER
                                     else "两端对齐" if pf.alignment == JUSTIFY
                                     else ("左对齐" if pf.alignment is None
                                           else str(pf.alignment))))
        if abs((cur_chars or 0.0) - b_chars) > 0.01:
            drift.append("首行缩进 %.2f 字符" % (cur_chars or 0.0))
        if drift:
            report.want("blank.mark", label, "、".join(drift),
                        "%s %g 磅、%s、首行缩进 %g 字符"
                        % (b_font, b_size,
                           "居中" if title_side else "两端对齐", b_chars),
                        (lambda p=par, k=kind: _style_blank(p, cfg, k)))
        _clear_indent_chars(par, report, label, kind, b_chars)
        return

    if kind == "caption_spacer":
        report.stats["para.blank"] += 1
        if not _spacer_props_ok(par._p, cfg):
            report.want("table.captionSpacer", label, "空行行高/字号不符",
                        "行高最小值 0 磅、5 号字",
                        lambda: _write_spacer_props(par._p, cfg))
        return

    want_align = None
    want_font = fonts["body"]
    want_size = sizes["body"]
    want_first_chars = indent_chars
    want_left_chars = 0.0
    want_hang_chars = 0.0
    want_line_single = False       # 图片行用单倍行距，不套 29 磅固定值

    if kind in ("title", "title_first", "title_mid", "title_last", "title_only"):
        want_align, want_font, want_size = CENTER, fonts["title"], sizes["title"]
        want_first_chars = None
    elif kind == "subtitle":
        # 副标题：仿宋三号、居中、无首行缩进，行距仍固定 29 磅
        want_align, want_font, want_size = CENTER, fonts["body"], sizes["body"]
        want_first_chars = None
    elif kind == "table_caption":
        want_align, want_font, want_size = CENTER, fonts["tableCaption"], sizes["title"]
        want_first_chars = None
    elif kind == "h1":
        want_font = fonts["h1"]
    elif kind == "h2":
        want_font = fonts["h2"]
    elif kind == "h3":
        want_font = fonts["h3"]
    elif kind == "attach_tag":
        want_align, want_font, want_first_chars = JUSTIFY, fonts["h1"], None
    elif kind == "attach_line":
        # 「附件：1．……」：左/右/悬挂全 0，只有首行缩进 2 字符
        want_align = JUSTIFY
        want_first_chars = float(cfg["attachLineIndentChars"])
    elif kind == "attach_cont":
        # 多附件的“2．……”：首行缩进 5 字符，与冒号后序号同一竖线
        want_align = JUSTIFY
        want_first_chars = float(cfg["attachContIndentChars"])
    elif kind == "attach_wrap":
        # 手工换行出来的续行：首行缩进 6 字符
        want_align = JUSTIFY
        want_first_chars = float(cfg["attachWrapIndentChars"])
    elif kind == "sign_org":
        if cfg["signature"]["enabled"]:
            want_align, want_first_chars = RIGHT, None
        else:
            want_align, want_first_chars = None, None   # 不写落款：不改造版式
    elif kind == "sign_date":
        if cfg["signature"]["enabled"]:
            want_align, want_first_chars = RIGHT, None
        else:
            want_align, want_first_chars = None, None
    elif kind == "attach_title":
        # 文末「附件1」下面那一行附件标题：同大标题（方正小标宋二号、居中、无首行缩进）
        want_align, want_font, want_size = CENTER, fonts["title"], sizes["title"]
        want_first_chars = None
    elif kind == "addressee":
        # 主送机关（"各学院、各学生组织："这类称呼）：顶格书写、不缩进
        want_align, want_font, want_size = LEFT, fonts["body"], sizes["body"]
        want_first_chars = None
    elif kind == "fig_caption":
        # 图注：另起一段、居中、五号仿宋、无首行缩进，行距仍固定 29 磅
        want_align, want_font = CENTER, fonts["body"]
        want_size = sizes.get("figCaption", 10.5)
        want_first_chars = None
    elif kind == "image":
        # 图片所在行：居中、无缩进（首行/左右/悬挂全 0）、单倍行距
        want_align, want_first_chars = CENTER, None
        want_line_single = True
        image_inline(par, report, label)
    else:
        want_align = JUSTIFY

    report.stats["para." + kind] += 1

    def do_align():
        pf.alignment = want_align

    if want_align is None:
        if pf.alignment is not None:
            report.want("par.alignment", label, ename(pf.alignment), "无",
                        lambda: setattr(pf, "alignment", None))
    else:
        report.want("par.alignment", label, ename(pf.alignment), ename(want_align), do_align)

    if want_line_single:
        _check_single_line(par, pf, report, label)
    else:
        cur_rule = ename(pf.line_spacing_rule)
        if pf.line_spacing_rule != WD_LINE_SPACING.EXACTLY:
            report.want("par.lineRule", label, cur_rule, "EXACTLY",
                        lambda: setattr(pf, "line_spacing_rule", WD_LINE_SPACING.EXACTLY))
        if pt_of(pf.line_spacing) != float(cfg["linePt"]):
            report.want("par.lineSpacing", label, pt_of(pf.line_spacing), float(cfg["linePt"]),
                        lambda: setattr(pf, "line_spacing", line))

    # 首行缩进：以字符为单位（w:firstLineChars=200 → 2 字符）
    cur_chars = _get_chars(par, "w:firstLineChars")
    if want_first_chars is None:
        if cur_chars is not None:
            report.want("par.firstLineChars", label, "%.2f 字符" % cur_chars, "无",
                        (lambda p=par: set_chars_indent(p, "firstLine", None, body_pt)))
        if not want_hang_chars and pf.first_line_indent is not None \
                and abs(pt_of(pf.first_line_indent) or 0) > 1:
            # 悬挂缩进时 python-docx 会把 w:hanging 读成负的 first_line_indent，
            # 那种情况由下面的 hanging 分支管，这里不能判它违规。
            report.want("par.firstLineIndent", label, pt_of(pf.first_line_indent), "无",
                        (lambda p=par: set_chars_indent(p, "firstLine", None, body_pt)))
    else:
        if cur_chars is None or abs(cur_chars - want_first_chars) > 0.01:
            report.want("par.firstLineChars", label,
                        "无（%.1f 磅）" % (pt_of(pf.first_line_indent) or 0) if cur_chars is None
                        else "%.2f 字符" % cur_chars,
                        "%.2f 字符" % want_first_chars,
                        (lambda p=par, c=want_first_chars, b=body_pt:
                         set_chars_indent(p, "firstLine", c, Pt(c * b.pt))))

    # 侧边缩进（左 / 右 / 悬挂）一律 0：可用的“特殊缩进”只有首行缩进
    ind = _get_ind(par)
    cur_left_chars = _get_chars(par, "w:leftChars")
    cur_left_pt = _get_pt(ind, "w:left")
    cur_hang_chars = _get_chars(par, "w:hangingChars")
    cur_hang_pt = _get_pt(ind, "w:hanging")

    if want_left_chars:
        if (cur_left_chars is None or abs(cur_left_chars - want_left_chars) > 0.01
                or _has_side(None, cur_left_pt)):
            report.want("par.leftChars", label, _side_desc(cur_left_chars, cur_left_pt),
                        "%.2f 字符" % want_left_chars,
                        (lambda p=par, c=want_left_chars, b=body_pt:
                         set_chars_indent(p, "left", c, Pt(c * b.pt))))
    elif _has_side(cur_left_chars, cur_left_pt):
        report.want("par.leftIndent", label, _side_desc(cur_left_chars, cur_left_pt), "0",
                    (lambda p=par: clear_side_indent(p)))

    if want_hang_chars:
        if (cur_hang_chars is None or abs(cur_hang_chars - want_hang_chars) > 0.01
                or _has_side(None, cur_hang_pt)):
            report.want("par.hangingChars", label, _side_desc(cur_hang_chars, cur_hang_pt),
                        "%.2f 字符" % want_hang_chars,
                        (lambda p=par, c=want_hang_chars, b=body_pt:
                         set_chars_indent(p, "hanging", c, Pt(c * b.pt))))
    elif _has_side(cur_hang_chars, cur_hang_pt):
        report.want("par.hangingIndent", label, _side_desc(cur_hang_chars, cur_hang_pt), "0",
                    (lambda p=par: clear_side_indent(p)))

    _zero_spacing(par, pf, report, label, cfg)

    cur_right = pt_of(pf.right_indent) or 0
    # 落款两行用右缩进摆位（署名超出/中心轴对齐），不参与「缩进一律 0」
    if kind not in ("sign_org", "sign_date") and abs(cur_right) > 1:
        report.want("par.rightIndent", label, cur_right, 0,
                    (lambda: setattr(pf, "right_indent", None)))

    # 中英文之间不留自动间距（否则没有空格字符也看着像有空格）
    auto_space_off(par, report, label)

    if par.runs:
        style_runs(par, want_font, want_size, fonts["latin"], report, label, cfg)


def _spacings(pPr):
    return [] if pPr is None else pPr.findall(qn("w:spacing"))


def _effective_spacing(pPr):
    """Word 渲染时真正生效的 w:spacing。

    同一个 pPr 下可能有多个 w:spacing（Word 自己不会这么写，但不少生成器会），
    Word 以**最后一个**为准；python-docx 的 `pPr.spacing` 返回的却是第一个。
    于是「第一个写对了、最后一个还是旧值」的文档会被误判为合规 —— 表标题与表格
    之间那一行的行距就是这么丢的。
    """
    items = _spacings(pPr)
    return items[-1] if items else None


def _sole_spacing(p_el):
    """返回唯一的 w:spacing：有重复就先合并成一个（以最后一个为基准）。"""
    pPr = p_el.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        p_el.insert(0, pPr)
    items = pPr.findall(qn("w:spacing"))
    if not items:
        sp = OxmlElement("w:spacing")
        pPr.insert_element_before(sp, *SPACING_SUCCESSORS)
        return sp
    keep = items[-1]
    for extra in items[:-1]:
        for attr in SPACING_KEEP_ATTRS:
            if keep.get(qn(attr)) is None and extra.get(qn(attr)) is not None:
                keep.set(qn(attr), extra.get(qn(attr)))
        pPr.remove(extra)
    return keep


def _spacing_desc(sp):
    if sp is None:
        return "无"
    return "line=%s %s / before=%s / after=%s" % (
        sp.get(qn("w:line")) or "-", sp.get(qn("w:lineRule")) or "auto",
        sp.get(qn("w:before")) or "-", sp.get(qn("w:after")) or "-")


def dedupe_spacing(par, report, label):
    """一个 pPr 只留一个 w:spacing；重复的以最后一个（真正生效的那个）为准合并。"""
    pPr = par._p.find(qn("w:pPr"))
    items = _spacings(pPr)
    if len(items) <= 1:
        return False
    keep = items[-1]
    report.want("par.spacingDup", label,
                "%d 个 w:spacing（生效的是最后一个：%s）" % (len(items), _spacing_desc(keep)),
                "只保留 1 个 w:spacing",
                (lambda p=par: _sole_spacing(p._p)))
    return True


def _spacing_elm(par):
    return _effective_spacing(par._p.find(qn("w:pPr")))


def _is_single_line(par):
    """w:line="240" + w:lineRule="auto"（或缺省 lineRule）即标准单倍行距。"""
    sp = _spacing_elm(par)
    if sp is None or sp.get(qn("w:line")) != "240":
        return False
    return (sp.get(qn("w:lineRule")) or "auto") == "auto"


def _write_single_line(par):
    sp = _sole_spacing(par._p)
    sp.set(qn("w:line"), "240")
    sp.set(qn("w:lineRule"), "auto")


def _check_single_line(par, pf, report, label):
    """图片行：要求单倍行距（w:line=240 lineRule=auto），不套 29 磅固定值。"""
    if _is_single_line(par):
        return
    sp = _spacing_elm(par)
    raw = None if sp is None else sp.get(qn("w:line"))
    if raw is None:
        cur = "继承样式"
    else:
        try:
            val = float(raw)
        except (TypeError, ValueError):
            val = 0.0
        rule = sp.get(qn("w:lineRule")) or "auto"
        if rule == "exact":
            cur = "%g 磅(固定)" % round(val / 20.0, 1)
        elif rule == "atLeast":
            cur = "%g 磅(最小值)" % round(val / 20.0, 1)
        else:
            cur = "%g 倍" % round(val / 240.0, 2)
    report.want("par.lineSingle", label, cur, "单倍行距",
                lambda: _write_single_line(par))


def _write_zero_spacing(par):
    """把段前/段后显式写成 0 磅，并清掉“行”单位与自动间距。

    只把 w:before / w:after 删掉是不行的：Word 会继续继承 docDefaults 或
    Normal 样式里的段后间距（常见 w:after="200" = 10 磅），用户看到的就是
    “段前段后不是 0 磅”。
    """
    pPr = par._p.get_or_add_pPr()
    sp = _sole_spacing(par._p)
    for key in ("w:beforeLines", "w:afterLines",
                "w:beforeAutospacing", "w:afterAutospacing"):
        sp.attrib.pop(qn(key), None)
    sp.set(qn("w:before"), "0")
    sp.set(qn("w:after"), "0")


def _zero_spacing(par, pf, report, label, cfg):
    """段前/段后一律 0 磅：间距由独立空段落表达，不用段前段后。"""
    pPr = par._p.get_or_add_pPr()
    sp = _effective_spacing(pPr)
    need = sp is None
    if sp is not None:
        for key in ("w:before", "w:after"):
            if (sp.get(qn(key)) or "") != "0":
                need = True
        for key in ("w:beforeLines", "w:afterLines",
                    "w:beforeAutospacing", "w:afterAutospacing"):
            if sp.get(qn(key)) is not None:
                need = True
    if need:
        report.want("par.spaceZero", label, "非 0 磅或继承样式", "0 磅",
                    lambda: _write_zero_spacing(par))


def _clear_indent_chars(par, report, label, kind, first_chars=0.0):
    """清掉不该有的字符缩进；*first_chars* 是本 kind 期望的首行缩进（默认 0）。"""
    for attr, name in (("firstLine", "firstLineChars"), ("left", "leftChars"),
                       ("right", "rightChars")):
        cur = _get_chars(par, "w:" + name)
        if cur is None:
            continue
        if attr == "firstLine" and abs(cur - float(first_chars)) <= 0.01:
            continue
        report.want("par." + name, label, "%.2f 字符" % cur, "无",
                    (lambda p=par, a=attr: set_chars_indent(p, a, None, Pt(0))))
    cur_hang = _get_chars(par, "w:hangingChars")
    if cur_hang is not None:
        report.want("par.hangingChars", label, "%.2f 字符" % cur_hang, "无",
                    (lambda p=par: set_hanging_chars(p, None, Pt(0))))


# ---------------------------------------------------------------- 文本规范化
CJK_PUNCT = "，。、；：？！“”‘’（）【】《》〈〉「」『』…—～·　"
ASCII_LATIN = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
ASCII_SPACE = " \t\u3000"


def is_cjk_char(ch):
    """汉字或中文标点。"""
    if not ch:
        return False
    if ch in CJK_PUNCT:
        return True
    return "\u4e00" <= ch <= "\u9fff" or "\u3400" <= ch <= "\u4dbf"


ALNUM_CHARS = ASCII_LATIN + "0123456789"


def _token_bounds(text, pos, step):
    """从 pos 起沿 step 方向吃掉连续的字母数字，返回区间 [start, end)。"""
    i = pos
    while 0 <= i < len(text) and text[i] in ALNUM_CHARS:
        i += step
    return (i + 1, pos + 1) if step < 0 else (pos, i)


def _token_has_latin(text, start, end):
    return any(ch in ASCII_LATIN for ch in text[start:end])


def cjk_latin_gaps(text):
    """中文与英文之间的空白区间 [(start, end)]；纯中文/纯数字相邻不返回。

    英文按“词”判定：含字母的字母数字串（AI、Word、5G、A4）算英文，
    纯数字（2026）不算，所以「2026 年」保持不变。
    """
    gaps = []
    i = 0
    n = len(text)
    while i < n:
        if text[i] not in ASCII_SPACE:
            i += 1
            continue
        j = i
        while j < n and text[j] in ASCII_SPACE:
            j += 1
        left_cjk = i > 0 and is_cjk_char(text[i - 1])
        right_cjk = j < n and is_cjk_char(text[j])
        left_latin = right_latin = False
        if i > 0 and text[i - 1] in ALNUM_CHARS:
            s, e = _token_bounds(text, i - 1, -1)
            left_latin = _token_has_latin(text, s, e)
        if j < n and text[j] in ALNUM_CHARS:
            s, e = _token_bounds(text, j, 1)
            right_latin = _token_has_latin(text, s, e)
        if (left_cjk and right_latin) or (left_latin and right_cjk):
            gaps.append((i, j))
        i = j
    return gaps


def drop_char_ranges(runs, gaps):
    """按全局字符位置删除 runs 里对应的字符（空格可能落在任一 run）。"""
    drop = set()
    for start, end in gaps:
        drop.update(range(start, end))
    pos = 0
    for run in runs:
        t = run.text
        if not t:
            continue
        kept = "".join(ch for k, ch in enumerate(t) if (pos + k) not in drop)
        pos += len(t)
        if kept != t:
            run.text = kept


# 「换行和分页」里那 6 个开关对应的 OOXML 属性，顺序即 pPr 的 schema 顺序
PAGINATION_ORDER = ("w:keepNext", "w:keepLines", "w:pageBreakBefore", "w:framePr",
                    "w:widowControl", "w:numPr", "w:suppressLineNumbers", "w:pBdr",
                    "w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku",
                    "w:wordWrap", "w:overflowPunct", "w:topLinePunct",
                    "w:autoSpaceDE", "w:autoSpaceDN", "w:bidi", "w:adjustRightInd",
                    "w:snapToGrid", "w:spacing", "w:ind", "w:contextualSpacing",
                    "w:mirrorIndents", "w:suppressOverlap", "w:jc",
                    "w:textDirection", "w:textAlignment", "w:textboxTightWrap",
                    "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr",
                    "w:pPrChange")
PAGINATION_TAGS = ("w:widowControl", "w:keepNext", "w:keepLines",
                   "w:pageBreakBefore", "w:suppressLineNumbers",
                   "w:suppressAutoHyphens")
PAGINATION_LABELS = {"w:widowControl": "孤行控制", "w:keepNext": "与下段同页",
                     "w:keepLines": "段中不分页", "w:pageBreakBefore": "段前分页",
                     "w:suppressLineNumbers": "取消行号",
                     "w:suppressAutoHyphens": "取消断字"}

# 中文引号：不用「」『』，用“”‘’
QUOTE_MAP = {"「": "“", "」": "”", "『": "‘", "』": "’"}


def _pagination_successors(tag):
    return PAGINATION_ORDER[PAGINATION_ORDER.index(tag) + 1:]


def clear_pagination(par, report, label):
    """段落「换行和分页」里的 6 个开关一律不启用（对话框里复选框全部留空）。

    这 6 项都不在 python-docx 的 API 里，直接写 pPr：
    孤行控制 / 与下段同页 / 段中不分页 / 段前分页 / 取消行号 / 取消断字。
    widowControl 的 Word 默认值是“开”，所以必须显式写 0，不能靠删除元素。
    """
    pPr = par._p.get_or_add_pPr()
    for tag in PAGINATION_TAGS:
        el = pPr.find(qn(tag))
        val = None if el is None else (el.get(qn("w:val")) or "1")
        if el is not None and str(val).lower() in ("0", "false", "off"):
            continue

        def apply(t=tag, e=el, p=pPr):
            node = e
            if node is None:
                node = OxmlElement(t)
                try:
                    p.insert_element_before(node, *_pagination_successors(t))
                except Exception:
                    p.append(node)
            node.set(qn("w:val"), "0")

        report.want("par.pagination." + tag.split(":")[1], label,
                    "%s：启用" % PAGINATION_LABELS[tag],
                    "%s：不启用" % PAGINATION_LABELS[tag], apply)


def quote_normalize(par, report, label):
    """中文里的引号写成“”‘’，不写「」『』。"""
    for idx, run in enumerate(par.runs):
        t = run.text or ""
        if not any(ch in QUOTE_MAP for ch in t):
            continue
        new = "".join(QUOTE_MAP.get(ch, ch) for ch in t)
        report.want("text.quote", "%s#r%d" % (label, idx),
                    "使用「」『』", "使用“”‘’",
                    (lambda r=run, v=new: setattr(r, "text", v)))


def _has_page_break(par):
    """段落里是否已有分页符（w:br w:type="page"）。"""
    for br in par._p.iter(qn("w:br")):
        if br.get(qn("w:type")) == "page":
            return True
    return False


def _insert_page_break(par):
    r = OxmlElement("w:r")
    # 分页符 run 也要带上字体字号，否则检查器会把它算成「字体不符」
    src = None
    for other in par.runs:
        if other._element.find(qn("w:rPr")) is not None:
            src = other._element.find(qn("w:rPr"))
            break
    if src is not None:
        r.append(copy.deepcopy(src))
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    r.append(br)
    pPr = par._p.find(qn("w:pPr"))
    if pPr is not None:
        pPr.addnext(r)
    else:
        par._p.insert(0, r)


def ensure_page_break(par, report, label):
    """正文与附件之间、附件与附件之间用分页符断开。

    不用 w:pageBreakBefore —— 那属于「换行和分页」六项，规范要求一律关闭；
    分页符是正文里的 w:br 节点，与那六项互不影响。
    """
    if _has_page_break(par):
        return
    report.want("par.pageBreak", label, "本段之前没有分页符", "正文/附件之间分页",
                lambda: _insert_page_break(par))


def _iter_all_paragraphs(doc):
    """文档里所有段落（正文 + 表格单元格 + 文本框 + 页眉页脚），每个 XML 节点只回一次。"""
    roots = [doc.element.body]
    for section in doc.sections:
        for part in (section.header, section.footer,
                     section.first_page_header, section.first_page_footer,
                     section.even_page_header, section.even_page_footer):
            try:
                if part is not None and not part.is_linked_to_previous:
                    roots.append(part._element)
            except Exception:
                continue
    for root in roots:
        for node in root.iter(qn("w:p")):
            yield Paragraph(node, doc)


SNAP_SUCCESSORS = (
    "w:spacing", "w:ind", "w:contextualSpacing", "w:mirrorIndents",
    "w:suppressOverlap", "w:jc", "w:textDirection", "w:textAlignment",
    "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr",
    "w:sectPr", "w:pPrChange",
)


def snap_to_grid_off(par, report, label):
    """不挂在文档网格上（「行与行字对齐」关掉），否则字距会被网格拉散。"""
    pPr = par._p.find(qn("w:pPr"))
    el = None if pPr is None else pPr.find(qn("w:snapToGrid"))
    if el is not None and (el.get(qn("w:val")) or "1") in ("0", "false", "off"):
        return

    def apply():
        p = par._p.get_or_add_pPr()
        node = p.find(qn("w:snapToGrid"))
        if node is None:
            node = OxmlElement("w:snapToGrid")
            p.insert_element_before(node, *SNAP_SUCCESSORS)
        node.set(qn("w:val"), "0")

    report.want("par.snapToGrid", label, "跟随文档网格", "不挂网格（snapToGrid=0）", apply)


def image_inline(par, report, label):
    """浮动图片（wp:anchor）改成嵌入型（wp:inline），否则整段居中不生效、图会跑版。"""
    WPD = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    move = []
    for dr in par._p.iter(qn("w:drawing")):
        anchor = dr.find("{%s}anchor" % WPD)
        if anchor is not None:
            move.append(anchor)
    if not move:
        return

    def apply():
        for a in move:
            inline = OxmlElement("wp:inline")
            for k in ("distT", "distB", "distL", "distR"):
                inline.set(k, "0")
            for el in list(a):
                if el.tag in (qn("wp:extent"), qn("wp:effectExtent"), qn("wp:docPr"),
                              qn("wp:cNvGraphicFramePr"), qn("a:graphic")):
                    inline.append(el)          # 直接搬节点，anchor 随后被替换
            a.getparent().replace(a, inline)

    report.want("image.inline", label, "浮动（wp:anchor）", "嵌入型（wp:inline）", apply)


def auto_space_off(par, report, label):
    """关闭 Word 的“自动调整中西文间距”（w:autoSpaceDE / w:autoSpaceDN）。

    Word 默认会在中英文之间自动塞约 1/4 空格：文本里没有空格字符，
    渲染出来依然像「使用 AI 工具」。要做到真正的“英文前后不加空格”，
    必须把这两个开关置 0。
    """
    pPr = par._p.get_or_add_pPr()
    for tag in ("w:autoSpaceDE", "w:autoSpaceDN"):
        el = pPr.find(qn(tag))
        if el is not None and (el.get(qn("w:val")) or "").lower() in ("0", "false", "off"):
            continue

        def apply(t=tag, e=el, p=pPr):
            node = e
            if node is None:
                node = OxmlElement(t)
                p.insert_element_before(node, *AUTOSPACE_SUCCESSORS)
            node.set(qn("w:val"), "0")

        report.want("par.autoSpace", label,
                    "开启（Word 自动在中英文之间加间距）", "关闭（autoSpaceDE/DN=0）", apply)


def _normalize_cjk_latin(par, report, label, cfg):
    """中文文段里英文词前后不留空格（「使用AI工具」）；纯英文语句不动。"""
    if not cfg["normalize"].get("cjkLatinSpace", True):
        return
    runs = par.runs
    if not runs:
        return
    joined = "".join(r.text for r in runs)
    gaps = cjk_latin_gaps(joined)
    if gaps:
        report.want("text.cjkLatinSpace", label,
                    "中文与英文之间夹着空白（共 %d 处）" % len(gaps),
                    "中文内英文前后不加空格（如「使用AI工具」）",
                    (lambda rs=runs, gs=gaps: drop_char_ranges(rs, gs)))


RPR_SPACING_SUCCESSORS = (
    "w:w", "w:kern", "w:position", "w:sz", "w:szCs", "w:highlight", "w:u",
    "w:effect", "w:bdr", "w:shd", "w:fitText", "w:vertAlign", "w:rtl", "w:cs",
    "w:em", "w:lang", "w:eastAsianLayout", "w:specVanish", "w:oMath",
)


def _write_run_char_spacing(run, val):
    rPr = run._r.get_or_add_rPr()
    sp = rPr.find(qn("w:spacing"))
    if sp is None:
        sp = OxmlElement("w:spacing")
        rPr.insert_element_before(sp, *RPR_SPACING_SUCCESSORS)
    sp.set(qn("w:val"), val)


def char_spacing_std(par, report, label):
    """字符间距「标准」：run 上 w:spacing w:val="0"（Word 默认是 0，但被改过的要拉回）。"""
    for run in par.runs:
        rPr = run._r.find(qn("w:rPr"))
        sp = None if rPr is None else rPr.find(qn("w:spacing"))
        if sp is None:
            continue                      # 没写就是标准值 0，不必到处塞节点
        if (sp.get(qn("w:val")) or "0") == "0":
            continue
        report.want("run.charSpacing", label, "非标准（%s）" % sp.get(qn("w:val")),
                    "标准（w:spacing=0）",
                    (lambda r=run: _write_run_char_spacing(r, "0")))


def _replace_par_text(par, text):
    """整段换文字：写在第一个 run 上，其余清空（保留第一个 run 的格式）。"""
    runs = par.runs
    if not runs:
        return
    runs[0].text = text
    for r in runs[1:]:
        r.text = ""


def _text_rules(par, report, label, cfg):
    """与版式无关的四条文本规则：时间冒号、小数点半角、附件标记冒号、日期格式。"""
    n = cfg["normalize"]
    runs = par.runs
    if not runs:
        return
    if n.get("timeColon", True):
        for run in runs:
            new = RE_TIME_FULL.sub(
                lambda m: "%s:%s%s" % (m.group(1), m.group(2),
                                       (":" + m.group(3)) if m.group(3) else ""),
                run.text)
            if new != run.text:
                report.want("text.timeColon", label, "时间用了全角冒号", "时分之间用半角冒号",
                            (lambda r=run, v=new: setattr(r, "text", v)))
    if n.get("numberHalfDot", True):
        for run in runs:
            txt = run.text
            # 附件序号的点必须保持全角：附件：1．2026年度… / 2．2026年度…
            # 这里的「．」后面正好跟四位年份，会被当成小数点误改成半角。
            prefix = ""
            m = re.match(r"^(附件\s*[：:]\s*\d{1,2})．", txt) or \
                re.match(r"^(\d{1,2})．(?=\d{4})", txt)
            if m:
                prefix, txt = m.group(1) + "．", txt[m.end():]
            new = RE_NUMBER_FULL_DOT.sub(".", txt)
            if new != txt:
                report.want("text.numberHalfDot", label, "小数点是全角．", "小数点用半角.",
                            (lambda r=run, v=prefix + new: setattr(r, "text", v)))
    if n.get("dateFormat", True):
        # 只规范化“整段就是一个日期”的段落（落款、日期行），正文里顺手写的时间不动
        whole = par.text.strip()
        m = RE_DATE_LOOSE.fullmatch(whole) if whole else None
        if m and runs:
            new = "%s年%d月%d日" % (m.group(1), int(m.group(2)), int(m.group(3)))
            if new != whole:
                report.want("text.dateFormat", label, whole, new,
                            (lambda: _replace_par_text(par, new)))
        elif runs:
            # 正文里顺手写的点分日期：2024.10.27 → 2024年10月27日（四位年打头，不会误伤 3.5 这类小数）
            for run in runs:
                new = RE_DATE_INLINE.sub(
                    lambda x: "%s年%d月%d日" % (x.group(1), int(x.group(2)), int(x.group(3))),
                    run.text)
                if new != run.text:
                    report.want("text.dateFormat", label, "日期用了点/斜杠分隔", "2024年10月27日",
                                (lambda r=run, v=new: setattr(r, "text", v)))
    if n.get("attachTagColon", True):
        if RE_ATTACH_TAG_COLON.match(par.text.strip()) and runs:
            last = runs[-1]
            new = re.sub(r"\s*[：:]\s*$", "", last.text)
            if new != last.text:
                report.want("text.attachTagColon", label, "附件标记后面带冒号", "去掉冒号",
                            (lambda r=last, v=new: setattr(r, "text", v)))


def normalize_text(par, report, label, cfg, light=False):
    p = par
    text = p.text
    if not text:
        return
    n = cfg["normalize"]
    if light:
        # 轻量模式：只做与版式无关的几件事（引号、时间/小数/日期写法、中英之间不留空格）
        _text_rules(p, report, label, cfg)
        _normalize_cjk_latin(p, report, label, cfg)
        return
    _text_rules(p, report, label, cfg)
    if n["stripLeadingSpace"] and RE_LEAD_SPACE.match(text):
        stripped = text.lstrip("\u3000 ")
        if stripped and p.runs:
            first = p.runs[0]
            report.want("text.leadingSpace", label, "存在行首空格", "无",
                        lambda: setattr(first, "text", first.text.lstrip("\u3000 ")))
    if n["fullwidthNumberDot"]:
        m = re.match(r"^(\s*)(\d{1,2})\.(?!\d)", p.text)
        if m and p.runs:
            runs = p.runs
            for run in runs:
                if "." in run.text and re.match(r"^\s*\d{1,2}\.\D*", run.text):
                    new = re.sub(r"^(\s*\d{1,2})\.(?!\d)\s*", "\\1．", run.text)
                    if new != run.text:
                        report.want("text.numberDot", label, "半角序号点", "全角．",
                                    (lambda r=run, v=new: setattr(r, "text", v)))
                    break
        # 「2．2026年度…」这类附件条目：点后紧跟四位年份，是序号不是小数，点也全角
        s2 = p.text.strip()
        if re.match(r"^\d{1,2}\s*\.\s*(?=\d{4})", s2):
            new2 = re.sub(r"^(\d{1,2})\s*\.\s*(?=\d{4})", "\\1．", s2)
            if new2 != p.text:
                report.want("text.numberDot", label, s2[:10], new2[:10],
                            (lambda: _replace_par_text(p, new2)))
        # 「附件：1. xxx」里的序号点也要全角，且不留空格（序号常被拆在别的 run 里，按整段处理）
        s = p.text.strip()
        if re.match(r"^附件\s*[：:]\s*\d{1,2}\s*[.．、]", s):
            new = re.sub(r"^(附件\s*[：:]\s*)(\d{1,2})\s*[.．、]\s*", "\\1\\2．", s)
            if new != p.text:
                report.want("text.attachNumberDot", label, s[:12], new[:12],
                            (lambda: _replace_par_text(p, new)))
    if n.get("cjkLatinSpace", True):
        _normalize_cjk_latin(p, report, label, cfg)


# ---------------------------------------------------------------- 表格
BORDER_ORDER = ["top", "left", "bottom", "right", "insideH", "insideV"]


def set_borders(tbl, report, label, cfg):
    """外边框 1.5 磅、内边框 0.5 磅（sz 单位为 1/8 磅）。"""
    tblPr = tbl._element.tblPr
    want_sz = {}
    for edge in BORDER_ORDER:
        kind = "outer" if edge in ("top", "left", "bottom", "right") else "inner"
        want_sz[edge] = str(int(round(cfg["tableBordersPt"][kind] * 8)))

    cur_el = tblPr.find(qn("w:tblBorders"))
    diffs = []
    for edge in BORDER_ORDER:
        cur = None if cur_el is None else cur_el.find(qn("w:" + edge))
        val = None if cur is None else cur.get(qn("w:val"))
        sz = None if cur is None else cur.get(qn("w:sz"))
        if val in (None, "nil", "none") or sz != want_sz[edge]:
            diffs.append((edge, "%s/%s" % (val, sz), "single/%s" % want_sz[edge]))
    def rebuild():
        old = tblPr.find(qn("w:tblBorders"))
        if old is not None:
            tblPr.remove(old)
        el = OxmlElement("w:tblBorders")
        for edge in BORDER_ORDER:
            b = OxmlElement("w:" + edge)
            b.set(qn("w:val"), "single")
            b.set(qn("w:sz"), want_sz[edge])
            b.set(qn("w:space"), "0")
            b.set(qn("w:color"), "auto")
            el.append(b)
        tblPr.insert_element_before(el, "w:shd", "w:tblLayout", "w:tblCellMar",
                                    "w:tblLook", "w:tblCaption", "w:tblDescription")
    for edge, cur_s, want_s in diffs:
        report.want("table.border." + edge, label, cur_s, want_s, rebuild)

    # 单元格级 tcBorders 会盖住 tblBorders（表头下边、首列右边就变成 1.5 磅）
    if _count_cell_borders(tbl):
        report.want("table.cellBorders", label,
                    "单元格自带边框（覆盖表的边框粗细）", "删掉，只留 tblBorders",
                    (lambda: _strip_cell_borders(tbl)))
    cs = tblPr.find(qn("w:tblCellSpacing"))
    if cs is not None:
        report.want("table.cellSpacing", label, cs.get(qn("w:w")), "无",
                    (lambda e=cs: tblPr.remove(e)))


def _count_cell_borders(tbl):
    n = 0
    for row in tbl.rows:
        for cell in row.cells:
            tcPr = cell._tc.tcPr
            if tcPr is None or _cell_spans(tcPr):
                continue
            if tcPr.find(qn("w:tcBorders")) is not None:
                n += 1
    return n


def _strip_cell_borders(tbl):
    n = 0
    for row in tbl.rows:
        for cell in row.cells:
            tcPr = cell._tc.tcPr
            if tcPr is None or _cell_spans(tcPr):
                continue
            el = tcPr.find(qn("w:tcBorders"))
            if el is not None:
                tcPr.remove(el)
                n += 1
    return n


def _cell_spans(tcPr):
    return (tcPr.find(qn("w:gridSpan")) is not None
            or tcPr.find(qn("w:vMerge")) is not None)


def clear_floating(tbl, report, label):
    """表格不浮动、不叠放：删 tblpPr / tblOverlap。"""
    tblPr = tbl._element.tblPr
    for tag in ("w:tblpPr", "w:tblOverlap"):
        el = tblPr.find(qn(tag))
        if el is not None:
            report.want("table.floating", label, "浮动/叠放", "无文字环绕",
                        (lambda p=tblPr, e=el: p.remove(e)))


def layout_table_width(tbl, report, label, cfg, usable_cm=None):
    """固定列宽、整表宽度 = 版心 − 两侧各 0.19cm，按内容权重分配，单列下限 1.2cm。"""
    if usable_cm is None:
        usable_cm = 21.0 - cfg["marginsCm"]["left"] - cfg["marginsCm"]["right"]
    want_cm = max(3.0, usable_cm - 2 * 0.19)
    tblPr = tbl._element.tblPr

    lay = tblPr.find(qn("w:tblLayout"))
    if lay is None or lay.get(qn("w:type")) != "fixed":
        report.want("table.layout", label,
                    None if lay is None else lay.get(qn("w:type")), "fixed",
                    (lambda: _set_or_make(tblPr, "w:tblLayout", {"w:type": "fixed"})))

    want_tw = str(int(round(Cm(want_cm).twips)))
    w = tblPr.find(qn("w:tblW"))
    if w is None or w.get(qn("w:type")) != "dxa" or (w.get(qn("w:w")) or "") != want_tw:
        report.want("table.width", label,
                    "auto" if w is None else "%s/%s" % (w.get(qn("w:type")), w.get(qn("w:w"))),
                    "dxa/%s（%.2fcm）" % (want_tw, want_cm),
                    (lambda: _set_or_make(tblPr, "w:tblW", {"w:type": "dxa", "w:w": want_tw})))

    mar = tblPr.find(qn("w:tblCellMar"))
    if mar is not None:
        report.want("table.cellMar", label, "单元格内边距偏大", "左右 0.19cm",
                    (lambda e=mar: tblPr.remove(e)))

    # 列宽：按每列最长文本加权
    ncol = 0
    for row in tbl.rows:
        ncol = max(ncol, len(row.cells))
    if not ncol:
        return
    def _vis(txt):
        # 视觉宽度：汉字算 1，数字/字母/半角符号算 0.55（TNR 半角窄），
        # 先把两字词补的空格去掉，否则列宽会来回震荡。
        # 注意：落款算空格数用的是另一个 text_width()（按 ord < 0x2000 折算 0.5，
        # 语义是「半角/全角」而不是「数字更窄」）。两者阈值不同，别合并。
        s = txt.replace(" ", "").replace("\u3000", "").strip()
        return sum(1.0 if ord(ch) > 0x2E7F else 0.55 for ch in s)

    weights = []
    for c in range(ncol):
        widest = 1.0
        for row in tbl.rows:
            cells = row.cells
            if c < len(cells):
                widest = max(widest, _vis(cells[c].text))
        weights.append(min(widest, 14.0))
    total = float(sum(weights))
    floor_cm = 2.0
    widths = [max(floor_cm, want_cm * wt / total) for wt in weights]
    scale = want_cm / sum(widths)
    widths = [x * scale for x in widths]

    # 一律按整数 twips 落地：各自取整后把余数补到最后一列，
    # 保证「列宽合计正好等于整表宽」，不会因为各自四舍五入差 0.01cm 而永远报一条偏差。
    twips = [int(round(Cm(x).twips)) for x in widths]
    twips[-1] += int(round(Cm(want_cm).twips)) - sum(twips)

    cols = tbl.columns
    cur = [None if cols[i].width is None else int(round(cols[i].width.twips))
           for i in range(min(ncol, len(cols)))]
    # 回读时 Twips→EMU→twips 会差几个 twip（约 0.02cm），按容差判定，避免永远报一条偏差
    def _near(a, b):
        return a is not None and abs(a - b) <= 20

    if not (len(cur) == len(twips) and all(_near(a, b) for a, b in zip(cur, twips))):
        def apply_widths():
            for i, tw in enumerate(twips):
                tbl.columns[i].width = Twips(tw)
                for row in tbl.rows:
                    if i < len(row.cells):
                        row.cells[i].width = Twips(tw)
        report.want("table.columnWidths", label, str(cur), str(twips), apply_widths)


def style_table(tbl, report, label, cfg, usable_cm=None):
    report.stats["tables"] += 1
    tblPr = tbl._element.tblPr
    jc = tblPr.find(qn("w:jc"))
    if jc is None or jc.get(qn("w:val")) != "center":
        report.want("table.alignment", label, None if jc is None else jc.get(qn("w:val")), "center",
                    lambda: (_set_or_make(tblPr, "w:jc", {"w:val": "center"})))
    clear_floating(tbl, report, label)
    layout_table_width(tbl, report, label, cfg, usable_cm)
    set_borders(tbl, report, label, cfg)

    # 表格本体缩进清零，避免 tblInd 把整表推偏导致“看起来没居中”
    tblInd = tblPr.find(qn("w:tblInd"))
    if tblInd is not None and (tblInd.get(qn("w:w")) or "0") != "0":
        report.want("table.indent", label, tblInd.get(qn("w:w")), 0,
                    lambda: tblPr.remove(tblInd))

    fonts, sizes = cfg["fonts"], cfg["sizes"]
    for r_i, row in enumerate(tbl.rows):
        trPr = row._tr.get_or_add_trPr()
        h = trPr.find(qn("w:trHeight"))
        want_twips = str(int(round(Cm(cfg["rowHeightCm"]).twips)))
        if h is None:
            report.want("table.rowHeight", "%s.r%d" % (label, r_i), "无", cfg["rowHeightCm"],
                        lambda r=trPr, t=want_twips: _set_or_make(r, "w:trHeight",
                                                                  {"w:val": t, "w:hRule": "atLeast"}))
        else:
            if h.get(qn("w:hRule")) != "atLeast":
                report.want("table.rowHeight.rule", "%s.r%d" % (label, r_i),
                            h.get(qn("w:hRule")), "atLeast",
                            lambda e=h: e.set(qn("w:hRule"), "atLeast"))
            if (h.get(qn("w:val")) or "0") != want_twips:
                report.want("table.rowHeight.value", "%s.r%d" % (label, r_i),
                            h.get(qn("w:val")), cfg["rowHeightCm"],
                            lambda e=h, t=want_twips: e.set(qn("w:val"), t))
        if r_i == 0:
            if trPr.find(qn("w:tblHeader")) is None:
                report.want("table.headerRepeat", "%s.r0" % label, "未重复标题行", "重复",
                            lambda r=trPr: _set_or_make(r, "w:tblHeader", {}))
        head = r_i == 0
        for c_i, cell in enumerate(row.cells):
            tcPr = cell._tc.get_or_add_tcPr()
            va = tcPr.find(qn("w:vAlign"))
            if va is None or va.get(qn("w:val")) != "center":
                report.want("table.vAlign", "%s.r%d.c%d" % (label, r_i, c_i), None, "center",
                            lambda p=tcPr: _set_or_make(p, "w:vAlign", {"w:val": "center"}))
            font = fonts["tableHead"] if head else fonts["tableBody"]
            for p_i, par in enumerate(cell.paragraphs):
                where = "%s.r%d.c%d.p%d" % (label, r_i, c_i, p_i)
                # 单元格也要处理中英混排（「AI 平台」→「AI平台」）
                normalize_text(par, report, where, cfg)
                auto_space_off(par, report, where)
                pf = par.paragraph_format
                if pf.line_spacing_rule != WD_LINE_SPACING.SINGLE:
                    report.want("table.lineRule", where, ename(pf.line_spacing_rule), "SINGLE",
                                lambda x=pf: setattr(x, "line_spacing_rule", WD_LINE_SPACING.SINGLE))
                if pt_of(pf.line_spacing) not in (None, 1.0):
                    report.want("table.lineSpacing", where, pt_of(pf.line_spacing), "单倍",
                                lambda x=pf: setattr(x, "line_spacing", 1.0))
                if pf.first_line_indent not in (None, 0):
                    report.want("table.firstLineIndent", where, pt_of(pf.first_line_indent), "无",
                                lambda x=pf: setattr(x, "first_line_indent", None))
                for attr in ("firstLine", "left", "right"):
                    cur_chars = _get_chars(par, "w:" + attr + "Chars")
                    if cur_chars is not None:
                        report.want("table.%sChars" % attr, where, "%.2f 字符" % cur_chars, "无",
                                    (lambda p=par, a=attr: set_chars_indent(p, a, None, Pt(0))))
                _zero_spacing(par, pf, report, where, cfg)
                # 单元格内容（含表头与正文）一律水平居中，并清掉左右缩进
                if pf.alignment != CENTER:
                    report.want("table.cellAlign", where, ename(pf.alignment), "CENTER",
                                lambda x=pf: setattr(x, "alignment", CENTER))
                if (pt_of(pf.left_indent) or 0) != 0:
                    report.want("table.cellLeftIndent", where, pt_of(pf.left_indent), 0,
                                lambda x=pf: setattr(x, "left_indent", None))
                if (pt_of(pf.right_indent) or 0) != 0:
                    report.want("table.cellRightIndent", where, pt_of(pf.right_indent), 0,
                                lambda x=pf: setattr(x, "right_indent", None))
                if cfg["normalize"]["twoCharCellSpace"] and par.text and re.fullmatch(r"[\u4e00-\u9fff]{2}", par.text):
                    _spread_two_chars(par, report, where, font, sizes["table"], cfg)
                if par.runs:
                    style_runs(par, font, sizes["table"], fonts["latin"], report, where, cfg)


def _spread_two_chars(par, report, where, font, size, cfg):
    """表格内两个字的词语中间加两个半格空格。"""
    target = par.text[0] + "  " + par.text[1]
    nonempty = [r for r in par.runs if r.text]
    keep_el = nonempty[0]._element if nonempty else None

    def apply():
        if keep_el is None:
            run = par.add_run()
            set_run(run, font, size, cfg["fonts"]["latin"], _NullReport(), where, cfg)
            run.text = target
            return
        for r in par.runs:
            r.text = target if r._element is keep_el else ""

    report.want("text.twoCharCell", where, par.text, target, apply)


class _NullReport:
    def want(self, *a, **k):
        return False


def _set_or_make(parent, tag, attrs):
    el = parent.find(qn(tag))
    if el is None:
        el = OxmlElement(tag)
        parent.append(el)
    for k, v in attrs.items():
        el.set(qn(k), v)
    return el


# ---------------------------------------------------------------- 页脚页码
def set_footer(section, report, cfg, fix=True):
    footer = section.footer
    if not fix:
        has_field = "PAGE" in footer._element.xml
        par0 = footer.paragraphs[0] if footer.paragraphs else None
        align_ok = par0 is not None and par0.paragraph_format.alignment == CENTER
        if not (has_field and align_ok):
            report.want("footer.pageNumber", "footer",
                        "有" if has_field else "无页码域", "底端居中 — 1 —")
        return
    if footer.is_linked_to_previous:
        footer.is_linked_to_previous = False
    par = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    has_field = "PAGE" in par._element.xml or True
    for child in list(par._p):
        if child.tag != qn("w:pPr"):
            par._p.remove(child)
    pf = par.paragraph_format
    pf.alignment = CENTER
    if pf.first_line_indent not in (None, 0):
        report.want("footer.firstLineIndent", "footer", pt_of(pf.first_line_indent), "无",
                    lambda: setattr(pf, "first_line_indent", None))

    def mk_run(text=None):
        r = par.add_run()
        rPr = r._element.get_or_add_rPr()
        rFonts = OxmlElement("w:rFonts")
        rFonts.set(qn("w:ascii"), cfg["fonts"]["latin"])
        rFonts.set(qn("w:hAnsi"), cfg["fonts"]["latin"])
        rFonts.set(qn("w:eastAsia"), cfg["fonts"]["footer"])
        rPr.append(rFonts)
        r.font.size = Pt(cfg["sizes"]["footer"])
        if text is not None:
            r.text = text
        return r

    def add_field_run(fld_type=None, instr=None):
        r = mk_run()
        if instr is not None:
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = instr
        else:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), fld_type)
        r._element.append(el)
        return r

    mk_run("— ")
    add_field_run("begin")
    add_field_run(instr=" PAGE ")
    add_field_run("separate")
    mk_run("1")
    add_field_run("end")
    mk_run(" —")
    report.stats["footer"] = report.stats.get("footer", 0) + 1


# ---------------------------------------------------------------- 主流程
def _looks_like_attach_wrap(par, cfg):
    """判断是否是附件条目被拆出来的续行：首行缩进正好等于 attachWrapIndentChars。"""
    want = float(cfg["attachWrapIndentChars"])
    cur = _get_chars(par, "w:firstLineChars")
    return cur is not None and abs(cur - want) < 0.05


def _gap_all_blank(items, a, b):
    """items[a] 与 items[b] 之间（不含两端）是否全是空段落。"""
    for k in range(a + 1, b):
        t, obj = items[k]
        if t != "p" or (obj.text or "").strip():
            return False
    return True


def classify(items, cfg):
    """返回 [(kind, obj, text)]，kinds: title/table_caption/h1/h2/h3/attach_tag/attach_line/
    sign_org/sign_date/image/body/blank/tbl"""
    out = []
    paras = [(i, obj) for i, (t, obj) in enumerate(items) if t == "p"]
    nonempty = [(i, p) for i, p in paras if p.text.strip()]
    title_map = {}
    if nonempty:
        i0 = nonempty[0][0]
        title_map[i0] = "title"
        # 副标题：紧跟主标题（中间最多隔 gapTolerance 个空段）的另一段短文本
        sub = cfg.get("subtitle") or {}
        if sub.get("enabled", True):
            tol = max(0, int(sub.get("gapTolerance", 1)))
            main_max = int(sub.get("mainMaxChars", 18))
            sub_max = int(sub.get("maxChars", 40))
            prev_len = len(nonempty[0][1].text.strip())
            for i, p in nonempty[1:3]:
                txt = p.text.strip()
                gap_ok = (i - i0 - 1) <= tol and _gap_all_blank(items, i0, i)
                # 破折号开头基本就是副标题；否则要求主标题够短，避免把正文首段吃进来
                dash_lead = bool(RE_DASH_LEAD.match(txt))
                if (gap_ok and len(txt) <= sub_max
                        and (dash_lead or prev_len <= main_max)
                        and not txt.endswith(("。", "：", "，", "；", "！", "？", "、", ":"))
                        and not RE_H1.match(txt) and not RE_H2.match(txt)
                        and not RE_ATTACH_LINE.match(txt) and not RE_ATTACH_TAG.match(txt)
                        and not RE_CAPTION_LEAD.match(txt)      # 「表 1 xxx」是表标题，不是副标题
                        and not RE_DATE.search(txt)):
                    title_map[i] = "subtitle"
                    prev_len = len(txt)
                else:
                    break

    # 落款：整段就是一个日期，且它后面紧跟「附件N」标记或文档结束。
    # 不能只看“末尾 6 段”——带附件正文的公文，落款在文档中间。
    sign = {}
    ne = [(i, p) for i, p in paras if p.text.strip()]
    cands = []
    for pos, (i, p) in enumerate(ne):
        txt = p.text.strip()
        if not (RE_DATE.fullmatch(txt) or RE_DATE_LOOSE.fullmatch(txt)):
            continue
        if text_width(txt) > 26:
            continue
        nxt = ne[pos + 1][1].text.strip() if pos + 1 < len(ne) else ""
        if nxt and not RE_ATTACH_TAG.match(nxt):
            continue        # 后面还有正文 → 不是落款日期
        cands.append((pos, i, p))
    if cands:
        pos, i, p = cands[-1]
        sign[i] = "sign_date"
        if pos > 0:
            prev_i, prev_p = ne[pos - 1]
            if (text_width(prev_p.text.strip()) <= 24 and not RE_DATE.search(prev_p.text)
                    and not prev_p.text.strip().endswith(("。", "．", "，", "；"))):
                sign[prev_i] = "sign_org"

    for i, (t, obj) in enumerate(items):
        if t == "tbl":
            out.append(("tbl", obj, ""))
            continue
        p = obj
        txt = p.text.strip()
        if i in title_map:
            out.append((title_map[i], p, txt))
        elif i in sign:
            out.append((sign[i], p, txt))
        elif not txt:
            has_img = bool(p._element.findall(".//" + qn("w:drawing")))
            out.append(("image" if has_img else "blank", p, ""))
        elif RE_ATTACH_TAG.match(p.text or txt):
            out.append(("attach_tag", p, txt))
        elif RE_ATTACH_LINE.match(p.text or txt):
            out.append(("attach_line", p, txt))
        elif RE_H1.match(p.text):
            out.append(("h1", p, txt))
        elif RE_H2.match(p.text):
            out.append(("h2", p, txt))
        elif RE_H3.match(p.text):
            out.append(("h3", p, txt))
        elif RE_FIG_CAPTION.match(p.text or txt):
            out.append(("fig_caption", p, txt))
        elif p._element.findall(".//" + qn("w:drawing")):
            out.append(("image", p, txt))
        else:
            out.append(("body", p, txt))
    # 主送机关：标题（含副标题）之后第一个「……：」短段，顶格不缩进
    seen_title = False
    for idx, (t, obj, txt) in enumerate(out):
        if t in ("title", "title_first", "title_mid", "title_last", "title_only", "subtitle"):
            seen_title = True
            continue
        if not seen_title:
            continue
        if t in BLANK_KINDS:
            continue
        if t == "body" and txt and RE_ADDRESSEE.match(txt) and text_width(txt) <= 30:
            out[idx] = ("addressee", obj, txt)
        break
    # 附件续行：紧接“附件：XXX”之后的附件条目，需与冒号后首字对齐
    in_attach = False
    for idx, (t, obj, txt) in enumerate(out):
        if t == "tbl":
            in_attach = False
            continue
        kind = t
        if kind == "attach_line":
            in_attach = True
            continue
        if not in_attach:
            continue
        if kind in BLANK_KINDS:
            continue        # 空行不打断附件条目序列
        if kind in ("sign_org", "sign_date", "attach_tag", "title"):
            in_attach = False
        elif RE_ATTACH_ITEM.match(txt):
            out[idx] = ("attach_cont", obj, txt)
        elif _looks_like_attach_wrap(obj, cfg):
            # 我们自己拆出来的续行（首行缩进 6 字符）→ 保持附件归属
            out[idx] = ("attach_wrap", obj, txt)
        else:
            in_attach = False
    # 文末「附件1」下面那一行是附件标题：按大标题排（方正小标宋二号、居中）
    for idx, (t, obj, txt) in enumerate(out):
        if t != "attach_tag":
            continue
        j = idx + 1
        while j < len(out) and out[j][0] in BLANK_KINDS:
            j += 1
        if j < len(out) and out[j][0] == "body":
            out[j] = ("attach_title", out[j][1], out[j][2])
    # 表标题：紧邻表格之前的段落
    for idx, (t, obj, txt) in enumerate(out):
        if t != "tbl":
            continue
        j = idx - 1
        while j >= 0 and out[j][0] in BLANK_KINDS:
            j -= 1
        if j >= 0:
            kind, prev, ptxt = out[j]
            if kind == "body" and ptxt and text_width(ptxt) <= 40 and not ptxt.endswith("。"):
                out[j] = ("table_caption", prev, ptxt)
    # 表标题与表格之间那一段：专用“0 行高、5 号字”空段，不能按普通空行排 29 磅
    for idx, (t, obj, txt) in enumerate(out):
        if t != "tbl":
            continue
        j = idx - 1
        while j >= 0 and out[j][0] in BLANK_KINDS + ("caption_spacer",):
            j -= 1
        if j >= 0 and out[j][0] == "table_caption":
            for k in range(j + 1, idx):
                out[k] = ("caption_spacer", out[k][1], "")
    # 空行也带格式：标题上空、多行标题之间、标题下固定那 1 行 → 沿用标题格式；
    # 其余空行沿用正文格式——和“手工回车”继承出来的结果一致。
    title_positions = [i for i, (t, _o, _x) in enumerate(out) if t in ("title", "subtitle")]
    if title_positions:
        first_t, last_t = title_positions[0], title_positions[-1]
        title_side = set(range(0, first_t)) | set(range(first_t + 1, last_t))
        j = last_t + 1
        room = int(cfg["titleBlankAfterLines"])
        while j < len(out) and out[j][0] == "blank" and room > 0:
            title_side.add(j)
            room -= 1
            j += 1
        for idx in title_side:
            if out[idx][0] == "blank":
                out[idx] = ("blank_title", out[idx][1], "")
    # 主送机关上面的空行跟着顶格（标题侧那几行已经标记为 blank_title，不再改）
    for idx, (t, _o, _x) in enumerate(out):
        if t != "addressee":
            continue
        k = idx - 1
        while k >= 0 and out[k][0] == "blank":
            out[k] = ("blank_add", out[k][1], out[k][2])
            k -= 1
    return out


def _spacer_size_val(cfg):
    return str(int(round(cfg["sizes"]["spacer"] * 2)))


def _spacer_props_ok(p_el, cfg):
    pPr = p_el.find(qn("w:pPr"))
    if pPr is None:
        return False
    sp = _effective_spacing(pPr)
    if sp is None:
        return False
    if sp.get(qn("w:line")) != "0" or sp.get(qn("w:lineRule")) != "atLeast":
        return False
    if (sp.get(qn("w:before")) or "") != "0" or (sp.get(qn("w:after")) or "") != "0":
        return False
    rPr = pPr.find(qn("w:rPr"))
    sz = None if rPr is None else rPr.find(qn("w:sz"))
    return sz is not None and sz.get(qn("w:val")) == _spacer_size_val(cfg)


def _write_spacer_props(p_el, cfg):
    """行高最小值 0 磅、字号 5 号、段前段后 0 —— 表标题与表格之间那一行。"""
    pPr = p_el.get_or_add_pPr()
    sp = _sole_spacing(p_el)
    for key in ("w:beforeLines", "w:afterLines",
                "w:beforeAutospacing", "w:afterAutospacing"):
        sp.attrib.pop(qn(key), None)
    sp.set(qn("w:line"), "0")
    sp.set(qn("w:lineRule"), "atLeast")
    sp.set(qn("w:before"), "0")
    sp.set(qn("w:after"), "0")
    rPr = pPr.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        pPr.append(rPr)
    val = _spacer_size_val(cfg)
    for tag in ("w:sz", "w:szCs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = OxmlElement(tag)
            rPr.append(el)
        el.set(qn("w:val"), val)


def format_table_caption_spacer(items, kinds, report, cfg):
    """表标题与表格之间插入一行（行高最小值 0，字号 5 号）；已有则只校正属性。"""
    n = 0
    for idx, (t, obj) in enumerate(items):
        if t != "tbl":
            continue
        j = idx - 1
        while j >= 0 and kinds[j] in BLANK_KINDS + ("caption_spacer",):
            j -= 1
        if j < 0 or kinds[j] != "table_caption":
            continue
        n += 1
        prev_p = items[j][1]
        nxt = prev_p._p.getnext()
        if nxt is not None and nxt.tag == qn("w:p") and _is_blank_el(nxt):
            if not _spacer_props_ok(nxt, cfg):
                report.want("table.captionSpacer", "table%d" % n,
                            "空行行高/字号不符", "行高最小值 0 磅、5 号字",
                            lambda e=nxt: _write_spacer_props(e, cfg))
            continue
        spacer = OxmlElement("w:p")
        _write_spacer_props(spacer, cfg)
        prev_p._p.addnext(spacer)
        report.want("table.captionSpacer", "table%d" % n,
                    "表标题与表格间无空行", "插入空行",
                    lambda: None)
        report.stats["table.spacer"] += 1


def _is_blank_el(el):
    if el.tag != qn("w:p"):
        return False
    if el.findall(".//" + qn("w:drawing")):
        return False
    return not "".join(el.itertext()).strip()


def _count_blanks_before(par):
    n = 0
    prev = par._p.getprevious()
    while prev is not None and _is_blank_el(prev):
        n += 1
        prev = prev.getprevious()
    return n


def _next_anchor_par(par):
    """其后第一个非空段落（跳过空段落；表格不作为锚点）。"""
    cur = par._p.getnext()
    while cur is not None:
        if cur.tag == qn("w:p") and not _is_blank_el(cur):
            return Paragraph(cur, par._parent)
        cur = cur.getnext()
    return None


def _anchor_after_attach(seq, idx):
    """「附件：…」之后第一个非附件条目、非空行的段落。

    附件条目（attach_cont）属于同一段附件说明，不能把它们和“附件：”拆开插空行。
    """
    for j in range(idx + 1, len(seq)):
        kind, obj, _txt = seq[j]
        if kind in BLANK_KINDS + ("attach_cont", "attach_wrap"):
            continue
        if kind == "tbl":
            return None
        return obj
    return None


def ensure_blank_lines(kinds, report, cfg, dry_run=False):
    """段前段后一律 0 磅，空行改用独立空段落（整行）表达。

    按参考要求保证：标题上空 2 行、下空 1 行；附件说明前空 1 行、后空 2 行；
    文末「附件1」标记前空 1 行；落款前空 2 行（仅 signature.enabled 时）；
    表标题前不留空行（0 行）。

    dry_run=True 时只统计偏差、不改文档——check 模式靠它把空行问题一并报出来，
    保证「check 通过 ⇒ apply 无改动」。
    """
    plan = {}

    def bump(par, want, label, blank_kind="blank", rule="doc.blankLines"):
        """want=0 表示「这里不许有空行」（表标题前），同样要登记，不能当成「无需处理」。"""
        if par is None or want is None:
            return
        key = id(par._p)
        old = plan.get(key)
        if old is None or old[1] < want:
            plan[key] = (par, want, label, blank_kind, rule)

    titles = [obj for k, obj, _t in kinds if k in ("title", "subtitle")]
    if titles:
        bump(titles[0], cfg["titleBlankBeforeLines"], "标题前", "blank_title")
        bump(_next_anchor_par(titles[-1]), cfg["titleBlankAfterLines"], "标题后", "blank_title")
    cap_no = 0
    for idx, (k, obj, _t) in enumerate(kinds):
        if k == "attach_line":
            bump(obj, cfg["attachBlankBeforeLines"], "附件说明前")
            bump(_anchor_after_attach(kinds, idx), cfg["attachBlankAfterLines"], "附件说明后")
        elif k == "attach_tag":
            bump(_next_anchor_par(obj), cfg["attachTagBlankAfterLines"], "附件标记与附件标题之间")
        elif k == "sign_org" and cfg["signature"]["enabled"]:
            bump(obj, cfg["signature"]["blankLinesBefore"], "落款前")
        elif k == "table_caption":
            # 正文与表标题之间不留空行：参考件只规定「表标题与表格之间空一行」，
            # 表格前面多敲的回车必须清掉，否则表标题会与正文拉开一整行。
            cap_no += 1
            bump(obj, 0, "表标题前（表%d）" % cap_no, "blank", "table.captionBlankBefore")

    touched = 0
    for par, want, label, blank_kind, rule in plan.values():
        cur = _count_blanks_before(par)
        if cur == want:
            continue
        if not dry_run:
            if cur < want:
                for _ in range(want - cur):
                    _style_blank(par.insert_paragraph_before(), cfg, blank_kind)
            else:
                for _ in range(cur - want):
                    prev = par._p.getprevious()
                    if prev is None or not _is_blank_el(prev):
                        break
                    prev.getparent().remove(prev)
        report.want(rule, "doc." + label,
                    "%d 个空行" % cur, "%d 个空行" % want)
        touched += 1
    return touched


def _collect_body(doc):
    body = doc.element.body
    items = []
    for child in list(body.iterchildren()):
        if child.tag == qn("w:p"):
            items.append(("p", Paragraph(child, doc)))
        elif child.tag == qn("w:tbl"):
            items.append(("tbl", Table(child, doc)))
    return items


# --------------------------------------------------------------- 附件自动分段
# “附件：1．……”这类条目往往长到要折行。公文做法是把它拆成真正的多个段落：
#   第 1 段（“附件：1．……首行”）首行缩进 2 字符；
#   每个折行出来的续行段落首行缩进 6 字符（“附件：”3 字 + “1．”2 字 + 半字），
#   后续条目“2．……”首行缩进 5 字符。左/右/悬挂缩进一律 0，段前段后 0。
#
# 折行位置不靠字宽估算：把待测段落单独存成临时文档交给 Word/WPS 排版，用
# Range.Information(wdFirstCharacterLineNumber) 读回每个字符落在第几行，
# 按实测行边界拆段，因此和最终打印效果完全一致。

_WORD_MEASURE_PS = r"""
$ErrorActionPreference = 'Stop'
$src = '@SRC@'
$out = '@OUT@'
$app = $null
$doc = $null
try {
  $app = New-Object -ComObject Word.Application
  $app.Visible = $false
  $app.DisplayAlerts = 0
  $doc = $app.Documents.Open($src, $false, $true)
  $parts = New-Object System.Collections.Generic.List[string]
  $count = $doc.Paragraphs.Count
  for ($p = 1; $p -le $count; $p++) {
    $r = $doc.Paragraphs($p).Range
    $t = $r.Text -replace "`r", "" -replace "`a", ""
    $len = $t.Length
    if ($len -eq 0) { $parts.Add('[]'); continue }
    if ([int]$r.ComputeStatistics(1) -le 1) { $parts.Add('[[' + 0 + ',' + ($len - 1) + ']]'); continue }
    $segs = New-Object System.Collections.Generic.List[string]
    $prev = -1
    $start = 0
    for ($i = 0; $i -lt $len; $i++) {
      $ln = [int]$doc.Range($r.Start + $i, $r.Start + $i + 1).Information(10)
      if ($ln -ne $prev) {
        if ($prev -ne -1) { $segs.Add('[' + $start + ',' + ($i - 1) + ']') }
        $start = $i
        $prev = $ln
      }
    }
    $segs.Add('[' + $start + ',' + ($len - 1) + ']')
    $parts.Add('[' + ($segs -join ',') + ']')
  }
  [System.IO.File]::WriteAllText($out, '[' + ($parts -join ',') + ']', [System.Text.UTF8Encoding]::new($false))
  'OK'
} catch {
  'ERR ' + $_.Exception.Message
} finally {
  if ($doc) { try { $doc.Close($false) } catch {} }
  if ($app) { try { $app.Quit() } catch {} }
}
"""


def _powershell():
    for name in ("powershell.exe", "powershell", "pwsh"):
        found = shutil.which(name)
        if found:
            return found
    return None


def _word_measure(path):
    """调用 Word/WPS 读回每段的行区间；不可用时返回 None。"""
    shell = _powershell()
    if not shell:
        return None
    fd, out_json = tempfile.mkstemp(suffix=".json")
    os.close(fd)
    script = _WORD_MEASURE_PS.replace("@SRC@", path).replace("@OUT@", out_json)
    enc = base64.b64encode(script.encode("utf-16-le")).decode("ascii")
    try:
        proc = subprocess.run([shell, "-NoProfile", "-NonInteractive",
                               "-EncodedCommand", enc],
                              capture_output=True, timeout=180)
        head = (proc.stdout or b"").decode("utf-8", "replace").strip()
        if not head.startswith("OK"):
            return None
        return json.loads(open(out_json, encoding="utf-8").read())
    except Exception:
        return None
    finally:
        try:
            os.unlink(out_json)
        except OSError:
            pass


def _sibling_pos(el):
    """元素在 body 里的位置（前面有几个元素兄弟）。lxml 代理对象不保证同一性，用 XPath 数最稳。"""
    return int(el.xpath("count(preceding-sibling::*)"))


def _measure_attach_lines(doc, entries, wrap_chars, body_pt=16.0):
    """entries: [(body 位置, Paragraph, label, is_wrap)]（已按位置升序）；返回等长行区间列表。

    两种测量几何：
      * 起始段（`附件：1．` / `2．`）：保持它自己的 firstLineChars，量出“首行写到哪换行”；
      * 续行段（已经拆出来的 wrap 段）：把缩进临时设成 `leftChars = 6 字符`、firstLine 清零，
        量成“整段 6 字符缩进”的均匀版式——拆出来的每一段都从 6 字符开始，只有这样量才准。
    """
    if not entries:
        return []
    tmpdir = tempfile.mkdtemp(prefix="docxfmt-measure-")
    src = os.path.join(tmpdir, "measure.docx")
    try:
        doc.save(src)
        probe = Document(src)
        keep = dict((pos, is_wrap) for pos, _p, _l, is_wrap in entries)
        body = probe.element.body
        cursor = 0
        for child in list(body.iterchildren()):
            if child.tag == qn("w:sectPr"):
                continue
            pos = cursor
            cursor += 1
            if pos not in keep:
                body.remove(child)
                continue
            if not keep[pos]:
                continue
            par = Paragraph(child, probe)
            ind = _get_ind(par, make=True)
            # 探测段落只用磅值左缩进：实测 `w:leftChars` 与 `w:left` 并存时 Word 会按
            # 别的方式解释（6 字符被当成 30 磅，折行点偏后 2 个字），只写磅值才准。
            ind.set(qn("w:left"), str(int(round(float(wrap_chars) * float(body_pt) * 20))))
            for name in ("w:leftChars", "w:firstLineChars", "w:firstLine",
                         "w:hangingChars", "w:hanging"):
                if ind.get(qn(name)) is not None:
                    del ind.attrib[qn(name)]
        probe.save(src)
        res = _word_measure(src)
        if res is None or len(res) != len(entries):
            return None
        return res
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def _do_split_par(par, segs, cfg):
    """按给定文本切片把一个段落拆成多个段落，续行首行缩进 attachWrapIndentChars。"""
    p_el = par._p
    parent = p_el.getparent()
    runs = list(par.runs)
    spans = []
    pos = 0
    for seg in segs:
        spans.append((pos, pos + len(seg)))
        pos += len(seg)
    els = []
    for a, b in spans:
        el = copy.deepcopy(p_el)
        for child in list(el):
            if child.tag != qn("w:pPr"):
                el.remove(child)
        cursor = 0
        for r in runs:
            rtxt = r.text
            rs = cursor
            cursor += len(rtxt)
            lo, hi = max(rs, a), min(cursor, b)
            if hi <= lo:
                continue
            new_r = copy.deepcopy(r._r)
            for t in new_r.findall(qn("w:t")):
                new_r.remove(t)
            t = OxmlElement("w:t")
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            t.text = rtxt[lo - rs:hi - rs]
            new_r.append(t)
            el.append(new_r)
        els.append(el)
    anchor = p_el
    for el in els:
        anchor.addnext(el)
        anchor = el
    parent.remove(p_el)
    wrap_chars = float(cfg["attachWrapIndentChars"])
    body_pt = float(cfg["sizes"]["body"])
    for el in els[1:]:
        sub = Paragraph(el, par._parent)
        set_chars_indent(sub, "firstLine", wrap_chars, Pt(wrap_chars * body_pt))


def split_attach_paragraphs(doc, report, cfg):
    """把超过一行的附件条目按 Word 实测折行点拆段；返回是否有改动。"""
    if not cfg["attachSplit"]["enabled"]:
        return False
    if cfg["attachSplit"]["mode"] != "word":
        return False   # indent-only：只排 2/5 字符缩进，折行交给使用者手工处理
    kinds = classify(_collect_body(doc), cfg)
    targets = []
    for i, (kind, obj, _txt) in enumerate(kinds):
        if kind not in ("attach_line", "attach_cont", "attach_wrap"):
            continue
        style_par(obj, kind, report, "p%d" % i, cfg)   # 先落缩进，测量才准
        # 版心宽 15.6cm ≈ 27.6 个三号汉字，按 6 字符缩进算也还有 21 字；
        # 短于此的条目不可能折行，不必启动 Word（省 3~6 秒）
        if len(obj.text) < int(cfg["attachSplit"]["minCharsToMeasure"]):
            continue
        targets.append((_sibling_pos(obj._p), obj, "p%d" % i, kind == "attach_wrap"))
    targets.sort(key=lambda t: t[0])
    if not targets:
        return False
    ranges = _measure_attach_lines(doc, targets, cfg["attachWrapIndentChars"],
                                    cfg["sizes"]["body"])
    if ranges is None:
        report.warn("附件自动分段不可用：本机没有能调用的 Word/WPS 排版引擎，"
                    "已按“附件条目首行缩进 2/5 字符、左右悬挂 0”排版，折行请手工换行。")
        return False
    changed = False
    for (_pidx, par, label, is_wrap), rng in zip(targets, ranges):
        if not rng or len(rng) < 2:
            continue
        text = par.text
        if not text:
            continue
        if is_wrap:
            segs = [text[s:e + 1] for s, e in rng if 0 <= s <= e < len(text)]
        else:
            # 起始段只切第一处折行；剩下的尾巴留给下一轮按 6 字符缩进量
            cut = rng[0][1] + 1
            segs = [text[:cut], text[cut:]]
        if len(segs) < 2 or "".join(segs) != text or not all(segs):
            continue
        report.want("attach.split", label, "1 段 %d 字（自动折行）" % len(text),
                    "%d 段（续行首行缩进 %g 字符）"
                    % (len(segs), cfg["attachWrapIndentChars"]),
                    (lambda p=par, s=segs: _do_split_par(p, s, cfg)))
        changed = True
    return changed


def _save(doc, path):
    tmp = path + ".tmp.docx"
    doc.save(tmp)
    try:
        os.replace(tmp, path)
    except PermissionError:
        # 目标是只读文件（或刚被复制过来的只读副本）：去掉只读位再替换一次
        try:
            os.chmod(path, stat.S_IWRITE | stat.S_IREAD)
        except OSError:
            pass
        try:
            os.replace(tmp, path)
        except OSError:
            os.remove(tmp)
            raise


def light_pass(doc, report, cfg):
    """与版式无关的全局规范化：不论是否套用公文版式都执行。

    - 一个段落只留一个 w:spacing（重复的会把行距“写对了却显示成旧值”）
    - 「换行和分页」里 6 个开关一律不启用
    - 中文里不写「」『』，改“”‘’
    - 字符间距标准、时间冒号/小数点半角、日期统一、附件标记去冒号
    - 中文文段里的英文前后不留空格（并关闭 Word 自动中西文间距）
    """
    n = cfg["normalize"]
    pagination = n.get("paginationOff", True)
    quote = n.get("quotePair", True)
    char_spacing = n.get("charSpacing", True)
    for i, par in enumerate(_iter_all_paragraphs(doc)):
        label = "l%d" % i
        dedupe_spacing(par, report, label)
        if pagination:
            clear_pagination(par, report, label)
        if quote:
            quote_normalize(par, report, label)
        _text_rules(par, report, label, cfg)
        if char_spacing:
            char_spacing_std(par, report, label)
        snap_to_grid_off(par, report, label)
        _normalize_cjk_latin(par, report, label, cfg)
        auto_space_off(par, report, label)
    return report


def drop_blanks_between_titles(items, kinds):
    """主标题与副标题之间不留空行（空行从副标题之后开始算）。"""
    idxs = [i for i, (k, _o, _t) in enumerate(kinds) if k in ("title", "subtitle")]
    if len(idxs) < 2:
        return 0
    removed = 0
    for a, b in zip(idxs, idxs[1:]):
        for k in range(b - 1, a, -1):
            if kinds[k][0] in BLANK_KINDS:
                el = kinds[k][1]._p
                if el.getparent() is not None:
                    el.getparent().remove(el)
                    removed += 1
    return removed


# ---------------------------------------------------------------- 字体档
def apply_font_profile(cfg, profile=None):
    """按档位覆盖 cfg["fonts"]，返回实际生效的档位名。

    standard 用常见简体字体；strict 用 GBK/GB2312 全字库。
    """
    p = str(profile or cfg.get("fontProfile") or "standard").strip().lower()
    if p not in ("standard", "strict"):
        p = "standard"
    if p == "strict":
        fonts = cfg.setdefault("fonts", {})
        for k, v in (cfg.get("fontsStrict") or {}).items():
            if v:
                fonts[k] = v
    return p


def check_fonts(report, cfg):
    """检测缺哪些字体。

    **不联网**：只读注册表与字体目录。check 模式把缺失记成偏差；
    apply 模式只写 warnings（不影响排版，也不阻断写回）。
    """
    fc = cfg.get("fontCheck") or {}
    if not fc.get("enabled", True):
        return None
    try:
        here = os.path.dirname(os.path.abspath(__file__))
        if here not in sys.path:
            sys.path.insert(0, here)
        import fontcheck
    except Exception:                                      # noqa: BLE001
        return None
    try:
        info = fontcheck.check(cfg)
    except Exception:                                      # noqa: BLE001
        return None
    if not info.get("missing"):
        return info
    cr = info.get("consentRequired") or {}
    ask = cr.get("ask") or ("缺少字体：%s" % "、".join(info["missing"]))
    if report.mode == "check":
        for name in info["missing"]:
            report.want("font.missing:" + name, "字体", "未安装", "已安装")
    else:
        report.warnings.append(ask)
        if cr.get("command"):
            report.warnings.append("修复：" + cr["command"])
    report.stats["missingFonts"] = len(info["missing"])
    return info


def process(path, mode, cfg, full=True, profile=None):
    t0 = time.time()
    fix = mode == "apply"
    doc = Document(path)
    report = Report(os.path.abspath(path), mode, fix)

    profile = apply_font_profile(cfg, profile)
    fontinfo = None
    if (cfg.get("fontCheck") or {}).get("onRun", True):
        fontinfo = check_fonts(report, cfg)

    light_pass(doc, report, cfg)

    if not full:
        changed = report.total > 0
        if fix and changed:
            _save(doc, path)
        out = report.as_dict(changed)
        out["light"] = True
        out["writeBack"] = bool(fix and changed)
        out["elapsedMs"] = int((time.time() - t0) * 1000)
        out["fontProfile"] = profile
        if fontinfo:
            out["fonts"] = fontinfo
        return out

    for section in doc.sections:
        m = cfg["marginsCm"]
        for attr, key in (("top_margin", "top"), ("bottom_margin", "bottom"),
                          ("left_margin", "left"), ("right_margin", "right")):
            cur = getattr(section, attr)
            want = Cm(m[key])
            if cur is None or abs(cur.cm - m[key]) > 0.02:
                report.want("page.margin." + key, "section", None if cur is None else round(cur.cm, 2), m[key],
                            (lambda s=section, a=attr, w=want: setattr(s, a, w)))
        if cfg.get("pageSize", "A4") != "keep":
            cur_w, cur_h = section.page_width, section.page_height
            if (cur_w is None or abs(cur_w.cm - 21.0) > 0.05
                    or cur_h is None or abs(cur_h.cm - 29.7) > 0.05):
                desc = "%.2f×%.2f cm" % (cur_w.cm if cur_w else 0.0,
                                         cur_h.cm if cur_h else 0.0)
                report.want("page.size", "section", desc, "A4（21.00×29.70 cm）",
                            (lambda s=section: (setattr(s, "page_width", Cm(21.0)),
                                                setattr(s, "page_height", Cm(29.7)))))
        if cfg["pageNumber"]:
            set_footer(section, report, cfg, fix)

    usable_cm = None
    try:
        sec0 = doc.sections[0]
        if sec0.page_width is not None:
            usable_cm = sec0.page_width.cm - cfg["marginsCm"]["left"] - cfg["marginsCm"]["right"]
    except Exception:
        usable_cm = None

    body = doc.element.body
    items = _collect_body(doc)

    kinds = classify(items, cfg)
    if fix:
        if cfg["attachSplit"]["enabled"]:
            for _round in range(max(1, int(cfg["attachSplit"]["maxRounds"]))):
                if not split_attach_paragraphs(doc, report, cfg):
                    break
            items = _collect_body(doc)
            kinds = classify(items, cfg)
        removed = drop_blanks_between_titles(items, kinds)
        if removed:
            report.want("title.subtitleGap", "doc.标题",
                        "%d 个空行" % removed, "0 个空行（主标题与副标题紧邻）")
            items = _collect_body(doc)
            kinds = classify(items, cfg)
        ensure_blank_lines(kinds, report, cfg)
        # 插完空行段落列表就变了：重新收集再定位“表标题 → 表格”的相邻关系
        items = _collect_body(doc)
        kinds = classify(items, cfg)
        format_table_caption_spacer(items, [k for k, _o, _t in kinds], report, cfg)
        items = _collect_body(doc)
        kinds = classify(items, cfg)
    else:
        # check：只报空行偏差，不动文档
        ensure_blank_lines(kinds, report, cfg, dry_run=True)

    if not fix and cfg["attachSplit"]["enabled"]:
        split_attach_paragraphs(doc, report, cfg)

    title_positions = [i for i, (k, _o, _t) in enumerate(kinds) if k == "title"]
    if title_positions:
        if len(title_positions) == 1:
            kinds[title_positions[0]] = ("title_only",) + tuple(kinds[title_positions[0]][1:])
        else:
            for pos, idx in enumerate(title_positions):
                name = "title_first" if pos == 0 else ("title_last" if pos == len(title_positions) - 1 else "title_mid")
                kinds[idx] = (name,) + tuple(kinds[idx][1:])
    table_idx = 0
    for i, (kind, obj, txt) in enumerate(kinds):
        label = "p%d" % i
        if kind == "tbl":
            style_table(obj, report, "t%d" % table_idx, cfg, usable_cm)
            table_idx += 1
            continue
        if kind in BLANK_KINDS:
            style_par(obj, kind, report, label, cfg)
            continue
        if kind in ("body", "h1", "h2", "h3", "attach_line", "attach_cont",
                    "attach_wrap", "table_caption", "fig_caption", "addressee",
                    "attach_title") or kind.startswith("title"):
            normalize_text(obj, report, label, cfg)
        style_par(obj, kind, report, label, cfg)
        if kind == "attach_tag":
            ensure_page_break(obj, report, label)
        if kind == "sign_org" and cfg["signature"]["enabled"]:
            _pad_signature(obj, report, label, cfg, kinds, i)
        if kind == "sign_date" and cfg["signature"]["enabled"]:
            _pad_date(obj, report, label, cfg)

    if not fix:
        # check 模式：只统计表标题与表格之间是否缺空行，内存里插入不写回
        format_table_caption_spacer(items, [k for k, _o, _t in kinds], report, cfg)

    # 前面插过空段、拆过附件段，新段落还没被全局规范化覆盖：收尾再扫一遍
    light_pass(doc, report, cfg)

    changed = report.total > 0
    if fix and changed:
        _save(doc, path)
    out = report.as_dict(changed)
    out["writeBack"] = bool(fix and changed)
    out["elapsedMs"] = int((time.time() - t0) * 1000)
    out["fontProfile"] = profile
    if fontinfo:
        out["fonts"] = fontinfo
    return out


def _trail_spaces(txt):
    """结尾全角空格的个数。"""
    n = 0
    for ch in reversed(txt):
        if ch == "　":
            n += 1
        else:
            break
    return n


WORDWRAP_SUCCESSORS = ("w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN",
                       "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
                       "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc",
                       "w:textDirection", "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl",
                       "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange")


def _word_wrap_off(par, report, label):
    """落款段落必须写 w:wordWrap="0"。

    Word 右对齐默认把行尾空格排到版心外（等于没写），只有关掉 wordWrap
    （界面上的「允许西文在单词中间换行」）行尾全角空格才计入行宽、把文字挤左。
    参考文档的落款段落正是这个写法，所以手工敲的空格有效。
    """
    pPr = par._p.get_or_add_pPr()
    el = pPr.find(qn("w:wordWrap"))
    cur = None if el is None else el.get(qn("w:val"))
    if cur in ("0", "false", "off"):
        return

    def apply():
        p2 = par._p.get_or_add_pPr()
        e2 = p2.find(qn("w:wordWrap"))
        if e2 is None:
            e2 = OxmlElement("w:wordWrap")
            p2.insert_element_before(e2, *WORDWRAP_SUCCESSORS)
        e2.set(qn("w:val"), "0")

    report.want("signature.wordWrap", label,
                "未关闭（Word 会忽略行尾空格）" if cur is None else "wordWrap=%s" % cur,
                "wordWrap=0（行尾空格参与对齐）", apply)


def _right_indent_pt(par):
    pPr = par._p.find(qn("w:pPr"))
    if pPr is None:
        return 0.0
    ind = pPr.find(qn("w:ind"))
    if ind is None:
        return 0.0
    raw = ind.get(qn("w:right"))
    return int(raw) / 20.0 if raw is not None else 0.0


def _clear_right_indent(par):
    pPr = par._p.find(qn("w:pPr"))
    if pPr is None:
        return
    ind = pPr.find(qn("w:ind"))
    if ind is None:
        return
    for a in ("w:right", "w:rightChars"):
        if ind.get(qn(a)) is not None:
            del ind.attrib[qn(a)]


def _set_tail_ideo(par, count, report, label, rule):
    """把段落尾部的全角空格设为 count 个（单独一个 run，仿宋三号）。"""
    count = int(max(0, round(float(count))))
    txt = par.text
    base = txt.rstrip("　 ")
    ok_count = (_trail_spaces(txt) == count and len(txt) == len(base) + count)
    ok_indent = _right_indent_pt(par) == 0.0
    if ok_count and ok_indent:
        return

    def apply():
        _replace_par_text(par, base)
        _clear_right_indent(par)
        if not count:
            return
        r = OxmlElement("w:r")
        rPr = OxmlElement("w:rPr")
        rf = OxmlElement("w:rFonts")
        rf.set(qn("w:ascii"), "Times New Roman")
        rf.set(qn("w:hAnsi"), "Times New Roman")
        rf.set(qn("w:cs"), "Times New Roman")
        rf.set(qn("w:eastAsia"), "仿宋")
        rPr.append(rf)
        hint = OxmlElement("w:hint")
        hint.set(qn("w:val"), "eastAsia")
        rPr.append(hint)
        for tag in ("w:sz", "w:szCs"):
            e = OxmlElement(tag)
            e.set(qn("w:val"), "32")
            rPr.append(e)
        r.append(rPr)
        t = OxmlElement("w:t")
        t.set(qn("xml:space"), "preserve")
        t.text = "　" * count
        r.append(t)
        par._p.append(r)

    report.want(rule, label,
                "尾部 %d 个全角空格%s" % (_trail_spaces(txt),
                                          "" if ok_indent else "、右缩进 %g 磅" % round(_right_indent_pt(par), 1)),
                "尾部 %d 个全角空格、无右缩进" % count, apply)


def _pad_date(par, report, label, cfg):
    """日期行：尾部留 datePadSpaces 个全角空格（对齐靠 wordWrap=0 + 右对齐）。"""
    pad = float(cfg["signature"]["datePadSpaces"])
    _word_wrap_off(par, report, label)
    _set_tail_ideo(par, pad, report, label, "signature.datePad")


def _norm_date_text(s):
    """把日期文本规范化成「YYYY年M月D日」，用于量宽度。

    署名要按日期宽度算空格数，而日期文本的规范化可能在本次流程稍后才发生，
    直接量原文会算错（`2024.10.27` 是 5 字，`2024年10月27日` 是 7 字）。
    """
    t = s.strip()
    m = RE_DATE_LOOSE.fullmatch(t)
    if m:
        return "%s年%d月%d日" % (m.group(1), int(m.group(2)), int(m.group(3)))
    return t


def _pad_signature(par, report, label, cfg, kinds, i):
    """按规范第 10 条排署名（右对齐 + 右缩进）：
    无需盖章 → 署名第一个字超出日期 orgOverhangSpaces 个全角；
    需盖章   → 署名与日期中心轴对齐。
    """
    date_txt = None
    for j in range(i + 1, len(kinds)):
        if kinds[j][0] == "sign_date":
            date_txt = kinds[j][2]
            break
    if not date_txt:
        return
    dp = float(cfg["signature"]["datePadSpaces"])
    ov = float(cfg["signature"].get("orgOverhangSpaces", 2))
    dw = text_width(_norm_date_text(date_txt))
    base = par.text.rstrip("　 ")
    sw = text_width(base)
    if cfg["signature"]["seal"]:
        pad = dp + dw / 2.0 - sw / 2.0
    else:
        pad = dp + dw + ov - sw
    pad = max(0.0, pad)
    _word_wrap_off(par, report, label)
    _set_tail_ideo(par, pad, report, label, "signature.orgPad")


def doc_text(path):
    doc = Document(path)
    parts = [p.text for p in doc.paragraphs]
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                parts.append(cell.text)
    return "\n".join(parts)


def skip_reason(path, cfg):
    """保留的扩展点：默认不按正文内容跳过任何文档。"""
    kws = cfg.get("skipContentKeywords") or []
    if not kws:
        return None
    try:
        text = doc_text(path)
    except Exception:
        return None
    for kw in kws:
        if kw and kw in text:
            return kw
    return None


def _collect(paths, recursive):
    out = []
    for p in paths:
        if os.path.isdir(p):
            if recursive:
                for root, _dirs, files in os.walk(p):
                    for f in sorted(files):
                        if f.lower().endswith(".docx") and not f.startswith("~$"):
                            out.append(os.path.join(root, f))
            else:
                for f in sorted(os.listdir(p)):
                    fp = os.path.join(p, f)
                    if os.path.isfile(fp) and f.lower().endswith(".docx") and not f.startswith("~$"):
                        out.append(fp)
        else:
            out.append(p)
    return out


def _fonts_cli(args, cfg):
    """--fonts / --install-fonts 的处理入口。"""
    import tempfile
    import shutil

    here = os.path.dirname(os.path.abspath(__file__))
    if here not in sys.path:
        sys.path.insert(0, here)
    import fontcheck

    want_install = (args.install_fonts is not None or args.preset
                    or args.from_zip)

    if args.confirm:
        # 用户授权令牌：满足它等同于 --yes，且令牌只能用一次
        rec = fontcheck.consume_consent(args.confirm)
        if not rec:
            msg = "授权令牌无效或已过期（请重新运行 --fonts 取新的 token）"
            if args.json:
                print(json.dumps({"ok": False, "error": msg,
                                  "reason": "bad-consent"}, ensure_ascii=False))
            else:
                print(msg)
            return 3
        args.yes = True
        if not want_install:
            # 用户只回传了 token：按当初登记的来源装
            src = str(rec.get("source") or "")
            if src.lower().endswith(".zip") or src.startswith(("http://", "https://")):
                args.from_zip = src
            elif src in fontcheck.PRESETS:
                args.preset = src

    if not want_install and not args.yes:
        info = fontcheck.check(cfg)
        if args.json:
            print(json.dumps(info, ensure_ascii=False, indent=2))
        else:
            print("平台：%s    安装目录：%s"
                  % (info["platform"], info["installDir"]))
            for name in info["required"]:
                hit = info["found"].get(name)
                print("  %-4s %s" % ("OK" if hit else "缺失", name))
            if info["missing"]:
                cr = info.get("consentRequired") or {}
                print("")
                print(info["hint"])
                if cr:
                    print(cr["ask"])
                    print("命令：%s" % cr["command"])
        return 0 if info["ok"] else 2

    if not args.yes:
        args.dry_run = True
        print("需要授权：加 --yes 才会下载安装（当前只做 dry-run）")

    tmpdir = tempfile.mkdtemp(prefix="docxfmt-fonts-")
    results = []
    try:
        if args.from_zip:
            results += fontcheck.extract_zip(args.from_zip, tmpdir)
        if args.preset:
            results += fontcheck.download_preset(args.preset, tmpdir)
        for src in (args.install_fonts or []):
            if os.path.isdir(src):
                for fn in sorted(os.listdir(src)):
                    if fn.lower().endswith(fontcheck.FONT_EXTS):
                        results.append({"file": os.path.join(src, fn),
                                        "ok": True})
            elif os.path.isfile(src):
                results.append({"file": src, "ok": True})
            else:
                results.append({"file": src, "ok": False,
                                "reason": "路径不存在"})
        good = [r["file"] for r in results if r.get("ok")]
        if not good:
            print("没有可安装的字体文件")
            if args.json:
                print(json.dumps({"results": results}, ensure_ascii=False,
                                 indent=2))
            return 1
        inst = fontcheck.install(good, dry_run=args.dry_run)
        if args.json:
            print(json.dumps({"downloaded": results, "installed": inst,
                              "dryRun": args.dry_run},
                             ensure_ascii=False, indent=2))
        else:
            for it in results:
                if not it.get("ok"):
                    print("  [跳过] %s  %s" % (it.get("file"), it.get("reason")))
            for it in inst:
                if not it.get("ok"):
                    print("  [失败] %s  %s" % (it.get("file"), it.get("reason")))
                    continue
                verb = {"installed": "已安装", "exists": "已存在",
                        "copy": "将安装"}.get(it["action"], it["action"])
                print("  [%s] %s → %s" % (verb, it.get("family"),
                                          it.get("dest")))
            if not args.dry_run:
                fams = [i["family"] for i in inst if i.get("ok")]
                if fams:
                    fontcheck.save_consent(
                        fams,
                        args.from_zip or args.preset or "local",
                        [i["dest"] for i in inst if i.get("ok")])
                    print("\n重启 Word 后生效；以管理员身份运行的 Word "
                          "看不到用户级字体。")
        return 0 if all(i.get("ok") for i in inst) else 1
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def main(argv=None):
    ap = argparse.ArgumentParser(description="docx 公文格式规范化")
    ap.add_argument("path", nargs="*")
    ap.add_argument("--mode", choices=["apply", "check"], default="apply")
    ap.add_argument("--light", action="store_true",
                    help="只做全局规范化（分页开关/中文引号/中英之间不留空格），不套公文版式")
    ap.add_argument("--recursive", action="store_true")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--config", default=None, help="JSON 配置字符串或 .json 文件路径")
    ap.add_argument("--fonts", action="store_true",
                    help="只报告字体安装情况（不处理文档）")
    ap.add_argument("--install-fonts", nargs="*", default=None,
                    help="安装字体：本地文件/目录（可留空配合 --preset）")
    ap.add_argument("--preset", default=None,
                    help="预置清单：gongwen（普通档）/ gongwen-strict（严格档）")
    ap.add_argument("--from-zip", default=None,
                    help="从 zip（URL 或本地路径）提取字体后安装")
    ap.add_argument("--yes", action="store_true",
                    help="确认下载安装；缺省时只做 dry-run")
    ap.add_argument("--confirm", default=None,
                    help="授权令牌（由 --fonts --json 的 consentRequired.token 给出，"
                         "用户同意后回传；一次性、10 分钟内有效）")
    ap.add_argument("--dry-run", action="store_true", help="只报告不落盘")
    ap.add_argument("--font-profile", choices=["standard", "strict"], default=None,
                    help="字体档：standard（默认）/ strict（GBK/GB2312）")
    ap.add_argument("--no-font-check", action="store_true",
                    help="跳过字体检测")
    args = ap.parse_args(argv)

    cfg = DEFAULTS
    if args.config:
        raw = args.config
        if os.path.isfile(raw):
            with open(raw, "r", encoding="utf-8") as fh:
                raw = fh.read()
        cfg = merge_config(json.loads(raw))

    if args.font_profile:
        cfg["fontProfile"] = args.font_profile
    if args.no_font_check:
        cfg.setdefault("fontCheck", {})["enabled"] = False
    if args.config and args.font_profile == "strict":
        cfg.setdefault("fontCheck", {})
    apply_font_profile(cfg, cfg.get("fontProfile"))

    if args.fonts or args.install_fonts is not None or args.preset \
            or args.from_zip:
        return _fonts_cli(args, cfg)

    if not args.path:
        ap.error("缺少要处理的 docx 路径（或用 --fonts 查看字体）")

    results = []
    failed = False
    violations = 0
    for path in _collect(args.path, args.recursive):
        if not path.lower().endswith(".docx"):
            continue
        reason = skip_reason(path, cfg)
        if reason:
            results.append({
                "ok": True, "file": os.path.abspath(path), "mode": args.mode,
                "skipped": True, "skipReason": "正文含白名单关键词「%s」" % reason,
                "changed": False, "violations": 0, "rules": {}, "stats": {}, "warnings": [],
            })
            continue
        try:
            res = process(path, args.mode, cfg, full=not args.light)
        except Exception as exc:
            import traceback
            res = {"ok": False, "file": os.path.abspath(path), "mode": args.mode,
                   "error": "%s: %s" % (type(exc).__name__, exc),
                   "trace": traceback.format_exc()[-1200:]}
            failed = True
        violations += res.get("violations", 0) or 0
        results.append(res)

    payload = results[0] if len(results) == 1 else {"ok": not failed, "files": results,
                                                    "count": len(results), "violations": violations}
    if args.json:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    else:
        for res in (results if len(results) > 1 else [results[0]]):
            if not res.get("ok"):
                print("[失败] %s\n  %s" % (res.get("file"), res.get("error")))
                continue
            if res.get("skipped"):
                print("[已跳过] %s  (%s)" % (res.get("file"), res.get("skipReason")))
                continue
            tag = "已规范化" if res.get("writeBack") else ("无需修改" if not res["changed"] else "存在偏差")
            print("[%s] %s  (偏差 %d 处, %d ms)" % (tag, res["file"], res["violations"], res.get("elapsedMs", 0)))
            for rule, info in res["rules"].items():
                smp = info["samples"][0]
                print("   - %s ×%d  例: %s: %s → %s" % (rule, info["count"], smp["where"],
                                                        smp["current"], smp["expected"]))
            for w in res["warnings"]:
                print("   ! " + w)
    if failed:
        return 1
    if args.mode == "check" and violations:
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
