#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""字体检测 / 授权下载 / 安装（纯标准库，跨平台）

三条纪律：

1. **检测永远不联网** —— 只读注册表、扫字体目录。
2. **下载必须带授权** —— 由检测结果发出的一次性 token（内存、10 分钟），
   或命令行显式 `--yes`；没有授权只做 dry-run 报告。
3. **字体文件不随包分发** —— 只在用户同意后按 URL 取回，源可配置。

命令行（也可以由 docxfmt.py 的 --fonts / --install-fonts 间接调用）：

    python fontcheck.py --check
    python fontcheck.py --check --json
    python fontcheck.py --install --preset gongwen --yes
    python fontcheck.py --install --from-zip <URL|PATH> --yes
    python fontcheck.py --install /path/to/fonts --yes --dry-run
"""

import argparse
import hashlib
import io
import json
import os
import platform
import re
import shutil
import struct
import sys
import tempfile
import time
import zipfile

# ---------------------------------------------------------------- 常量

# 单文件下载源（{file} 会被替换成 URL 编码后的文件名）
DEFAULT_SOURCES = [
    "https://cdn.jsdelivr.net/gh/luoyoulai-collab/gongwen-fonts@main/fonts/{file}",
    "https://raw.githubusercontent.com/luoyoulai-collab/gongwen-fonts/main/fonts/{file}",
]

# 整包下载源：晋城市财政局「常用字体」下载中心，一份覆盖两档
GONGWEN_ZIP = ("http://czj.jcgov.gov.cn/ggfw/xzzx/202506/"
               "P020250624600205358001.zip")

# 预置清单：文件名 → 期望的字体族名
PRESETS = {
    # 普通档：小标宋简体 + 仿宋_GB2312 + 楷体_GB2312（约 12 MB）
    "gongwen": [
        ("方正小标宋简体.ttf", "方正小标宋简体"),
        ("仿宋_GB2312.ttf", "仿宋_GB2312"),
        ("楷体_GB2312.ttf", "楷体_GB2312"),
    ],
    # 严格档：GBK 小标宋/黑体 + GB2312 仿宋/楷体（约 33 MB）
    "gongwen-strict": [
        ("方正小标宋_GBK.ttf", "方正小标宋_GBK"),
        ("方正黑体_GBK.ttf", "方正黑体_GBK"),
        ("仿宋_GB2312.ttf", "仿宋_GB2312"),
        ("楷体_GB2312.ttf", "楷体_GB2312"),
    ],
}

# 期望字体 → 可接受的等价族名（规范化后比较）。
# 注意方向：GB2312/GBK 版本可以顶替普通名（更全），反过来不行。
ALIASES = {
    "仿宋": {"仿宋", "仿宋_gb2312", "fangsong", "fangsong_gb2312", "simfang"},
    "仿宋_gb2312": {"仿宋_gb2312", "fangsong_gb2312", "fangsong"},
    "楷体": {"楷体", "楷体_gb2312", "kaiti", "kaiti_gb2312", "simkai"},
    "楷体_gb2312": {"楷体_gb2312", "kaiti_gb2312", "kaiti"},
    "黑体": {"黑体", "simhei", "heiti", "方正黑体_gbk"},
    "方正黑体_gbk": {"方正黑体_gbk", "fzheiti_gbk"},
    "方正小标宋简体": {"方正小标宋简体", "fzxiaobiaosong_b05s"},
    "方正小标宋_gbk": {"方正小标宋_gbk", "fzxiaobiaosong_gbk"},
    "宋体": {"宋体", "simsun", "songti"},
    "times new roman": {"times new roman", "timesnewroman", "timesnewromanpsmt"},
}

# 严格档的字体必须精确命中，不吃普通名兜底
STRICT_NAMES = {"方正小标宋_gbk", "仿宋_gb2312", "楷体_gb2312", "方正黑体_gbk"}

# 字体文件的 sfnt 头
SFNT_MAGIC = {
    b"\x00\x01\x00\x00": "truetype",
    b"OTTO": "cff",
    b"true": "truetype",
    b"ttcf": "collection",
}

FONT_EXTS = (".ttf", ".otf", ".ttc")


# ---------------------------------------------------------------- 规范化

def _norm(name):
    """字体名规范化：小写、去空白与分隔符。"""
    if not name:
        return ""
    s = str(name).strip().lower()
    s = s.replace("\u3000", " ")
    for ch in (" ", "-", "_", "."):
        s = s.replace(ch, "")
    # 去掉注册表里的 “(TrueType)” 尾巴
    s = re.sub(r"\((truetype|opentype|all res)\)$", "", s)
    return s.strip()


def resolve_alias(name):
    """返回该字体名可接受的等价名集合（规范化后）。"""
    key = str(name or "").strip()
    nk = _norm(key)
    low = key.lower()
    cands = {nk}
    for alias_key, group in ALIASES.items():
        if _norm(alias_key) == nk or alias_key.lower() == low:
            cands |= {_norm(g) for g in group}
            break
    return {c for c in cands if c}


def is_strict_name(name):
    return _norm(name) in {_norm(s) for s in STRICT_NAMES}


# ---------------------------------------------------------------- 平台

def _platform_tag():
    if sys.platform.startswith("win"):
        return "windows"
    if sys.platform == "darwin":
        return "macos"
    return "linux"


def font_dirs():
    """返回会去扫描的字体目录列表。"""
    out = []
    home = os.path.expanduser("~")
    if sys.platform.startswith("win"):
        windir = os.environ.get("WINDIR") or r"C:\Windows"
        out.append(os.path.join(windir, "Fonts"))
        local = os.environ.get("LOCALAPPDATA") or os.path.join(
            home, "AppData", "Local")
        out.append(os.path.join(local, "Microsoft", "Windows", "Fonts"))
    elif sys.platform == "darwin":
        out.append(os.path.join(home, "Library", "Fonts"))
        out.append("/Library/Fonts")
        out.append("/System/Library/Fonts")
    else:
        out.append(os.path.join(home, ".local", "share", "fonts"))
        out.append(os.path.join(home, ".fonts"))
        out.append("/usr/share/fonts")
        out.append("/usr/local/share/fonts")
    return [d for d in out if os.path.isdir(d)]


def install_dir():
    """用户级安装目录（写入不需要管理员）。"""
    home = os.path.expanduser("~")
    if sys.platform.startswith("win"):
        local = os.environ.get("LOCALAPPDATA") or os.path.join(
            home, "AppData", "Local")
        return os.path.join(local, "Microsoft", "Windows", "Fonts")
    if sys.platform == "darwin":
        return os.path.join(home, "Library", "Fonts")
    return os.path.join(home, ".local", "share", "fonts")


# ---------------------------------------------------------------- 已装字体

_WIN_FONTS_KEY = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"


def _win_registry_fonts():
    """从 HKLM + HKCU 读已注册字体：{规范化族名: 文件名}。"""
    out = {}
    try:
        import winreg
    except ImportError:
        return out
    for hive, sub in (
        (winreg.HKEY_LOCAL_MACHINE, _WIN_FONTS_KEY),
        (winreg.HKEY_CURRENT_USER, _WIN_FONTS_KEY),
    ):
        try:
            key = winreg.OpenKey(hive, sub, 0, winreg.KEY_READ)
        except OSError:
            continue
        try:
            i = 0
            while True:
                try:
                    name, value, _kind = winreg.EnumValue(key, i)
                except OSError:
                    break
                i += 1
                family = re.sub(r"\s*\((TrueType|OpenType|All Res)\)\s*$",
                                "", str(name), flags=re.I).strip()
                if family:
                    out.setdefault(_norm(family), str(value))
        finally:
            key.Close()
    return out


def installed_fonts():
    """{规范化族名: 来源}。注册表优先，其次字体文件名。"""
    found = {}
    for base in font_dirs():
        try:
            names = os.listdir(base)
        except OSError:
            continue
        for fn in names:
            if not fn.lower().endswith(FONT_EXTS):
                continue
            stem = os.path.splitext(fn)[0]
            found.setdefault(_norm(stem), fn)
    if sys.platform.startswith("win"):
        for k, v in _win_registry_fonts().items():
            found[k] = v
    return found


# ---------------------------------------------------------------- 期望字体

def required_fonts(cfg):
    """按配置算出这条文档需要的字体名（去重、保序）。"""
    fonts = (cfg or {}).get("fonts") or {}
    keys = ("title", "tableCaption", "subtitle", "body", "h3", "h1", "h2",
            "tableHead", "tableBody", "footer", "latin")
    out = []
    for k in keys:
        v = fonts.get(k)
        if v and v not in out:
            out.append(v)
    return out


def _match(installed, want, strict=False):
    """want 是否已装。strict=True 时不吃普通名兜底。"""
    cands = resolve_alias(want)
    if strict or is_strict_name(want):
        cands = {_norm(want)} | {c for c in cands if c == _norm(want)}
    for c in cands:
        if c in installed:
            return installed[c]
    return None


def check(cfg, *, with_consent=True):
    """检测缺哪些字体。**不联网**。

    返回 dict：ok / required / found / missing / platform / installDir / hint
    missing 非空且未授权时附带 consentRequired。
    """
    req = required_fonts(cfg)
    inst = installed_fonts()
    found, missing = {}, []
    for name in req:
        hit = _match(inst, name, strict=is_strict_name(name))
        if hit:
            found[name] = hit
        else:
            missing.append(name)

    out = {
        "ok": not missing,
        "platform": _platform_tag(),
        "profile": (cfg or {}).get("fontProfile") or "standard",
        "required": req,
        "found": found,
        "missing": missing,
        "installDir": install_dir(),
        "hint": "",
    }
    if missing:
        out["hint"] = ("缺 %d 个字体；可先 `--fonts` 看清单，"
                       "授权后 `--install-fonts --preset gongwen --yes` 安装。"
                       % len(missing))
        if with_consent:
            source = (cfg or {}).get("fontCheck", {}).get("sources") or []
            src = source[0] if source else GONGWEN_ZIP
            est = _estimate_bytes(missing)
            token = issue_consent(missing, src)
            out["consentRequired"] = {
                "fonts": missing,
                "bytes": est,
                "mb": round(est / 1048576.0, 1),
                "source": src,
                "token": token,
                "ask": ("系统缺少 %s 共 %d 个字体（约 %.1f MB），"
                        "来源：%s，是否下载安装？"
                        % ("、".join(missing), len(missing),
                           est / 1048576.0, src)),
                "command": ("python lib/docxfmt.py --install-fonts "
                            "--preset gongwen --yes"),
            }
    return out


def _estimate_bytes(names):
    """按预置清单里已知文件大小粗估（拿不到就按 4 MB/个）。"""
    known = {
        "方正小标宋简体": 3806520,
        "方正小标宋_gbk": 13904802,
        "仿宋_gb2312": 3996872,
        "楷体_gb2312": 4135804,
        "方正黑体_gbk": 9961614,
        "仿宋": 3996872,
        "楷体": 4135804,
        "黑体": 9961614,
        "宋体": 10500000,
        "times new roman": 1600000,
    }
    total = 0
    for n in names:
        total += known.get(_norm(n), 4 * 1024 * 1024)
    return total


# ---------------------------------------------------------------- 授权

_CONSENT_TTL = 600.0        # 10 分钟
_consent_tokens = {}        # token -> {"fonts":[...], "source":str, "at":ts}
_consent_file = None


def consent_path():
    global _consent_file
    if _consent_file:
        return _consent_file
    home = os.path.expanduser("~")
    _consent_file = os.path.join(home, ".dsh", "dsh-docx-format",
                                 "font-consent.json")
    return _consent_file


def issue_consent(fonts, source):
    """生成一次性 token（只存内存）。"""
    raw = "|".join(sorted(fonts)) + "|" + str(source) + "|" + str(time.time())
    token = hashlib.sha256(raw.encode("utf-8")).hexdigest()[:16]
    _consent_tokens[token] = {
        "fonts": list(fonts), "source": str(source), "at": time.time(),
    }
    # 顺手清掉过期 token
    now = time.time()
    for k in [k for k, v in _consent_tokens.items()
              if now - v["at"] > _CONSENT_TTL]:
        _consent_tokens.pop(k, None)
    return token


def verify_consent(token, consume=True):
    """校验 token；consume=True 时用掉即失效。"""
    if not token:
        return False
    rec = _consent_tokens.get(token)
    if not rec:
        return False
    if time.time() - rec["at"] > _CONSENT_TTL:
        _consent_tokens.pop(token, None)
        return False
    if consume:
        _consent_tokens.pop(token, None)
    return True


def consume_consent(token):
    """校验并取出 token 记录（一次性）；无效/过期返回 None。"""
    if not token:
        return None
    rec = _consent_tokens.get(token)
    if not rec:
        return None
    if time.time() - rec["at"] > _CONSENT_TTL:
        _consent_tokens.pop(token, None)
        return None
    _consent_tokens.pop(token, None)
    return rec


def load_consent():
    """读已授权记录：{ "<font>@<source>": {...} }"""
    path = consent_path()
    try:
        with io.open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def save_consent(fonts, source, files=None):
    """落盘授权凭据。返回写入后的完整字典。"""
    data = load_consent()
    stamp = time.strftime("%Y-%m-%dT%H:%M:%S")
    for name in fonts:
        key = "%s@%s" % (name, source)
        data[key] = {
            "grantedAt": stamp,
            "source": str(source),
            "files": list(files or []),
        }
    path = consent_path()
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        tmp = path + ".tmp"
        with io.open(tmp, "w", encoding="utf-8") as fh:
            json.dump(data, fh, ensure_ascii=False, indent=2)
        os.replace(tmp, path)
    except OSError:
        pass
    return data


def consented_fonts(cfg=None):
    """已授权过的字体名集合（同批不再重复提问）。"""
    data = load_consent()
    out = set()
    for key, rec in data.items():
        name = key.rsplit("@", 1)[0]
        if rec and rec.get("grantedAt"):
            out.add(name)
    return out


# ---------------------------------------------------------------- 字体文件校验

def read_family(path):
    """解析 sfnt name 表，返回字体族名；解析失败返回 None。"""
    try:
        with open(path, "rb") as fh:
            head = fh.read(12)
            if len(head) < 12:
                return None
            magic = head[:4]
            if magic not in SFNT_MAGIC:
                return None
            if magic == b"ttcf":
                fh.seek(12)
                first = fh.read(4)
                if len(first) < 4:
                    return None
                base = struct.unpack(">I", first)[0]
                fh.seek(base)
                head = fh.read(12)
                if len(head) < 12 or head[:4] not in SFNT_MAGIC:
                    return None
            num_tables = struct.unpack(">H", head[4:6])[0]
            fh.seek(12)
            directory = fh.read(16 * num_tables)
            name_off = name_len = None
            for i in range(num_tables):
                rec = directory[i * 16:(i + 1) * 16]
                if len(rec) < 16:
                    break
                tag, _sum, off, length = struct.unpack(">4sIII", rec)
                if tag == b"name":
                    name_off, name_len = off, length
                    break
            if name_off is None:
                return None
            fh.seek(name_off)
            table = fh.read(min(name_len, 65536))
            if len(table) < 6:
                return None
            _fmt, count, str_off = struct.unpack(">HHH", table[:6])
            best = {}
            for i in range(count):
                rec = table[6 + i * 12: 6 + (i + 1) * 12]
                if len(rec) < 12:
                    break
                pid, eid, _lid, nid, ln, off = struct.unpack(">HHHHHH", rec)
                if nid not in (1, 16):
                    continue
                raw = table[str_off + off: str_off + off + ln]
                if pid == 3:
                    try:
                        text = raw.decode("utf-16-be", "ignore")
                    except Exception:
                        continue
                elif pid == 1:
                    text = raw.decode("latin-1", "ignore")
                else:
                    continue
                text = text.strip("\x00").strip()
                if not text:
                    continue
                # 16 是 typographic family，优先
                if nid not in best or nid == 16:
                    best[nid] = text
            return best.get(16) or best.get(1) or None
    except (OSError, struct.error, IndexError):
        return None


def valid_font_file(path, expect_family=None):
    """校验字体文件：magic + 能解析出族名（可选比对期望名）。"""
    try:
        if os.path.getsize(path) < 1024:
            return False, "文件过小"
    except OSError:
        return False, "文件不存在"
    try:
        with open(path, "rb") as fh:
            magic = fh.read(4)
    except OSError:
        return False, "读取失败"
    if magic not in SFNT_MAGIC:
        return False, "不是字体文件（magic=%r）" % magic
    family = read_family(path)
    if not family:
        return False, "字体名解析失败"
    if expect_family:
        wanted = resolve_alias(expect_family)
        if _norm(family) not in wanted:
            return False, "字体名不符（实际 %s，期望 %s）" % (family, expect_family)
    return True, family


# ---------------------------------------------------------------- 下载

def _http_get(url, timeout=120, headers=None):
    import urllib.request
    req = urllib.request.Request(url, headers=headers or {
        "User-Agent": "dsh-docx-format-fontcheck/1.0",
        "Accept": "*/*",
    })
    return urllib.request.urlopen(req, timeout=timeout)


def _download_to(url, dest, timeout=300):
    """下载到 dest，返回字节数。"""
    tmp = dest + ".part"
    total = 0
    with _http_get(url, timeout=timeout) as resp, open(tmp, "wb") as fh:
        while True:
            chunk = resp.read(262144)
            if not chunk:
                break
            fh.write(chunk)
            total += len(chunk)
    os.replace(tmp, dest)
    return total


def _source_urls(fname, sources):
    """按源模板拼 URL；模板不含 {file} 时当目录前缀。"""
    from urllib.parse import quote
    urls = []
    quoted = quote(fname)
    for s in (sources or DEFAULT_SOURCES):
        if "{file}" in s:
            urls.append(s.replace("{file}", quoted))
        elif "{name}" in s:
            urls.append(s.replace("{name}", quoted))
        else:
            urls.append(s.rstrip("/") + "/" + quoted)
    return urls


def download_preset(preset, dest_dir, sources=None, timeout=300):
    """按预置清单逐文件下载。返回 [{file, ok, reason, bytes, family}]。"""
    items = PRESETS.get(preset)
    if not items:
        return [{"file": preset, "ok": False,
                 "reason": "未知预置 %s（可用：%s）"
                           % (preset, ", ".join(PRESETS))}]
    os.makedirs(dest_dir, exist_ok=True)
    results = []
    for fname, family in items:
        dest = os.path.join(dest_dir, fname)
        if os.path.isfile(dest):
            ok, info = valid_font_file(dest, family)
            if ok:
                results.append({"file": dest, "ok": True, "family": info,
                                "bytes": os.path.getsize(dest),
                                "reason": "已存在"})
                continue
        last = None
        for url in _source_urls(fname, sources):
            try:
                n = _download_to(url, dest, timeout=timeout)
                ok, info = valid_font_file(dest, family)
                if ok:
                    results.append({"file": dest, "ok": True, "family": info,
                                    "bytes": n, "reason": url})
                    last = None
                    break
                last = "校验失败：%s" % info
                try:
                    os.remove(dest)
                except OSError:
                    pass
            except Exception as exc:                       # noqa: BLE001
                last = "%s: %s" % (type(exc).__name__, exc)
        else:
            results.append({"file": fname, "ok": False,
                            "reason": last or "无可用源"})
    return results


# ---------------------------------------------------------------- zip

def _safe_zip_member(name):
    """zip 条目是否安全：拒绝绝对路径、盘符、../ 穿越。"""
    n = str(name).replace("\\", "/")
    if not n or n.endswith("/"):
        return None
    if n.startswith("/") or re.match(r"^[A-Za-z]:", n):
        return None
    parts = [p for p in n.split("/") if p not in ("", ".")]
    if any(p == ".." for p in parts):
        return None
    return parts[-1] if parts else None


def extract_zip(src, dest_dir, allow=None):
    """从 zip（本地路径或 URL）提取字体文件。返回 [{file, ok, reason}]。

    只认 .ttf/.otf/.ttc；路径穿越、绝对路径、超大条目一律丢弃。
    """
    os.makedirs(dest_dir, exist_ok=True)
    allow = allow or FONT_EXTS
    tmp_zip = None
    results = []
    try:
        if re.match(r"^https?://", str(src), flags=re.I):
            fd, tmp_zip = tempfile.mkstemp(suffix=".zip")
            os.close(fd)
            _download_to(src, tmp_zip)
            path = tmp_zip
        else:
            path = src
            if not os.path.isfile(path):
                return [{"file": str(src), "ok": False, "reason": "文件不存在"}]
        with zipfile.ZipFile(path) as zf:
            for info in zf.infolist():
                base = _safe_zip_member(info.filename)
                if not base:
                    if info.filename.lower().endswith(FONT_EXTS):
                        results.append({"file": info.filename, "ok": False,
                                        "reason": "拒绝：路径不安全"})
                    continue
                if not base.lower().endswith(tuple(allow)):
                    continue
                if info.file_size > 64 * 1024 * 1024:
                    results.append({"file": base, "ok": False,
                                    "reason": "拒绝：文件过大"})
                    continue
                dest = os.path.join(dest_dir, base)
                try:
                    with zf.open(info) as fsrc, open(dest, "wb") as fdst:
                        shutil.copyfileobj(fsrc, fdst)
                except Exception as exc:                   # noqa: BLE001
                    results.append({"file": base, "ok": False,
                                    "reason": "%s: %s" % (type(exc).__name__, exc)})
                    continue
                ok, info2 = valid_font_file(dest)
                if ok:
                    results.append({"file": dest, "ok": True, "family": info2,
                                    "bytes": os.path.getsize(dest)})
                else:
                    try:
                        os.remove(dest)
                    except OSError:
                        pass
                    results.append({"file": base, "ok": False, "reason": info2})
    except zipfile.BadZipFile:
        results.append({"file": str(src), "ok": False, "reason": "不是有效 zip"})
    except Exception as exc:                               # noqa: BLE001
        results.append({"file": str(src), "ok": False,
                        "reason": "%s: %s" % (type(exc).__name__, exc)})
    finally:
        if tmp_zip:
            try:
                os.remove(tmp_zip)
            except OSError:
                pass
    return results


# ---------------------------------------------------------------- 安装

def _win_register(path, family):
    """写 HKCU 注册表 + 让 GDI 立即生效。"""
    import winreg
    import ctypes

    target = os.path.join(install_dir(), os.path.basename(path))
    if os.path.abspath(path) != os.path.abspath(target):
        os.makedirs(install_dir(), exist_ok=True)
        shutil.copy2(path, target)
    value_name = "%s (TrueType)" % (family or os.path.splitext(
        os.path.basename(target))[0])
    key = winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, _WIN_FONTS_KEY, 0,
                             winreg.KEY_SET_VALUE)
    try:
        winreg.SetValueEx(key, value_name, 0, winreg.REG_SZ, target)
    finally:
        key.Close()
    try:
        ctypes.windll.gdi32.AddFontResourceW(target)
        HWND_BROADCAST, WM_FONTCHANGE, SMTO_ABORTIFHUNG = 0xFFFF, 0x001D, 0x0002
        ctypes.windll.user32.SendMessageTimeoutW(
            HWND_BROADCAST, WM_FONTCHANGE, 0, 0, SMTO_ABORTIFHUNG, 1000, None)
    except Exception:                                      # noqa: BLE001
        pass
    return target


def _posix_install(path, family):
    dest_dir = install_dir()
    os.makedirs(dest_dir, exist_ok=True)
    target = os.path.join(dest_dir, os.path.basename(path))
    if os.path.abspath(path) != os.path.abspath(target):
        shutil.copy2(path, target)
    if sys.platform != "darwin":
        try:
            import subprocess
            subprocess.run(["fc-cache", "-f", dest_dir],
                           check=False, capture_output=True, timeout=60)
        except Exception:                                  # noqa: BLE001
            pass
    return target


def install(paths, *, dry_run=False):
    """安装字体文件到用户级目录。返回 [{file, family, action, dest}]。"""
    out = []
    for path in paths:
        if not os.path.isfile(path):
            out.append({"file": str(path), "ok": False, "action": "missing",
                        "reason": "文件不存在"})
            continue
        ok, info = valid_font_file(path)
        if not ok:
            out.append({"file": path, "ok": False, "action": "invalid",
                        "reason": info})
            continue
        family = info
        if sys.platform.startswith("win"):
            target = os.path.join(install_dir(), os.path.basename(path))
            exists = os.path.isfile(target) and \
                os.path.getsize(target) == os.path.getsize(path)
            if dry_run:
                out.append({"file": path, "family": family, "ok": True,
                            "action": "exists" if exists else "copy",
                            "dest": target})
                continue
            if exists:
                # 文件在，但注册表可能缺，补一次注册
                dest = _win_register(path, family)
                out.append({"file": path, "family": family, "ok": True,
                            "action": "exists", "dest": dest})
            else:
                dest = _win_register(path, family)
                out.append({"file": path, "family": family, "ok": True,
                            "action": "installed", "dest": dest})
        else:
            target = os.path.join(install_dir(), os.path.basename(path))
            exists = os.path.isfile(target) and \
                os.path.getsize(target) == os.path.getsize(path)
            if dry_run:
                out.append({"file": path, "family": family, "ok": True,
                            "action": "exists" if exists else "copy",
                            "dest": target})
                continue
            dest = _posix_install(path, family)
            out.append({"file": path, "family": family, "ok": True,
                        "action": "exists" if exists else "installed",
                        "dest": dest})
    return out


# ---------------------------------------------------------------- CLI

def _cli(argv=None):
    ap = argparse.ArgumentParser(description="dsh-docx-format 字体检测/安装")
    ap.add_argument("source", nargs="*", help="本地字体文件或目录")
    ap.add_argument("--check", action="store_true", help="只看缺什么")
    ap.add_argument("--install", action="store_true", help="安装字体")
    ap.add_argument("--preset", default=None,
                    help="预置清单：%s" % ", ".join(PRESETS))
    ap.add_argument("--from-zip", default=None,
                    help="从 zip（URL 或本地路径）提取字体")
    ap.add_argument("--yes", action="store_true", help="确认下载安装（必需）")
    ap.add_argument("--dry-run", action="store_true", help="只报告不落盘")
    ap.add_argument("--json", action="store_true")
    args = ap.parse_args(argv)

    cfg = {"fonts": {}}
    # 借用 docxfmt 的 DEFAULTS（拿不到就用最小集）
    try:
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        import docxfmt
        cfg = docxfmt.DEFAULTS
    except Exception:                                      # noqa: BLE001
        cfg = {"fonts": {"title": "方正小标宋简体", "body": "仿宋",
                         "h1": "黑体", "h2": "楷体", "h3": "仿宋",
                         "footer": "宋体", "tableHead": "黑体",
                         "tableBody": "仿宋",
                         "tableCaption": "方正小标宋简体",
                         "latin": "Times New Roman"}}

    if args.check or not args.install:
        info = check(cfg)
        if args.json:
            print(json.dumps(info, ensure_ascii=False, indent=2))
        else:
            print("平台：%s    安装目录：%s" % (info["platform"], info["installDir"]))
            for name in info["required"]:
                hit = info["found"].get(name)
                print("  %s %s%s" % ("OK  " if hit else "缺失", name,
                                     "" if hit else "   ← %s" % _suggest(name)))
            if info["missing"]:
                print("\n%s" % info["hint"])
                cr = info.get("consentRequired") or {}
                if cr:
                    print(cr["ask"])
                    print("命令：%s" % cr["command"])
        return 0 if info["ok"] else 2

    if not args.yes:
        print("需要授权：加 --yes 才会下载安装（当前只做 dry-run）")
        args.dry_run = True

    tmpdir = tempfile.mkdtemp(prefix="docxfmt-fonts-")
    results = []
    try:
        if args.from_zip:
            results += extract_zip(args.from_zip, tmpdir)
        if args.preset:
            results += download_preset(args.preset, tmpdir)
        for src in args.source:
            if os.path.isdir(src):
                for fn in os.listdir(src):
                    if fn.lower().endswith(FONT_EXTS):
                        results.append({"file": os.path.join(src, fn),
                                        "ok": True})
            elif os.path.isfile(src):
                results.append({"file": src, "ok": True})
        good = [r["file"] for r in results if r.get("ok")]
        if not good:
            print("没有可安装的字体文件")
            if args.json:
                print(json.dumps({"results": results}, ensure_ascii=False, indent=2))
            return 1
        inst = install(good, dry_run=args.dry_run)
        if args.json:
            print(json.dumps({"downloaded": results, "installed": inst,
                              "dryRun": args.dry_run},
                             ensure_ascii=False, indent=2))
        else:
            for item in results:
                if not item.get("ok"):
                    print("  [跳过] %s  %s" % (item.get("file"), item.get("reason")))
            for item in inst:
                if not item.get("ok"):
                    print("  [失败] %s  %s" % (item.get("file"), item.get("reason")))
                    continue
                verb = {"installed": "已安装", "exists": "已存在",
                        "copy": "将安装"}.get(item["action"], item["action"])
                print("  [%s] %s → %s" % (verb, item.get("family"), item.get("dest")))
            if not args.dry_run:
                names = [i["family"] for i in inst if i.get("ok")]
                if names:
                    save_consent(names, args.from_zip or args.preset or "local",
                                 [i["dest"] for i in inst if i.get("ok")])
                    print("\n重启 Word 后生效。以管理员身份运行的 Word 看不到"
                          "用户级字体，请用普通账户启动。")
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)
    return 0


def _suggest(name):
    n = _norm(name)
    if "gbk" in n or "gb2312" in n:
        return "严格档字体，可 --install-fonts --preset gongwen-strict"
    return "可 --install-fonts --preset gongwen"


if __name__ == "__main__":
    sys.exit(_cli())
