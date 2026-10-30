#!/usr/bin/env python3
"""Docx 正文抽取器（纯标准库，不依赖 python-docx、不联网）。

用法:
    python docx_read.py <file.docx> [--offset 0] [--max-chars 12000] [--json]

输出 JSON（--json）:
    {"ok": true, "path": ..., "total_chars": N, "offset": 0, "next_offset": N,
     "text": "...", "paragraphs": N, "tables": N, "truncated": false, "notes": [...]}
"""

import argparse
import json
import sys
import zipfile
from xml.etree import ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def q(tag):
    return W + tag


def para_text(p):
    parts = []
    for node in p.iter():
        tag = node.tag
        if tag == q('t'):
            if node.text:
                parts.append(node.text)
        elif tag == q('tab'):
            parts.append('\t')
        elif tag in (q('br'), q('cr')):
            parts.append('\n')
    return ''.join(parts).strip()


def table_text(tbl):
    rows = []
    for tr in tbl.findall(q('tr')):
        cells = []
        for tc in tr.findall(q('tc')):
            inner = [para_text(p) for p in tc.iter(q('p'))]
            cells.append(' '.join(x for x in inner if x).strip())
        if any(cells):
            rows.append(' | '.join(cells))
    return rows


def extract(path):
    notes = []
    with zipfile.ZipFile(path) as zf:
        names = zf.namelist()
        if 'word/document.xml' not in names:
            return None, notes + ['zip 内没有 word/document.xml（可能是 .doc 或损坏文件）']
        raw = zf.read('word/document.xml')
        has_media = any(n.startswith('word/media/') for n in names)
        has_footnote = 'word/footnotes.xml' in names
    if has_media:
        notes.append('文档含图片（word/media/），本抽取只取文字')
    if has_footnote:
        notes.append('文档含脚注，本抽取不含脚注文字')
    try:
        root = ET.fromstring(raw)
    except ET.ParseError as exc:
        return None, notes + ['document.xml 解析失败: %s' % exc]

    body = root.find(q('body'))
    if body is None:
        return None, notes + ['document.xml 没有 w:body']

    blocks = []
    paragraphs = 0
    tables = 0
    for child in list(body):
        tag = child.tag
        if tag == q('p'):
            text = para_text(child)
            if text:
                blocks.append(text)
                paragraphs += 1
        elif tag == q('tbl'):
            rows = table_text(child)
            if rows:
                blocks.append('\n'.join('[表] ' + r for r in rows))
                tables += 1
        elif tag == q('sdt'):
            for p in child.iter(q('p')):
                text = para_text(p)
                if text:
                    blocks.append(text)
                    paragraphs += 1

    text = '\n'.join(blocks)
    stats = {'paragraphs': paragraphs, 'tables': tables, 'chars': len(text)}
    return (text, stats), notes


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument('path')
    ap.add_argument('--offset', type=int, default=0)
    ap.add_argument('--max-chars', type=int, default=12000)
    ap.add_argument('--json', action='store_true')
    args = ap.parse_args(argv)

    try:
        result, notes = extract(args.path)
    except FileNotFoundError:
        result, notes = None, ['文件不存在']
    except zipfile.BadZipFile:
        result, notes = None, ['不是有效的 docx（zip）文件']
    except OSError as exc:
        result, notes = None, ['读取失败: %s' % exc]

    if result is None:
        payload = {'ok': False, 'path': args.path, 'error': notes[0] if notes else '未知错误', 'notes': notes}
        print(json.dumps(payload, ensure_ascii=False))
        return 1

    text, stats = result
    offset = max(0, min(args.offset, len(text)))
    chunk = text[offset:offset + max(200, args.max_chars)]
    next_offset = offset + len(chunk)
    payload = {
        'ok': True,
        'path': args.path,
        'total_chars': len(text),
        'paragraphs': stats['paragraphs'],
        'tables': stats['tables'],
        'offset': offset,
        'next_offset': next_offset if next_offset < len(text) else None,
        'truncated': next_offset < len(text),
        'text': chunk,
        'notes': notes,
    }
    print(json.dumps(payload, ensure_ascii=False))
    return 0


if __name__ == '__main__':
    sys.exit(main())
