#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""degree_adapt.py — 学位层次适配（agent 后处理步骤，不改引擎）。

在「硕士」格式生成物上，根据 --degree 把封面 / 页眉 / 页脚 / 后置小节标题中的
「硕士 / master」改写为本科 / bachelor 或 博士 / doctor：

  - 封面（中文封面「硕士学位论文」+「硕士研究生：」、英文封面
    「Dissertation for the Master Degree」「Master Degree in Engineering」等）
  - 页眉（默认页眉「哈尔滨工业大学硕士学位论文」、后置部分小节页眉
    「攻读硕士学位期间取得的科研成果」）
  - 页脚（如有）
  - 正文后置小节标题「攻读硕士学位期间取得的科研成果」

degree=master（默认）时为空操作，文档本就是硕士格式，原样保留。

设计要点（对照铁律#1：不侵入式改引擎）：
  - 直接对 docx 包内 document.xml 与各 header/footer 的 XML 做文本改写，
    因此无论「硕士研究生：」在普通段落、表格单元格还是封面文本框里都能命中。
  - **关键**：封面模板常把「硕士」拆成多个 <w:t> 跑（如「硕」+「士研究生」），
    逐跑替换会漏掉。本脚本在每个块内做「逻辑文本重组 + 位置 preserving 重排」，
    既修正拆分，又尽量保留原有字形（仅被替换 token 所在跑的字形会被合并）。
  - 正文一般内容（若偶然出现「硕士」）不动：封面区以「首个一级标题 / 摘要 / 声明」
    为界，界外仅改写学位相关的小节标题，避免误伤作者正文。

用法：
  python skill/scripts/degree_adapt.py <in.docx> --degree bachelor --out <out.docx>
  python skill/scripts/degree_adapt.py <in.docx> --degree doctor          # 覆盖写回
  python skill/scripts/degree_adapt.py <in.docx> --degree master          # 空操作
"""
from __future__ import annotations

import argparse
import os
import re
import sys
import tempfile
import zipfile

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# 中文学位层次映射（master 保留原值）
ZH = {"bachelor": "本科", "doctor": "博士", "master": "硕士"}
EN_CAP = {"bachelor": "Bachelor", "doctor": "Doctor", "master": "Master"}
EN_LOW = {"bachelor": "bachelor", "doctor": "doctor", "master": "master"}

# 正文后置部分小节标题（学位相关，需整体一致）
ACHIEVEMENTS_TITLE = "攻读硕士学位期间取得的科研成果"

# 封面区在「第一个一级标题 / 摘要 / 声明」之前结束
STOP_TEXTS = {"摘要", "ABSTRACT", "Abstract", "声明", "原创性声明"}

# HIT 学位专属短语（几乎不可能出现在论文正文中），用于全文匹配。
# 引擎常把「封面信息表」放在 摘要 之后，导致封面区边界提前关闭而漏改；
# 用这些短语在全文范围内兜底，确保 bachelor/doctor 下无残留「硕士/master」。
DEGREE_PHRASES = [
    "硕士学位论文",
    "硕士研究生",
    "Dissertation for the Master Degree",
    "Master Degree in",
    "Master's Study",
    "攻读硕士学位期间取得的科研成果",
]

# 声明区起始标记（原创性声明 / 版权使用授权书）
DECL_START_MARKERS = [
    "原创性声明",
    "学位论文版权使用授权书",
    "学位论文版权",
    "版权使用授权书",
]

# 机构名保护：声明区「研究生」→「本科生」时，绝不能把「深圳研究生院」等
# 机构名误改成「深圳本科生院」。
INSTITUTION_PROTECT = "研究生院"


def w(tag: str) -> str:
    return f"{{{W}}}{tag}"


def find_spans(logical: str, zh: str, en_cap: str, en_low: str):
    """扫描 logical，返回有序、不重叠的 (start, end, replacement) 列表。

    - 中文优先匹配长词「硕士研究生」(5字)，再「硕士」(2字)；
      硕士→zh；硕士研究生→「本科生研究生」/「博士研究生」/「硕士研究生」。
    - 英文匹配 master('s)?（大小写不敏感），按首字母大小写选择 Bachelor/Doctor
      或 bachelor/doctor，并保留 's。
    """
    spans = []
    i, n = 0, len(logical)
    while i < n:
        m = re.match(r"master('s)?", logical[i:], re.IGNORECASE)
        if m:
            orig = m.group(0)
            suf = m.group(1) or ""
            rep = (en_cap if orig[0].isupper() else en_low) + suf
            spans.append((i, i + len(orig), rep))
            i += len(orig)
            continue
        if logical[i:i + 5] == "硕士研究生":
            rep = {"本科": "本科生研究生", "博士": "博士研究生",
                   "硕士": "硕士研究生"}[zh]
            spans.append((i, i + 5, rep))
            i += 5
            continue
        if logical[i:i + 2] == "硕士":
            spans.append((i, i + 2, zh))
            i += 2
            continue
        i += 1
    return spans


def reflow_block(runs, spans, logical: str) -> None:
    """把替换结果写回各 w:t 跑，保留未被替换部分的原有字形。

    runs: 该块内所有 <w:t> 元素（按文档顺序）。spans 来自 find_spans。
    """
    if not spans:
        return
    n = len(logical)
    consumed = [False] * n
    span_start = {}
    for (s, e, rep) in spans:
        for p in range(s, e):
            consumed[p] = True
        span_start[s] = (e, rep)

    # 每块跑的起始逻辑位置
    run_start = []
    pos = 0
    for r in runs:
        run_start.append(pos)
        pos += len(r.text or "")

    out = [[] for _ in runs]
    p = 0
    for j, r in enumerate(runs):
        t = r.text or ""
        k = 0
        while k < len(t) and p < n:
            # 先处理替换起点：把 adapted token 写入本跑（覆盖逻辑位置 s..e-1）
            if p in span_start:
                e, rep = span_start[p]
                out[j].extend(list(rep))
                while p < e:
                    p += 1
                    k += 1
                continue
            # 其余被替换 token 占用的位置（token 内部、非起点）直接跳过
            if consumed[p]:
                p += 1
                k += 1
                continue
            out[j].append(t[k])
            p += 1
            k += 1
    for j, r in enumerate(runs):
        r.text = "".join(out[j])


def block_runs(elem):
    """返回 elem 下按文档顺序排列的 <w:t> 元素列表。"""
    return list(elem.iter(w("t")))


def block_logical(runs) -> str:
    return "".join(r.text or "" for r in runs)


def adapt_block(elem, zh: str, en_cap: str, en_low: str) -> None:
    runs = block_runs(elem)
    logical = block_logical(runs)
    spans = find_spans(logical, zh, en_cap, en_low)
    if spans:
        reflow_block(runs, spans, logical)


def _para_text(p_elem) -> str:
    return "".join(t.text or "" for t in p_elem.iter(w("t")))


def _tbl_text(tbl_elem) -> str:
    return "".join(t.text or "" for t in tbl_elem.iter(w("t")))


def _contains_degree_phrase(text: str) -> bool:
    """块内是否含 HIT 学位专属短语（封面信息表 / 英文科研成果标题等）。"""
    return any(p in text for p in DEGREE_PHRASES)


def _is_decl_start(p_elem) -> bool:
    t = _para_text(p_elem).strip()
    return any(m in t for m in DECL_START_MARKERS)


def _decl_replace(text: str) -> str:
    """声明区「研究生」→「本科生」，保护机构名「研究生院」。"""
    return (text.replace(INSTITUTION_PROTECT, "\x00")
                .replace("研究生", "本科生")
                .replace("\x00", INSTITUTION_PROTECT))


def adapt_declaration_block(elem) -> None:
    """声明区逐跑替换「研究生」→「本科生」，保护机构名。"""
    for r in block_runs(elem):
        if r.text and "研究生" in r.text:
            r.text = _decl_replace(r.text)


def _is_stop(p_elem) -> bool:
    style_id = ""
    ppr = p_elem.find(w("pPr"))
    if ppr is not None:
        pstyle = ppr.find(w("pStyle"))
        if pstyle is not None:
            style_id = pstyle.get(w("val")) or ""
    if style_id.startswith("HitHeading1") or style_id.startswith("Heading1"):
        return True
    return _para_text(p_elem).strip() in STOP_TEXTS


def _adapt_document(xml_bytes: bytes, zh: str, en_cap: str, en_low: str) -> bytes:
    root = etree.fromstring(xml_bytes)
    body = root.find(w("body"))
    if body is None:
        return xml_bytes
    in_cover = True
    in_declaration = False
    for child in list(body):
        tag = etree.QName(child).localname
        if tag == "p":
            if in_cover and _is_stop(child):
                in_cover = False
            if _is_decl_start(child):
                in_declaration = True
            if in_cover:
                adapt_block(child, zh, en_cap, en_low)
            elif _contains_degree_phrase(_para_text(child)):
                # 封面信息表等落在 摘要 之后，仍要改写学位字样
                adapt_block(child, zh, en_cap, en_low)
            elif in_declaration:
                # 声明区「研究生」→「本科生」，保护机构名「研究生院」
                adapt_declaration_block(child)
        elif tag == "tbl":
            if in_cover:
                adapt_block(child, zh, en_cap, en_low)
            elif _contains_degree_phrase(_tbl_text(child)):
                adapt_block(child, zh, en_cap, en_low)
        # sectPr 等其他元素跳过
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _adapt_part(xml_bytes: bytes, zh: str, en_cap: str, en_low: str) -> bytes:
    root = etree.fromstring(xml_bytes)
    adapt_block(root, zh, en_cap, en_low)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def main() -> int:
    ap = argparse.ArgumentParser(
        description="学位层次适配：硕士→本科/博士，master→bachelor/doctor（封面/页眉/页脚/后置小节标题）")
    ap.add_argument("docx", help="输入 docx（硕士格式生成物）")
    ap.add_argument("--degree", choices=["bachelor", "master", "doctor"], default="master")
    ap.add_argument("--out", help="输出 docx（默认覆盖输入）")
    args = ap.parse_args()

    if args.degree == "master":
        print("[degree-adapt] degree=master，无需适配，原样保留。")
        return 0

    zh, en_cap, en_low = ZH[args.degree], EN_CAP[args.degree], EN_LOW[args.degree]

    src = zipfile.ZipFile(args.docx, "r")
    out_path = args.out or args.docx
    tmp_names = src.namelist()
    header_re = re.compile(r"word/header\d*\.xml$")
    footer_re = re.compile(r"word/footer\d*\.xml$")

    fd, tmp_path = tempfile.mkstemp(suffix=".docx", dir=os.path.dirname(out_path) or ".")
    os.close(fd)
    with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as out:
        for name in tmp_names:
            data = src.read(name)
            if name == "word/document.xml":
                data = _adapt_document(data, zh, en_cap, en_low)
            elif header_re.search(name) or footer_re.search(name):
                data = _adapt_part(data, zh, en_cap, en_low)
            out.writestr(name, data)
    src.close()
    os.replace(tmp_path, out_path)
    print(f"[degree-adapt] degree={args.degree}：封面/页眉/页脚/后置小节标题 已适配为"
          f"「{zh}」（master→{en_low}）。输出 {out_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
