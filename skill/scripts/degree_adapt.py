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

# 机构名保护：声明区「研究生」改写时，绝不能把「深圳研究生院」等
# 机构名误改（本科→深圳本科生院 / 博士→深圳博士生院）。
INSTITUTION_PROTECT = "研究生院"

# 声明区通用「研究生」按目标学位改写（不能写死为「本科生」，否则博士论文会被误改）
DECL_GRAD_MAP = {"本科": "本科生", "博士": "博士生", "硕士": "研究生"}


def w(tag: str) -> str:
    return f"{{{W}}}{tag}"


def find_spans(logical: str, zh: str, en_cap: str, en_low: str):
    """扫描 logical，返回有序、不重叠的 (start, end, replacement) 列表。

    - 中文优先匹配长词「硕士研究生」(5字)，再「硕士」(2字)；
      硕士→zh；硕士研究生→「本科生」/「博士研究生」/「硕士研究生」。
      （注意：本科封面官方模板写的是「本科生」而非「本科生研究生」）
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
            rep = {"本科": "本科生", "博士": "博士研究生",
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
    """把替换结果写回各 <w:t> 跑，保留未被替换部分的原有字形。

    原实现同时递增 logical 位置 p 和 run 内位置 k，在 replacement 长度与被替换
    长度不一致（例如"硕士研究生"→"本科生"）时会导致 p 与后续 run 错位，出现
    "本科生士研究生"等残留。新版按"每个 run 负责输出其原始 logical 区间对应的
    字符"重新分配，遇到 span 起点时把完整 rep 写入该 run，并跳过 span 其余位置。
    """
    if not spans:
        return
    n = len(logical)
    sorted_spans = sorted(spans, key=lambda x: x[0])

    # 标记被替换位置，并记录每个 span 起点对应的替换文本
    replaced = [False] * n
    replacements = {}
    for s, e, rep in sorted_spans:
        for p in range(s, e):
            replaced[p] = True
        replacements[s] = rep

    # 每个 run 在原始 logical 字符串中的起止位置
    run_start = []
    pos = 0
    for r in runs:
        run_start.append(pos)
        pos += len(r.text or "")

    out = [[] for _ in runs]
    p = 0
    for j, r in enumerate(runs):
        t = r.text or ""
        start = run_start[j]
        end = start + len(t)
        # 如果前面因 span 替换产生跳跃，p 可能已超过当前 run 起点，直接跳过
        if p < start:
            p = start
        while p < end and p < n:
            if p in replacements:
                out[j].extend(list(replacements[p]))
                # 跳过整个 span 覆盖的原始位置
                p = next((e for s, e, _ in sorted_spans if s == p), p + 1)
            else:
                out[j].append(logical[p])
                p += 1

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


def _decl_replace(text: str, target: str) -> str:
    """声明区「研究生」→ 目标学位对应词（本科→本科生 / 博士→博士生），保护机构名「研究生院」。"""
    rep = DECL_GRAD_MAP[target]
    return (text.replace(INSTITUTION_PROTECT, "\x00")
                .replace("研究生", rep)
                .replace("\x00", INSTITUTION_PROTECT))


def adapt_declaration_block(elem, target: str) -> None:
    """声明区逐跑替换「研究生」→ 目标学位对应词，保护机构名。"""
    for r in block_runs(elem):
        if r.text and "研究生" in r.text:
            r.text = _decl_replace(r.text, target)


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
                # 声明区「研究生」→ 目标学位对应词（本科→本科生 / 博士→博士生），保护机构名「研究生院」
                adapt_declaration_block(child, zh)
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


# ---------------------------------------------------------------------------
# 引擎内置封面检测（cover_inject 失败守卫）
# ---------------------------------------------------------------------------
_ENGINE_TABLE_MARKERS = [
    "硕士研究生：导",       # 中文表格封面特征行（旧版 profile）
    "Candidate：Supervisor：",  # 英文表格封面特征行（旧版 profile）
    "申请学位：",             # 引擎 frontmatter 表格字段
    "Academic Degree Applied for：",
]

_ENGINE_PARA_MARKERS = [
    "硕士研究生：【请填写】",   # 段落型封面占位符（当前 hit-master-thesis）
    "【请填写】",
    "【Please Fill】",
]

def _has_engine_builtin_cover(docx_path: str) -> bool:
    """检测 document.xml 是否包含引擎内置的未替换封面。

    引擎生成的封面有两种形式：
      1. 段落型（当前 hit-master-thesis profile）：含「【请填写】」/「【Please Fill】」
         占位符，如「硕士研究生：【请填写】」、「Candidate：【Please Fill】」。
      2. 表格型（旧版 profile）：含「硕士研究生：导师」等未替換字段名的表格。

    官方 封面.docx 注入后，这些占位符会被 replace_cover_titles 替换为实际论文
    题目和作者信息。若检测到占位符仍存在，说明 cover_inject 未能成功执行。

    返回 True 表示 cover_inject 很可能未成功，degree_adapt 不应在错误封面上改字。
    """
    try:
        with zipfile.ZipFile(docx_path, "r") as zf:
            data = zf.read("word/document.xml")
    except (zipfile.BadZipFile, KeyError):
        return False
    root = etree.fromstring(data)

    # 注意：官方封面模板的表格本身含「硕士研究生：导」「Candidate：Supervisor：」
    # 等字样，这些并非失败信号；唯一可靠的 cover_inject 失败信号是引擎生成的
    # 未填占位符「【请填写】/【Please Fill】」，因此只对段落占位符做检测，
    # 避免把正确的官方模板表格误判为「引擎内置封面」而阻断层次适配。

    # 检测段落型当前封面（检查前 40 个段落的文本）—— 仅以未填占位符为信号
    paras = list(root.iter(w("p")))[:50]
    for p in paras:
        text = "".join(p.itertext())
        for marker in _ENGINE_PARA_MARKERS:
            if marker in text:
                return True

    return False


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

    # ---- 守卫：检测引擎内置封面（cover_inject 失败）----
    if _has_engine_builtin_cover(args.docx):
        print("[degree-adapt] ❌ 错误：检测到文档中仍存在引擎内置的未替换封面！", file=sys.stderr)
        print("  这说明 cover_inject.py（封面模板注入）未能成功替换掉引擎生成的错误封面。", file=sys.stderr)
        print(f"  如果在错误的封面上执行 degree={args.degree} 文字替换，会产生乱码封面。", file=sys.stderr)
        print("  请检查：", file=sys.stderr)
        print("    1. input/封面.docx 是否存在且未损坏", file=sys.stderr)
        print("    2. 上一步 [2/3] 是否显示「封面区注入: 成功」", file=sys.stderr)
        print("    3. 是否使用了 generate.py（而非直接跑 md2docx 命令）", file=sys.stderr)
        return 1

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
