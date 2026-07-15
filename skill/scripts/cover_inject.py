#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
cover_inject.py — HIT 硕士论文"封面与前置部分版式注入"（标准后处理步骤，不改动引擎）

职责（全部依赖外部官方模板 封面.docx，引擎本身不生成这种精确版式）：
  1. 整页复制 封面.docx 的封面区（含 3 页：中文封面 / 中文扉页表格 / 英文封面表格 +
     分页符 + 分节 sectPr）到论文开头，替换其中的中/英文题目为你论文的题目，
     白色批注箭头原样保留。
  2. 在中/英文目录的章节条目之前，插入 `摘  要 → I`、`Abstract → II`（罗马数字页码）。
  3. 对齐前置部分（封面 + 摘要页）行距到官方《书写范例》：摘要标题段距、摘要正文行距。

题目来源（按优先级）：
  a) 命令行 --cn-title / --en-title
  b) front_matter（默认 input/front_matter_hit.md 的"论文题目："/"英文题目："）
  c) 从输入 docx 现有封面区提取（自包含兜底）

健壮性：
  - 找不到 封面.docx 模板 → 跳过封面复制，仅做目录/行距，不报错（流程可无模板运行）。
  - 幂等：每次都基于模板重新注入封面区（先删旧封面区）；目录条目/行距检测已存在则跳过。

用法：
  python cover_inject.py <thesis.docx> [--cover-template input/封面.docx]
                              [--front-matter input/front_matter_hit.md]
                              [--out output/<论文>.docx] [--no-spacing]
"""
from __future__ import annotations
import argparse
import copy
import os
import re
import sys

sys.stdout.reconfigure(encoding="utf-8")

from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
Q = lambda t: f"{{{W}}}{t}"

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ENGINE_ROOT = os.path.dirname(os.path.dirname(SCRIPT_DIR))   # HITmd2docx/ (scripts -> skill -> HITmd2docx)
DEFAULT_COVER = os.path.join(ENGINE_ROOT, "input", "封面.docx")
DEFAULT_FRONT = os.path.join(ENGINE_ROOT, "input", "front_matter_hit.md")


# ---------------------------------------------------------------------------
# 题目提取
# ---------------------------------------------------------------------------
def _read_front_titles(front_path: str):
    cn = en = None
    if front_path and os.path.exists(front_path):
        with open(front_path, encoding="utf-8") as f:
            for line in f:
                s = line.strip()
                if s.startswith("论文题目：") or s.startswith("论文题目:"):
                    cn = s.split("：", 1)[-1].split(":", 1)[-1].strip()
                elif s.startswith("英文题目：") or s.startswith("英文题目:"):
                    en = s.split("：", 1)[-1].split(":", 1)[-1].strip()
    return cn, en


def _extract_titles_from_doc(doc: Document):
    """从 docx 现有封面区（开头到第一个摘要）提取中/英文题目，作兜底。"""
    cn = en_lines = None
    cn_cands, en_cands = [], []
    FIXED = ("硕士学位论文", "学术学位", "哈尔滨工业大学", "国内图书分类号",
             "国际图书分类号", "学校代码", "密级", "Classified", "U.D.C",
             "Dissertation", "Candidate", "Supervisor", "Academic Degree",
             "Speciality", "Affiliation", "Date of Defence",
             "Degree-Conferring", "授予学位", "所在单位", "答辩日期", "申请学位", "导师")
    for p in doc.paragraphs:
        t = p.text.strip()
        if not t:
            continue
        if "摘" in t and "要" in t and "Abstract" not in t:
            break  # 到达摘要，封面区结束
        if any(k in t for k in FIXED):
            continue
        if re.search(r"[一-鿿]", t) and len(t) >= 8 and not t.startswith("（"):
            cn_cands.append(t)
        if re.match(r"^[A-Z][a-z]", t) and ("Resilience" in t or "Research" in t
                                            or "Convergence" in t or "Impact" in t):
            en_cands.append(t)
    if cn_cands:
        cn = cn_cands[0]
    if en_cands:
        en = " ".join(en_cands).replace("↑", "").strip()
    return cn, en


def extract_titles(front_path, doc, cli_cn=None, cli_en=None):
    cn, en = _read_front_titles(front_path)
    if cli_cn:
        cn = cli_cn
    if cli_en:
        en = cli_en
    if not cn or not en:
        dcn, den = _extract_titles_from_doc(doc)
        cn = cn or dcn
        en = en or den
    return cn, en


# ---------------------------------------------------------------------------
# 题目替换（保留白色批注）
# ---------------------------------------------------------------------------
def _replace_title_text(para, new_text):
    runs = para.runs
    black = []
    for i, r in enumerate(runs):
        rPr = r._r.find(Q("rPr"))
        color = None
        if rPr is not None:
            c = rPr.find(Q("color"))
            if c is not None:
                color = c.get(Q("val"))
        if color and color.upper() == "FFFFFF":
            continue
        black.append(i)
    if not black:
        return False
    runs[black[0]].text = new_text
    for i in black[1:]:
        runs[i].text = ""
    return True


def _split_en(en, maxc=56):
    """把英文题目断成恰好 2 行（第二行含剩余所有词，绝不丢字）。

    优先在复合词 'Industrial Technology' 后断开（本论文题目前半段结构）；
    否则贪心到 maxc，剩余全部放入第二行。模板备注：题目过长可用小 2 号字。
    """
    marker = "Industrial Technology"
    if marker in en:
        i = en.index(marker) + len(marker)
        return en[:i].strip(), en[i:].strip()
    words = en.split()
    if len(words) <= 1:
        return en, ""
    line1 = words[0]
    j = 1
    while j < len(words) and len(line1) + 1 + len(words[j]) <= maxc:
        line1 += " " + words[j]
        j += 1
    return line1, " ".join(words[j:])


def replace_cover_titles(cover: Document, cn, en):
    if not cn and not en:
        return 0
    paras = cover.paragraphs
    en1, en2 = _split_en(en) if en else (None, None)
    n = 0
    for i, p in enumerate(paras):
        t = p.text.strip()
        if not t:
            continue
        if cn and ("网络互动" in t or "女性弱势群体" in t or "抗逆力" in t):
            if _replace_title_text(p, cn):
                n += 1
            continue
        if en and t.upper().startswith("RESEARCH") and (
                "PROMOTION" in t.upper() or "VULNERABLE" in t.upper()
                or "FEMALE" in t.upper()):
            if _replace_title_text(p, en1):
                j = i + 1
                while j < len(paras) and not paras[j].text.strip():
                    j += 1
                if j < len(paras):
                    _replace_title_text(paras[j], en2)
                n += 1
            continue
    return n


# ---------------------------------------------------------------------------
# 封面区元素构建 & 注入
# ---------------------------------------------------------------------------
def build_cover_elems(cover: Document):
    cover_body = cover.element.body
    children = list(cover_body)
    if etree.QName(children[-1]).localname == "sectPr":
        body_sect = children[-1]
        last_para = None
        for child in reversed(children[:-1]):
            if etree.QName(child).localname == "p" and "".join(child.itertext()).strip():
                last_para = child
                break
        if last_para is not None:
            pPr = last_para.find(Q("pPr"))
            if pPr is None:
                pPr = OxmlElement("w:pPr")
                last_para.insert(0, pPr)
            if pPr.find(Q("sectPr")) is None:
                pPr.append(copy.deepcopy(body_sect))
        cover_body.remove(body_sect)
        children = list(cover_body)
    return [copy.deepcopy(c) for c in children]


def inject_cover(target: Document, cover_elems):
    body = target.element.body
    cut = None
    for child in body:
        if etree.QName(child).localname == "p":
            txt = "".join(child.itertext()).strip()
            if "摘" in txt and "要" in txt and "Abstract" not in txt and "ABSTRACT" not in txt.upper():
                cut = child
                break
    if cut is None:
        print("  [跳过封面替换] 未找到摘要标题，不删除封面区")
        return False
    to_remove = []
    for child in body:
        if child is cut:
            break
        to_remove.append(child)
    for c in to_remove:
        body.remove(c)
    for el in reversed(cover_elems):
        body.insert(0, el)
    return True


# ---------------------------------------------------------------------------
# 目录摘要条目
# ---------------------------------------------------------------------------
def _find_toc_title(paras, labels):
    for i, p in enumerate(paras):
        t = p.text.strip().replace(" ", "")
        if t in labels:
            return i
    return None


def _first_chapter(paras, start):
    for i in range(start, len(paras)):
        t = paras[i].text.strip()
        if re.match(r"^第.+章", t) or re.match(r"^\d+\s", t) or re.match(r"^[一二三四五六七八九十]+、", t):
            return i
    return None


def _first_chapter_en(paras, start):
    for i in range(start, len(paras)):
        if paras[i].text.strip().startswith("Chapter"):
            return i
    return None


def _has_entry(paras, start, end, text):
    for i in range(start, end):
        if paras[i].text.strip().startswith(text):
            return True
    return False


def _make_entry_para(before_para, text, page):
    p = OxmlElement("w:p")
    pPr = before_para._p.find(Q("pPr"))
    if pPr is not None:
        p.append(copy.deepcopy(pPr))
    r1 = OxmlElement("w:r")
    t1 = OxmlElement("w:t")
    t1.set(Q("space"), "preserve")
    t1.text = text
    r1.append(t1)
    rt = OxmlElement("w:r")
    rt.append(OxmlElement("w:tab"))
    r2 = OxmlElement("w:r")
    t2 = OxmlElement("w:t")
    t2.set(Q("space"), "preserve")
    t2.text = page
    r2.append(t2)
    p.append(r1)
    p.append(rt)
    p.append(r2)
    return p


def _insert_toc_entries(target, before_para, entries):
    body = target.element.body
    bp = before_para._p
    idx = list(body).index(bp)
    for text, page in reversed(entries):
        p = _make_entry_para(before_para, text, page)
        body.insert(idx, p)


def _is_toc_para(p) -> bool:
    """判断段落是否为目录条目（带 TOC1/2/3 样式或分页字段）。"""
    pPr = p._p.find(Q("pPr"))
    if pPr is None:
        return False
    ps = pPr.find(Q("pStyle"))
    if ps is not None:
        style_id = ps.get(Q("val")) or ""
        if style_id.upper().startswith("TOC"):
            return True
    # 含 TOC 分页字段也视为目录条目
    for r in p._p.findall(Q("r")):
        for fc in r.iter(Q("fldChar")):
            return True
    return False


def ensure_toc_abstract(target: Document):
    paras = target.paragraphs
    n = 0
    ci = _find_toc_title(paras, ("目  录", "目录", "目 录"))
    if ci is not None:
        fc = _first_chapter(paras, ci + 1)
        if fc is not None and not _has_entry(paras, ci, fc, "摘  要"):
            _insert_toc_entries(target, paras[fc], [("摘  要", "I"), ("Abstract", "II")])
            n += 2
    # 英文目录（Contents）由引擎直接生成 Abstract(In Chinese) / Abstract(In English)，
    # 不再插入 摘 要/Abstract 中文标签，避免重复。
    return n


# ---------------------------------------------------------------------------
# 前置部分行距对齐（官方书写范例）
# ---------------------------------------------------------------------------
def _set_para_spacing(p, before=None, after=None, line=None, rule=None):
    pPr = p._p.get_or_add_pPr()
    sp = pPr.find(Q("spacing"))
    if sp is None:
        sp = OxmlElement("w:spacing")
        pPr.append(sp)
    if before is not None:
        sp.set(Q("before"), str(int(before)))
    if after is not None:
        sp.set(Q("after"), str(int(after)))
    if line is not None:
        sp.set(Q("line"), str(int(line)))
    if rule is not None:
        sp.set(Q("lineRule"), rule)


def align_abstract_spacing(target: Document):
    """对齐摘要页标题段距与摘要正文行距到书写范例。幂等。"""
    paras = target.paragraphs
    # 找正文区（摘要目录之后）的"摘  要"标题
    toc_seen = False
    abs_title = None
    for p in paras:
        t = p.text.strip().replace(" ", "")
        if t in ("目  录", "目录", "Contents"):
            toc_seen = True
            continue
        if toc_seen and "摘要" in t and "Abstract" not in t:
            if _is_toc_para(p):
                continue  # 跳过目录条目，找真正的摘要标题
            abs_title = p
            break
    if abs_title is None:
        return 0
    # 摘要标题段距
    n = 0
    cur = abs_title
    before_val = cur._p.find(Q("pPr"))
    need = True
    if before_val is not None:
        sp = before_val.find(Q("spacing"))
        if sp is not None and sp.get(Q("before")) == "391" and sp.get(Q("after")) == "312":
            need = False
    if need:
        _set_para_spacing(abs_title, before=391, after=312)
        n += 1
    # 摘要正文（标题之后到 关键词/Keywords 之前）：中文 1.2 倍(288)、英文 1.5 倍(360)
    body_start = False
    for p in paras[paras.index(abs_title) + 1:]:
        t = p.text.strip()
        if not t:
            continue
        if t.startswith("关键词") or t.startswith("Keywords") or t.startswith("Abstract"):
            break
        pPr = p._p.get_or_add_pPr()
        sp = pPr.find(Q("spacing"))
        cur_line = sp.get(Q("line")) if sp is not None else None
        # 摘要正文统一 1.5 倍行距（360），与正文/标题一致；不再区分中英文。
        target_line = 360
        if cur_line != str(target_line):
            _set_para_spacing(p, line=target_line, rule="auto")
            n += 1
    return n


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="HIT 硕士论文封面与前置部分版式注入")
    ap.add_argument("docx", help="引擎生成的论文 docx")
    ap.add_argument("--cover-template", default=DEFAULT_COVER, help="官方封面模板 docx")
    ap.add_argument("--front-matter", default=DEFAULT_FRONT, help="front_matter_hit.md 路径")
    ap.add_argument("--cn-title", default=None, help="中文题目（优先于 front_matter）")
    ap.add_argument("--en-title", default=None, help="英文题目（优先于 front_matter）")
    ap.add_argument("--out", default=None, help="输出路径，默认覆盖原文件")
    ap.add_argument("--no-spacing", action="store_true", help="跳过前置部分行距对齐")
    args = ap.parse_args()

    target = Document(args.docx)
    cn, en = extract_titles(args.front_matter, target, args.cn_title, args.en_title)
    print(f"题目: 中文={cn!r} 英文={en!r}")

    cover_changed = False
    if os.path.exists(args.cover_template):
        print(f"1. 读取封面模板: {args.cover_template}")
        cover = Document(args.cover_template)
        nn = replace_cover_titles(cover, cn, en)
        print(f"   题目替换 {nn} 处")
        elems = build_cover_elems(cover)
        ok = inject_cover(target, elems)
        cover_changed = ok
        print(f"   封面区注入: {'成功' if ok else '跳过'}")
    else:
        print(f"[警告] 未找到封面模板 {args.cover_template}，跳过封面复制（仅做目录/行距）")

    print("2. 目录摘要条目 ...")
    n_toc = ensure_toc_abstract(target)
    print(f"   插入 {n_toc} 条")

    if not args.no_spacing:
        print("3. 前置部分行距对齐 ...")
        n_sp = align_abstract_spacing(target)
        print(f"   调整 {n_sp} 处")

    out = args.out or args.docx
    target.save(out)
    print(f"保存: {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
