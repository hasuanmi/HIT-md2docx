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

import math

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
    cn = en = en_lines = None
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
# 题目替换
# ---------------------------------------------------------------------------
def _replace_title_text(para, new_text):
    """替换段落文本，并清空其余 run，避免模板残留的格式说明/乱码符号被保留。

    只改文本，绝不改字号/字体/间距（题目字号须严格保持模板规范：2号字）。
    """
    runs = para.runs
    if not runs:
        return False
    runs[0].text = new_text
    for r in runs[1:]:
        r.text = ""
    return True


def _split_en(en, maxc=56):
    """把英文题目按模板两行结构拆分。

    如果整题能在 56 字符以内放下，则只放第一行，第二行留空；
    否则优先在第一行结束时断开，剩余全部放入第二行。
    """
    if len(en) <= maxc:
        return en, ""
    words = en.split()
    if len(words) <= 1:
        return en, ""
    line1 = words[0]
    j = 1
    while j < len(words) and len(line1) + 1 + len(words[j]) <= maxc:
        line1 += " " + words[j]
        j += 1
    return line1, " ".join(words[j:])


# 封面固定标签（非题目），用于从模板中识别并排除，避免误把标签当题目替换
_COVER_FIXED_LABELS = (
    "哈尔滨工业大学", "本科学位论文", "硕士学位论文", "博士学位论文",
    "国内图书分类号", "国际图书分类号", "学校代码", "密级",
    "Classified Index", "U.D.C", "Dissertation for the Master Degree",
    "Candidate", "Supervisor", "Academic Degree Applied for", "Speciality",
    "Affiliation", "Date of Defence", "Degree-Conferring", "授予学位单位",
    "所在单位", "申请学位", "导师", "答辩日期", "Harbin Institute of Technology",
    "（Times New Roman", "Times New Roman", "2号", "小2号", "加粗", "居中",
    "年 月", "年月", "答辩日期",
)


def _is_cover_label(t: str) -> bool:
    return any(k in t for k in _COVER_FIXED_LABELS)


def _para_is_bold_title_style(p) -> bool:
    """判断段落是否符合封面题目（中文/英文）的格式特征：至少有一个非白色 run 加粗，
    且字号 >= 18pt（小二号）= 342000 EMU。"""
    for r in p.runs:
        if not r.text:
            continue
        # 跳过白色批注 run
        rPr = r._r.find(Q("rPr"))
        if rPr is not None:
            c = rPr.find(Q("color"))
            if c is not None and c.get(Q("val"), "").upper() == "FFFFFF":
                continue
        if r.bold:
            sz = r.font.size
            if sz is not None and sz >= 342000:  # 18pt
                return True
    return False


def _cjk_len(t: str) -> int:
    return sum(1 for c in t if "一" <= c <= "鿿")


def _en_len(t: str) -> int:
    return sum(1 for c in t if c.isascii() and c.isalpha())


# ---------------------------------------------------------------------------
# 封面第一页排版守恒：题目变长后压缩空行，保证"哈尔滨工业大学 / 年 月"留在第一页底部
#
# 思路（不改字号，严格保持模板 2 号题字规范）：
#   模板用一串空行把校名/年月块顶到第一页底部。题目越长、换行越多，就把校名块
#   往下挤到第二页。因此按"题目相对模板原题多出的视觉行数"，等比例删掉题目区与
#   校名块之间的空行即可把校名块拉回第一页底部。
#   - 题目行(2号=22pt) 行高 ≈ 27.5pt；空行(正文五号=10.5pt) 行高 ≈ 12.6pt。
#   - 每多 1 行题目 ≈ 需删 27.5/12.6 ≈ 2.2 个空行来补偿。
# ---------------------------------------------------------------------------
_TITLE_LINE_PT = 27.5      # 2号题目单行高度(近似)
_EMPTY_LINE_PT = 12.6      # 五号空行单行高度(近似)
_TITLE_FONT_PT = 22.0      # 2号
_EN_CHAR_W_RATIO = 0.5     # 英文字符平均宽 ≈ 0.5×字号


def _cover_usable_width_pt(cover: Document) -> float:
    """封面页面可用文字宽度(pt) = 页宽 - 左右页边距。"""
    body = cover.element.body
    for sect in body.iter(Q("sectPr")):
        pgSz = sect.find(Q("pgSz"))
        pgMar = sect.find(Q("pgMar"))
        if pgSz is not None and pgMar is not None:
            try:
                w = int(pgSz.get(Q("w")))
                l = int(pgMar.get(Q("left")))
                r = int(pgMar.get(Q("right")))
                return (w - l - r) / 20.0  # twips -> pt
            except (TypeError, ValueError):
                pass
    return 425.0  # A4 默认(11906-1701-1701)/20


def _text_width_units(t: str) -> float:
    """估算文本"全角宽度单位"：中日韩/全角标点=1.0，其余(含 ASCII)=0.5。"""
    u = 0.0
    fullwidth_punct = "《》【】「」『』（）—…·、，。；：！？"
    for c in t:
        if ("\u4e00" <= c <= "\u9fff") or ("\u3000" <= c <= "\u303f") \
                or ("\uff00" <= c <= "\uffef") or c in fullwidth_punct:
            u += 1.0
        else:
            u += 0.5
    return u


def _cjk_title_lines(text: str, usable_pt: float) -> int:
    """中文题目在封面题目区的视觉行数。"""
    if not text:
        return 0
    units = _text_width_units(text)
    per_line = max(1.0, usable_pt / _TITLE_FONT_PT)
    return max(1, math.ceil(units / per_line))


def _en_title_lines(text: str, usable_pt: float) -> int:
    """英文题目在封面题目区的视觉行数。"""
    n = len((text or "").strip())
    if not n:
        return 0
    per_line = max(1.0, usable_pt / (_TITLE_FONT_PT * _EN_CHAR_W_RATIO))
    return max(1, math.ceil(n / per_line))


def compress_cover_gap(cover: Document, remove_n: int, keep_min: int = 3) -> int:
    """删除"哈尔滨工业大学"上方紧邻的若干空行，把校名/年月块拉回第一页底部。

    只删除紧邻校名块上方的连续空段落（题目区与校名块之间的排版留白），
    保留至少 keep_min 个空行以维持底部间距；从最靠近校名块的一侧开始删。
    """
    if remove_n <= 0:
        return 0
    paras = cover.paragraphs
    anchor = None
    for i, p in enumerate(paras):
        if "哈尔滨工业大学" in p.text:
            anchor = i
            break
    if anchor is None:
        return 0
    empties = []
    j = anchor - 1
    while j >= 0 and paras[j].text.strip() == "":
        empties.append(paras[j])
        j -= 1
    removable = max(0, len(empties) - keep_min)
    to_remove = min(remove_n, removable)
    removed = 0
    for k in range(to_remove):
        el = empties[k]._p
        el.getparent().remove(el)
        removed += 1
    return removed


def replace_cover_titles(cover: Document, cn, en):
    """把模板封面的中/英文题目替换为论文题目。

    定位策略基于「题目 = 封面中长度达标的非固定标签文本段落」：
      - 中文题目：替换所有 CJK 长度 >= 8、非固定标签、非格式说明的段落；
      - 英文题目：替换所有纯英文长度 >= 20、非固定标签、非格式说明的段落，
        并把其后续连续英文续行清空或填入第二行。
    模板底部的格式说明（如「年 月、Times New Roman...」）已被 _COVER_FIXED_LABELS
    排除，避免被误替换。
    """
    if not cn and not en:
        return 0
    paras = cover.paragraphs
    n = 0

    # —— 替换前先记录模板原题目，用于估算题目变长后多出的视觉行数 ——
    usable_pt = _cover_usable_width_pt(cover)
    orig_cn = None
    for p in paras:
        t = p.text.strip()
        if not t or _is_cover_label(t) or t.lstrip().startswith("（"):
            continue
        if _cjk_len(t) >= 8:
            orig_cn = t
            break
    orig_en_parts = []
    for i0 in range(len(paras)):
        t = paras[i0].text.strip()
        if t and not _is_cover_label(t) and _en_len(t) >= 20 and _cjk_len(t) == 0:
            orig_en_parts.append(t)
            k = i0 + 1
            while k < len(paras):
                tt = paras[k].text.strip()
                if not tt:
                    k += 1
                    continue
                if _is_cover_label(tt):
                    break
                if _en_len(tt) > 0 and _cjk_len(tt) == 0:
                    orig_en_parts.append(tt)
                    k += 1
                    continue
                break
            break
    orig_en = " ".join(orig_en_parts)

    if cn:
        for p in paras:
            t = p.text.strip()
            if not t or _is_cover_label(t) or t.lstrip().startswith("（"):
                continue
            if _cjk_len(t) >= 8:
                if _replace_title_text(p, cn):
                    n += 1

    if en:
        en1, en2 = _split_en(en)
        i = 0
        while i < len(paras):
            p = paras[i]
            t = p.text.strip()
            if not t or _is_cover_label(t):
                i += 1
                continue
            if _en_len(t) >= 20 and _cjk_len(t) == 0:
                # 命中英文标题第一行
                if _replace_title_text(p, en1):
                    n += 1
                # 处理续行：下一个连续英文段落作为第二行
                j = i + 1
                second_filled = False
                while j < len(paras):
                    tt = paras[j].text.strip()
                    if not tt:
                        j += 1
                        continue
                    if _is_cover_label(tt):
                        break
                    if _en_len(tt) > 0 and _cjk_len(tt) == 0:
                        if en2 and not second_filled:
                            if _replace_title_text(paras[j], en2):
                                n += 1
                            second_filled = True
                        else:
                            for r in paras[j].runs:
                                r.text = ""
                        j += 1
                        continue
                    # 遇到非英文格式说明行，停止清空
                    break
                i = j
            else:
                i += 1

    # —— 题目变长 → 等比例压缩校名块上方空行，把"哈尔滨工业大学 / 年 月"拉回第一页底部 ——
    base_lines = _cjk_title_lines(orig_cn or "", usable_pt) + _en_title_lines(orig_en, usable_pt)
    new_lines = _cjk_title_lines(cn or orig_cn or "", usable_pt) + _en_title_lines(en or orig_en, usable_pt)
    extra_lines = new_lines - base_lines
    if extra_lines > 0:
        remove_n = round(extra_lines * (_TITLE_LINE_PT / _EMPTY_LINE_PT))
        removed = compress_cover_gap(cover, remove_n)
        print(f"   题目多出 {extra_lines} 行 → 压缩封面空行 {removed} 个（保校名块于首页底部）")

    return n


# ---------------------------------------------------------------------------
# 封面区元素构建 & 注入
# ---------------------------------------------------------------------------
def build_cover_elems(cover: Document):
    """返回封面模板 body 下所有子元素的深拷贝，并保留其原始分节结构。

    封面模板 ``input/封面.docx`` 的 body 末尾会带一个 ``w:sectPr``（body sectPr）。
    在目标文档中，这个 sectPr 若作为 body 子元素直接插入到封面与正文之间，会
    破坏分节逻辑，导致英文封面与摘要之间的空白页丢失或分节错位。

    这里把它从 body 子元素中取出，作为 inline sectPr 挂到封面区最后一个段落
    （即空白页段落）的 pPr 上：
      - 封面区最后一个 inline sectPr（模板中已有的英文封面节结束）会先产生新页，
        进入空白页；
      - 移到这个末段的 body sectPr 再产生新页，使摘要从新页开始；
      - 这样英文封面与摘要之间自然夹着一页空白页。
    """
    cover_body = cover.element.body
    children = list(cover_body)
    elems = [copy.deepcopy(c) for c in children]

    # 把模板末尾的 body sectPr 取下来，准备挂到封面区最后一个段落
    last_sectPr = None
    if elems and etree.QName(elems[-1]).localname == "sectPr":
        last_sectPr = elems.pop()

    # 确保封面区最后一个元素是段落（模板正常是末尾的空白页段落）
    if not elems or etree.QName(elems[-1]).localname != "p":
        blank = OxmlElement("w:p")
        elems.append(blank)
    last_p = elems[-1]

    if last_sectPr is not None:
        pPr = last_p.find("{%s}pPr" % W)
        if pPr is None:
            pPr = OxmlElement("w:pPr")
            last_p.insert(0, pPr)
        # 避免和已有 sectPr 冲突：先移除已有的 inline sectPr
        old = pPr.find("{%s}sectPr" % W)
        if old is not None:
            pPr.remove(old)
        pPr.append(last_sectPr)

    return elems


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

    # ---- 安全网：清理残留的引擎内置前置封面表 ----
    # 引擎可能在段落式封面之外还生成了表格型前置封面（含
    # 「硕士研究生：」「Candidate：Supervisor：」等字段）。
    # 这些表格位于注入的官方模板和正文摘要之间，需要额外清除。
    # 策略：移除 body 下所有位于 摘要段落之前的表格（正文数据表只会在摘要之后）。
    n_rm = 0
    _abstract_seen = False
    for child in list(body):
        tag = etree.QName(child).localname
        if tag == "p":
            txt = "".join(child.itertext()).strip()
            if "摘" in txt and "要" in txt and "Abstract" not in txt:
                _abstract_seen = True
        if _abstract_seen:
            break
        if tag == "tbl":
            body.remove(child)
            n_rm += 1
    if n_rm:
        print(f"   清理 {n_rm} 个残留引擎前置封面表")

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
