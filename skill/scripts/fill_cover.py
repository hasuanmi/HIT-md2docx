#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fill_cover.py — 哈工大论文封面表格字段补填（生成后处理）

背景：
  cover_inject.py 只替换「题目」，封面表格的其余字段（作者 / 导师 / 申请学位 /
  学科 / 所在单位 / 答辩日期 + 英文 6 字段）在官方模板里是空的。本脚本按
  「标签去空格匹配 + 末单元格填值」从 front_matter md 把这些已知字段填进去。

  ⚠️ 真实姓名 / 导师姓名 在 front_matter 里是「待填写 / To Be Filled」占位符，
     本脚本如实填入占位符，由学生最终手动替换为真实信息。

用法：
  python fill_cover.py <论文>.docx --front-matter <front_matter>.md \
      [--degree master|bachelor|doctor] [--out <out>.docx]

标签为模板真实文本（codepoint 校验）：
  中文：硕士研究生 / 导 师 / 申请学位 / 学 科 / 所在单位 / 答辩日期 / 授予学位单位
  英文：Candidate / Supervisor / Academic Degree Applied for / Speciality /
        Affiliation / Date of Defence / Degree-Conferring-Institution
"""
import argparse
import os
import re
import sys
import zipfile
import shutil
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass


def norm(s: str) -> str:
    """去掉空白与冒号，便于标签匹配。"""
    return re.sub(r"[\s：:]", "", s or "")


def parse_front_matter(path: str) -> dict:
    """解析 front_matter md：'键：值' 形式。"""
    out = {}
    if not path or not os.path.isfile(path):
        return out
    for line in open(path, encoding="utf-8"):
        line = line.rstrip("\n")
        if "：" in line or ":" in line:
            sep = "：" if "：" in line else ":"
            k, v = line.split(sep, 1)
            out[k.strip()] = v.strip()
    return out


def build_label_map(fm: dict, degree: str) -> dict:
    """front_matter 字典 -> 规范化标签 -> 值。"""
    author_label = "硕士研究生" if degree == "master" else (
        "博士研究生" if degree == "doctor" else "本科生")
    # 学位类别词
    deg_word = fm.get("学位类别", "")
    eng_deg = fm.get("英文学位", "")

    cn = {
        author_label: fm.get("作者", ""),
        "导师": fm.get("指导教师", ""),
        "申请学位": deg_word,
        "学科": fm.get("学科专业", ""),
        "所在单位": fm.get("所在单位", ""),
        "答辩日期": fm.get("答辩日期", ""),
        "授予学位单位": "哈尔滨工业大学",
    }
    en = {
        "Candidate": fm.get("英文作者", ""),
        "Supervisor": fm.get("英文导师", ""),
        "AcademicDegreeAppliedfor": eng_deg,
        "Speciality": fm.get("英文学科", ""),
        "Affiliation": fm.get("英文单位", ""),
        "DateofDefence": fm.get("英文答辩日期", ""),
        "DegreeConferringInstitution": "Harbin Institute of Technology",
    }
    # 合并（英文键已去空格）
    m = {}
    for k, v in cn.items():
        m[norm(k)] = v
    for k, v in en.items():
        m[norm(k)] = v
    return m


def set_cell_text(tc, value: str):
    """把 value 写到单元格末段落的首个 run（保留已有 rPr 字体）。"""
    if value is None:
        return
    # 找首个有 <w:t> 的 run，复用其字体；否则新建
    runs = tc.findall(f"{{{W}}}p/{{{W}}}r")
    target_t = None
    for r in runs:
        t = r.find(f"{{{W}}}t")
        if t is not None:
            target_t = t
            break
    if target_t is not None:
        target_t.text = value
        return
    # 没有现成 <w:t>：建一个 run 追加到首个段落
    p = tc.find(f"{{{W}}}p")
    if p is None:
        p = etree.SubElement(tc, f"{{{W}}}p")
    r = etree.SubElement(p, f"{{{W}}}r")
    t = etree.SubElement(r, f"{{{W}}}t")
    # xml:space 保留空格
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = value


def fill(docx_path: str, fm: dict, degree: str, out_path: str):
    label_map = build_label_map(fm, degree)
    tmp = docx_path + ".filltmp"
    with zipfile.ZipFile(docx_path, "r") as zin:
        names = zin.namelist()
        data = {n: zin.read(n) for n in names}

    xml_bytes = data["word/document.xml"]
    root = etree.fromstring(xml_bytes)

    filled = []
    for tbl in root.iter(f"{{{W}}}tbl"):
        for row in tbl.iter(f"{{{W}}}tr"):
            cells = row.findall(f"{{{W}}}tc")
            if not cells:
                continue
            label_cell = cells[0]
            label = "".join(t.text or "" for t in label_cell.iter(f"{{{W}}}t"))
            key = norm(label)
            if key in label_map:
                value = label_map[key]
                if value:
                    last_cell = cells[-1]
                    set_cell_text(last_cell, value)
                    filled.append((label.strip(), value))

    data["word/document.xml"] = etree.tostring(root, xml_declaration=True,
                                                encoding="UTF-8", standalone=True)
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            zout.writestr(n, data[n])
    shutil.move(tmp, out_path)
    return filled


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("docx")
    ap.add_argument("--front-matter", default=None)
    ap.add_argument("--degree", default="master",
                    choices=["master", "bachelor", "doctor"])
    ap.add_argument("--out", default=None)
    args = ap.parse_args()

    out = args.out or args.docx
    fm = parse_front_matter(args.front_matter) if args.front_matter else {}
    filled = fill(args.docx, fm, args.degree, out)
    if filled:
        print(f"[fill_cover] 已补填 {len(filled)} 个封面字段：")
        for k, v in filled:
            print(f"  - {k} -> {v}")
    else:
        print("[fill_cover] 未匹配到任何封面字段（请检查 front_matter 或模板标签）")


if __name__ == "__main__":
    main()
