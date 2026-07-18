#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""md_to_fordocx.py — HIT 硕士论文 Markdown 预处理（生成 *_for_docx.md）。

背景
----
thesis_md2docx 的 HIT 硕士 profile 解析器只把 `# 数字`（如 `# 1 绪论`）视为
正文起点，且标题样式自带“第一章 / 1.1”自动编号。若源 md 用中文章号写法
（`# 第一章 绪论`），解析器无法识别正文起点，会把整篇正文当作前置内容丢弃。

因此正确工作流是：先在 md 层把中文章号归一为数字章号，产出 *_for_docx.md，
再喂给引擎。本脚本只做这一步纯文本预处理，不改动任何引擎代码，零回归。

“固定部分/可变部分”分离（--front-matter）
------------------------------------------
论文格式天然分两类：
  1. 固定部分——封面信息（`## 封面信息`）、原创性声明（`## 声明`）等，几乎与
     模板一致、每次生成不变，用户只需维护一份（见 input/front_matter_hit.md）。
  2. 可变部分——摘要、目录、正文等，随论文内容变化。
很多同学的正文 md 里只写题目 + 摘要 + 正文，没有 `## 封面信息` / `## 声明` 块。
此时引擎读不到封面数据 → 英文题目变空行、各字段回退成占位符（封面“生成不对”）。

本脚本在预处理阶段自动补齐固定部分（md 层注入，不碰引擎）：
  - 若正文 md 缺 `## 封面信息`，则把固定前置模板注入到题目行之后、正文之前；
  - 固定模板里的 `论文题目：` 会用正文 H1 题目自动同步，无需手改两处；
  - 若正文 md 已含 `## 封面信息`，则原样保留，不重复注入。

转换规则（最小改动）
--------------------
- 仅转换一级章标题：`# 第X章 标题` → `# N 标题`（X 为中文/阿拉伯数字，N 为阿拉伯数字）。
- 其余一切不动：论文标题、`# 参考文献`、`# 附录 A`、`# 后置内容`、
  前置 `## 摘要` / `## Abstract`、节 `## 1.1`、小节 `### 1.1.1` 全部原样保留。

用法
----
    python md_to_fordocx.py <输入.md> [-o <输出.md>] [--front-matter <固定前置.md>]
不指定 -o 时，输出到同目录下的 <名>_for_docx.md。
不指定 --front-matter 时，自动在脚本同级 input/ 或输入文件同目录寻找
front_matter_hit.md；找到且正文缺封面信息则注入，否则跳过（可用 --no-front-matter 关闭）。
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

_CN_DIGITS = {"零": 0, "一": 1, "二": 2, "两": 2, "三": 3, "四": 4,
              "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}


def cn_to_int(s: str) -> int:
    """将中文数字（一/十二/廿…）或阿拉伯数字字符串转为 int。"""
    s = s.strip()
    if s.isdigit():
        return int(s)
    if "十" not in s:
        return sum(_CN_DIGITS.get(ch, 0) for ch in s)
    if s == "十":
        return 10
    if s.startswith("十"):
        return 10 + cn_to_int(s[1:])
    if s.endswith("十"):
        return cn_to_int(s[:-1]) * 10
    left, right = s.split("十", 1)
    return cn_to_int(left) * 10 + (cn_to_int(right) if right else 0)


_CHAPTER = re.compile(r"^#\s+第([一二三四五六七八九十零\d]+)章(?:\s+(.*))?$")

# 末端（后置）章节：在哈工大硕士论文中应为顶层一级标题（`#`），不应作为
# 正文章节的二级子标题。常见错误是把 `## 参考文献` / `## 附录 A` 写在
# `# N 结论` 章之下，导致：① 目录里参考文献被当成结论的子节；② 校验器
# 找不到顶层“参考文献”标题而误报。预处理阶段把这类二级写法提升为一级。
_BACK_MATTER = re.compile(
    r"^##\s+(参考文献|致谢|个人简历|附录.*|攻读.{0,24}?(成果|科研成果))\s*$"
)

# 正文 H1 题目：`# 标题`（排除 `# 数字…` / `# 第X章…` / `## …`）
_H1_TITLE = re.compile(r"^#\s+(?!第[一二三四五六七八九十零\d]+章)(?!\d)(.+)$")
# 固定前置里的“论文题目：xxx”行，用于与正文 H1 同步
_FM_TITLE_LINE = re.compile(r"^(论文题目\s*[:：]).*$")

# 这些一级标题不是论文题目，extract_title 必须跳过，否则会把摘要误当题目
_TITLE_SKIP = {"摘要", "ABSTRACT", "Abstract"}


def extract_title(text: str) -> str:
    """从正文 md 取第一处 H1 题目（非数字章、非二级、非 摘要/ABSTRACT）。取不到返回空串。"""
    for line in text.splitlines():
        m = _H1_TITLE.match(line.strip())
        if m:
            cand = m.group(1).strip()
            if cand in _TITLE_SKIP:
                continue
            return cand
    return ""


def has_cover_info(text: str) -> bool:
    """正文是否已含 `## 封面信息` 块（含则不再注入固定前置）。"""
    for line in text.splitlines():
        s = line.strip()
        if s.startswith("## ") and "封面信息" in s:
            return True
    return False


def inject_front_matter(text: str, front_matter: str) -> str:
    """把固定前置（封面信息+声明）注入到正文 H1 题目行之后、其余内容之前。

    - 固定前置里的“论文题目：”会用正文 H1 题目同步；
    - 若正文没有 H1，则把固定前置放在最前面。
    """
    title = extract_title(text)
    # 同步固定前置里的“论文题目：”
    if title:
        fm_lines = []
        for line in front_matter.splitlines():
            if _FM_TITLE_LINE.match(line.strip()):
                fm_lines.append(f"论文题目：{title}")
            else:
                fm_lines.append(line)
        front_matter = "\n".join(fm_lines)
    fm_block = front_matter.strip("\n")

    lines = text.splitlines()
    insert_at = None
    for i, line in enumerate(lines):
        if _H1_TITLE.match(line.strip()):
            insert_at = i + 1
            break
    if insert_at is None:
        # 正文无 H1 题目（如用户把 `# 摘要` 写在最前）：把固定前置置顶，并尝试
        # 从封面信息里的「论文题目：」生成 `# 题目` 一级标题，使引擎拿到正确的
        # 封面题目（与封面信息一致），而不是把 摘要 误判为题目。
        m = re.search(r"论文题目\s*[:：]\s*(.+)", front_matter)
        title_val = m.group(1).strip() if m else ""
        head = ["# " + title_val, ""] if title_val else []
        tail_nl = "\n" if text.endswith("\n") else ""
        return "\n".join(head + [fm_block, "", text]) + tail_nl
    head = lines[:insert_at]
    tail = lines[insert_at:]
    merged = head + ["", fm_block, ""] + tail
    tail_nl = "\n" if text.endswith("\n") else ""
    return "\n".join(merged) + tail_nl


def convert(text: str) -> tuple[str, int]:
    """把 `# 第X章 标题` 归一为 `# N 标题`。返回 (新文本, 转换的章数)。"""
    out: list[str] = []
    n_chapters = 0
    for line in text.splitlines():
        m = _CHAPTER.match(line)
        if m:
            num = cn_to_int(m.group(1))
            title = (m.group(2) or "").strip()
            out.append(f"# {num} {title}".rstrip())
            n_chapters += 1
        else:
            out.append(line)
    tail = "\n" if text.endswith("\n") else ""
    return "\n".join(out) + tail, n_chapters


def promote_back_matter(text: str) -> tuple[str, int]:
    """把误写为二级（`##`）的末端章节提升为顶层一级（`#`）。

    覆盖：参考文献、致谢、个人简历、附录（含 附录 A/B）、攻读…成果。
    这些节在哈工大硕士论文里都是顶层一级标题，规范写法见 HIT.md 的
    `# 参考文献`。仅当某行完全匹配时才改写，不影响 `## 3.1` 等正文子节。
    返回 (新文本, 提升的节数)。
    """
    out: list[str] = []
    n = 0
    for line in text.splitlines():
        m = _BACK_MATTER.match(line)
        if m:
            title = m.group(1).strip()
            out.append("# " + title)
            n += 1
        else:
            out.append(line)
    tail = "\n" if text.endswith("\n") else ""
    return "\n".join(out) + tail, n


# 单行展示公式 `$$ 公式 $$`：thesis_md2docx 的公式收集器只识别“独占一行的
# `$$` 块”，单行 `$$...$$` 会被误当作行内 `$...$` 解析，导致 LaTeX 串尾多出一个
# `$`、转换器产出非法 OMML（7 个公式转换失败的根因）。预处理阶段把整行的
# `$$ 公式 $$` 归一为多行块形式，零引擎改动即可让公式正确转成 Word 公式。
_DISPLAY_MATH_SINGLELINE = re.compile(r"^\s*\$\$(.+?)\$\$\s*$", re.M)


def normalize_display_math(text: str) -> tuple[str, int]:
    """把整行的 `$$ 公式 $$` 归一为多行 `$$` 块（引擎已支持的形式）。

    仅匹配“整行仅由 `$$…$$` 构成”的行，不触碰行内 `$x$` 或本就是多行的
    `$$` 块。返回 (新文本, 归一的公式数)。
    """
    count = 0

    def _sub(m: re.Match[str]) -> str:
        nonlocal count
        count += 1
        # 闭 `$$` 后补一个换行，避免连续展示公式的 `$$` 被拼成 `$$$$` 而合并成一项
        return "$$\n" + m.group(1).strip() + "\n$$\n"

    new_text = _DISPLAY_MATH_SINGLELINE.sub(_sub, text)
    return new_text, count


# 章前分隔 horizontal rule：解析器把 `---` / `***` 当成分页符（PageBreakBlock），
# 若它紧贴一级标题 `# ...`（章/参考文献等）之前，会与引擎在一级标题处自行插入
# 的“分节符+分页”叠加，导致 `append_chapter_page_break` 的 `need_break` 守卫判定
# “上一段已是分页”而跳过本节分节符——该章“首页章名页眉”的 rId 永远写不进 sectPr
# （仅最后一章靠文档末尾 sectPr 兜底）。预处理层删掉章前 `---`，零引擎改动即修复
# 全部章名页眉。引擎对一级标题本就会分节+分页，故删除冗余 `---` 不影响排版。
_HR = re.compile(r"^\s*(-\s*){3,}\s*$|^\s*(\*\s*){3,}\s*$")
_H1_LINE = re.compile(r"^#\s+\S")


def strip_hr_before_heading(text: str) -> tuple[str, int]:
    """删掉紧贴一级标题 `# ...` 之前的 horizontal rule（`---`/`***`）。

    仅当 hr 与一级标题之间只隔空白行时才删除（即“章前分隔线”语义），
    不影响正文内部的 `---`（其后不是一级标题）。返回 (新文本, 删除条数)。
    """
    lines = text.splitlines()
    out: list[str] = []
    n = 0
    i = 0
    while i < len(lines):
        line = lines[i]
        if _HR.match(line) and not line.lstrip().startswith("#"):
            j = i + 1
            while j < len(lines) and lines[j].strip() == "":
                j += 1
            if j < len(lines) and _H1_LINE.match(lines[j]):
                n += 1
                # 删掉 hr 行 + 其与一级标题之间的空行；保留 hr 之前的空行不变
                i = j
                continue
        out.append(line)
        i += 1
    tail = "\n" if text.endswith("\n") else ""
    return "\n".join(out) + tail, n


# 摘要/ABSTRACT 误写为一级标题 → 降级为二级
# ---------------------------------------------------------------------------
# 引擎只把 `## 摘要` 当作摘要小节。若用户把 `# 摘要` 写成一级标题：
#   - extract_title 会把它误判为论文题目（封面题目变成「摘要」）；
#   - 其后的摘要正文被当成正文/标题区，摘要整段丢失。
# 预处理阶段把 `# 摘要` / `# ABSTRACT` / `# Abstract` 降为 `##`，零引擎改动
# 即让摘要正确归位。已为 `##` 的不受影响（正则只匹配单个 `#`）。
_ABSTRACT_H1 = re.compile(r"^#\s+(摘要|ABSTRACT|Abstract)\s*$", re.IGNORECASE)


def normalize_abstract_headings(text: str) -> tuple[str, int]:
    """把误写为一级标题的 `# 摘要`/`# ABSTRACT`/`# Abstract` 降为二级 `##`。

    返回 (新文本, 降级处数)。幂等：已为二级的不重复处理。
    """
    out: list[str] = []
    n = 0
    for line in text.splitlines():
        if _ABSTRACT_H1.match(line.strip()):
            out.append("## " + line.lstrip("#").strip())
            n += 1
        else:
            out.append(line)
    tail = "\n" if text.endswith("\n") else ""
    return "\n".join(out) + tail, n


# 参考文献按正文首次引用顺序重排（GB/T 7714 顺序编码制）
# ---------------------------------------------------------------------------
# 仅匹配“纯数字”引用标记，避免误伤 [图 1-1] / [NQP] 等含字母或中文的方括号。
_CITE_MARKER = re.compile(r"\[\s*(\d+(?:\s*[-,，]\s*\d+)*)\s*\]")
_REF_ENTRY = re.compile(r"^\[\s*(\d+)\s*\]\s*(.*)$")
_REF_HEADING = re.compile(r"^#\s+参考文献\s*$")


def sort_references_by_citation(text: str) -> tuple[str, int]:
    """按正文首次引用顺序重排参考文献列表（GB/T 7714 顺序编码制），并重编号。

    - 扫描参考文献之前的全部 [N] / [N,N] / [N-N] 标记，按首次出现顺序得到引用序列；
    - 未被引用的条目保留在末尾（按原相对顺序）；
    - 重编号：被引条目 1..k，未被引 k+1..，同时改写正文标记与列表条目的编号。

    幂等：若正文引用顺序与列表编号顺序已一致，则原样返回（重排数 0）。
    返回 (新文本, 实际重排/重编号的条目数)。
    """
    lines = text.splitlines()
    ref_idx = None
    for i, ln in enumerate(lines):
        if _REF_HEADING.match(ln.strip()):
            ref_idx = i
            break
    if ref_idx is None:
        return text, 0

    body_lines = lines[:ref_idx]
    ref_lines = lines[ref_idx:]
    body_text = "\n".join(body_lines)

    # 1) 正文引用首次出现顺序（展开逗号/连字符范围）
    order: list[int] = []
    seen: set[int] = set()
    for m in _CITE_MARKER.finditer(body_text):
        for ns in re.split(r"\s*[-,，]\s*", m.group(1)):
            ns = ns.strip()
            if ns.isdigit():
                n = int(ns)
                if n not in seen:
                    seen.add(n)
                    order.append(n)
    if not order:
        return text, 0

    # 2) 解析参考文献条目
    entries = [(int(em.group(1)), em.group(2)) for ln in ref_lines
               if (em := _REF_ENTRY.match(ln))]
    if not entries:
        return text, 0
    by_old = {num: body for num, body in entries}

    # 3) 构建新顺序与编号映射
    cited = [n for n in order if n in by_old]
    uncited = [n for n, _ in entries if n not in seen]
    new_order = cited + uncited
    old_to_new = {old: new for new, old in enumerate(new_order, start=1)}
    new_to_old = {v: k for k, v in old_to_new.items()}

    # 4) 改写正文标记编号（仅替换方括号内的数字，保留分隔符）
    def _remap_inner(inner: str) -> str:
        return re.sub(r"\d+",
                      lambda mm: str(old_to_new.get(int(mm.group()), int(mm.group()))),
                      inner)
    new_body = _CITE_MARKER.sub(lambda m: "[" + _remap_inner(m.group(1)) + "]", body_text)

    # 5) 重排并重编号参考文献条目（按新号 1..N 升序输出）
    #    条目之间插入空行，确保 markdown 解析器把每条文献拆成独立段落，
    #    引擎才会为每条生成独立的 HitReference 段落与书签（否则会被合并成
    #    一个段落，仅首条有书签，正文 [2]–[N] 超链接指向失效书签）。
    heading = ref_lines[0] if ref_lines else "# 参考文献"
    new_entries = []
    for new_num in range(1, len(new_order) + 1):
        old_num = new_to_old[new_num]
        old_body = by_old[old_num]
        new_entries.append(f"[{new_num}] {old_body}")
    new_text = new_body + "\n" + heading + "\n\n" + "\n\n".join(new_entries)
    tail = "\n" if text.endswith("\n") else ""
    return new_text + tail, len(new_order)


# ---------------------------------------------------------------------------
# 直引号 → 中文引号（中文论文排版规范）
# ---------------------------------------------------------------------------
def normalize_quotes(text: str) -> tuple[str, int]:
    """把正文 ASCII 直引号 "（U+0022）配对替换为中文引号 “ ”（U+201C/U+201D）。

    中文论文正文须用中文引号；源稿常因输入法/编辑器原因使用半角直引号。
    本函数在 md 预处理层统一规范化（零引擎改动）。

    保护范围（内部引号保持原样、不被替换）：
      - ``` 围栏代码块 ```、行内 `code`
      - $...$ 行内公式 与 $$...$$ 块级公式（LaTeX 内引号/语义符号不动）
    配对：按段落（空行分隔）独立配对，段内首个 " → “（左），下一个 → ”（右），
    交替进行，段落边界重置。已存在的中文引号不重复处理（幂等）。
    不处理单引号 '，避免破坏英文缩写（don't / it's）。
    仅处理「# 参考文献」之前的正文（与引用转换边界一致），文献列表保持原样。
    """
    lines = text.splitlines()
    ref_idx = None
    for i, ln in enumerate(lines):
        if re.match(r"^#\s+参考文献\s*$", ln.strip()):
            ref_idx = i
            break
    if ref_idx is not None:
        body_text = "\n".join(lines[:ref_idx])
        ref_text = "\n".join(lines[ref_idx:])
        new_body, n = _normalize_quotes_body(body_text)
        return new_body + "\n" + ref_text, n
    return _normalize_quotes_body(text)


def _normalize_quotes_body(body: str) -> tuple[str, int]:
    protected = []
    def _protect(m):
        protected.append(m.group(0))
        return f"\x00Q{len(protected) - 1}\x00"

    # 1) 暂存需保护的区域
    body = re.sub(r"```.*?```", _protect, body, flags=re.S)   # 围栏代码块
    body = re.sub(r"`[^`\n]+`", _protect, body)               # 行内 code
    body = re.sub(r"\$\$.*?\$\$", _protect, body, flags=re.S) # 块级公式
    body = re.sub(r"(?<![\$])\$(?!\$)[^$\n]+?\$(?!\$)", _protect, body)  # 行内公式

    # 2) 按段落配对替换直引号
    paragraphs = re.split(r"\n{2,}", body)
    out_paras = []
    cnt = 0
    for para in paragraphs:
        buf = []
        open_q = False
        for ch in para:
            if ch == '"':
                buf.append("\u201c" if not open_q else "\u201d")
                open_q = not open_q
                cnt += 1
            else:
                buf.append(ch)
        out_paras.append("".join(buf))
    body = "\n\n".join(out_paras)

    # 3) 还原保护区域
    for i, seg in enumerate(protected):
        body = body.replace(f"\x00Q{i}\x00", seg)
    return body, cnt


# ---------------------------------------------------------------------------
# 作者-年份制 → GB/T 7714 顺序编码制（双语鲁棒版）
# ---------------------------------------------------------------------------
# 作用：把正文中的 （作者，年份）/ 作者（年份）/ Author (year) / 中文（English，year）
# 等著者-出版年引文，按「首次出现顺序」编号为 [1][2]…，并把参考文献列表重排、
# 重编号，使其与正文顺序一致。引擎随后会把正文 [N] 渲染为右上角上标、把列表
# 条目做成可点击书签。未匹配到参考文献条目的引文（如法律、方法说明、或正文
# 引用但文献表缺失）保持原样，列入报告供人工核对。
_CN_CONNECTORS = ["而", "但", "且", "该", "上述", "这些", "此", "其", "这一", "这种",
                  "那种", "后", "前", "内", "上", "下", "外", "中", "各", "这", "那",
                  "本", "同"]


def _norm_author(a: str) -> str:
    return re.sub(r"[\s,，、．・·.]+", "", a.lower().strip())


def _strip_cn(s: str) -> str:
    s = s.strip()
    changed = True
    while changed:
        changed = False
        for c in _CN_CONNECTORS:
            if s.startswith(c):
                s = s[len(c):]
                changed = True
                break
    s = re.sub(r"^的+", "", s)
    s = re.sub(r"等$", "", s)
    s = re.sub(r"[，,][\u4e00-\u9fff]+$", "", s)
    s = re.sub(r"、[\u4e00-\u9fff]+$", "", s)
    s = re.sub(r"和[\u4e00-\u9fff]+$", "", s)
    return s.strip()


def _authors_before(pre: str) -> list[tuple[str, str]]:
    """从括号前的文本尾部提取著者候选（first-author 优先）。

    兼容多种"作者（年份）"写法，专治旧逻辑的两个盲区：
      - 英文作者 + 等（年）：`Gao等（2023）` / `Aghion等（2020）`
        （括号内只有年份，作者在括号外、被"等"隔开）
      - 中文双作者 A和B（年）：`吴超鹏和唐菂（2016）`
        （旧逻辑抓到最近的第二作者"唐菂"，此处按连接符拆出首作者"吴超鹏"）

    仅抓取"紧贴括号、由著者连接符(和/与/、/，/,/&/and/空格)连成的著者串"，
    避免误吸正文普通词；返回 [(author, kind), ...]，kind ∈ {"cn","en"}，
    英文 token 按出现序、中文按分段给出，供上层逐一匹配参考文献表。
    """
    s = pre.rstrip()
    # 去掉尾部的 "等 / 团队 / 课题组 / et al." 等群体后缀
    s = re.sub(r"(?:等|团队|课题组)$", "", s)
    s = re.sub(r"(?:et\s+al\.?)$", "", s, flags=re.I).rstrip()
    # 尾部著者串：中英文名，以 和/与/、/，/,/&/and/空格 连接，锚定到末尾
    m = re.search(
        r"((?:[A-Z][a-zA-Z]+|[\u4e00-\u9fff]{2,4})"
        r"(?:\s*(?:和|与|、|，|,|&|and)\s*(?:[A-Z][a-zA-Z]+|[\u4e00-\u9fff]{2,4}))*)$",
        s,
    )
    if not m:
        return []
    chunk = m.group(1)
    cands: list[tuple[str, str]] = []
    for t in re.findall(r"[A-Z][a-zA-Z]+", chunk):
        cands.append((t, "en"))
    for seg in re.split(r"\s*(?:和|与|、|，|,|&|and)\s*", chunk):
        if re.match(r"[\u4e00-\u9fff]", seg):
            seg2 = _strip_cn(seg)
            if len(seg2) >= 2:
                cands.append((seg2, "cn"))
    return cands


def convert_author_year_to_gbt7714(text: str) -> tuple[str, int]:
    """作者-年份制正文 → GB/T 7714 顺序编码制（[N] 上标 + 按首现重排）。

    返回 (新文本, 实际转换的引文处数)。幂等：若正文本无著者-出版年引文则返回原样。
    """
    lines = text.splitlines()
    ref_idx = None
    for i, ln in enumerate(lines):
        if re.match(r"^#\s+参考文献\s*$", ln.strip()):
            ref_idx = i
            break
    if ref_idx is None:
        return text, 0

    body_lines = lines[:ref_idx]
    ref_lines = lines[ref_idx:]
    body_text = "\n".join(body_lines)

    # 1) 解析参考文献列表，建双语键（中文首作者 / 英文首作者, year）
    ref_entries = []
    for ln in ref_lines:
        m = re.match(r"^\[(\d+[a-z]?)\]\s*(.*)$", ln)
        if m:
            bodytext = m.group(2)
            # 收集条目内出现的全部年份：出版年可能不在最前（标题里也可能含年份，
            # 如 "...the 2021 antitrust crackdown...[J]. Journal, 2022"）。取第一个
            # 作主 year（向后兼容），同时保留全部年份供匹配时任一命中即可。
            years = re.findall(r"(?:19|20)\d{2}", bodytext)
            year = years[0] if years else None
            if re.match(r"[\u4e00-\u9fff]", bodytext):
                fa = re.match(r"([\u4e00-\u9fff]{2,4})", bodytext)
                fa = fa.group(1) if fa else ""
            else:
                fa_m = re.match(r"([A-Za-z]+)", bodytext)
                fa = fa_m.group(1) if fa_m else ""
            ref_entries.append({"body": bodytext, "fa": fa,
                                "year": year, "years": years})
    if not ref_entries:
        return text, 0

    # 一个文献条目在其【全部年份】下都建键：正文引用年份命中任一即算匹配，
    # 解决"标题含其他年份导致出版年被首年份覆盖"的误配（如 Haque 2022）。
    ref_keys: dict[tuple[str, str], list[int]] = {}
    for i, e in enumerate(ref_entries):
        if e["fa"] and e["years"]:
            for y in dict.fromkeys(e["years"]):
                ref_keys.setdefault((_norm_author(e["fa"]), y), []).append(i)

    # 2) 检测所有含年份的括号引文
    repl = []
    for m in re.finditer(r"[（(]([^）)]*?(?:19|20)\d{2}[^）)]*?)[）)]", body_text):
        inner = m.group(1)
        ym = re.search(r"(?:19|20)\d{2}", inner)
        if not ym:
            continue
        year = ym.group(0)
        pairs = []
        # (a) 括号前的著者串：Gao等（2023）/ 吴超鹏和唐菂（2016）/ Katz & Shapiro …
        #     由 _authors_before 统一处理英文名、中文首作者、多作者连接等盲区。
        for (a, kind) in _authors_before(body_text[:m.start()]):
            pairs.append((a, year, kind))
        # (b) 括号内的著者：（作者，2023）/（Katz & Shapiro，1985）
        for en in re.findall(r"[A-Z][a-zA-Z]+", inner):
            pairs.append((en, year, "en"))
        # 括号内中文著者：可能是 "A和B，year"（花合凤和谢莉娟，2022），
        # 在连接符处切出首作者，避免把"花合凤和"整体当作者名。
        cn_inside = re.search(r"^([\u4e00-\u9fff]{2,}(?:[和与、，][\u4e00-\u9fff]{2,})*)", inner)
        if cn_inside:
            ci = _strip_cn(re.split(r"[和与、，]", cn_inside.group(1))[0])
            if len(ci) >= 2:
                pairs.append((ci, year, "cn"))
        # 去重保序（同一著者可能被括号内外重复捕获）
        seen_pair = set()
        uniq_pairs = []
        for p in pairs:
            key = (_norm_author(p[0]), p[1], p[2])
            if key not in seen_pair:
                seen_pair.add(key)
                uniq_pairs.append(p)
        if uniq_pairs:
            repl.append((m.start(), m.end(), uniq_pairs))

    repl.sort(key=lambda x: x[0])
    cleaned = []
    for r in repl:
        if cleaned and r[0] < cleaned[-1][1]:
            continue
        cleaned.append(r)
    repl = cleaned
    if not repl:
        return text, 0

    # 3) 映射（中文名优先，英文名兜底；取首个命中参考文献条目的候选）
    ref_to_num: dict[int, int] = {}
    next_num = 1
    repl_num = []
    for (s, e, pairs) in repl:
        nums = []
        chosen_ak = pairs[0][0]
        chosen_yr = pairs[0][1]
        for (ak, year, kind) in pairs:
            if kind == "en":
                chosen_ak, chosen_yr = ak, year
                break
        matched = False
        for (ak, year, kind) in pairs:
            cand = ref_keys.get((_norm_author(ak), year), [])
            if not cand:
                continue
            chosen = next((c for c in cand if c not in ref_to_num), cand[0])
            if chosen not in ref_to_num:
                ref_to_num[chosen] = next_num
                next_num += 1
            nums.append(ref_to_num[chosen])
            matched = True
            break
        repl_num.append((s, e, nums if matched else None))

    # 4) 替换正文
    result = []
    cursor = 0
    for (s, e, nums) in repl_num:
        result.append(body_text[cursor:s])
        if nums is None:
            result.append(body_text[s:e])
        elif len(nums) == 1:
            result.append(f"[{nums[0]}]")
        else:
            result.append("[" + ", ".join(str(n) for n in nums) + "]")
        cursor = e
    result.append(body_text[cursor:])
    new_body = "".join(result)

    # 5) 重排参考文献（被引者按首现顺序 1..k，未引者续接）
    uncited = [i for i in range(len(ref_entries)) if i not in ref_to_num]
    for i in uncited:
        ref_to_num[i] = next_num
        next_num += 1
    num_entry = sorted((ref_to_num[i], ref_entries[i]) for i in range(len(ref_entries)))
    new_ref_lines = ["# 参考文献", ""]
    for num, e in num_entry:
        new_ref_lines.append(f"[{num}] {e['body']}")
        new_ref_lines.append("")
    tail = "\n" if text.endswith("\n") else ""
    return new_body.rstrip("\n") + "\n\n" + "\n".join(new_ref_lines) + tail, len(repl)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="HIT 硕士论文 Markdown 预处理：中文章号 → 数字章号，生成 *_for_docx.md")
    parser.add_argument("input", help="输入 Markdown 文件路径")
    parser.add_argument("-o", "--output", help="输出路径（默认 <名>_for_docx.md）")
    parser.add_argument("--front-matter", help="固定前置模板（封面信息+声明）路径")
    parser.add_argument("--no-front-matter", action="store_true",
                        help="关闭固定前置自动注入")
    args = parser.parse_args(argv)

    src = Path(args.input)
    if not src.exists():
        print(f"[error] 输入文件不存在: {src}", file=sys.stderr)
        return 1

    text = src.read_text(encoding="utf-8")

    # 摘要/ABSTRACT 若被误写成一级标题，先降级为二级（避免被误当论文题目、摘要丢失）
    text, a = normalize_abstract_headings(text)

    # ---- 公式预处理：纯文本计量公式 → $$...$$ 块 ----
    # 引擎的 MathConverter（temml → mathml2omml）只识别 $$...$$ 展示公式块，
    # 源稿里大量公式是“RD_it = α₀ + …  (1)”这种纯文本写法，若不转换就会被
    # 当成普通段落、以纯文本渲染（即用户反馈的“公式没有被提取出来”）。
    # 这里调用引擎同款的 preprocess_equations，把公式转成 $$ 块，编号 (N) 保留
    # 在 latex 末尾交由引擎提取为右对齐公式号。
    _here = Path(__file__).resolve()
    _engine_scripts_candidates = [
        _here.parent.parent.parent / "scripts",   # HIT-md2docx/skill/scripts → HIT-md2docx/scripts
        _here.parent / "scripts",                 # 兜底：同目录
    ]
    _engine_scripts = next((p for p in _engine_scripts_candidates if p.exists()), None)
    if _engine_scripts is not None and str(_engine_scripts) not in sys.path:
        sys.path.insert(0, str(_engine_scripts))
    convert_equations = None
    try:
        from preprocess_equations import convert_text as convert_equations  # noqa: E402
    except Exception as _e:  # pragma: no cover - 依赖缺失时跳过，不阻断主流程
        print(f"[warn] 公式预处理器不可用，跳过公式提取: {_e}", file=sys.stderr)

    # ---- 固定前置注入（仅当正文缺 ## 封面信息 时）----
    injected = False
    if not args.no_front_matter:
        if has_cover_info(text):
            print("[skip] 正文已含 `## 封面信息`，不注入固定前置")
        else:
            fm_path = None
            if args.front_matter:
                fm_path = Path(args.front_matter)
            else:
                # 自动定位固定前置模板，兼容两种目录布局：
                #  - 规范仓库: HIT-md2docx/skill/scripts/ → ../../../input/
                #  - 注册专家包: .../hitmd2docx/scripts/ → ../input/
                here = Path(__file__).resolve()
                candidates = [
                    here.parent.parent.parent / "input" / "front_matter_hit.md",
                    here.parent.parent / "input" / "front_matter_hit.md",
                    here.parent / "front_matter_hit.md",
                    src.with_name("front_matter_hit.md"),
                    src.parent / "input" / "front_matter_hit.md",
                ]
                fm_path = next((p for p in candidates if p.exists()), None)
            if fm_path and fm_path.exists():
                fm_text = fm_path.read_text(encoding="utf-8")
                text = inject_front_matter(text, fm_text)
                injected = True
                print(f"[ok] 已注入固定前置（封面信息+声明）: {fm_path}")
            else:
                print("[warn] 未找到固定前置模板，封面字段将回退占位符；"
                      "可用 --front-matter 指定或在正文补 `## 封面信息` 块")

    converted, n = convert(text)

    # 公式预处理：纯文本计量公式 → $$...$$ 块（编号 (N) 保留，供引擎右对齐提取）
    eq_count = 0
    if convert_equations is not None:
        before_math = converted.count("$$")
        converted = convert_equations(converted)
        eq_count = (converted.count("$$") - before_math) // 2

    converted, m = promote_back_matter(converted)
    converted, k = normalize_display_math(converted)
    converted, qn = normalize_quotes(converted)
    converted, g = convert_author_year_to_gbt7714(converted)
    converted, r = sort_references_by_citation(converted)
    converted, h = strip_hr_before_heading(converted)

    if args.output:
        dst = Path(args.output)
    else:
        dst = src.with_name(f"{src.stem}_for_docx{src.suffix}")
    dst.write_text(converted, encoding="utf-8")

    print(f"[ok] 转换 {n} 个章标题（中文章号→数字章号）")
    if a:
        print(f"[ok] 降级 {a} 个误写的一级摘要标题（# 摘要→## 摘要，避免被误当题目）")
    if eq_count:
        print(f"[ok] 提取 {eq_count} 个纯文本计量公式 → $$...$$ 展示公式块（引擎转 OMML 原生公式）")
    if m:
        print(f"[ok] 提升 {m} 个末端章节为顶层一级标题（参考文献/附录/成果等）")
    if k:
        print(f"[ok] 归一 {k} 个单行展示公式 `$$ 公式 $$` → 多行块（修复 OMML 转换失败）")
    if g:
        print(f"[ok] 作者-年份引文 → GB/T 7714 顺序编码 [N] 上标并据首现重排文献 {g} 处")
    if qn:
        print(f"[ok] 规范化 {qn} 个半角直引号 → 中文引号“ ”（正文，参考文献列表保持原样）")
    if r:
        print(f"[ok] 按正文首次引用顺序重排并重编号 {r} 条参考文献（GB/T 7714 顺序编码制）")
    if h:
        print(f"[ok] 删除 {h} 处章前分隔线 `---`（避免与引擎分节符叠加导致章名页眉丢失）")
    if injected:
        print("[ok] 固定/可变两部分已在 md 层合并（引擎零改动）")
    print(f"[ok] 已写出: {dst}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
