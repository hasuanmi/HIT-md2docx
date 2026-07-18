from __future__ import annotations

import re
from dataclasses import dataclass

# 行内双语标题分隔符：作者可在 Markdown 标题后用 ` | ` 或 ` :: ` 附上英文，
# 例如 `## 1.1 研究背景与意义 | Research Background`。英文目录使用英文部分，
# 正文/中文目录/页眉自动剥除。要求英文部分以拉丁字母开头，避免误拆中文里的 `|`。
_BILINGUAL_RE = re.compile(r"\s*(?:[|｜]|:{2})\s*([A-Za-z].*)$")


def split_bilingual(label: str) -> tuple[str, str | None]:
    """把 `中文 | English` 拆成 (中文, 英文)。无有效英文后缀时返回 (原文, None)。"""
    label = (label or "").strip()
    m = _BILINGUAL_RE.search(label)
    if not m:
        return label, None
    cn = label[: m.start()].strip()
    en = m.group(1).strip()
    if not en:
        return label, None
    return cn, en


def strip_bilingual(label: str) -> str:
    """返回行内标题的中文部分（无后缀时原样返回）。用于正文/中文目录/页眉展示。"""
    return split_bilingual(label)[0]


# 标题字典「规范键」：把各种引号 / 全半角歧义标点 / 空白归一成稳定形式，
# 使 heading_translations.json 的 key 不必与被翻译论文标题逐字节一致。
# 例如 ASCII 直引号 "、"弯引号 "、中文「」/『』/«»、全角（）：，等变体都能命中同一 key。
_QUOTE_TO_ASCII = {
    "\u2018": "'", "\u2019": "'", "\u201a": "'", "\u201b": "'",
    "\u2032": "'", "\uff07": "'",
    "\u201c": "\"", "\u201d": "\"", "\u201e": "\"", "\u201f": "\"",
    "\u300c": "\"", "\u300d": "\"", "\u300e": "\"", "\u300f": "\"",
    "\u00ab": "\"", "\u00bb": "\"", "\uff02": "\"",
}
_FULLWIDTH_PUNCT = {
    "\uff08": "(", "\uff09": ")", "\uff1a": ":", "\uff0c": ",",
    "\uff1b": ";",
}


def normalize_heading_key(label: str) -> str:
    """将标题文本归一成字典查表用的规范键。

    1) 全部引号变体归一成 ASCII 直引号 " / 单引号 '；
    2) 常见全角歧义标点（（）：；，）归一成半角；
    3) 折叠连续空白并去首尾空格。
    这样无论论文标题用哪种引号、字典 key 用哪种引号，都会归一到同一 key，
    自动命中——无需手工对齐码点，换篇论文也不会因引号样式漏翻。
    """
    if not label:
        return label
    s = "".join(_QUOTE_TO_ASCII.get(ch, ch) for ch in label)
    s = "".join(_FULLWIDTH_PUNCT.get(ch, ch) for ch in s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


TOC_BOOKMARK_ID_START = 5000
TOC_BOOKMARK_NAME_START = 102000000


@dataclass(frozen=True)
class TocEntry:
    level: int
    text: str
    anchor: str
    bookmark_id: int


def make_toc_entry(index: int, *, level: int, text: str) -> TocEntry:
    return TocEntry(
        level=max(1, min(level, 3)),
        text=text,
        anchor=f"_Toc{TOC_BOOKMARK_NAME_START + index}",
        bookmark_id=TOC_BOOKMARK_ID_START + index,
    )
