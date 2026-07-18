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
