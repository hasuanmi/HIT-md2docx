from __future__ import annotations

from dataclasses import dataclass
import re

from .constants import INLINE_MATH_PATTERN


@dataclass(frozen=True)
class InlineSegment:
    kind: str
    value: str
    bold: bool = False
    italic: bool = False


def split_inline_code(text: str) -> list[tuple[str, str]]:
    parts: list[tuple[str, str]] = []
    i = 0
    last = 0
    while i < len(text):
        if text[i] != "`":
            i += 1
            continue

        tick_count = 1
        while i + tick_count < len(text) and text[i + tick_count] == "`":
            tick_count += 1

        marker = "`" * tick_count
        closing = text.find(marker, i + tick_count)
        if closing == -1:
            i += tick_count
            continue

        if i > last:
            parts.append(("text", text[last:i]))
        parts.append(("code", text[i + tick_count : closing]))
        i = closing + tick_count
        last = i

    if last < len(text):
        parts.append(("text", text[last:]))

    return parts if parts else [("text", text)]


def split_inline_emphasis(text: str) -> list[tuple[str, str]]:
    parts: list[tuple[str, str]] = []
    pattern = re.compile(r"\*\*.+?\*\*|\*[^*\n][^*\n]*?\*")
    last = 0
    for match in pattern.finditer(text):
        if match.start() > last:
            parts.append(("text", text[last:match.start()]))
        token = match.group(0)
        if token.startswith("**") and token.endswith("**"):
            parts.append(("bold", token[2:-2]))
        else:
            parts.append(("italic", token[1:-1]))
        last = match.end()
    if last < len(text):
        parts.append(("text", text[last:]))
    return parts if parts else [("text", text)]


def split_inline_math(text: str) -> list[tuple[str, str]]:
    parts: list[tuple[str, str]] = []
    last = 0
    for match in INLINE_MATH_PATTERN.finditer(text):
        if match.start() > last:
            parts.append(("text", text[last:match.start()]))
        latex = match.group(1).strip()
        if latex:
            parts.append(("math", latex))
        else:
            parts.append(("text", "$$"))
        last = match.end()
    if last < len(text):
        parts.append(("text", text[last:]))
    return [(kind, value.replace(r"\$", "$")) for kind, value in parts if value]


def _is_escaped(text: str, index: int) -> bool:
    slash_count = 0
    cursor = index - 1
    while cursor >= 0 and text[cursor] == "\\":
        slash_count += 1
        cursor -= 1
    return slash_count % 2 == 1


def _code_marker_at(text: str, index: int) -> str | None:
    if index >= len(text) or text[index] != "`" or _is_escaped(text, index):
        return None
    tick_count = 1
    while index + tick_count < len(text) and text[index + tick_count] == "`":
        tick_count += 1
    return "`" * tick_count


def _is_inline_math_delimiter(text: str, index: int) -> bool:
    return (
        index < len(text)
        and text[index] == "$"
        and not _is_escaped(text, index)
        and (index + 1 >= len(text) or text[index + 1] != "$")
    )


def _find_closing_math(text: str, start: int) -> int:
    cursor = start
    while cursor < len(text):
        if text[cursor] == "\n":
            return -1
        if _is_inline_math_delimiter(text, cursor):
            return cursor
        cursor += 1
    return -1


def _find_closing_marker(text: str, marker: str, start: int) -> int:
    cursor = start
    while cursor < len(text):
        if text[cursor] == "\n":
            return -1

        code_marker = _code_marker_at(text, cursor)
        if code_marker is not None:
            closing = text.find(code_marker, cursor + len(code_marker))
            if closing != -1:
                cursor = closing + len(code_marker)
                continue

        if _is_inline_math_delimiter(text, cursor):
            closing = _find_closing_math(text, cursor + 1)
            if closing != -1:
                cursor = closing + 1
                continue

        if marker == "*" and text.startswith("**", cursor) and not _is_escaped(text, cursor):
            cursor += 2
            continue

        if text.startswith(marker, cursor) and not _is_escaped(text, cursor):
            return cursor
        cursor += 1
    return -1


def _append_segment(segments: list[InlineSegment], segment: InlineSegment) -> None:
    if (
        segment.kind == "text"
        and segments
        and segments[-1].kind == "text"
        and segments[-1].bold == segment.bold
        and segments[-1].italic == segment.italic
    ):
        previous = segments[-1]
        segments[-1] = InlineSegment("text", previous.value + segment.value, previous.bold, previous.italic)
        return
    segments.append(segment)


def _parse_inline_segments(text: str, *, bold: bool = False, italic: bool = False) -> list[InlineSegment]:
    segments: list[InlineSegment] = []
    buffer: list[str] = []
    cursor = 0

    def flush_text() -> None:
        if not buffer:
            return
        _append_segment(segments, InlineSegment("text", "".join(buffer), bold, italic))
        buffer.clear()

    while cursor < len(text):
        code_marker = _code_marker_at(text, cursor)
        if code_marker is not None:
            closing = text.find(code_marker, cursor + len(code_marker))
            if closing != -1:
                flush_text()
                _append_segment(
                    segments,
                    InlineSegment("code", text[cursor + len(code_marker) : closing], bold, italic),
                )
                cursor = closing + len(code_marker)
                continue

        if _is_inline_math_delimiter(text, cursor):
            closing = _find_closing_math(text, cursor + 1)
            if closing != -1:
                latex = text[cursor + 1 : closing].strip()
                if latex:
                    flush_text()
                    _append_segment(segments, InlineSegment("math", latex.replace(r"\$", "$"), bold, italic))
                else:
                    buffer.append(text[cursor : closing + 1])
                cursor = closing + 1
                continue

        if text.startswith("**", cursor) and not _is_escaped(text, cursor):
            closing = _find_closing_marker(text, "**", cursor + 2)
            if closing > cursor + 2:
                flush_text()
                for segment in _parse_inline_segments(text[cursor + 2 : closing], bold=True, italic=italic):
                    _append_segment(segments, segment)
                cursor = closing + 2
                continue
            buffer.append("**")
            cursor += 2
            continue

        if text[cursor] == "*" and not _is_escaped(text, cursor):
            closing = _find_closing_marker(text, "*", cursor + 1)
            if closing > cursor + 1:
                flush_text()
                for segment in _parse_inline_segments(text[cursor + 1 : closing], bold=bold, italic=True):
                    _append_segment(segments, segment)
                cursor = closing + 1
                continue

        buffer.append(text[cursor])
        cursor += 1

    flush_text()
    return segments


def split_inline_segments(text: str) -> list[InlineSegment]:
    return _parse_inline_segments(text)
