from __future__ import annotations

import re
from pathlib import Path
from typing import Mapping
from xml.sax.saxutils import escape

from ..body_rules import BodyParseRules, _strip_markdown_bold
from ..constants import *
from ..ir import (
    BlankBlock,
    Block,
    CodeBlock,
    FigureRowBlock,
    HeadingBlock,
    ImageBlock,
    MathBlock,
    PageBreakBlock,
    ParagraphBlock,
    QuoteBlock,
    TableBlock,
    TableCell,
    TableSplitBlock,
)
from ..math.converter import MathConverter
from ..media import MediaImage, MediaManager
from ..ooxml.render import (
    bookmark_paragraph_xml,
    figure_row_xml,
    image_paragraph_xml,
    math_image_paragraph_xml,
    math_paragraph_xml,
    page_break_xml,
    paragraph_with_inline_math_xml,
    paragraph_xml,
    section_break_paragraph_xml,
    table_xml,
)
from ..ooxml.xml import indent_xml, spacing_xml
from ..parser import parse_body_blocks
from ..styles import BodyRenderProfile
from ..table_utils import parse_bool_option, parse_int_option, parse_table_split_spec, split_table_rows
from ..toc import TocEntry, make_toc_entry


_CN_DIGITS = "零一二三四五六七八九"


def chinese_cardinal(n: int) -> str:
    """1..99 的中文数字，与 Word ``chineseCounting`` 在个位~十位一致。"""
    if not 1 <= n <= 99:
        return str(n)
    if n < 10:
        return _CN_DIGITS[n]
    if n < 20:
        return "十" + (_CN_DIGITS[n - 10] if n > 10 else "")
    tens, ones = divmod(n, 10)
    return _CN_DIGITS[tens] + "十" + (_CN_DIGITS[ones] if ones else "")


def strip_heading_prefix(text: str) -> str:
    """去掉标题开头的小节编号前缀（如 ``1 `` / ``1.1 ``），用于页眉标题。"""
    stripped = re.sub(r"^\d+(?:\.\d+)*\s+", "", text).strip()
    return stripped or text.strip()


def _require_hook(profile: BodyRenderProfile, name: str):
    hook = getattr(profile, name)
    if hook is None:
        raise ValueError(f"body render profile is missing hook: {name}")
    return hook


_OMIT_OPTION_VALUES = {"", "none", "omit", "inherit", "null"}


def _option_text(options: Mapping[str, str] | None, *names: str) -> str | None:
    if not options:
        return None
    for name in names:
        if name in options:
            return options[name].strip()
    return None


def _style_option(options: Mapping[str, str] | None, default_style: str | None) -> str | None:
    value = _option_text(options, "style", "p_style")
    if value is None:
        return default_style
    return None if value.lower() in _OMIT_OPTION_VALUES else value


def _int_option(options: Mapping[str, str] | None, *names: str) -> int | None:
    if not options:
        return None
    for name in names:
        if name in options:
            return parse_int_option(options, name)
    return None


def _bool_option(options: Mapping[str, str] | None, name: str) -> bool | None:
    if not options or name not in options:
        return None
    return parse_bool_option(options, name)


def _spacing_override_xml(options: Mapping[str, str] | None) -> str | None:
    if not options:
        return None
    names = (
        "spacing_line",
        "line",
        "spacing_before",
        "before",
        "spacing_after",
        "after",
        "spacing_before_lines",
        "before_lines",
        "spacing_after_lines",
        "after_lines",
        "spacing_line_rule",
        "line_rule",
    )
    if not any(name in options for name in names):
        return None
    line_rule = _option_text(options, "spacing_line_rule", "line_rule") or "auto"
    return spacing_xml(
        line=_int_option(options, "spacing_line", "line"),
        before=_int_option(options, "spacing_before", "before"),
        after=_int_option(options, "spacing_after", "after"),
        before_lines=_int_option(options, "spacing_before_lines", "before_lines"),
        after_lines=_int_option(options, "spacing_after_lines", "after_lines"),
        line_rule=line_rule,
    )


def _indent_override_xml(options: Mapping[str, str] | None) -> str | None:
    if not options:
        return None
    if _option_text(options, "indent") and (_option_text(options, "indent") or "").lower() in _OMIT_OPTION_VALUES:
        return ""
    names = (
        "left",
        "paragraph_left",
        "first_line",
        "paragraph_first_line",
        "first_line_chars",
        "paragraph_first_line_chars",
        "right",
        "hanging",
    )
    if not any(name in options for name in names):
        return None
    return indent_xml(
        left=_int_option(options, "left", "paragraph_left"),
        first_line=_int_option(options, "first_line", "paragraph_first_line"),
        first_line_chars=_int_option(options, "first_line_chars", "paragraph_first_line_chars"),
        right=_int_option(options, "right"),
        hanging=_int_option(options, "hanging"),
    )


def _paragraph_mark_rpr_xml(options: Mapping[str, str] | None) -> str:
    if not options:
        return ""
    parts: list[str] = []
    rstyle = _option_text(options, "ppr_rstyle", "paragraph_run_style", "mark_style")
    if rstyle and rstyle.lower() not in _OMIT_OPTION_VALUES:
        parts.append(f'<w:rStyle w:val="{escape(rstyle)}"/>')
    if _bool_option(options, "ppr_bold") is True or _bool_option(options, "mark_bold") is True:
        parts.append("<w:b/>")
    if _bool_option(options, "ppr_bold_cs") is True or _bool_option(options, "mark_bold_cs") is True:
        parts.append("<w:bCs/>")
    size = _int_option(options, "ppr_size", "mark_size")
    if size is not None:
        parts.append(f'<w:sz w:val="{size}"/>')
    size_cs = _int_option(options, "ppr_size_cs", "mark_size_cs")
    if size_cs is not None:
        parts.append(f'<w:szCs w:val="{size_cs}"/>')
    kern = _int_option(options, "ppr_kern", "mark_kern")
    if kern is not None:
        parts.append(f'<w:kern w:val="{kern}"/>')
    return f"<w:rPr>{''.join(parts)}</w:rPr>" if parts else ""


def _ppr_override_xml(
    options: Mapping[str, str] | None,
    *,
    default_ppr_extra: str,
    default_keep_next: bool = False,
) -> str:
    if not options:
        return default_ppr_extra + ("<w:keepNext/>" if default_keep_next else "")

    spacing_override = _spacing_override_xml(options)
    indent_override = _indent_override_xml(options)
    has_override = spacing_override is not None or indent_override is not None or any(
        key in options
        for key in (
            "keep_next",
            "keep_lines",
            "widow_control",
            "snap_to_grid",
            "ppr_rstyle",
            "paragraph_run_style",
            "mark_style",
            "ppr_bold",
            "ppr_bold_cs",
            "mark_bold",
            "mark_bold_cs",
            "ppr_size",
            "ppr_size_cs",
            "ppr_kern",
        )
    )
    if not has_override:
        return default_ppr_extra + ("<w:keepNext/>" if default_keep_next else "")

    parts: list[str] = []
    if spacing_override is not None:
        parts.append(spacing_override)
    if indent_override is not None:
        parts.append(indent_override)
    keep_next = _bool_option(options, "keep_next")
    if keep_next if keep_next is not None else default_keep_next:
        parts.append("<w:keepNext/>")
    if _bool_option(options, "keep_lines") is True:
        parts.append("<w:keepLines/>")
    widow_control = _bool_option(options, "widow_control")
    if widow_control is not None:
        parts.append("<w:widowControl/>" if widow_control else '<w:widowControl w:val="0"/>')
    snap_to_grid = _bool_option(options, "snap_to_grid")
    if snap_to_grid is not None:
        parts.append("<w:snapToGrid/>" if snap_to_grid else '<w:snapToGrid w:val="0"/>')
    parts.append(_paragraph_mark_rpr_xml(options))
    return "".join(parts)


def _run_kwargs_from_options(
    options: Mapping[str, str] | None,
    default_run_kwargs: dict[str, object] | None,
) -> dict[str, object] | None:
    if not options:
        return default_run_kwargs
    run_mode = _option_text(options, "run", "run_props")
    if run_mode and run_mode.lower() in _OMIT_OPTION_VALUES:
        return None
    keys = (
        "font_ascii",
        "font_hansi",
        "font_eastasia",
        "font_cs",
        "font_hint",
        "run_style",
        "rstyle",
        "size",
        "run_size",
        "size_cs",
        "run_size_cs",
        "spacing",
        "char_spacing",
        "bold",
        "bold_cs",
        "italic",
        "italic_cs",
    )
    if not any(key in options for key in keys):
        return default_run_kwargs
    run_kwargs = dict(default_run_kwargs or {})
    for option_name, kwarg_name in (
        ("font_ascii", "font_ascii"),
        ("font_hansi", "font_hansi"),
        ("font_eastasia", "font_eastasia"),
        ("font_cs", "font_cs"),
        ("font_hint", "font_hint"),
        ("run_style", "run_style"),
        ("rstyle", "run_style"),
    ):
        value = _option_text(options, option_name)
        if value and value.lower() in _OMIT_OPTION_VALUES:
            run_kwargs.pop(kwarg_name, None)
        elif value:
            run_kwargs[kwarg_name] = value
    size = _int_option(options, "size", "run_size")
    if size is not None:
        run_kwargs["size"] = size
    size_cs_text = _option_text(options, "size_cs", "run_size_cs")
    if size_cs_text is not None:
        if size_cs_text.lower() in {"0", "false", "no", "off", "omit", "none"}:
            run_kwargs["size_cs"] = False
        else:
            size_cs = _int_option(options, "size_cs", "run_size_cs")
            if size_cs is not None:
                run_kwargs["size_cs"] = size_cs
    char_spacing = _int_option(options, "char_spacing")
    if char_spacing is not None:
        run_kwargs["spacing"] = char_spacing
    for option_name in ("bold", "bold_cs", "italic", "italic_cs"):
        value = _bool_option(options, option_name)
        if value is not None:
            run_kwargs[option_name] = value
    return run_kwargs


def _paragraph_with_options_xml(
    text: str,
    options: Mapping[str, str] | None,
    *,
    default_style: str | None,
    default_align: str | None = None,
    default_ppr_extra: str = "",
    default_keep_next: bool = False,
    default_run_kwargs: dict[str, object] | None = None,
    preserve_breaks: bool = True,
    math_converter: MathConverter | None = None,
    reference_anchors: dict[str, str] | None = None,
    allow_bold: bool = True,
) -> str:
    align = _option_text(options, "align", "jc") if options else None
    omit_align = bool(align and align.lower() in _OMIT_OPTION_VALUES)
    if omit_align:
        align = None
    return paragraph_with_inline_math_xml(
        text,
        style=_style_option(options, default_style),
        align=align if align is not None else None if omit_align else default_align,
        ppr_extra=_ppr_override_xml(
            options,
            default_ppr_extra=default_ppr_extra,
            default_keep_next=default_keep_next,
        ),
        preserve_breaks=preserve_breaks,
        run_kwargs=_run_kwargs_from_options(options, default_run_kwargs),
        math_converter=math_converter,
        reference_anchors=reference_anchors,
        allow_bold=allow_bold,
    )


def build_document_elements(
    text: str,
    profile: BodyRenderProfile,
    *,
    rules: BodyParseRules | None = None,
    treat_first_heading_as_title: bool = True,
    math_converter: MathConverter | None = None,
    reference_anchors: dict[str, str] | None = None,
    markdown_dir: Path | None = None,
    media_manager: MediaManager | None = None,
    chapter_title_headers: list[tuple[str, str]] | None = None,
    initial_section_title_header_rid: str | None = None,
    initial_section_page_format: str | None = None,
    initial_section_page_start: int | None = None,
    initial_level1_header_index: int = 1,
) -> tuple[list[str], str, list[TocEntry]]:
    rules = rules or BodyParseRules()
    blocks = parse_body_blocks(text, rules=rules)
    return build_document_blocks(
        blocks,
        profile,
        rules=rules,
        treat_first_heading_as_title=treat_first_heading_as_title,
        math_converter=math_converter,
        reference_anchors=reference_anchors,
        markdown_dir=markdown_dir,
        media_manager=media_manager,
        chapter_title_headers=chapter_title_headers,
        initial_section_title_header_rid=initial_section_title_header_rid,
        initial_section_page_format=initial_section_page_format,
        initial_section_page_start=initial_section_page_start,
        initial_level1_header_index=initial_level1_header_index,
    )


def collect_toc_entries(
    text: str,
    *,
    rules: BodyParseRules | None = None,
    appendix_heading_normalizer=None,
) -> list[TocEntry]:
    rules = rules or BodyParseRules()
    blocks = parse_body_blocks(text, rules=rules)
    entries: list[TocEntry] = []
    current_top_heading = ""
    in_appendix = False
    current_appendix_index = 0

    for block in blocks:
        if not isinstance(block, HeadingBlock):
            continue
        raw_level = block.raw_level
        level = block.level
        heading_text = block.text

        if raw_level == 1:
            current_top_heading = heading_text
            if rules.is_appendix_heading(heading_text):
                in_appendix = True
                current_appendix_index = 0
                if rules.is_toc_heading(heading_text, in_appendix=in_appendix, level=level):
                    entries.append(
                        make_toc_entry(
                            len(entries) + 1,
                            level=level,
                            text=rules.display_heading_text(heading_text),
                        )
                    )
                continue
            in_appendix = False
        elif rules.is_appendix_heading(current_top_heading):
            in_appendix = True

        if rules.is_appendix_heading(current_top_heading) and rules.is_appendix_item_heading(raw_level):
            current_appendix_index += 1
            if appendix_heading_normalizer:
                display_text = appendix_heading_normalizer(heading_text, current_appendix_index)
            else:
                display_text = heading_text
            if rules.is_toc_heading(heading_text, in_appendix=in_appendix, level=1):
                entries.append(make_toc_entry(len(entries) + 1, level=1, text=display_text))
            continue

        if rules.is_toc_heading(heading_text, in_appendix=in_appendix, level=level):
            entries.append(
                make_toc_entry(
                    len(entries) + 1,
                    level=level,
                    text=rules.display_heading_text(heading_text),
                )
            )

    return entries


def build_document_blocks(
    blocks: list[Block],
    profile: BodyRenderProfile,
    *,
    rules: BodyParseRules | None = None,
    treat_first_heading_as_title: bool = True,
    math_converter: MathConverter | None = None,
    reference_anchors: dict[str, str] | None = None,
    markdown_dir: Path | None = None,
    media_manager: MediaManager | None = None,
    chapter_title_headers: list[tuple[str, str]] | None = None,
    initial_section_title_header_rid: str | None = None,
    initial_section_page_format: str | None = None,
    initial_section_page_start: int | None = None,
    initial_level1_header_index: int = 1,
) -> tuple[list[str], str, list[TocEntry]]:
    rules = rules or BodyParseRules()
    elements: list[str] = []
    current_top_heading = ""
    in_appendix = False
    in_reference_entries = False
    current_chapter_number = ""
    current_appendix_index = 0
    formula_counters: dict[str, int] = {}
    page_break_marker = page_break_xml()
    chapter_section_break_count = 0
    first_body_section_seen = False
    next_level1_header_index = initial_level1_header_index
    # 当前正在写入的“节”的页眉/页码属性，会在下一处分节符处被写出去（因为
    # OOXML 中段落内的 sectPr 描述的是“以该段落结尾的那一节”，所以当前节的
    # 属性必须延迟到下一节开始前的分节符里写入）。
    current_section_title_rid: str | None = initial_section_title_header_rid
    current_section_page_format: str | None = initial_section_page_format
    current_section_page_start: int | None = initial_section_page_start
    pending_table_split: list[int] | None = None
    last_table_caption_text: str | None = None
    toc_entries: list[TocEntry] = []

    heading_builder = _require_hook(profile, "heading_builder")
    acknowledgement_heading_builder = _require_hook(profile, "acknowledgement_heading_builder")
    caption_builder = _require_hook(profile, "caption_builder")
    reference_builder = _require_hook(profile, "reference_builder")
    special_paragraph_builder = profile.special_paragraph_builder
    image_builder = profile.image_builder or image_paragraph_xml
    table_builder = profile.table_builder or table_xml
    appendix_heading_normalizer = _require_hook(profile, "appendix_heading_normalizer")
    appendix_reference_normalizer = _require_hook(profile, "appendix_reference_normalizer")
    section_pr_builder = _require_hook(profile, "section_pr_builder")
    chapter_section_break_builder = profile.chapter_section_break_builder or section_break_paragraph_xml

    def is_chapter_section_break(element: str) -> bool:
        return "<w:sectPr>" in element and 'w:type w:val="nextPage"' in element

    def append_chapter_page_break(new_section_title_header_rid: str | None = None) -> None:
        # 延迟模型：OOXML 段落内的 <w:sectPr> 描述的是“以该段落结尾的那一节”。
        # 因此，位于新节开头处的这个分节符，应该携带【上一节】的页眉/页码属性；
        # 当前节的属性则延迟到“再下一个分节符”才写出去。这样可保证：
        # 1) 第一章首页由第二章前的分节符写入 decimal/start=1 + 第一章标题页眉；
        # 2) 前置节（英文目录）之后不再重复插入分节符，避免目录与第一章之间
        #    出现空白页。
        nonlocal first_body_section_seen, chapter_section_break_count, current_section_title_rid
        nonlocal current_section_page_format, current_section_page_start
        # 当前这个分节符描述的是【上一节】；若上一节末尾已紧接分页/分节，则
        # 不再重复插入，避免产生空白节。
        need_break = (
            elements and elements[-1] != page_break_marker and not is_chapter_section_break(elements[-1])
        )
        if need_break:
            rid = current_section_title_rid
            sect_pr = section_pr_builder(
                section_type="nextPage",
                with_header=True,
                footer_kind="page",
                page_number_format=current_section_page_format,
                page_number_start=current_section_page_start,
                with_title_header=bool(rid),
                title_header_rid=rid,
            )
            elements.append(chapter_section_break_builder(sect_pr))
            chapter_section_break_count += 1
        # 更新当前节属性为【新节】，供下一次分节符写出
        current_section_title_rid = new_section_title_header_rid
        current_section_page_format = "decimal"
        current_section_page_start = 1 if not first_body_section_seen else None
        first_body_section_seen = True

    def append_heading_page_break_without_section() -> None:
        if elements and elements[-1] != page_break_marker and not is_chapter_section_break(elements[-1]):
            elements.append(page_break_marker)

    def next_formula_number() -> str | None:
        scope = rules.formula_scope(
            in_appendix=in_appendix,
            current_appendix_index=current_appendix_index,
            current_chapter_number=current_chapter_number,
        )
        if not scope:
            return None
        formula_counters[scope] = formula_counters.get(scope, 0) + 1
        return rules.format_formula_number(scope, formula_counters[scope])

    def append_paragraph(paragraph: str, options: Mapping[str, str] | None = None) -> None:
        nonlocal in_reference_entries, last_table_caption_text
        paragraph = paragraph.strip()
        if not paragraph:
            return

        if in_appendix and current_appendix_index > 0:
            paragraph = appendix_reference_normalizer(paragraph, current_appendix_index)

        if rules.is_reference_heading(current_top_heading):
            if rules.should_skip_reference_paragraph(paragraph):
                return
            if rules.is_reference_entry(paragraph):
                in_reference_entries = True
                elements.append(reference_builder(paragraph, reference_anchors=reference_anchors))
                return
            if in_reference_entries:
                elements.append(reference_builder(paragraph, reference_anchors=reference_anchors))
                return

        if rules.is_caption_paragraph(paragraph):
            # 图/表题在源 markdown 里常以 `**图6.1** 描述` 的加粗形式书写，
            # 先剥掉 Markdown 加粗标记得到干净文本；题注不加粗、居中。
            caption_text = _strip_markdown_bold(paragraph)
            # Table captions appear immediately above the table they label, so
            # `keepNext` keeps them on the same page as the table when possible.
            # Figure captions appear after the image, so they do not need it.
            keep_next_caption = rules.is_table_caption(caption_text)
            if keep_next_caption:
                last_table_caption_text = caption_text
            caption_style = profile.styles.caption or profile.styles.normal
            if options:
                elements.append(
                    _paragraph_with_options_xml(
                        caption_text,
                        options,
                        default_style=caption_style,
                        default_align="center",
                        default_ppr_extra=spacing_xml(line=360, before=120, after=120)
                        + indent_xml(left=0, first_line=0),
                        default_keep_next=keep_next_caption,
                        default_run_kwargs={
                            "font_ascii": "Times New Roman",
                            "font_hansi": "Times New Roman",
                            "font_eastasia": "宋体",
                            "size": 21,
                            "bold": False,
                        },
                        math_converter=math_converter,
                        reference_anchors=reference_anchors,
                    )
                )
            else:
                elements.append(
                    caption_builder(
                        caption_text,
                        style=caption_style,
                        math_converter=math_converter,
                        reference_anchors=reference_anchors,
                        keep_next=keep_next_caption,
                    )
                )
            return

        if special_paragraph_builder is not None and not options:
            special_paragraph = special_paragraph_builder(
                paragraph,
                math_converter=math_converter,
                reference_anchors=reference_anchors,
            )
            if special_paragraph is not None:
                elements.append(special_paragraph)
                return

        normal_run = profile.normal_run.to_kwargs() if profile.normal_run else None
        normal_paragraph = profile.normal_paragraph
        allow_bold = profile.allow_bold
        if options:
            elements.append(
                _paragraph_with_options_xml(
                    paragraph,
                    options,
                    default_style=profile.styles.normal,
                    default_ppr_extra=normal_paragraph.ppr_extra,
                    default_run_kwargs=normal_run,
                    preserve_breaks=True,
                    math_converter=math_converter,
                    reference_anchors=reference_anchors,
                    allow_bold=allow_bold,
                )
            )
        elif normal_run:
            elements.append(
                paragraph_with_inline_math_xml(
                    paragraph,
                    style=profile.styles.normal,
                    ppr_extra=normal_paragraph.ppr_extra,
                    first_line_chars=normal_paragraph.first_line_chars,
                    first_line=normal_paragraph.first_line,
                    preserve_breaks=True,
                    run_kwargs=normal_run,
                    math_converter=math_converter,
                    reference_anchors=reference_anchors,
                    allow_bold=allow_bold,
                )
            )
        else:
            elements.append(
                paragraph_with_inline_math_xml(
                    paragraph,
                    style=profile.styles.normal,
                    first_line_chars=normal_paragraph.first_line_chars,
                    first_line=normal_paragraph.first_line,
                    preserve_breaks=True,
                    ppr_extra=normal_paragraph.ppr_extra,
                    math_converter=math_converter,
                    reference_anchors=reference_anchors,
                )
            )

    def resolve_image(target: str) -> MediaImage | None:
        image_path = markdown_dir / target if markdown_dir else Path(target)
        return media_manager.register_image(image_path) if media_manager else None

    def next_heading_text_after(index: int) -> str:
        for following in blocks[index + 1 :]:
            if isinstance(following, HeadingBlock):
                return following.text
            if not isinstance(following, PageBreakBlock):
                return ""
        return ""

    for block_index, block in enumerate(blocks):
        if isinstance(block, TableSplitBlock):
            pending_table_split = parse_table_split_spec(block.spec)
            continue

        if isinstance(block, FigureRowBlock):
            figure_items = [(resolve_image(image.target), image.alt_text) for image in block.images]
            if figure_items and media_manager is not None:
                figure_xml = figure_row_xml(figure_items, media_manager)
                if figure_xml:
                    elements.append(figure_xml)
            else:
                for raw_line in block.raw_lines:
                    append_paragraph(raw_line)
            continue

        if isinstance(block, ImageBlock):
            item = resolve_image(block.target)
            if item is not None:
                elements.append(
                    image_builder(
                        item,
                        media_manager,
                        alt_text=block.alt_text,
                        width_emu=block.width_emu,
                        height_emu=block.height_emu,
                        crop_top=block.crop_top,
                        crop_right=block.crop_right,
                        crop_bottom=block.crop_bottom,
                        crop_left=block.crop_left,
                        options=block.options,
                    )
                )
            else:
                append_paragraph(block.raw_text)
            continue

        if isinstance(block, PageBreakBlock):
            if block.before_heading_level == 1:
                next_heading_text = next_heading_text_after(block_index)
                if next_heading_text and rules.should_skip_section_break(next_heading_text):
                    elements.append(page_break_marker)
                else:
                    # 真正分节符由后续标题统一处理，此处仅插入普通分页，避免
                    # 与标题块产生的分节符叠加导致页眉/页码错位。
                    append_heading_page_break_without_section()
            else:
                elements.append(page_break_xml())
            continue

        if isinstance(block, BlankBlock):
            count = parse_int_option(block.options or {}, "count") or 1
            style = _style_option(block.options, None)
            ppr_extra = _ppr_override_xml(block.options, default_ppr_extra="")
            for _ in range(max(1, count)):
                elements.append(paragraph_xml(style=style, runs=[], ppr_extra=ppr_extra))
            continue

        if isinstance(block, HeadingBlock):
            raw_level = block.raw_level
            level = block.level
            heading_text = block.text

            if raw_level == 1:
                current_top_heading = heading_text
                in_reference_entries = False
                if rules.is_appendix_heading(heading_text):
                    in_appendix = True
                    current_appendix_index = 0
                    current_chapter_number = ""
                    # 附录前的分节符结束“最后一章”的节，并自动使用其标题页眉
                    # rId；附录本身无专用首页眉。
                    append_chapter_page_break(new_section_title_header_rid=None)
                    appendix_toc_entry: TocEntry | None = None
                    if rules.is_toc_heading(heading_text, in_appendix=in_appendix, level=level):
                        appendix_toc_entry = make_toc_entry(
                            len(toc_entries) + 1,
                            level=level,
                            text=rules.display_heading_text(heading_text),
                        )
                    appendix_heading = heading_builder(
                        rules.display_heading_text(heading_text),
                        level,
                        profile,
                        numbered=False,
                    )
                    if appendix_toc_entry is not None:
                        appendix_heading = bookmark_paragraph_xml(
                            appendix_heading,
                            bookmark_id=appendix_toc_entry.bookmark_id,
                            anchor=appendix_toc_entry.anchor,
                        )
                        toc_entries.append(appendix_toc_entry)
                    elements.append(appendix_heading)
                    continue
                in_appendix = False
                current_chapter_number = rules.extract_chapter_number(heading_text)
            elif rules.is_appendix_heading(current_top_heading):
                in_appendix = True

            if len(elements) == 0 and treat_first_heading_as_title:
                elements.append(paragraph_xml(heading_text, style=profile.styles.title, align="center"))
                continue

            toc_entry: TocEntry | None = None
            if rules.is_toc_heading(heading_text, in_appendix=in_appendix, level=level) and not (
                rules.is_appendix_heading(current_top_heading) and rules.is_appendix_item_heading(raw_level)
            ):
                display_for_toc = rules.display_heading_text(heading_text)
                toc_entry = make_toc_entry(len(toc_entries) + 1, level=level, text=display_for_toc)

            if rules.is_appendix_heading(current_top_heading) and rules.is_appendix_item_heading(raw_level):
                current_appendix_index += 1
                append_chapter_page_break()
                normalized_heading = appendix_heading_normalizer(heading_text, current_appendix_index)
                if rules.is_toc_heading(heading_text, in_appendix=in_appendix, level=1):
                    toc_entry = make_toc_entry(len(toc_entries) + 1, level=1, text=normalized_heading)
                appendix_heading = heading_builder(
                    normalized_heading,
                    1,
                    profile,
                    numbered=False,
                )
                if toc_entry is not None:
                    appendix_heading = bookmark_paragraph_xml(
                        appendix_heading,
                        bookmark_id=toc_entry.bookmark_id,
                        anchor=toc_entry.anchor,
                    )
                    toc_entries.append(toc_entry)
                elements.append(appendix_heading)
                continue

            if raw_level == 1:
                if not rules.should_skip_section_break(heading_text):
                    # 本节/本章首页“标题页眉”的 rId：OOXML 段落内的 sectPr 描述的是
                    # “开启它后面那一节”，因此该 rId 直接写进本次分节符，用于本章/
                    # 本节首页页眉。编号章节用“第X章 名称”；参考文献、致谢、附录等
                    # 非编号节直接用“名称”（如“参考文献”），各自拥有独立标题页眉，
                    # 不再复用上一章（或上一节）的 rId，避免出现多余/错位的重复分节符。
                    this_chapter_title_rid = None
                    if chapter_title_headers is not None:
                        is_numbered_chapter = not (
                            rules.is_reference_heading(heading_text)
                            or rules.is_acknowledgement_heading(heading_text)
                            or rules.is_appendix_heading(heading_text)
                            or rules.is_unnumbered_heading(
                                heading_text, in_appendix=in_appendix, level=level
                            )
                        )
                        this_chapter_title_rid = f"rId{TITLE_HEADER_RID_BASE + next_level1_header_index}"
                        if is_numbered_chapter:
                            chapter_label = f"第{chinese_cardinal(next_level1_header_index)}章"
                            chapter_full = f"{chapter_label} {strip_heading_prefix(heading_text)}"
                        else:
                            chapter_full = strip_heading_prefix(heading_text)
                        chapter_title_headers.append((this_chapter_title_rid, chapter_full))
                        next_level1_header_index += 1
                    append_chapter_page_break(new_section_title_header_rid=this_chapter_title_rid)
                else:
                    append_heading_page_break_without_section()

            is_unnumbered = rules.is_unnumbered_heading(heading_text, in_appendix=in_appendix, level=level)
            display_heading_text = rules.display_heading_text(heading_text)

            caption_style_id = profile.styles.caption or ""
            previous_is_caption = bool(
                caption_style_id and elements and f'w:pStyle w:val="{caption_style_id}"' in elements[-1]
            )
            if rules.is_acknowledgement_heading(heading_text) and raw_level == 1:
                heading_xml = acknowledgement_heading_builder(display_heading_text, profile)
            else:
                heading_xml = heading_builder(
                    display_heading_text,
                    level,
                    profile,
                    numbered=not is_unnumbered,
                    keep_with_next=not previous_is_caption,
                )
            if toc_entry is not None:
                heading_xml = bookmark_paragraph_xml(
                    heading_xml,
                    bookmark_id=toc_entry.bookmark_id,
                    anchor=toc_entry.anchor,
                )
                toc_entries.append(toc_entry)
            elements.append(heading_xml)
            continue

        if isinstance(block, QuoteBlock):
            elements.append(
                paragraph_xml(
                    block.text,
                    style=profile.styles.quote,
                )
            )
            continue

        if isinstance(block, TableBlock):
            rows = [list(row) for row in block.rows]
            if rows:
                rich_table = bool(block.options) or any(
                    cell.colspan > 1 or cell.rowspan > 1 or cell.align or cell.header
                    for row in rows
                    for cell in row
                )
                if rich_table:
                    normalized = rows
                else:
                    width = len(rows[0])
                    normalized = [row[:width] + [TableCell("")] * max(0, width - len(row)) for row in rows]
                table_chunks = split_table_rows(normalized, pending_table_split or [])
                caption_style = profile.styles.caption or profile.styles.normal
                for chunk_idx, table_chunk in enumerate(table_chunks):
                    if chunk_idx > 0 and last_table_caption_text:
                        elements.append(
                            caption_builder(
                                f"{last_table_caption_text}（续）",
                                style=caption_style,
                                math_converter=math_converter,
                                reference_anchors=reference_anchors,
                                keep_next=True,
                            )
                        )
                    elements.append(
                        table_builder(
                            table_chunk,
                            cell_style=profile.styles.table_cell,
                            options=block.options,
                            math_converter=math_converter,
                            reference_anchors=reference_anchors,
                        )
                    )
                pending_table_split = None
                last_table_caption_text = None
            continue

        if isinstance(block, CodeBlock):
            elements.append(
                paragraph_xml(
                    block.text,
                    style=profile.styles.code,
                    preserve_breaks=True,
                    ppr_extra=profile.code_paragraph.ppr_extra,
                )
            )
            continue

        if isinstance(block, MathBlock):
            # 公式号一律按章节自动编号（如（4-1）、（4-2）），不使用源稿里的
            # 全局 (N) 编号——HIT 模板要求公式按“章-序号”编号。
            # 兜底：若公式文本仍残留行尾 (N)（旧版预处理会写入），先剥离，避免重复/错位。
            raw_text = (block.text or "").strip()
            m_eqnum = re.search(r"\(\s*(\d{1,3})\s*\)\s*$", raw_text)
            math_text = raw_text[: m_eqnum.start()].rstrip() if m_eqnum else raw_text
            auto_eq_num = next_formula_number()
            if block.image:
                image_item = resolve_image(block.image)
                if image_item is not None and media_manager is not None:
                    elements.append(
                        math_image_paragraph_xml(
                            image_item,
                            media_manager,
                            alt_text=block.image_alt,
                            width_emu=block.image_width_emu,
                            height_emu=block.image_height_emu,
                            equation_number=auto_eq_num,
                            style=profile.styles.math,
                            image_mode=block.image_mode,
                            first_line=block.image_first_line,
                            first_line_chars=block.image_first_line_chars,
                            position=block.image_position,
                            include_shapetype=block.image_include_shapetype,
                        )
                    )
                    continue
            elements.append(
                math_paragraph_xml(
                    math_text,
                    style=profile.styles.math,
                    math_converter=math_converter,
                    equation_number=auto_eq_num,
                )
            )
            continue

        if isinstance(block, ParagraphBlock):
            append_paragraph(block.text, block.options)

    # 文档末尾的 sectPr 描述最后一节。把当前节的页眉/页码属性写入最终
    # sectPr，避免在文档末尾多插入一个空白页。
    final_sect_pr = section_pr_builder(
        with_header=True,
        footer_kind="page",
        page_number_format=current_section_page_format,
        page_number_start=current_section_page_start,
        with_title_header=bool(current_section_title_rid),
        title_header_rid=current_section_title_rid,
    )

    return elements, final_sect_pr, toc_entries
