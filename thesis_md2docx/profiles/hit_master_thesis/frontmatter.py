from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

from ...constants import (
    DEFAULT_LOCAL_COVER_ASSETS_REL,
    EMU_PER_INCH,
    SIGNATURE_IMAGE_HEIGHT_EMU,
    SIGNATURE_IMAGE_WIDTH_EMU,
)
from ...frontmatter import split_statement_content
from ...markdown import parse_cover_info, split_cover_title_lines
from ...math.converter import MathConverter
from ...media import MediaImage, MediaManager
from ...ooxml.render import (
    add_page_break_before_paragraph_xml,
    formatted_paragraph_xml,
    image_run_xml,
    page_break_xml,
    paragraph_with_inline_math_xml,
    paragraph_xml,
)
from ...ooxml.xml import indent_xml, run_text_xml, spacing_xml
from .styles import STYLE_BODY, STYLE_FRONT_HEADING


@dataclass(frozen=True)
class CoverFieldSpec:
    source_key: str
    label: str


@dataclass(frozen=True)
class CoverInfoRow:
    label: str
    value: str
    draw_top_border: bool = True


@dataclass(frozen=True)
class DeclarationSignatureSpec:
    author_label: str = "作者签名："
    date_label: str = "签字日期："
    signature_alt: str = "电子签名"
    blank_count_without_image: int = 17
    blank_count_with_image: int = 13

    def blank_count(self, *, has_signature_image: bool) -> int:
        return self.blank_count_with_image if has_signature_image else self.blank_count_without_image


# 硕士学位论文封面字段（深圳研究生院）
HIT_COVER_FIELDS: tuple[CoverFieldSpec, ...] = (
    CoverFieldSpec("作者", "作    者："),
    CoverFieldSpec("学号", "学    号："),
    CoverFieldSpec("所在单位", "所在单位："),
    CoverFieldSpec("学科专业", "学科专业："),
    CoverFieldSpec("指导教师", "指导教师："),
    CoverFieldSpec("答辩日期", "答辩日期："),
    CoverFieldSpec("学位类别", "学位类别："),
    CoverFieldSpec("学校代码", "学校代码："),
    CoverFieldSpec("密级", "密    级："),
)

# 英文内封字段（与 HIT_COVER_FIELDS 一一对应）
HIT_COVER_FIELDS_EN: tuple[CoverFieldSpec, ...] = (
    CoverFieldSpec("作者", "Candidate:"),
    CoverFieldSpec("学号", "Student ID:"),
    CoverFieldSpec("所在单位", "Affiliation:"),
    CoverFieldSpec("学科专业", "Speciality:"),
    CoverFieldSpec("指导教师", "Supervisor:"),
    CoverFieldSpec("答辩日期", "Date of Defence:"),
    CoverFieldSpec("学位类别", "Academic Degree Applied for:"),
    CoverFieldSpec("学校代码", "School Code:"),
    CoverFieldSpec("密级", "Confidentiality:"),
)


XJU_DECLARATION_SIGNATURE = DeclarationSignatureSpec()


def _contains_cjk_text(text: str) -> bool:
    return any("\u4e00" <= char <= "\u9fff" for char in text)


def resolve_cover_assets_dir(markdown_path: Path, assets_dir: Path | None, *, use_cover_assets: bool) -> Path | None:
    if not use_cover_assets:
        return None
    candidates: list[Path] = []
    if assets_dir is not None:
        candidates.append(assets_dir)
    local_assets_dir = markdown_path.parent / DEFAULT_LOCAL_COVER_ASSETS_REL
    if local_assets_dir not in candidates:
        candidates.append(local_assets_dir)
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return candidates[0] if candidates else None


def cover_info_table_xml(title: str, cover_info: dict[str, str], *, english: bool = False) -> str:
    title_lines = split_cover_title_lines(title)
    info_rows: list[CoverInfoRow] = []
    fields = HIT_COVER_FIELDS_EN if english else HIT_COVER_FIELDS

    title_label = "" if english else "论文题目："
    if title_lines:
        info_rows.append(CoverInfoRow(title_label, title_lines[0], draw_top_border=False))
        for extra_line in title_lines[1:]:
            info_rows.append(CoverInfoRow("", extra_line))

    for field in fields:
        value = cover_info.get(field.source_key)
        if value:
            info_rows.append(CoverInfoRow(field.label, value))

    tbl_pr = (
        "<w:tblPr>"
        '<w:tblW w:w="6943" w:type="dxa"/>'
        '<w:jc w:val="center"/>'
        '<w:tblLayout w:type="fixed"/>'
        '<w:tblLook w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:val="04A0"/>'
        "</w:tblPr>"
    )
    tbl_grid = '<w:tblGrid><w:gridCol w:w="1948"/><w:gridCol w:w="4995"/></w:tblGrid>'

    if english:
        label_run = {
            "font_ascii": "Times New Roman",
            "font_hansi": "Times New Roman",
            "font_eastasia": "Times New Roman",
            "bold": True,
            "bold_cs": False,
            "size": 36,
        }
        value_run = {
            "font_ascii": "Times New Roman",
            "font_hansi": "Times New Roman",
            "font_eastasia": "Times New Roman",
            "bold": True,
            "bold_cs": False,
            "size": 36,
        }
    else:
        label_run = {
            "font_eastasia": "宋体",
            "font_hint": "eastAsia",
            "bold": True,
            "bold_cs": False,
            "size": 36,
        }
        value_run = {
            "font_eastasia": "宋体",
            "font_hint": "eastAsia",
            "bold": True,
            "bold_cs": False,
            "size": 36,
        }

    rows_xml: list[str] = []
    for idx, row in enumerate(info_rows):
        label_para = formatted_paragraph_xml(
            row.label,
            align="center" if row.label else None,
            ppr_extra="",
            run_kwargs=label_run,
        )
        value_para = formatted_paragraph_xml(
            row.value,
            align="center",
            ppr_extra="",
            run_kwargs=value_run,
        )
        value_borders = ["<w:tcBorders>"]
        if row.draw_top_border and idx > 0:
            value_borders.append('<w:top w:val="single" w:color="auto" w:sz="4" w:space="0"/>')
        value_borders.append('<w:bottom w:val="single" w:color="auto" w:sz="4" w:space="0"/>')
        value_borders.append("</w:tcBorders>")

        rows_xml.append(
            "<w:tr>"
            '<w:trPr><w:trHeight w:val="680"/></w:trPr>'
            '<w:tc><w:tcPr><w:tcW w:w="1948" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>'
            + label_para
            + "</w:tc>"
            + '<w:tc><w:tcPr><w:tcW w:w="4995" w:type="dxa"/>'
            + "".join(value_borders)
            + '<w:vAlign w:val="center"/></w:tcPr>'
            + value_para
            + "</w:tc>"
            + "</w:tr>"
        )

    return f"<w:tbl>{tbl_pr}{tbl_grid}{''.join(rows_xml)}</w:tbl>"


def _cover_line(
    text: str,
    size: int,
    *,
    bold: bool = False,
    ascii_font: str = "宋体",
    eastasia_font: str = "宋体",
    line_val: int | None = None,
    line_rule: str = "auto",
    before: int | None = None,
    after: int | None = None,
) -> str:
    ppr = ""
    if before is not None:
        ppr += spacing_xml(before=before)
    if after is not None:
        ppr += spacing_xml(after=after)
    if line_val is not None:
        ppr += spacing_xml(line=line_val, line_rule=line_rule)
    return formatted_paragraph_xml(
        text,
        align="center",
        ppr_extra=ppr,
        run_kwargs={
            "font_ascii": ascii_font,
            "font_hansi": ascii_font,
            "font_eastasia": eastasia_font,
            "bold": bold,
            "size": size,
        },
    )


def _cover_field_para(label: str, value: str, size: int = 28) -> str:
    """封面元数据行：冒号左侧黑体、右侧宋体（哈工大规范：冒号左黑体4号、右宋体4号，行距1.5）。"""
    label_run = run_text_xml(label, font_eastasia="黑体", font_hint="eastAsia", bold=True, size_cs=False, size=size)
    value_run = run_text_xml(value, font_eastasia="宋体", font_hint="eastAsia", size_cs=False, size=size)
    return paragraph_xml(runs=[label_run, value_run], align="center", ppr_extra=spacing_xml(line=360))


def build_cover_elements(
    title: str,
    cover_info: dict[str, str],
    *,
    cover_assets_dir: Path | None = None,
    media_manager: MediaManager | None = None,
) -> list[str]:
    elements: list[str] = []
    spacer = lambda after=240: paragraph_xml(" ", ppr_extra=spacing_xml(after=after))

    # ===== 校名 + 论文性质 =====
    elements.append(_cover_line("哈尔滨工业大学", 48, bold=True, eastasia_font="宋体", line_val=600, line_rule="exact"))
    elements.append(_cover_line("硕士学位论文", 36, bold=True, eastasia_font="宋体", line_val=600, line_rule="exact", after=200))

    # ===== 题目（中 / 英）=====
    elements.append(_cover_line(title, 44, bold=True, eastasia_font="黑体", line_val=600, line_rule="exact"))
    en_title = (cover_info.get("英文题目") or "").strip().upper()
    if en_title:
        elements.append(
            _cover_line(en_title, 44, bold=True, ascii_font="Times New Roman", eastasia_font="Times New Roman", line_val=600, line_rule="exact")
        )
    elements.append(spacer(200))

    # ===== 分类号行（宋体小4）=====
    cn_class = cover_info.get("国内图书分类号") or "××××"
    school_code = cover_info.get("学校代码") or "×××"
    intl_class = cover_info.get("国际图书分类号") or "××××"
    secrecy = cover_info.get("密级") or "公开"
    elements.append(_cover_line(f"国内图书分类号：{cn_class}        学校代码：{school_code}", 24, eastasia_font="宋体", line_val=300))
    elements.append(_cover_line(f"国际图书分类号：{intl_class}           密级：{secrecy}", 24, eastasia_font="宋体", line_val=300))
    elements.append(spacer(200))

    # 第 1 页（中文封面）结束 → 空白页（第 2 页）→ 中文扉页（第 3 页）
    elements.append(page_break_xml())
    elements.append(page_break_xml())

    # ===== 中文详细信息块 =====
    elements.append(_cover_line("硕士学位论文", 36, bold=True, eastasia_font="宋体", line_val=600, line_rule="exact"))
    elements.append(_cover_line(title, 44, bold=True, eastasia_font="黑体", line_val=600, line_rule="exact", after=120))

    elements.append(_cover_field_para("硕士研究生：", cover_info.get("作者") or "×××"))
    elements.append(_cover_field_para("导    师：", cover_info.get("指导教师") or "×××"))
    elements.append(_cover_field_para("申请学位：", cover_info.get("学位类别") or "×××"))
    elements.append(_cover_field_para("学    科：", cover_info.get("学科专业") or "×××"))
    elements.append(_cover_field_para("所 在 单 位：", cover_info.get("所在单位") or "×××"))
    elements.append(_cover_field_para("答 辩 日 期：", cover_info.get("答辩日期") or "×××"))
    elements.append(_cover_field_para("授予学位单位：", "哈尔滨工业大学"))
    elements.append(spacer(200))

    # 第 3 页（中文扉页）结束 → 空白页（第 4 页）→ 英文扉页（第 5 页）
    elements.append(page_break_xml())
    elements.append(page_break_xml())

    # ===== 英文详细信息块 =====
    elements.append(_cover_line("Classified Index: " + cn_class, 24, ascii_font="Times New Roman", eastasia_font="Times New Roman", line_val=300))
    elements.append(_cover_line("U.D.C: " + intl_class, 24, ascii_font="Times New Roman", eastasia_font="Times New Roman", line_val=300))
    elements.append(_cover_line("Dissertation for the Master Degree", 36, ascii_font="Times New Roman", eastasia_font="Times New Roman", line_val=600, line_rule="exact", after=120))
    if en_title:
        elements.append(_cover_line(en_title, 44, bold=True, ascii_font="Times New Roman", eastasia_font="Times New Roman", line_val=600, line_rule="exact", after=120))

    elements.append(_cover_line("Candidate：" + (cover_info.get("英文作者") or cover_info.get("作者") or "XXX"), 28, ascii_font="Times New Roman", eastasia_font="Times New Roman", line_val=360))
    elements.append(_cover_line("Supervisor：" + (cover_info.get("英文导师") or cover_info.get("指导教师") or "XXX"), 28, ascii_font="Times New Roman", eastasia_font="Times New Roman", line_val=360))
    elements.append(_cover_line("Academic Degree Applied for：" + (cover_info.get("英文学位") or cover_info.get("学位类别") or "XXX"), 28, ascii_font="Times New Roman", eastasia_font="Times New Roman", line_val=360))
    elements.append(_cover_line("Speciality：" + (cover_info.get("英文学科") or cover_info.get("学科专业") or "XXX"), 28, ascii_font="Times New Roman", eastasia_font="Times New Roman", line_val=360))
    elements.append(_cover_line("Affiliation：" + (cover_info.get("英文单位") or cover_info.get("所在单位") or "XXX"), 28, ascii_font="Times New Roman", eastasia_font="Times New Roman", line_val=360))
    elements.append(_cover_line("Date of Defence：" + (cover_info.get("英文答辩日期") or cover_info.get("答辩日期") or "XXX"), 28, ascii_font="Times New Roman", eastasia_font="Times New Roman", line_val=360))
    elements.append(_cover_line("Degree-Conferring-Institution：Harbin Institute of Technology", 28, ascii_font="Times New Roman", eastasia_font="Times New Roman", line_val=360))

    # 第 5 页（英文扉页）结束 → 空白页（第 6 页）；正文分节符将把摘要推到第 7 页
    elements.append(page_break_xml())
    return elements



def build_front_heading(
    text: str,
    *,
    english: bool = False,
    toc: bool = False,
    statement: bool = False,
    bold: bool = False,
    page_break_before: bool = False,
) -> str:
    if toc:
        # 目录标题直接来自 FrontMatterSpec.toc_title（默认“目  录”），
        # 居中对齐、黑体、与正文前标题同字号。
        paragraph = formatted_paragraph_xml(
            text,
            style="afa",
            align="center",
            ppr_extra=spacing_xml(before=390, after=312, line=360, line_rule="auto"),
            run_kwargs={
                "font_ascii": "黑体",
                "font_hansi": "黑体",
                "font_eastasia": "黑体",
                "size": 36,
            },
        )
        return add_page_break_before_paragraph_xml(paragraph) if page_break_before else paragraph

    if statement:
        run_kwargs = {
            "font_ascii": "黑体",
            "font_hansi": "黑体",
            "font_eastasia": "黑体",
            "font_hint": "eastAsia",
            "size": 36,
        }
        ppr_extra = (
            '<w:snapToGrid w:val="0"/>'
            + spacing_xml(before=390, after=312, line=360, line_rule="auto")
            + '<w:rPr><w:rFonts w:ascii="黑体" w:hAnsi="黑体" w:eastAsia="黑体" w:hint="eastAsia"/>'
            '<w:sz w:val="36"/><w:szCs w:val="36"/></w:rPr>'
        )
    elif english:
        # 英文标题（Abstract / Contents）：拉丁字母用 Times New Roman，
        # 中文（若有）用黑体，与正文英文统一字体。小二(18pt=36半磅)。
        # Contents 等目录级标题按规范使用粗体。
        run_kwargs = {
            "font_ascii": "Times New Roman",
            "font_hansi": "Times New Roman",
            "font_eastasia": "黑体",
            "size": 36,
            "bold": bold,
        }
        bpr = '<w:b/><w:bCs/>' if bold else ''
        ppr_extra = (
            '<w:snapToGrid w:val="0"/>'
            + spacing_xml(before=390, after=312, line=360, line_rule="auto")
            + '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="黑体" w:hint="eastAsia"/>'
            f'{bpr}<w:sz w:val="36"/><w:szCs w:val="36"/></w:rPr>'
        )
    else:
        run_kwargs = {
            "font_ascii": "黑体",
            "font_hansi": "黑体",
            "font_eastasia": "黑体",
            "font_hint": "eastAsia",
            "size": 36,
        }
        ppr_extra = (
            '<w:snapToGrid w:val="0"/>'
            + spacing_xml(before=390, after=312, line=360, line_rule="auto")
            + '<w:rPr><w:rFonts w:ascii="黑体" w:hAnsi="黑体" w:eastAsia="黑体" w:hint="eastAsia"/>'
            '<w:sz w:val="36"/><w:szCs w:val="36"/></w:rPr>'
        )

    if english and text.strip().upper() == "ABSTRACT":
        # 哈工大规范 1.3(3)：英文摘要标题首字母大写 "Abstract"（非全大写）
        text = "Abstract"
    paragraph = formatted_paragraph_xml(
        text,
        align="center",
        ppr_extra=ppr_extra,
        run_kwargs=run_kwargs,
    )
    return add_page_break_before_paragraph_xml(paragraph) if page_break_before else paragraph


def build_body_paragraph(
    text: str,
    *,
    english: bool = False,
    math_converter: MathConverter | None = None,
    reference_anchors: dict[str, str] | None = None,
) -> str:
    if english:
        run_kwargs = {"size": 24, "size_cs": False}
        if _contains_cjk_text(text):
            run_kwargs["font_hint"] = "eastAsia"
        paragraph_mark = '<w:rPr><w:sz w:val="24"/></w:rPr>'
    else:
        run_kwargs = {
            "font_ascii": "Times New Roman",
            "font_hansi": "Times New Roman",
            "font_eastasia": "宋体",
            "size": 24,
            "size_cs": False,
        }
        paragraph_mark = (
            '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/>'
            '<w:sz w:val="24"/></w:rPr>'
        )
    ppr_extra = spacing_xml(line=360) + paragraph_mark

    return paragraph_with_inline_math_xml(
        text,
        ppr_extra=ppr_extra,
        first_line_chars=200,
        first_line=480,
        preserve_breaks=True,
        run_kwargs=run_kwargs,
        math_converter=math_converter,
        reference_anchors=reference_anchors,
    )


def build_keyword_paragraph(keywords: str, *, english: bool = False) -> str | None:
    if not keywords:
        return None
    if english:
        # 去掉可能残留的 “KEY WORDS / Keywords” 标签（源 md 写法不一），
        # 并以半角逗号分隔英文关键词（模板要求 Keywords 小写加粗、逗号隔开）。
        kw = re.sub(
            r'^\s*(?:\*\*)?(?:KEY\s*WORDS|Keywords)\b\s*[:：]\s*',
            '',
            keywords,
            flags=re.IGNORECASE,
        )
        parts = re.split(r'[;；、,]+', kw)
        parts = [p.strip() for p in parts if p.strip()]
        kw = ', '.join(parts)
        runs = [
            run_text_xml("Keywords", bold=True, size=24, size_cs=False),
            run_text_xml(": ", bold=True, size=24, size_cs=False),
            run_text_xml(kw, size=24, size_cs=False),
        ]
    else:
        run_kwargs = {
            "font_ascii": "Times New Roman",
            "font_hansi": "Times New Roman",
            "font_eastasia": "宋体",
            "size": 24,
            "size_cs": False,
        }
        runs = [
            run_text_xml("关键词", bold=True, **run_kwargs),
            run_text_xml("：", bold=True, **run_kwargs),
            run_text_xml(keywords, **run_kwargs),
        ]
    ppr_extra = spacing_xml(line=360)
    if not english:
        ppr_extra += (
            '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/>'
            '<w:sz w:val="24"/></w:rPr>'
        )
    return paragraph_xml(runs=runs, ppr_extra=ppr_extra)


def build_blank_paragraph(*, style: str = STYLE_BODY, line: int = 300, run_size: int | None = None) -> str:
    if run_size is None:
        return paragraph_xml(" ", style=style, ppr_extra=spacing_xml(line=line))
    return formatted_paragraph_xml(
        " ",
        style=style,
        ppr_extra=spacing_xml(line=line),
        run_kwargs={
            "font_ascii": "Times New Roman",
            "font_hansi": "Times New Roman",
            "font_eastasia": "宋体",
            "size": run_size,
        },
    )


def build_statement_body_paragraph(
    text: str,
    *,
    math_converter: MathConverter | None = None,
    reference_anchors: dict[str, str] | None = None,
) -> str:
    run_kwargs = {
        "font_ascii": "Times New Roman",
        "font_hansi": "Times New Roman",
        "font_eastasia": "宋体",
        "size": 24,
    }
    return paragraph_with_inline_math_xml(
        text,
        style=STYLE_BODY,
        ppr_extra=spacing_xml(line=360),
        first_line_chars=200,
        first_line=480,
        run_kwargs=run_kwargs,
        math_converter=math_converter,
        reference_anchors=reference_anchors,
    )


def build_statement_signature_paragraph(
    label: str,
    value: str = "",
    *,
    is_date: bool = False,
    signature_image: MediaImage | None = None,
    media_manager: MediaManager | None = None,
    signature_alt: str = "电子签名",
) -> str:
    normalized = value.strip().strip("_").strip()
    if not normalized:
        normalized = "   年   月   日" if is_date else "\u00a0" * 7
    run_kwargs = {
        "font_ascii": "Times New Roman",
        "font_hansi": "Times New Roman",
        "font_eastasia": "宋体",
        "size": 24,
    }
    ppr_extra = spacing_xml(line=360)
    if not is_date:
        ppr_extra += indent_xml(right=280)
    if signature_image is not None and media_manager is not None:
        runs = [
            run_text_xml(label, **run_kwargs),
            image_run_xml(
                signature_image,
                docpr_id=media_manager.next_drawing_id(),
                alt_text=signature_alt,
                width_emu=SIGNATURE_IMAGE_WIDTH_EMU,
                height_emu=SIGNATURE_IMAGE_HEIGHT_EMU,
            ),
        ]
        return paragraph_xml(runs=runs, align="right", ppr_extra=ppr_extra)
    return formatted_paragraph_xml(
        f"{label}{normalized}",
        align="right",
        ppr_extra=ppr_extra,
        run_kwargs=run_kwargs,
    )


# 任务书：硕士学位论文不使用任务书，保留以兼容 document.py 导入
def build_taskbook_elements(taskbook_text: str, cover_info: dict[str, str]) -> list[str]:
    return []
