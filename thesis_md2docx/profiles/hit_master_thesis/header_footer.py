from __future__ import annotations

from ...constants import R_NS, W_NS
from ...ooxml.render import formatted_paragraph_xml, paragraph_xml
from ...ooxml.xml import field_char_run_xml, instr_text_run_xml, run_text_xml
from .styles import STYLE_FOOTER, STYLE_HEADER


def header_xml() -> str:
    # 页眉：粗/细双线（双线由 STYLE_HEADER 段落下边框实现），小5号宋体居中
    paragraph = formatted_paragraph_xml(
        "哈尔滨工业大学硕士学位论文",
        style=STYLE_HEADER,
        align="center",
        ppr_extra="",
        run_kwargs={
            "font_ascii": "Times New Roman",
            "font_hansi": "Times New Roman",
            "font_eastasia": "宋体",
            "size": 18,
        },
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:hdr xmlns:w="{W_NS}" xmlns:r="{R_NS}">{paragraph}</w:hdr>'
    )


def empty_footer_xml() -> str:
    paragraph = paragraph_xml("", style=STYLE_FOOTER, first_line=0, first_line_chars=0)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:ftr xmlns:w="{W_NS}" xmlns:r="{R_NS}">{paragraph}</w:ftr>'
    )


def title_header_xml(title: str) -> str:
    # 每章首页页眉：显示章节标题，复用 HitHeader 样式（含"上粗下细"下边框）
    paragraph = formatted_paragraph_xml(
        title,
        style=STYLE_HEADER,
        align="center",
        ppr_extra="",
        run_kwargs={
            "font_ascii": "Times New Roman",
            "font_hansi": "Times New Roman",
            "font_eastasia": "宋体",
            "size": 18,
        },
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:hdr xmlns:w="{W_NS}" xmlns:r="{R_NS}">{paragraph}</w:hdr>'
    )


def empty_header_xml() -> str:
    paragraph = paragraph_xml("", first_line=0, first_line_chars=0)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:hdr xmlns:w="{W_NS}" xmlns:r="{R_NS}">{paragraph}</w:hdr>'
    )


def page_footer_xml() -> str:
    runs = [
        field_char_run_xml("begin"),
        instr_text_run_xml("PAGE   \\* MERGEFORMAT"),
        field_char_run_xml("separate"),
        run_text_xml(
            "1",
            font_ascii="Times New Roman",
            font_hansi="Times New Roman",
            font_eastasia="宋体",
            size=18,
        ),
        field_char_run_xml("end"),
    ]
    paragraph = paragraph_xml(runs=runs, style=STYLE_FOOTER, align="center", first_line=0, first_line_chars=0)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:ftr xmlns:w="{W_NS}" xmlns:r="{R_NS}">{paragraph}</w:ftr>'
    )
