from __future__ import annotations

from ...layout import StyleBundle
from ...ooxml.parts import settings_xml
from ...styles import DocumentDefaultsSpec, StyleCatalog, StyleRole, StyleRoleMap, StyleSpec
from ...styles.fonts import FontSpec, FontTableSpec, font_table_xml
from ...styles.numbering import (
    AbstractNumberingSpec,
    NumberingCatalog,
    NumberingInstanceSpec,
    NumberingLevelSpec,
    numbering_xml,
)
from ...styles.ooxml import styles_xml as render_styles_xml
from ...styles.properties import (
    bold,
    bold_complex,
    indent,
    italic,
    justification,
    kern,
    numbering,
    outline_level,
    paragraph_bottom_border,
    run_fonts,
    run_size,
    shading,
    snap_to_grid,
    spacing,
    style_props,
    tab,
    tabs,
    widow_control,
)

# 字号（半磅 half-point）：二号=44，小3=30，4号=28，小4=24，小2=32，5号=20，小5=18
STYLE_BODY = "HitBody"
STYLE_HEADING_1 = "HitHeading1"
STYLE_HEADING_2 = "HitHeading2"
STYLE_HEADING_3 = "HitHeading3"
STYLE_FRONT_HEADING = "HitFrontHeading"
STYLE_TOC_FIELD = "HitTocField"
STYLE_CAPTION = "HitCaption"
STYLE_REFERENCE = "HitReference"
STYLE_QUOTE = "HitQuote"
STYLE_CODE_BLOCK = "HitCodeBlock"
STYLE_MATH_BLOCK = "HitMathBlock"
STYLE_FIGURE_IMAGE = "HitFigureImage"
STYLE_TABLE_TEXT = "HitTableText"
STYLE_HEADER = "HitHeader"
STYLE_FOOTER = "HitFooter"

# 多倍行距 1.5 -> 360/240
LINE_MULT = 360


def hit_numbering_catalog() -> NumberingCatalog:
    # 哈工大硕士格式：一级“第一章”、二级“一、”、三级“（一）”，均用中文数字
    return NumberingCatalog(
        abstract_numbers=(
            AbstractNumberingSpec(
                abstract_num_id=0,
                levels=(
                    NumberingLevelSpec(
                        level=0,
                        text="第%1章 ",
                        paragraph_style=STYLE_HEADING_1,
                        num_format="chineseCounting",
                        suffix="space",
                    ),
                    NumberingLevelSpec(
                        level=1,
                        text="%2、",
                        paragraph_style=STYLE_HEADING_2,
                        num_format="chineseCounting",
                        suffix="nothing",
                    ),
                    NumberingLevelSpec(
                        level=2,
                        text="（%3）",
                        paragraph_style=STYLE_HEADING_3,
                        num_format="chineseCounting",
                        suffix="nothing",
                    ),
                ),
            ),
        ),
        instances=(NumberingInstanceSpec(num_id=1, abstract_num_id=0),),
    )


def hit_numbering_xml() -> str:
    return numbering_xml(hit_numbering_catalog())


def hit_font_table() -> FontTableSpec:
    return FontTableSpec(
        fonts=(
            FontSpec("Times New Roman"),
            FontSpec("宋体", alt_name="SimSun", charset="86", family="auto", pitch="variable"),
            FontSpec("黑体", alt_name="SimHei", charset="86", family="modern", pitch="fixed"),
            FontSpec("楷体_GB2312", alt_name="楷体", charset="86", family="modern", pitch="default"),
            FontSpec("Cambria Math"),
            FontSpec("Courier New"),
            FontSpec("等线", alt_name="DengXian"),
        ),
    )


def hit_font_table_xml() -> str:
    return font_table_xml(hit_font_table())


def hit_style_roles() -> StyleRoleMap:
    return StyleRoleMap(
        {
            StyleRole.BASE_NORMAL: "Normal",
            StyleRole.BODY_TITLE: STYLE_HEADING_1,
            StyleRole.BODY_NORMAL: STYLE_BODY,
            StyleRole.BODY_HEADING_LEVEL1: STYLE_HEADING_1,
            StyleRole.BODY_HEADING_LEVEL2: STYLE_HEADING_2,
            StyleRole.BODY_HEADING_LEVEL3: STYLE_HEADING_3,
            StyleRole.FRONT_HEADING: STYLE_FRONT_HEADING,
            StyleRole.TOC_FIELD: STYLE_TOC_FIELD,
            StyleRole.TOC_LEVEL1: "TOC1",
            StyleRole.TOC_LEVEL2: "TOC2",
            StyleRole.TOC_LEVEL3: "TOC3",
            StyleRole.CAPTION_DEFAULT: STYLE_CAPTION,
            StyleRole.REFERENCE_ITEM: STYLE_REFERENCE,
            StyleRole.QUOTE_BLOCK: STYLE_QUOTE,
            StyleRole.CODE_BLOCK: STYLE_CODE_BLOCK,
            StyleRole.MATH_BLOCK: STYLE_MATH_BLOCK,
            StyleRole.BODY_IMAGE: STYLE_FIGURE_IMAGE,
            StyleRole.TABLE_CELL: STYLE_TABLE_TEXT,
            StyleRole.HEADER_DEFAULT: STYLE_HEADER,
            StyleRole.FOOTER_DEFAULT: STYLE_FOOTER,
        }
    )


def hit_style_catalog() -> StyleCatalog:
    return StyleCatalog(
        defaults=DocumentDefaultsSpec(
            run_props=style_props(
                run_fonts(ascii="Times New Roman", hansi="Times New Roman", eastasia="宋体"),
                run_size(24),
            ),
            # 关闭文档网格吸附：每个段落使用自身 spacing（1.25倍=300twips），
            # 而非被 docGrid 吸成 391/384，解决"行距设置一致但视觉更挤"的问题。
            # 对照研究生学位论文写作模板(2020版).doc：其所有段落均 snapToGrid=0。
            paragraph_props=style_props(snap_to_grid(False)),
        ),
        styles=(
            StyleSpec(
                style_id="Normal",
                name="Normal",
                default=True,
                paragraph_props=style_props(widow_control(False), justification("both")),
                run_props=style_props(
                    run_fonts(
                        ascii="Times New Roman",
                        hansi="Times New Roman",
                        eastasia="宋体",
                        complex_script="Times New Roman",
                    ),
                    kern(2),
                    run_size(21, complex_size=24),
                ),
            ),
            StyleSpec(
                style_id=STYLE_BODY,
                name="HIT Body",
                based_on="Normal",
                q_format=True,
                paragraph_props=style_props(
                    widow_control(False),
                    justification("both"),
                    spacing(after=0, line=LINE_MULT),
                    indent(first_line_chars=200, first_line=480),
                ),
                run_props=style_props(
                    run_fonts(ascii="Times New Roman", hansi="Times New Roman", eastasia="宋体"),
                    kern(2),
                    run_size(24),
                ),
            ),
            StyleSpec(
                style_id=STYLE_HEADING_1,
                name="HIT Heading 1",
                based_on="Normal",
                next_style="Normal",
                q_format=True,
                paragraph_props=style_props(
                    "<w:keepNext/>",
                    "<w:keepLines/>",
                    numbering(num_id=1),
                    spacing(before=390, after=312, line=360, line_rule="auto"),
                    justification("center"),
                    outline_level(1),
                ),
                run_props=style_props(
                    bold_complex(),
                    snap_to_grid(False),
                    run_fonts(ascii="Times New Roman", hansi="Times New Roman", eastasia="黑体"),
                    kern(44),
                    run_size(36, complex_size=36),
                ),
            ),
            StyleSpec(
                style_id=STYLE_HEADING_2,
                name="HIT Heading 2",
                based_on=STYLE_HEADING_1,
                next_style="Normal",
                q_format=True,
                paragraph_props=style_props(
                    numbering(level=1),
                    spacing(before=195, after=195, line=360, line_rule="auto"),
                    justification("both"),
                    outline_level(2),
                ),
                run_props=style_props(
                    bold_complex(),
                    run_fonts(ascii="Times New Roman", hansi="Times New Roman", eastasia="黑体"),
                    run_size(30, complex_size=30),
                ),
            ),
            StyleSpec(
                style_id=STYLE_HEADING_3,
                name="HIT Heading 3",
                based_on=STYLE_HEADING_2,
                next_style="Normal",
                q_format=True,
                paragraph_props=style_props(
                    numbering(level=2),
                    spacing(before=195, after=195, line=360, line_rule="auto"),
                    outline_level(3),
                ),
                run_props=style_props(
                    bold_complex(),
                    run_fonts(ascii="Times New Roman", hansi="Times New Roman", eastasia="黑体"),
                    run_size(28, complex_size=28),
                ),
            ),
            StyleSpec(
                style_id=STYLE_FRONT_HEADING,
                name="HIT Front Heading",
                based_on="Normal",
                q_format=True,
                paragraph_props=style_props(
                    justification("center"),
                    spacing(before=390, after=312, line=360, line_rule="auto"),
                ),
                run_props=style_props(
                    run_fonts(ascii="Times New Roman", hansi="Times New Roman", eastasia="黑体"),
                    run_size(36),
                ),
            ),
            StyleSpec(
                style_id="afa",
                name="目录标题",
                based_on="Normal",
                q_format=True,
                paragraph_props=style_props(
                    spacing(before=390, after=312, line=360, line_rule="auto"),
                    justification("center"),
                ),
                run_props=style_props(run_size(36)),
            ),
            StyleSpec(
                style_id=STYLE_TOC_FIELD,
                name="HIT TOC Field",
                based_on="Normal",
                paragraph_props=style_props(spacing(after=0, line=LINE_MULT)),
                run_props=style_props(
                    run_fonts(ascii="Times New Roman", hansi="Times New Roman", eastasia="宋体"),
                    run_size(24),
                ),
            ),
            StyleSpec(
                style_id="TOC1",
                name="toc 1",
                based_on="Normal",
                paragraph_props=style_props(
                    tabs(tab("right", 8303, leader="dot")),
                    spacing(after=0, line=LINE_MULT),
                ),
                run_props=style_props(
                    run_fonts(ascii="Times New Roman", hansi="Times New Roman", eastasia="黑体"),
                    run_size(24),
                ),
            ),
            StyleSpec(
                style_id="TOC2",
                name="toc 2",
                based_on="Normal",
                paragraph_props=style_props(
                    tabs(tab("right", 8303, leader="dot")),
                    indent(left=210),
                    spacing(after=0, line=LINE_MULT),
                ),
                run_props=style_props(
                    run_fonts(ascii="Times New Roman", hansi="Times New Roman", eastasia="宋体"),
                    run_size(24),
                ),
            ),
            StyleSpec(
                style_id="TOC3",
                name="toc 3",
                based_on="Normal",
                paragraph_props=style_props(
                    tabs(tab("right", 8303, leader="dot")),
                    indent(left=420),
                    spacing(after=0, line=LINE_MULT),
                ),
                run_props=style_props(
                    run_fonts(ascii="Times New Roman", hansi="Times New Roman", eastasia="宋体"),
                    run_size(24),
                ),
            ),
            StyleSpec(
                style_id=STYLE_HEADER,
                name="HIT Header",
                based_on="Normal",
                paragraph_props=style_props(
                    # 上粗下细：粗线在上、细线在下。OOXML 中 thinThickMediumGap =
                    # “a thin line below a thick line”，即粗线靠近正文（上方）、
                    # 细线在下方；thickThinMediumGap 反之（细上粗下），故此处用 thinThick。
                    paragraph_bottom_border(value="thinThickMediumGap", size=18),
                    tabs(tab("center", 4153), tab("right", 8306)),
                    snap_to_grid(False),
                    justification("center"),
                ),
                run_props=style_props(
                    run_fonts(ascii="宋体", hansi="宋体", eastasia="宋体"),
                    run_size(18),
                ),
            ),
            StyleSpec(
                style_id=STYLE_FOOTER,
                name="HIT Footer",
                based_on="Normal",
                paragraph_props=style_props(
                    tabs(tab("center", 4153), tab("right", 8306)),
                    snap_to_grid(False),
                    spacing(line=LINE_MULT),
                    indent(first_line_chars=200, first_line=200),
                    justification("center"),
                ),
                run_props=style_props(
                    run_fonts(ascii="Times New Roman", hansi="Times New Roman", eastasia="宋体"),
                    run_size(18),
                ),
            ),
            StyleSpec(
                style_id=STYLE_REFERENCE,
                name="HIT Reference",
                based_on="Normal",
                paragraph_props=style_props(
                    '<w:overflowPunct w:val="0"/>',
                    '<w:autoSpaceDN w:val="0"/>',
                    '<w:adjustRightInd w:val="0"/>',
                    snap_to_grid(False),
                    spacing(line=LINE_MULT),
                ),
                run_props=style_props(
                    run_fonts(complex_script="Courier New"),
                    snap_to_grid(False),
                    kern(0),
                    '<w:szCs w:val="21"/>',
                ),
            ),
            StyleSpec(
                style_id=STYLE_QUOTE,
                name="HIT Quote",
                based_on=STYLE_BODY,
                paragraph_props=style_props(
                    indent(left=720),
                    spacing(after=120, line=LINE_MULT),
                ),
                run_props=style_props(italic()),
            ),
            StyleSpec(
                style_id=STYLE_CODE_BLOCK,
                name="HIT Code Block",
                based_on="Normal",
                paragraph_props=style_props(
                    spacing(after=120),
                    shading("clear", "F5F5F5"),
                ),
                run_props=style_props(
                    run_fonts(ascii="Courier New", hansi="Courier New", eastasia="等线"),
                    run_size(20),
                ),
            ),
            StyleSpec(
                style_id=STYLE_MATH_BLOCK,
                name="公式",
                based_on="Normal",
                paragraph_props=style_props(
                    tabs(tab("right", 8971)),
                    spacing(before_lines=50, before=120, after_lines=50, after=120),
                    indent(first_line_chars=200, first_line=480),
                ),
                run_props=style_props(
                    run_size(24, complex_size=21),
                ),
            ),
            StyleSpec(
                style_id=STYLE_FIGURE_IMAGE,
                name="HIT Figure Image",
                based_on="Normal",
                paragraph_props=style_props(
                    justification("center"),
                    spacing(before_lines=50, before=50, after_lines=50, after=50, line=LINE_MULT),
                ),
                run_props=style_props(run_size(21)),
            ),
            StyleSpec(
                style_id=STYLE_CAPTION,
                name="HIT Caption",
                based_on="Normal",
                paragraph_props=style_props(
                    justification("center"),
                    spacing(before_lines=0, after_lines=0, line=288, line_rule="auto"),
                    indent(left=0, first_line=0),
                ),
                run_props=style_props(
                    run_fonts(ascii="Times New Roman", hansi="Times New Roman", eastasia="宋体"),
                    run_size(21),
                ),
            ),
            StyleSpec(
                style_id=STYLE_TABLE_TEXT,
                name="HIT Table Text",
                based_on="Normal",
                paragraph_props=style_props(spacing(after=0, line=LINE_MULT)),
                run_props=style_props(
                    run_fonts(ascii="Times New Roman", hansi="Times New Roman", eastasia="宋体"),
                    run_size(21),
                ),
            ),
        ),
    )


def hit_styles_xml() -> str:
    return render_styles_xml(hit_style_catalog())


def hit_style_bundle() -> StyleBundle:
    from .header_footer import empty_footer_xml, empty_header_xml, header_xml, page_footer_xml

    return StyleBundle(
        styles_xml=hit_styles_xml(),
        numbering_xml=hit_numbering_xml(),
        settings_xml=settings_xml(),
        font_table_xml=hit_font_table_xml(),
        header_xml=header_xml(),
        empty_header_xml=empty_header_xml(),
        empty_footer_xml=empty_footer_xml(),
        page_footer_xml=page_footer_xml(),
    )
