from __future__ import annotations

import re
import shutil
import subprocess
import tempfile
from pathlib import Path

from ...builders.document import (
    build_document_elements,
    chinese_cardinal,
    collect_toc_entries,
)
from ...constants import (
    REL_ID_EMPTY_FOOTER,
    REL_ID_EMPTY_HEADER,
    REL_ID_PAGE_FOOTER,
    TITLE_HEADER_RID_BASE,
)
from .body import strip_heading_prefix
from .frontmatter import (
    build_blank_paragraph,
    build_body_paragraph,
    build_cover_elements,
    build_front_heading,
    build_keyword_paragraph,
    build_statement_body_paragraph,
    build_statement_signature_paragraph,
    build_taskbook_elements,
    XJU_DECLARATION_SIGNATURE,
)
from .header_footer import title_header_xml
from ...frontmatter import parse_inline_image_value, split_statement_content
from ...layout import FrontMatterPageSpec
from ...markdown import (
    extract_abstract_keyword_blocks,
    parse_cover_info,
    parse_markdown_document,
)
from ...math.converter import MathConverter
from ...media import MediaManager
from ...ooxml.render import (
    add_section_to_paragraph_xml,
    bookmark_paragraph_xml,
    formatted_paragraph_xml,
    image_run_xml,
    page_break_xml,
    paragraph_xml,
    section_break_paragraph_xml,
    toc_cache_entry_paragraph_xml,
    toc_field_paragraph_xml,
)
from ...ooxml.parts import native_sect_pr_xml
from ...ooxml.xml import spacing_xml
from ...ooxml.xml import indent_xml
from ...styles import StyleRole
from ...toc import TocEntry, make_toc_entry
from ..base import ThesisProfile

A4_WIDTH_TWIPS = 11907
A4_HEIGHT_TWIPS = 16840
A4_WIDTH_EMU = A4_WIDTH_TWIPS * 635
A4_HEIGHT_EMU = A4_HEIGHT_TWIPS * 635
EXTERNAL_TASKBOOK_PDF_KEYS = (
    "任务书PDF",
    "任务书 PDF",
    "外部任务书PDF",
    "外部任务书 PDF",
    "任务书文件",
)
PDF_PAGE_IMAGE_DPI = 300


def _front_text(front_sections: dict[str, str], page: FrontMatterPageSpec) -> str:
    if page.source_key is None:
        return ""
    key = page.source_key
    if key in front_sections:
        return front_sections[key].strip()
    # 大小写不敏感兜底：源 md 里前置标题可能写成 “Abstract” 而非 “ABSTRACT”，
    # 直接按 source_key 精确查找会取空，导致英文摘要整页（含分节符）被吞掉。
    lk = key.lower()
    for k, v in front_sections.items():
        if k.lower() == lk:
            return v.strip()
    return ""


# ---------------------------------------------------------------------------
# 原创性声明和使用权限（后置，位于致谢之前）
# 内容直接照抄哈工大《学位论文书写范例》，与论文题目无关，故作为固定模板渲染。
# ---------------------------------------------------------------------------


HIT_DECLARATION_TITLE = "哈尔滨工业大学学位论文原创性声明和使用权限"


def _hit_decl_subheading(text: str) -> str:
    """声明内的二级小标题（如“学位论文原创性声明”），居中黑体。"""
    return formatted_paragraph_xml(
        text,
        align="center",
        ppr_extra=spacing_xml(before=240, after=120, line=360),
        run_kwargs={
            "font_ascii": "Times New Roman",
            "font_hansi": "Times New Roman",
            "font_eastasia": "黑体",
            "size": 28,
        },
    )


def _hit_signature_line(label: str, date: str = "日期：    年   月   日") -> str:
    """右对齐的签名行：作者签名：……日期：……年 月 日。"""
    return formatted_paragraph_xml(
        f"{label}\u3000\u3000\u3000\u3000\u3000\u3000\u3000\u3000\u3000\u3000\u3000\u3000\u3000\u3000{date}",
        align="right",
        ppr_extra=spacing_xml(line=360),
        run_kwargs={
            "font_ascii": "Times New Roman",
            "font_hansi": "Times New Roman",
            "font_eastasia": "宋体",
            "size": 24,
        },
    )


def _hit_declaration_elements(thesis_title: str) -> list[str]:
    """生成“学位论文原创性声明和使用权限”整页元素，照抄哈工大模板。"""
    els: list[str] = []
    els.append(build_front_heading(HIT_DECLARATION_TITLE, statement=True))
    els.append(_hit_decl_subheading("学位论文原创性声明"))
    els.append(
        build_statement_body_paragraph(
            f"本人郑重声明：此处所提交的学位论文《{thesis_title}》，是本人在导师指导下，"
            "在哈尔滨工业大学攻读学位期间独立进行研究工作所取得的成果，且学位论文中除已标注"
            "引用文献的部分外不包含他人完成或已发表的研究成果。对本学位论文的研究工作做出"
            "重要贡献的个人和集体，均已在文中以明确方式注明。"
        )
    )
    els.append(_hit_signature_line("作者签名："))
    els.append(_hit_decl_subheading("学位论文使用权限"))
    els.append(
        build_statement_body_paragraph(
            "学位论文是研究生在哈尔滨工业大学攻读学位期间完成的成果，知识产权归属哈尔滨工业大学。"
            "学位论文的使用权限如下："
        )
    )
    els.append(
        build_statement_body_paragraph(
            "（1）学校可以采用影印、缩印或其他复制手段保存研究生上交的学位论文，并向国家图书馆报送"
            "学位论文；（2）学校可以将学位论文部分或全部内容编入有关数据库进行检索和提供相应阅览"
            "服务；（3）研究生毕业后发表与此学位论文研究成果相关的学术论文和其他成果时，应征得导师"
            "同意，且第一署名单位为哈尔滨工业大学。"
        )
    )
    els.append(
        build_statement_body_paragraph("保密论文在保密期内遵守有关保密规定，解密后适用于此使用权限规定。")
    )
    els.append(build_statement_body_paragraph("本人知悉学位论文的使用权限，并将遵守有关规定。"))
    els.append(_hit_signature_line("作者签名："))
    els.append(_hit_signature_line("导师签名："))
    return els


# 哈工大硕士论文：后置“攻读硕士学位期间取得的科研成果”节（即博士论文里的
# “创新性成果”节）。本文件此前误用博士措辞，已纠正。集中为常量避免散落笔误。
HIT_MASTER_ACHIEVEMENTS_TITLE = "攻读硕士学位期间取得的科研成果"


def _md_has_heading(text: str, title: str) -> bool:
    """判断正文 markdown 是否已用任意层级标题提供某后置节（避免引擎重复插入）。"""
    want = title.strip()
    for line in text.splitlines():
        m = re.match(r"^#{1,3}\s+(.+?)\s*$", line)
        if m and m.group(1).strip() == want:
            return True
    return False



def _innovative_achievements_elements() -> list[str]:
    """攻读硕士学位期间取得的科研成果（占位，请补充具体内容）。"""
    return [
        build_body_paragraph(
            "（攻读硕士学位期间取得的科研成果，如已发表的学术论文、授权专利、科研获奖、"
            "参与的科研项目等，请在此处补充。）"
        )
    ]


def _acknowledgement_elements() -> list[str]:
    """致谢（占位，请补充具体内容）。"""
    return [
        build_body_paragraph(
            "（衷心感谢导师在论文选题、研究方法与写作过程中的悉心指导，感谢同门与好友的"
            "支持，也感谢家人在求学路上的陪伴。请在此处补充完整的致谢内容。）"
        )
    ]


def _resume_elements() -> list[str]:
    """个人简历（占位，请补充具体内容）。"""
    return [
        build_body_paragraph(
            "（个人简历：包括出生年月、籍贯、教育经历、研究方向等。请在此处补充。）"
        )
    ]


def _append_post_reference_sections(
    elements: list[str],
    *,
    thesis_profile: "ThesisProfile",
    profile: "BodyRenderProfile",
    thesis_title: str,
    post_ref_entries: list[TocEntry],
    md_provided_titles: set[str] | None = None,
    title_header_rids: dict[str, str] | None = None,
    body_sect_pr: str | None = None,
) -> str | None:
    """在参考文献之后按模板顺序插入后置部分，各部分均另起一节。

    每节首页页眉显示该节标题（原创性声明 / 致谢 / 个人简历 / 科研成果）。
    第一个后置节之前会插入 body_sect_pr 以结束正文最后一节；返回的 sectPr
    用于最后一个后置节，调用方应将其作为文档末尾的 sectPr 写入。

    若作者已在正文中提供某后置节，则该节已由正文渲染、目录也已登记，此处
    跳过引擎占位，避免重复。decl_title（原创性声明）始终由生成器插入。
    """
    decl_title = thesis_profile.front_matter_spec().declaration_title
    md_provided = md_provided_titles or set()
    title_header_rids = title_header_rids or {}
    layout = thesis_profile.document_layout()
    post_ref_titles = [
        HIT_MASTER_ACHIEVEMENTS_TITLE,
        decl_title,
        "致谢",
        "个人简历",
    ]
    titles_to_insert = [
        title for title in post_ref_titles
        if title not in md_provided or title == decl_title
    ]
    if not titles_to_insert:
        return None

    # 延迟模型：每个分节符描述的是“以它结尾的那一节”。
    # 第一个分节符用 body_sect_pr 描述正文最后一节；第二个分节符描述第一个后置节；
    # 以此类推；最终返回的 sectPr 描述最后一个后置节。
    for i, title in enumerate(titles_to_insert):
        if i == 0:
            # 正文最后一节（如参考文献）的 sectPr 作为第一个后置节之前的分节符
            if body_sect_pr:
                elements.append(section_break_paragraph_xml(body_sect_pr))
        else:
            # 描述上一个后置节
            prev_rid = title_header_rids.get(titles_to_insert[i - 1])
            elements.append(
                section_break_paragraph_xml(
                    thesis_profile.section_from_spec(
                        layout.body_continue,
                        with_title_header=bool(prev_rid),
                        title_header_rid=prev_rid,
                    )
                )
            )
        entry = next((e for e in post_ref_entries if e.text == title), None)
        if title == decl_title:
            decl = _hit_declaration_elements(thesis_title)
            decl_heading = decl[0]
            if entry is not None:
                decl_heading = bookmark_paragraph_xml(
                    decl_heading, bookmark_id=entry.bookmark_id, anchor=entry.anchor
                )
            elements.append(decl_heading)
            elements.extend(decl[1:])
            continue
        heading_xml = profile.heading_builder(title, 1, profile, numbered=False, keep_with_next=True)
        if entry is not None:
            heading_xml = bookmark_paragraph_xml(
                heading_xml, bookmark_id=entry.bookmark_id, anchor=entry.anchor
            )
        elements.append(heading_xml)
        if title == HIT_MASTER_ACHIEVEMENTS_TITLE:
            elements.extend(_innovative_achievements_elements())
        elif title == "致谢":
            elements.extend(_acknowledgement_elements())
        elif title == "个人简历":
            elements.extend(_resume_elements())

    # 最终 sectPr 描述最后一个后置节
    last_rid = title_header_rids.get(titles_to_insert[-1])
    return thesis_profile.section_from_spec(
        layout.body_continue,
        with_title_header=bool(last_rid),
        title_header_rid=last_rid,
    )


def _append_cover_page(
    elements: list[str],
    *,
    thesis_title: str,
    cover_info: dict[str, str],
    cover_sect: str,
    cover_assets_dir: Path | None,
    media_manager: MediaManager | None,
) -> None:
    elements.extend(
        build_cover_elements(
            thesis_title,
            cover_info,
            cover_assets_dir=cover_assets_dir,
            media_manager=media_manager,
        )
    )
    # 封面占据物理第 1/3/5 页，第 2/4/6 页为空白（由 build_cover_elements 内部
    # 的分页符实现）。此处仅追加一个“纯分节”段落（不含额外分页）以结束封面节，
    # cover_sect 的 nextPage 会把中文摘要推到第 7 页，并从此处起以罗马数字 I 编页。
    elements.append(section_break_paragraph_xml(cover_sect))


def _append_declaration_page(
    elements: list[str],
    *,
    page: FrontMatterPageSpec,
    declaration: str,
    math_converter: MathConverter | None,
    reference_anchors: dict[str, str] | None,
    markdown_dir: Path | None,
    media_manager: MediaManager | None,
    section_break_after: str | None = None,
) -> None:
    if not declaration:
        return
    elements.append(build_front_heading(page.title, statement=True))
    statement_paragraphs, author_value, date_value = split_statement_content(declaration)
    for paragraph in statement_paragraphs:
        elements.append(
            build_statement_body_paragraph(
                paragraph,
                math_converter=math_converter,
                reference_anchors=reference_anchors,
            )
        )
    signature_image = None
    signature_alt = XJU_DECLARATION_SIGNATURE.signature_alt
    inline_signature = parse_inline_image_value(author_value)
    if inline_signature is not None and media_manager is not None and markdown_dir is not None:
        signature_alt, signature_target = inline_signature
        signature_image = media_manager.register_image(markdown_dir / signature_target)
        if signature_image is not None:
            author_value = ""
    for _ in range(XJU_DECLARATION_SIGNATURE.blank_count(has_signature_image=signature_image is not None)):
        elements.append(build_blank_paragraph(run_size=24))
    elements.append(
        build_statement_signature_paragraph(
            XJU_DECLARATION_SIGNATURE.author_label,
            author_value,
            signature_image=signature_image,
            media_manager=media_manager,
            signature_alt=signature_alt or XJU_DECLARATION_SIGNATURE.signature_alt,
        )
    )
    elements.append(
        build_statement_signature_paragraph(XJU_DECLARATION_SIGNATURE.date_label, date_value, is_date=True)
    )
    if page.page_break_after:
        if section_break_after:
            elements.append(section_break_paragraph_xml(section_break_after))
        else:
            elements.append(page_break_xml())


def _append_taskbook_page(
    elements: list[str],
    *,
    taskbook: str,
    cover_info: dict[str, str],
) -> tuple[bool, bool]:
    if not taskbook:
        return False, False
    elements.extend(build_taskbook_elements(taskbook, cover_info))
    return True, False


def _taskbook_image_section_xml() -> str:
    return (
        '<w:sectPr><w:type w:val="nextPage"/>'
        f'<w:headerReference w:type="default" r:id="{REL_ID_EMPTY_HEADER}"/>'
        f'<w:footerReference w:type="default" r:id="{REL_ID_EMPTY_FOOTER}"/>'
        '<w:pgNumType w:fmt="upperRoman"/>'
        f'<w:pgSz w:w="{A4_WIDTH_TWIPS}" w:h="{A4_HEIGHT_TWIPS}"/>'
        '<w:pgMar w:top="0" w:right="0" w:bottom="0" w:left="0" '
        'w:header="0" w:footer="0" w:gutter="0"/>'
        '<w:cols w:space="720"/>'
        '<w:docGrid w:linePitch="384"/>'
        '</w:sectPr>'
    )


def _front_matter_continue_section_xml(thesis_profile: ThesisProfile) -> str:
    return thesis_profile.section_pr_xml(
        with_header=True,
        footer_kind="page",
        section_type="nextPage",
        page_number_format="upperRoman",
    )


def _resolve_external_taskbook_pdf(taskbook: str, markdown_dir: Path | None) -> Path | None:
    task_info = parse_cover_info(taskbook)
    raw_path = ""
    for key in EXTERNAL_TASKBOOK_PDF_KEYS:
        raw_path = task_info.get(key, "").strip()
        if raw_path:
            break
    if not raw_path:
        return None
    path = Path(raw_path).expanduser()
    if not path.is_absolute() and markdown_dir is not None:
        path = markdown_dir / path
    path = path.resolve()
    if path.suffix.lower() != ".pdf":
        raise ValueError(f"外部任务书必须使用 PDF 文件：{path}")
    if not path.is_file():
        raise FileNotFoundError(f"外部任务书 PDF 不存在：{path}")
    return path


def _pdf_page_sort_key(path: Path) -> tuple[int, str]:
    match = re.search(r"-(\d+)\.png$", path.name)
    return (int(match.group(1)) if match else 0, path.name)


def _render_pdf_pages_to_png(pdf_path: Path) -> list[Path]:
    pdftoppm = shutil.which("pdftoppm")
    if not pdftoppm:
        raise RuntimeError("pdftoppm not found; install poppler-utils or add pdftoppm to PATH")
    output_dir = Path(tempfile.mkdtemp(prefix="xju-taskbook-pdf-"))
    output_prefix = output_dir / "page"
    result = subprocess.run(
        [pdftoppm, "-png", "-r", str(PDF_PAGE_IMAGE_DPI), str(pdf_path), str(output_prefix)],
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
    )
    if result.returncode != 0:
        raise RuntimeError((result.stdout or "").strip() or f"pdftoppm exited with {result.returncode}")
    pages = sorted(output_dir.glob("page-*.png"), key=_pdf_page_sort_key)
    if not pages:
        raise RuntimeError(f"未能从外部任务书 PDF 渲染出页面：{pdf_path}")
    return pages


def _taskbook_pdf_page_xml(
    image_item,
    media_manager: MediaManager,
    *,
    page_break_after: bool = False,
    section_after: str | None = None,
) -> str:
    runs = [
        image_run_xml(
            image_item,
            docpr_id=media_manager.next_drawing_id(),
            alt_text=image_item.filename,
            width_emu=A4_WIDTH_EMU,
            height_emu=A4_HEIGHT_EMU,
        )
    ]
    if page_break_after:
        runs.append('<w:r><w:br w:type="page"/></w:r>')
    paragraph = paragraph_xml(
        runs=runs,
        ppr_extra='<w:snapToGrid w:val="0"/>' + spacing_xml(before=0, after=0),
    )
    if section_after:
        paragraph = add_section_to_paragraph_xml(paragraph, section_after)
    return paragraph


def _append_external_taskbook_pdf(
    elements: list[str],
    *,
    pdf_path: Path,
    media_manager: MediaManager | None,
    section_after: str,
) -> tuple[bool, bool]:
    if media_manager is None:
        raise RuntimeError("外部任务书 PDF 需要 media_manager 才能插入 DOCX")
    page_images = _render_pdf_pages_to_png(pdf_path)
    if page_images:
        media_manager.register_temp_dir(page_images[0].parent)
    for index, image_path in enumerate(page_images):
        image_item = media_manager.register_image(image_path)
        if image_item is None:
            raise RuntimeError(f"无法注册外部任务书页面图片：{image_path}")
        is_last = index == len(page_images) - 1
        elements.append(
            _taskbook_pdf_page_xml(
                image_item,
                media_manager,
                page_break_after=not is_last,
                section_after=section_after if is_last else None,
            )
        )
    return True, True


def _append_abstract_page(
    elements: list[str],
    *,
    page: FrontMatterPageSpec,
    text: str,
    page_break_before: bool,
    math_converter: MathConverter | None,
    reference_anchors: dict[str, str] | None,
    section_break_after: str | None = None,
    toc_entry: TocEntry | None = None,
) -> None:
    kw_text = text
    if page.english:
        # 兼容源 md 中 “KEY WORDS:” / “**Keywords**:” / “**Keywords**：” 等多种写法，
        # 统一规整为 “Keywords：”，使关键词能被正确剥离（否则整行会被当作正文原样输出）。
        # 注意：关键词前后都可能有 Markdown 加粗 “**”，正则需同时吃掉两侧的 “**”。
        # 关键：用 [ \t] 而非 \s，否则 re.MULTILINE 下 \s 会吞掉关键词行前面的换行符，
        # 导致空行消失、关键词行被并入上一段，startswith 匹配失败。
        kw_text = re.sub(
            r'^[ \t]*\*{0,2}(?:KEY[ \t]*WORDS|Keywords)\b\*{0,2}[ \t]*[:：]',
            'Keywords：',
            text,
            flags=re.MULTILINE | re.IGNORECASE,
        )
    else:
        # 中文关键词同理规整 “**关键词**：” → “关键词：”，确保能走关键词样式段落。
        kw_text = re.sub(
            r'^[ \t]*\*{0,2}关键词\b\*{0,2}[ \t]*[:：]',
            '关键词：',
            text,
            flags=re.MULTILINE,
        )
    paragraphs, keywords, after_keyword = extract_abstract_keyword_blocks(kw_text, page.keyword_prefix)
    if not paragraphs and not keywords and not after_keyword:
        return
    heading_xml = build_front_heading(
        page.title,
        english=page.english,
        page_break_before=page_break_before,
    )
    if toc_entry is not None:
        heading_xml = bookmark_paragraph_xml(
            heading_xml, bookmark_id=toc_entry.bookmark_id, anchor=toc_entry.anchor
        )
    elements.append(heading_xml)
    for paragraph in paragraphs:
        elements.append(
            build_body_paragraph(
                paragraph,
                english=page.english,
                math_converter=math_converter,
                reference_anchors=reference_anchors,
            )
        )
    keyword_paragraph = build_keyword_paragraph(keywords, english=page.english)
    if keyword_paragraph:
        elements.append(build_blank_paragraph(line=360 if page.english else 405))
        elements.append(keyword_paragraph)
    for paragraph in after_keyword:
        elements.append(
            build_body_paragraph(
                paragraph,
                english=page.english,
                math_converter=math_converter,
                reference_anchors=reference_anchors,
            )
        )
    if section_break_after:
        # 仅用“分节符”（nextPage）切到下一节，去掉多余的 page_break run，
        # 避免摘要与下一节之间多出空白页。
        elements.append(section_break_paragraph_xml(section_break_after))
    elif page.page_break_after:
        elements.append(page_break_xml())


def _append_toc_page(
    elements: list[str],
    *,
    thesis_profile: ThesisProfile,
    page: FrontMatterPageSpec,
    front_sect: str,
    toc_entries: list[TocEntry] | None = None,
) -> None:
    elements.append(build_front_heading(page.title, toc=True))
    toc_style = thesis_profile.style_roles().require(StyleRole.TOC_FIELD)
    toc_level_styles = {
        1: thesis_profile.style_roles().require(StyleRole.TOC_LEVEL1),
        2: thesis_profile.style_roles().require(StyleRole.TOC_LEVEL2),
        3: thesis_profile.style_roles().require(StyleRole.TOC_LEVEL3),
    }
    if not toc_entries:
        # 没有缓存条目时：中文目录页本身即结束前置节
        elements.append(add_section_to_paragraph_xml(toc_field_paragraph_xml(style=toc_style), front_sect))
        return
    for index, entry in enumerate(toc_entries):
        is_last = index == len(toc_entries) - 1
        paragraph = toc_cache_entry_paragraph_xml(
            entry,
            first=index == 0,
            close_field=is_last,
            toc_field_style=toc_style,
            toc_level_styles=toc_level_styles,
        )
        elements.append(paragraph)
    # 中文目录末尾仅用“分节符”（nextPage）切入英文目录：去掉多余的 page_break
    # run，否则会与英文目录开头的分页叠加，导致目录与英文目录之间多出空白页。
    elements.append(section_break_paragraph_xml(front_sect))


def _append_english_toc_page(
    elements: list[str],
    *,
    thesis_profile: ThesisProfile,
    front_sect: str,
    toc_entries: list[TocEntry] | None = None,
) -> None:
    """英文目录（Contents）页：紧接中文目录之后（由中文目录末尾的分节符切入，
    已另起一页），复用中文目录的同一批 _Toc 书签，页码完全一致；编号转为英文
    样式。英文目录结尾不再挂分节符——正文第一章自身的分节符（decimal、start 1）
    会完成“罗马→阿拉伯”的切换，避免目录与正文之间多出空白页。"""
    elements.append(build_front_heading("Contents", english=True, bold=True))
    toc_style = thesis_profile.style_roles().require(StyleRole.TOC_FIELD)
    toc_level_styles = {
        1: thesis_profile.style_roles().require(StyleRole.TOC_LEVEL1),
        2: thesis_profile.style_roles().require(StyleRole.TOC_LEVEL2),
        3: thesis_profile.style_roles().require(StyleRole.TOC_LEVEL3),
    }
    if not toc_entries:
        elements.append(
            paragraph_xml(
                style=thesis_profile.style_roles().require(StyleRole.BODY_HEADING_LEVEL1),
                ppr_extra=spacing_xml(line=240) + indent_xml(left=0, first_line=0) + front_sect,
            )
        )
        return
    for index, entry in enumerate(toc_entries):
        is_last = index == len(toc_entries) - 1
        paragraph = toc_cache_entry_paragraph_xml(
            entry,
            first=index == 0,
            close_field=is_last,
            toc_field_style=toc_style,
            toc_level_styles=toc_level_styles,
            extra_ppr="",
        )
        elements.append(paragraph)
    # 英文目录（Contents）结尾也要挂分节符：让第一章从英文目录之后的新页开始，
    # 避免第一章紧跟目录。延迟模型下，该分节符描述英文目录节（upperRoman、
    # rId1012），第一章的 decimal/start=1 属性延迟到第二章前的分节符写出；
    # 同时第一章处的 append_chapter_page_break 会因 elements[-1] 已是分节符
    # 而跳过重复插入，不会引入空白页。
    elements.append(section_break_paragraph_xml(front_sect))


_TOC_NUM_RE = re.compile(r"^(\d+(?:\.\d+)*)\s+(.*)$")

# 英文目录一级标题用的英文序数词（Chapter One / Two / ...）
ENGLISH_ORDINALS = {
    1: "One", 2: "Two", 3: "Three", 4: "Four", 5: "Five",
    6: "Six", 7: "Seven", 8: "Eight", 9: "Nine", 10: "Ten",
    11: "Eleven", 12: "Twelve",
}


def _load_heading_translations(markdown_dir: Path | None) -> dict[str, str]:
    """从 markdown 所在目录加载 heading_translations.json（中英文目录标题映射）。"""
    if markdown_dir is None:
        return {}
    path = markdown_dir / "heading_translations.json"
    if not path.exists():
        return {}
    try:
        import json

        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _format_toc_chinese_number(entry: TocEntry) -> TocEntry:
    """把目录缓存条目里的十进制编号（1 / 1.1 / 1.1.1）替换成中文编号
    （第一章 / 一、 / （一）），与正文标题的编号样式保持一致。"""
    m = _TOC_NUM_RE.match(entry.text)
    if not m:
        return entry
    num_parts = m.group(1).split(".")
    label = m.group(2).strip()
    try:
        if entry.level == 1:
            prefix = f"第{chinese_cardinal(int(num_parts[0]))}章 "
        elif entry.level == 2:
            prefix = f"{chinese_cardinal(int(num_parts[1]))}、"
        else:
            prefix = f"（{chinese_cardinal(int(num_parts[2]))}）"
    except (IndexError, ValueError):
        return entry
    return TocEntry(
        level=entry.level,
        text=prefix + label,
        anchor=entry.anchor,
        bookmark_id=entry.bookmark_id,
    )


def _format_toc_english_number(
    entry: TocEntry,
    translations: dict[str, str] | None = None,
) -> TocEntry:
    """英文目录（Contents）：一级标题加粗并显示为 Chapter One，二级/三级
    保留中文编号（一、/（一）），标题文字按 heading_translations.json 翻译成英文；
    未命中时回退中文。复用与中文目录相同的 _Toc 书签，页码完全一致。"""
    translations = translations or {}
    m = _TOC_NUM_RE.match(entry.text)
    if not m:
        # 无编号的条目标题（如“参考文献”“致谢”“附录”）直接按映射翻译并加粗
        label = entry.text.strip()
        english_label = translations.get(label, label)
        if english_label == label:
            return entry
        return TocEntry(
            level=entry.level,
            text=f"**{english_label}**",
            anchor=entry.anchor,
            bookmark_id=entry.bookmark_id,
        )
    num = m.group(1).strip()  # 例如 "1" / "1.1" / "1.1.1"
    label = m.group(2).strip()
    english_label = translations.get(label, label)
    try:
        parts = num.split(".")
        if entry.level == 1:
            ch = int(parts[0])
            # 哈工大模板：最后一章固定章节名为“Conclusions”，不显示“Chapter Five”。
            if english_label == "Conclusions":
                prefix = ""
            else:
                prefix = f"Chapter {ENGLISH_ORDINALS.get(ch, ch)} "
            text = f"**{prefix}{english_label}**"
        elif entry.level == 2:
            text = f"{chinese_cardinal(int(parts[1]))}、{english_label}"
        else:
            text = f"（{chinese_cardinal(int(parts[2]))}）{english_label}"
    except (IndexError, ValueError):
        return entry
    return TocEntry(
        level=entry.level,
        text=text,
        anchor=entry.anchor,
        bookmark_id=entry.bookmark_id,
    )


def build_document(
    text: str,
    *,
    thesis_profile: ThesisProfile,
    math_converter: MathConverter | None = None,
    reference_anchors: dict[str, str] | None = None,
    markdown_dir: Path | None = None,
    cover_assets_dir: Path | None = None,
    media_manager: MediaManager | None = None,
) -> tuple[list[str], str, str]:
    markdown_title, front_sections, body_text = parse_markdown_document(text)
    front_spec = thesis_profile.front_matter_spec()
    cover_info = parse_cover_info(front_sections.get(front_spec.cover_info_key, ""))
    thesis_title = cover_info.get("论文题目") or markdown_title or front_spec.default_title
    profile = thesis_profile.body_style_profile()
    layout = thesis_profile.document_layout()

    # 封面结束处定义“前置节”：摘要与目录统一使用罗马数字页码（I, II, III…），
    # 正文（第一章）另起一节用阿拉伯数字从 1 开始。封面本身用 titlePg 隐藏页眉
    # 页脚，使封面不显示页码、摘要页起为 I。
    front_sect = thesis_profile.section_from_spec(layout.front_matter)
    # 封面独占 6 页（3 个封面 + 3 个空白页），全程无页眉页脚、无页码；
    # 摘要起（cover_sect 之后的第一节）才以罗马数字 I 开始编页码。
    cover_sect = thesis_profile.section_from_spec(layout.cover)
    front_continue_sect = _front_matter_continue_section_xml(thesis_profile)
    taskbook_image_sect = _taskbook_image_section_xml()
    body_start_sect = thesis_profile.section_from_spec(layout.body_start)
    body_continue_sect = thesis_profile.section_from_spec(layout.body_continue)

    # 前置页（摘要/目录）专用首页眉：把中文摘要、英文摘要、中文目录、英文目录
    # 各拆成独立“前置节”，首页(first)显示该页名称（摘 要 / Abstract / 目 录），
    # 其余页显示默认页眉（哈尔滨工业大学硕士学位论文）。中文摘要从罗马数字 I
    # 起（start=1）；英文摘要、目录罗马数字续编（无 start）。正文第一章另起节
    # 用阿拉伯数字从 1 开始（由 body 渲染器的分节符处理）。
    ABSTRACT_CN_RID = "rId1010"
    ABSTRACT_EN_RID = "rId1011"
    TOC_RID = "rId1012"

    def _front_title_sect(rid: str) -> str:
        # 罗马数字续编（去掉 start）：用于英文摘要、目录等继中文摘要之后的页，
        # 页码延续（II、III、IV…）而不重置。
        s = front_sect.replace(
            '<w:pgNumType w:fmt="upperRoman" w:start="1"/>',
            '<w:pgNumType w:fmt="upperRoman"/>',
        )
        s = s.replace(
            '<w:headerReference w:type="default"',
            f'<w:headerReference w:type="first" r:id="{rid}"/>'
            '<w:headerReference w:type="default"',
        )
        # 前置节首页(first)也挂带页码的 footer，避免摘要/目录首页缺失页码
        s = s.replace(
            '<w:footerReference w:type="default"',
            f'<w:footerReference w:type="first" r:id="{REL_ID_PAGE_FOOTER}"/>'
            '<w:footerReference w:type="default"',
        )
        if "<w:titlePg/>" not in s:
            s = s.replace(
                "<w:type w:val=\"nextPage\"/>",
                "<w:titlePg/><w:type w:val=\"nextPage\"/>",
            )
        return s

    def _front_title_sect_with_start(rid: str) -> str:
        # 罗马数字从 I 起（保留 start=1）：用于中文摘要这一“罗马数字首页”，
        # 显式重置页码为 I，使摘要页脚显示罗马数字 1。
        s = front_sect
        s = s.replace(
            '<w:headerReference w:type="default"',
            f'<w:headerReference w:type="first" r:id="{rid}"/>'
            '<w:headerReference w:type="default"',
        )
        # 前置节首页(first)也挂带页码的 footer，避免摘要首页缺失页码
        s = s.replace(
            '<w:footerReference w:type="default"',
            f'<w:footerReference w:type="first" r:id="{REL_ID_PAGE_FOOTER}"/>'
            '<w:footerReference w:type="default"',
        )
        if "<w:titlePg/>" not in s:
            s = s.replace(
                "<w:type w:val=\"nextPage\"/>",
                "<w:titlePg/><w:type w:val=\"nextPage\"/>",
            )
        return s

    abstract_cn_sect = _front_title_sect_with_start(ABSTRACT_CN_RID)
    abstract_en_sect = _front_title_sect(ABSTRACT_EN_RID)
    toc_cn_sect = _front_title_sect(TOC_RID)
    toc_en_sect = _front_title_sect(TOC_RID)

    taskbook_text = front_sections.get(front_spec.taskbook_key or "", "").strip()
    external_taskbook_pdf = _resolve_external_taskbook_pdf(taskbook_text, markdown_dir)

    body_rules = thesis_profile.body_parse_rules()
    toc_entries_raw = collect_toc_entries(
        body_text,
        rules=body_rules,
        appendix_heading_normalizer=profile.appendix_heading_normalizer,
    )
    # 摘要/英文摘要也要出现在目录中。它们在前置页渲染，因此需要单独构造书签；
    # 正文标题的书签从 index 3 开始，避免冲突。
    abstract_cn_entry = make_toc_entry(1, level=1, text=front_spec.cn_abstract_title)
    abstract_en_entry = make_toc_entry(2, level=1, text=front_spec.en_abstract_title)
    # 中文目录：编号转为中文（第一章 / 一、 / （一）），与正文标题一致；前置页加入摘要
    toc_entries = [
        abstract_cn_entry,
        abstract_en_entry,
        *[_format_toc_chinese_number(e) for e in toc_entries_raw],
    ]
    # 英文目录（Contents）：复用同一批书签，编号转为英文样式，标题按 JSON 翻译
    heading_translations = _load_heading_translations(markdown_dir)
    toc_entries_en = [
        TocEntry(level=1, text="**Abstract(In Chinese)**", anchor=abstract_cn_entry.anchor, bookmark_id=abstract_cn_entry.bookmark_id),
        TocEntry(level=1, text="**Abstract(In English)**", anchor=abstract_en_entry.anchor, bookmark_id=abstract_en_entry.bookmark_id),
        *[_format_toc_english_number(e, heading_translations) for e in toc_entries_raw],
    ]

    # 后置部分（参考文献之后）：攻读硕士学位期间取得的科研成果 →
    # 原创性声明和使用权限 → 致谢 → 个人简历。顺序与哈工大（硕士）模板目录一致。
    post_ref_titles = [
        HIT_MASTER_ACHIEVEMENTS_TITLE,
        "哈尔滨工业大学学位论文原创性声明和使用权限",
        "致谢",
        "个人简历",
    ]
    # 作者已在正文中提供的后置节（标题精确匹配），用于跳过引擎占位 + 目录去重
    md_provided_titles = {t for t in post_ref_titles if _md_has_heading(body_text, t)}
    post_ref_entries: list[TocEntry] = []
    for _title in post_ref_titles:
        if _title in md_provided_titles:
            # 正文已渲染该节并登记目录，不再添加合成目录条目，避免重复
            continue
        _raw = make_toc_entry(len(toc_entries) + len(post_ref_entries) + 1, level=1, text=_title)
        post_ref_entries.append(_raw)
        toc_entries.append(_format_toc_chinese_number(_raw))
        toc_entries_en.append(_format_toc_english_number(_raw, heading_translations))

    elements: list[str] = []
    taskbook_added = False
    taskbook_ended_with_section = False
    for page in thesis_profile.front_matter_plan().pages:
        if page.kind == "cover":
            _append_cover_page(
                elements,
                thesis_title=thesis_title,
                cover_info=cover_info,
                cover_sect=cover_sect,
                cover_assets_dir=cover_assets_dir,
                media_manager=media_manager,
            )
            continue
        if page.kind == "declaration":
            _append_declaration_page(
                elements,
                page=page,
                declaration=_front_text(front_sections, page),
                math_converter=math_converter,
                reference_anchors=reference_anchors,
                markdown_dir=markdown_dir,
                media_manager=media_manager,
                section_break_after=front_sect if external_taskbook_pdf is not None else None,
            )
            continue
        if page.kind == "taskbook":
            if external_taskbook_pdf is not None:
                taskbook_added, taskbook_ended_with_section = _append_external_taskbook_pdf(
                    elements,
                    pdf_path=external_taskbook_pdf,
                    media_manager=media_manager,
                    section_after=taskbook_image_sect,
                )
            else:
                taskbook_added, taskbook_ended_with_section = _append_taskbook_page(
                    elements,
                    taskbook=_front_text(front_sections, page),
                    cover_info=cover_info,
                )
                if taskbook_added:
                    elements.append(page_break_xml())
            continue
        if page.kind == "abstract":
            _append_abstract_page(
                elements,
                page=page,
                text=_front_text(front_sections, page),
                page_break_before=page.page_break_before,
                math_converter=math_converter,
                reference_anchors=reference_anchors,
                section_break_after=abstract_cn_sect if not page.english else abstract_en_sect,
                toc_entry=abstract_cn_entry if not page.english else abstract_en_entry,
            )
            continue
        if page.kind == "toc":
            _append_toc_page(
                elements,
                thesis_profile=thesis_profile,
                page=page,
                front_sect=toc_cn_sect,
                toc_entries=toc_entries,
            )
            # 英文目录（Contents）：紧接中文目录之后，复用同一批书签，页码一致
            _append_english_toc_page(
                elements,
                thesis_profile=thesis_profile,
                front_sect=toc_en_sect,
                toc_entries=toc_entries_en,
            )
            continue
    chapter_title_headers: list[tuple[str, str]] = []
    body_elements, body_final_sect_pr, _ = build_document_elements(
        body_text,
        profile=profile,
        rules=body_rules,
        treat_first_heading_as_title=False,
        math_converter=math_converter,
        reference_anchors=reference_anchors,
        markdown_dir=markdown_dir,
        media_manager=media_manager,
        chapter_title_headers=chapter_title_headers,
    )
    elements.extend(body_elements)

    # 后置部分（参考文献之后）的标题页眉 rId 映射：科研成果 / 声明 / 致谢 / 简历。
    POST_ACHV_RID = "rId1020"
    POST_DECL_RID = "rId1021"
    POST_ACK_RID = "rId1022"
    POST_RESUME_RID = "rId1023"
    post_title_header_rids = {
        HIT_MASTER_ACHIEVEMENTS_TITLE: POST_ACHV_RID,
        thesis_profile.front_matter_spec().declaration_title: POST_DECL_RID,
        "致谢": POST_ACK_RID,
        "个人简历": POST_RESUME_RID,
    }
    # 后置部分：参考文献之后按哈工大（硕士）模板顺序插入
    # 科研成果 → 原创性声明和使用权限 → 致谢 → 个人简历；
    # 各部分均另起一节，首页显示本节标题页眉；目录条目已在上文登记。
    post_final_sect = _append_post_reference_sections(
        elements,
        thesis_profile=thesis_profile,
        profile=profile,
        thesis_title=thesis_title,
        post_ref_entries=post_ref_entries,
        md_provided_titles=md_provided_titles,
        title_header_rids=post_title_header_rids,
        body_sect_pr=body_final_sect_pr,
    )

    # 每章首页页眉：为每章生成独立的标题页眉部件
    extra_parts: dict[str, str] = {}
    extra_relationships: list[tuple[str, str]] = []
    extra_overrides: list[str] = []
    for rid, full_title in chapter_title_headers:
        part_num = rid.replace("rId", "")
        part_name = f"header{part_num}.xml"
        extra_parts[part_name] = title_header_xml(full_title)
        extra_relationships.append((rid, part_name))
        extra_overrides.append(part_name)

    # 前置页专用首页眉部件：摘 要 / Abstract / 目 录
    extra_parts["header1010.xml"] = title_header_xml(front_spec.cn_abstract_title)
    extra_relationships.append((ABSTRACT_CN_RID, "header1010.xml"))
    extra_overrides.append("header1010.xml")
    extra_parts["header1011.xml"] = title_header_xml(front_spec.en_abstract_title)
    extra_relationships.append((ABSTRACT_EN_RID, "header1011.xml"))
    extra_overrides.append("header1011.xml")
    extra_parts["header1012.xml"] = title_header_xml(front_spec.toc_title)
    extra_relationships.append((TOC_RID, "header1012.xml"))
    extra_overrides.append("header1012.xml")

    # 后置部分专用首页眉部件：科研成果 / 声明 / 致谢 / 个人简历
    extra_parts["header1020.xml"] = title_header_xml(HIT_MASTER_ACHIEVEMENTS_TITLE)
    extra_relationships.append((POST_ACHV_RID, "header1020.xml"))
    extra_overrides.append("header1020.xml")
    extra_parts["header1021.xml"] = title_header_xml(thesis_profile.front_matter_spec().declaration_title)
    extra_relationships.append((POST_DECL_RID, "header1021.xml"))
    extra_overrides.append("header1021.xml")
    extra_parts["header1022.xml"] = title_header_xml("致谢")
    extra_relationships.append((POST_ACK_RID, "header1022.xml"))
    extra_overrides.append("header1022.xml")
    extra_parts["header1023.xml"] = title_header_xml("个人简历")
    extra_relationships.append((POST_RESUME_RID, "header1023.xml"))
    extra_overrides.append("header1023.xml")

    # 文档末尾 sectPr：优先用后置部分返回的最后一节（含其标题页眉）；若没有
    # 后置部分（全部由正文提供），则沿用正文最后一节。
    body_sect = post_final_sect or body_final_sect_pr
    return elements, body_sect, thesis_title, {
        "extra_parts": extra_parts,
        "extra_relationships": extra_relationships,
        "extra_overrides": extra_overrides,
    }
