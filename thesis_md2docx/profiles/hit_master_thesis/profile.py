from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from ...body_rules import BodyParseRules
from ...constants import DEFAULT_COVER_ASSETS_DIR
from ...layout import (
    DocumentLayout,
    FrontMatterPageSpec,
    FrontMatterPlan,
    FrontMatterSpec,
    SectionSpec,
    StyleBundle,
)
from ...math.converter import MathConverter
from ...media import MediaManager
from ...ooxml.parts import native_sect_pr_xml
from ...styles import BodyRenderProfile, StyleCatalog, StyleRoleMap
from ..base import ThesisProfile
from .body import (
    body_parse_rules,
    body_style_profile,
    HIT_PAGE_MARGIN_XML,
    HIT_DOC_GRID_XML,
)
from .document import build_document as build_profile_document
from .frontmatter import resolve_cover_assets_dir
from .styles import hit_style_bundle, hit_style_catalog, hit_style_roles


@dataclass(frozen=True)
class HitMasterThesisProfile(ThesisProfile):
    name: str = "hit-master-thesis"
    display_name: str = "哈尔滨工业大学硕士学位论文"
    default_cover_assets_dir: Path | None = DEFAULT_COVER_ASSETS_DIR

    def resolve_cover_assets_dir(
        self,
        markdown_path: Path,
        assets_dir: Path | None,
        *,
        use_cover_assets: bool,
    ) -> Path | None:
        return resolve_cover_assets_dir(
            markdown_path,
            assets_dir or self.default_cover_assets_dir,
            use_cover_assets=use_cover_assets,
        )

    def body_style_profile(self) -> BodyRenderProfile:
        return body_style_profile()

    def body_parse_rules(self) -> BodyParseRules:
        return body_parse_rules()

    def document_layout(self) -> DocumentLayout:
        return DocumentLayout(
            # 封面无页眉页脚
            cover=SectionSpec(with_header=False, footer_kind="empty", section_type="nextPage"),
            front_matter=SectionSpec(
                with_header=True,
                footer_kind="page",
                section_type="nextPage",
                page_number_format="upperRoman",
                page_number_start=1,
            ),
            body_start=SectionSpec(
                with_header=True,
                footer_kind="page",
                page_number_format="decimal",
                page_number_start=1,
            ),
            body_continue=SectionSpec(with_header=True, footer_kind="page"),
        )

    def front_matter_spec(self) -> FrontMatterSpec:
        return FrontMatterSpec(
            cover_info_key="封面信息",
            declaration_key="声明",
            declaration_title="哈尔滨工业大学学位论文原创性声明和使用权限",
            taskbook_key=None,
            cn_abstract_key="摘要",
            cn_abstract_title="摘  要",
            cn_keyword_prefix="关键词：",
            en_abstract_key="ABSTRACT",
            en_abstract_title="Abstract",
            en_keyword_prefix="Keywords：",
            toc_title="目  录",
            default_title="哈尔滨工业大学硕士学位论文",
        )

    def front_matter_plan(self) -> FrontMatterPlan:
        # 哈工大模板顺序：封面 → 摘要（中/英） → 目录 → 正文 →
        # 原创性声明和使用权限 → 致谢。声明不放在前置部分，
        # 而是在正文之后、致谢之前单独成页（见 document.build_document）。
        spec = self.front_matter_spec()
        pages: list[FrontMatterPageSpec] = [
            FrontMatterPageSpec(kind="cover", source_key=spec.cover_info_key),
        ]
        if spec.cn_abstract_key:
            pages.append(
                FrontMatterPageSpec(
                    kind="abstract",
                    source_key=spec.cn_abstract_key,
                    title=spec.cn_abstract_title,
                    keyword_prefix=spec.cn_keyword_prefix,
                    page_break_after=True,
                )
            )
        if spec.en_abstract_key:
            pages.append(
                FrontMatterPageSpec(
                    kind="abstract",
                    source_key=spec.en_abstract_key,
                    title=spec.en_abstract_title,
                    keyword_prefix=spec.en_keyword_prefix,
                    english=True,
                    page_break_after=True,
                )
            )
        pages.append(FrontMatterPageSpec(kind="toc", title=spec.toc_title))
        return FrontMatterPlan(tuple(pages))

    def section_from_spec(
        self,
        spec: SectionSpec,
        *,
        with_title_header: bool = False,
        title_header_rid: str | None = None,
    ) -> str:
        # 哈工大模板页边距/网格（对照 研究生学位论文写作模板(2020版).doc 实测）：
        #   上 2155(38mm) / 下·左右 1701(30mm) / 页眉 1701 / 页脚 1304(23mm)
        #   网格 type=linesAndChars linePitch=391 charSpace=1861
        # 正文段落显式关闭 snapToGrid（见 styles.py pPrDefault），故行高由段落
        # 自身 spacing(1.25倍) 决定，而非被网格吸成 384，解决"行距看着一样实际更挤"问题。
        return native_sect_pr_xml(
            with_header=spec.with_header,
            footer_kind=spec.footer_kind,
            section_type=spec.section_type,
            page_number_format=spec.page_number_format,
            page_number_start=spec.page_number_start,
            with_title_header=with_title_header,
            title_header_rid=title_header_rid,
            pg_mar_xml=HIT_PAGE_MARGIN_XML,
            doc_grid_xml=HIT_DOC_GRID_XML,
        )

    def style_catalog(self) -> StyleCatalog:
        return hit_style_catalog()

    def style_roles(self) -> StyleRoleMap:
        return hit_style_roles()

    def style_bundle(self) -> StyleBundle:
        return hit_style_bundle()

    def build_document(
        self,
        text: str,
        *,
        math_converter: MathConverter | None = None,
        reference_anchors: dict[str, str] | None = None,
        markdown_dir: Path | None = None,
        cover_assets_dir: Path | None = None,
        media_manager: MediaManager | None = None,
    ) -> tuple[list[str], str, str] | tuple[list[str], str, str, dict]:
        return build_profile_document(
            text,
            thesis_profile=self,
            math_converter=math_converter,
            reference_anchors=reference_anchors,
            markdown_dir=markdown_dir,
            cover_assets_dir=cover_assets_dir,
            media_manager=media_manager,
        )
