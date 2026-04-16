#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
扬州大学（YZU）论文格式规格 + 格式化器。
"""

from __future__ import annotations

from pathlib import Path

from docx import Document

from app.services.formatters.base import BaseFormatter
from app.services.formatters.engine import ThesisFormatEngine
from app.services.formatters.report import FormatReport
from app.services.formatters.spec import (
    AbstractRules,
    CaptionRules,
    CoverPageRules,
    FontAssignments,
    FontSizes,
    HeadingNumbering,
    HeadingRule,
    PageLayout,
    ReferenceRules,
    SignatureRules,
    SpacingRules,
    ThesisFormatSpec,
)
from app.services.formatters.schools import register_formatter
from app.schemas.documents import ReviewResponse
from app.services.review_report import build_review_response


def create_yzu_spec(thesis_type: str) -> ThesisFormatSpec:
    """创建扬州大学论文格式规格。"""
    # 按论文类型解析差异项
    if thesis_type == "design_report":
        abstract_font = "宋体"
        h1_alignment = "left"
    else:
        abstract_font = "楷体"
        h1_alignment = "center"

    return ThesisFormatSpec(
        school_id="yzu",
        school_name="扬州大学",
        header_text="扬州大学本科生毕业论文",
        page_layout=PageLayout(
            top_margin=2.2, bottom_margin=2.2, left_margin=2.5,
            right_margin=2.0, gutter=0.5, header_distance=1.2, footer_distance=1.5,
        ),
        fonts=FontAssignments(
            chinese_font="宋体",
            english_font="Times New Roman",
            heading_font="黑体",
            abstract_font=abstract_font,
            signature_font="仿宋体",
        ),
        font_sizes=FontSizes(
            title=18, heading_1=15, heading_2=14, heading_3=12,
            body=12, abstract_label=12, signature=12, header_footer=9, caption=10.5,
        ),
        spacing=SpacingRules(
            title_line_spacing=1.0, heading_line_spacing=1.0,
            body_line_spacing=1.5, paragraph_spacing=12,
        ),
        headings=(
            HeadingRule(alignment="center", font="黑体", size=18, bold=True),   # 0: 论文标题
            HeadingRule(alignment=h1_alignment, font="黑体", size=15, bold=True),  # 1: H1
            HeadingRule(alignment="left", font="黑体", size=14, bold=True),    # 2: H2
            HeadingRule(alignment="left", font="黑体", size=12, bold=True),    # 3: H3
            HeadingRule(alignment="left", font="黑体", size=12, bold=True),    # 4: H4
        ),
        heading_numbering=HeadingNumbering(
            style="chinese",
            level_1=r"^第[一二三四五六七八九十百]+章",
            level_2=r"^[一二三四五六七八九十百]+、",
            level_3=r"^（[一二三四五六七八九十百]+）",
            level_4=r"^\d+\.",
        ),
        abstract=AbstractRules(
            min_chars=100, label_font="黑体", font_size=12,
            keywords=("摘要", "abstract"),
            boundary_keywords=("关键词", "keywords", "目录"),
            first_line_indent=0.74,
        ),
        references=ReferenceRules(
            standard="GB7714-87", min_count=8, max_typical_count=50,
            numbering_id=99, title_font="黑体", title_size=15,
            section_end_keywords=("致谢", "附录", "结论"),
            hanging_indent=0.74,
        ),
        signature=SignatureRules(
            keywords=("年级专业", "学生姓名", "指导教师", "学院", "学号", "届别"),
            max_position=30, font="仿宋体", size=12, alignment="center",
            prefix_patterns=(r"^(学生|教师|专业|班级|学号)[:：]",),
        ),
        caption=CaptionRules(
            pattern=r"^#?\s*(图|表)\s*\d+[-\.]\d+",
            alignment="center", font_size=10.5,
            line_spacing=1.0, space_before=6, space_after=6,
        ),
        cover=CoverPageRules(
            university_name="扬州大学",
            thesis_type_label="本科毕业论文（设计）",
            fields=(
                ("学院", "________"),
                ("专  业", "________"),
                ("班  级", "________"),
                ("学生姓名", "________"),
                ("学  号", "________"),
                ("指导教师", "________"),
                ("完成日期", "________"),
            ),
            university_font="华文行楷",
            university_size=26,
            thesis_type_font="宋体",
            thesis_type_size=26,
            title_font="宋体",
            title_size=22,
            field_label_font="宋体",
            field_label_size=15,
            underline_width=6.0,
        ),
        special_titles=("摘要", "abstract", "参考文献", "references", "目录", "contents"),
        body_first_line_indent=0.74,
    )


class YZUFormatter(BaseFormatter):
    school_id = "yzu"

    def review_document(self, input_path: Path, thesis_type: str) -> ReviewResponse:
        spec = create_yzu_spec(thesis_type)
        engine = ThesisFormatEngine(spec)
        doc = Document(str(input_path))
        report = engine.process_document(doc, filename=input_path.name)
        report_text = engine.generate_report(report)
        return build_review_response(report, report_text, self.school_id, thesis_type)

    def format_document(
        self,
        input_path: Path,
        output_path: Path,
        thesis_type: str,
    ) -> tuple[Path, ReviewResponse]:
        spec = create_yzu_spec(thesis_type)
        engine = ThesisFormatEngine(spec)
        doc = Document(str(input_path))
        report = engine.process_document(doc, filename=input_path.name)
        doc.save(str(output_path))
        report_text = engine.generate_report(report)
        return Path(output_path), build_review_response(report, report_text, self.school_id, thesis_type)


# 注册 YZU
register_formatter("yzu", YZUFormatter)
# sdfmu_ai 使用 YZU 格式（暂无独立规格）
register_formatter("sdfmu_ai", YZUFormatter)
