#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
山东第一医科大学（SDFMU）医学信息与人工智能学院 — 论文格式规格 + 格式化器。
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
from app.services.formatters.schools import register_formatter, register_spec_factory
from app.schemas.documents import ReviewResponse
from app.services.review_report import build_review_response


def create_sdfmu_spec(thesis_type: str) -> ThesisFormatSpec:
    """创建山东第一医科大学论文格式规格。"""
    return ThesisFormatSpec(
        school_id="sdfmu",
        school_name="山东第一医科大学（山东省医学科学院）",
        header_text="山东第一医科大学（山东省医学科学院）本科毕业论文(设计)",
        page_layout=PageLayout(
            top_margin=2.54, bottom_margin=2.54, left_margin=3.17,
            right_margin=3.17, gutter=0, header_distance=1.5, footer_distance=1.75,
        ),
        fonts=FontAssignments(
            chinese_font="宋体",
            english_font="Times New Roman",
            heading_font="黑体",
            abstract_font="宋体",
            signature_font="宋体",
        ),
        font_sizes=FontSizes(
            title=24, heading_1=15, heading_2=14, heading_3=12,
            body=12, abstract_label=14, signature=12, header_footer=10.5, caption=12,
        ),
        spacing=SpacingRules(
            title_line_spacing=1.5, heading_line_spacing=1.5,
            body_line_spacing=1.5, paragraph_spacing=6,
        ),
        headings=(
            HeadingRule(alignment="center", font="黑体", size=24, bold=True),   # 0: 论文标题（小一）
            HeadingRule(alignment="center", font="黑体", size=15, bold=True),   # 1: H1（小三）
            HeadingRule(alignment="left", font="黑体", size=14, bold=True),     # 2: H2（四号）
            HeadingRule(alignment="left", font="宋体", size=12, bold=False),    # 3: H3（小四）
            HeadingRule(alignment="left", font="宋体", size=12, bold=False),    # 4: H4（小四）
        ),
        heading_numbering=HeadingNumbering(
            style="arabic",
            level_1=r"^第\d+章",
            level_2=r"^\d+\.",
            level_3=r"^\d+\.\d+",
            level_4=r"^\d+\.\d+\.\d+",
        ),
        abstract=AbstractRules(
            min_chars=200, label_font="黑体", font_size=14,
            keywords=("摘要", "abstract"),
            boundary_keywords=("关键词", "keywords", "目录"),
            first_line_indent=0.74,
        ),
        references=ReferenceRules(
            standard="GB7714-87", min_count=15, max_typical_count=50,
            numbering_id=99, title_font="黑体", title_size=16,
            section_end_keywords=("致谢", "附录"),
            hanging_indent=0.74,
        ),
        signature=SignatureRules(
            keywords=("教学机构", "专业", "学生姓名", "指导教师", "学号"),
            max_position=30, font="宋体", size=12, alignment="center",
            prefix_patterns=(r"^(学生|教师|专业|班级|学号)[:：]",),
        ),
        caption=CaptionRules(
            pattern=r"^#?\s*(图|表)\s*\d+[-\.]\d+",
            alignment="center", font_size=12,
            line_spacing=1.0, space_before=6, space_after=6,
        ),
        cover=CoverPageRules(
            university_name="山东第一医科大学（山东省医学科学院）",
            thesis_type_label="毕业论文（设计）",
            fields=(
                ("教学机构", "医学信息与人工智能学院"),
                ("专业", "________"),
                ("年级、班级", "________"),
                ("学号", "________"),
                ("学生姓名", "________"),
                ("指导教师", "________"),
                ("企业导师", "（校企合作填写，普通本科删除）"),
                ("完成日期", "________"),
            ),
            university_font="黑体",
            university_size=24,
            thesis_type_font="黑体",
            thesis_type_size=24,
            title_font="黑体",
            title_size=24,
            field_label_font="宋体",
            field_label_size=16,
            underline_width=6.0,
        ),
        special_titles=("摘要", "abstract", "参考文献", "references", "目录", "contents", "前言", "结论", "致谢", "附录", "文献综述"),
        body_first_line_indent=0.74,
        use_roman_front_matter=False,
    )


class SDFMUFormatter(BaseFormatter):
    school_id = "sdfmu"

    def review_document(self, input_path: Path, thesis_type: str) -> ReviewResponse:
        spec = create_sdfmu_spec(thesis_type)
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
        spec = create_sdfmu_spec(thesis_type)
        engine = ThesisFormatEngine(spec)
        doc = Document(str(input_path))
        report = engine.process_document(doc, filename=input_path.name)
        doc.save(str(output_path))
        report_text = engine.generate_report(report)
        return Path(output_path), build_review_response(report, report_text, self.school_id, thesis_type)


# 注册学校
register_formatter("sdfmu", SDFMUFormatter)
register_formatter("sdfmu_ai", SDFMUFormatter)
register_spec_factory("sdfmu", create_sdfmu_spec)
register_spec_factory("sdfmu_ai", create_sdfmu_spec)
