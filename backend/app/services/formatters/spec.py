#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
论文格式规格数据类。
纯数据定义，不依赖 python-docx。
"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class PageLayout:
    top_margin: float       # cm
    bottom_margin: float
    left_margin: float
    right_margin: float
    gutter: float
    header_distance: float  # cm
    footer_distance: float


@dataclass(frozen=True)
class FontAssignments:
    chinese_font: str       # 正文中文
    english_font: str       # 正文英文
    heading_font: str       # 标题
    abstract_font: str      # 摘要内容（已按论文类型解析）
    signature_font: str     # 封面签名


@dataclass(frozen=True)
class FontSizes:
    title: float            # pt
    heading_1: float
    heading_2: float
    heading_3: float
    body: float
    abstract_label: float
    signature: float
    header_footer: float
    caption: float


@dataclass(frozen=True)
class SpacingRules:
    title_line_spacing: float       # 倍数
    heading_line_spacing: float
    body_line_spacing: float
    paragraph_spacing: float        # pt 段前段后
    body_line_spacing_rule: str = "multiple"   # "multiple" / "exact"


@dataclass(frozen=True)
class HeadingRule:
    alignment: str          # "center" / "left" / "right"
    font: str
    size: float             # pt
    bold: bool = True


@dataclass(frozen=True)
class HeadingNumbering:
    style: str              # "chinese" | "arabic" — 决定标题编号生成格式
    level_1: str            # 正则，如 r'^第[一二三四五六七八九十百]+章'
    level_2: str
    level_3: str
    level_4: str


@dataclass(frozen=True)
class AbstractRules:
    min_chars: int
    label_font: str
    font_size: float        # pt
    keywords: tuple[str, ...]
    boundary_keywords: tuple[str, ...]
    first_line_indent: float  # cm
    keywords_label: str = "关键词"
    english_keywords_label: str = "Keywords"
    require_english_abstract: bool = False


@dataclass(frozen=True)
class ReferenceRules:
    standard: str           # 如 "GB7714-87"
    min_count: int
    max_typical_count: int
    numbering_id: int       # Word numbering definition ID
    title_font: str
    title_size: float
    section_end_keywords: tuple[str, ...]
    hanging_indent: float   # cm
    min_foreign_count: int = 0
    min_recent_count: int = 0
    recent_year_span: int = 5


@dataclass(frozen=True)
class SignatureRules:
    keywords: tuple[str, ...]
    max_position: int       # 从文档开头计算的段落数
    font: str
    size: float
    alignment: str
    prefix_patterns: tuple[str, ...]


@dataclass(frozen=True)
class CaptionRules:
    pattern: str            # 正则
    alignment: str
    font_size: float
    line_spacing: float
    space_before: float     # pt
    space_after: float


@dataclass(frozen=True)
class CoverPageRules:
    university_name: str            # 学校全称
    thesis_type_label: str          # 如 "本科毕业论文（设计）"
    fields: tuple[tuple[str, str], ...]  # ((标签, 占位文字), ...) 如 (("学院", "______"), ...)
    university_font: str
    university_size: float          # pt
    thesis_type_font: str
    thesis_type_size: float
    title_font: str
    title_size: float
    field_label_font: str
    field_label_size: float
    underline_width: float          # cm 下划线长度
    declaration_title: str = ""
    declaration_body: tuple[str, ...] = ()
    declaration_fields: tuple[str, ...] = ()
    declaration_date_placeholder: str = ""
    logo_path: str = ""
    logo_width: float = 0.0


@dataclass(frozen=True)
class ThesisFormatSpec:
    school_id: str
    school_name: str
    header_text: str
    page_layout: PageLayout
    fonts: FontAssignments
    font_sizes: FontSizes
    spacing: SpacingRules
    headings: tuple[HeadingRule, ...]  # 索引 0=标题, 1=H1, 2=H2, 3=H3, 4=H4
    heading_numbering: HeadingNumbering
    abstract: AbstractRules
    references: ReferenceRules
    signature: SignatureRules
    caption: CaptionRules
    cover: CoverPageRules
    special_titles: tuple[str, ...]
    body_first_line_indent: float  # cm
    special_title_display_map: tuple[tuple[str, str], ...] = ()
    required_sections: tuple[str, ...] = ()
    use_roman_front_matter: bool = True   # 前言部分是否使用罗马数字页码（YZU=True, SDFMU=False）
