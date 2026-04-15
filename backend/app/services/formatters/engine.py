#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
论文格式化引擎（规格驱动）。
从 docs/data/yzu_thesis_formatter.py 参考文件移植，所有格式参数从 ThesisFormatSpec 读取。
"""

from __future__ import annotations

import re
from datetime import datetime
from typing import TYPE_CHECKING

from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from app.services.formatters.report import FormatIssueLevel, FormatReport
from app.services.formatters.spec import ThesisFormatSpec

if TYPE_CHECKING:
    from docx import Document

# 对齐方式映射
_ALIGNMENT_MAP = {
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


class ThesisFormatEngine:
    """论文格式化引擎，接收 ThesisFormatSpec 规格作用于 Word 文档。"""

    def __init__(self, spec: ThesisFormatSpec) -> None:
        self.spec = spec

    # ── 主入口 ────────────────────────────────────────────

    def process_document(self, doc: Document, filename: str) -> FormatReport:
        """处理整个文档：应用格式 + 审查校验，返回 FormatReport。"""
        report = FormatReport(filename=filename)
        report.add_issue(FormatIssueLevel.INFO, "文档加载", f"成功加载文档（按{self.spec.school_name}格式处理）", "")

        self._apply_page_setup(doc, report)
        self._apply_header_footer(doc, report)

        paragraphs = doc.paragraphs
        for i, para in enumerate(paragraphs):
            self._format_paragraph(para, para_index=i, all_paragraphs=paragraphs, report=report)

        self._check_abstract(doc, report)
        self._format_references(doc, report)
        self._check_references(doc, report)
        self._check_signature(doc, report)

        return report

    def generate_report(self, report: FormatReport) -> str:
        """生成文本格式审查报告。"""
        s = self.spec
        lines = [
            "=" * 60,
            f"{s.school_name}毕业论文格式审查报告",
            "=" * 60,
            f"文件名: {report.filename}",
            f"学校模板: {s.school_name} ({s.school_id})",
            f"检查时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            "",
            "-" * 60, "统计信息", "-" * 60,
            f"问题总数: {report.total_issues}",
            f"  - 错误: {report.error_count}",
            f"  - 警告: {report.warning_count}",
            f"  - 提示: {report.info_count}",
            "",
        ]

        for level in (FormatIssueLevel.ERROR, FormatIssueLevel.WARNING, FormatIssueLevel.INFO):
            issues = [i for i in report.issues if i.level == level]
            if issues:
                lines += ["-" * 60, f"{level.value} ({len(issues)}项)", "-" * 60]
                for idx, iss in enumerate(issues, 1):
                    lines.append(f"\n[{idx}] 位置: {iss.location}")
                    lines.append(f"    问题: {iss.description}")
                    if iss.current_value and iss.expected_value:
                        lines.append(f"    当前值: {iss.current_value}")
                        lines.append(f"    期望值: {iss.expected_value}")
                    lines.append(f"    建议: {iss.suggestion}")

        lines += ["", "=" * 60, "格式规范参考", "=" * 60]
        lines.append(self._get_format_reference())
        return "\n".join(lines)

    # ── 页面设置 ──────────────────────────────────────────

    def _apply_page_setup(self, doc: Document, report: FormatReport) -> None:
        p = self.spec.page_layout
        for section in doc.sections:
            section.top_margin = Cm(p.top_margin)
            section.bottom_margin = Cm(p.bottom_margin)
            section.left_margin = Cm(p.left_margin)
            section.right_margin = Cm(p.right_margin)
            section.gutter = Cm(p.gutter)
            section.header_distance = Cm(p.header_distance)
            section.footer_distance = Cm(p.footer_distance)
        report.add_issue(FormatIssueLevel.INFO, "页面设置", f"已应用{self.spec.school_name}页面格式", "")

    def _apply_header_footer(self, doc: Document, report: FormatReport) -> None:
        for section in doc.sections:
            # 页眉
            header = section.header
            h_para = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
            h_para.text = self.spec.header_text
            h_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in h_para.runs:
                run.font.name = self.spec.fonts.chinese_font
                run.font.size = Pt(self.spec.font_sizes.header_footer)
                run._element.rPr.rFonts.set(qn("w:eastAsia"), self.spec.fonts.chinese_font)

            # 页脚（页码）
            footer = section.footer
            f_para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
            f_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            self._add_page_number(f_para)

        report.add_issue(FormatIssueLevel.INFO, "页眉页脚", "已添加页眉和页码", "")

    def _add_page_number(self, paragraph) -> None:
        run = paragraph.add_run()
        fld1 = OxmlElement("w:fldChar"); fld1.set(qn("w:fldCharType"), "begin")
        instr = OxmlElement("w:instrText"); instr.set(qn("xml:space"), "preserve"); instr.text = "PAGE"
        fld2 = OxmlElement("w:fldChar"); fld2.set(qn("w:fldCharType"), "end")
        run._r.append(fld1); run._r.append(instr); run._r.append(fld2)
        run.font.name = self.spec.fonts.english_font
        run.font.size = Pt(self.spec.font_sizes.body)

    # ── 段落格式化 ────────────────────────────────────────

    def _format_paragraph(self, para, para_index: int, all_paragraphs: list, report: FormatReport) -> None:
        text = para.text.strip()
        heading_level = self._detect_heading_level(text)

        if self._is_special_title(text):
            para.paragraph_format.line_spacing = self.spec.spacing.body_line_spacing
        elif self._is_signature_paragraph(text, para_index):
            self._format_signature(para)
        elif re.match(self.spec.caption.pattern, text):
            self._format_caption(para)
        elif self.spec.fonts.abstract_font != self.spec.fonts.chinese_font and all_paragraphs and self._is_abstract_content(para_index, all_paragraphs):
            self._format_abstract_content(para)
        elif heading_level >= 0:
            self._format_heading(para, heading_level)
        else:
            self._format_body(para)

    def _format_heading(self, para, level: int) -> None:
        headings = self.spec.headings
        if level < 0 or level >= len(headings):
            return
        rule = headings[level]
        para.alignment = _ALIGNMENT_MAP.get(rule.alignment, WD_ALIGN_PARAGRAPH.LEFT)
        para.paragraph_format.space_before = Pt(self.spec.spacing.paragraph_spacing)
        para.paragraph_format.space_after = Pt(self.spec.spacing.paragraph_spacing)
        para.paragraph_format.line_spacing = self.spec.spacing.heading_line_spacing
        for run in para.runs:
            run.font.name = rule.font
            run.font.size = Pt(rule.size)
            run.font.bold = rule.bold
            run._element.rPr.rFonts.set(qn("w:eastAsia"), rule.font)

    def _format_signature(self, para) -> None:
        sig = self.spec.signature
        para.alignment = _ALIGNMENT_MAP.get(sig.alignment, WD_ALIGN_PARAGRAPH.CENTER)
        para.paragraph_format.line_spacing = self.spec.spacing.body_line_spacing
        for run in para.runs:
            run.font.name = sig.font
            run.font.size = Pt(sig.size)
            run._element.rPr.rFonts.set(qn("w:eastAsia"), sig.font)

    def _format_abstract_content(self, para) -> None:
        para.paragraph_format.line_spacing = self.spec.spacing.body_line_spacing
        para.paragraph_format.first_line_indent = Cm(self.spec.abstract.first_line_indent)
        abstract_font = self.spec.fonts.abstract_font
        for run in para.runs:
            run.font.size = Pt(self.spec.font_sizes.body)
            run.font.name = self.spec.fonts.english_font
            if run.text and any("\u4e00" <= c <= "\u9fff" for c in run.text):
                run._element.rPr.rFonts.set(qn("w:eastAsia"), abstract_font)

    def _format_body(self, para) -> None:
        para.paragraph_format.line_spacing = self.spec.spacing.body_line_spacing
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after = Pt(0)
        para.paragraph_format.first_line_indent = Cm(self.spec.body_first_line_indent)
        for run in para.runs:
            run.font.size = Pt(self.spec.font_sizes.body)
            run.font.name = self.spec.fonts.english_font
            if run.text and any("\u4e00" <= c <= "\u9fff" for c in run.text):
                run._element.rPr.rFonts.set(qn("w:eastAsia"), self.spec.fonts.chinese_font)

    def _format_caption(self, para) -> None:
        c = self.spec.caption
        para.alignment = _ALIGNMENT_MAP.get(c.alignment, WD_ALIGN_PARAGRAPH.CENTER)
        para.paragraph_format.line_spacing = c.line_spacing
        para.paragraph_format.space_before = Pt(c.space_before)
        para.paragraph_format.space_after = Pt(c.space_after)
        for run in para.runs:
            run.font.size = Pt(c.font_size)
            run.font.name = self.spec.fonts.english_font
            run._element.rPr.rFonts.set(qn("w:eastAsia"), self.spec.fonts.chinese_font)

    # ── 检测辅助 ──────────────────────────────────────────

    def _detect_heading_level(self, text: str) -> int:
        text = text.strip()
        n = self.spec.heading_numbering
        if re.match(n.level_1, text): return 1
        if re.match(n.level_2, text): return 2
        if re.match(n.level_3, text): return 3
        if re.match(n.level_4, text): return 4
        # 论文标题启发式检测
        if len(text) < 50 and not any(c in text for c in "。，；：\u201c\u201d\u2018\u2019") and text:
            return 0
        return -1

    def _is_signature_paragraph(self, text: str, para_index: int) -> bool:
        sig = self.spec.signature
        if para_index > sig.max_position:
            return False
        text = text.strip()
        count = sum(1 for k in sig.keywords if k in text)
        if count >= 2:
            return True
        for pat in sig.prefix_patterns:
            if re.match(pat, text):
                return True
        return False

    def _is_abstract_content(self, para_index: int, paragraphs: list) -> bool:
        start_idx = -1
        for i in range(min(50, len(paragraphs))):
            t = paragraphs[i].text.strip().replace(" ", "").lower()
            for kw in self.spec.abstract.keywords:
                if kw in t and (len(t) < 15 or t.startswith(kw)):
                    start_idx = i
                    break
            if start_idx >= 0:
                break

        if start_idx == -1 or para_index <= start_idx:
            return False

        boundary_idx = len(paragraphs)
        for j in range(start_idx + 1, len(paragraphs)):
            t = paragraphs[j].text.strip().replace(" ", "").lower()
            for bk in self.spec.abstract.boundary_keywords:
                if bk in t:
                    boundary_idx = j
                    break
            if boundary_idx < len(paragraphs):
                break
        return start_idx < para_index < boundary_idx

    def _is_special_title(self, text: str) -> bool:
        text_clean = text.strip().replace(" ", "").lower()
        if len(text) < 20:
            for title in self.spec.special_titles:
                if title in text_clean:
                    return True
        return False

    # ── 审查校验 ──────────────────────────────────────────

    def _check_abstract(self, doc: Document, report: FormatReport) -> None:
        found = False
        for para in doc.paragraphs[:20]:
            text = para.text.strip().replace(" ", "")
            for kw in self.spec.abstract.keywords:
                if kw in text and len(text) < 20:
                    found = True
                    for run in para.runs:
                        if kw in run.text:
                            if run.font.name != self.spec.abstract.label_font:
                                report.add_issue(
                                    FormatIssueLevel.WARNING, "中文摘要",
                                    f"'{kw}'标签字体应为{self.spec.abstract.label_font}",
                                    f"修改字体为{self.spec.abstract.label_font}",
                                    run.font.name or "未设置", self.spec.abstract.label_font,
                                )
                    content = text.replace(kw, "").replace("：", "").replace(":", "").strip()
                    if 0 < len(content) < self.spec.abstract.min_chars:
                        report.add_issue(
                            FormatIssueLevel.WARNING, "中文摘要",
                            f"摘要内容可能过短（{len(content)}字），建议约300字",
                            "补充摘要内容", f"{len(content)}字", "约300字",
                        )
                    break
            if found:
                break
        if not found:
            report.add_issue(FormatIssueLevel.WARNING, "文档结构", "未检测到中文摘要", "请检查文档结构")

    def _check_references(self, doc: Document, report: FormatReport) -> None:
        ref_found = False
        ref_section_started = False
        ref_items: list[tuple[int, str]] = []

        for para in doc.paragraphs:
            text = para.text.strip()
            text_ns = text.replace(" ", "")
            if ("参考文献" in text_ns or "References" in text) and len(text) < 20:
                ref_found = True
                ref_section_started = True
                for run in para.runs:
                    if run.font.name != self.spec.fonts.heading_font:
                        report.add_issue(
                            FormatIssueLevel.WARNING, "参考文献标题",
                            f"'参考文献'标题字体应为{self.spec.fonts.heading_font}",
                            f"修改字体为{self.spec.fonts.heading_font}",
                            run.font.name or "未设置", self.spec.fonts.heading_font,
                        )
                continue
            if ref_section_started and text and len(text) > 10:
                ref_items.append((0, text))

        if not ref_found:
            report.add_issue(FormatIssueLevel.WARNING, "文档结构", "未检测到参考文献章节", "请添加参考文献")
            return

        self._analyze_reference_format(ref_items, report)

    def _analyze_reference_format(self, ref_items: list[tuple[int, str]], report: FormatReport) -> None:
        if not ref_items:
            report.add_issue(FormatIssueLevel.WARNING, "参考文献", "参考文献章节无具体条目", "请添加参考文献条目")
            return

        journal = book = conf = other = 0
        issues = []
        for idx, (_, text) in enumerate(ref_items[:20]):
            if re.search(r"[\u4e00-\u9fa5]+.*\d{4}.*\d+.*\d+.*\d+", text) and ("卷" in text or "期" in text):
                journal += 1
                if not re.search(r"\d{4}.*[\(（]\d+[\)）]", text):
                    issues.append(f"第{idx+1}条期刊文献可能缺少期号格式")
            elif re.search(r"(出版社|出版|Press|Publishing)", text, re.IGNORECASE) and re.search(r"\d{4}", text):
                book += 1
            elif re.search(r"(会议|论文集|Proceedings|Conference|Symposium)", text, re.IGNORECASE):
                conf += 1
            else:
                other += 1

        total = len(ref_items)
        report.add_issue(
            FormatIssueLevel.INFO, "参考文献统计",
            f"检测到参考文献共{total}条：期刊约{journal}条，图书约{book}条，会议约{conf}条，其他约{other}条", "",
        )

        for iss in issues[:5]:
            report.add_issue(FormatIssueLevel.WARNING, "参考文献格式", iss, f"请参考{self.spec.references.standard}标准格式")

        if total < self.spec.references.min_count:
            report.add_issue(
                FormatIssueLevel.WARNING, "参考文献数量",
                f"参考文献数量较少（{total}条），本科论文通常要求10-20篇",
                "建议补充相关文献", f"{total}条", "10-20条",
            )
        elif total > self.spec.references.max_typical_count:
            report.add_issue(FormatIssueLevel.INFO, "参考文献数量", f"参考文献数量较多（{total}条），请确保质量", "")

    def _check_signature(self, doc: Document, report: FormatReport) -> None:
        found = False
        for idx, para in enumerate(doc.paragraphs[:10]):
            if self._is_signature_paragraph(para.text, idx):
                found = True
                break
        if not found:
            report.add_issue(FormatIssueLevel.INFO, "文档结构", "未检测到署名（年级专业、学生姓名、指导教师）", "毕业论文需要包含署名")

    # ── 参考文献格式化 ────────────────────────────────────

    def _format_references(self, doc: Document, report: FormatReport) -> None:
        ref_section_started = False
        ref_items: list[tuple[object, str]] = []

        for para in doc.paragraphs:
            text = para.text.strip()
            text_ns = text.replace(" ", "")
            if ("参考文献" in text_ns or "References" in text) and len(text) < 20:
                ref_section_started = True
                para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                for run in para.runs:
                    run.font.name = self.spec.references.title_font
                    run.font.size = Pt(self.spec.references.title_size)
                    run.font.bold = True
                    run._element.rPr.rFonts.set(qn("w:eastAsia"), self.spec.references.title_font)
                continue
            if ref_section_started and text and len(text) > 10:
                clean = re.sub(r"^\[\d+\]\s*", "", text)
                clean = re.sub(r"^\d+[.．]\s*", "", clean)
                ref_items.append((para, clean))
            if ref_section_started and text and len(text) < 10:
                if any(kw in text for kw in self.spec.references.section_end_keywords):
                    break

        if ref_items:
            self._create_reference_numbering(doc)

        for idx, (para, clean_text) in enumerate(ref_items, 1):
            para.clear()
            self._add_reference_text_with_mixed_fonts(para, clean_text)
            para.paragraph_format.line_spacing = self.spec.spacing.body_line_spacing
            para.paragraph_format.first_line_indent = Cm(-self.spec.references.hanging_indent)
            para.paragraph_format.left_indent = Cm(self.spec.references.hanging_indent)
            para.paragraph_format.space_after = Pt(3)
            self._apply_reference_numbering(para, idx)

    def _apply_reference_numbering(self, paragraph, number: int) -> None:
        from docx.oxml import parse_xml
        from docx.oxml.ns import nsdecls
        pPr = paragraph._p.get_or_add_pPr()
        num_id = str(self.spec.references.numbering_id)
        numPr = parse_xml(
            f'<w:numPr {nsdecls("w")}><w:ilvl w:val="0"/><w:numId w:val="{num_id}"/></w:numPr>'
        )
        for child in list(pPr):
            if child.tag.endswith("numPr"):
                pPr.remove(child)
        pPr.append(numPr)
        paragraph.paragraph_format.first_line_indent = None
        paragraph.paragraph_format.left_indent = None

    def _add_reference_text_with_mixed_fonts(self, paragraph, text: str) -> None:
        current_text = ""
        is_current_chinese = None
        fs = self.spec.fonts
        body_size = self.spec.font_sizes.body

        for char in text:
            is_chinese = "\u4e00" <= char <= "\u9fff"
            if is_current_chinese is None:
                is_current_chinese = is_chinese
                current_text = char
            elif is_current_chinese == is_chinese:
                current_text += char
            else:
                if current_text:
                    run = paragraph.add_run(current_text)
                    run.font.size = Pt(body_size)
                    if is_current_chinese:
                        run.font.name = fs.chinese_font
                        run._element.rPr.rFonts.set(qn("w:eastAsia"), fs.chinese_font)
                    else:
                        run.font.name = fs.english_font
                        run._element.rPr.rFonts.set(qn("w:eastAsia"), fs.english_font)
                is_current_chinese = is_chinese
                current_text = char

        if current_text:
            run = paragraph.add_run(current_text)
            run.font.size = Pt(body_size)
            if is_current_chinese:
                run.font.name = fs.chinese_font
                run._element.rPr.rFonts.set(qn("w:eastAsia"), fs.chinese_font)
            else:
                run.font.name = fs.english_font
                run._element.rPr.rFonts.set(qn("w:eastAsia"), fs.english_font)

    def _create_reference_numbering(self, doc: Document) -> None:
        from docx.oxml import parse_xml
        from docx.oxml.ns import nsdecls
        try:
            numbering_part = doc.part.numbering_part
        except Exception:
            from docx.parts.numbering import NumberingPart
            numbering_part = NumberingPart.new()
            doc.part._package.parts.append(numbering_part)

        num_id = str(self.spec.references.numbering_id)
        abstract_num = parse_xml(
            f'<w:abstractNum {nsdecls("w")} w:abstractNumId="{num_id}">'
            '<w:multiLevelType w:val="singleLevel"/>'
            '<w:lvl w:ilvl="0">'
            '<w:numFmt w:val="decimal"/>'
            '<w:lvlText w:val="[%1]"/>'
            '<w:lvlJc w:val="left"/>'
            '<w:pPr><w:ind w:left="0" w:firstLine="0"/></w:pPr>'
            '<w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="Times New Roman"/></w:rPr>'
            '</w:lvl>'
            '</w:abstractNum>'
        )
        num_elem = parse_xml(
            f'<w:num {nsdecls("w")} w:numId="{num_id}">'
            f'<w:abstractNumId w:val="{num_id}"/>'
            '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride>'
            '</w:num>'
        )
        numbering_element = numbering_part._element
        numbering_element.append(abstract_num)
        numbering_element.append(num_elem)

    # ── 报告参考信息 ──────────────────────────────────────

    def _get_format_reference(self) -> str:
        s = self.spec
        p = s.page_layout
        abstract_font = s.fonts.abstract_font
        h1_align = s.headings[1].alignment
        h1_pos = "居中" if h1_align == "center" else "靠左顶格"
        return f"""
【学校模板】{s.school_name}

【页面设置】
- 纸张: A4
- 上边距: {p.top_margin}cm, 下边距: {p.bottom_margin}cm
- 左边距: {p.left_margin}cm, 右边距: {p.right_margin}cm
- 装订线: {p.gutter}cm
- 页眉: {p.header_distance}cm, 页脚: {p.footer_distance}cm

【字体字号】
- 论文题目: {s.headings[0].font} {s.font_sizes.title}pt, 居中
- 一级标题: {s.headings[1].font} {s.font_sizes.heading_1}pt, {h1_pos}
- 二级标题: {s.headings[2].font} {s.font_sizes.heading_2}pt, 靠左顶格
- 三级标题: {s.headings[3].font} {s.font_sizes.heading_3}pt, 靠左顶格
- 正文: {s.fonts.chinese_font} {s.font_sizes.body}pt, {s.spacing.body_line_spacing}倍行距
- 英文/数字: {s.fonts.english_font}
- 摘要内容: {abstract_font} {s.font_sizes.body}pt
- 署名: {s.signature.font} {s.signature.size}pt, 居中

【摘要要求】
- 中文摘要约300字
- "摘要"二字用{s.abstract.label_font} {s.abstract.font_size}pt, 居中
- 内容用{abstract_font} {s.font_sizes.body}pt

【参考文献】
- 格式标准: {s.references.standard}
- 数量要求: 本科论文通常10-20篇
- 期刊格式: 作者.文章题目名[J].期刊名,年份,卷号(期数):页码
- 图书格式: 作者.书名[M].出版地:出版单位,年份:页码
- 会议格式: 作者.文章题目名[C].会议名(论文集),年份:页码
"""
