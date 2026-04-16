#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
将 Markdown 格式的论文初稿转换为 Word 文档（支持数学公式）
使用标准 Word Heading 样式，方便后续在 Word 中直接修改样式
公式使用 OMML 格式，可在 Word 中直接编辑
"""

import os
import re
import sys
import argparse
from pathlib import Path
from urllib.parse import urlparse
from docx import Document
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from app.services.math_utils import latex_to_omml
from app.services.formatters.spec import ThesisFormatSpec
from app.services.formatters.schools.yzu import create_yzu_spec
from app.services.formatters.schools.sdfmu import create_sdfmu_spec


MATH_FONT = 'Times New Roman'

# 对齐方式映射（与 engine.py 保持一致）
_ALIGNMENT_MAP = {
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


def _resolve_spec(school_id: str, thesis_type: str) -> ThesisFormatSpec:
    """根据 school_id 和 thesis_type 加载格式规格。"""
    if school_id == "sdfmu":
        return create_sdfmu_spec(thesis_type)
    return create_yzu_spec(thesis_type)


def _add_block_equation(doc, latex_str, eq_label: str = ""):
    """向 Word 文档中添加一个居中的块级 OMML 公式段落。

    eq_label: 如 "(3-1)"，为空则不编号。
    """
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para.paragraph_format.line_spacing = 1.5
    try:
        omml_elem = latex_to_omml(latex_str, display='block')
        para._element.append(omml_elem)
    except Exception:
        run = para.add_run(latex_str)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)

    # 右对齐的公式编号（使用制表符右对齐）
    if eq_label:
        run_num = para.add_run(f"\t{eq_label}")
        run_num.font.name = 'Times New Roman'
        run_num.font.size = Pt(12)
        # 设置右对齐制表位
        pPr = para._element.get_or_add_pPr()
        tabs = OxmlElement('w:tabs')
        tab = OxmlElement('w:tab')
        tab.set(qn('w:val'), 'right')
        tab.set(qn('w:pos'), '8300')  # 大致右侧位置 (twips)
        tabs.append(tab)
        pPr.append(tabs)


def _add_inline_equation(para, latex_str):
    """向已有段落中插入一个行内 OMML 公式。"""
    try:
        omml_elem = latex_to_omml(latex_str, display='inline')
        para._element.append(omml_elem)
    except Exception:
        run = para.add_run(f'${latex_str}$')
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)


def set_math_font(doc):
    """设置文档的默认公式字体为 Times New Roman。"""
    from lxml import etree
    M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
    # 在 settings.xml 中添加/修改 <m:mathPr><m:mathFont>
    settings = doc.settings.element
    mathPr = settings.find(qn('m:mathPr'))
    if mathPr is None:
        mathPr = OxmlElement('m:mathPr')
        settings.append(mathPr)
    mathFont = mathPr.find(qn('m:mathFont'))
    if mathFont is None:
        mathFont = OxmlElement('m:mathFont')
        mathPr.insert(0, mathFont)
    mathFont.set(qn('m:val'), MATH_FONT)


def add_page_break(paragraph):
    """在段落前添加分页符"""
    run = paragraph.add_run()
    run._r.append(OxmlElement('w:br'))
    run._r[-1].set(qn('w:type'), 'page')


def _insert_toc_field(doc):
    """在当前位置插入 Word 目录域（TOC field）。

    生成一个包含 TOC 域代码的段落，Word 打开时会自动更新目录。
    """
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # 创建域代码结构: <w:fldSimple w:instr=" TOC \\o &quot;1-3&quot; \\h \\z \\u ">
    fldSimple = OxmlElement('w:fldSimple')
    fldSimple.set(qn('w:instr'), ' TOC \\o "1-3" \\h \\z \\u ')

    # 域内占位文字
    fldContent = OxmlElement('w:fldContent')
    r = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), '宋体')
    rFonts.set(qn('w:eastAsia'), '宋体')
    rFonts.set(qn('w:hAnsi'), '宋体')
    rPr.append(rFonts)
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), '24')  # 12pt
    rPr.append(sz)
    r.append(rPr)
    t = OxmlElement('w:t')
    t.text = '（请在 Word 中右键点击此处，选择"更新域"以生成目录）'
    r.append(t)
    fldContent.append(r)
    fldSimple.append(fldContent)
    para._element.append(fldSimple)
    return para


def _generate_cover_page(doc, spec: ThesisFormatSpec):
    """生成论文封面页模板。"""
    from docx.shared import Emu
    cv = spec.cover

    # 空行（顶部留白）
    for _ in range(3):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)

    # 学校名称
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(cv.university_name)
    run.font.name = cv.university_font
    run.font.size = Pt(cv.university_size)
    run._element.rPr.rFonts.set(qn('w:eastAsia'), cv.university_font)

    # 空行
    doc.add_paragraph()

    # 论文类型标签（如"本科毕业论文（设计）"）
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(cv.thesis_type_label)
    run.font.name = cv.thesis_type_font
    run.font.size = Pt(cv.thesis_type_size)
    run._element.rPr.rFonts.set(qn('w:eastAsia'), cv.thesis_type_font)

    # 空行
    for _ in range(2):
        doc.add_paragraph()

    # 论文标题（占位）
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('（论文题目）')
    run.font.name = cv.title_font
    run.font.size = Pt(cv.title_size)
    run.font.bold = True
    run._element.rPr.rFonts.set(qn('w:eastAsia'), cv.title_font)

    # 空行
    for _ in range(3):
        doc.add_paragraph()

    # 字段行（如：学院 ________）
    for label, placeholder in cv.fields:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(6)

        run_label = p.add_run(f'{label}：')
        run_label.font.name = cv.field_label_font
        run_label.font.size = Pt(cv.field_label_size)
        run_label._element.rPr.rFonts.set(qn('w:eastAsia'), cv.field_label_font)

        # 使用下划线占位
        run_underline = p.add_run(placeholder)
        run_underline.font.name = cv.field_label_font
        run_underline.font.size = Pt(cv.field_label_size)
        run_underline.font.underline = True
        run_underline._element.rPr.rFonts.set(qn('w:eastAsia'), cv.field_label_font)


def _add_section_break(doc):
    """添加分节符（下一页），用于罗马/阿拉伯双页码。"""
    from docx.oxml.ns import nsdecls
    from docx.oxml import parse_xml

    # 在最后一段后添加分节符
    sectPr = OxmlElement('w:sectPr')
    sectType = OxmlElement('w:type')
    sectType.set(qn('w:val'), 'nextPage')
    sectPr.append(sectType)

    # 添加到文档最后一个段落的 pPr 中，或直接加到 document body
    body = doc.element.body
    # 在最终 sectPr 之前插入
    final_sectPr = body.find(qn('w:sectPr'))
    if final_sectPr is not None:
        body.insert(list(body).index(final_sectPr), sectPr)
    else:
        body.append(sectPr)
    return sectPr


def _set_page_number_roman(section):
    """设置某节的页码格式为罗马数字（i, ii, iii）。"""
    sectPr = section._sectPr
    pgNumType = sectPr.find(qn('w:pgNumType'))
    if pgNumType is None:
        pgNumType = OxmlElement('w:pgNumType')
        sectPr.append(pgNumType)
    pgNumType.set(qn('w:fmt'), 'lowerLetter')  # Word 用 lowerRoman 不被识别，用自定义
    pgNumType.set(qn('w:start'), '1')


def _set_page_number_arabic(section):
    """设置某节的页码格式为阿拉伯数字，从 1 开始。"""
    sectPr = section._sectPr
    pgNumType = sectPr.find(qn('w:pgNumType'))
    if pgNumType is None:
        pgNumType = OxmlElement('w:pgNumType')
        sectPr.append(pgNumType)
    pgNumType.set(qn('w:fmt'), 'decimal')
    pgNumType.set(qn('w:start'), '1')


def _add_page_number_to_footer(section, spec: ThesisFormatSpec):
    """向页脚添加居中页码。"""
    from docx.shared import Emu
    ft = spec.fonts
    fs = spec.font_sizes

    footer = section.footer
    footer.is_linked_to_previous = False

    # 清空现有内容
    for p in footer.paragraphs:
        p.clear()

    if not footer.paragraphs:
        para = footer.add_paragraph()
    else:
        para = footer.paragraphs[0]

    para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 添加 PAGE 域代码
    run = para.add_run()
    run.font.name = ft.english_font
    run.font.size = Pt(fs.header_footer)

    fldChar_begin = OxmlElement('w:fldChar')
    fldChar_begin.set(qn('w:fldCharType'), 'begin')
    run._r.append(fldChar_begin)

    run2 = para.add_run()
    run2.font.name = ft.english_font
    run2.font.size = Pt(fs.header_footer)
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = ' PAGE '
    run2._r.append(instrText)

    run3 = para.add_run()
    run3.font.name = ft.english_font
    run3.font.size = Pt(fs.header_footer)
    fldChar_end = OxmlElement('w:fldChar')
    fldChar_end.set(qn('w:fldCharType'), 'end')
    run3._r.append(fldChar_end)


def parse_image_meta(alt_text: str):
    """支持 `caption|width=8cm|align=center` 语法。"""
    parts = [part.strip() for part in alt_text.split('|') if part.strip()]
    caption = parts[0] if parts else "（请在此补充图名）"
    width = Cm(12)
    alignment = WD_ALIGN_PARAGRAPH.CENTER

    for option in parts[1:]:
        if '=' not in option:
            continue

        key, value = [item.strip().lower() for item in option.split('=', 1)]
        if key in {'width', 'w'}:
            width_match = re.match(r'^(\d+(?:\.\d+)?)(cm|mm|in|pt)?$', value)
            if not width_match:
                continue

            number = float(width_match.group(1))
            unit = width_match.group(2) or 'cm'
            if unit == 'cm':
                width = Cm(number)
            elif unit == 'mm':
                width = Cm(number / 10)
            elif unit == 'in':
                width = Inches(number)
            elif unit == 'pt':
                width = Pt(number)
        elif key in {'align', 'a'}:
            if value in {'left', '左'}:
                alignment = WD_ALIGN_PARAGRAPH.LEFT
            elif value in {'right', '右'}:
                alignment = WD_ALIGN_PARAGRAPH.RIGHT
            else:
                alignment = WD_ALIGN_PARAGRAPH.CENTER

    return caption, width, alignment


def set_table_borders(table, border_type='three_line'):
    """
    设置表格边框为三线表格式
    三线表：顶线、栏目线、底线
    """
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    
    # 获取表格的 tbl 元素
    tbl = table._tbl
    
    # 创建表格属性
    tblPr = tbl.tblPr if tbl.tblPr is not None else parse_xml(r'<w:tblPr {}></w:tblPr>'.format(nsdecls('w')))
    
    # 清除所有边框
    tblBorders = parse_xml(r'''
        <w:tblBorders {}>
            <w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>
            <w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>
            <w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>
            <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>
            <w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>
            <w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>
        </w:tblBorders>
    '''.format(nsdecls('w')))
    
    # 移除旧的边框设置
    for child in list(tblPr):
        if child.tag.endswith('tblBorders'):
            tblPr.remove(child)
    
    tblPr.append(tblBorders)
    
    # 如果没有 tblPr，添加它
    if tbl.tblPr is None:
        tbl.insert(0, tblPr)
    
    # 设置顶线（第一行顶部）和栏目线（第一行底部）
    if len(table.rows) > 0:
        header_row = table.rows[0]
        for cell in header_row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tcBorders = parse_xml(r'''
                <w:tcBorders {}>
                    <w:top w:val="single" w:sz="12" w:space="0" w:color="000000"/>
                    <w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>
                    <w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>
                    <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>
                </w:tcBorders>
            '''.format(nsdecls('w')))
            # 移除旧的单元格边框设置
            for child in list(tcPr):
                if child.tag.endswith('tcBorders'):
                    tcPr.remove(child)
            tcPr.append(tcBorders)
    
    # 设置底线（最后一行底部）
    if len(table.rows) > 1:
        bottom_row = table.rows[-1]
        for cell in bottom_row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tcBorders = parse_xml(r'''
                <w:tcBorders {}>
                    <w:bottom w:val="single" w:sz="12" w:space="0" w:color="000000"/>
                    <w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>
                    <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>
                </w:tcBorders>
            '''.format(nsdecls('w')))
            # 移除旧的单元格边框设置
            for child in list(tcPr):
                if child.tag.endswith('tcBorders'):
                    tcPr.remove(child)
            tcPr.append(tcBorders)
    
    # 中间行无边框
    for row_idx in range(1, len(table.rows) - 1):
        row = table.rows[row_idx]
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tcBorders = parse_xml(r'''
                <w:tcBorders {}>
                    <w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>
                    <w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>
                    <w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>
                    <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>
                </w:tcBorders>
            '''.format(nsdecls('w')))
            # 移除旧的单元格边框设置
            for child in list(tcPr):
                if child.tag.endswith('tcBorders'):
                    tcPr.remove(child)
            tcPr.append(tcBorders)


def format_reference_text(para, text, spec: ThesisFormatSpec):
    """
    格式化参考文献文本，确保数字和英文使用指定英文字体
    中文字符使用正文字体
    """
    import re
    ft = spec.fonts
    fs = spec.font_sizes

    # 匹配模式：序号、作者、标题、期刊、年份、卷期、页码等
    # 首先处理序号 [数字]
    parts = re.split(r'(\[\d+\])', text)

    for part in parts:
        if not part:
            continue

        # 如果是序号 [数字]
        if re.match(r'^\[\d+\]$', part):
            run = para.add_run(part)
            run.font.name = ft.english_font
            run.font.size = Pt(fs.body)
            run._element.rPr.rFonts.set(qn('w:eastAsia'), ft.english_font)
        else:
            # 处理剩余部分，区分中英文
            current_run_text = ""
            is_current_ascii = None

            for char in part:
                is_ascii = ord(char) < 128 and char.isprintable() or char.isspace()

                if is_current_ascii is None:
                    is_current_ascii = is_ascii
                    current_run_text = char
                elif is_current_ascii == is_ascii:
                    current_run_text += char
                else:
                    if current_run_text:
                        run = para.add_run(current_run_text)
                        run.font.size = Pt(fs.body)
                        if is_current_ascii:
                            run.font.name = ft.english_font
                            run._element.rPr.rFonts.set(qn('w:eastAsia'), ft.english_font)
                        else:
                            run.font.name = ft.chinese_font
                            run._element.rPr.rFonts.set(qn('w:eastAsia'), ft.chinese_font)
                    is_current_ascii = is_ascii
                    current_run_text = char

            if current_run_text:
                run = para.add_run(current_run_text)
                run.font.size = Pt(fs.body)
                if is_current_ascii:
                    run.font.name = ft.english_font
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), ft.english_font)
                else:
                    run.font.name = ft.chinese_font
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), ft.chinese_font)


def setup_styles(doc, spec: ThesisFormatSpec):
    """设置文档的标准样式（从 spec 读取参数）"""
    ft = spec.fonts
    fs = spec.font_sizes
    sp = spec.spacing

    # 设置正文样式
    style = doc.styles['Normal']
    style.font.name = ft.english_font
    style.font.size = Pt(fs.body)
    style._element.rPr.rFonts.set(qn('w:eastAsia'), ft.chinese_font)

    # 设置 Heading 1 样式（一级标题）
    heading1 = doc.styles['Heading 1']
    h1 = spec.headings[1]
    heading1.font.name = ft.english_font
    heading1.font.size = Pt(h1.size)
    heading1.font.bold = h1.bold
    heading1.font.color.rgb = RGBColor(0, 0, 0)
    heading1._element.rPr.rFonts.set(qn('w:eastAsia'), h1.font)
    heading1.paragraph_format.alignment = _ALIGNMENT_MAP.get(h1.alignment, WD_ALIGN_PARAGRAPH.CENTER)
    heading1.paragraph_format.space_before = Pt(sp.paragraph_spacing)
    heading1.paragraph_format.space_after = Pt(sp.paragraph_spacing)
    heading1.paragraph_format.line_spacing = sp.heading_line_spacing

    # 设置 Heading 2 样式（二级标题）
    heading2 = doc.styles['Heading 2']
    h2 = spec.headings[2]
    heading2.font.name = ft.english_font
    heading2.font.size = Pt(h2.size)
    heading2.font.bold = h2.bold
    heading2.font.color.rgb = RGBColor(0, 0, 0)
    heading2._element.rPr.rFonts.set(qn('w:eastAsia'), h2.font)
    heading2.paragraph_format.alignment = _ALIGNMENT_MAP.get(h2.alignment, WD_ALIGN_PARAGRAPH.LEFT)
    heading2.paragraph_format.space_before = Pt(sp.paragraph_spacing)
    heading2.paragraph_format.space_after = Pt(sp.paragraph_spacing)
    heading2.paragraph_format.line_spacing = sp.heading_line_spacing

    # 设置 Heading 3 样式（三级标题）
    heading3 = doc.styles['Heading 3']
    h3 = spec.headings[3]
    heading3.font.name = ft.english_font
    heading3.font.size = Pt(h3.size)
    heading3.font.bold = h3.bold
    heading3.font.color.rgb = RGBColor(0, 0, 0)
    heading3._element.rPr.rFonts.set(qn('w:eastAsia'), h3.font)
    heading3.paragraph_format.alignment = _ALIGNMENT_MAP.get(h3.alignment, WD_ALIGN_PARAGRAPH.LEFT)
    heading3.paragraph_format.space_before = Pt(sp.paragraph_spacing)
    heading3.paragraph_format.line_spacing = sp.heading_line_spacing

    # 设置论文标题样式（自定义样式）
    h0 = spec.headings[0]
    if '论文标题' not in doc.styles:
        title_style = doc.styles.add_style('论文标题', WD_STYLE_TYPE.PARAGRAPH)
    else:
        title_style = doc.styles['论文标题']
    title_style.font.name = ft.english_font
    title_style.font.size = Pt(h0.size)
    title_style.font.bold = h0.bold
    title_style.font.color.rgb = RGBColor(0, 0, 0)
    title_style._element.rPr.rFonts.set(qn('w:eastAsia'), h0.font)
    title_style.paragraph_format.alignment = _ALIGNMENT_MAP.get(h0.alignment, WD_ALIGN_PARAGRAPH.CENTER)
    title_style.paragraph_format.space_before = Pt(sp.paragraph_spacing)
    title_style.paragraph_format.space_after = Pt(sp.paragraph_spacing)
    title_style.paragraph_format.line_spacing = sp.title_line_spacing

    # 设置 Caption 样式（题注）
    cap = spec.caption
    if 'Caption' not in doc.styles:
        caption_style = doc.styles.add_style('Caption', WD_STYLE_TYPE.PARAGRAPH)
    else:
        caption_style = doc.styles['Caption']
    caption_style.font.name = ft.english_font
    caption_style.font.size = Pt(cap.font_size)
    caption_style.font.color.rgb = RGBColor(0, 0, 0)
    caption_style._element.rPr.rFonts.set(qn('w:eastAsia'), ft.chinese_font)
    caption_style.paragraph_format.alignment = _ALIGNMENT_MAP.get(cap.alignment, WD_ALIGN_PARAGRAPH.CENTER)
    caption_style.paragraph_format.line_spacing = cap.line_spacing
    caption_style.paragraph_format.space_before = Pt(6)
    caption_style.paragraph_format.space_after = Pt(6)


def int_to_chinese(num):
    chinese_digits = ["", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十"]
    if num == 0: return "零"
    if num <= 10: return chinese_digits[num]
    elif num < 20: return "十" + chinese_digits[num - 10]
    elif num % 10 == 0: return chinese_digits[num // 10] + "十"
    else: return chinese_digits[num // 10] + "十" + chinese_digits[num % 10]

def clean_heading(text):
    text = re.sub(r'^#{1,4}\s*', '', text)
    text = re.sub(r'^[第一二三四五六七八九十百]+章\s*', '', text)
    text = re.sub(r'^第[一二三四五六七八九十百]+节\s*', '', text)
    text = re.sub(r'^[一二三四五六七八九十]+\s*[、\.]?\s*', '', text)
    text = re.sub(r'^（[一二三四五六七八九十]+）\s*', '', text)
    text = re.sub(r'^[\d.]+、?\s*', '', text)
    text = re.sub(r'^（\d+）\s*', '', text)
    return text.strip()

def convert_markdown_to_word(markdown_path: str, output_path: str,
                             school_id: str = "yzu", thesis_type: str = "thesis"):
    """将 Markdown 转换为 Word 文档，使用标准 Heading 样式"""

    # 加载学校格式规格
    spec = _resolve_spec(school_id, thesis_type)

    # 读取 Markdown 文件
    with open(markdown_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # 创建 Word 文档
    doc = Document()

    # 设置页面（从 spec 读取）
    pg = spec.page_layout
    for section in doc.sections:
        section.top_margin = Cm(pg.top_margin)
        section.bottom_margin = Cm(pg.bottom_margin)
        section.left_margin = Cm(pg.left_margin)
        section.right_margin = Cm(pg.right_margin)
        section.gutter = Cm(pg.gutter)

    # 设置标准样式
    setup_styles(doc, spec)
    set_math_font(doc)

    # 解析 Markdown 内容
    lines = content.split('\n')
    in_code_block = False
    in_math_block = False
    math_buffer = []
    code_buffer = []
    in_abstract = False
    in_references = False
    in_front_matter = True   # 前言部分（封面、摘要、目录）用罗马数字页码
    first_heading1 = True  # 标记是否是第一个一级标题（论文题目）
    chap_num = 0
    sec_num = 0
    subsec_num = 0
    subsubsec_num = 0
    fig_num = 0
    tab_num = 0
    eq_num = 0  # 公式编号（按章节重置）
    
    i = 0
    while i < len(lines):
        line = lines[i].rstrip()
        
        # 跳过空行
        if not line.strip():
            i += 1
            continue
        
        stripped = line.strip()
        
        # 代码块处理
        if stripped.startswith('```'):
            if not in_code_block:
                # 代码块开始
                in_code_block = True
                code_buffer = []
            else:
                # 代码块结束，输出已收集的代码内容
                in_code_block = False
                if code_buffer:
                    para = doc.add_paragraph()
                    para.paragraph_format.left_indent = Cm(1)
                    run = para.add_run('\n'.join(code_buffer))
                    run.font.name = 'Courier New'
                    run.font.size = Pt(10)
                    code_buffer = []
            i += 1
            continue

        if in_code_block:
            code_buffer.append(line)
            i += 1
            continue

        # ── 数学公式块处理 ($$...$$) ──
        # 单行公式：$$ E = mc^2 $$
        if stripped.startswith('$$') and stripped.endswith('$$') and len(stripped) > 4 and not in_math_block:
            latex = stripped[2:-2].strip()
            eq_num += 1
            eq_label = f"({chap_num}-{eq_num})" if chap_num > 0 else f"({eq_num})"
            _add_block_equation(doc, latex, eq_label)
            i += 1
            continue

        # 多行公式块开始/结束
        if stripped == '$$':
            if not in_math_block:
                in_math_block = True
                math_buffer = []
            else:
                in_math_block = False
                if math_buffer:
                    latex = '\n'.join(math_buffer)
                    eq_num += 1
                    eq_label = f"({chap_num}-{eq_num})" if chap_num > 0 else f"({eq_num})"
                    _add_block_equation(doc, latex, eq_label)
                    math_buffer = []
            i += 1
            continue

        if in_math_block:
            math_buffer.append(stripped)
            i += 1
            continue

        # -- 图片插图题注 --
        if stripped.startswith('!['):
            img_match = re.match(r'^!\[(.*?)\]\((.*?)\)$', stripped)
            if img_match:
                caption, image_width, image_alignment = parse_image_meta(img_match.group(1).strip())
                image_target = img_match.group(2).strip()
                fig_num += 1
                caption_text = f"图 {chap_num}-{fig_num} {caption}" if chap_num > 0 else f"图 {fig_num} {caption}"
                
                para_img = doc.add_paragraph()
                para_img.alignment = image_alignment
                run_img = para_img.add_run()

                image_path = Path(image_target)
                if image_path.exists():
                    try:
                        run_img.add_picture(str(image_path), width=image_width)
                    except Exception:
                        # python-docx 可能不支持某些 JPEG，用 PIL 转为 PNG 后插入
                        try:
                            from PIL import Image
                            import io
                            img = Image.open(str(image_path))
                            buf = io.BytesIO()
                            img.save(buf, format='PNG')
                            buf.seek(0)
                            run_img.add_picture(buf, width=image_width)
                        except Exception:
                            run_img = para_img.add_run(f"[插图格式不支持: {image_target}]")
                            run_img.font.color.rgb = RGBColor(128, 128, 128)
                else:
                    run_img = para_img.add_run(f"[插图: {image_target}]")
                    run_img.font.color.rgb = RGBColor(128, 128, 128)
                
                # 图题注 (图片下方)
                para = doc.add_paragraph()
                para.alignment = image_alignment
                run = para.add_run(caption_text)
                run.font.name = spec.fonts.chinese_font
                run.font.size = Pt(spec.font_sizes.caption)
                run._element.rPr.rFonts.set(qn('w:eastAsia'), spec.fonts.chinese_font)
                
                just_parsed_table_caption = True
                i += 1
                continue

        # -- 图/表 显示题注 --
        cap_match = re.match(r'^#*\s*(\*\*?)?(图|表)\s*(?:\d+[-\.]\d+|\d+)?[:：\s]+\s*(.+?)(\*\*?)?$', stripped)
        if cap_match and '格式' not in stripped:
            ctype = cap_match.group(2)
            caption = cap_match.group(3).strip()
            if ctype == '图':
                fig_num += 1
                caption_text = f"图 {chap_num}-{fig_num} {caption}" if chap_num > 0 else f"图 {fig_num} {caption}"
            else:
                tab_num += 1
                caption_text = f"表 {chap_num}-{tab_num} {caption}" if chap_num > 0 else f"表 {tab_num} {caption}"
            
            para = doc.add_paragraph(caption_text, style='Caption')
            just_parsed_table_caption = (ctype == '表')
            i += 1
            continue

        # -- 各类结构标题 --
        if stripped.startswith('#'):
            h_match = re.match(r'^(#{1,4})\s+(.*)', stripped)
            if h_match:
                level_marks = h_match.group(1)
                raw_title = h_match.group(2).strip()
                title_clean = clean_heading(raw_title)

                # ── 封面页 ──
                if '封面' in title_clean or 'Cover' in title_clean.lower():
                    _generate_cover_page(doc, spec)
                    first_heading1 = False
                    i += 1
                    continue

                # 特殊区块标题
                if '参考文献' in title_clean or 'References' in title_clean:
                    in_abstract = False
                    in_references = True
                    para = doc.add_paragraph('参考文献', style='Heading 1')
                    if not first_heading1: para.paragraph_format.page_break_before = True
                    first_heading1 = False
                    i += 1
                    continue
                
                if '摘要' in title_clean or 'Abstract' in title_clean.lower() and len(title_clean) < 15:
                    if '摘要' in title_clean: in_abstract = True
                    para = doc.add_paragraph()
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run = para.add_run(title_clean)
                    run.font.name = spec.abstract.label_font
                    run.font.size = Pt(spec.abstract.font_size)
                    run.font.bold = True
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), spec.abstract.label_font)
                    first_heading1 = False
                    i += 1
                    continue

                if '目录' in title_clean or 'Contents' in title_clean.lower() and len(title_clean) < 15:
                    in_abstract = False
                    para = doc.add_paragraph()
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run = para.add_run('目录')
                    run.font.name = spec.fonts.heading_font
                    run.font.size = Pt(spec.font_sizes.heading_1)
                    run.font.bold = True
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), spec.fonts.heading_font)
                    # 插入真正的 TOC 域
                    _insert_toc_field(doc)
                    first_heading1 = False
                    i += 1
                    continue
                
                # 论文题目（第一个一级标题，且不是"摘要"、"目录"等特殊标题）
                if first_heading1 and title_clean and len(level_marks) <= 2:
                    # 检查是否是特殊标题（摘要、目录等），这些不应该作为论文题目
                    special_titles = ['摘要', 'abstract', '目录', 'contents', '参考文献', 'references']
                    is_special = any(s in title_clean.lower() for s in special_titles)
                    
                    if not is_special:
                        in_abstract = False
                        # 使用论文标题样式（黑体小二）
                        para = doc.add_paragraph(title_clean, style='论文标题')
                        first_heading1 = False
                        i += 1
                        continue
                
                # 正规编号的标题
                in_abstract = False
                level = len(level_marks)
                # 兼容使用者错误的一阶标注习惯
                if level == 1 or (level == 2 and raw_title.startswith('第一章')):
                    # ── 前言/正文分节（罗马→阿拉伯页码）──
                    if in_front_matter:
                        _add_section_break(doc)
                        in_front_matter = False

                    chap_num += 1
                    sec_num = 0
                    subsec_num = 0
                    subsubsec_num = 0
                    fig_num = 0
                    tab_num = 0
                    eq_num = 0
                    num_style = spec.heading_numbering.style
                    if num_style == "arabic":
                        heading_text = f"第{chap_num}章 {title_clean}"
                    else:
                        heading_text = f"第一章 {title_clean}" if chap_num == 1 else f"第{int_to_chinese(chap_num)}章 {title_clean}"
                    para = doc.add_paragraph(heading_text, style='Heading 1')
                    para.paragraph_format.page_break_before = True
                elif level == 2:
                    sec_num += 1
                    subsec_num = 0
                    subsubsec_num = 0
                    if num_style == "arabic":
                        heading_text = f"{sec_num}. {title_clean}"
                    else:
                        heading_text = f"{int_to_chinese(sec_num)}、{title_clean}"
                    doc.add_paragraph(heading_text, style='Heading 2')
                elif level == 3:
                    subsec_num += 1
                    subsubsec_num = 0
                    if num_style == "arabic":
                        heading_text = f"{sec_num}.{subsec_num} {title_clean}"
                    else:
                        heading_text = f"（{int_to_chinese(subsec_num)}）{title_clean}"
                    doc.add_paragraph(heading_text, style='Heading 3')
                elif level >= 4:
                    subsubsec_num += 1
                    if num_style == "arabic":
                        heading_text = f"{sec_num}.{subsec_num}.{subsubsec_num} {title_clean}"
                    else:
                        heading_text = f"{subsubsec_num}. {title_clean}"
                    para = doc.add_paragraph(heading_text, style='Heading 3')
                    para.runs[0].font.size = Pt(spec.headings[4].size)
                
                first_heading1 = False
                just_parsed_table_caption = False
                i += 1
                continue
        
        # 参考文献条目（支持 [1]、[ 1 ] 等格式）
        if re.match(r'^\s*\[\s*\d+\s*\]', stripped):
            ref_text = stripped.strip()
            para = doc.add_paragraph()
            hang = spec.references.hanging_indent
            para.paragraph_format.left_indent = Cm(hang)
            para.paragraph_format.first_line_indent = Cm(-hang)
            # 使用新的格式化函数，确保数字和英文使用指定英文字体
            format_reference_text(para, ref_text, spec)
            i += 1
            continue
        
        # 表格处理（使用三线表格式）
        if stripped.startswith('|'):
            # 若之前没有出现专门的表题标注，则自动补充一个缺失的名字
            if i == 0 or not just_parsed_table_caption:
                tab_num += 1
                caption_text = f"表 {chap_num}-{tab_num} （请补充表名）" if chap_num > 0 else f"表 {tab_num} （请补充表名）"
                doc.add_paragraph(caption_text, style='Caption')
            
            table_lines = []
            while i < len(lines) and lines[i].strip().startswith('|'):
                table_lines.append(lines[i].strip())
                i += 1
            if len(table_lines) > 2:
                header_cells = [c.strip() for c in table_lines[0].split('|') if c.strip()]
                num_cols = len(header_cells)
                table = doc.add_table(rows=1, cols=num_cols)
                # 使用 Light Grid 样式作为基础，然后自定义边框
                table.style = 'Light Grid Accent 1'
                # 表格整体居中
                tblPr = table._tbl.tblPr
                if tblPr is None:
                    tblPr = OxmlElement('w:tblPr')
                    table._tbl.insert(0, tblPr)
                jc = OxmlElement('w:jc')
                jc.set(qn('w:val'), 'center')
                tblPr.append(jc)
                
                # 填充表头
                for col_idx, cell_text in enumerate(header_cells):
                    cell = table.cell(0, col_idx)
                    cell.text = cell_text
                    # 设置表头字体和居中
                    for paragraph in cell.paragraphs:
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        for run in paragraph.runs:
                            run.font.name = spec.fonts.english_font
                            run.font.size = Pt(spec.font_sizes.caption)
                            run._element.rPr.rFonts.set(qn('w:eastAsia'), spec.fonts.chinese_font)
                            run.font.bold = True

                # 填充数据行
                for row_line in table_lines[2:]:
                    row_cells = [c.strip() for c in row_line.split('|')]
                    if row_line.startswith('|') and len(row_cells) > 0: row_cells.pop(0)
                    if row_line.endswith('|') and len(row_cells) > 0: row_cells.pop()
                    row = table.add_row()
                    for col_idx in range(min(num_cols, len(row_cells))):
                        cell = row.cells[col_idx]
                        cell.text = row_cells[col_idx]
                        # 设置单元格字体和居中
                        for paragraph in cell.paragraphs:
                            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            for run in paragraph.runs:
                                run.font.name = spec.fonts.english_font
                                run.font.size = Pt(spec.font_sizes.caption)
                                run._element.rPr.rFonts.set(qn('w:eastAsia'), spec.fonts.chinese_font)
                
                # 应用三线表格式
                set_table_borders(table)
                        
            just_parsed_table_caption = False
            continue
        
        # 分隔线
        if stripped.startswith('---'):
            i += 1
            continue
        
        # 引用块
        if stripped.startswith('>'):
            quote_text = stripped[1:].strip()
            para = doc.add_paragraph()
            para.paragraph_format.left_indent = Cm(1)
            run = para.add_run(quote_text)
            run.font.name = spec.fonts.abstract_font
            run.font.size = Pt(spec.font_sizes.body)
            run.font.italic = True
            run._element.rPr.rFonts.set(qn('w:eastAsia'), spec.fonts.abstract_font)
            i += 1
            continue
        
        # 列表项与无标签四级标题内联文本 (1. xxx: ...)
        if re.match(r'^([\*\-\+]|\d+\.)\s', stripped):
            list_text = re.sub(r'^([\*\-\+]|\d+\.)\s', '', stripped)
            prefix = re.match(r'^([\*\-\+]|\d+\.)\s', stripped).group(1)
            para = doc.add_paragraph()
            para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            para.paragraph_format.left_indent = Pt(24)
            
            # 支持 “1. 标题：内容” 的行内黑体样式
            inline_title_match = re.match(r'^([^：:]+[：:])(.*)$', list_text)
            if prefix.endswith('.'):
                text_prefix = f"{prefix} "
            else:
                text_prefix = "• "
                
            if inline_title_match:
                run_title = para.add_run(text_prefix + inline_title_match.group(1))
                run_title.font.name = spec.fonts.heading_font
                run_title.font.size = Pt(spec.font_sizes.body)
                run_title.font.bold = True
                run_title._element.rPr.rFonts.set(qn('w:eastAsia'), spec.fonts.heading_font)

                run_content = para.add_run(inline_title_match.group(2))
                run_content.font.name = spec.fonts.chinese_font
                run_content.font.size = Pt(spec.font_sizes.body)
                run_content._element.rPr.rFonts.set(qn('w:eastAsia'), spec.fonts.chinese_font)
            else:
                run = para.add_run(text_prefix + list_text)
                run.font.name = spec.fonts.chinese_font
                run.font.size = Pt(spec.font_sizes.body)
                run._element.rPr.rFonts.set(qn('w:eastAsia'), spec.fonts.chinese_font)
                
            just_parsed_table_caption = False
            i += 1
            continue
        
        # 正文段落
        if stripped and not stripped.startswith('#') and not stripped.startswith('['):
            # 特殊检测：加粗版摘要或关键词（针对截图1的情况）
            if re.match(r'^\s*(\*\*|__)\s*(摘要|关键词)\s*(\*\*|__)\s*[:：]', stripped):
                is_keywords_line = '关键词' in stripped
                if not is_keywords_line:
                    in_abstract = True
                para = doc.add_paragraph()
                if not is_keywords_line: para.alignment = WD_ALIGN_PARAGRAPH.CENTER

                parts = re.split(r'([：:])', stripped, 1) # 分割前缀
                # 前缀（摘要/关键词）
                run_label = para.add_run(parts[0].replace('*','').replace('_',''))
                run_label.font.name = spec.abstract.label_font
                run_label.font.size = Pt(spec.abstract.font_size)
                run_label.font.bold = True
                run_label._element.rPr.rFonts.set(qn('w:eastAsia'), spec.abstract.label_font)

                # 冒号及之后内容
                if len(parts) > 1:
                    run_rest = para.add_run(''.join(parts[1:]))
                    run_rest.font.size = Pt(spec.abstract.font_size)
                    run_rest.font.name = spec.fonts.english_font
                    if not is_keywords_line and in_abstract:
                        run_rest._element.rPr.rFonts.set(qn('w:eastAsia'), spec.fonts.abstract_font)
                    else:
                        run_rest._element.rPr.rFonts.set(qn('w:eastAsia'), spec.fonts.chinese_font if is_keywords_line else spec.fonts.abstract_font)
                
                just_parsed_table_caption = False
                i += 1
                continue

            para = doc.add_paragraph()
            para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            para.paragraph_format.line_spacing = spec.spacing.body_line_spacing
            para.paragraph_format.first_line_indent = Pt(spec.font_sizes.body * 2)

            # 判断是否为摘要内容
            is_abstract_content = False
            if in_abstract and '关键词' not in stripped:
                is_abstract_content = True

            # 处理正文中的行内公式和引用上标
            parts = re.split(r'(\$[^$]+\$|\[\d+(?:,\s*\d+)*\])', stripped)
            for part in parts:
                if not part:
                    continue
                # 行内公式 $...$
                if part.startswith('$') and part.endswith('$') and len(part) > 2:
                    _add_inline_equation(para, part[1:-1])
                # 引用上标 [n]
                elif re.match(r'^\[\d+(?:,\s*\d+)*\]$', part):
                    run = para.add_run(part)
                    run.font.size = Pt(spec.font_sizes.body)
                    run.font.name = spec.fonts.english_font
                    run.font.superscript = True
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), spec.fonts.abstract_font if is_abstract_content else spec.fonts.chinese_font)
                else:
                    run = para.add_run(part)
                    run.font.size = Pt(spec.font_sizes.body)
                    run.font.name = spec.fonts.english_font
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), spec.fonts.abstract_font if is_abstract_content else spec.fonts.chinese_font)
            just_parsed_table_caption = False
            i += 1
            continue
        
        i += 1

    # ── 后处理：设置页码（前言罗马，正文阿拉伯）──
    sections = doc.sections
    if len(sections) >= 2:
        # 第一节（前言）：罗马数字页码
        _set_page_number_roman(sections[0])
        _add_page_number_to_footer(sections[0], spec)
        # 第二节（正文）：阿拉伯数字页码，从 1 开始
        _set_page_number_arabic(sections[1])
        _add_page_number_to_footer(sections[1], spec)
    elif len(sections) == 1:
        # 只有一个节，直接设置阿拉伯页码
        _set_page_number_arabic(sections[0])
        _add_page_number_to_footer(sections[0], spec)

    # 保存文档
    doc.save(output_path)
    print("[OK] 转换完成！")
    print(f"输出文件：{output_path}")
    print("\n样式说明：")
    print("   - 一级标题：应用 'Heading 1' 样式（黑体小三号，居中），自动分页")
    print("   - 二级标题：应用 'Heading 2' 样式（黑体四号，靠左）")
    print("   - 三级标题：应用 'Heading 3' 样式（黑体小四号，靠左）")
    print("   - 正文：宋体小四号，1.5倍行距，首行缩进2字符")
    print("   - 参考文献：悬挂缩进格式")
    print("\n提示：在 Word 中可以通过'样式'面板直接修改各级标题样式")


def main():
    parser = argparse.ArgumentParser(
        description='将 Markdown 论文转换为 Word 文档，使用标准 Heading 样式',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
使用示例:
  python convert_to_word.py input.md
  python convert_to_word.py input.md -o output.docx
        '''
    )
    
    parser.add_argument('input', help='输入的 Markdown 文件路径')
    parser.add_argument('-o', '--output', help='输出的 Word 文件路径（默认：input.docx）')
    
    args = parser.parse_args()
    
    # 检查输入文件
    if not os.path.exists(args.input):
        print(f"[ERROR] 文件不存在 - {args.input}")
        sys.exit(1)
    
    # 确定输出路径
    if args.output:
        output_path = args.output
    else:
        base_name = os.path.splitext(args.input)[0]
        output_path = f"{base_name}.docx"
    
    # 执行转换
    convert_markdown_to_word(args.input, output_path)


if __name__ == "__main__":
    main()
