#!/usr/bin/env python3
"""用 python-docx 生成课程报告时的常用小工具。

用法示例::

    from docx import Document
    import docx_helpers as h

    doc = Document()
    h.setup_document(doc)                    # A4、2.5 cm 边距、宋体五号、1.5 倍行距、标题样式
    h.add_heading(doc, "一、实验目的", 1)
    h.add_body(doc, "本实验测量……")
    h.add_figure(doc, "curve.png", "电压随时间变化曲线")   # 自动按像素与页宽算尺寸，配图题
    h.add_table(doc, [["时间/s", "电压/V"], ["0", "0.00"]], caption="原始数据")
    h.add_code(doc, "print('hello')")
    doc.save("课程名_报告标题_姓名.docx")

所有字号不小于 10 pt；图片宽度默认不超过 15 cm、高度不超过 18 cm，并按 150 dpi 下限压尺寸。
"""
from __future__ import annotations

import os
from typing import Iterable, Sequence

from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

try:
    from PIL import Image
except ImportError:  # pragma: no cover
    Image = None

BODY_CN = "宋体"
BODY_EN = "Times New Roman"
HEADING_CN = {1: "宋体", 2: "黑体", 3: "黑体"}
HEADING_PT = {1: 15, 2: 12, 3: 10.5}
HEADING_BOLD = {1: True, 2: False, 3: False}
CAPTION_CN = "楷体"
CODE_FONT = "Consolas"
BODY_PT = 10.5
CAPTION_PT = 10.5
CODE_PT = 10
MIN_PT = 10

MAX_FIG_W_CM = 15.0
MAX_FIG_H_CM = 18.0
MIN_DPI = 150

_counters = {"图": 0, "表": 0}


# ---------- 字体 ----------

def set_run_font(run, cn: str = BODY_CN, en: str = BODY_EN, size_pt: float = BODY_PT,
                 bold: bool | None = None, italic: bool | None = None):
    """同时设置中文（eastAsia）和西文字体。字号低于 10 pt 会被抬到 10 pt。"""
    size_pt = max(size_pt, MIN_PT)
    run.font.name = en
    run.font.size = Pt(size_pt)
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    rfonts.set(qn("w:ascii"), en)
    rfonts.set(qn("w:hAnsi"), en)
    rfonts.set(qn("w:eastAsia"), cn)
    rfonts.set(qn("w:cs"), en)
    if bold is not None:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic
    run.font.color.rgb = RGBColor(0, 0, 0)


def _style_fonts(style, cn: str, en: str, size_pt: float, bold: bool | None = None):
    style.font.name = en
    style.font.size = Pt(max(size_pt, MIN_PT))
    if bold is not None:
        style.font.bold = bold
    style.font.color.rgb = RGBColor(0, 0, 0)
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for k in ("w:ascii", "w:hAnsi", "w:cs"):
        rfonts.set(qn(k), en)
    rfonts.set(qn("w:eastAsia"), cn)
    for k in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        if rfonts.get(qn(k)) is not None:
            del rfonts.attrib[qn(k)]


def _set_first_line_chars(paragraph, chars: int):
    """首行缩进按“字符”设置，Word 里显示为“首行缩进 2 字符”。"""
    ppr = paragraph._p.get_or_add_pPr()
    ind = ppr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        ppr.append(ind)
    if chars:
        ind.set(qn("w:firstLineChars"), str(chars * 100))
    else:
        for k in ("w:firstLineChars", "w:firstLine"):
            if ind.get(qn(k)) is not None:
                del ind.attrib[qn(k)]


# ---------- 文档级设置 ----------

def setup_document(doc, margins_cm: float = 2.5, landscape: bool = False):
    """A4、四边等宽页边距、正文与标题样式。"""
    for section in doc.sections:
        if landscape:
            section.page_width, section.page_height = Cm(29.7), Cm(21.0)
        else:
            section.page_width, section.page_height = Cm(21.0), Cm(29.7)
        section.left_margin = section.right_margin = Cm(margins_cm)
        section.top_margin = section.bottom_margin = Cm(margins_cm)
        section.gutter = Cm(0)

    normal = doc.styles["Normal"]
    _style_fonts(normal, BODY_CN, BODY_EN, BODY_PT)
    pf = normal.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = 1.5
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    for level in (1, 2, 3):
        try:
            st = doc.styles[f"Heading {level}"]
        except KeyError:
            continue
        _style_fonts(st, HEADING_CN[level], BODY_EN, HEADING_PT[level], HEADING_BOLD[level])
        hp = st.paragraph_format
        hp.alignment = WD_ALIGN_PARAGRAPH.LEFT
        hp.keep_with_next = True
        hp.space_before = Pt({1: 12, 2: 8, 3: 6}[level])
        hp.space_after = Pt({1: 6, 2: 4, 3: 3}[level])
        hp.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        hp.line_spacing = 1.5
        hp.first_line_indent = Cm(0)
    return doc


def text_width_cm(doc, section_index: int = 0) -> float:
    s = doc.sections[section_index]
    return (s.page_width - s.left_margin - s.right_margin) / 360000


# ---------- 段落 ----------

def add_heading(doc, text: str, level: int = 1):
    p = doc.add_heading(level=level)
    run = p.add_run(text)
    set_run_font(run, HEADING_CN.get(level, "黑体"), BODY_EN, HEADING_PT.get(level, 10.5),
                 bold=HEADING_BOLD.get(level, False))
    p.paragraph_format.keep_with_next = True
    _set_first_line_chars(p, 0)
    return p


def add_body(doc, text: str, indent_chars: int = 2, bold: bool = False):
    p = doc.add_paragraph()
    run = p.add_run(text)
    set_run_font(run, BODY_CN, BODY_EN, BODY_PT, bold=bold)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _set_first_line_chars(p, indent_chars)
    return p


def add_caption(doc, kind: str, text: str, number: int | None = None):
    """图题/表题：楷体五号居中。kind 为 '图' 或 '表'。自动编号。"""
    if number is None:
        _counters[kind] = _counters.get(kind, 0) + 1
        number = _counters[kind]
    p = doc.add_paragraph()
    run = p.add_run(f"{kind} {number} {text}")
    set_run_font(run, CAPTION_CN, BODY_EN, CAPTION_PT)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_first_line_chars(p, 0)
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(6)
    return p, number


def reset_counters():
    for k in _counters:
        _counters[k] = 0


# ---------- 图片 ----------

def image_pixels(path: str):
    if Image is None:
        return None
    try:
        with Image.open(path) as im:
            return im.size
    except Exception:
        return None


def fit_size(px_w: int, px_h: int, max_w_cm: float = MAX_FIG_W_CM, max_h_cm: float = MAX_FIG_H_CM,
             min_dpi: float = MIN_DPI, target_w_cm: float | None = None):
    """按像素和限制算出插入尺寸（cm）。保持比例；宽高都不超上限；有效 dpi 不低于 min_dpi。"""
    ratio = px_h / px_w
    w = target_w_cm if target_w_cm else max_w_cm
    w = min(w, max_w_cm)
    dpi_cap_w = px_w / min_dpi * 2.54 if min_dpi else w
    if dpi_cap_w < w:
        w = max(dpi_cap_w, 6.0) if dpi_cap_w < 6.0 else dpi_cap_w
    h = w * ratio
    if h > max_h_cm:
        h = max_h_cm
        w = h / ratio
    return round(w, 2), round(h, 2)


def add_figure(doc, image_path: str, caption: str, width_cm: float | None = None,
               max_w_cm: float = MAX_FIG_W_CM, max_h_cm: float = MAX_FIG_H_CM, min_dpi: float = MIN_DPI):
    """居中插入图片并在下方加图题。width_cm 是期望宽度，最终仍受上限约束。"""
    px = image_pixels(image_path)
    if px:
        w, h = fit_size(px[0], px[1], max_w_cm, max_h_cm, min_dpi, width_cm)
    else:
        w, h = min(width_cm or max_w_cm, max_w_cm), None
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_first_line_chars(p, 0)
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.space_before = Pt(6)
    run = p.add_run()
    if h:
        run.add_picture(image_path, width=Cm(w), height=Cm(h))
    else:
        run.add_picture(image_path, width=Cm(w))
    cap, n = add_caption(doc, "图", caption)
    return p, n


def add_side_by_side_figures(doc, image_paths: Sequence[str], caption: str, total_w_cm: float = 15.0,
                             gap_cm: float = 0.3, sub_labels: bool = True):
    """两三张小图并排：无边框单行表格，共用一个图题，子图用（a）（b）标注。"""
    n = len(image_paths)
    each = (total_w_cm - gap_cm * (n - 1)) / n
    table = doc.add_table(rows=2 if sub_labels else 1, cols=n)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    for i, path in enumerate(image_paths):
        cell = table.cell(0, i)
        cell.width = Cm(each)
        para = cell.paragraphs[0]
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        px = image_pixels(path)
        if px:
            w, h = fit_size(px[0], px[1], each, MAX_FIG_H_CM, MIN_DPI, each)
            para.add_run().add_picture(path, width=Cm(w), height=Cm(h))
        else:
            para.add_run().add_picture(path, width=Cm(each))
        if sub_labels:
            lab = table.cell(1, i).paragraphs[0]
            lab.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_run_font(lab.add_run(f"（{chr(ord('a') + i)}）"), CAPTION_CN, BODY_EN, CAPTION_PT)
    cap, num = add_caption(doc, "图", caption)
    return table, num


# ---------- 表格 ----------

def _cell_shading(cell, hex_fill: str):
    tcpr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tcpr.append(shd)


def add_table(doc, rows: Sequence[Sequence[str]], caption: str | None = None, header: bool = True,
              col_widths_cm: Sequence[float] | None = None, font_pt: float = BODY_PT,
              center_cols: Iterable[int] | None = None, style: str = "Table Grid"):
    """原生表格：表题在上、居中、五号字、表头加粗居中、重复表头。"""
    if caption:
        cap, _ = add_caption(doc, "表", caption)
        cap.paragraph_format.keep_with_next = True
    n_rows, n_cols = len(rows), max(len(r) for r in rows)
    table = doc.add_table(rows=n_rows, cols=n_cols)
    table.style = style
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    center_cols = set(center_cols) if center_cols is not None else set()
    for i, row in enumerate(rows):
        for j in range(n_cols):
            text = str(row[j]) if j < len(row) else ""
            cell = table.cell(i, j)
            cell.text = ""
            para = cell.paragraphs[0]
            run = para.add_run(text)
            is_header = header and i == 0
            set_run_font(run, BODY_CN, BODY_EN, font_pt, bold=is_header)
            para.paragraph_format.space_before = Pt(0)
            para.paragraph_format.space_after = Pt(0)
            para.paragraph_format.line_spacing = 1.0
            if is_header or j in center_cols or _looks_numeric(text):
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            if col_widths_cm and j < len(col_widths_cm):
                cell.width = Cm(col_widths_cm[j])
    if header:
        trpr = table.rows[0]._tr.get_or_add_trPr()
        th = OxmlElement("w:tblHeader")
        th.set(qn("w:val"), "true")
        trpr.append(th)
    doc.add_paragraph().paragraph_format.space_after = Pt(0)
    return table


def _looks_numeric(text: str) -> bool:
    t = text.replace(",", "").replace("%", "").replace("−", "-").strip()
    if not t:
        return False
    try:
        float(t)
        return True
    except ValueError:
        return False


# ---------- 代码块 ----------

def add_code(doc, code: str, font_pt: float = CODE_PT, shade: str = "F5F5F5", max_line: int = 90):
    """等宽字体、浅灰底纹、保留缩进；超长行自动折行并加 ↩ 标记。"""
    lines = []
    for raw in code.rstrip("\n").split("\n"):
        line = raw.replace("\t", "    ")
        while len(line) > max_line:
            lines.append(line[:max_line] + " ↩")
            line = "    " + line[max_line:]
        lines.append(line)
    paras = []
    for line in lines:
        p = doc.add_paragraph()
        run = p.add_run(line if line else " ")
        set_run_font(run, CODE_FONT, CODE_FONT, font_pt)
        run._element.rPr.rFonts.set(qn("w:eastAsia"), BODY_CN)
        pf = p.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing = 1.0
        pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf.keep_with_next = True
        _set_first_line_chars(p, 0)
        ppr = p._p.get_or_add_pPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), shade)
        ppr.append(shd)
        paras.append(p)
    if paras:
        paras[-1].paragraph_format.keep_with_next = False
        paras[-1].paragraph_format.space_after = Pt(6)
    return paras


def add_page_break(doc):
    from docx.enum.text import WD_BREAK
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def add_page_number_footer(doc, size_pt: float = 10.5):
    """页脚居中页码（PAGE 域）。"""
    for section in doc.sections:
        footer = section.footer
        p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        set_run_font(run, BODY_CN, BODY_EN, size_pt)
        for tag, text in (("begin", None), (None, "PAGE"), ("end", None)):
            if tag:
                el = OxmlElement("w:fldChar")
                el.set(qn("w:fldCharType"), tag)
            else:
                el = OxmlElement("w:instrText")
                el.set(qn("xml:space"), "preserve")
                el.text = text
            run._element.append(el)
