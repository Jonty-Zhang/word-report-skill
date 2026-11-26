#!/usr/bin/env python3
"""用 python-docx 生成课程报告时的常用小工具。

用法示例::

    from docx import Document
    import docx_helpers as h

    doc = Document()
    h.reset_counters(by_chapter=False)       # 论文或 ≥10 页的文档用 True：图 3.2；短报告 False：图 1
    h.setup_document(doc)                    # A4、2.5 cm 边距、宋体五号、1.5 倍行距、标题样式
    h.add_heading(doc, "一、实验目的", 1)      # 标题带大纲级别
    h.add_body(doc, "电压变化见[[fig:curve]]，原始数据见[[tab:raw]]。")   # 交叉引用，可先于图表出现
    h.add_figure(doc, "curve.png", "电压随时间变化曲线", key="curve")     # 题注用 SEQ 域，自动算尺寸
    h.add_table(doc, [["时间/s", "电压/V"], ["0", "0.00"]], caption="原始数据", key="raw")   # 三线表
    h.add_table(doc, rows, caption="充放电对比", group_header=[("", 1), ("充电", 2), ("放电", 2)],
                note="注：电压为三次测量平均值。", total_row=True)
    h.add_code(doc, "print('hello')")
    h.add_body(doc, "已有工作多用阶跃响应拟合[[cite:zhang2020,li2019]]。")   # 上标文献引用 [1,2]
    h.add_references(doc, [                                                 # 悬挂缩进，按首次引用排序
        ("zhang2020", h.gbt7714("J", ["张三"], "题名", journal="刊名", year=2020, volume=3, issue=2, pages="1-5")),
        ("li2019", h.gbt7714("M", ["李四"], "书名", place="北京", publisher="出版社", year=2019, pages="10-12"))])
    h.finalize(doc)                          # 回填交叉引用；引用不存在的 key 会报错
    doc.save("课程名_报告标题_姓名.docx")

所有字号不小于 10 pt；图片宽度默认不超过 15 cm、高度不超过 18 cm，并按 150 dpi 下限压尺寸；
表格一律三线表；标题段落写入大纲级别，导航窗格可见。
"""
from __future__ import annotations

import os
import re
from typing import Iterable, Sequence

from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor, Twips

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

# 题注编号状态。CAPTION_BY_CHAPTER=True 时按章编号（图 3.2），否则连续编号（图 1）。
CAPTION_BY_CHAPTER = False
UNNUMBERED_HEADINGS = {"摘要", "摘 要", "abstract", "目录", "目 录", "参考文献", "致谢", "致 谢", "附录", "后记"}
_state = {"chapter": 0, "图": 0, "表": 0, "bm": 0, "refs": {}, "cite_order": [], "cite_bm": {}, "bib_done": False}


# ---------- OOXML 子元素顺序 ----------
# Word 对 pPr / tblPr / tcPr / trPr / tcBorders 的子元素顺序有严格要求，顺序错了会被忽略或报“无法读取的内容”。
_ORDERS = {
    "pPr": ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
            "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
            "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
            "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc",
            "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
            "sectPr", "pPrChange"],
    "tblPr": ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize",
              "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar",
              "tblLook", "tblCaption", "tblDescription"],
    "tcPr": ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar",
             "textDirection", "tcFitText", "vAlign", "hideMark"],
    "trPr": ["cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore", "wAfter", "cantSplit", "trHeight",
             "tblHeader", "tblCellSpacing", "jc", "hidden"],
    "tcBorders": ["top", "start", "left", "bottom", "end", "right", "insideH", "insideV", "tl2br", "tr2bl"],
}


def _local(tag: str) -> str:
    return tag.split("}", 1)[1] if "}" in tag else tag


def _place(parent, el):
    """按 schema 顺序把 el 放进 parent；同名旧元素先删掉。"""
    order = _ORDERS[_local(parent.tag)]
    name = _local(el.tag)
    for old in parent.findall(el.tag):
        parent.remove(old)
    idx = order.index(name)
    for i, child in enumerate(list(parent)):
        cname = _local(child.tag)
        if cname in order and order.index(cname) > idx:
            child.addprevious(el)
            return el
    parent.append(el)
    return el


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
        ind = _place(ppr, OxmlElement("w:ind"))
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

def set_outline_level(paragraph, level: int):
    """给段落写入大纲级别（1 起算）。Word 导航窗格和自动目录都靠它识别层级。"""
    ppr = paragraph._p.get_or_add_pPr()
    el = OxmlElement("w:outlineLvl")
    el.set(qn("w:val"), str(max(0, min(8, level - 1))))
    _place(ppr, el)


def add_heading(doc, text: str, level: int = 1):
    """标题：用标题样式，并在段落上显式写入大纲级别（套用模板时样式里未必有）。"""
    p = doc.add_heading(level=level)
    run = p.add_run(text)
    set_run_font(run, HEADING_CN.get(level, "黑体"), BODY_EN, HEADING_PT.get(level, 10.5),
                 bold=HEADING_BOLD.get(level, False))
    p.paragraph_format.keep_with_next = True
    _set_first_line_chars(p, 0)
    set_outline_level(p, level)
    if level == 1:
        key = text.strip().lower()
        unnumbered = any(key.startswith(u) for u in UNNUMBERED_HEADINGS)
        if not unnumbered:
            _state["chapter"] += 1
            # 隐藏的章计数域：题注里的 {SEQ 章 \c} 取它的值
            _append_field(p, " SEQ 章 \\h ", "", HEADING_CN.get(1, "宋体"), HEADING_PT.get(1, 15))
        _state["图"] = 0
        _state["表"] = 0
    return p


_REF_RE = re.compile(r"\[\[(fig|tab|cite|cite-text):([^\]]+)\]\]")


def add_body(doc, text: str, indent_chars: int = 2, bold: bool = False):
    """正文段落。标记会变成交叉引用（REF 域）：
    [[fig:key]] / [[tab:key]]   显示为“图 3.2”“表 1”，可以出现在图表之前，finalize(doc) 统一回填
    [[cite:a,b,c]]              上标文献引用，如 ^[1]、^[2-4]、^[1,3]
    [[cite-text:a]]             不上标的引用，用于“见文献[3]”这种作句子成分的场合"""
    p = doc.add_paragraph()
    pos = 0
    for m in _REF_RE.finditer(text):
        if m.start() > pos:
            set_run_font(p.add_run(text[pos:m.start()]), BODY_CN, BODY_EN, BODY_PT, bold=bold)
        if m.group(1) in ("fig", "tab"):
            add_ref(p, m.group(1), m.group(2).strip(), bold=bold)
        else:
            keys = [k.strip() for k in re.split(r"[,，]", m.group(2)) if k.strip()]
            add_cite(p, keys, superscript=(m.group(1) == "cite"), bold=bold)
        pos = m.end()
    if pos < len(text):
        set_run_font(p.add_run(text[pos:]), BODY_CN, BODY_EN, BODY_PT, bold=bold)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _set_first_line_chars(p, indent_chars)
    return p


# ---------- 题注与交叉引用（域） ----------

def _new_run(paragraph, cn: str, size_pt: float, bold: bool = False, superscript: bool = False):
    run = paragraph.add_run()
    set_run_font(run, cn, BODY_EN, size_pt, bold=bold)
    if superscript:
        run.font.superscript = True
    return run


def _append_field(paragraph, instr: str, result: str, cn: str = BODY_CN, size_pt: float = BODY_PT,
                  bold: bool = False, superscript: bool = False):
    """追加一个复杂域：begin / instrText / separate / 缓存结果 / end。返回结果文字所在的 run。"""
    def fld(kind):
        r = _new_run(paragraph, cn, size_pt, bold, superscript)
        el = OxmlElement("w:fldChar")
        el.set(qn("w:fldCharType"), kind)
        r._element.append(el)
        return r
    fld("begin")
    r = _new_run(paragraph, cn, size_pt, bold, superscript)
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = instr
    r._element.append(it)
    fld("separate")
    res = _new_run(paragraph, cn, size_pt, bold, superscript)
    res.text = result
    fld("end")
    return res


def _bookmark(paragraph, name: str, start: bool):
    if start:
        _state["bm"] += 1
        el = OxmlElement("w:bookmarkStart")
        el.set(qn("w:id"), str(_state["bm"]))
        el.set(qn("w:name"), name)
        paragraph._p.append(el)
        return _state["bm"]
    return None


def _bookmark_end(paragraph, bm_id: int):
    el = OxmlElement("w:bookmarkEnd")
    el.set(qn("w:id"), str(bm_id))
    paragraph._p.append(el)


def add_caption(doc, kind: str, text: str, key: str | None = None):
    """图题/表题（Word 题注）：楷体五号居中。
    编号用 SEQ 域，按章编号时为 “图 {SEQ 章 \\c}.{SEQ 图 \\s 1}”；
    “标签 + 编号”包在书签里，正文用 REF 域引用它（等同于 Word 的“交叉引用 → 只有标签和编号”）。"""
    _state[kind] += 1
    n = _state[kind]
    by_chapter = CAPTION_BY_CHAPTER and _state["chapter"] > 0
    label = f"{kind} {_state['chapter']}.{n}" if by_chapter else f"{kind} {n}"
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_first_line_chars(p, 0)
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.first_line_indent = Cm(0)
    bm_name = "_Ref%09d" % (100000000 + len(_state["refs"]) + _state["bm"] + 1)
    bm_id = _bookmark(p, bm_name, True)
    r = _new_run(p, CAPTION_CN, CAPTION_PT)
    r.text = f"{kind} "
    if by_chapter:
        _append_field(p, " SEQ 章 \\c ", str(_state["chapter"]), CAPTION_CN, CAPTION_PT)
        r = _new_run(p, CAPTION_CN, CAPTION_PT)
        r.text = "."
        _append_field(p, f" SEQ {kind} \\* ARABIC \\s 1 ", str(n), CAPTION_CN, CAPTION_PT)
    else:
        _append_field(p, f" SEQ {kind} \\* ARABIC ", str(n), CAPTION_CN, CAPTION_PT)
    _bookmark_end(p, bm_id)
    r = _new_run(p, CAPTION_CN, CAPTION_PT)
    r.text = f" {text}"
    ref_key = f"{'fig' if kind == '图' else 'tab'}:{key or bm_name}"
    _state["refs"][ref_key] = (bm_name, label)
    return p, label


def add_ref(paragraph, kind: str, key: str, bold: bool = False):
    """在段落里插入对图/表的交叉引用（REF 域）。kind 为 'fig' 或 'tab'。"""
    k = f"{kind}:{key}"
    bm, label = _state["refs"].get(k, (f"__pending__{k}", f"[[{k}]]"))
    return _append_field(paragraph, f" REF {bm} \\h ", label, BODY_CN, BODY_PT, bold)


# ---------- GB/T 7714—2025 著录格式 ----------

GBT_TYPES = {"M", "J", "N", "C", "D", "R", "S", "P", "EB", "A", "CM", "DS", "PP", "G", "CP", "DB", "Z"}
GBT_CARRIERS = {"OL", "CD", "DK", "MT", "MM"}
_CJK = re.compile(r"[一-鿿]")


def _dash(pages):
    return re.sub(r"\s*[–—~～－]\s*", "-", str(pages)) if pages is not None else None


def gbt7714(kind: str, authors=None, title: str = "", *, subtitle=None, journal=None, year=None,
            volume=None, issue=None, pages=None, edition=None, place=None, publisher=None, school=None,
            conference=None, report_no=None, date=None, std_no=None, patent_no=None, newspaper=None,
            version=None, platform=None, updated=None, accessed=None, url=None, pid=None,
            host_authors=None, host_title=None, online=None, lang=None) -> str:
    """按 GB/T 7714—2025 生成一条著录文本（不含 [n] 序号），交给 add_references。

    kind: J 期刊 | M 图书 | M// 图书析出 | D 学位论文 | C 会议录析出 | R 报告 | S 标准 | P 专利
          | N 报纸 | EB 网页 | DS 数据集 | PP 预印本
    authors: 列表，超过 3 个自动写成“前 3 个,等”/“, et al.”。西文名按“Smith J A”传入（姓只首字母大写）。
    accessed: 引用日期 YYYY-MM-DD，只用于网络资源（EB/DS/PP 必填；其他类型有 url 时可填）。
    url / pid: 获取和访问路径、永久标识符（DOI、CSTR）；pid 已包含在 url 里时不要重复传。
    中文条目符号后不加空格，西文条目加空格（lang 不传时按题名有无汉字判断）。
    缺必备项会直接抛 ValueError，不让残缺条目进文档。
    """
    kind = kind.upper()
    zh = (lang == "zh") if lang else bool(_CJK.search(title or "") or _CJK.search(str(journal or "")))
    sp = "" if zh else " "

    def need(**kw):
        miss = [k for k, v in kw.items() if v in (None, "", [])]
        if miss:
            raise ValueError(f"[{kind}] 缺必备项：{', '.join(miss)}（题名：{title}）")

    def names(lst):
        if not lst:
            return ""
        lst = [str(a).strip() for a in lst]
        if len(lst) > 3:
            lst = lst[:3] + (["等"] if zh else ["et al."])
        return ("," if zh else ", ").join(lst)

    is_online = online if online is not None else bool(url) or kind in ("EB", "DS", "PP")
    tag = kind.rstrip("/")
    if tag == "M//":
        tag = "M"
    code = f"[{tag}/OL]" if is_online and tag not in ("S",) else f"[{tag}]"
    if tag == "S" and is_online:
        code = "[S/OL]"
    title = re.sub(r"\s*：\s*", ":" + sp, title.strip())   # 题名与其他题名信息之间用半角冒号
    full_title = f"{title}:{sp}{subtitle}" if subtitle else title
    head = f"{names(authors)}.{sp}" if authors else ""
    p = _dash(pages)

    if kind == "J":
        need(title=title, journal=journal, year=year)
        vi = f"{volume}({issue})" if volume and issue else (str(volume) if volume else (f"({issue})" if issue else ""))
        body = f"{journal},{sp}{year}" + (f",{sp}{vi}" if volume else (vi if issue else ""))
        body += f":{sp}{p}" if p else ""
        acc = f"[{accessed}]" if accessed and is_online else ""
        return f"{head}{full_title}{code}.{sp}{body}{acc}." + "".join(f"{sp}{e}." for e in (url, pid) if e)
    if kind in ("M", "M//"):
        need(title=title, publisher=publisher, year=year)
        ed = f"{edition}.{sp}" if edition else ""
        pub = f"{place}:{sp}{publisher},{sp}{year}" if place else f"{publisher},{sp}{year}"
        pub += f":{sp}{p}" if p else ""
        if kind == "M//":
            need(host_title=host_title)
            host = (f"{names(host_authors)}.{sp}" if host_authors else "") + host_title
            return f"{head}{full_title}{code}//{host}.{sp}{ed}{pub}." + "".join(f"{sp}{e}." for e in (url, pid) if e)
        return f"{head}{full_title}{code}.{sp}{ed}{pub}." + "".join(f"{sp}{e}." for e in (url, pid) if e)
    if kind == "D":
        need(authors=authors, title=title, place=place, school=school, year=year)
        return (f"{head}{full_title}{code}.{sp}{place}:{sp}{school},{sp}{year}" + (f":{sp}{p}" if p else "") + "."
                + "".join(f"{sp}{e}." for e in (url, pid) if e))
    if kind == "C":
        need(authors=authors, title=title, conference=conference, year=year)
        return (f"{head}{full_title}{code}//{conference},{sp}{year}" + (f":{sp}{p}" if p else "") + "."
                + "".join(f"{sp}{e}." for e in (url, pid) if e))
    if kind == "R":
        need(title=title, date=date)
        t = f"{full_title}:{sp}{report_no}" if report_no else full_title
        return (f"{head}{t}{code}.{sp}{date}" + (f":{sp}{p}" if p else "") + "."
                + "".join(f"{sp}{e}." for e in (url, pid) if e))
    if kind == "S":
        need(std_no=std_no, title=title)
        no = re.sub(r"(\d)\s*[-–－]\s*(\d{4})$", r"\1—\2", std_no.strip())
        return f"{no} {full_title}{code}." + "".join(f"{sp}{e}." for e in (url, pid) if e)
    if kind == "P":
        need(authors=authors, title=title, patent_no=patent_no, date=date)
        acc = f"[{accessed}]" if accessed and is_online else ""
        return f"{head}{full_title}:{sp}{patent_no}{code}.{sp}{date}{acc}." + "".join(f"{sp}{e}." for e in (url, pid) if e)
    if kind == "N":
        need(title=title, newspaper=newspaper, date=date)
        ver = f"({pages})" if pages else ""
        return f"{head}{full_title}{code}.{sp}{newspaper},{sp}{date}{ver}." + "".join(f"{sp}{e}." for e in (url, pid) if e)
    if kind == "EB":
        need(title=title, accessed=accessed, url=url)
        upd = f"({updated})" if updated else ""
        return f"{head}{full_title}{code}.{sp}{upd}[{accessed}].{sp}{url}." + (f"{sp}{pid}." if pid else "")
    if kind in ("DS", "PP"):
        need(authors=authors, title=title, platform=platform, accessed=accessed, url=url)
        ver = f"{version}.{sp}" if version else ""
        upd = f"({updated})" if updated else ""
        return (f"{head}{full_title}{code}.{sp}{ver}{platform}{upd}[{accessed}].{sp}{url}."
                + (f"{sp}{pid}." if pid else ""))
    raise ValueError(f"不支持的文献类型：{kind}")


# ---------- 参考文献与引用（域） ----------
# 文献序号是 SEQ 域，书签只包住数字；正文引用是 REF 域，方括号写在正文里。
# 这样多篇引用能合并成 [1,3]、[2-4]（GB/T 7714 顺序编码制），增删、调序后全选按 F9 即可重排。

BIB_HANG_CHARS = 2      # 悬挂缩进 2 字符，[1]～[99] 都放得下；≥100 条时自动用 2.5 字符


def _cite_info(key: str):
    if key not in _state["cite_order"]:
        _state["cite_order"].append(key)
    n = _state["cite_order"].index(key) + 1
    bm = _state["cite_bm"].setdefault(key, "Cite%03d_%s" % (n, re.sub(r"[^A-Za-z0-9]", "", key)[:24]))
    return n, bm


def _plain(paragraph, text: str, superscript: bool, bold: bool = False):
    r = _new_run(paragraph, BODY_CN, BODY_PT, bold, superscript)
    r.text = text
    return r


def add_cite(paragraph, keys: Sequence[str], superscript: bool = True, bold: bool = False):
    """插入文献引用。序号按首次引用的先后分配；连续 3 篇及以上写成范围 [2-4]，其余用逗号 [1,3]。"""
    items = sorted({_cite_info(k) for k in keys})
    groups, cur = [], [items[0]]
    for it in items[1:]:
        if it[0] == cur[-1][0] + 1:
            cur.append(it)
        else:
            groups.append(cur)
            cur = [it]
    groups.append(cur)
    _plain(paragraph, "[", superscript, bold)
    for gi, g in enumerate(groups):
        if gi:
            _plain(paragraph, ",", superscript, bold)
        if len(g) >= 3:
            seq = [g[0], None, g[-1]]
        else:
            seq = []
            for i, it in enumerate(g):
                if i:
                    seq.append("sep")
                seq.append(it)
        for it in seq:
            if it is None:
                _plain(paragraph, "-", superscript, bold)
            elif it == "sep":
                _plain(paragraph, ",", superscript, bold)
            else:
                n, bm = it
                _append_field(paragraph, f" REF {bm} \\h \\* MERGEFORMAT ", str(n), BODY_CN, BODY_PT, bold, superscript)
    _plain(paragraph, "]", superscript, bold)


def add_references(doc, entries, heading: str = "参考文献", page_break: bool = False,
                   allow_uncited: bool = False, font_pt: float = BODY_PT):
    """生成参考文献表（放在正文所有引用之后调用）。

    entries        [(key, 著录文本), ...] 或 {key: 文本}；著录文本用 gbt7714(...) 生成（GB/T 7714—2025）。顺序无所谓，
                   最终按正文首次引用的先后排序（顺序编码制）
    page_break     论文类另起一页
    allow_uncited  默认不允许列出正文没引用的文献（会报错）；True 时排在已引用文献之后
    版式：不首行缩进，悬挂缩进 2 字符，序号后接制表符，续行与第一行文字对齐；五号，两端对齐。
    """
    items = list(entries.items()) if isinstance(entries, dict) else list(entries)
    texts = dict(items)
    missing = [k for k in _state["cite_order"] if k not in texts]
    if missing:
        raise KeyError(f"正文引用了参考文献表里没有的条目：{missing}")
    uncited = [k for k, _ in items if k not in _state["cite_order"]]
    if uncited and not allow_uncited:
        raise ValueError(f"这些文献正文没有引用，删掉或在正文引用：{uncited}")
    for k in uncited:
        _cite_info(k)
    order = list(_state["cite_order"])
    if page_break:
        add_page_break(doc)
    add_heading(doc, heading, 1)
    hang_chars = BIB_HANG_CHARS if len(order) < 100 else BIB_HANG_CHARS + 0.5
    hang_twips = int(round(hang_chars * font_pt * 20))
    for key in order:
        n, bm = _cite_info(key)
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf = p.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        ppr = p._p.get_or_add_pPr()
        ind = OxmlElement("w:ind")
        ind.set(qn("w:left"), str(hang_twips))
        ind.set(qn("w:leftChars"), "0")
        ind.set(qn("w:hanging"), str(hang_twips))
        ind.set(qn("w:hangingChars"), str(int(hang_chars * 100)))
        _place(ppr, ind)
        pf.tab_stops.add_tab_stop(Twips(hang_twips))   # 显式制表位，WPS/LibreOffice 也能对齐
        r = _new_run(p, BODY_CN, font_pt)
        r.text = "["
        bm_id = _bookmark(p, bm, True)
        _append_field(p, " SEQ 参考文献 \\* ARABIC ", str(n), BODY_CN, font_pt)
        _bookmark_end(p, bm_id)
        r = _new_run(p, BODY_CN, font_pt)
        r.text = "]\t" + texts[key].strip()
    _state["bib_done"] = True
    return order


def finalize(doc):
    """回填所有交叉引用：把先于图表出现的 [[fig:key]] 换成真实书签和编号。
    引用了不存在的 key 会抛错，避免成品里出现“错误！未找到引用源”。"""
    missing = []
    for it in doc.element.body.iter(qn("w:instrText")):
        txt = it.text or ""
        if "__pending__" not in txt:
            continue
        k = txt.split("__pending__", 1)[1].split()[0]
        if k not in _state["refs"]:
            missing.append(k)
            continue
        bm, label = _state["refs"][k]
        it.text = f" REF {bm} \\h "
        # 找到该域的缓存结果 run，替换文字
        run = it.getparent()
        sib = run.getnext()
        while sib is not None:
            fc = sib.find(qn("w:fldChar"))
            if fc is not None and fc.get(qn("w:fldCharType")) == "separate":
                res = sib.getnext()
                t = res.find(qn("w:t")) if res is not None else None
                if t is not None:
                    t.text = label
                break
            sib = sib.getnext()
    if missing:
        raise KeyError(f"交叉引用找不到对应的图表：{sorted(set(missing))}")
    if _state["cite_order"] and not _state["bib_done"]:
        raise RuntimeError("正文有文献引用，但还没有调用 add_references 生成参考文献表")
    return doc


def reset_counters(by_chapter: bool | None = None):
    """新建文档前调用。by_chapter=True 按章编号（论文、≥10 页的文档），False 连续编号。"""
    global CAPTION_BY_CHAPTER
    if by_chapter is not None:
        CAPTION_BY_CHAPTER = by_chapter
    _state.update({"chapter": 0, "图": 0, "表": 0, "bm": 0, "refs": {}, "cite_order": [], "cite_bm": {},
                   "bib_done": False})


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
               max_w_cm: float = MAX_FIG_W_CM, max_h_cm: float = MAX_FIG_H_CM, min_dpi: float = MIN_DPI,
               key: str | None = None):
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
    cap, n = add_caption(doc, "图", caption, key=key)
    return p, n


def add_side_by_side_figures(doc, image_paths: Sequence[str], caption: str, total_w_cm: float = 15.0,
                             gap_cm: float = 0.3, sub_labels: bool = True, key: str | None = None):
    """两三张小图并排：无边框单行表格，共用一个图题，子图用（a）（b）标注。"""
    n = len(image_paths)
    each = (total_w_cm - gap_cm * (n - 1)) / n
    table = doc.add_table(rows=2 if sub_labels else 1, cols=n)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    _plain_table_style(doc, table)
    _set_tbl_borders(table, top=0, bottom=0)  # 排版用的无边框表
    for tag in ("w:top", "w:bottom"):
        el = table._tbl.tblPr.find(qn("w:tblBorders")).find(qn(tag))
        el.set(qn("w:val"), "nil")
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
    cap, num = add_caption(doc, "图", caption, key=key)
    return table, num


# ---------- 表格（三线表） ----------

RULE_THICK = 12    # 1.5 磅，单位 1/8 磅
RULE_MID = 6       # 0.75 磅
RULE_THIN = 4      # 0.5 磅
CHAR_CM = 0.37     # 五号汉字约 0.37 cm 宽
CELL_PAD_LR_CM = 0.19
CELL_PAD_TB_CM = 0.10
MIN_COL_CM = 1.8
LONG_TEXT_CHARS = 12


def _cell_shading(cell, hex_fill: str):
    tcpr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    _place(tcpr, shd)


def _display_len(text: str) -> float:
    """汉字算 1，西文、数字算 0.55。"""
    n = 0.0
    for ch in text:
        n += 1.0 if ord(ch) > 0x2E80 else 0.55
    return n


def _border_el(tag: str, val: str = "single", sz: int = 0):
    el = OxmlElement(tag)
    el.set(qn("w:val"), val)
    if val not in ("nil", "none"):
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
    return el


def _plain_table_style(doc, table):
    """显式套用“普通表格”样式并关闭全部条件格式。
    不这样做时，表格会继承文档（模板）的默认表格样式，首行/末行/镶边行的条件格式可能自动加线，变成双线。"""
    tblpr = table._tbl.tblPr
    try:
        table.style = doc.styles["Normal Table"]
    except KeyError:
        for old in tblpr.findall(qn("w:tblStyle")):
            tblpr.remove(old)
    look = OxmlElement("w:tblLook")
    look.set(qn("w:val"), "0000")
    for k in ("firstRow", "lastRow", "firstColumn", "lastColumn"):
        look.set(qn(f"w:{k}"), "0")
    look.set(qn("w:noHBand"), "1")
    look.set(qn("w:noVBand"), "1")
    _place(tblpr, look)
    for cs in tblpr.findall(qn("w:tblCellSpacing")):   # 单元格间距会把每条线画成双线
        tblpr.remove(cs)


def _set_tbl_borders(table, top=RULE_THICK, bottom=RULE_THICK):
    tblpr = table._tbl.tblPr
    b = OxmlElement("w:tblBorders")
    b.append(_border_el("w:top", "single", top))
    b.append(_border_el("w:left", "nil"))
    b.append(_border_el("w:bottom", "single", bottom))
    b.append(_border_el("w:right", "nil"))
    b.append(_border_el("w:insideH", "nil"))
    b.append(_border_el("w:insideV", "nil"))
    _place(tblpr, b)


def _set_cell_border(cell, edge: str, sz: int):
    tcpr = cell._tc.get_or_add_tcPr()
    borders = tcpr.find(qn("w:tcBorders"))
    if borders is None:
        borders = _place(tcpr, OxmlElement("w:tcBorders"))
    _place(borders, _border_el(f"w:{edge}", "single", sz))


def _set_cell_margins(table, tb_cm: float = CELL_PAD_TB_CM, lr_cm: float = CELL_PAD_LR_CM):
    tblpr = table._tbl.tblPr
    mar = OxmlElement("w:tblCellMar")
    for edge, cm in (("top", tb_cm), ("left", lr_cm), ("bottom", tb_cm), ("right", lr_cm)):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:w"), str(int(round(cm * 567))))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    _place(tblpr, mar)


def _set_table_width(table, widths_cm: Sequence[float]):
    tblpr = table._tbl.tblPr
    tw = OxmlElement("w:tblW")
    tw.set(qn("w:w"), str(int(round(sum(widths_cm) * 567))))
    tw.set(qn("w:type"), "dxa")
    _place(tblpr, tw)
    lay = OxmlElement("w:tblLayout")
    lay.set(qn("w:type"), "fixed")
    _place(tblpr, lay)
    grid = table._tbl.tblGrid
    for gc, w in zip(grid.findall(qn("w:gridCol")), widths_cm):
        gc.set(qn("w:w"), str(int(round(w * 567))))
    for row in table.rows:
        for j, cell in enumerate(row.cells):
            if j < len(widths_cm):
                cell.width = Cm(widths_cm[j])


def auto_col_widths(rows: Sequence[Sequence[str]], text_w_cm: float, n_cols: int) -> list[float]:
    """按每列最长内容分配列宽；总宽在正文宽 60%～100% 之间；每列 ≥1.8 cm。"""
    longest = [0.0] * n_cols
    for row in rows:
        for j in range(n_cols):
            t = str(row[j]) if j < len(row) else ""
            longest[j] = max(longest[j], _display_len(t))
    capped = [min(max(x, 2.0), 20.0) for x in longest]
    raw = [c * CHAR_CM + 2 * CELL_PAD_LR_CM + 0.3 for c in capped]
    natural = sum(raw)
    if n_cols >= 5 or max(longest) > 20:
        target = text_w_cm
    else:
        target = min(max(natural, 0.6 * text_w_cm), text_w_cm)
    widths = [w * target / natural for w in raw]
    # 保底宽度：不够的抬到 1.8 cm，再从宽列里扣回来
    for _ in range(5):
        short = [i for i, w in enumerate(widths) if w < MIN_COL_CM]
        if not short:
            break
        deficit = sum(MIN_COL_CM - widths[i] for i in short)
        for i in short:
            widths[i] = MIN_COL_CM
        donors = [i for i, w in enumerate(widths) if w > MIN_COL_CM and i not in short]
        pool = sum(widths[i] - MIN_COL_CM for i in donors)
        if pool <= 0:
            break
        for i in donors:
            widths[i] -= deficit * (widths[i] - MIN_COL_CM) / pool
    return [round(w, 2) for w in widths]


def _fill_cell(cell, text: str, font_pt: float, align):
    cell.text = ""
    para = cell.paragraphs[0]
    run = para.add_run(text)
    set_run_font(run, BODY_CN, BODY_EN, font_pt, bold=False)
    pf = para.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    pf.first_line_indent = Cm(0)
    _set_first_line_chars(para, 0)
    para.alignment = align
    from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    return para


def add_table(doc, rows: Sequence[Sequence[str]], caption: str | None = None, header: bool = True,
              col_widths_cm: Sequence[float] | None = None, font_pt: float = BODY_PT,
              center_cols: Iterable[int] | None = None, left_cols: Iterable[int] | None = None,
              group_header: Sequence[tuple[str, int]] | None = None, total_row: bool = False,
              note: str | None = None, empty_mark: str = "—", key: str | None = None):
    """三线表。

    rows          第一行是表头（header=True 时），其余是表身
    group_header  二级表头上层，如 [("", 1), ("充电", 2), ("放电", 2)]；组名为空的列与下层表头纵向合并
    total_row     最后一行是合计/平均，上沿加 0.5 磅合计线
    note          表注，以“注：”开头，放在表下
    key           交叉引用用的名字，正文里写 [[tab:key]]
    其余：表题在上；顶底线 1.5 磅，栏目线 0.75 磅，分组线 0.5 磅，全部单实线；无竖线、无底纹；
    “普通表格”样式并关闭条件格式（不继承模板的默认表格样式），不设单元格间距；
    表头居中不加粗；数字列和短列居中，长文字列左对齐；单倍行距、段前段后 0；
    行不跨页、表头重复；总宽 60%～100% 正文宽，自动列宽。
    """
    if caption:
        cap, _ = add_caption(doc, "表", caption, key=key)
        cap.paragraph_format.space_before = Pt(6)
        cap.paragraph_format.space_after = Pt(3)
        cap.paragraph_format.keep_with_next = True
    n_cols = max(len(r) for r in rows)
    if group_header:
        if sum(span for _, span in group_header) != n_cols:
            raise ValueError("group_header 的跨列数之和必须等于列数")
    n_head = (2 if group_header else 1) if header else 0
    n_rows = len(rows) + (1 if group_header else 0)
    table = doc.add_table(rows=n_rows, cols=n_cols)
    _plain_table_style(doc, table)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    widths = list(col_widths_cm) if col_widths_cm else auto_col_widths(rows, text_width_cm(doc), n_cols)
    center_cols = set(center_cols or [])
    left_cols = set(left_cols or [])
    body = rows[1:] if header else rows
    long_cols = {j for j in range(n_cols)
                 if any(_display_len(str(r[j]) if j < len(r) else "") > LONG_TEXT_CHARS for r in body)}

    offset = 1 if group_header else 0
    for i, row in enumerate(rows):
        ti = i + offset
        is_head = header and i == 0
        for j in range(n_cols):
            text = str(row[j]).strip() if j < len(row) else ""
            if not text and not is_head:
                text = empty_mark
            if is_head or j in center_cols:
                align = WD_ALIGN_PARAGRAPH.CENTER
            elif j in left_cols or (j in long_cols and not _looks_numeric(text)):
                align = WD_ALIGN_PARAGRAPH.LEFT
            else:
                align = WD_ALIGN_PARAGRAPH.CENTER
            _fill_cell(table.cell(ti, j), text, font_pt, align)

    if group_header:
        j = 0
        for label, span in group_header:
            top = table.cell(0, j)
            if label:
                merged = top.merge(table.cell(0, j + span - 1)) if span > 1 else top
                _fill_cell(merged, label, font_pt, WD_ALIGN_PARAGRAPH.CENTER)
                for k in range(j, j + span):
                    _set_cell_border(table.cell(0, k), "bottom", RULE_THIN)
            else:
                lower_text = str(rows[0][j]) if header else ""
                merged = top.merge(table.cell(1, j))
                _fill_cell(merged, lower_text, font_pt, WD_ALIGN_PARAGRAPH.CENTER)
            j += span

    _set_tbl_borders(table)
    if header:
        for j in range(n_cols):
            _set_cell_border(table.cell(n_head - 1, j), "bottom", RULE_MID)
    if total_row and len(body) >= 2:
        for j in range(n_cols):
            _set_cell_border(table.cell(n_rows - 1, j), "top", RULE_THIN)
    _set_cell_margins(table)
    _set_table_width(table, widths)

    short_table = n_rows <= 15
    for ri, row in enumerate(table.rows):
        trpr = row._tr.get_or_add_trPr()
        cs = OxmlElement("w:cantSplit")
        cs.set(qn("w:val"), "true")
        _place(trpr, cs)
        if ri < n_head:
            th = OxmlElement("w:tblHeader")
            th.set(qn("w:val"), "true")
            _place(trpr, th)
        if short_table and ri < n_rows - 1:
            for cell in row.cells:
                for para in cell.paragraphs:
                    para.paragraph_format.keep_with_next = True

    if note:
        np_ = doc.add_paragraph()
        set_run_font(np_.add_run(note if note.startswith("注") else "注：" + note), BODY_CN, BODY_EN, 10)
        np_.alignment = WD_ALIGN_PARAGRAPH.LEFT
        _set_first_line_chars(np_, 0)
        np_.paragraph_format.space_before = Pt(3)
        np_.paragraph_format.space_after = Pt(6)
        np_.paragraph_format.line_spacing = 1.0
    else:
        spacer = doc.add_paragraph()
        spacer.paragraph_format.space_before = Pt(0)
        spacer.paragraph_format.space_after = Pt(0)
        spacer.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        spacer.paragraph_format.line_spacing = Pt(6)
        _set_first_line_chars(spacer, 0)
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
        _place(ppr, shd)
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
