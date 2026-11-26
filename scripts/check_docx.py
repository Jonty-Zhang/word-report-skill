#!/usr/bin/env python3
"""检查课程类 Word 报告的常见问题。

用法::

    python3 check_docx.py 报告.docx
    选项：--min-font 10  --max-fig-width 15  --max-fig-height 18  --min-dpi 150
          --no-layout（有模板时跳过页面/边距检查）  --template（同上，并放宽字体检查提示）

检查项（代码 → 含义）：
  SMALL_FONT      字号小于下限（含表格、页眉页脚、脚注、文本框）
  IMG_TOO_WIDE    图片宽度超过正文区
  IMG_OVER_LIMIT  图片宽度超过建议上限（默认 15 cm）
  IMG_TOO_TALL    图片高度超过上限（默认 18 cm）
  IMG_DISTORTED   图片被拉伸变形（显示比例与像素比例差 >3%）
  IMG_LOW_DPI     图片有效分辨率低于下限（默认 150 dpi）
  IMG_FLOATING    浮动图片（建议嵌入型）
  IMG_NO_CAPTION  图片下方两段内没有“图 N”图题
  IMG_TINY        图片宽度小于 4 cm（可能是误插的小图）
  FORBIDDEN       元话语 / 占位符 / AI 自曝等禁用表述
  HEDGE           让步式自贬（“更好的做法是…但…”）
  MARKDOWN        Markdown 残留（#、**、反引号、```）
  AI_WORD         AI 味词汇超过阈值
  FLAW_SPREAD     缺点表述分散在多个章节
  PRONOUN_MIX     自称混用（笔者/本文 与 我 混用）
  READER_ADDRESS  正文对读者说“您”
  EMPTY_PARAS     连续空段
  EMOJI           表情符号
  EXCLAIM         感叹号过多
  DASH            破折号过多
  PAGE_SETUP      纸张 / 页边距与默认值不符
  TABLE_NOT_3LINE 数据表不是三线表：双线/虚线等非单实线、单元格间距、段落边框、带条件格式边框的表格样式
                  （含文档默认表格样式）、竖线、表身横线、缺顶/底/栏目线、底纹
  TABLE_RULE_WEIGHT 三线的线宽不对（顶底 1.5 磅、栏目线 0.75 磅）
  TABLE_SPACING   表格内段落有段前/段后间距（含从 Normal 样式继承来的）
  NO_OUTLINE      长得像标题（一、/（一）/1.1/第X章）却没有大纲级别
  OUTLINE_SKIP    标题跳级（一级下面直接三级）
  CAPTION_NO_FIELD 图题/表题编号是手打的，不是 SEQ 域（题注）
  CAPTION_SEQ     图/表编号重复或不连续
  XREF_MANUAL     正文里“图 N / 表 N”是手打的，不是交叉引用
  XREF_DANGLING   正文引用了不存在的图/表编号，或出现“错误！未找到引用源”
  CITE_MANUAL     正文文献引用 [n] 是手打的，不是指向参考文献的交叉引用
  CITE_DANGLING   引用了参考文献表里不存在的序号
  CITE_ORDER      首次引用顺序与文献序号不一致（顺序编码制应按首次引用先后编号）
  BIB_UNCITED     参考文献表里有正文没引用的条目
  BIB_SEQ         参考文献序号重复或断号
  BIB_NO_FIELD    文献序号是手打的（不是 SEQ 域或自动编号）
  BIB_INDENT      参考文献条目首行缩进了，或没有悬挂缩进（续行不对齐）
  BIB_FORMAT      著录不符合 GB/T 7714—2025：缺文献类型标识、代码不在代码表、印刷文献写了引用日期、
                  网页/数据集/预印本缺引用日期、全角著录符号、西文姓全大写（2015 写法）
ERROR 必须清零；WARN 逐条判断。退出码：有 ERROR 为 1，否则 0。
"""
from __future__ import annotations

import argparse
import io
import os
import re
import sys
from collections import Counter, defaultdict

from docx import Document
from docx.oxml.ns import qn
from lxml import etree

try:
    from PIL import Image
except ImportError:  # pragma: no cover
    Image = None

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
V = "urn:schemas-microsoft-com:vml"
EMU_PER_CM = 360000

# ---- 词表 ----
FORBIDDEN_PHRASES = [  # 高精度：出现即 ERROR
    "按照用户", "根据用户", "用户要求", "用户提供", "用户未提供", "用户没有提供", "按用户",
    "课堂要求", "老师要求", "老师的要求", "题目要求我们", "作业要求我们", "本次作业要求",
    "以下是", "如下是", "接下来给出",
    "这里插入", "此处插入", "待补充", "待填写", "待确认", "待完善", "（略）", "(略)",
    "占位符", "placeholder", "lorem ipsum",
    "作为AI", "作为一个AI", "作为人工智能", "作为一名AI", "我是一个语言模型", "语言模型生成", "知识截止",
    "希望对您有帮助", "希望对你有帮助", "请告诉我", "如您所见", "如你所见",
    "已满足要求", "满足了要求", "覆盖了全部要求",
    "生成本文", "使用python-docx", "由matplotlib绘制", "使用AI绘图",
]
FORBIDDEN_CASE_SENSITIVE = ["TODO", "TBD"]  # 区分大小写，出现即 ERROR
FORBIDDEN_SOFT = [  # 上下文相关：WARN，人工判断
    "按照要求", "根据要求", "字数要求", "严格按照格式", "AI生成", "如有需要", "XXX", "下面是本文", "此处放",
]
WORDCOUNT_RE = re.compile(r"不少于\s*\d+\s*字")
MONO_FONTS = ("consolas", "courier", "menlo", "monaco", "source code", "jetbrains", "fira code", "sf mono", "cascadia")
HEDGE_PATTERNS = [
    r"更好的做法", r"更好的选择", r"更合理的做法", r"理想情况下应", r"本应", r"严格来说应",
    r"为了简便", r"为了简单起见", r"为简便起见", r"限于时间", r"由于时间仓促", r"时间有限，", r"水平有限",
    r"才疏学浅", r"难免疏漏", r"限于篇幅",
]
AI_WORDS = {
    "此外": 3, "至关重要": 1, "深入探讨": 1, "赋能": 1, "标志着": 1, "彰显": 1, "凸显": 2, "不可或缺": 1,
    "奠定了坚实": 1, "有力支撑": 1, "具有重要意义": 1, "举足轻重": 1, "值得注意的是": 2, "值得一提的是": 2,
    "众所周知": 1, "毋庸置疑": 1, "显而易见": 1, "不仅": 3, "一系列": 3, "诸多": 3, "格局": 2, "抓手": 1,
    "闭环": 2, "深受启发": 1, "受益匪浅": 1, "圆满成功": 1, "完美": 1, "经过反复调试": 1, "不懈努力": 1,
}
SUMMARY_WORDS = ["综上所述", "总而言之", "总的来说", "总之"]
FLAW_WORDS = ["不足", "局限", "缺陷", "缺点", "有待", "尚未", "仍需", "仍有", "不够", "欠缺", "改进空间", "遗憾", "未能", "有限"]
FORMAL_SELF = ["笔者", "本文"]
CASUAL_SELF_RE = re.compile(r"我(?!们|国|校|院|方|省|市|党|军)")
WE_RE = re.compile(r"我们")
EMOJI_RE = re.compile("[\U0001F300-\U0001FAFF☀-➿⭐✅❌]")
MD_RE = [
    (re.compile(r"^\s*#{1,6}\s"), "标题井号"),
    (re.compile(r"\*\*[^*]+\*\*"), "**加粗**"),
    (re.compile(r"```"), "代码围栏"),
    (re.compile(r"^\s*[-*+]\s+\S"), "列表符号"),
    (re.compile(r"^\s*\|.*\|\s*$"), "Markdown 表格"),
    (re.compile(r"\[[^\]]+\]\(https?://[^)]+\)"), "Markdown 链接"),
]
BACKTICK_RE = re.compile(r"`[^`]+`")


class Report:
    def __init__(self):
        self.items = []  # (level, code, message)

    def error(self, code, msg):
        self.items.append(("ERROR", code, msg))

    def warn(self, code, msg):
        self.items.append(("WARN", code, msg))

    @property
    def errors(self):
        return [i for i in self.items if i[0] == "ERROR"]

    @property
    def warns(self):
        return [i for i in self.items if i[0] == "WARN"]


# ---- 字号解析 ----

def build_style_info(styles_el):
    info = {}
    default_para = None
    for st in styles_el.findall(qn("w:style")):
        sid = st.get(qn("w:styleId"))
        based = st.find(qn("w:basedOn"))
        sz = st.find(f"{{{W}}}rPr/{{{W}}}sz")
        name_el = st.find(qn("w:name"))
        info[sid] = {
            "based": based.get(qn("w:val")) if based is not None else None,
            "sz": int(sz.get(qn("w:val"))) if sz is not None and sz.get(qn("w:val")) else None,
            "name": name_el.get(qn("w:val")) if name_el is not None else sid,
            "type": st.get(qn("w:type")),
        }
        if st.get(qn("w:default")) in ("1", "true") and st.get(qn("w:type")) == "paragraph":
            default_para = sid
    dd = styles_el.find(f"{{{W}}}docDefaults/{{{W}}}rPrDefault/{{{W}}}rPr/{{{W}}}sz")
    default_sz = int(dd.get(qn("w:val"))) if dd is not None else 20
    return info, default_para, default_sz


def style_sz(sid, info, depth=0):
    while sid and sid in info and depth < 20:
        if info[sid]["sz"]:
            return info[sid]["sz"]
        sid = info[sid]["based"]
        depth += 1
    return None


def run_size_halfpts(run_el, para_el, info, default_para, default_sz):
    rpr = run_el.find(qn("w:rPr"))
    if rpr is not None:
        sz = rpr.find(qn("w:sz"))
        if sz is not None and sz.get(qn("w:val")):
            return int(sz.get(qn("w:val")))
        rs = rpr.find(qn("w:rStyle"))
        if rs is not None:
            v = style_sz(rs.get(qn("w:val")), info)
            if v:
                return v
    ppr = para_el.find(qn("w:pPr"))
    if ppr is not None:
        ps = ppr.find(qn("w:pStyle"))
        if ps is not None:
            v = style_sz(ps.get(qn("w:val")), info)
            if v:
                return v
    if default_para:
        v = style_sz(default_para, info)
        if v:
            return v
    return default_sz


def run_text(run_el):
    return "".join(t.text or "" for t in run_el.iter(qn("w:t")))


def para_text(p_el):
    return "".join(t.text or "" for t in p_el.iter(qn("w:t")))


def is_code_like(p_el):
    """等宽字体或带底纹的段落视为代码块，跳过 Markdown/占位符检查。"""
    ppr = p_el.find(qn("w:pPr"))
    if ppr is not None and ppr.find(qn("w:shd")) is not None:
        return True
    for rf in p_el.iter(qn("w:rFonts")):
        for k in ("w:ascii", "w:hAnsi"):
            v = (rf.get(qn(k)) or "").lower()
            if any(m in v for m in MONO_FONTS):
                return True
    return False


def is_heading(p_el, info):
    ppr = p_el.find(qn("w:pPr"))
    if ppr is None:
        return False
    if ppr.find(qn("w:outlineLvl")) is not None:
        return True
    ps = ppr.find(qn("w:pStyle"))
    if ps is None:
        return False
    sid = ps.get(qn("w:val"))
    name = info.get(sid, {}).get("name", sid) or sid
    return sid.lower().startswith("heading") or "标题" in name or name.lower().startswith("heading")


def location_of(p_el, body_index):
    anc = p_el
    while anc is not None:
        if anc.tag == qn("w:txbxContent"):
            return "文本框"
        if anc.tag == qn("w:tbl"):
            return "表格内"
        if anc.tag == qn("w:body"):
            break
        anc = anc.getparent()
    idx = body_index.get(id(p_el))
    return f"正文第 {idx} 段" if idx is not None else "正文"


# ---- 表格 / 大纲 / 题注 ----

HEADING_LIKE_RES = [
    re.compile(r"^[一二三四五六七八九十]{1,3}、\S"),
    re.compile(r"^（[一二三四五六七八九十]{1,3}）\S"),
    re.compile(r"^\d{1,2}(\.\d{1,2}){1,2}\s*\S"),
    re.compile(r"^第[一二三四五六七八九十\d]{1,3}[章节]"),
]
CAPTION_RE = re.compile(r"^(图|表)\s*(\d+)(?:[.．-](\d+))?(?=\s|[^\d.．-]|$)")
MENTION_RE = re.compile(r"(图|表)\s*(\d+(?:[.．]\d+)?)(?![\d.．])")
REF_ERROR_STRS = ["错误！未找到引用源", "错误!未找到引用源", "Error! Reference source not found", "错误！书签未定义"]


def style_outline(sid, styles_el, depth=0):
    """样式链上的大纲级别（0 起算），没有返回 None。"""
    while sid and depth < 20:
        st = None
        for x in styles_el.findall(qn("w:style")):
            if x.get(qn("w:styleId")) == sid:
                st = x
                break
        if st is None:
            return None
        ol = st.find(f"{{{W}}}pPr/{{{W}}}outlineLvl")
        if ol is not None:
            return int(ol.get(qn("w:val")))
        name = st.find(qn("w:name"))
        nm = (name.get(qn("w:val")) if name is not None else "").lower()
        m = re.match(r"heading (\d)", nm)
        if m:  # Word 内置标题样式自带大纲级别
            return int(m.group(1)) - 1
        b = st.find(qn("w:basedOn"))
        sid = b.get(qn("w:val")) if b is not None else None
        depth += 1
    return None


def para_outline(p_el, styles_el):
    ppr = p_el.find(qn("w:pPr"))
    if ppr is not None:
        ol = ppr.find(qn("w:outlineLvl"))
        if ol is not None:
            v = int(ol.get(qn("w:val")))
            return None if v >= 9 else v
        ps = ppr.find(qn("w:pStyle"))
        if ps is not None:
            return style_outline(ps.get(qn("w:val")), styles_el)
    return None


def in_table(el):
    anc = el.getparent()
    while anc is not None and anc.tag != qn("w:body"):
        if anc.tag == qn("w:tbl"):
            return True
        anc = anc.getparent()
    return False


def instr_texts(p_el):
    out = [(it.text or "") for it in p_el.iter(qn("w:instrText"))]
    out += [fs.get(qn("w:instr")) or "" for fs in p_el.iter(qn("w:fldSimple"))]
    return out


def _border_visible(el):
    if el is None:
        return None
    val = el.get(qn("w:val"))
    if val in ("nil", "none"):
        return 0
    try:
        return int(el.get(qn("w:sz")) or 4)
    except ValueError:
        return 4


SOLID = ("single",)
EDGES = ("top", "left", "bottom", "right", "insideH", "insideV", "start", "end")
LOOK_BITS = {"firstRow": 0x0020, "lastRow": 0x0040, "firstColumn": 0x0080, "lastColumn": 0x0100,
             "noHBand": 0x0200, "noVBand": 0x0400}


def _table_style_chain(tbl, styles_el):
    """表格样式链；表格没指定样式时用文档的默认表格样式（模板常把它设成带线的样式）。"""
    tblpr = tbl.find(qn("w:tblPr"))
    st = tblpr.find(qn("w:tblStyle")) if tblpr is not None else None
    sid = st.get(qn("w:val")) if st is not None else None
    if sid is None:
        for x in styles_el.findall(qn("w:style")):
            if x.get(qn("w:type")) == "table" and x.get(qn("w:default")) in ("1", "true"):
                sid = x.get(qn("w:styleId"))
    chain, depth = [], 0
    while sid and depth < 10:
        node = next((x for x in styles_el.findall(qn("w:style")) if x.get(qn("w:styleId")) == sid), None)
        if node is None:
            break
        chain.append(node)
        b = node.find(qn("w:basedOn"))
        sid = b.get(qn("w:val")) if b is not None else None
        depth += 1
    return chain


def _table_style_borders(tbl, styles_el, chain=None):
    tblpr = tbl.find(qn("w:tblPr"))
    chain = chain if chain is not None else _table_style_chain(tbl, styles_el)
    eff, vals = {}, {}
    sources = [node.find(f"{{{W}}}tblPr/{{{W}}}tblBorders") for node in reversed(chain)]
    if tblpr is not None:
        sources.append(tblpr.find(qn("w:tblBorders")))
    for tb in sources:
        if tb is None:
            continue
        for edge in EDGES:
            e = tb.find(qn(f"w:{edge}"))
            if e is not None:
                eff[edge] = _border_visible(e)
                vals[edge] = e.get(qn("w:val"))
    eff.setdefault("left", eff.get("start", 0) or 0)
    eff.setdefault("right", eff.get("end", 0) or 0)
    return eff, vals


def _look_flags(tbl):
    """tblLook：哪些条件格式开着。没写 tblLook 时按 Word 默认（首行、首列开）。"""
    look = tbl.find(f"{{{W}}}tblPr/{{{W}}}tblLook")
    flags = {"firstRow": True, "lastRow": False, "firstColumn": True, "lastColumn": False,
             "noHBand": False, "noVBand": True}
    if look is None:
        return flags
    v = look.get(qn("w:val"))
    if v:
        try:
            n = int(v, 16)
            flags = {k: bool(n & bit) for k, bit in LOOK_BITS.items()}
        except ValueError:
            pass
    for k in LOOK_BITS:
        a = look.get(qn(f"w:{k}"))
        if a is not None:
            flags[k] = a in ("1", "true", "on")
    return flags


def _style_conditional_lines(tbl, chain):
    """样式里的条件格式（首行、末行、镶边行……）会不会给表格加线。"""
    flags = _look_flags(tbl)
    active = {"wholeTable": True, "firstRow": flags["firstRow"], "lastRow": flags["lastRow"],
              "firstCol": flags["firstColumn"], "lastCol": flags["lastColumn"],
              "band1Horz": not flags["noHBand"], "band2Horz": not flags["noHBand"],
              "band1Vert": not flags["noVBand"], "band2Vert": not flags["noVBand"],
              "nwCell": flags["firstRow"] and flags["firstColumn"], "neCell": flags["firstRow"] and flags["lastColumn"],
              "swCell": flags["lastRow"] and flags["firstColumn"], "seCell": flags["lastRow"] and flags["lastColumn"]}
    hits = []
    for node in chain:
        name_el = node.find(qn("w:name"))
        sname = name_el.get(qn("w:val")) if name_el is not None else node.get(qn("w:styleId"))
        for cond in node.findall(qn("w:tblStylePr")):
            typ = cond.get(qn("w:type"))
            if not active.get(typ, False):
                continue
            for b in list(cond.iter(qn("w:tcBorders"))) + list(cond.iter(qn("w:tblBorders"))):
                if any(_border_visible(e) for e in b):
                    hits.append(f"{sname}/{typ}")
                    break
    return sorted(set(hits))


def check_tables(doc, rep):
    styles_el = doc.styles.element
    normal_after = None
    for x in styles_el.findall(qn("w:style")):
        if x.get(qn("w:type")) == "paragraph" and x.get(qn("w:default")) in ("1", "true"):
            sp = x.find(f"{{{W}}}pPr/{{{W}}}spacing")
            if sp is not None and sp.get(qn("w:after")):
                normal_after = int(sp.get(qn("w:after")))
    if normal_after is None:
        dd = styles_el.find(f"{{{W}}}docDefaults/{{{W}}}pPrDefault/{{{W}}}pPr/{{{W}}}spacing")
        if dd is not None and dd.get(qn("w:after")):
            normal_after = int(dd.get(qn("w:after")))
    spacing_hits = 0
    for ti, tbl in enumerate(doc.element.body.iter(qn("w:tbl")), 1):
        if tbl.find(f".//{{{WP}}}inline") is not None or tbl.find(f".//{{{WP}}}anchor") is not None:
            continue  # 排版用的图片表格
        rows = tbl.findall(qn("w:tr"))
        if not rows:
            continue
        chain = _table_style_chain(tbl, styles_el)
        eff, eff_vals = _table_style_borders(tbl, styles_el, chain)
        n_rows = len(rows)
        n_head = sum(1 for r in rows if r.find(f"{{{W}}}trPr/{{{W}}}tblHeader") is not None) or 1
        problems = []
        vertical = False
        top_sz = bottom_sz = header_sz = 0
        body_lines = 0
        shaded = False
        nonsolid = set()
        # 线型：凡是画出来的线，都只能是单实线
        for edge, v in eff_vals.items():
            if v not in (None, "nil", "none") and v not in SOLID and eff.get(edge):
                nonsolid.add(v)
        for b in tbl.iter(qn("w:tcBorders")):
            for e in b:
                v = e.get(qn("w:val"))
                if v not in (None, "nil", "none") and v not in SOLID:
                    nonsolid.add(v)
        # 单元格间距：会把每条线画成两条
        spacing_set = any(int(cs.get(qn("w:w")) or 0) > 0 for cs in tbl.iter(qn("w:tblCellSpacing")))
        # 单元格里的段落边框：常被拿来“补线”，结果出现双线
        pbdr = any(p.find(f"{{{W}}}pPr/{{{W}}}pBdr") is not None for p in tbl.iter(qn("w:p")))
        cond_lines = _style_conditional_lines(tbl, chain)
        for ri, tr in enumerate(rows):
            cells = tr.findall(qn("w:tc"))
            row_has_top = row_has_bottom = False
            for ci, tc in enumerate(cells):
                tcpr = tc.find(qn("w:tcPr"))
                tcb = tcpr.find(qn("w:tcBorders")) if tcpr is not None else None

                def edge(name, fallback):
                    if tcb is not None:
                        e = tcb.find(qn(f"w:{name}"))
                        if e is None and name in ("left", "right"):
                            e = tcb.find(qn("w:start" if name == "left" else "w:end"))
                        if e is not None:
                            return _border_visible(e)
                    return eff.get(fallback, 0) or 0

                l = edge("left", "left" if ci == 0 else "insideV")
                r = edge("right", "right" if ci == len(cells) - 1 else "insideV")
                if l or r:
                    vertical = True
                t = edge("top", "top" if ri == 0 else "insideH")
                b = edge("bottom", "bottom" if ri == n_rows - 1 else "insideH")
                if ri == 0:
                    top_sz = max(top_sz, t)
                if ri == n_rows - 1:
                    bottom_sz = max(bottom_sz, b)
                if ri == n_head - 1 and b > 0:
                    header_sz = max(header_sz, b)
                if ri == n_head and t > 0:
                    header_sz = max(header_sz, t)
                if 0 < ri < n_rows - 1 and ri >= n_head and b > 0:
                    row_has_bottom = True
                if ri > n_head and t > 0 and ri != n_rows - 1:
                    row_has_top = True
                if tcpr is not None:
                    shd = tcpr.find(qn("w:shd"))
                    if shd is not None and (shd.get(qn("w:fill")) or "auto").lower() not in ("auto", "ffffff", "none"):
                        shaded = True
                for p in tc.iter(qn("w:p")):
                    sp = p.find(f"{{{W}}}pPr/{{{W}}}spacing")
                    after = sp.get(qn("w:after")) if sp is not None else None
                    before = sp.get(qn("w:before")) if sp is not None else None
                    if (after is not None and int(after) > 0) or (before is not None and int(before) > 0) \
                            or (after is None and normal_after and normal_after > 0):
                        spacing_hits += 1
            if row_has_top or row_has_bottom:
                body_lines += 1
        if nonsolid:
            problems.append(f"线型不是单实线（{'、'.join(sorted(nonsolid))}），双线表不允许")
        if spacing_set:
            problems.append("设置了单元格间距，每条线都会显示成双线")
        if pbdr:
            problems.append("单元格段落带段落边框，会和表格线叠成双线")
        if cond_lines:
            problems.append(f"套用的表格样式带条件格式边框（{'、'.join(cond_lines[:4])}），会自动加线；改用“普通表格”并关闭条件格式")
        if vertical:
            problems.append("有竖线")
        if top_sz < 8:
            problems.append("缺 1 磅以上的顶线")
        if bottom_sz < 8:
            problems.append("缺 1 磅以上的底线")
        if n_rows > 1 and not header_sz:
            problems.append("缺栏目线（只有上下两条线不算三线表）")
        if body_lines > 1:
            problems.append(f"表身有 {body_lines} 处横线")
        if shaded:
            problems.append("有底纹")
        if problems:
            rep.error("TABLE_NOT_3LINE", f"第 {ti} 个表格不是三线表：{'；'.join(problems)}")
        else:
            weights = []
            if not 10 <= top_sz <= 14 or not 10 <= bottom_sz <= 14:
                weights.append(f"顶/底线 {top_sz / 8:g}/{bottom_sz / 8:g} 磅（应 1.5 磅）")
            if n_rows > 1 and not 4 <= header_sz <= 8:
                weights.append(f"栏目线 {header_sz / 8:g} 磅（应 0.75 磅）")
            if n_rows > 1 and header_sz >= min(top_sz, bottom_sz):
                weights.append("栏目线不比顶底线细，层次不清")
            if weights:
                rep.warn("TABLE_RULE_WEIGHT", f"第 {ti} 个表格线宽不合规范：{'；'.join(weights)}")
    if spacing_hits:
        rep.warn("TABLE_SPACING", f"表格内有 {spacing_hits} 个段落带段前/段后间距（含继承自 Normal 样式的），应为 0")


def check_outline(doc, rep):
    styles_el = doc.styles.element
    paras = [p for p in doc.element.body.iter(qn("w:p")) if not in_table(p)]
    fake = []
    prev = None
    skips = []
    for i, p in enumerate(paras):
        t = para_text(p).strip()
        if not t:
            continue
        lvl = para_outline(p, styles_el)
        if lvl is not None:
            if prev is not None and lvl > prev + 1:
                skips.append(t[:20])
            prev = lvl
            continue
        if len(t) <= 30 and not re.search(r"[。；：:，,]$", t) and any(rx.match(t) for rx in HEADING_LIKE_RES):
            if not t.endswith("。"):
                fake.append(t[:20])
    if fake:
        rep.warn("NO_OUTLINE", f"{len(fake)} 个段落像标题但没有大纲级别，导航窗格里看不到：" + "；".join(fake[:6]))
    if skips:
        rep.warn("OUTLINE_SKIP", "标题跳级：" + "；".join(skips[:6]))


def check_captions(doc, rep):
    paras = list(doc.element.body.iter(qn("w:p")))
    _, bib_entries, bib_bms = bib_section(doc)
    bib_ids = {id(p) for p in bib_entries}
    caps = {"图": [], "表": []}
    no_field = 0
    cap_ids = set()
    for p in paras:
        if in_table(p) and p.find(f".//{{{WP}}}inline") is None:
            continue
        t = para_text(p).strip()
        m = CAPTION_RE.match(t)
        if not m or len(t) > 60:
            continue
        kind = m.group(1)
        a, b = m.group(2), m.group(3)
        caps[kind].append((int(a), int(b)) if b else (None, int(a)))
        cap_ids.add(id(p))
        if not any("SEQ" in x.upper() for x in instr_texts(p)):
            no_field += 1
    if no_field:
        rep.warn("CAPTION_NO_FIELD", f"{no_field} 个图题/表题的编号是手打的，应改用题注（SEQ 域），增删图表后才能自动重排")
    known = {"图": set(), "表": set()}
    for kind, lst in caps.items():
        seen = set()
        expect = {}
        bad = []
        for ch, n in lst:
            known[kind].add(f"{ch}.{n}" if ch is not None else str(n))
            if (ch, n) in seen:
                bad.append(f"重复 {kind}{'' if ch is None else str(ch) + '.'}{n}")
            seen.add((ch, n))
            want = expect.get(ch, 0) + 1
            if n != want:
                bad.append(f"{kind}{'' if ch is None else str(ch) + '.'}{n} 前缺 {kind}{'' if ch is None else str(ch) + '.'}{want}")
            expect[ch] = n
        if bad:
            rep.error("CAPTION_SEQ", f"{kind}编号有问题：" + "；".join(bad[:6]))
    manual = 0
    dangling = []
    for p in paras:
        if id(p) in cap_ids or id(p) in bib_ids:
            continue
        t = para_text(p)
        for s_ in REF_ERROR_STRS:
            if s_ in t:
                rep.error("XREF_DANGLING", f"出现“{s_}”：“{t.strip()[:40]}”")
        mentions = MENTION_RE.findall(t)
        if not mentions:
            continue
        refs = sum(1 for x in instr_texts(p)
                   if x.strip().upper().startswith("REF") and not (len(x.split()) > 1 and x.split()[1] in bib_bms))
        manual += max(0, len(mentions) - refs)
        for kind, num in mentions:
            num = num.replace("．", ".")
            if known[kind] and num not in known[kind]:
                dangling.append(f"{kind} {num}")
    if dangling:
        rep.error("XREF_DANGLING", "正文引用了不存在的编号：" + "、".join(sorted(set(dangling))[:8]))
    if manual:
        rep.warn("XREF_MANUAL", f"正文里有 {manual} 处“图 N / 表 N”是手打的，应改用交叉引用（REF 域）")


CITE_RE = re.compile(r"\[(\d+(?:\s*[,，\-–~]\s*\d+)*)\]")
BIB_HEAD_RE = re.compile(r"^参\s*考\s*文\s*献$|^references$|^bibliography$", re.I)


def _expand(group: str):
    out = []
    for part in re.split(r"\s*[,，]\s*", group.strip()):
        m = re.match(r"^(\d+)\s*[\-–~]\s*(\d+)$", part)
        if m:
            a, b = int(m.group(1)), int(m.group(2))
            out.extend(range(a, b + 1) if b >= a else [a, b])
        elif part.isdigit():
            out.append(int(part))
    return out


def _effective_ind(p_el, styles_el):
    """段落实际缩进：样式链上的 w:ind 被段落自己的 w:ind 按属性覆盖。"""
    attrs = {}
    chain = []
    ppr = p_el.find(qn("w:pPr"))
    ps = ppr.find(qn("w:pStyle")) if ppr is not None else None
    sid = ps.get(qn("w:val")) if ps is not None else None
    if sid is None:
        for x in styles_el.findall(qn("w:style")):
            if x.get(qn("w:type")) == "paragraph" and x.get(qn("w:default")) in ("1", "true"):
                sid = x.get(qn("w:styleId"))
    depth = 0
    while sid and depth < 20:
        node = next((x for x in styles_el.findall(qn("w:style")) if x.get(qn("w:styleId")) == sid), None)
        if node is None:
            break
        chain.append(node)
        b = node.find(qn("w:basedOn"))
        sid = b.get(qn("w:val")) if b is not None else None
        depth += 1
    for node in reversed(chain):
        ind = node.find(f"{{{W}}}pPr/{{{W}}}ind")
        if ind is not None:
            attrs.update({k.split("}")[1]: v for k, v in ind.attrib.items()})
    if ppr is not None:
        ind = ppr.find(qn("w:ind"))
        if ind is not None:
            own = {k.split("}")[1]: v for k, v in ind.attrib.items()}
            # 字符单位与绝对单位互斥：段落设了一种，就覆盖样式里同类的另一种
            if "hanging" in own or "hangingChars" in own:
                attrs.pop("firstLine", None)
                attrs.pop("firstLineChars", None)
            if "firstLine" in own or "firstLineChars" in own:
                attrs.pop("hanging", None)
                attrs.pop("hangingChars", None)
            attrs.update(own)
    return attrs


def bib_section(doc):
    """返回 (参考文献标题之前的正文段落, 参考文献条目段落, 条目里的书签名集合)。"""
    paras = [p for p in doc.element.body.iter(qn("w:p")) if not in_table(p)]
    styles_el = doc.styles.element
    head = None
    for i, p in enumerate(paras):
        if BIB_HEAD_RE.match(para_text(p).strip().replace(" ", "")):
            head = i
    if head is None:
        return paras, [], set()
    entries = []
    for p in paras[head + 1:]:
        t = para_text(p).strip()
        if para_outline(p, styles_el) is not None and t:
            break
        if t:
            entries.append(p)
    bms = {b.get(qn("w:name")) for p in entries for b in p.iter(qn("w:bookmarkStart"))}
    return paras[:head], entries, bms


GBT_TYPES = {"M", "J", "N", "C", "D", "R", "S", "P", "EB", "A", "CM", "DS", "PP", "G", "CP", "DB", "Z"}
GBT_CARRIERS = {"OL", "CD", "DK", "MT", "MM"}
TYPE_RE = re.compile(r"\[([A-Z]{1,2})(?:/([A-Z]{2}))?\](?://)?")
DATE_BRACKET_RE = re.compile(r"\[\d{4}-\d{2}-\d{2}\]")
STD_NO_RE = re.compile(r"^(GB|GB/T|GB/Z|ISO|IEC|IEEE|JB|JB/T|HJ|YY|DL/T|QB|SJ|GJB)[\s/]")
CAPS_SURNAME_RE = re.compile(r"(?:^|,\s*)([A-Z]{2,}(?:-[A-Z]{2,})?)\s[A-Z](?:\s[A-Z])*(?=[,.])")


def check_bib_format(entries, rep):
    """GB/T 7714—2025 著录检查（只查能用规则判断的部分）。"""
    no_type, bad_code, print_date, no_date, fullwidth, caps = [], [], [], [], [], []
    for k, p in enumerate(entries, 1):
        t = re.sub(r"^\[\d+\]\s*", "", para_text(p).strip())
        if re.search(r"[，：；．]", t) or re.search(r"[一-鿿\]）)]。$", t):
            fullwidth.append(k)
        if not re.search(r"[一-鿿]", t.split("[")[0]):
            m = CAPS_SURNAME_RE.search(t.split("[")[0])
            if m:
                caps.append(f"{k}:{m.group(1)}")
        codes = TYPE_RE.findall(t)
        is_std = bool(STD_NO_RE.match(t))
        if not codes:
            if not is_std:
                no_type.append(k)
            continue
        typ, carrier = codes[0]
        if typ not in GBT_TYPES or (carrier and carrier not in GBT_CARRIERS):
            bad_code.append(f"{k}:[{typ}{'/' + carrier if carrier else ''}]")
        online = carrier == "OL" or "http" in t.lower()
        has_date = bool(DATE_BRACKET_RE.search(t))
        if has_date and not online:
            print_date.append(k)
        if typ in ("EB", "DS", "PP") and not has_date:
            no_date.append(k)
    if no_type:
        rep.warn("BIB_FORMAT", f"第 {no_type[:10]} 条缺文献类型标识（2025 版为必备项，只有标准可省）")
    if bad_code:
        rep.warn("BIB_FORMAT", f"类型/载体代码不在 GB/T 7714—2025 代码表里：{bad_code[:8]}")
    if print_date:
        rep.warn("BIB_FORMAT", f"第 {print_date[:10]} 条是非网络文献却写了引用日期，2025 版只有网络资源著录引用日期")
    if no_date:
        rep.warn("BIB_FORMAT", f"第 {no_date[:10]} 条是网页/数据集/预印本，缺 [引用日期]")
    if fullwidth:
        rep.warn("BIB_FORMAT", f"第 {fullwidth[:10]} 条用了全角著录符号（，：；．或句号结尾），著录用符号应为半角")
    if caps:
        rep.warn("BIB_FORMAT", f"西文姓全大写是 2015 写法，2025 版只首字母大写（Einstein A）：{caps[:6]}；若是汉语拼音姓名可保留全大写")


def check_citations(doc, rep):
    styles_el = doc.styles.element
    body_paras, entries, bib_bms = bib_section(doc)
    if not entries:
        return
    nums, no_field, bad_indent, no_hang = [], 0, [], []
    for k, p in enumerate(entries, 1):
        t = para_text(p).strip()
        numbered = p.find(f"{{{W}}}pPr/{{{W}}}numPr") is not None
        m = re.match(r"^\[(\d+)\]", t)
        nums.append(int(m.group(1)) if m else (k if numbered else None))
        if not numbered and not any("SEQ" in x.upper() for x in instr_texts(p)):
            no_field += 1
        ind = _effective_ind(p, styles_el)
        first = int(ind.get("firstLine", 0) or 0) + int(ind.get("firstLineChars", 0) or 0)
        hang = int(ind.get("hanging", 0) or 0) + int(ind.get("hangingChars", 0) or 0)
        if first > 0:
            bad_indent.append(k)
        elif hang <= 0 and not numbered:
            no_hang.append(k)
    if bad_indent:
        rep.error("BIB_INDENT", f"参考文献第 {', '.join(map(str, bad_indent[:8]))} 条用了首行缩进，应为悬挂缩进（序号顶格，续行与文字对齐）")
    if no_hang:
        rep.warn("BIB_INDENT", f"参考文献第 {', '.join(map(str, no_hang[:8]))} 条没有悬挂缩进，续行会顶到序号下面")
    if no_field:
        rep.warn("BIB_NO_FIELD", f"{no_field} 条参考文献的序号是手打的，增删文献后不会自动重排")
    valid = [n for n in nums if n is not None]
    if valid != list(range(1, len(valid) + 1)):
        rep.error("BIB_SEQ", f"参考文献序号不连续或重复：{valid[:12]}")
    total = len(entries)
    check_bib_format(entries, rep)

    cited_first = []
    manual = 0
    dangling = set()
    for p in body_paras:
        if is_code_like(p):
            continue  # 代码里的 a[0] 不是引用
        t = para_text(p)
        groups = CITE_RE.findall(t)
        if not groups:
            continue
        refs_to_bib = sum(1 for x in instr_texts(p)
                          if x.strip().upper().startswith("REF") and len(x.split()) > 1 and x.split()[1] in bib_bms)
        if refs_to_bib == 0:
            manual += len(groups)
        for g in groups:
            for n in _expand(g):
                if n < 1 or n > total:
                    dangling.add(n)
                elif n not in cited_first:
                    cited_first.append(n)
    if dangling:
        rep.error("CITE_DANGLING", f"正文引用了不存在的文献序号：{sorted(dangling)[:10]}（共 {total} 条）")
    if manual:
        rep.warn("CITE_MANUAL", f"正文有 {manual} 处文献引用是手打的，应改为指向参考文献序号的交叉引用（REF 域）")
    uncited = [n for n in range(1, total + 1) if n not in cited_first]
    if uncited:
        rep.warn("BIB_UNCITED", f"参考文献第 {uncited[:10]} 条正文没有引用")
    if cited_first and cited_first != sorted(cited_first):
        rep.warn("CITE_ORDER", f"首次引用顺序为 {cited_first[:12]}，顺序编码制应按首次引用先后编号")


# ---- 检查 ----

def check_fonts(doc, rep, min_pt, extra_parts):
    info, default_para, default_sz = build_style_info(doc.styles.element)
    min_half = int(round(min_pt * 2))
    hits = []
    body_index = {id(p): i + 1 for i, p in enumerate(doc.element.body.iter(qn("w:p")))}

    def scan(root, where_prefix):
        for p in root.iter(qn("w:p")):
            for r in p.iter(qn("w:r")):
                text = run_text(r).strip()
                if not text:
                    continue
                sz = run_size_halfpts(r, p, info, default_para, default_sz)
                if sz < min_half:
                    where = where_prefix or location_of(p, body_index)
                    hits.append((sz / 2, where, text[:30]))

    scan(doc.element.body, None)
    for name, root in extra_parts:
        scan(root, name)
    if hits:
        by_size = Counter(h[0] for h in hits)
        summary = "，".join(f"{s:g} pt×{n}" for s, n in sorted(by_size.items()))
        rep.error("SMALL_FONT", f"发现 {len(hits)} 处字号小于 {min_pt:g} pt（{summary}）")
        for s, where, text in hits[:12]:
            rep.error("SMALL_FONT", f"  {s:g} pt @ {where}: “{text}”")
        if len(hits) > 12:
            rep.error("SMALL_FONT", f"  …另有 {len(hits) - 12} 处，见上面的汇总")


def check_images(doc, rep, text_w_cm, max_w, max_h, min_dpi):
    body = doc.element.body
    paras = list(body.iter(qn("w:p")))
    p_index = {id(p): i for i, p in enumerate(paras)}
    n = 0
    for tag in (f"{{{WP}}}inline", f"{{{WP}}}anchor"):
        for dr in body.iter(tag):
            n += 1
            ext = dr.find(f"{{{WP}}}extent")
            if ext is None:
                continue
            w = int(ext.get("cx")) / EMU_PER_CM
            h = int(ext.get("cy")) / EMU_PER_CM
            label = f"图片 {n}（{w:.1f}×{h:.1f} cm）"
            if tag.endswith("anchor"):
                rep.warn("IMG_FLOATING", f"{label} 是浮动图片，建议改为嵌入型")
            if w > text_w_cm + 0.05:
                rep.error("IMG_TOO_WIDE", f"{label} 宽度超过正文区 {text_w_cm:.1f} cm")
            elif w > max_w + 0.05:
                rep.warn("IMG_OVER_LIMIT", f"{label} 宽度超过建议上限 {max_w:g} cm")
            if h > max_h + 0.05:
                rep.error("IMG_TOO_TALL", f"{label} 高度超过 {max_h:g} cm，可能与图题分页")
            if w < 4:
                rep.warn("IMG_TINY", f"{label} 宽度不足 4 cm，确认是否为有意的小图")
            px = None
            blip = dr.find(f".//{{{A}}}blip")
            if blip is not None and Image is not None:
                rid = blip.get(f"{{{R}}}embed")
                part = doc.part.related_parts.get(rid) if rid else None
                if part is not None:
                    try:
                        with Image.open(io.BytesIO(part.blob)) as im:
                            px = im.size
                    except Exception:
                        px = None
            if px and px[0] and px[1] and w > 0 and h > 0:
                shown, real = w / h, px[0] / px[1]
                if abs(shown / real - 1) > 0.03:
                    rep.error("IMG_DISTORTED", f"{label} 被拉伸：显示比 {shown:.2f}，像素比 {real:.2f}（{px[0]}×{px[1]} px）")
                dpi = px[0] / (w / 2.54)
                if dpi < min_dpi * 0.66:
                    rep.error("IMG_LOW_DPI", f"{label} 有效分辨率仅 {dpi:.0f} dpi（{px[0]} px 宽），会明显发糊，请缩小或重新出图")
                elif dpi < min_dpi:
                    rep.warn("IMG_LOW_DPI", f"{label} 有效分辨率 {dpi:.0f} dpi，低于 {min_dpi:g}，建议缩小到 {px[0] / min_dpi * 2.54:.1f} cm 以内或重新出图")
            # 图题
            p = dr
            while p is not None and p.tag != qn("w:p"):
                p = p.getparent()
            if p is not None and id(p) in p_index:
                i = p_index[id(p)]
                found = False
                for j in range(i, min(i + 3, len(paras))):
                    t = para_text(paras[j]).strip()
                    if re.match(r"^(图|Fig\.?|Figure)\s*\d", t) or re.match(r"^图\s*\d", t):
                        found = True
                        break
                if not found:
                    rep.warn("IMG_NO_CAPTION", f"{label} 下方两段内没有“图 N …”图题")
    for _ in body.iter(f"{{{V}}}imagedata"):
        rep.warn("IMG_FLOATING", "存在 VML 旧式图片，无法自动检查尺寸，请在 Word 里确认")


def check_text(doc, rep):
    info, _, _ = build_style_info(doc.styles.element)
    body_paras = list(doc.element.body.iter(qn("w:p")))
    texts = [para_text(p) for p in body_paras]
    full = "\n".join(texts)

    code_like = [is_code_like(p) for p in body_paras]

    # 禁用短语
    def collect(phrases, case_sensitive=False):
        cnt, ex = Counter(), {}
        for i, t in enumerate(texts):
            hay = t if case_sensitive else t.lower()
            for ph in phrases:
                needle = ph if case_sensitive else ph.lower()
                if needle in hay:
                    cnt[ph] += 1
                    ex.setdefault(ph, (i + 1, t.strip()[:40]))
        return cnt, ex

    cnt, ex = collect(FORBIDDEN_PHRASES)
    for ph, c in cnt.most_common():
        i, e = ex[ph]
        rep.error("FORBIDDEN", f"“{ph}” 出现 {c} 次，如第 {i} 段：“{e}”")
    cnt, ex = collect(FORBIDDEN_CASE_SENSITIVE, case_sensitive=True)
    for ph, c in cnt.most_common():
        i, e = ex[ph]
        if code_like[i - 1]:
            rep.warn("FORBIDDEN", f"代码段里出现 “{ph}” {c} 次，如第 {i} 段：“{e}”，确认不是遗留占位")
        else:
            rep.error("FORBIDDEN", f"“{ph}” 出现 {c} 次，如第 {i} 段：“{e}”")
    cnt, ex = collect(FORBIDDEN_SOFT)
    for ph, c in cnt.most_common():
        i, e = ex[ph]
        rep.warn("FORBIDDEN", f"“{ph}” 出现 {c} 次（需人工判断），如第 {i} 段：“{e}”")
    for i, t in enumerate(texts):
        if WORDCOUNT_RE.search(t):
            rep.error("FORBIDDEN", f"第 {i + 1} 段复述字数要求：“{t.strip()[:40]}”")

    # 让步式自贬
    for i, t in enumerate(texts):
        for pat in HEDGE_PATTERNS:
            if re.search(pat, t):
                rep.warn("HEDGE", f"第 {i + 1} 段疑似让步式自贬（{pat}）：“{t.strip()[:40]}”")
                break

    # Markdown 残留（代码段跳过）
    md = Counter()
    md_ex = {}
    for i, t in enumerate(texts):
        if code_like[i]:
            continue
        for rx, name in MD_RE:
            if rx.search(t):
                md[name] += 1
                md_ex.setdefault(name, (i + 1, t.strip()[:40]))
        if BACKTICK_RE.search(t):
            md["反引号"] += 1
            md_ex.setdefault("反引号", (i + 1, t.strip()[:40]))
    for name, c in md.items():
        i, e = md_ex[name]
        level = rep.warn if name in ("反引号", "列表符号") else rep.error
        level("MARKDOWN", f"{name} 残留 {c} 处，如第 {i} 段：“{e}”")

    # AI 味词
    for w, thr in AI_WORDS.items():
        c = full.count(w)
        if c >= thr:
            rep.warn("AI_WORD", f"“{w}” 出现 {c} 次（阈值 {thr}）")
    sc = sum(full.count(w) for w in SUMMARY_WORDS)
    if sc >= 2:
        rep.warn("AI_WORD", f"总结套话（综上所述/总而言之/总的来说/总之）共 {sc} 次，一篇最多一次")
    trip = min(full.count("首先"), full.count("其次"), full.count("最后"))
    if trip >= 3:
        rep.warn("AI_WORD", f"“首先…其次…最后”套路出现约 {trip} 组")

    # 缺点分布
    section = "（开头）"
    sec_hits = defaultdict(list)
    total = 0
    for i, (p, t) in enumerate(zip(body_paras, texts)):
        if is_heading(p, info) and t.strip():
            section = t.strip()[:20]
            continue
        if any(w in t for w in FLAW_WORDS):
            sec_hits[section].append(i + 1)
            total += 1
    if len(sec_hits) >= 3 or total >= 6:
        where = "；".join(f"{s}（{len(v)} 段）" for s, v in sec_hits.items())
        rep.warn("FLAW_SPREAD", f"缺点/局限类表述出现在 {len(sec_hits)} 个章节共 {total} 段，建议集中到一处：{where}")

    # 自称
    formal = sum(full.count(w) for w in FORMAL_SELF)
    we = len(WE_RE.findall(full))
    me = len(CASUAL_SELF_RE.findall(full))
    parts = [f"笔者/本文 {formal}", f"我们 {we}", f"我 {me}"]
    if formal > 0 and me >= 3:
        rep.warn("PRONOUN_MIX", f"自称混用：{'，'.join(parts)}。论文体统一用“本文/笔者”，心得/汇报体统一用“我/我们”")
    if full.count("笔者") > 0 and full.count("本文") > 0 and me == 0:
        pass  # 笔者+本文 可以共存
    nin = full.count("您")
    if nin > 0:
        rep.warn("READER_ADDRESS", f"正文出现“您” {nin} 次，报告正文一般不对读者说话")

    # 空段
    run_len = 0
    start = 0
    for i, (p, t) in enumerate(zip(body_paras, texts)):
        has_drawing = p.find(f".//{{{WP}}}inline") is not None or p.find(f".//{{{WP}}}anchor") is not None
        has_break = p.find(f".//{{{W}}}br[@{{{W}}}type='page']") is not None
        in_table = False
        anc = p.getparent()
        while anc is not None and anc.tag != qn("w:body"):
            if anc.tag == qn("w:tbl"):
                in_table = True
                break
            anc = anc.getparent()
        if in_table:
            continue
        if not t.strip() and not has_drawing and not has_break:
            if run_len == 0:
                start = i + 1
            run_len += 1
        else:
            if run_len >= 2:
                rep.warn("EMPTY_PARAS", f"第 {start} 段起连续 {run_len} 个空段")
            run_len = 0
    if run_len >= 2:
        rep.warn("EMPTY_PARAS", f"第 {start} 段起连续 {run_len} 个空段（文末）")

    # 符号
    em = len(EMOJI_RE.findall(full))
    if em:
        rep.error("EMOJI", f"出现表情符号 {em} 处")
    ex = full.count("！") + full.count("!")
    if ex >= 3:
        rep.warn("EXCLAIM", f"感叹号 {ex} 处，学术文体基本不用")
    dash = full.count("——")
    if dash >= 4:
        rep.warn("DASH", f"破折号 {dash} 处，建议不超过 3 处")


def check_layout(doc, rep):
    for i, s in enumerate(doc.sections, 1):
        pw, ph = s.page_width / EMU_PER_CM, s.page_height / EMU_PER_CM
        a4 = (abs(pw - 21.0) < 0.1 and abs(ph - 29.7) < 0.1) or (abs(pw - 29.7) < 0.1 and abs(ph - 21.0) < 0.1)
        if not a4:
            rep.warn("PAGE_SETUP", f"第 {i} 节纸张 {pw:.1f}×{ph:.1f} cm，不是 A4")
        m = [s.top_margin, s.bottom_margin, s.left_margin, s.right_margin]
        m_cm = [x / EMU_PER_CM if x is not None else 0 for x in m]
        if any(abs(x - 2.5) > 0.06 for x in m_cm):
            rep.warn("PAGE_SETUP", f"第 {i} 节页边距 上{m_cm[0]:.2f} 下{m_cm[1]:.2f} 左{m_cm[2]:.2f} 右{m_cm[3]:.2f} cm，默认应四边 2.5 cm（有模板则以模板为准）")
        if s.gutter and s.gutter > 0:
            rep.warn("PAGE_SETUP", f"第 {i} 节设置了装订线 {s.gutter / EMU_PER_CM:.2f} cm")


def extra_xml_parts(doc):
    out = []
    for rel in doc.part.rels.values():
        if rel.is_external:
            continue
        rt = rel.reltype
        name = None
        if rt.endswith("/header"):
            name = "页眉"
        elif rt.endswith("/footer"):
            name = "页脚"
        elif rt.endswith("/footnotes"):
            name = "脚注"
        elif rt.endswith("/endnotes"):
            name = "尾注"
        if name:
            part = rel.target_part
            el = getattr(part, "element", None)
            if el is None:
                try:
                    el = etree.fromstring(part.blob)
                except Exception:
                    continue
            out.append((name, el))
    return out


def main(argv=None):
    ap = argparse.ArgumentParser(description="检查课程报告 .docx 的字号、图片、语言、版面问题")
    ap.add_argument("docx")
    ap.add_argument("--min-font", type=float, default=10.0)
    ap.add_argument("--max-fig-width", type=float, default=15.0)
    ap.add_argument("--max-fig-height", type=float, default=18.0)
    ap.add_argument("--min-dpi", type=float, default=150.0)
    ap.add_argument("--no-layout", action="store_true", help="跳过纸张/页边距检查")
    ap.add_argument("--template", action="store_true", help="有模板：等同 --no-layout")
    args = ap.parse_args(argv)

    if not os.path.exists(args.docx):
        print(f"文件不存在：{args.docx}", file=sys.stderr)
        return 2
    doc = Document(args.docx)
    rep = Report()
    s0 = doc.sections[0]
    text_w = (s0.page_width - s0.left_margin - s0.right_margin) / EMU_PER_CM

    check_fonts(doc, rep, args.min_font, extra_xml_parts(doc))
    check_images(doc, rep, text_w, args.max_fig_width, args.max_fig_height, args.min_dpi)
    check_text(doc, rep)
    check_tables(doc, rep)
    check_outline(doc, rep)
    check_captions(doc, rep)
    check_citations(doc, rep)
    if not (args.no_layout or args.template):
        check_layout(doc, rep)

    print(f"文件：{args.docx}")
    print(f"正文区宽 {text_w:.1f} cm；字号下限 {args.min_font:g} pt；图宽上限 {args.max_fig_width:g} cm\n")
    for level in ("ERROR", "WARN"):
        items = [i for i in rep.items if i[0] == level]
        if not items:
            continue
        print(f"== {level} ({len(items)}) ==")
        for _, code, msg in items:
            print(f"[{level}][{code}] {msg}")
        print()
    ne, nw = len(rep.errors), len(rep.warns)
    if ne == 0 and nw == 0:
        print("未发现问题。")
    else:
        print(f"合计：ERROR {ne}，WARN {nw}")
    return 1 if ne else 0


if __name__ == "__main__":
    sys.exit(main())
