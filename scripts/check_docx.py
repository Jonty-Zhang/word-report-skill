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
