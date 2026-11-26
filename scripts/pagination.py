#!/usr/bin/env python3
"""减少“图表在本页放不下、整块挪到下一页、本页留下大片空白”。

Word 里，插图和不拆分的表格在当前页放不下时，会整块移到下一页，当前页剩下的位置就空着。
本模块在保存前粗略模拟分页（不需要 Word），发现这种空白时，按下面的顺序处理：

1. 插图缩小一点就能放下（不小于原尺寸的 85%）：缩小插图。
2. 把图表后面的正文段落移到图表前面填空。效果类似 LaTeX 的浮动体：
   - 只移普通正文段落；
   - 遇到标题、公式、其他图表、分页符就停；
   - 图表所在的连续段落里有标题时，不移动。
3. 插图缩小到不小于原尺寸的 75% 能放下：缩小插图。
4. 数据行不少于 6 行的表格允许跨页，有两种方式（split_mode）：
   - "continued"（论文，默认）：在估算的页尾处拆成两段，后一段上方加“表×（续）”，
     并重复表头（GB/T 7713.2—2022 5.5.5）；拆分后仍超过一页的表格，继续往下拆。
   - "flow"（报告）：不拆表，只解除前几行之后的“与下段同页”，由 Word 自己断页，表头行自动重复，
     不加“续表”字样。断页位置由 Word 决定，不受估算误差影响。

估算依据是字号、行距、缩进、图片尺寸和表格行数，和 Word 的实际排版有误差，所以每一步都留了余量。
reduce_float_gaps 返回改动说明。交付时把改动告诉用户，并提醒在 Word 里看一眼分页。

用法::

    import pagination
    log = pagination.reduce_float_gaps(doc)   # 在所有内容写完、保存之前调用
"""
from __future__ import annotations

import copy
import math
import re

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
EMU_PER_PT = 12700

GAP_RATIO = 0.18          # 空白超过版心高度的 18% 才处理（A4、2.5 cm 页边距时约 4.4 cm）
SLACK_LINES = 1.5         # 估算误差余量，按正文行高计
SHRINK_SOFT = 0.85        # 先试：缩小后不小于原尺寸的 85%
SHRINK_HARD = 0.75        # 再试：不小于 75%
MIN_SPLIT_ROWS = 6        # 数据行不少于 6 行的表格才拆
MIN_PART_ROWS = 3         # 拆开后每段至少 3 行数据
MAX_MOVE_PARAS = 8        # 最多移动 8 段正文
MAX_ROUNDS = 60

_CJK = re.compile(r"[⺀-鿿　-〿＀-￯“”‘’]")
_TABLE_LABEL = re.compile(r"^\s*(表\s*[A-Z]?[.．]?\d+(?:[.\-－]\d+)?)")
_FIG_LABEL = re.compile(r"^\s*(图\s*[A-Z]?[.．]?\d+(?:[.\-－]\d+)?)")


def _on(el):
    """布尔开关元素（keepNext 等）：None=未设置，否则 True/False。"""
    if el is None:
        return None
    return el.get(qn("w:val")) not in ("0", "false", "off")


def _int(el, attr):
    if el is None:
        return None
    v = el.get(qn(attr))
    try:
        return int(float(v)) if v is not None else None
    except ValueError:
        return None


class _Styles:
    """沿 段落属性 → 段落样式（basedOn 链）→ docDefaults 取值。"""

    def __init__(self, doc):
        root = doc.styles.element
        self.by_id = {s.get(qn("w:styleId")): s for s in root.findall(qn("w:style"))}
        self.default_p = None
        for s in root.findall(qn("w:style")):
            if s.get(qn("w:type")) == "paragraph" and s.get(qn("w:default")) in ("1", "true"):
                self.default_p = s.get(qn("w:styleId"))
        dd = root.find(qn("w:docDefaults"))
        self.d_ppr = dd.find(f"{qn('w:pPrDefault')}/{qn('w:pPr')}") if dd is not None else None
        self.d_rpr = dd.find(f"{qn('w:rPrDefault')}/{qn('w:rPr')}") if dd is not None else None

    def _chain(self, sid):
        seen = set()
        while sid and sid in self.by_id and sid not in seen:
            seen.add(sid)
            s = self.by_id[sid]
            yield s
            b = s.find(qn("w:basedOn"))
            sid = b.get(qn("w:val")) if b is not None else None

    def _pstyle(self, p):
        ppr = p.find(qn("w:pPr"))
        ps = ppr.find(qn("w:pStyle")) if ppr is not None else None
        return ps.get(qn("w:val")) if ps is not None else self.default_p

    def ppr(self, p, getter):
        ppr = p.find(qn("w:pPr"))
        if ppr is not None:
            v = getter(ppr)
            if v is not None:
                return v
        for s in self._chain(self._pstyle(p)):
            sp = s.find(qn("w:pPr"))
            if sp is not None:
                v = getter(sp)
                if v is not None:
                    return v
        return getter(self.d_ppr) if self.d_ppr is not None else None

    def size(self, p, run=None):
        """字号（pt）：run → run 样式 → 段落样式 → docDefaults。"""
        if run is not None:
            rpr = run.find(qn("w:rPr"))
            v = _int(rpr.find(qn("w:sz")), "w:val") if rpr is not None else None
            if v:
                return v / 2
            rs = rpr.find(qn("w:rStyle")) if rpr is not None else None
            if rs is not None:
                for s in self._chain(rs.get(qn("w:val"))):
                    srpr = s.find(qn("w:rPr"))
                    v = _int(srpr.find(qn("w:sz")), "w:val") if srpr is not None else None
                    if v:
                        return v / 2
        for s in self._chain(self._pstyle(p)):
            srpr = s.find(qn("w:rPr"))
            v = _int(srpr.find(qn("w:sz")), "w:val") if srpr is not None else None
            if v:
                return v / 2
        v = _int(self.d_rpr.find(qn("w:sz")), "w:val") if self.d_rpr is not None else None
        return v / 2 if v else 10.0


class _Block:
    __slots__ = ("el", "kind", "h", "min_h", "lh", "keep", "page_before", "break_after", "img", "heading",
                 "math", "text", "rows", "kept_table", "empty")

    def __init__(self, el, kind):
        self.el, self.kind = el, kind
        self.h = self.min_h = self.lh = 0.0
        self.keep = self.page_before = self.break_after = self.img = self.heading = self.math = False
        self.kept_table = self.empty = False
        self.text = ""
        self.rows = []


class _Layout:
    def __init__(self, doc):
        self.doc = doc
        self.st = _Styles(doc)
        sect = doc.sections[0]
        self.page_h = (sect.page_height - sect.top_margin - sect.bottom_margin) / EMU_PER_PT
        self.text_w = (sect.page_width - sect.left_margin - sect.right_margin) / EMU_PER_PT
        self.body_lh = self._normal_line_height()

    def _normal_line_height(self):
        p = OxmlElement("w:p")
        size = self.st.size(p)
        line = self.st.ppr(p, lambda x: _int(x.find(qn("w:spacing")), "w:line"))
        rule = self.st.ppr(p, lambda x: (x.find(qn("w:spacing")).get(qn("w:lineRule"))
                                         if x.find(qn("w:spacing")) is not None else None))
        return self._line_h(size, True, line, rule)

    @staticmethod
    def _line_h(size, cjk, line, rule):
        base = size * (1.3 if cjk else 1.17)
        if rule == "exact" and line:
            return line / 20
        if rule == "atLeast" and line:
            return max(line / 20, base)
        return base * ((line / 240) if line else 1.0)

    # ---------- 段落 ----------
    def para(self, p, width=None):
        b = _Block(p, "p")
        width = width or self.text_w
        st = self.st
        texts = "".join(t.text or "" for t in p.iter(qn("w:t")))
        tabs = len(p.findall(f".//{qn('w:tab')}")) - len(p.findall(f"{qn('w:pPr')}/{qn('w:tabs')}/{qn('w:tab')}"))
        mtext = "".join(t.text or "" for t in p.iter(f"{{{M_NS}}}t"))
        b.text = texts.strip()
        b.math = bool(mtext) or p.find(f".//{{{M_NS}}}oMath") is not None
        runs = [r for r in p.iter(qn("w:r")) if r.find(qn("w:t")) is not None]
        size = max([st.size(p, r) for r in runs] or [st.size(p)])
        cjk = bool(_CJK.search(texts)) or not texts.strip()
        sp = lambda a: (lambda x: _int(x.find(qn("w:spacing")), a))  # noqa: E731
        line = st.ppr(p, sp("w:line"))
        rule = st.ppr(p, lambda x: (x.find(qn("w:spacing")).get(qn("w:lineRule"))
                                    if x.find(qn("w:spacing")) is not None and
                                    x.find(qn("w:spacing")).get(qn("w:lineRule")) else None))
        before = (st.ppr(p, sp("w:before")) or 0) / 20
        after = (st.ppr(p, sp("w:after")) or 0) / 20
        lh = self._line_h(size, cjk, line, rule)
        if b.math and (p.find(f".//{{{M_NS}}}f") is not None or p.find(f".//{{{M_NS}}}nary") is not None):
            lh *= 1.8
        ind = lambda a: (lambda x: _int(x.find(qn("w:ind")), a))  # noqa: E731
        first_chars = st.ppr(p, ind("w:firstLineChars"))
        first = (first_chars / 100 * size) if first_chars else (st.ppr(p, ind("w:firstLine")) or 0) / 20
        left = (st.ppr(p, ind("w:left")) or st.ppr(p, ind("w:start")) or 0) / 20
        right = (st.ppr(p, ind("w:right")) or st.ppr(p, ind("w:end")) or 0) / 20
        hanging = (st.ppr(p, ind("w:hanging")) or 0) / 20
        avail = max(width - left - right, 40)
        em = sum(1.0 if _CJK.match(c) else 0.55 for c in texts) + 2.0 * max(tabs, 0) + 0.6 * len(mtext)
        n = max(1, math.ceil((em * size + first - hanging) / avail)) if (texts or mtext) else 1
        img = 0.0
        for ext in p.iter(f"{{{WP_NS}}}extent"):
            if ext.getparent().tag == f"{{{WP_NS}}}inline":
                img = max(img, int(ext.get("cy", "0")) / EMU_PER_PT)
        b.img = img > 0
        if b.img:
            mult = (line / 240) if (line and rule in (None, "auto")) else 1.0
            body = img * mult + size * 0.3 + (n - 1) * lh if texts.strip() else img * mult + size * 0.3
            if rule == "exact" and line:
                body = line / 20
        else:
            body = n * lh
        b.lh = lh
        b.h = before + body + after
        keep_lines = st.ppr(p, lambda x: _on(x.find(qn("w:keepLines"))))
        b.min_h = before + 2 * lh if (n >= 4 and not keep_lines and not b.img) else b.h
        b.keep = bool(st.ppr(p, lambda x: _on(x.find(qn("w:keepNext")))))
        b.page_before = bool(st.ppr(p, lambda x: _on(x.find(qn("w:pageBreakBefore")))))
        b.break_after = any(br.get(qn("w:type")) == "page" for br in p.iter(qn("w:br"))) or \
            p.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is not None
        lvl = st.ppr(p, lambda x: _int(x.find(qn("w:outlineLvl")), "w:val"))
        b.heading = lvl is not None and lvl < 9
        b.empty = not texts.strip() and not b.img and not b.math
        return b

    # ---------- 表格 ----------
    def table(self, tbl):
        b = _Block(tbl, "tbl")
        tblpr = tbl.find(qn("w:tblPr"))
        mar = tblpr.find(qn("w:tblCellMar")) if tblpr is not None else None
        mt = (_int(mar.find(qn("w:top")), "w:w") or 0) / 20 if mar is not None else 0
        mb = (_int(mar.find(qn("w:bottom")), "w:w") or 0) / 20 if mar is not None else 0
        ml = (_int(mar.find(qn("w:left")), "w:w") or 108) / 20 if mar is not None else 5.4
        mr = (_int(mar.find(qn("w:right")), "w:w") or 108) / 20 if mar is not None else 5.4
        grid = [(_int(g, "w:w") or 0) / 20 for g in tbl.findall(f"{qn('w:tblGrid')}/{qn('w:gridCol')}")]
        rows = []
        for tr in tbl.findall(qn("w:tr")):
            col, rh, keep_all, any_p = 0, 0.0, True, False
            for tc in tr.findall(qn("w:tc")):
                tcpr = tc.find(qn("w:tcPr"))
                span = _int(tcpr.find(qn("w:gridSpan")), "w:val") if tcpr is not None else None
                span = span or 1
                w = sum(grid[col:col + span]) if grid[col:col + span] else (_int(
                    tcpr.find(qn("w:tcW")) if tcpr is not None else None, "w:w") or 1440) / 20
                col += span
                ch = 0.0
                for p in tc.findall(qn("w:p")):
                    pb = self.para(p, width=max(w - ml - mr, 20))
                    ch += pb.h
                    any_p = True
                    keep_all = keep_all and pb.keep
                rh = max(rh, ch)
            trpr = tr.find(qn("w:trPr"))
            trh = trpr.find(qn("w:trHeight")) if trpr is not None else None
            if trh is not None and _int(trh, "w:val"):
                v = _int(trh, "w:val") / 20
                rh = v if trh.get(qn("w:hRule")) == "exact" else max(rh, v)
            header = trpr is not None and _on(trpr.find(qn("w:tblHeader"))) is True
            rows.append({"el": tr, "h": rh + mt + mb, "keep": keep_all and any_p, "header": header})
        b.rows = rows
        b.h = sum(r["h"] for r in rows) + 2
        b.kept_table = len(rows) > 1 and all(r["keep"] for r in rows[:-1])
        b.keep = bool(rows) and rows[-1]["keep"]
        b.min_h = b.h if b.kept_table else (rows[0]["h"] + (rows[1]["h"] if len(rows) > 1 else 0))
        return b

    def blocks(self):
        out = []
        for el in self.doc.element.body:
            if el.tag == qn("w:p"):
                out.append(self.para(el))
            elif el.tag == qn("w:tbl"):
                out.append(self.table(el))
        return out

    # ---------- 模拟分页 ----------
    def simulate(self, blocks):
        """返回 [(i, j, gap, page)]：blocks[i..j] 这一串因放不下被整体挪到下一页，留下 gap pt 空白。"""
        P = self.page_h
        y, page, i, n, events = 0.0, 1, 0, len(blocks), []
        while i < n:
            if blocks[i].page_before and y > 0:
                y, page = 0.0, page + 1
            j = i
            while j < n - 1 and blocks[j].keep and not blocks[j + 1].page_before and not blocks[j].break_after:
                j += 1
            chain = blocks[i:j + 1]
            total = sum(b.h for b in chain)
            need = sum(b.h for b in chain[:-1]) + chain[-1].min_h
            if y > 0 and y + need > P:
                events.append((i, j, P - y, page))
                y, page = 0.0, page + 1
            y += total
            while y > P:
                y -= P
                page += 1
            if chain[-1].break_after:
                y, page = 0.0, page + 1
            i = j + 1
        return events


# ---------- 修正手段 ----------

def _scale_picture(p, ratio):
    for ext in p.iter(f"{{{WP_NS}}}extent"):
        for k in ("cx", "cy"):
            ext.set(k, str(int(int(ext.get(k)) * ratio)))
    for ext in p.iter(f"{{{A_NS}}}ext"):
        if ext.get("cx") is not None:
            for k in ("cx", "cy"):
                ext.set(k, str(int(int(ext.get(k)) * ratio)))


def _label(chain, pattern, default):
    for b in chain:
        m = pattern.match(b.text or "")
        if m:
            return m.group(1).strip()         # 保留原题注的写法：“表1”或“表 1”
    return default


def _movable_after(blocks, j):
    """chain 之后可以挪到前面的正文段落下标（跳过紧跟的空白段和“注”）。"""
    k = j + 1
    while k < len(blocks) and blocks[k].kind == "p" and (blocks[k].empty or (blocks[k].text or "").startswith("注")):
        if blocks[k].break_after or blocks[k].page_before:
            return []
        k += 1
    out = []
    while k < len(blocks) and len(out) < MAX_MOVE_PARAS:
        b = blocks[k]
        if (b.kind != "p" or b.heading or b.img or b.math or b.keep or b.page_before or b.empty
                or b.text.startswith("式中")):
            break
        out.append(k)
        if b.break_after:
            break
        k += 1
    return out


def _plan_move(blocks, cand, fill):
    """在 fill pt 空白里放得下的段落；最后一段可以跨页（它至少有 4 行时）。返回 (下标列表, 填充高度)。"""
    acc, chosen = 0.0, []
    for k in cand:
        b = blocks[k]
        if acc + b.h <= fill:
            chosen.append(k)
            acc += b.h
            continue
        if b.min_h < b.h and acc + b.min_h <= fill:
            chosen.append(k)
            acc = fill
        break
    return chosen, acc


def _clear_keep(tr):
    for p in tr.iter(qn("w:p")):
        ppr = p.find(qn("w:pPr"))
        if ppr is not None:
            for kn in ppr.findall(qn("w:keepNext")):
                ppr.remove(kn)


def _continued_caption(cap_el, text):
    """按表题段落的格式做一个“表×（续）”段落（不含书签和域）。"""
    p = OxmlElement("w:p")
    ppr = cap_el.find(qn("w:pPr")) if cap_el is not None else None
    if ppr is not None:
        p.append(copy.deepcopy(ppr))
    else:
        ppr = OxmlElement("w:pPr")
        jc = OxmlElement("w:jc")
        jc.set(qn("w:val"), "center")
        ppr.append(jc)
        p.append(ppr)
    npr = p.find(qn("w:pPr"))
    if npr.find(qn("w:keepNext")) is None:
        npr.insert(1 if npr.find(qn("w:pStyle")) is not None else 0, OxmlElement("w:keepNext"))
    rpr = None
    if cap_el is not None:
        for r in cap_el.iter(qn("w:r")):
            if r.find(qn("w:t")) is not None and r.find(qn("w:rPr")) is not None:
                rpr = r.find(qn("w:rPr"))
                break
    for seg in re.findall(r"[⺀-鿿＀-￯]+|[^⺀-鿿＀-￯]+", text):
        r = OxmlElement("w:r")
        if rpr is not None:
            nr = copy.deepcopy(rpr)
            if not _CJK.match(seg) and cap_el is not None:
                # 西文、数字沿用表题里西文 run 的加粗设置
                for rr in cap_el.iter(qn("w:r")):
                    t = rr.find(qn("w:t"))
                    if t is not None and t.text and re.search(r"[0-9A-Za-z]", t.text) and rr.find(qn("w:rPr")) is not None:
                        nr = copy.deepcopy(rr.find(qn("w:rPr")))
                        break
            r.append(nr)
        t = OxmlElement("w:t")
        t.text = seg
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        r.append(t)
        p.append(r)
    return p


def _split_table(_lay, chain, tb, room):
    """把表格在 room pt 内能放下的行之后拆开。成功返回 (前段数据行数, 续表段落)。"""
    rows = tb.rows
    head = [r for r in rows if r["header"]]
    data = [r for r in rows if not r["header"]]
    if len(data) < MIN_SPLIT_ROWS:
        return None
    used = sum(r["h"] for r in head) + 2
    k = 0
    for r in data:
        if used + r["h"] > room:
            break
        used += r["h"]
        k += 1
    if k < MIN_PART_ROWS or len(data) - k < 2:
        return None
    tbl = tb.el
    new_tbl = copy.deepcopy(tbl)
    new_rows = [tr for tr in new_tbl.findall(qn("w:tr"))]
    for tr in [r["el"] for r in data[k:]]:
        tbl.remove(tr)
    _clear_keep(data[k - 1]["el"])
    n_head = len(head)
    for tr in new_rows[n_head:n_head + k]:
        new_tbl.remove(tr)
    cap_el = None
    for b in chain:
        if b.kind == "p" and _TABLE_LABEL.match(b.text or ""):
            cap_el = b.el
    label = _label(chain, _TABLE_LABEL, "")
    cont = _continued_caption(cap_el, f"{label}（续）" if label else "续表")
    tbl.addnext(cont)
    cont.addnext(new_tbl)
    return k, label


def _flow_table(tb):
    """让表格从第 MIN_PART_ROWS 行数据之后可以跨页：表题、表头和前几行仍在一起。"""
    data = [r for r in tb.rows if not r["header"]]
    if len(data) < MIN_SPLIT_ROWS:
        return False
    for r in data[MIN_PART_ROWS - 1:-1]:
        _clear_keep(r["el"])
    return True


def reduce_float_gaps(doc, gap_ratio: float = GAP_RATIO, move_text: bool = True, shrink: bool = True,
                      split_tables: bool = True, split_mode: str = "continued") -> list[str]:
    """估算分页并修正图表造成的大块空白。返回改动说明（中文），没有改动时返回空列表。
    split_mode："continued" 拆成“表×（续）”；"flow" 只让表格自然跨页（表头重复，不加续表字样）。"""
    if split_mode not in ("continued", "flow"):
        raise ValueError("split_mode 只能是 continued 或 flow")

    def split(chain, tb, room, page):
        name = _label(chain, _TABLE_LABEL, "") or "表格"
        if split_mode == "flow":
            if not _flow_table(tb):
                return False
            log.append(f"{name}在约第 {page} 页放不下，改为允许跨页（表头在下一页重复）")
            return True
        res = _split_table(None, chain, tb, room)
        if not res:
            return False
        log.append(f"{name}在约第 {page} 页末拆开：前 {res[0]} 行数据留在本页，其余转下页，"
                   f"上方加“{name}（续）”并重复表头")
        return True

    log: list[str] = []
    done_move, done_shrink, reported = set(), set(), set()
    for _ in range(MAX_ROUNDS):
        lay = _Layout(doc)
        blocks = lay.blocks()
        threshold = gap_ratio * lay.page_h
        slack = SLACK_LINES * lay.body_lh
        changed = False
        for i, j, gap, page in lay.simulate(blocks):
            chain = blocks[i:j + 1]
            imgs = [b for b in chain if b.img]
            tables = [b for b in chain if b.kind == "tbl" and b.kept_table]
            if gap < threshold or not (imgs or tables):
                continue
            total = sum(b.h for b in chain)
            room = gap - slack
            has_heading = any(b.heading for b in chain)
            key = chain[0].el
            name = _label(chain, _FIG_LABEL, "") or _label(chain, _TABLE_LABEL, "") or ("插图" if imgs else "表格")
            # 1. 插图略缩
            if imgs and shrink and len(imgs) == 1 and key not in done_shrink:
                img_h = imgs[0].h
                need = total - room
                ratio = (img_h - need) / img_h if img_h else 0
                if ratio >= SHRINK_SOFT:
                    _scale_picture(imgs[0].el, ratio)
                    done_shrink.add(key)
                    log.append(f"{name}缩小到原尺寸的 {ratio:.0%}，放进约第 {page} 页底部的空白")
                    changed = True
                    break
            # 2. 移正文填空
            cand = _movable_after(blocks, j) if (move_text and not has_heading and key not in done_move) else []
            chosen, filled = _plan_move(blocks, cand, room) if cand else ([], 0.0)
            fills = filled >= room - lay.body_lh
            if tables and split_tables:
                # 表格：超过一页直接拆；挪 1～2 段正文能填满才挪，否则先拆表
                if total > lay.page_h or not (fills and len(chosen) <= 2):
                    tb = tables[0]
                    if len(tables) == 1 and split(chain, tb, room - sum(b.h for b in chain[:chain.index(tb)]), page):
                        changed = True
                        break

            def move():
                anchor = chain[0].el
                for k in chosen:
                    anchor.addprevious(blocks[k].el)
                done_move.add(key)
                log.append(f"{name}后面的 {len(chosen)} 段正文移到{name}前面，填补约第 {page} 页底部"
                           f"约 {gap / 28.35:.1f} cm 的空白")

            if chosen and (fills or not imgs):
                move()
                changed = True
                break
            # 3. 插图再缩
            if imgs and shrink and len(imgs) == 1:
                img_h = imgs[0].h
                need = total - room
                ratio = (img_h - need) / img_h if img_h else 0
                if ratio >= SHRINK_HARD and key not in done_shrink:
                    _scale_picture(imgs[0].el, ratio)
                    done_shrink.add(key)
                    log.append(f"{name}缩小到原尺寸的 {ratio:.0%}，放进约第 {page} 页底部的空白")
                    changed = True
                    break
            # 3b. 插图：缩小也放不下时，能挪多少正文就挪多少
            if imgs and chosen:
                move()
                changed = True
                break
            # 4. 拆表
            if tables and split_tables and len(tables) == 1:
                tb = tables[0]
                if split(chain, tb, room - sum(b.h for b in chain[:chain.index(tb)]), page):
                    changed = True
                    break
            if key not in reported:
                reported.add(key)
                why = "后面紧接标题" if has_heading or not cand else "后面可挪的正文不够"
                extra = "缩小到 75% 也放不下" if imgs else f"表格不足 {MIN_SPLIT_ROWS} 行数据，不宜拆"
                log.append(f"{name}放不下，约第 {page} 页底部留空约 {gap / 28.35:.1f} cm（{why}，"
                           f"{extra}），未自动处理，请在 Word 里确认")
        if not changed:
            # 超过一页的不拆分表格：从页首开始也放不下时，继续拆
            for b in blocks:
                if b.kind == "tbl" and b.kept_table and b.h > lay.page_h and split_tables:
                    idx = blocks.index(b)
                    chain = blocks[max(0, idx - 2):idx + 1]
                    if split(chain, b, lay.page_h - SLACK_LINES * lay.body_lh - 40, "?"):
                        log[-1] = log[-1].replace("在约第 ? 页", "超过一页，")
                        changed = True
                        break
        if not changed:
            break
    return log
