#!/usr/bin/env python3
"""自测：造一份故意违规的文档，验证 check_docx.py 能查出各类问题；
用 fit_images.py 修图后图片类错误消失；用 docx_helpers 建的干净文档 0 ERROR，
并直接检查 XML：三线表的线、标题的大纲级别、题注 SEQ 域、交叉引用 REF 域与书签。

用法::  python3 selftest.py [工作目录]
"""
from __future__ import annotations

import os
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from docx import Document  # noqa: E402
from docx.shared import Cm, Pt  # noqa: E402
from PIL import Image, ImageDraw  # noqa: E402

import docx_helpers as h  # noqa: E402

PY = sys.executable


def run(script, *args):
    proc = subprocess.run([PY, os.path.join(HERE, script), *args], capture_output=True, text=True)
    return proc.returncode, proc.stdout + proc.stderr


def make_png(path, w, ht):
    im = Image.new("RGB", (w, ht), "white")
    d = ImageDraw.Draw(im)
    d.rectangle([10, 10, w - 10, ht - 10], outline="black", width=4)
    d.line([10, ht - 10, w - 10, 10], fill="blue", width=3)
    im.save(path)


def build_bad(path, big_png, small_png):
    doc = Document()
    doc.add_heading("一、引言", 1)
    doc.add_paragraph("按照用户要求，以下是本文的引言。本文存在一定不足。")
    p = doc.add_paragraph()
    r = p.add_run("这是一段 9 号字，用来塞表格。")
    r.font.size = Pt(9)
    doc.add_paragraph("**方法**：更好的做法是使用深度模型，但为了简便本文选择了线性回归。此外，此外，此外，这至关重要。")
    doc.add_paragraph("")
    doc.add_paragraph("")
    doc.add_heading("二、方法", 1)
    doc.add_paragraph("笔者认为方法仍有局限。我觉得我做得还不够好，我会继续努力，我我我。")
    doc.add_paragraph().add_run().add_picture(big_png, width=Cm(20), height=Cm(5))
    doc.add_paragraph("上图没有图题。")
    doc.add_heading("三、结果", 1)
    doc.add_paragraph("结果有待提高，仍需改进，效果不够理想。")
    doc.add_paragraph().add_run().add_picture(small_png, width=Cm(12))
    doc.add_paragraph("图 1 小图")
    doc.add_paragraph("如图 3 所示，数据见表 1。")
    fake = doc.add_paragraph()
    fake.add_run("（一）实验原理").bold = True
    t = doc.add_table(rows=3, cols=2)
    t.style = "Table Grid"
    for i in range(3):
        for j in range(2):
            t.cell(i, j).text = f"{i}{j}"
    doc.add_heading("四、总结", 1)
    doc.add_paragraph("综上所述，本文尚未完成全部目标。总而言之，遗憾之处很多！！！")
    doc.add_paragraph("前人用阶跃响应拟合时间常数[1]，也有用频域方法的[5]。")
    doc.add_heading("参考文献", 1)
    for t in ("[1] EINSTEIN A. Zur Elektrodynamik bewegter Körper[J]. Annalen der Physik, 1905, 17: 891-921.",
              "[2] 李四.书名[M].北京:出版社,2019[2024-01-01].",
              "[3] 王五，赵六．没有类型标识的文献．刊名，2020．",
              "[4] 某中心.实验指导[EB/OL].https://example.cn/a.html."):
        bp = doc.add_paragraph(t)
        bp.paragraph_format.first_line_indent = Cm(0.74)
    doc.save(path)


def build_good(path, big_png):
    doc = Document()
    h.reset_counters(by_chapter=False)
    h.setup_document(doc)
    h.add_heading(doc, "一、实验目的", 1)
    h.add_body(doc, "本实验测量 RC 电路的充放电曲线，并由曲线拟合得到时间常数。")
    h.add_heading(doc, "二、数据记录与处理", 1)
    h.add_heading(doc, "（一）充电过程", 2)
    h.add_body(doc, "充电过程中电压随时间的变化如[[fig:charge]]所示，原始读数见[[tab:raw]]。")
    h.add_figure(doc, big_png, "充电过程电压随时间的变化", key="charge")
    h.add_table(doc, [["时间/s", "电压/V"], ["0", "0.00"], ["5", "3.16"], ["10", "4.32"]],
                caption="充电过程原始数据", key="raw")
    h.add_heading(doc, "（二）充放电对比", 2)
    h.add_body(doc, "两个过程的对比见[[tab:cmp]]，[[fig:charge]]中的曲线与之一致。")
    h.add_table(doc, [["组别", "时间/s", "电压/V", "时间/s", "电压/V"],
                      ["第 1 组", "5", "3.16", "5", "1.84"],
                      ["第 2 组", "5", "3.18", "5", "1.82"],
                      ["平均", "5", "3.17", "5", "1.83"]],
                caption="充放电过程对比", key="cmp",
                group_header=[("", 1), ("充电", 2), ("放电", 2)], total_row=True,
                note="注：电压为三次读数的平均值。")
    h.add_body(doc, "拟合所用代码如下。")
    h.add_code(doc, "import numpy as np\ntau = np.polyfit(t, np.log(v_inf - v), 1)[0]")
    h.add_heading(doc, "三、实验总结", 1)
    h.add_body(doc, "拟合得到时间常数 9.8 s，与标称值 10 s 相差 2%。本次只测了一组电阻电容组合，后续可换两组元件复核拟合方法的稳定性。")
    h.add_body(doc, "用阶跃响应拟合时间常数的做法见[[cite:zhang2020]]，频域方法见[[cite:li2019,wang2021,zhao2022]]，"
                    "两类方法的比较见[[cite:zhang2020,wang2021]]，其中误差推导可参考文献[[cite-text:li2019]]。")
    g = h.gbt7714
    h.add_references(doc, [
        ("wang2021", g("J", ["王五", "赵六"], "一阶电路时间常数的频域测量", journal="物理实验", year=2021,
                       volume=41, issue=5, pages="22-26")),
        ("zhang2020", g("J", ["Smith J A", "Lee K"], "Time-constant estimation in first-order circuits",
                        journal="Measurement", year=2021, volume=172, pages="108-115",
                        url="https://doi.org/10.1016/j.example.2021.108")),
        ("zhao2022", g("EB", ["某某大学物理实验中心"], "RC电路实验指导", updated="2024-03-01",
                       accessed="2026-10-08", url="https://example.edu.cn/rc.html")),
        ("li2019", g("M", ["李四"], "电子测量误差分析：一本书名故意写得很长用来检查悬挂缩进下续行是否与第一行文字对齐",
                     edition="2版", place="北京", publisher="科学出版社", year=2019, pages="33-40")),
    ])
    h.add_page_number_footer(doc)
    h.finalize(doc)
    doc.save(path)


def build_chapter(path, big_png):
    """论文式：按章编号，摘要不计章号，引用先于图表出现。"""
    doc = Document()
    h.reset_counters(by_chapter=True)
    h.setup_document(doc)
    h.add_heading(doc, "摘要", 1)
    h.add_body(doc, "本文分析了某数据集上两种拟合方法的误差。")
    h.add_heading(doc, "1 引言", 1)
    h.add_body(doc, "实验装置见[[fig:setup]]，测量曲线见[[fig:curve]]。")
    h.add_figure(doc, big_png, "实验装置", key="setup")
    h.add_figure(doc, big_png, "测量曲线", key="curve")
    h.add_heading(doc, "2 方法", 1)
    h.add_heading(doc, "2.1 数据", 2)
    h.add_body(doc, "参数见[[tab:param]]，拟合结果见[[fig:fit]]；装置仍沿用[[fig:setup]]。")
    h.add_table(doc, [["参数", "取值"], ["采样率/Hz", "100"], ["时长/s", "60"]], caption="采集参数", key="param")
    h.add_figure(doc, big_png, "拟合结果", key="fit")
    h.add_body(doc, "采样方法沿用已有工作[[cite:a]]。")
    h.add_references(doc, [("a", h.gbt7714("D", ["钱七"], "一阶电路参数辨识方法研究", place="武汉",
                                            school="某某大学", year=2022, pages=27))], page_break=True)
    h.finalize(doc)
    doc.save(path)


W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def field_result(instr_el):
    run = instr_el.getparent().getnext()
    while run is not None:
        fc = run.find(f"{W}fldChar")
        if fc is not None and fc.get(f"{W}fldCharType") == "separate":
            return run.getnext()
        run = run.getnext()
    return None


def cite_checks(path, expect_cites, expect_bib_starts):
    """文献引用：正文显示、上标、书签存在；参考文献表：顺序、SEQ 域、悬挂缩进且无首行缩进。"""
    errs = []
    d = Document(path)
    body = d.element.body
    bms = {b.get(f"{W}name") for b in body.iter(f"{W}bookmarkStart")}
    shown = []
    for p in d.paragraphs:
        if not any((x.text or "").strip().startswith("REF Cite") for x in p._p.iter(f"{W}instrText")):
            continue
        import re as _re
        shown += _re.findall(r"\[[\d,\-]+\]", p.text)
        for it in p._p.iter(f"{W}instrText"):
            if (it.text or "").strip().startswith("REF Cite"):
                if it.text.split()[1] not in bms:
                    errs.append(f"文献引用指向不存在的书签 {it.text.split()[1]}")
    if shown != expect_cites:
        errs.append(f"文献引用显示 {shown} ≠ 期望 {expect_cites}")
    heads = [i for i, p in enumerate(d.paragraphs) if p.text.strip() == "参考文献"]
    entries = d.paragraphs[heads[-1] + 1:] if heads else []
    starts = [e.text[:e.text.index("\t")] if "\t" in e.text else e.text[:4] for e in entries]
    if starts != expect_bib_starts:
        errs.append(f"参考文献序号 {starts} ≠ 期望 {expect_bib_starts}")
    for e in entries:
        ind = e._p.find(f"{W}pPr/{W}ind")
        if ind is None or ind.get(f"{W}hanging") in (None, "0") or ind.get(f"{W}firstLine") or ind.get(f"{W}firstLineChars"):
            errs.append(f"参考文献“{e.text[:12]}”不是悬挂缩进")
        if ind is not None and ind.get(f"{W}left") != ind.get(f"{W}hanging"):
            errs.append(f"参考文献“{e.text[:12]}”左缩进≠悬挂量，序号没有顶格")
        if not any("SEQ 参考文献" in (x.text or "") for x in e._p.iter(f"{W}instrText")):
            errs.append(f"参考文献“{e.text[:12]}”序号不是 SEQ 域")
    # 上标
    for it in body.iter(f"{W}instrText"):
        if (it.text or "").strip().startswith("REF Cite"):
            res = field_result(it)
            sup = res.find(f"{W}rPr/{W}vertAlign") if res is not None else None
            p_text = "".join(t.text or "" for t in it.getparent().getparent().iter(f"{W}t"))
            if sup is None and "参考文献[" not in p_text:
                errs.append("上标引用的结果没有设为上标")
                break
    return errs


def xml_checks(path, expect_caps, expect_refs):
    """直接读 XML 断言：三线表、大纲级别、题注域、交叉引用。"""
    errs = []
    d = Document(path)
    body = d.element.body
    # 标题大纲级别
    for p in d.paragraphs:
        if p.style.name.startswith("Heading"):
            ol = p._p.find(f"{W}pPr/{W}outlineLvl")
            want = int(p.style.name.split()[-1]) - 1
            if ol is None or int(ol.get(f"{W}val")) != want:
                errs.append(f"标题“{p.text}”缺大纲级别 {want + 1}")
    # 三线表
    for ti, tbl in enumerate(body.iter(f"{W}tbl"), 1):
        tb = tbl.find(f"{W}tblPr/{W}tblBorders")
        get = lambda e: tb.find(f"{W}{e}")
        if get("top").get(f"{W}sz") != "12" or get("bottom").get(f"{W}sz") != "12":
            errs.append(f"表 {ti} 顶/底线不是 1.5 磅")
        for e in ("left", "right", "insideH", "insideV"):
            if get(e).get(f"{W}val") != "nil":
                errs.append(f"表 {ti} 的 {e} 线没去掉")
        szs = [b.get(f"{W}sz") for b in tbl.iter(f"{W}bottom") if b.getparent().tag == f"{W}tcBorders"]
        if "6" not in szs:
            errs.append(f"表 {ti} 没有 0.75 磅栏目线")
        if tbl.find(f".//{W}tblHeader") is None or tbl.find(f".//{W}cantSplit") is None:
            errs.append(f"表 {ti} 缺表头重复或行不跨页")
        for sp in tbl.iter(f"{W}spacing"):
            if sp.getparent().getparent().tag == f"{W}p" and (sp.get(f"{W}after") not in (None, "0")):
                errs.append(f"表 {ti} 单元格有段后间距")
                break
    # 题注
    caps = [p.text for p in d.paragraphs if p.text[:1] in "图表" and any("SEQ" in (x.text or "") for x in p._p.iter(f"{W}instrText"))]
    got = [c.split()[0] + " " + c.split()[1] for c in caps]
    if got != expect_caps:
        errs.append(f"题注编号 {got} ≠ 期望 {expect_caps}")
    # 交叉引用：REF 域的缓存结果
    refs = []
    for r_instr in body.iter(f"{W}instrText"):
        if (r_instr.text or "").strip().startswith("REF _Ref"):
            run = r_instr.getparent().getnext()
            while run is not None:
                fc = run.find(f"{W}fldChar")
                if fc is not None and fc.get(f"{W}fldCharType") == "separate":
                    refs.append(run.getnext().find(f"{W}t").text)
                    break
                run = run.getnext()
            bm = r_instr.text.split()[1]
            if not any(b.get(f"{W}name") == bm for b in body.iter(f"{W}bookmarkStart")):
                errs.append(f"REF 指向不存在的书签 {bm}")
    if refs != expect_refs:
        errs.append(f"交叉引用显示 {refs} ≠ 期望 {expect_refs}")
    # OOXML 子元素顺序（Word 对顺序敏感）
    for tag, order in h._ORDERS.items():
        for el in body.iter(f"{W}{tag}"):
            idx = [order.index(h._local(c.tag)) for c in el if h._local(c.tag) in order]
            if idx != sorted(idx):
                errs.append(f"{tag} 子元素顺序不合 schema：{[h._local(c.tag) for c in el]}")
                break
    return errs


def table_variant_checks(work):
    """各种“双线表”必须被查出；模板默认表格样式带线时，add_table 仍要产出干净的三线表。"""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    rows = [["时间/s", "电压/V"], ["0", "0.00"], ["5", "3.16"]]
    errs = []

    def make(name, mutate):
        d = Document()
        h.reset_counters()
        h.setup_document(d)
        t = h.add_table(d, rows, caption="测试")
        mutate(d, t)
        path = os.path.join(work, f"tbl_{name}.docx")
        d.save(path)
        return path

    def double_rules(d, t):
        for e in ("top", "bottom"):
            t._tbl.tblPr.find(qn("w:tblBorders")).find(qn("w:" + e)).set(qn("w:val"), "double")

    def header_double(d, t):
        for c in t.rows[0].cells:
            c._tc.tcPr.find(qn("w:tcBorders")).find(qn("w:bottom")).set(qn("w:val"), "double")

    def cell_spacing(d, t):
        cs = OxmlElement("w:tblCellSpacing")
        cs.set(qn("w:w"), "20")
        cs.set(qn("w:type"), "dxa")
        h._place(t._tbl.tblPr, cs)

    def styled(d, t):
        t.style = d.styles["Light List Accent 1"]
        t._tbl.tblPr.remove(t._tbl.tblPr.find(qn("w:tblLook")))

    def para_border(d, t):
        ppr = t.cell(0, 0).paragraphs[0]._p.get_or_add_pPr()
        bdr = OxmlElement("w:pBdr")
        b = OxmlElement("w:bottom")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), "6")
        bdr.append(b)
        h._place(ppr, bdr)

    def two_lines(d, t):
        for c in t.rows[0].cells:
            tcb = c._tc.tcPr.find(qn("w:tcBorders"))
            tcb.remove(tcb.find(qn("w:bottom")))

    for name, fn in (("double_rules", double_rules), ("header_double", header_double),
                     ("cell_spacing", cell_spacing), ("style_cond", styled), ("para_border", para_border),
                     ("two_lines", two_lines)):
        code, out = run("check_docx.py", make(name, fn))
        if "[TABLE_NOT_3LINE]" not in out:
            errs.append(f"{name} 这种双线/非三线表没有被查出")

    # 模板的默认表格样式是带线的样式：add_table 必须不受影响
    d = Document()
    for st in d.styles.element.findall(qn("w:style")):
        if st.get(qn("w:type")) == "table":
            if st.get(qn("w:styleId")) == "LightList-Accent1":
                st.set(qn("w:default"), "1")
            elif st.get(qn("w:default")):
                del st.attrib[qn("w:default")]
    h.reset_counters()
    h.setup_document(d)
    h.add_table(d, rows, caption="模板默认样式带线")
    path = os.path.join(work, "tbl_template_default.docx")
    d.save(path)
    code, out = run("check_docx.py", path)
    if "[TABLE_" in out:
        errs.append("模板默认表格样式带线时，add_table 生成的表没通过三线表检查")
    # 反过来：同一模板下 doc.add_table 不处理，要被查出
    d2 = Document(path)
    d2.add_table(rows=2, cols=2).cell(0, 0).text = "x"
    p2 = os.path.join(work, "tbl_template_raw.docx")
    d2.save(p2)
    code, out = run("check_docx.py", p2)
    if "第 2 个表格不是三线表" not in out:
        errs.append("模板默认样式下直接 doc.add_table 的表没被查出")
    return errs


def gbt_checks():
    """gbt7714() 输出与 GB/T 7714—2025 模板逐字比对。"""
    g = h.gbt7714
    expect = [
        (g("J", ["张三", "李四", "王五", "赵六"], "阶跃响应法测量RC电路时间常数", journal="大学物理", year=2020,
           volume=39, issue=2, pages="45–48"),
         "张三,李四,王五,等.阶跃响应法测量RC电路时间常数[J].大学物理,2020,39(2):45-48."),
        (g("J", ["Smith J A", "Lee K"], "Title", journal="Measurement", year=2021, volume=172, pages="108-115",
           url="https://doi.org/10.1/x"),
         "Smith J A, Lee K. Title[J/OL]. Measurement, 2021, 172: 108-115. https://doi.org/10.1/x."),
        (g("M//", ["孙八"], "第3章 一阶电路", host_authors=["周九"], host_title="电路分析基础", place="北京",
           publisher="高等教育出版社", year=2018, pages="60-75"),
         "孙八.第3章 一阶电路[M]//周九.电路分析基础.北京:高等教育出版社,2018:60-75."),
        (g("C", ["孙八"], "RC电路实验教学改进", conference="全国高校物理实验教学研讨会", year=2023, pages="12-15"),
         "孙八.RC电路实验教学改进[C]//全国高校物理实验教学研讨会,2023:12-15."),
        (g("R", ["某研究所"], "年度报告", report_no="ABC-01", date="2023-09-30"),
         "某研究所.年度报告:ABC-01[R].2023-09-30."),
        (g("S", None, "电工术语 基本术语", std_no="GB/T 2900.1-2008"), "GB/T 2900.1—2008 电工术语 基本术语[S]."),
        (g("DS", ["张三"], "测量数据", version="V1.0", platform="某数据中心", updated="2023-09-10",
           accessed="2025-07-15", url="https://e.cn/1", pid="CSTR:1.2"),
         "张三.测量数据[DS/OL].V1.0.某数据中心(2023-09-10)[2025-07-15].https://e.cn/1.CSTR:1.2."),
    ]
    errs = [f"gbt7714 输出 {got} ≠ {want}" for got, want in expect if got != want]
    try:
        g("EB", None, "无引用日期的网页", url="https://x.cn")
        errs.append("gbt7714 网页缺引用日期没有报错")
    except ValueError:
        pass
    return errs


def pagination_checks(work):
    """图表在页底放不下：插图靠挪正文或缩图消掉大块空白；短表改为允许跨页（报告不加“续表”字样）。"""
    import pagination
    errs = []
    png = os.path.join(work, "blank.png")
    Image.new("RGB", (1500, 1000), "white").save(png, dpi=(300, 300))
    line = "单行正文用于调整插图位置。"

    def build(n, kind, follow="text", rows=70):
        h.reset_counters(by_chapter=False)
        d = Document()
        h.setup_document(d)
        h.add_heading(d, "1 测试", 1)
        for _ in range(n):
            h.add_body(d, line)
        if kind == "fig":
            h.add_body(d, "见[[fig:a]]。")
            h.add_figure(d, png, "测试插图", width_cm=14, key="a")
        else:
            h.add_body(d, "见[[tab:a]]。")
            h.add_table(d, [["序号", "x/mm"]] + [[str(i), f"{i}.00"] for i in range(1, rows + 1)],
                        caption="测试表格", key="a", note="测试表注")
        if follow == "heading":
            h.add_heading(d, "2 下一章", 1)
        for _ in range(6):
            h.add_body(d, "后文。" + line * 6)
        h.finalize(d, fix_gaps=False)
        return d

    def big_gaps(d):
        lay = pagination._Layout(d)
        return [e for e in lay.simulate(lay.blocks()) if e[2] > pagination.GAP_RATIO * lay.page_h]

    moved = shrunk = 0
    for follow in ("text", "heading"):
        for n in range(4, 44):
            d = build(n, "fig", follow)
            before = big_gaps(d)
            log = pagination.reduce_float_gaps(d)
            if before and not log:
                errs.append(f"{follow} n={n}：有大块留白却没有任何处理或说明")
            if big_gaps(d) and not any("未自动处理" in x for x in log):
                errs.append(f"{follow} n={n}：留白没消掉，也没记进日志：{log}")
            if follow == "heading" and any("移到" in x for x in log):
                errs.append(f"n={n}：后面紧接标题时不应挪正文")
            moved += any("移到" in x for x in log)
            shrunk += any("缩小到" in x for x in log)
    if not moved or not shrunk:
        errs.append(f"分页用例没有覆盖挪正文（{moved}）和缩图（{shrunk}）两种处理")
    flowed = 0
    for n in range(4, 44):
        d = build(n, "tab", rows=12)
        log = pagination.reduce_float_gaps(d, split_mode="flow")
        if big_gaps(d) and not any("未自动处理" in x for x in log):
            errs.append(f"表格 n={n}：留白没消掉，也没记进日志：{log}")
        if any(p.text.endswith("（续）") for p in d.paragraphs) or len(d.tables) != 1:
            errs.append(f"表格 n={n}：报告不应拆表或加“续表”字样")
        flowed += any("允许跨页" in x for x in log)
    if not flowed:
        errs.append("分页用例没有覆盖短表改为允许跨页")
    out = os.path.join(work, "pagination.docx")
    d.save(out)
    Document(out)
    return ["分页：" + e for e in errs]


def main():
    work = sys.argv[1] if len(sys.argv) > 1 else tempfile.mkdtemp(prefix="word-report-selftest-")
    os.makedirs(work, exist_ok=True)
    big = os.path.join(work, "big.png")
    small = os.path.join(work, "small.png")
    make_png(big, 2400, 1200)
    make_png(small, 300, 200)
    bad = os.path.join(work, "bad.docx")
    good = os.path.join(work, "good.docx")
    build_bad(bad, big, small)
    build_good(good, big)
    chap = os.path.join(work, "chapter.docx")
    build_chapter(chap, big)

    failures = []

    code, out = run("check_docx.py", bad)
    print("---- check bad.docx ----")
    print(out)
    expected = ["SMALL_FONT", "FORBIDDEN", "HEDGE", "MARKDOWN", "AI_WORD", "EMPTY_PARAS", "FLAW_SPREAD",
                "PRONOUN_MIX", "IMG_TOO_WIDE", "IMG_DISTORTED", "IMG_LOW_DPI", "IMG_NO_CAPTION", "EXCLAIM",
                "TABLE_NOT_3LINE", "TABLE_SPACING", "NO_OUTLINE", "CAPTION_NO_FIELD", "XREF_MANUAL", "XREF_DANGLING",
                "CITE_MANUAL", "CITE_DANGLING", "BIB_UNCITED", "BIB_NO_FIELD", "BIB_INDENT", "BIB_FORMAT",
                "PAGE_NUMBER"]
    for frag in ("缺文献类型标识", "非网络文献却写了引用日期", "缺 [引用日期]", "全角著录符号", "EINSTEIN"):
        if frag not in out:
            failures.append(f"bad.docx 的 BIB_FORMAT 没查出：{frag}")
    for c in expected:
        if f"[{c}]" not in out:
            failures.append(f"bad.docx 未查出 {c}")
    if code != 1:
        failures.append(f"bad.docx 退出码应为 1，实际 {code}")

    code, out = run("fit_images.py", bad, "--inplace")
    print("---- fit_images bad.docx ----")
    print(out)
    code, out = run("check_docx.py", bad)
    print("---- re-check bad.docx ----")
    print(out)
    for c in ("IMG_TOO_WIDE", "IMG_DISTORTED"):
        if f"[{c}]" in out:
            failures.append(f"fit_images 之后仍有 {c}")

    code, out = run("check_docx.py", good)
    print("---- check good.docx ----")
    print(out)
    if code != 0 or "[ERROR]" in out:
        failures.append("good.docx 仍有 ERROR")

    code, out = run("check_docx.py", chap)
    print("---- check chapter.docx ----")
    print(out)
    if code != 0 or "[ERROR]" in out:
        failures.append("chapter.docx 仍有 ERROR")

    failures += ["good.docx: " + e for e in xml_checks(
        good, ["图 1", "表 1", "表 2"], ["图 1", "表 1", "表 2", "图 1"])]
    failures += ["good.docx: " + e for e in cite_checks(
        good, ["[1]", "[2-4]", "[1,3]", "[2]"], ["[1]", "[2]", "[3]", "[4]"])]
    failures += ["chapter.docx: " + e for e in cite_checks(chap, ["[1]"], ["[1]"])]
    failures += ["chapter.docx: " + e for e in xml_checks(
        chap, ["图 1.1", "图 1.2", "表 2.1", "图 2.1"], ["图 1.1", "图 1.2", "表 2.1", "图 2.1", "图 1.1"])]

    failures += gbt_checks()
    failures += table_variant_checks(work)
    failures += pagination_checks(work)

    print("工作目录：", work)
    if failures:
        print("SELFTEST FAIL")
        for f in failures:
            print(" -", f)
        return 1
    print("SELFTEST PASS")
    return 0


if __name__ == "__main__":
    sys.exit(main())
