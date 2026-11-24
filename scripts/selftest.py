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
    doc.add_paragraph("前人用阶跃响应拟合时间常数[1]，也有用频域方法的[3]。")
    doc.add_heading("参考文献", 1)
    for t in ("[1] 张三. 题名[J]. 刊名, 2020, 3(2): 1-5.", "[2] 李四. 书名[M]. 北京: 出版社, 2019."):
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
    h.add_references(doc, [
        ("wang2021", "王五, 赵六. 一阶电路时间常数的频域测量[J]. 物理实验, 2021, 41(5): 22-26."),
        ("zhang2020", "张三. 阶跃响应法测量 RC 电路时间常数[J]. 大学物理, 2020, 39(2): 45-48."),
        ("zhao2022", "赵七. 电路测量中的数据拟合方法[M]. 北京: 高等教育出版社, 2022: 101-118."),
        ("li2019", "李四. 电子测量误差分析[M]. 2 版. 北京: 科学出版社, 2019: 33-40. 这一条故意写得很长，"
                   "用来检查悬挂缩进下续行是否与第一行文字对齐，而不是顶到序号下面。"),
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
    h.add_references(doc, [("a", "张三. 题名[J]. 刊名, 2020, 3(2): 1-5.")], page_break=True)
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
                "CITE_MANUAL", "CITE_DANGLING", "BIB_UNCITED", "BIB_NO_FIELD", "BIB_INDENT"]
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
