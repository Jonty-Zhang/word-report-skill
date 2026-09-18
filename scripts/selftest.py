#!/usr/bin/env python3
"""自测：造一份故意违规的文档，验证 check_docx.py 能查出各类问题；
用 fit_images.py 修图后图片类错误消失；用 docx_helpers 建的干净文档 0 ERROR。

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
    doc.add_heading("四、总结", 1)
    doc.add_paragraph("综上所述，本文尚未完成全部目标。总而言之，遗憾之处很多！！！")
    doc.save(path)


def build_good(path, big_png):
    doc = Document()
    h.setup_document(doc)
    h.reset_counters()
    h.add_heading(doc, "一、实验目的", 1)
    h.add_body(doc, "本实验测量 RC 电路的充放电曲线，并由曲线拟合得到时间常数。")
    h.add_heading(doc, "二、数据记录与处理", 1)
    h.add_body(doc, "充电过程中电压随时间的变化如图 1 所示，原始读数见表 1。")
    h.add_figure(doc, big_png, "充电过程电压随时间的变化")
    h.add_table(doc, [["时间/s", "电压/V"], ["0", "0.00"], ["5", "3.16"], ["10", "4.32"]], caption="充电过程原始数据")
    h.add_body(doc, "拟合所用代码如下。")
    h.add_code(doc, "import numpy as np\ntau = np.polyfit(t, np.log(v_inf - v), 1)[0]")
    h.add_heading(doc, "三、实验总结", 1)
    h.add_body(doc, "拟合得到时间常数 9.8 s，与标称值 10 s 相差 2%。本次只测了一组电阻电容组合，后续可换两组元件复核拟合方法的稳定性。")
    h.add_page_number_footer(doc)
    doc.save(path)


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

    failures = []

    code, out = run("check_docx.py", bad)
    print("---- check bad.docx ----")
    print(out)
    expected = ["SMALL_FONT", "FORBIDDEN", "HEDGE", "MARKDOWN", "AI_WORD", "EMPTY_PARAS", "FLAW_SPREAD",
                "PRONOUN_MIX", "IMG_TOO_WIDE", "IMG_DISTORTED", "IMG_LOW_DPI", "IMG_NO_CAPTION", "EXCLAIM"]
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
