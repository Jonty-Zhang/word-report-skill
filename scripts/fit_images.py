#!/usr/bin/env python3
"""把 .docx 里超尺寸、变形的图片按比例修正。

用法::

    python3 fit_images.py 报告.docx                    # 输出 报告_fitted.docx
    python3 fit_images.py 报告.docx --inplace          # 原地覆盖
    python3 fit_images.py 报告.docx -o 新文件.docx
    选项：--max-width 15 --max-height 18 --no-fix-aspect

规则：宽 > max-width 或高 > max-height 时等比缩小；像素宽高比与显示宽高比相差 3% 以上视为拉伸变形，
按像素比重算高度（--no-fix-aspect 关闭）。只改嵌入式/浮动式 DrawingML 图片，不动 VML 旧式图片。
"""
from __future__ import annotations

import argparse
import io
import os
import sys

from docx import Document
from docx.oxml.ns import qn

try:
    from PIL import Image
except ImportError:  # pragma: no cover
    Image = None

EMU_PER_CM = 360000
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def iter_drawings(doc):
    body = doc.element.body
    for tag in (f"{{{WP}}}inline", f"{{{WP}}}anchor"):
        for el in body.iter(tag):
            yield el


def pixel_size(doc, drawing):
    blip = drawing.find(f".//{{{A}}}blip")
    if blip is None or Image is None:
        return None
    rid = blip.get(f"{{{R}}}embed")
    part = doc.part.related_parts.get(rid) if rid else None
    if part is None:
        return None
    try:
        with Image.open(io.BytesIO(part.blob)) as im:
            return im.size
    except Exception:
        return None


def set_size(drawing, cx: int, cy: int):
    ext = drawing.find(f"{{{WP}}}extent")
    if ext is not None:
        ext.set("cx", str(cx))
        ext.set("cy", str(cy))
    for aext in drawing.iter(f"{{{A}}}ext"):
        parent = aext.getparent()
        if parent is not None and parent.tag == f"{{{A}}}xfrm":
            aext.set("cx", str(cx))
            aext.set("cy", str(cy))


def fit(doc, max_w_cm: float, max_h_cm: float, fix_aspect: bool = True):
    changes = []
    for idx, dr in enumerate(iter_drawings(doc), 1):
        ext = dr.find(f"{{{WP}}}extent")
        if ext is None:
            continue
        cx, cy = int(ext.get("cx")), int(ext.get("cy"))
        w, h = cx / EMU_PER_CM, cy / EMU_PER_CM
        new_w, new_h = w, h
        note = []
        px = pixel_size(doc, dr)
        if fix_aspect and px and px[0] and px[1] and w > 0 and h > 0:
            shown = w / h
            real = px[0] / px[1]
            if abs(shown / real - 1) > 0.03:
                new_h = new_w / real
                note.append(f"变形 {shown:.2f}→{real:.2f}")
        if new_w > max_w_cm:
            scale = max_w_cm / new_w
            new_w, new_h = max_w_cm, new_h * scale
            note.append(f"宽超 {max_w_cm} cm")
        if new_h > max_h_cm:
            scale = max_h_cm / new_h
            new_w, new_h = new_w * scale, max_h_cm
            note.append(f"高超 {max_h_cm} cm")
        if abs(new_w - w) > 0.01 or abs(new_h - h) > 0.01:
            set_size(dr, int(round(new_w * EMU_PER_CM)), int(round(new_h * EMU_PER_CM)))
            changes.append((idx, w, h, new_w, new_h, "，".join(note)))
    return changes


def main(argv=None):
    ap = argparse.ArgumentParser(description="按比例修正 docx 中的图片尺寸")
    ap.add_argument("docx")
    ap.add_argument("-o", "--output")
    ap.add_argument("--inplace", action="store_true")
    ap.add_argument("--max-width", type=float, default=15.0, help="宽度上限 cm，默认 15")
    ap.add_argument("--max-height", type=float, default=18.0, help="高度上限 cm，默认 18")
    ap.add_argument("--no-fix-aspect", action="store_true", help="不修正拉伸变形")
    args = ap.parse_args(argv)

    if not os.path.exists(args.docx):
        print(f"文件不存在：{args.docx}", file=sys.stderr)
        return 2
    doc = Document(args.docx)
    changes = fit(doc, args.max_width, args.max_height, not args.no_fix_aspect)
    if args.inplace:
        out = args.docx
    elif args.output:
        out = args.output
    else:
        root, ext = os.path.splitext(args.docx)
        out = f"{root}_fitted{ext}"
    if not changes:
        print("没有需要调整的图片。")
        if out != args.docx:
            doc.save(out)
            print(f"已另存：{out}")
        return 0
    for idx, w, h, nw, nh, note in changes:
        print(f"图片 {idx}: {w:.2f}×{h:.2f} cm → {nw:.2f}×{nh:.2f} cm（{note}）")
    doc.save(out)
    print(f"共调整 {len(changes)} 张，已保存：{out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
