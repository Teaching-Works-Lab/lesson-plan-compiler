# -*- coding: utf-8 -*-
"""生成教案空白样式模板 lesson-plan-master.docx（样式模板 / 可见表单格式）。

用法：
    py -3.12 make_master.py [--output assets/lesson-plan-master.docx]

该 master 只包含页面设置与命名样式，不含任何课程内容；渲染器加载它、清空正文后按数据模板重建内容。
"""
import argparse
from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt


def _set_east(style, east: str) -> None:
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    rfonts.set(qn("w:ascii"), style.font.name or "Times New Roman")
    rfonts.set(qn("w:hAnsi"), style.font.name or "Times New Roman")
    rfonts.set(qn("w:eastAsia"), east)


def add_style(doc, name, latin, east, size, bold=False, centered=False):
    style = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
    style.font.name = latin
    style.font.size = Pt(size)
    style.font.bold = bold
    if centered:
        style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_east(style, east)
    return style


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", default=str(Path(__file__).parent.parent / "assets" / "lesson-plan-master.docx"))
    args = parser.parse_args()

    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(1.9)
    section.bottom_margin = Cm(2.3)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.0)

    # Normal：正文中文宋体、西文 Times New Roman，10.5 pt（五号）
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(10.5)
    _set_east(normal, "宋体")

    # 教案专用样式
    add_style(doc, "教案封面标题", "华文中宋", "华文中宋", 56, bold=True, centered=True)
    add_style(doc, "教案封面落款", "华文中宋", "华文中宋", 16, bold=False, centered=True)
    add_style(doc, "教案节标题", "黑体", "黑体", 16, bold=True)
    add_style(doc, "教案表格文字", "Times New Roman", "宋体", 9, bold=False)
    add_style(doc, "教案注文", "黑体", "黑体", 9, bold=False)

    out = Path(args.output)
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
    print(f"已生成样式模板：{out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
