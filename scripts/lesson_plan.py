# -*- coding: utf-8 -*-
"""教案（lesson plan）确定性工具：check / render / inspect。

用法：
    py -3.12 lesson_plan.py check  --input 教案.md --report review.md [--level basic|changed|enhanced]
    py -3.12 lesson_plan.py render --input 教案.md --output 教案.docx
    py -3.12 lesson_plan.py inspect --input 教案.docx --report inspection.md

要求 Python 3.10+、python-docx、PyYAML。
"""
from __future__ import annotations

import argparse
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree

import yaml
from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

SCRIPT_VERSION = "0.1.0"
RULE_VERSION = "1"

# —— 稳定合同 ——
COURSE_REQUIRED = (
    "学院", "课程名称", "主讲教师", "授课专业", "授课学期", "制定日期",
    "课程类别", "课程性质", "授课语言", "学分", "总学时",
    "教材", "授课学院", "考核类型", "考核形式", "考核方式", "总评成绩比例",
)
LESSON_REQUIRED_FACTS = ("章节", "授课题目", "周次", "课时安排")
LESSON_REQUIRED_SECTIONS = ("教学目标", "教学重难点", "教学反思")
PROCESS_STEPS = ("课前任务", "互动导入", "传授新知", "过关检测", "课堂小结", "作业布置", "考勤")
VALID_LEVELS = {"basic", "changed", "enhanced"}
UNRESOLVED = ("[待填写]", "to_confirm", "待填写")
AI_AUTHORSHIP_PATTERNS = (
    re.compile(r"(?:本教案|本文档|本文件).{0,12}(?:由)?(?:AI|ChatGPT|Codex|模型).{0,8}(?:生成|编写|撰写)", re.I),
    re.compile(r"(?:AI|ChatGPT|Codex|模型)(?:自动)?(?:生成|编写|撰写)(?:的教案|本文档|本文件)?", re.I),
)


class Issue:
    def __init__(self, code: str, severity: str, message: str, location: str):
        self.code = code
        self.severity = severity  # "error" | "warning"
        self.message = message
        self.location = location


# ==================== 解析 ====================

def _front_matter(text: str) -> tuple[dict, str]:
    if not text.startswith("---"):
        raise ValueError("教案必须以 YAML front matter 开头")
    parts = text.split("---", 2)
    if len(parts) != 3:
        raise ValueError("YAML front matter 未闭合")
    metadata = yaml.safe_load(parts[1]) or {}
    if not isinstance(metadata, dict):
        raise ValueError("YAML front matter 必须是映射")
    return metadata, parts[2]


def _cells(line: str) -> list[str]:
    return [c.strip() for c in line.strip().strip("|").split("|")]


def _is_separator(line: str) -> bool:
    cells = _cells(line)
    return bool(cells) and all(re.fullmatch(r":?-{3,}:?", c) for c in cells)


def _parse_table_rows(headers: list[str], lines: list[str], start: int) -> tuple[list[dict], int]:
    rows = []
    i = start
    while i < len(lines) and lines[i].strip().startswith("|"):
        values = _cells(lines[i])
        if len(values) != len(headers):
            raise ValueError(f"表格列数与表头不一致：{' | '.join(headers)}")
        rows.append(dict(zip(headers, values)))
        i += 1
    return rows, i


def _missing(value: str) -> bool:
    return value is None or value.strip() == "" or value.strip() in UNRESOLVED


def parse_working(path: Path) -> dict:
    text = Path(path).read_text(encoding="utf-8")
    metadata, body = _front_matter(text)
    if "# 正式教案内容" in body:
        formal = body.split("# 正式教案内容", 1)[1]
    else:
        formal = body
    if "# 内部追踪区" in formal:
        formal = formal.split("# 内部追踪区", 1)[0]

    lines = formal.splitlines()
    course_info: dict[str, str] = {}
    course_intro_lines: list[str] = []
    lessons: list[dict] = []
    current_lesson: dict | None = None
    current_section: str | None = None

    i = 0
    while i < len(lines):
        line = lines[i].rstrip()
        if line.startswith("### "):
            lid = line[4:].strip()
            current_lesson = {"lesson_id": lid, "facts": {}, "sections": {}, "process": []}
            lessons.append(current_lesson)
            current_section = None
            i += 1
            continue
        if line.startswith("#### "):
            current_section = line[5:].strip()
            i += 1
            continue
        if line.startswith("## "):
            current_section = None
            i += 1
            continue
        if line.startswith("# "):
            i += 1
            continue
        if line.startswith("|") and i + 1 < len(lines) and _is_separator(lines[i + 1]):
            headers = _cells(line)
            rows, i = _parse_table_rows(headers, lines, i + 2)
            if headers[:2] == ["字段", "内容"]:
                for row in rows:
                    field = row.get("字段", "")
                    value = row.get("内容", "")
                    if current_lesson is None:
                        course_info[field] = value
                    else:
                        current_lesson["facts"][field] = value
            elif headers[:3] == ["教学步骤", "主要教学内容", "设计意图"] and current_lesson is not None:
                current_lesson["process"] = rows
            else:
                raise ValueError(f"无法识别的表格：{' | '.join(headers)}")
            continue
        # 普通文字行
        if line.strip():
            if current_lesson is None and current_section == "课程简介":
                course_intro_lines.append(line)
            elif current_lesson is not None and current_section:
                current_lesson["sections"].setdefault(current_section, []).append(line)
        i += 1

    return {
        "metadata": metadata,
        "course_info": course_info,
        "course_intro": "\n".join(course_intro_lines),
        "lessons": lessons,
    }


# ==================== check ====================

def _strip_md(text: str) -> str:
    return text.replace("**", "").replace("<br>", "\n")


def check_working(data: dict, level: str = "basic") -> list[Issue]:
    issues: list[Issue] = []
    info = data["course_info"]
    lessons = data["lessons"]

    if "schema_version" not in data["metadata"]:
        issues.append(Issue("missing-front-matter", "error", "缺少 schema_version", "front matter"))

    for field in COURSE_REQUIRED:
        if _missing(info.get(field, "")):
            issues.append(Issue("missing-course-field", "error", f"课程必填字段缺失或未确认：{field}", "课程基本信息"))

    if not lessons:
        issues.append(Issue("no-lesson", "error", "至少需要一个教学单元", "教学单元"))
        return issues

    for lesson in lessons:
        lid = lesson["lesson_id"]
        facts = lesson["facts"]
        sections = lesson["sections"]
        for field in LESSON_REQUIRED_FACTS:
            if _missing(facts.get(field, "")):
                issues.append(Issue("missing-lesson-fact", "error", f"教学单元必填字段缺失：{field}", lid))
        for section in LESSON_REQUIRED_SECTIONS:
            text = _strip_md("\n".join(sections.get(section, [])))
            if _missing(text):
                issues.append(Issue("missing-lesson-section", "error", f"教学单元缺少小节：{section}", lid))

        # 教学过程七步完整性
        step_labels = [row.get("教学步骤", "") for row in lesson["process"]]
        for step in PROCESS_STEPS:
            if not any(label.strip().startswith(step) for label in step_labels):
                issues.append(Issue("missing-process-step", "error", f"教学过程缺少步骤：{step}", lid))
        for row in lesson["process"]:
            if _missing(row.get("主要教学内容", "")):
                issues.append(Issue("empty-process-step", "error", f"教学过程步骤内容为空：{row.get('教学步骤', '')}", lid))

        # 时间一致性（警告，不阻断）
        minutes = [int(m) for m in re.findall(r"(\d+)\s*min", facts.get("教学设计", ""))]
        total_minutes = sum(minutes)
        if minutes and total_minutes and total_minutes != 100:
            issues.append(Issue("time-mismatch", "warning", f"教学设计各环节合计 {total_minutes} min，与默认 100 min 不一致", lid))

    # 正式内容占位符扫描
    formal_values = [info.get(k, "") for k in info]
    for lesson in lessons:
        formal_values += list(lesson["facts"].values())
        for sec_lines in lesson["sections"].values():
            formal_values.append("\n".join(sec_lines))
        for row in lesson["process"]:
            formal_values += [row.get("教学步骤", ""), row.get("主要教学内容", ""), row.get("设计意图", "")]
    for value in formal_values:
        for token in UNRESOLVED:
            if token in str(value):
                issues.append(Issue("unresolved-placeholder", "error", f"正式内容仍含未解决占位符：{token}", "正式教案内容"))
                break

    return issues


def check_file(input_path: Path, level: str) -> tuple[dict, list[Issue]]:
    data = parse_working(input_path)
    issues = check_working(data, level)
    return data, issues


def write_report(path: Path, input_path: Path, level: str, data: dict, issues: list[Issue]) -> None:
    errors = [i for i in issues if i.severity == "error"]
    warnings = [i for i in issues if i.severity == "warning"]
    lines = [
        f"# 教案检查报告",
        f"",
        f"- 输入：{input_path}",
        f"- 级别：{level}",
        f"- 教学单元数：{len(data['lessons'])}",
        f"- 阻断问题：{len(errors)}；警告：{len(warnings)}",
        f"",
    ]
    if errors:
        lines.append("## 阻断问题")
        for i in errors:
            lines.append(f"- `{i.code}` [{i.location}] {i.message}")
    if warnings:
        lines.append("## 警告")
        for i in warnings:
            lines.append(f"- `{i.code}` [{i.location}] {i.message}")
    if not errors and not warnings:
        lines.append("无阻断问题、无警告。")
    path.write_text("\n".join(lines), encoding="utf-8")


# ==================== render ====================

def _set_run_font(run, east_asia: str, latin: str, size: float, bold: bool | None = None) -> None:
    run.font.name = latin
    run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    rfonts.set(qn("w:ascii"), latin)
    rfonts.set(qn("w:hAnsi"), latin)
    rfonts.set(qn("w:eastAsia"), east_asia)


def _set_cell_text(cell, text: str, *, bold: bool = False, centered: bool = False,
                   size: float = 9, east: str = "宋体", latin: str = "Times New Roman") -> None:
    cell.text = ""
    segs = str(text).replace("<br>", "\n").split("\n")
    first = True
    for seg in segs:
        if first:
            p = cell.paragraphs[0]
            first = False
        else:
            p = cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if centered else WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.space_before = Pt(0)
        run = p.add_run(seg)
        _set_run_font(run, east, latin, size, bold)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def _keep_row_together(row) -> None:
    trpr = row._tr.get_or_add_trPr()
    if trpr.find(qn("w:cantSplit")) is None:
        trpr.append(OxmlElement("w:cantSplit"))


def _repeat_table_header(row) -> None:
    trpr = row._tr.get_or_add_trPr()
    header = OxmlElement("w:tblHeader")
    header.set(qn("w:val"), "true")
    trpr.append(header)


def _set_table_grid(table, widths_cm: list[float]) -> None:
    tbl = table._tbl
    tblPr = tbl.tblPr
    layout = tblPr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tblPr.append(layout)
    layout.set(qn("w:type"), "fixed")
    grid = tbl.find(qn("w:tblGrid"))
    if grid is None:
        grid = OxmlElement("w:tblGrid")
        tbl.insert(1, grid)
    for gc in list(grid):
        grid.remove(gc)
    for w in widths_cm:
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(int(round(w * 567))))
        grid.append(gc)


def _add_labeled_row(table, label: str, value: str, *, label_bold=True, value_centered=False) -> None:
    """单标签行：第 1 列标签，第 2–5 列合并为内容。"""
    row = table.add_row()
    _keep_row_together(row)
    _set_cell_text(row.cells[0], label, bold=label_bold, centered=True)
    merged = row.cells[1].merge(row.cells[4])
    _set_cell_text(merged, value, centered=value_centered)


def _build_cover(doc, info: dict) -> None:
    p = doc.add_paragraph(style="教案封面标题")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(50)
    p.paragraph_format.space_after = Pt(36)
    _set_run_font(p.add_run("教 学 设 计"), "华文中宋", "华文中宋", 56, True)

    table = doc.add_table(rows=7, cols=2)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_table_grid(table, [3.5, 10.5])
    for i, field in enumerate(("学院", "专业（教研室）", "课程名称", "主讲教师", "职称", "授课学期", "授课专业")):
        _set_cell_text(table.rows[i].cells[0], field, bold=True, centered=True, size=12)
        _set_cell_text(table.rows[i].cells[1], info.get(field, ""), centered=True, size=12)
        _keep_row_together(table.rows[i])

    p = doc.add_paragraph(style="教案封面落款")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(28)
    _set_run_font(p.add_run("教 务 处 制"), "华文中宋", "华文中宋", 16, False)
    p = doc.add_paragraph(style="教案封面落款")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_run_font(p.add_run(info.get("制定日期", "")), "华文中宋", "华文中宋", 16, False)
    doc.add_page_break()


def _build_course_info(doc, info: dict, course_intro: str) -> None:
    p = doc.add_paragraph(style="教案节标题")
    _set_run_font(p.add_run("课程基本信息"), "黑体", "黑体", 16, True)

    table = doc.add_table(rows=0, cols=6)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_table_grid(table, [2.6, 2.7, 2.7, 2.7, 2.7, 2.7])

    def full_row(label, value):
        r = table.add_row()
        _keep_row_together(r)
        _set_cell_text(r.cells[0], label, bold=True, centered=True)
        _set_cell_text(r.cells[1].merge(r.cells[5]), value, centered=True)

    full_row("课程名称（中文）", info.get("课程名称", ""))
    full_row("课程名称（英文）", info.get("课程名称（英文）", ""))

    r = table.add_row()
    _keep_row_together(r)
    _set_cell_text(r.cells[0], "课程类别", bold=True, centered=True)
    _set_cell_text(r.cells[1], info.get("课程类别", ""), centered=True)
    _set_cell_text(r.cells[2], "课程性质", bold=True, centered=True)
    _set_cell_text(r.cells[3], info.get("课程性质", ""), centered=True)
    _set_cell_text(r.cells[4], "授课语言", bold=True, centered=True)
    _set_cell_text(r.cells[5], info.get("授课语言", ""), centered=True)

    r = table.add_row()
    _keep_row_together(r)
    _set_cell_text(r.cells[0], "授课学期", bold=True, centered=True)
    _set_cell_text(r.cells[1].merge(r.cells[2]), info.get("授课学期", ""), centered=True)
    _set_cell_text(r.cells[3], "学分", bold=True, centered=True)
    _set_cell_text(r.cells[4].merge(r.cells[5]), info.get("学分", ""), centered=True)

    r = table.add_row()
    _keep_row_together(r)
    _set_cell_text(r.cells[0], "课程学时及分配", bold=True, centered=True)
    for j, h in enumerate(("总学时", "理论", "实验/上机", "实习", "其他"), start=1):
        _set_cell_text(r.cells[j], h, bold=True, centered=True)
    r2 = table.add_row()
    _keep_row_together(r2)
    r2.cells[0].merge(r.cells[0])
    for j, field in enumerate(("总学时", "理论学时", "实验/上机学时", "实习学时", "其他学时"), start=1):
        _set_cell_text(r2.cells[j], info.get(field, "0"), centered=True)

    full_row("适用专业", info.get("授课专业", ""))
    full_row("教材", info.get("教材", ""))
    full_row("授课学院", info.get("授课学院", ""))
    full_row("先修课程", info.get("先修课程", ""))
    full_row("后续课程", info.get("后续课程", ""))
    full_row("考核类型", info.get("考核类型", ""))
    full_row("考核形式", info.get("考核形式", ""))
    full_row("考核方式", info.get("考核方式", ""))
    full_row("总评成绩比例", info.get("总评成绩比例", ""))

    intro_text = _strip_md(course_intro)
    r = table.add_row()
    _keep_row_together(r)
    _set_cell_text(r.cells[0], "课程简介", bold=True, centered=True)
    _set_cell_text(r.cells[1].merge(r.cells[5]), intro_text, centered=False)

    notes = (
        "注：1.课程类别：选填“通识教育类/学科基础类/专业类”；",
        "2.课程性质：选填“选修/必修”；",
        "3.授课语言：选填“中文、英文或其他语种”；",
        "4.若为集中实习类实践或课程设计类实践，以周为单位记学时。",
        "5.表中（ ）选项请打“√”。",
    )
    for note in notes:
        p = doc.add_paragraph(style="教案注文")
        p.paragraph_format.space_after = Pt(0)
        _set_run_font(p.add_run(note), "黑体", "黑体", 9, False)
    doc.add_paragraph()


def _build_lesson(doc, lesson: dict) -> None:
    facts = lesson["facts"]
    sections = lesson["sections"]
    process = lesson["process"]

    table = doc.add_table(rows=0, cols=5)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_table_grid(table, [2.0, 1.6, 1.8, 5.3, 5.3])

    r = table.add_row()
    _keep_row_together(r)
    _repeat_table_header(r)
    _set_cell_text(r.cells[0], "章节", bold=True, centered=True)
    _set_cell_text(r.cells[1], facts.get("章节", ""), centered=True)
    _set_cell_text(r.cells[2], "授课题目", bold=True, centered=True)
    _set_cell_text(r.cells[3].merge(r.cells[4]), facts.get("授课题目", ""), centered=True)

    r = table.add_row()
    _keep_row_together(r)
    _set_cell_text(r.cells[0], "周次", bold=True, centered=True)
    _set_cell_text(r.cells[1], facts.get("周次", ""), centered=True)
    _set_cell_text(r.cells[2], "课时安排", bold=True, centered=True)
    _set_cell_text(r.cells[3].merge(r.cells[4]), facts.get("课时安排", ""), centered=True)

    def sec(name: str) -> str:
        return _strip_md("\n".join(sections.get(name, [])))

    for label, value in (
        ("知识脉络图", facts.get("知识脉络图", "")),
        ("教学目标", sec("教学目标")),
        ("课程思政", sec("课程思政")),
        ("专创融合", sec("专创融合")),
        ("教学重难点", sec("教学重难点")),
        ("教学方法", facts.get("教学方法", "")),
        ("教学用具", facts.get("教学用具", "")),
        ("教学设计", facts.get("教学设计", "")),
    ):
        _add_labeled_row(table, label, value)

    r = table.add_row()
    _keep_row_together(r)
    _set_cell_text(r.cells[0], "教学过程", bold=True, centered=True)
    _set_cell_text(r.cells[1].merge(r.cells[3]), "教学步骤及主要教学内容", bold=True, centered=True)
    _set_cell_text(r.cells[4], "设计意图", bold=True, centered=True)

    for step in process:
        r = table.add_row()
        _keep_row_together(r)
        _set_cell_text(r.cells[0], step.get("教学步骤", ""), bold=True, centered=True)
        _set_cell_text(r.cells[1].merge(r.cells[3]), step.get("主要教学内容", ""), centered=False)
        _set_cell_text(r.cells[4], step.get("设计意图", ""), centered=False)

    r = table.add_row()
    _keep_row_together(r)
    _set_cell_text(r.cells[0], "课后\n教学反思", bold=True, centered=True)
    _set_cell_text(r.cells[1].merge(r.cells[4]), sec("教学反思"), centered=False)

    r = table.add_row()
    _keep_row_together(r)
    _set_cell_text(r.cells[0], "本章节\n参考文献", bold=True, centered=True)
    _set_cell_text(r.cells[1].merge(r.cells[4]), facts.get("参考文献", ""), centered=False)

    doc.add_paragraph()


def render_working(data: dict, output_path: Path, template_path: Path) -> None:
    doc = Document(template_path)
    body = doc._element.body
    for child in list(body):
        if child.tag != qn("w:sectPr"):
            body.remove(child)

    info = data["course_info"]
    _build_cover(doc, info)
    _build_course_info(doc, info, data.get("course_intro", ""))
    for lesson in data["lessons"]:
        _build_lesson(doc, lesson)

    cp = doc.core_properties
    cp.author = ""
    cp.last_modified_by = ""
    cp.comments = ""
    cp.keywords = ""
    doc.save(str(output_path))


# ==================== inspect ====================

class InspectionResult:
    def __init__(self, clean: bool, codes: tuple[str, ...], visible_text: str):
        self.clean = clean
        self.codes = codes
        self.visible_text = visible_text


def inspect_docx(path: Path) -> InspectionResult:
    path = Path(path)
    findings: list[str] = []
    visible_parts: list[str] = []
    with zipfile.ZipFile(path) as archive:
        names = set(archive.namelist())
        story_names = sorted(
            n for n in names
            if n == "word/document.xml"
            or re.fullmatch(r"word/(?:header|footer)\d+\.xml", n)
            or n in {"word/footnotes.xml", "word/endnotes.xml"}
        )
        joined_xml = []
        for name in story_names:
            payload = archive.read(name)
            joined_xml.append(payload.decode("utf-8", errors="replace"))
            root = ElementTree.fromstring(payload)
            for element in root.iter():
                if element.tag == qn("w:t") and element.text:
                    visible_parts.append(element.text)
        xml = "\n".join(joined_xml)
        visible_text = "\n".join(visible_parts)

        if any(n.startswith("word/comments") for n in names) or "<w:commentReference" in xml:
            findings.append("COMMENTS_PRESENT")
        if re.search(r"<w:(?:ins|del|moveFrom|moveTo)(?:\s|>)", xml):
            findings.append("TRACKED_CHANGES_PRESENT")
        if "<w:vanish" in xml or "<w:webHidden" in xml:
            findings.append("HIDDEN_TEXT_PRESENT")
        if "[待填写]" in visible_text or "to_confirm" in visible_text or "待填写" in visible_text:
            findings.append("UNRESOLVED_PLACEHOLDER")
        if any(p.search(visible_text) for p in AI_AUTHORSHIP_PATTERNS):
            findings.append("AI_AUTHORSHIP_MARKER")

        metadata = ""
        for name in ("docProps/core.xml", "docProps/custom.xml", "docProps/app.xml"):
            if name in names:
                metadata += archive.read(name).decode("utf-8", errors="replace")
        if re.search(r"OpenAI|ChatGPT|Codex|GPT-[0-9]", metadata, re.I):
            findings.append("TOOL_METADATA_PRESENT")

    codes = tuple(dict.fromkeys(findings))
    return InspectionResult(clean=not codes, codes=codes, visible_text=visible_text)


def write_inspection_report(path: Path, input_path: Path, result: InspectionResult) -> None:
    lines = [
        f"# Word 清洁检查报告",
        f"",
        f"- 输入：{input_path}",
        f"- 结论：{'通过' if result.clean else '未通过'}",
    ]
    if result.codes:
        lines.append("- 发现：")
        lines += [f"  - {c}" for c in result.codes]
    lines.append(f"- 可见文本长度：{len(result.visible_text)} 字符")
    path.write_text("\n".join(lines), encoding="utf-8")


# ==================== CLI ====================

def _parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="lesson_plan.py", description="教案确定性工具：check / render / inspect")
    sub = p.add_subparsers(dest="command", required=True)

    c = sub.add_parser("check", help="检查工作 Markdown 结构")
    c.add_argument("--input", required=True)
    c.add_argument("--report", default="review.md")
    c.add_argument("--level", default="basic", choices=sorted(VALID_LEVELS))

    r = sub.add_parser("render", help="把工作 Markdown 编译为 Word 教案")
    r.add_argument("--input", required=True)
    r.add_argument("--output", required=True)
    r.add_argument("--template", default=str(Path(__file__).parent.parent / "assets" / "lesson-plan-master.docx"))

    i = sub.add_parser("inspect", help="检查最终 DOCX 的清洁门禁")
    i.add_argument("--input", required=True)
    i.add_argument("--report", default="inspection.md")
    return p


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    if args.command == "check":
        data, issues = check_file(Path(args.input), args.level)
        write_report(Path(args.report), Path(args.input), args.level, data, issues)
        errors = [i for i in issues if i.severity == "error"]
        for i in issues:
            print(f"[{i.severity}] {i.code}: {i.message}")
        print(f"报告已写入 {args.report}")
        return 1 if errors else 0

    if args.command == "render":
        data, issues = check_file(Path(args.input), "basic")
        errors = [i for i in issues if i.severity == "error"]
        if errors:
            print("渲染被阻断，存在以下阻断问题：")
            for i in errors:
                print(f"  [{i.location}] {i.message}")
            return 1
        render_working(data, Path(args.output), Path(args.template))
        print(f"已生成 {args.output}")
        return 0

    if args.command == "inspect":
        result = inspect_docx(Path(args.input))
        write_inspection_report(Path(args.report), Path(args.input), result)
        print(f"清洁检查：{'通过' if result.clean else '未通过'}（{', '.join(result.codes) or '无' }）")
        print(f"报告已写入 {args.report}")
        return 0 if result.clean else 1

    return 2


if __name__ == "__main__":
    sys.exit(main())
