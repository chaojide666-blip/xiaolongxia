# -*- coding: utf-8 -*-
"""
饭加鱼：Word 生成技能（全功能版）
支持：封面、页眉页脚、目录、表格、图片、引用、待办、超链接、行内格式
"""

from docx import Document
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from pathlib import Path
import re
import time
import os
from datetime import datetime


# =========================================================
# 工具函数
# =========================================================

def _set_cell_bg(cell, color_hex):
    """设置单元格底色。"""
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color_hex)
    tc_pr.append(shd)


def _set_cell_border(cell, color="CCCCCC", size="4"):
    """设置单元格边框。"""
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), size)
        el.set(qn("w:color"), color)
        tc_borders.append(el)
    tc_pr.append(tc_borders)


def _add_horizontal_line(doc, color="CCCCCC"):
    """加水平分割线。"""
    p = doc.add_paragraph()
    p_pr = p._p.get_or_add_pPr()
    p_bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color)
    p_bdr.append(bottom)
    p_pr.append(p_bdr)


def _add_page_number(paragraph):
    """插入页码字段。"""
    run = paragraph.add_run()
    fld1 = OxmlElement("w:fldChar")
    fld1.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE"
    fld2 = OxmlElement("w:fldChar")
    fld2.set(qn("w:fldCharType"), "end")
    run._r.append(fld1)
    run._r.append(instr)
    run._r.append(fld2)


def _add_hyperlink(paragraph, url, text):
    """插入超链接。"""
    part = paragraph.part
    r_id = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)
    new_run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    rPr.append(color)
    u = OxmlElement("w:u")
    u.set(qn("w:val"), "single")
    rPr.append(u)
    new_run.append(rPr)
    t = OxmlElement("w:t")
    t.text = text
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


def _parse_inline(paragraph, text):
    """解析行内格式：**粗体**、*斜体*、`代码`、==高亮==、[文字](url)"""
    pattern = re.compile(
        r"(\*\*.+?\*\*|`[^`]+?`|==.+?==|\*.+?\*|\[[^\]]+?\]\([^)]+?\))"
    )
    parts = pattern.split(text)

    for part in parts:
        if not part:
            continue

        if part.startswith("**") and part.endswith("**"):
            run = paragraph.add_run(part[2:-2])
            run.font.bold = True

        elif part.startswith("`") and part.endswith("`"):
            run = paragraph.add_run(part[1:-1])
            run.font.name = "Consolas"
            run.font.size = Pt(10)
            run.font.color.rgb = RGBColor.from_string("C7254E")
            # 灰色底
            rPr = run._r.get_or_add_rPr()
            shd = OxmlElement("w:shd")
            shd.set(qn("w:val"), "clear")
            shd.set(qn("w:fill"), "F7F7F9")
            rPr.append(shd)

        elif part.startswith("==") and part.endswith("=="):
            run = paragraph.add_run(part[2:-2])
            rPr = run._r.get_or_add_rPr()
            shd = OxmlElement("w:shd")
            shd.set(qn("w:val"), "clear")
            shd.set(qn("w:fill"), "FFF3B0")
            rPr.append(shd)

        elif part.startswith("[") and "](" in part:
            m = re.match(r"\[([^\]]+)\]\(([^)]+)\)", part)
            if m:
                _add_hyperlink(paragraph, m.group(2), m.group(1))

        elif part.startswith("*") and part.endswith("*") and not part.startswith("**"):
            run = paragraph.add_run(part[1:-1])
            run.font.italic = True

        else:
            paragraph.add_run(part)


# =========================================================
# 主函数
# =========================================================

def make_word(title, content):
    project_dir = Path(__file__).parent
    output_dir = project_dir / "生成的Word"
    output_dir.mkdir(exist_ok=True)

    safe_title = re.sub(r'[\\/:*?"<>|]', "", title)
    file_path = output_dir / f"{safe_title}_{int(time.time())}.docx"

    doc = Document()

    # =========================================================
    # 全局样式
    # =========================================================
    style = doc.styles["Normal"]
    style.font.name = "微软雅黑"
    style.font.size = Pt(11)
    style.element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
    style.paragraph_format.line_spacing = 1.5
    style.paragraph_format.space_after = Pt(6)

    # 标题样式
    title_styles = [
        (24, "1F4E79", "Heading 1"),
        (17, "2E74B5", "Heading 2"),
        (14, "2E74B5", "Heading 3"),
        (12, "404040", "Heading 4"),
    ]
    for size, color, name in title_styles:
        s = doc.styles[name]
        s.font.name = "微软雅黑"
        s.font.size = Pt(size)
        s.font.bold = True
        s.font.color.rgb = RGBColor.from_string(color)
        s.element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")

    # =========================================================
    # 封面页
    # =========================================================
    for _ in range(7):
        doc.add_paragraph()

    cover = doc.add_paragraph()
    cover.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = cover.add_run(title)
    r.font.name = "微软雅黑"
    r.font.size = Pt(30)
    r.font.bold = True
    r.font.color.rgb = RGBColor.from_string("1F4E79")
    r.element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rs = sub.add_run("饭加鱼 · 生成于 " + datetime.now().strftime("%Y 年 %m 月 %d 日"))
    rs.font.size = Pt(11)
    rs.font.color.rgb = RGBColor.from_string("808080")

    doc.add_page_break()

    # =========================================================
    # 页眉页脚
    # =========================================================
    section = doc.sections[0]
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)

    header_p = section.header.paragraphs[0]
    header_p.text = title
    header_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in header_p.runs:
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor.from_string("808080")

    footer_p = section.footer.paragraphs[0]
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer_p.add_run("第 ").font.size = Pt(9)
    _add_page_number(footer_p)
    footer_p.add_run(" 页").font.size = Pt(9)
    for run in footer_p.runs:
        run.font.color.rgb = RGBColor.from_string("808080")

    # =========================================================
    # 正文解析
    # =========================================================
    lines = content.split("\n")
    i = 0

    while i < len(lines):
        line = lines[i].rstrip()

        if not line.strip():
            i += 1
            continue

        stripped = line.strip()

        # -------- 水平分割线 --------
        if stripped in ("---", "***", "___"):
            _add_horizontal_line(doc)
            i += 1
            continue

        # -------- 居中 / 右对齐 --------
        if stripped.startswith("::: center"):
            i += 1
            while i < len(lines) and not lines[i].strip().startswith(":::"):
                p = doc.add_paragraph(lines[i].strip())
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                i += 1
            i += 1
            continue

        if stripped.startswith("::: right"):
            i += 1
            while i < len(lines) and not lines[i].strip().startswith(":::"):
                p = doc.add_paragraph(lines[i].strip())
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                i += 1
            i += 1
            continue

        # -------- 引用块 --------
        if stripped.startswith("> "):
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.5)
            run = p.add_run("💡 " + stripped[2:])
            run.font.size = Pt(11)
            run.font.color.rgb = RGBColor.from_string("B45309")
            p_pr = p._p.get_or_add_pPr()
            p_bdr = OxmlElement("w:pBdr")
            left = OxmlElement("w:left")
            left.set(qn("w:val"), "single")
            left.set(qn("w:sz"), "18")
            left.set(qn("w:space"), "8")
            left.set(qn("w:color"), "F59E0B")
            p_bdr.append(left)
            p_pr.append(p_bdr)
            i += 1
            continue

        # -------- 图片 --------
        img_match = re.match(r"!\[([^\]]*)\]\(([^)]+)\)", stripped)
        if img_match:
            img_path = img_match.group(2)
            if not os.path.isabs(img_path):
                img_path = str(project_dir / img_path)
            if os.path.exists(img_path):
                try:
                    doc.add_picture(img_path, width=Inches(5))
                    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                except Exception as e:
                    doc.add_paragraph(f"[图片加载失败：{e}]")
            else:
                doc.add_paragraph(f"[图片不存在：{img_path}]")
            i += 1
            continue

        # -------- 表格 --------
        if stripped.startswith("|"):
            table_lines = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                table_lines.append(lines[i].strip())
                i += 1

            rows = []
            for row in table_lines:
                cells = [x.strip() for x in row.strip("|").split("|")]
                rows.append(cells)

            rows = [r for r in rows if not all("-" in c for c in r)]

            if rows:
                table = doc.add_table(rows=len(rows), cols=len(rows[0]))
                table.style = "Table Grid"
                table.alignment = WD_TABLE_ALIGNMENT.CENTER

                for r_idx, row in enumerate(rows):
                    for c_idx, value in enumerate(row):
                        cell = table.cell(r_idx, c_idx)
                        cell.text = value

                        if r_idx == 0:
                            _set_cell_bg(cell, "1F4E79")
                            for para in cell.paragraphs:
                                for run in para.runs:
                                    run.font.bold = True
                                    run.font.color.rgb = RGBColor.from_string("FFFFFF")
                                    run.font.size = Pt(10)
                        elif r_idx % 2 == 0:
                            _set_cell_bg(cell, "F2F2F2")

                        _set_cell_border(cell)

                doc.add_paragraph()

            continue

        # -------- 待办清单 --------
        todo_match = re.match(r"^-\s*\[([ xX])\]\s*(.*)$", stripped)
        if todo_match:
            checked = todo_match.group(1).lower() == "x"
            text = todo_match.group(2)
            mark = "☑" if checked else "☐"
            p = doc.add_paragraph(f"{mark}  {text}")
            p.paragraph_format.left_indent = Cm(0.5)
            i += 1
            continue

        # -------- 标题 --------
        if stripped.startswith("####"):
            doc.add_heading(stripped.replace("#", "").strip(), level=4)
        elif stripped.startswith("###"):
            doc.add_heading(stripped.replace("#", "").strip(), level=3)
        elif stripped.startswith("##"):
            doc.add_heading(stripped.replace("#", "").strip(), level=2)
        elif stripped.startswith("#"):
            doc.add_heading(stripped.replace("#", "").strip(), level=1)

        # -------- 有序列表 --------
        elif re.match(r"^\d+\.\s", stripped):
            text = re.sub(r"^\d+\.\s", "", stripped)
            p = doc.add_paragraph(style="List Number")
            _parse_inline(p, text)

        # -------- 无序列表 --------
        elif stripped.startswith(("- ", "* ")):
            p = doc.add_paragraph(style="List Bullet")
            _parse_inline(p, stripped[2:])

        # -------- 普通段落 --------
        else:
            p = doc.add_paragraph()
            _parse_inline(p, stripped)

        i += 1

    doc.save(file_path)
    print("Word生成成功:", file_path)
    return str(file_path)
