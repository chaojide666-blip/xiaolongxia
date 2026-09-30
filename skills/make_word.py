# -*- coding: utf-8 -*-
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from pathlib import Path
import re
import time
from datetime import datetime


def _set_cell_bg(cell, color_hex):
    """给表格单元格设置底色。"""
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color_hex)
    tc_pr.append(shd)


def _add_horizontal_line(doc):
    """加一条水平分割线。"""
    p = doc.add_paragraph()
    p_pr = p._p.get_or_add_pPr()
    p_bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "CCCCCC")
    p_bdr.append(bottom)
    p_pr.append(p_bdr)


def _add_page_number(paragraph):
    """在段落里插入页码字段。"""
    run = paragraph.add_run()
    fld_char1 = OxmlElement("w:fldChar")
    fld_char1.set(qn("w:fldCharType"), "begin")
    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = "PAGE"
    fld_char2 = OxmlElement("w:fldChar")
    fld_char2.set(qn("w:fldCharType"), "end")
    run._r.append(fld_char1)
    run._r.append(instr_text)
    run._r.append(fld_char2)


def make_word(title, content):
    # 项目目录
    project_dir = Path(__file__).parent
    output_dir = project_dir / "生成的Word"
    output_dir.mkdir(exist_ok=True)

    # 文件名
    safe_title = re.sub(r'[\\/:*?"<>|]', "", title)
    file_path = output_dir / f"{safe_title}_{int(time.time())}.docx"

    doc = Document()

    # =========================================================
    # 全局样式：正文
    # =========================================================
    style = doc.styles["Normal"]
    style.font.name = "微软雅黑"
    style.font.size = Pt(11)
    style.element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
    style.paragraph_format.line_spacing = 1.5
    style.paragraph_format.space_after = Pt(6)

    # 标题样式
    for i, (size, color) in enumerate([(22, "1F4E79"), (16, "2E74B5"), (13, "2E74B5")], start=1):
        h = doc.styles[f"Heading {i}"]
        h.font.name = "微软雅黑"
        h.font.size = Pt(size)
        h.font.bold = True
        h.font.color.rgb = RGBColor.from_string(color)
        h.element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")

    # =========================================================
    # 封面页
    # =========================================================
    for _ in range(6):
        doc.add_paragraph()

    cover_title = doc.add_paragraph()
    cover_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = cover_title.add_run(title)
    run.font.name = "微软雅黑"
    run.font.size = Pt(28)
    run.font.bold = True
    run.font.color.rgb = RGBColor.from_string("1F4E79")
    run.element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")

    # 分割线
    line_p = doc.add_paragraph()
    line_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    line_run = line_p.add_run("———————")
    line_run.font.size = Pt(12)
    line_run.font.color.rgb = RGBColor.from_string("CCCCCC")

    # 日期
    date_p = doc.add_paragraph()
    date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    date_run = date_p.add_run(datetime.now().strftime("%Y 年 %m 月 %d 日"))
    date_run.font.size = Pt(12)
    date_run.font.color.rgb = RGBColor.from_string("808080")

    # 分页
    doc.add_page_break()

    # =========================================================
    # 页眉页脚
    # =========================================================
    section = doc.sections[0]
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)

    header = section.header
    header_p = header.paragraphs[0]
    header_p.text = title
    header_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in header_p.runs:
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor.from_string("808080")

    footer = section.footer
    footer_p = footer.paragraphs[0]
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
        line = lines[i].strip()

        # 空行
        if not line:
            i += 1
            continue

        # =========================
        # 水平分割线
        # =========================
        if line in ("---", "***", "___"):
            _add_horizontal_line(doc)
            i += 1
            continue

        # =========================
        # 引用块 / 高亮提示
        # =========================
        if line.startswith("> "):
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.5)
            run = p.add_run("💡 " + line[2:])
            run.font.size = Pt(11)
            run.font.color.rgb = RGBColor.from_string("B45309")
            # 加边框
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

        # =========================
        # Markdown 表格
        # =========================
        if line.startswith("|"):
            table_lines = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                table_lines.append(lines[i].strip())
                i += 1

            rows = []
            for row in table_lines:
                cells = [x.strip() for x in row.strip("|").split("|")]
                rows.append(cells)

            # 删除分隔线
            rows = [
                r for r in rows
                if not all("-" in c for c in r)
            ]

            if rows:
                table = doc.add_table(rows=len(rows), cols=len(rows[0]))
                table.style = "Table Grid"
                table.alignment = WD_TABLE_ALIGNMENT.CENTER

                for r_idx, row in enumerate(rows):
                    for c_idx, value in enumerate(row):
                        cell = table.cell(r_idx, c_idx)
                        cell.text = value

                        # 表头：深蓝底 + 白字
                        if r_idx == 0:
                            _set_cell_bg(cell, "1F4E79")
                            for para in cell.paragraphs:
                                for run in para.runs:
                                    run.font.bold = True
                                    run.font.color.rgb = RGBColor.from_string("FFFFFF")
                                    run.font.size = Pt(10)
                        # 斑马纹：偶数行浅灰底
                        elif r_idx % 2 == 0:
                            _set_cell_bg(cell, "F2F2F2")

            doc.add_paragraph()
            continue

        # =========================
        # 标题
        # =========================
        if line.startswith("###"):
            doc.add_heading(line.replace("#", "").strip(), level=3)
        elif line.startswith("##"):
            doc.add_heading(line.replace("#", "").strip(), level=2)
        elif line.startswith("#"):
            doc.add_heading(line.replace("#", "").strip(), level=1)

        # =========================
        # 有序列表
        # =========================
        elif re.match(r"^\d+\.\s", line):
            text = re.sub(r"^\d+\.\s", "", line)
            doc.add_paragraph(text, style="List Number")

        # =========================
        # 无序列表
        # =========================
        elif line.startswith(("- ", "* ")):
            doc.add_paragraph(line[2:], style="List Bullet")

        # =========================
        # 加粗行（**xxx**）
        # =========================
        elif line.startswith("**") and line.endswith("**"):
            p = doc.add_paragraph()
            run = p.add_run(line.strip("*"))
            run.font.bold = True
            run.font.size = Pt(11)

        # =========================
        # 普通段落
        # =========================
        else:
            doc.add_paragraph(line)

        i += 1

    # 保存
    doc.save(file_path)
    print("Word生成成功:", file_path)
    return str(file_path)
