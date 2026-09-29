from docx import Document
from pathlib import Path
import re
import time


def make_word(title, content):

    # 项目目录
    project_dir = Path(__file__).parent

    output_dir = project_dir / "生成的Word"

    output_dir.mkdir(
        exist_ok=True
    )


    # 文件名
    safe_title = re.sub(
        r'[\\/:*?"<>|]',
        "",
        title
    )

    file_path = (
        output_dir /
        f"{safe_title}_{int(time.time())}.docx"
    )


    doc = Document()


    # 标题
    doc.add_heading(
        title,
        level=1
    )


    lines = content.split("\n")

    i = 0


    while i < len(lines):

        line = lines[i].strip()


        # 空行跳过
        if not line:
            i += 1
            continue


        # =========================
        # Markdown表格
        # =========================

        if line.startswith("|"):

            table_lines = []


            while (
                i < len(lines)
                and lines[i].strip().startswith("|")
            ):
                table_lines.append(
                    lines[i].strip()
                )
                i += 1


            rows = []

            for row in table_lines:

                cells = [
                    x.strip()
                    for x in row.strip("|").split("|")
                ]

                rows.append(cells)


            # 删除分隔线
            rows = [
                r for r in rows
                if not all(
                    "-" in c
                    for c in r
                )
            ]


            if rows:

                table = doc.add_table(
                    rows=len(rows),
                    cols=len(rows[0])
                )


                table.style = "Table Grid"


                for r, row in enumerate(rows):

                    for c, value in enumerate(row):

                        table.cell(
                            r,
                            c
                        ).text = value


            continue



        # =========================
        # 标题
        # =========================

        if line.startswith("###"):

            doc.add_heading(
                line.replace("#","").strip(),
                level=3
            )


        elif line.startswith("##"):

            doc.add_heading(
                line.replace("#","").strip(),
                level=2
            )


        elif line.startswith("#"):

            doc.add_heading(
                line.replace("#","").strip(),
                level=1
            )


        # =========================
        # 列表
        # =========================

        elif line.startswith(
            ("- ","* ")
        ):

            doc.add_paragraph(
                line[2:],
                style="List Bullet"
            )


        else:

            doc.add_paragraph(
                line
            )


        i += 1



    doc.save(
        file_path
    )


    print(
        "Word生成成功:",
        file_path
    )


    return str(file_path)