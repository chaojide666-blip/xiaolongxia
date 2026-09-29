# -*- coding: utf-8 -*-
"""
饭加鱼：Excel 数据技能 v2

支持：
- .xlsx / .xlsm：读取工作表、预览、汇总、搜索、区域
- .xls：读取旧版 Excel（BIFF），使用 xlrd
- 特殊目录：桌面 / 下载 / 文档
- 只读，不修改原文件
"""

from pathlib import Path
from typing import Any
import json


def _path(path: str) -> Path:
    if not path or not str(path).strip():
        raise ValueError("Excel 路径不能为空。")

    value = str(path).strip().strip('"').rstrip("\\/")

    home = Path.home()
    special_dirs = {
        "桌面": home / "Desktop",
        "desktop": home / "Desktop",
        "下载": home / "Downloads",
        "downloads": home / "Downloads",
        "文档": home / "Documents",
        "documents": home / "Documents",
    }

    if value in special_dirs:
        return special_dirs[value].resolve()

    return Path(value).expanduser().resolve()


def _check_excel(target: Path) -> dict | None:
    if not target.exists():
        return {"success": False, "error": f"Excel 文件不存在：{target}"}
    if not target.is_file():
        return {"success": False, "error": f"这不是文件：{target}"}
    if target.suffix.lower() not in {".xlsx", ".xlsm", ".xls"}:
        return {
            "success": False,
            "error": "当前版本支持 .xlsx / .xlsm / .xls。",
        }
    return None


def _load_xlsx(path: Path):
    try:
        from artifact_tool import Blob, SpreadsheetFile
    except ImportError as e:
        raise RuntimeError(
            "当前 Python 环境没有安装 artifact_tool，无法读取 .xlsx / .xlsm。"
        ) from e
    return SpreadsheetFile.import_xlsx(Blob.load(str(path)))


def _load_xls(path: Path):
    try:
        import xlrd
    except ImportError as e:
        raise RuntimeError(
            "读取旧版 .xls 需要 xlrd。请运行：python -m pip install xlrd"
        ) from e
    return xlrd.open_workbook(str(path), on_demand=True)


def _sheet_names_xlsx(wb):
    result = wb.inspect({
        "kind": "sheet",
        "include": "id,name",
    })

    names = []
    for line in result.ndjson.splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            obj = json.loads(line)
        except Exception:
            continue
        if isinstance(obj, dict) and obj.get("name"):
            names.append(str(obj["name"]))
    return names


def _xls_cell_value(cell):
    # xlrd 日期类型先按 Excel 序列值返回，避免未经上下文擅自改写。
    return cell.value


def _xls_sheet_values(sheet, max_rows: int, max_cols: int):
    rows = []
    row_limit = min(sheet.nrows, max_rows)
    col_limit = min(sheet.ncols, max_cols)

    for r in range(row_limit):
        row = []
        for c in range(col_limit):
            row.append(_xls_cell_value(sheet.cell(r, c)))
        rows.append(row)

    return rows


def list_excel_sheets(path: str) -> dict:
    target = _path(path)
    err = _check_excel(target)
    if err:
        return err

    try:
        if target.suffix.lower() == ".xls":
            wb = _load_xls(target)
            names = wb.sheet_names()
            wb.release_resources()
        else:
            wb = _load_xlsx(target)
            names = _sheet_names_xlsx(wb)

        return {
            "success": True,
            "path": str(target),
            "name": target.name,
            "sheet_count": len(names),
            "sheets": names,
        }
    except Exception as e:
        return {
            "success": False,
            "error": f"读取 Excel 工作表失败：{e}",
        }


def preview_excel(
    path: str,
    sheet_name: str = "",
    max_rows: int = 30,
    max_cols: int = 20,
) -> dict:
    target = _path(path)
    err = _check_excel(target)
    if err:
        return err

    try:
        max_rows = max(2, min(int(max_rows), 500))
        max_cols = max(2, min(int(max_cols), 100))
    except (TypeError, ValueError):
        max_rows, max_cols = 30, 20

    try:
        if target.suffix.lower() == ".xls":
            wb = _load_xls(target)
            names = wb.sheet_names()

            if not names:
                wb.release_resources()
                return {
                    "success": True,
                    "path": str(target),
                    "name": target.name,
                    "sheet_count": 0,
                    "sheets": [],
                    "data": [],
                }

            if not sheet_name:
                sheet_name = names[0]

            if sheet_name not in names:
                wb.release_resources()
                return {
                    "success": False,
                    "error": f"找不到工作表：{sheet_name}",
                    "available_sheets": names,
                }

            sheet = wb.sheet_by_name(sheet_name)
            values = _xls_sheet_values(sheet, max_rows, max_cols)
            row_count = sheet.nrows
            col_count = sheet.ncols
            wb.release_resources()

        else:
            wb = _load_xlsx(target)
            names = _sheet_names_xlsx(wb)

            if not names:
                return {
                    "success": True,
                    "path": str(target),
                    "name": target.name,
                    "sheet_count": 0,
                    "sheets": [],
                    "data": [],
                }

            if not sheet_name:
                sheet_name = names[0]

            if sheet_name not in names:
                return {
                    "success": False,
                    "error": f"找不到工作表：{sheet_name}",
                    "available_sheets": names,
                }

            sheet = wb.worksheets.get_item(sheet_name)
            values = sheet.get_range_by_indexes(
                0, 0, max_rows, max_cols
            ).values

            # artifact_tool 预览规模以实际取到的数据为准。
            row_count = len(values)
            col_count = max((len(row) for row in values), default=0)

        cleaned_rows = []
        for row in values:
            row_values = list(row)
            while row_values and row_values[-1] in (None, ""):
                row_values.pop()
            if any(v not in (None, "") for v in row_values):
                cleaned_rows.append(row_values)

        return {
            "success": True,
            "path": str(target),
            "name": target.name,
            "sheet_name": sheet_name,
            "sheet_count": len(names),
            "sheets": names,
            "rows_read": len(cleaned_rows),
            "source_rows": row_count,
            "source_cols": col_count,
            "max_rows": max_rows,
            "max_cols": max_cols,
            "data": cleaned_rows,
        }

    except Exception as e:
        return {
            "success": False,
            "error": f"读取 Excel 数据失败：{e}",
        }


def read_excel_file(
    path: str,
    sheet_name: str = "",
    max_rows: int = 100,
    max_cols: int = 30,
) -> dict:
    """
    给 Agent 用的统一 Excel 读取入口。
    .xls / .xlsx / .xlsm 自动选择解析方式。
    """
    return preview_excel(path, sheet_name, max_rows, max_cols)


def read_excel_range(path: str, sheet_name: str, cell_range: str) -> dict:
    target = _path(path)
    err = _check_excel(target)
    if err:
        return err

    if not sheet_name:
        return {"success": False, "error": "必须提供工作表名称。"}
    if not cell_range:
        return {"success": False, "error": "必须提供单元格区域，例如 A1:F20。"}

    try:
        if target.suffix.lower() == ".xls":
            # xlrd 没有统一的 A1 range API，这里做简单解析。
            import re
            match = re.fullmatch(
                r"([A-Za-z]+)(\d+):([A-Za-z]+)(\d+)",
                cell_range.strip(),
            )
            if not match:
                return {
                    "success": False,
                    "error": "当前 .xls 区域格式请使用例如 A1:F20。",
                }

            def col_to_num(s):
                n = 0
                for ch in s.upper():
                    n = n * 26 + (ord(ch) - 64)
                return n - 1

            c1, r1, c2, r2 = match.groups()
            c_start = col_to_num(c1)
            c_end = col_to_num(c2)
            r_start = int(r1) - 1
            r_end = int(r2)

            wb = _load_xls(target)
            if sheet_name not in wb.sheet_names():
                names = wb.sheet_names()
                wb.release_resources()
                return {
                    "success": False,
                    "error": f"找不到工作表：{sheet_name}",
                    "available_sheets": names,
                }

            sheet = wb.sheet_by_name(sheet_name)
            rows = []

            for r in range(r_start, min(r_end, sheet.nrows)):
                row = []
                for c in range(c_start, min(c_end + 1, sheet.ncols)):
                    row.append(sheet.cell_value(r, c))
                rows.append(row)

            wb.release_resources()

            return {
                "success": True,
                "path": str(target),
                "sheet_name": sheet_name,
                "range": cell_range,
                "rows": rows,
            }

        wb = _load_xlsx(target)
        sheet = wb.worksheets.get_item(sheet_name)
        values = sheet.get_range(cell_range).values

        return {
            "success": True,
            "path": str(target),
            "sheet_name": sheet_name,
            "range": cell_range,
            "rows": values,
        }

    except Exception as e:
        return {
            "success": False,
            "error": f"读取区域失败：{e}",
        }


def summarize_excel(path: str, sheet_name: str = "") -> dict:
    preview = preview_excel(
        path=path,
        sheet_name=sheet_name,
        max_rows=100,
        max_cols=30,
    )

    if not preview.get("success"):
        return preview

    data = preview.get("data", [])
    if not data:
        return {
            **preview,
            "summary": "工作表没有读取到数据。",
            "numeric_summary": [],
        }

    headers = []
    for idx, value in enumerate(data[0]):
        headers.append(
            str(value) if value not in (None, "") else f"列{idx + 1}"
        )

    numeric_summary = []

    for col_idx, header in enumerate(headers):
        nums = []

        for row in data[1:]:
            if col_idx >= len(row):
                continue

            value = row[col_idx]

            if isinstance(value, bool):
                continue

            if isinstance(value, (int, float)):
                nums.append(float(value))

        if nums:
            numeric_summary.append({
                "column": header,
                "count": len(nums),
                "sum": sum(nums),
                "average": sum(nums) / len(nums),
                "min": min(nums),
                "max": max(nums),
            })

    return {
        **preview,
        "summary": (
            f"工作表“{preview['sheet_name']}”读取到约 "
            f"{len(data)} 行、{max((len(r) for r in data), default=0)} 列的预览数据。"
        ),
        "numeric_summary": numeric_summary,
    }


def find_excel_keyword(
    path: str,
    keyword: str,
    sheet_name: str = "",
    max_rows: int = 500,
    max_cols: int = 50,
) -> dict:
    if not str(keyword or "").strip():
        return {"success": False, "error": "搜索关键词不能为空。"}

    preview = preview_excel(
        path=path,
        sheet_name=sheet_name,
        max_rows=max_rows,
        max_cols=max_cols,
    )

    if not preview.get("success"):
        return preview

    needle = str(keyword).lower()
    matches = []

    for row_index, row in enumerate(preview.get("data", []), start=1):
        row_text = " | ".join(
            "" if value is None else str(value) for value in row
        )

        if needle in row_text.lower():
            matches.append({
                "row": row_index,
                "values": row,
            })

    return {
        "success": True,
        "path": preview["path"],
        "sheet_name": preview["sheet_name"],
        "keyword": keyword,
        "count": len(matches),
        "matches": matches[:200],
    }


def read_excel_rows(
    path: str,
    sheet_name: str = "",
    max_rows: int = 100,
    max_cols: int = 30,
) -> dict:
    """兼容 Agent 常见工具名：读取 Excel 行数据。"""
    return read_excel_file(path, sheet_name, max_rows, max_cols)
