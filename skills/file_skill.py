# -*- coding: utf-8 -*-
"""
饭加鱼：文件处理技能 v3
"""

from pathlib import Path
import shutil
from datetime import datetime


def _path(path: str) -> Path:
    if not path or not str(path).strip():
        raise ValueError("文件路径不能为空。")

    value = str(path).strip().strip('"')
    home = Path.home()

    special_dirs = {
        "桌面": home / "Desktop",
        "desktop": home / "Desktop",
        "桌面文件夹": home / "Desktop",
        "下载": home / "Downloads",
        "downloads": home / "Downloads",
        "下载文件夹": home / "Downloads",
        "文档": home / "Documents",
        "documents": home / "Documents",
        "文档文件夹": home / "Documents",
    }

    if value in special_dirs:
        return special_dirs[value].resolve()

    return Path(value).expanduser().resolve()


def _human_size(size):
    """把字节数转成人类可读的大小。"""
    if size is None:
        return None
    size = float(size)
    for unit in ["B", "KB", "MB", "GB", "TB"]:
        if size < 1024:
            return f"{size:.1f} {unit}"
        size /= 1024
    return f"{size:.1f} PB"


def list_files(folder: str, recursive: bool = False) -> dict:
    root = _path(folder)
    if not root.exists():
        return {"success": False, "error": f"文件夹不存在：{root}"}
    if not root.is_dir():
        return {"success": False, "error": f"不是文件夹：{root}"}

    pattern = "**/*" if recursive else "*"
    items = []

    try:
        for item in root.glob(pattern):
            try:
                item_type = "folder" if item.is_dir() else "file"
                size = item.stat().st_size if item.is_file() else None
                items.append({
                    "name": item.name,
                    "path": str(item),
                    "type": item_type,
                    "size": size,
                    "size_readable": _human_size(size),
                })
            except (OSError, PermissionError):
                continue
    except (OSError, PermissionError) as e:
        return {"success": False, "error": f"无法读取文件夹：{e}"}

    items.sort(key=lambda x: (x["type"] != "folder", x["name"].lower()))
    return {
        "success": True,
        "folder": str(root),
        "count": len(items),
        "items": items,
    }


def search_files(folder: str, keyword: str, recursive: bool = True) -> dict:
    root = _path(folder)
    if not root.exists():
        return {"success": False, "error": f"文件夹不存在：{root}"}
    if not root.is_dir():
        return {"success": False, "error": f"不是文件夹：{root}"}

    keyword = str(keyword or "").strip().lower()
    if not keyword:
        return {"success": False, "error": "搜索关键词不能为空。"}

    pattern = "**/*" if recursive else "*"
    matches = []

    try:
        for item in root.glob(pattern):
            try:
                if item.is_file() and keyword in item.name.lower():
                    matches.append(str(item))
            except (OSError, PermissionError):
                continue
    except (OSError, PermissionError) as e:
        return {"success": False, "error": f"搜索文件时出错：{e}"}

    matches.sort(key=str.lower)
    return {
        "success": True,
        "folder": str(root),
        "keyword": keyword,
        "count": len(matches),
        "files": matches,
    }


def create_folder(folder: str) -> dict:
    target = _path(folder)
    try:
        target.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        return {"success": False, "error": f"创建文件夹失败：{e}"}
    return {
        "success": True,
        "path": str(target),
        "message": "文件夹已准备好。",
    }


def copy_file(source: str, destination: str) -> dict:
    src = _path(source)
    dst = _path(destination)

    if not src.exists():
        return {"success": False, "error": f"源文件不存在：{src}"}
    if not src.is_file():
        return {"success": False, "error": f"源路径不是文件：{src}"}

    if dst.exists() and dst.is_dir():
        dst = dst / src.name
    else:
        dst.parent.mkdir(parents=True, exist_ok=True)

    if dst.exists():
        return {
            "success": False,
            "error": f"目标文件已存在，为安全起见没有覆盖：{dst}",
        }

    try:
        shutil.copy2(src, dst)
    except Exception as e:
        return {"success": False, "error": f"复制失败：{e}"}

    return {"success": True, "source": str(src), "destination": str(dst)}


def move_file(source: str, destination: str) -> dict:
    src = _path(source)
    dst = _path(destination)

    if not src.exists():
        return {"success": False, "error": f"源路径不存在：{src}"}
    if not src.is_file():
        return {"success": False, "error": f"当前版本只允许移动文件：{src}"}

    if dst.exists() and dst.is_dir():
        dst = dst / src.name
    else:
        dst.parent.mkdir(parents=True, exist_ok=True)

    if dst.exists():
        return {
            "success": False,
            "error": f"目标文件已存在，为安全起见没有覆盖：{dst}",
        }

    try:
        shutil.move(str(src), str(dst))
    except Exception as e:
        return {"success": False, "error": f"移动失败：{e}"}

    return {"success": True, "source": str(src), "destination": str(dst)}


def rename_file(path: str, new_name: str) -> dict:
    src = _path(path)

    if not src.exists():
        return {"success": False, "error": f"文件不存在：{src}"}
    if not src.is_file():
        return {"success": False, "error": f"当前版本只允许重命名文件：{src}"}

    new_name = str(new_name or "").strip()
    if not new_name:
        return {"success": False, "error": "新文件名不能为空。"}
    if Path(new_name).name != new_name:
        return {"success": False, "error": "新名称只能是文件名，不能包含路径。"}

    target = src.with_name(new_name)
    if target.exists():
        return {
            "success": False,
            "error": f"目标文件已存在，为安全起见没有覆盖：{target}",
        }

    try:
        src.rename(target)
    except Exception as e:
        return {"success": False, "error": f"重命名失败：{e}"}

    return {"success": True, "old_path": str(src), "new_path": str(target)}


def get_file_info(path: str) -> dict:
    target = _path(path)
    if not target.exists():
        return {"success": False, "error": f"路径不存在：{target}"}

    try:
        stat = target.stat()
    except Exception as e:
        return {"success": False, "error": f"无法读取文件信息：{e}"}

    is_file = target.is_file()
    size = stat.st_size if is_file else None

    try:
        modified_time = datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        modified_time = None

    return {
        "success": True,
        "name": target.name,
        "path": str(target),
        "type": "folder" if target.is_dir() else "file",
        "size": size,
        "size_readable": _human_size(size),
        "suffix": target.suffix if is_file else None,
        "modified_time": modified_time,
    }


def delete_file(path: str, confirmed: bool = False) -> dict:
    if not confirmed:
        return {
            "success": False,
            "requires_confirmation": True,
            "error": "删除文件属于高风险操作，请先明确确认。",
        }

    target = _path(path)
    if not target.exists():
        return {"success": False, "error": f"路径不存在：{target}"}
    if not target.is_file():
        return {"success": False, "error": f"当前版本只允许删除文件：{target}"}

    try:
        target.unlink()
    except Exception as e:
        return {"success": False, "error": f"删除失败：{e}"}

    return {"success": True, "deleted": str(target)}