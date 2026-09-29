# -*- coding: utf-8 -*-
"""
饭加鱼：PDF 文档技能 v1

支持：
- 读取 PDF 页数
- 提取 PDF 文本
- 获取 PDF 基本信息
- 按页返回文本
- 只读，不修改原 PDF
"""

from pathlib import Path


def _path(path: str) -> Path:
    if not path or not str(path).strip():
        raise ValueError("PDF 路径不能为空。")

    value = str(path).strip().strip('"')
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


def _load_backend():
    """
    优先使用 PyMuPDF（fitz），失败则尝试 pypdf。
    返回 (backend_name, module)。
    """
    try:
        import fitz
        return "pymupdf", fitz
    except ImportError:
        pass

    try:
        import pypdf
        return "pypdf", pypdf
    except ImportError:
        pass

    return None, None


def _read_with_pymupdf(path: Path, max_chars: int):
    import fitz

    doc = fitz.open(str(path))

    pages = []
    all_parts = []

    for index, page in enumerate(doc):
        text = page.get_text("text") or ""
        text = text.strip()

        page_data = {
            "page": index + 1,
            "text": text,
            "char_count": len(text),
        }
        pages.append(page_data)

        if text:
            all_parts.append(f"[第 {index + 1} 页]\n{text}")

    content = "\n\n".join(all_parts)

    metadata = {}
    try:
        metadata = dict(doc.metadata or {})
    except Exception:
        metadata = {}

    result = {
        "success": True,
        "path": str(path),
        "name": path.name,
        "type": "pdf",
        "backend": "pymupdf",
        "page_count": len(doc),
        "metadata": metadata,
        "pages": pages,
        "content": content[:max_chars],
        "truncated": len(content) > max_chars,
    }

    doc.close()
    return result


def _read_with_pypdf(path: Path, max_chars: int):
    from pypdf import PdfReader

    reader = PdfReader(str(path))
    pages = []
    all_parts = []

    for index, page in enumerate(reader.pages):
        text = page.extract_text() or ""
        text = text.strip()

        pages.append({
            "page": index + 1,
            "text": text,
            "char_count": len(text),
        })

        if text:
            all_parts.append(f"[第 {index + 1} 页]\n{text}")

    content = "\n\n".join(all_parts)

    metadata = {}
    try:
        raw = reader.metadata
        metadata = {str(k): str(v) for k, v in (raw or {}).items()}
    except Exception:
        metadata = {}

    return {
        "success": True,
        "path": str(path),
        "name": path.name,
        "type": "pdf",
        "backend": "pypdf",
        "page_count": len(reader.pages),
        "metadata": metadata,
        "pages": pages,
        "content": content[:max_chars],
        "truncated": len(content) > max_chars,
    }


def read_pdf(path: str, max_chars: int = 40000) -> dict:
    """
    读取 PDF 文本。
    注意：扫描件/纯图片 PDF 可能没有可提取文字，需要后续增加 OCR 技能。
    """
    target = _path(path)

    if not target.exists():
        return {"success": False, "error": f"PDF 文件不存在：{target}"}

    if not target.is_file():
        return {"success": False, "error": f"这不是文件：{target}"}

    if target.suffix.lower() != ".pdf":
        return {"success": False, "error": "当前技能只接受 .pdf 文件。"}

    backend, module = _load_backend()

    if backend is None:
        return {
            "success": False,
            "error": (
                "当前 Python 环境没有安装 PDF 解析库。\n"
                "请运行：python -m pip install pymupdf"
            ),
        }

    try:
        max_chars = max(1000, min(int(max_chars), 80000))
    except (TypeError, ValueError):
        max_chars = 40000

    try:
        if backend == "pymupdf":
            return _read_with_pymupdf(target, max_chars)

        return _read_with_pypdf(target, max_chars)

    except Exception as e:
        return {
            "success": False,
            "error": f"读取 PDF 失败：{e}",
        }


def get_pdf_info(path: str) -> dict:
    target = _path(path)

    if not target.exists():
        return {"success": False, "error": f"PDF 文件不存在：{target}"}

    if target.suffix.lower() != ".pdf":
        return {"success": False, "error": "当前技能只接受 .pdf 文件。"}

    backend, _ = _load_backend()

    if backend is None:
        return {
            "success": False,
            "error": "没有安装 PDF 解析库，请运行：python -m pip install pymupdf",
        }

    try:
        if backend == "pymupdf":
            import fitz
            doc = fitz.open(str(target))
            info = {
                "success": True,
                "path": str(target),
                "name": target.name,
                "type": "pdf",
                "backend": "pymupdf",
                "page_count": len(doc),
                "metadata": dict(doc.metadata or {}),
                "size": target.stat().st_size,
            }
            doc.close()
            return info

        from pypdf import PdfReader
        reader = PdfReader(str(target))
        return {
            "success": True,
            "path": str(target),
            "name": target.name,
            "type": "pdf",
            "backend": "pypdf",
            "page_count": len(reader.pages),
            "metadata": {str(k): str(v) for k, v in (reader.metadata or {}).items()},
            "size": target.stat().st_size,
        }

    except Exception as e:
        return {"success": False, "error": f"读取 PDF 信息失败：{e}"}
