#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""word_diff.docx_reader —— 读取 .docx 文档的薄适配层。

本模块是唯一直接依赖 python-docx 的地方，负责把“外部 .docx 文件”
转换成内部无依赖的文本/段落结构，供 diff_engine 与 reporting 使用。
这样做的好处：核心逻辑可独立测试，替换 / 扩展解析器只改这里。
"""
from __future__ import annotations

import os
import sys
from collections.abc import Sequence
from pathlib import Path

from .models import ParagraphEntry


def _vendor_site() -> str | None:
    """返回项目根下 vendor/ 目录的绝对路径（若存在）。

    若全局未安装 python-docx，可自动降级到 vendor/（由 bootstrap_env.py 生成），
    从而让 word_diff 在本环境开箱即用。
    """
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    vendor = os.path.join(root, "vendor")
    return vendor if os.path.isdir(vendor) else None


def _require_docx() -> None:
    """解析并返回 docx 模块；缺失时尝试 vendor/ 兜底，仍无则报友好错误。"""
    try:
        import docx  # noqa: F401
        return
    except ImportError:
        pass

    vendor = _vendor_site()
    if vendor is not None:
        sys.path.insert(0, vendor)
        try:
            import docx  # noqa: F401
            return
        except ImportError:
            pass

    raise RuntimeError(
        "缺少 python-docx 依赖。可执行:  pip install python-docx "
        "或  python tests/bootstrap_env.py"
    )


def validate_docx_path(path: str) -> dict[str, object] | None:
    """校验路径是否为合法存在的 .docx 文件。

    返回 None 表示校验通过；否则返回 {status, error_type, message}。
    """
    p = Path(path)
    if not p.exists():
        return {
            "status": "error",
            "error_type": "file_not_found",
            "message": f"文件不存在: {path}",
        }
    if not p.is_file():
        return {
            "status": "error",
            "error_type": "invalid_format",
            "message": f"不是文件: {path}",
        }
    if p.suffix.lower() != ".docx":
        return {
            "status": "error",
            "error_type": "invalid_format",
            "message": f"仅支持 .docx 格式，当前文件: {p.name}",
        }
    return None


def read_paragraph_entries(path: str, skip_empty: bool = True) -> list[ParagraphEntry]:
    """从 .docx 中读取所有段落，返回 ParagraphEntry 列表。

    所有段落都会被读取（包括空段落），以便 position 与原始段落号一致；
    是否跳过空段落由 skip_empty 决定（在 diff 引擎层面生效）。
    """
    _require_docx()
    from docx import Document

    doc = Document(path)
    entries: list[ParagraphEntry] = []
    for idx, para in enumerate(doc.paragraphs):
        text = (para.text or "").strip()
        if skip_empty and not text:
            continue
        entries.append(ParagraphEntry(index=idx, text=text))
    return entries


def read_all_paragraphs(path: str) -> list[str]:
    """读取所有段落的原始文本（含空段），供进阶分析或调试使用。"""
    _require_docx()
    from docx import Document

    doc = Document(path)
    return [(p.text or "") for p in doc.paragraphs]


def read_doc_meta(path: str) -> dict[str, object]:
    """读取文档元信息（名称、段落总数）。"""
    _require_docx()
    from docx import Document

    doc = Document(path)
    name = Path(path).name
    return {"path": path, "name": name, "paragraph_count": len(doc.paragraphs)}
