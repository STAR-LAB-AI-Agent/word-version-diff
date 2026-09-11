#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""word_diff.reporting —— 结果组装、摘要生成与输出。

负责把 DiffResult 转成稳定、可程序化消费的 JSON 字典，
并提供中文摘要文本（便于智能体直接返回给用户）。
"""
from __future__ import annotations

from typing import Any

from .models import DiffResult
from .docx_reader import read_doc_meta


def build_result(
    doc1_path: str,
    doc2_path: str,
    diff: DiffResult,
) -> dict[str, Any]:
    """组装基础比较的完整结果字典。"""
    doc1 = read_doc_meta(doc1_path)
    doc2 = read_doc_meta(doc2_path)
    return {
        "status": "success",
        "doc1": doc1,
        "doc2": doc2,
        "is_identical": diff.is_identical,
        "summary": diff.summary(),
        "differences": diff.to_dict(),
    }


def build_error(error_type: str, message: str) -> dict[str, Any]:
    """统一错误结果结构。"""
    return {
        "status": "error",
        "error_type": error_type,
        "message": message,
    }


def build_summary_text(doc1_path: str, doc2_path: str, diff: DiffResult) -> str:
    """生成人类可读的中文修改摘要。

    title 段：
        “修改摘要”：共 N 处变动（新增 A，删除 B，修改 C）。
        “文档《X》与《Y》的差异统计：…”，并列示前若干处内容。
    """
    s = diff.summary()
    total = s["total"]
    header = (
        f"对比《{Path_name(doc1_path)}》与《{Path_name(doc2_path)}》"
        f"共发现 {total} 处变动（新增 {s['added']}，删除 {s['removed']}，修改 {s['modified']}）。"
    )
    if total == 0:
        return header + "两份文档内容一致。"

    lines: list[str] = [header, ""]

    def _clip(text: str | None, n: int = 60) -> str:
        if not text:
            return ""
        text = text.replace("\n", " ").strip()
        return text if len(text) <= n else text[: n - 1] + "…"

    for item in diff.modified[:10]:
        lines.append(f"- 修改（第{item.position + 1}段）：{_clip(item.old_text)} → {_clip(item.new_text)}")
    for item in diff.added[:10]:
        lines.append(f"- 新增（第{item.position + 1}段）：{_clip(item.new_text)}")
    for item in diff.removed[:10]:
        lines.append(f"- 删除（第{item.position + 1}段）：{_clip(item.old_text)}")

    if s["total"] > 30:
        lines.append("")
        lines.append("（内容较多，仅展示前若干处，完整差异见 differences 字段。）")

    return "\n".join(lines)


def Path_name(path: str) -> str:
    """提取文件名字，供中文摘要使用。"""
    import os
    return os.path.basename(path)
