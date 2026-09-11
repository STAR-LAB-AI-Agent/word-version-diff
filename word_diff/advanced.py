#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""word_diff.advanced —— 进阶分析：格式 / 图片 / 批注 (按需调用)。

设计说明：
* 只在用户明确要求时启用，避免无条件把大文件解析结果塞进上下文（低 Token）。
* 图片内容差异：优先尝试真实哈希（读取图片二进制），不可读时给出明确说明，
  不再像此前那样伪造一个占位的 “MD5 = 无法提取…” 却仍参与比较。
* 批注比较：保留顺序/作者/时间，避免用 set 丢失重复批注与作者信息。
* 与基础脚本共用 docx_reader 的校验与错误结构。
"""
from __future__ import annotations

import json
import sys
from typing import Any

from .docx_reader import validate_docx_path
from .reporting import build_error


def _require_docx() -> None:
    """复用 docx_reader 的解析逻辑（含 vendor/ 兜底）。"""
    from .docx_reader import _require_docx as _req
    _req()


# ==================== 格式变化检测 ====================

def extract_run_format_info(paragraph) -> list[dict[str, Any]]:
    """提取一个段落内所有 Run 的格式信息。"""
    _require_docx()
    runs_info: list[dict[str, Any]] = []
    for run in paragraph.runs:
        font = run.font
        runs_info.append({
            "text": run.text,
            "font_name": font.name if font.name else "默认",
            "font_size": font.size.pt if font.size else None,
            "bold": font.bold,
            "italic": font.italic,
            "underline": font.underline,
        })
    return runs_info


def compare_format(doc1_path: str, doc2_path: str) -> dict[str, Any]:
    """比较两份文档格式变化（仅针对文本完全相同的段落，聚焦格式差异）。"""
    _require_docx()
    from docx import Document

    doc1 = Document(doc1_path)
    doc2 = Document(doc2_path)

    changes = []
    min_len = min(len(doc1.paragraphs), len(doc2.paragraphs))

    for idx in range(min_len):
        para1 = doc1.paragraphs[idx]
        para2 = doc2.paragraphs[idx]

        # 文本不同则跳过（关注点是“同文本不同格式”）
        if (para1.text or "").strip() != (para2.text or "").strip():
            continue

        runs1 = extract_run_format_info(para1)
        runs2 = extract_run_format_info(para2)

        if len(runs1) != len(runs2):
            changes.append({
                "paragraph": idx,
                "kind": "structure_change",
                "detail": f"格式片段数量由 {len(runs1)} 变为 {len(runs2)}",
            })
            continue

        for run_idx, (r1, r2) in enumerate(zip(runs1, runs2)):
            detail = []
            if r1["font_name"] != r2["font_name"]:
                detail.append(f"字体: {r1['font_name']} → {r2['font_name']}")
            if r1["font_size"] != r2["font_size"]:
                detail.append(f"字号: {r1['font_size']}pt → {r2['font_size']}pt")
            if r1["bold"] != r2["bold"]:
                detail.append(f"粗体: {r1['bold']} → {r2['bold']}")
            if r1["italic"] != r2["italic"]:
                detail.append(f"斜体: {r1['italic']} → {r2['italic']}")
            if r1["underline"] != r2["underline"]:
                detail.append(f"下划线: {r1['underline']} → {r2['underline']}")

            if detail:
                changes.append({
                    "paragraph": idx,
                    "run_index": run_idx,
                    "kind": "property_change",
                    "detail": detail,
                    "text": r1["text"],
                })

    return {"format_changes": changes, "count": len(changes),
            "note": "仅对比文本相同的段落；文本不同的段落不纳入格式比较"}


# ==================== 图片差异检测 ====================

def _image_blob(shape) -> bytes | None:
    """尽力从 inline_shape 提取图片二进制；读取不到返回 None。"""
    try:
        blob = shape._inline.graphic.graphicData.pic.blipFill.blip
        # python-docx 不直接暴露图片二进制，这里通过 relationship 读取
        rId = blob.embed if hasattr(blob, "embed") else None
        if rId is None:
            return None
        part = shape.part
        image_part = part.related_parts[rId]
        return image_part.blob
    except Exception:
        return None


def extract_image_info(doc_path: str) -> list[dict[str, Any]]:
    """提取文档中所有内嵌图片的信息（索引、尺寸、内容哈希）。"""
    _require_docx()
    from docx import Document
    from hashlib import md5

    doc = Document(doc_path)
    images: list[dict[str, Any]] = []

    for idx, shape in enumerate(doc.inline_shapes):
        try:
            width = shape.width.cm if hasattr(shape, "width") and shape.width else None
            height = shape.height.cm if hasattr(shape, "height") and shape.height else None

            blob = _image_blob(shape)
            digest = md5(blob).hexdigest() if blob else None
            size = len(blob) if blob else None

            images.append({
                "index": idx,
                "width_cm": round(width, 2) if width else None,
                "height_cm": round(height, 2) if height else None,
                "md5": digest,
                "size_bytes": size,
                "hashed": blob is not None,
            })
        except Exception as exc:  # 单个图片失败不影响整体
            images.append({
                "index": idx,
                "error": str(exc)[:60],
            })

    return images


def compare_images(doc1_path: str, doc2_path: str) -> dict[str, Any]:
    """比较两份文档的图片数量与内容差异。"""
    images1 = extract_image_info(doc1_path)
    images2 = extract_image_info(doc2_path)

    diff = []
    if len(images1) != len(images2):
        diff.append({"kind": "count_change", "detail": f"图片数量 {len(images1)} → {len(images2)}"})

    for idx in range(min(len(images1), len(images2))):
        img1, img2 = images1[idx], images2[idx]
        detail = []

        # 尺寸
        w1, w2 = img1.get("width_cm"), img2.get("width_cm")
        h1, h2 = img1.get("height_cm"), img2.get("height_cm")
        if w1 and w2 and abs(w1 - w2) > 0.1:
            detail.append(f"宽度: {w1}cm → {w2}cm")
        if h1 and h2 and abs(h1 - h2) > 0.1:
            detail.append(f"高度: {h1}cm → {h2}cm")

        # 内容：仅当两边都成功哈希时才可能做比较
        m1, m2 = img1.get("md5"), img2.get("md5")
        if m1 and m2 and m1 != m2:
            detail.append("图片内容变化（MD5 不同）")

        if detail:
            diff.append({"image_index": idx, "detail": detail})

    return {
        "doc1_count": len(images1),
        "doc2_count": len(images2),
        "image_diff": diff,
        "changed": len(diff) > 0 or len(images1) != len(images2),
    }


# ==================== 批注差异检测 ====================

def extract_comments_info(doc_path: str) -> list[dict[str, Any]]:
    """提取文档中所有批注（作者、内容、缩写，保留顺序）。"""
    _require_docx()
    from docx import Document

    doc = Document(doc_path)
    comments: list[dict[str, Any]] = []
    for comment in doc.comments:
        text = comment.text or ""
        comments.append({
            "author": getattr(comment, "author", "未知"),
            "initials": getattr(comment, "initials", None),
            "text": text,
            "truncated": len(text) > 100,
        })
    return comments


def compare_comments(doc1_path: str, doc2_path: str) -> dict[str, Any]:
    """比较两份文档的批注差异（顺序与作者敏感，改用逐条对比而非 set）。"""
    comments1 = extract_comments_info(doc1_path)
    comments2 = extract_comments_info(doc2_path)

    c1 = [(c["author"], c["text"]) for c in comments1]
    c2 = [(c["author"], c["text"]) for c in comments2]

    diff: list[dict[str, Any]] = []
    if len(comments1) != len(comments2):
        diff.append({"kind": "count_change", "detail": f"批注数量 {len(comments1)} → {len(comments2)}"})

    # 简单的前缀/追加式比较：找公共前段与公共后段
    prefix = 0
    max_prefix = min(len(c1), len(c2))
    while prefix < max_prefix and c1[prefix] == c2[prefix]:
        prefix += 1

    suffix = 0
    max_suffix = min(len(c1), len(c2)) - prefix
    while suffix < max_suffix and c1[len(c1) - 1 - suffix] == c2[len(c2) - 1 - suffix]:
        suffix += 1

    added = c2[prefix : len(c2) - suffix if suffix else len(c2)]
    removed = c1[prefix : len(c1) - suffix if suffix else len(c1)]

    if removed:
        diff.append({"kind": "removed", "detail": [f"{a}: {t[:60]}" for a, t in removed]})
    if added:
        diff.append({"kind": "added", "detail": [f"{a}: {t[:60]}" for a, t in added]})

    return {
        "doc1_count": len(comments1),
        "doc2_count": len(comments2),
        "comment_diff": diff,
        "changed": len(diff) > 0 or len(comments1) != len(comments2),
        "note": "按(作者, 内容)逐条对比；批注时间信息在 python-docx 中不可用",
    }


# ==================== 主流程 ====================

def analyze_advanced(doc1_path: str, doc2_path: str) -> dict[str, Any]:
    """整合格式 / 图片 / 批注三项进阶检测。"""
    for path in (doc1_path, doc2_path):
        err = validate_docx_path(path)
        if err is not None:
            return err

    try:
        format_result = compare_format(doc1_path, doc2_path)
        image_result = compare_images(doc1_path, doc2_path)
        comment_result = compare_comments(doc1_path, doc2_path)

        total = (
            format_result["count"]
            + len(image_result["image_diff"])
            + len(comment_result["comment_diff"])
        )
        has_any = (
            format_result["count"] > 0
            or image_result["changed"]
            or comment_result["changed"]
        )

        return {
            "status": "success",
            "has_any_advanced_change": has_any,
            "total_changes": total,
            "format": format_result,
            "images": image_result,
            "comments": comment_result,
            "summary_text": (
                f"进阶分析：格式变化 {format_result['count']} 处，"
                f"图片差异 {len(image_result['image_diff'])} 处，"
                f"批注差异 {len(comment_result['comment_diff'])} 处"
            ),
        }
    except Exception as exc:
        return build_error("runtime_error", f"进阶分析失败: {exc}")


def main(argv=None) -> int:
    """进阶 CLI 入口（可独立运行测试），统一用 argparse 而非手动解析。"""
    import argparse

    parser = argparse.ArgumentParser(
        prog="word_diff_advanced",
        description="进阶比较两份 Word 文档的格式 / 图片 / 批注差异。",
        epilog="示例: python -m word_diff.advanced a.docx b.docx --pretty",
    )
    parser.add_argument("doc1", help="第一份文档路径")
    parser.add_argument("doc2", help="第二份文档路径")
    parser.add_argument("--pretty", action="store_true", help="美化输出")

    args = parser.parse_args(argv)
    result = analyze_advanced(args.doc1, args.doc2)
    if args.pretty:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print(json.dumps(result, ensure_ascii=False))
    return 0 if result.get("status") == "success" else 1


if __name__ == "__main__":
    sys.exit(main())
