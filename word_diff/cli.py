#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""word_diff.cli —— 命令行入口：基础版本比较。

用法示例：
    python -m word_diff.cli v1.docx v2.docx          # 输出结构化 JSON
    python -m word_diff.cli v1.docx v2.docx --pretty # 缩进美化
    python -m word_diff.cli v1.docx v2.docx --summary# 只输出中文修改摘要
    python -m word_diff.cli v1.docx v2.docx --advanced # 额外执行进阶分析
"""
from __future__ import annotations

import argparse
import json
import sys
from typing import Any

from . import __version__
from .diff_engine import compute_diff, collect_paragraphs, DEFAULT_SIMILARITY_THRESHOLD
from .docx_reader import read_paragraph_entries, validate_docx_path
from .reporting import build_result, build_error, build_summary_text


def _dump(result: dict[str, Any], pretty: bool = False) -> None:
    """统一 JSON 输出（ensure_ascii=False 保留中文）。"""
    if pretty:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print(json.dumps(result, ensure_ascii=False))


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="word_diff",
        description="比较两份 Word 文档(.docx)的段落差异，输出结构化 JSON。",
        epilog="示例: python -m word_diff.cli v1.docx v2.docx --pretty",
    )
    parser.add_argument("doc1", help="第一份文档路径")
    parser.add_argument("doc2", help="第二份文档路径")
    parser.add_argument("--pretty", action="store_true", help="美化输出（缩进 JSON）")
    parser.add_argument(
        "--summary",
        action="store_true",
        help="仅输出中文修改摘要（低 Token 场景推荐）",
    )
    parser.add_argument(
        "--advanced",
        action="store_true",
        help="在基础比较后额外执行进阶分析（格式/图片/批注），默认关闭以节省 Token",
    )
    parser.add_argument(
        "--similarity",
        type=float,
        default=DEFAULT_SIMILARITY_THRESHOLD,
        help=f"修改/删除/新增 配对时的段落相似度阈值（默认 {DEFAULT_SIMILARITY_THRESHOLD}）",
    )
    parser.add_argument(
        "--keep-empty",
        action="store_true",
        help="保留空段落参与比较（默认忽略空段落，降低噪声）",
    )
    parser.add_argument("--version", action="version", version=f"word_diff {__version__}")
    return parser


def run(
    doc1: str,
    doc2: str,
    similarity_threshold: float = DEFAULT_SIMILARITY_THRESHOLD,
    skip_empty: bool = True,
    advanced: bool = False,
    summary_only: bool = False,
) -> dict[str, Any]:
    """主流程：校验 -> 读取 -> 比较 -> 组装。返回结果字典。

    拆成 run() 便于编程调用与单元测试（不经过 CLI 参数解析）。
    """
    # 1. 前置校验
    for path in (doc1, doc2):
        err = validate_docx_path(path)
        if err is not None:
            return err

    try:
        # 2. 读取段落
        entries1 = read_paragraph_entries(doc1, skip_empty=skip_empty)
        entries2 = read_paragraph_entries(doc2, skip_empty=skip_empty)

        # 3. 计算差异
        diff = compute_diff(entries1, entries2, similarity_threshold=similarity_threshold)

        # 4. 组装结果
        result = build_result(doc1, doc2, diff)

        # 5. 附上中文摘要（基础结果即包含，供低 Token 场景使用）
        result["summary_text"] = build_summary_text(doc1, doc2, diff)

        # 6. 可选：进阶分析
        if advanced:
            from .advanced import analyze_advanced  # 延迟导入，避免无谓开销

            adv_result = analyze_advanced(doc1, doc2)
            if adv_result.get("status") == "success":
                result["advanced"] = adv_result
            else:
                result["advanced_warning"] = adv_result

        # 7. --summary 只保留摘要与统计
        if summary_only:
            return {
                "status": "success",
                "is_identical": diff.is_identical,
                "summary": diff.summary(),
                "summary_text": result["summary_text"],
            }

        return result

    except Exception as exc:  # 兜底：任何运行时异常都转成结构化错误
        return build_error("runtime_error", f"处理文档时发生错误: {exc}")


def main(argv=None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    result = run(
        doc1=args.doc1,
        doc2=args.doc2,
        similarity_threshold=args.similarity,
        skip_empty=not args.keep_empty,
        advanced=args.advanced,
        summary_only=args.summary,
    )

    _dump(result, pretty=args.pretty)
    return 0 if result.get("status") == "success" else 1


if __name__ == "__main__":
    sys.exit(main())
