#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""word_diff —— Word 文档版本比较（Script/CLI + Skill）。

选题：编号 05 · AI Word 版本比较
基础任务：比较两个 Word 文档(.docx)的新增、删除、修改，输出结构化差异。
参考开源项目：python-docx
可选功能：生成修改摘要（见 reporting.build_summary_text）。

本包分层设计：
    diff_engine   纯逻辑 diff 核心（不依赖 python-docx，可独立测试）
    docx_reader   读取 .docx 的薄适配层（唯一依赖 python-docx 处）
    reporting     结果组装 / 摘要 / 输出
    cli           命令行入口
    advanced      进阶分析（格式 / 图片 / 批注），按需启用，节省 Token
"""
from .models import DiffItem, DiffResult, ParagraphEntry
from .diff_engine import compute_diff, collect_paragraphs, DEFAULT_SIMILARITY_THRESHOLD
from .reporting import build_result, build_error, build_summary_text
from .docx_reader import read_paragraph_entries, validate_docx_path

__all__ = [
    "DiffItem",
    "DiffResult",
    "ParagraphEntry",
    "compute_diff",
    "collect_paragraphs",
    "DEFAULT_SIMILARITY_THRESHOLD",
    "build_result",
    "build_error",
    "build_summary_text",
    "read_paragraph_entries",
    "validate_docx_path",
]

__version__ = "1.1.0"
