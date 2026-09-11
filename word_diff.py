#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Word 文档版本比较入口（兼容旧用法）。

用法：
    python word_diff.py v1.docx v2.docx            # 结构化 JSON
    python word_diff.py v1.docx v2.docx --pretty   # 缩进美化
    python word_diff.py v1.docx v2.docx --summary  # 中文摘要（低 Token）
    python word_diff.py v1.docx v2.docx --advanced # 额外进阶分析

注意：本脚本与 `word_diff/` 包同名。为保证直接运行时能定位到包，
这里把脚本所在目录显式加入 sys.path。
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from word_diff.cli import main  # noqa: E402

if __name__ == "__main__":
    sys.exit(main())
