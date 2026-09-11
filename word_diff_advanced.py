#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Word 文档进阶比较入口（按需调用）。

用法：
    python word_diff_advanced.py v1.docx v2.docx            # 结构化 JSON
    python word_diff_advanced.py v1.docx v2.docx --pretty   # 缩进美化
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from word_diff.advanced import main  # noqa: E402

if __name__ == "__main__":
    sys.exit(main())
