# -*- coding: utf-8 -*-
"""用**你自己的真实文档**跑一遍（可选、可移植，不把机器路径写进代码）。

用法（PowerShell 举例）：
    $env:WORD_DIFF_REAL_DOC1="D:\\...\\v1.docx"
    $env:WORD_DIFF_REAL_DOC2="D:\\...\\v2.docx"
    python tests/run_tests.py real_documents

未设置环境变量（或文件不存在）时自动 **跳过**，所以换设备/换目录都不会变红。
"""
import os
import unittest
from pathlib import Path

from helpers import run_advanced, run_cli, skip_if_no_docx

ENV_DOC1 = "WORD_DIFF_REAL_DOC1"
ENV_DOC2 = "WORD_DIFF_REAL_DOC2"


def _real_pair() -> tuple[str, str]:
    skip_if_no_docx()
    doc1, doc2 = os.environ.get(ENV_DOC1), os.environ.get(ENV_DOC2)
    if not (doc1 and doc2):
        raise unittest.SkipTest(f"未设置 {ENV_DOC1} / {ENV_DOC2}，跳过真实文档用例")
    for path in (doc1, doc2):
        if not Path(path).is_file():
            raise unittest.SkipTest(f"文件不存在，跳过：{path}")
    return doc1, doc2


def test_real_documents_basic_contract():
    doc1, doc2 = _real_pair()
    code, data, out = run_cli([doc1, doc2])
    assert code == 0
    assert data["status"] == "success"
    assert {"added", "removed", "modified"} <= set(data["differences"])
    counted = sum(len(v) for v in data["differences"].values())
    assert data["summary"]["total"] == counted


def test_real_documents_advanced_contract():
    doc1, doc2 = _real_pair()
    result = run_advanced(doc1, doc2)
    assert result["status"] == "success"
    assert {"format", "images", "comments", "blank_paragraphs"} <= set(result)
    assert isinstance(result["summary_text"], str) and result["summary_text"]


def test_real_documents_summary_only_is_low_token():
    doc1, doc2 = _real_pair()
    code, data, _ = run_cli([doc1, doc2, "--summary"])
    assert code == 0
    assert set(data) == {"status", "is_identical", "summary", "summary_text"}
