# -*- coding: utf-8 -*-
"""第 2 层：读取层与路径校验（**需要 python-docx**，缺失时自动跳过）。

被测对象：`word_diff.docx_reader`
覆盖：空段落过滤与段落号保留、元信息、以及路径校验的四种异常分类
（不存在 / 不是文件 / 后缀不对 / 正常）。
"""
import tempfile
from pathlib import Path

from helpers import make_docx, skip_if_no_docx

from word_diff.docx_reader import (
    read_all_paragraphs,
    read_doc_meta,
    read_paragraph_entries,
    validate_docx_path,
)


def test_read_paragraph_entries_skips_empty_but_keeps_index():
    skip_if_no_docx()
    with tempfile.TemporaryDirectory() as tmp:
        path = make_docx(Path(tmp) / "a.docx", ["标题", "", "正文"])
        entries = read_paragraph_entries(path)
        assert [e.text for e in entries] == ["标题", "正文"]
        assert [e.index for e in entries] == [0, 2]  # 段落号仍与原文档一致


def test_read_paragraph_entries_can_keep_empty():
    skip_if_no_docx()
    with tempfile.TemporaryDirectory() as tmp:
        path = make_docx(Path(tmp) / "a.docx", ["标题", "", "正文"])
        entries = read_paragraph_entries(path, skip_empty=False)
        assert [e.index for e in entries] == [0, 1, 2]


def test_read_all_paragraphs_keeps_everything():
    skip_if_no_docx()
    with tempfile.TemporaryDirectory() as tmp:
        path = make_docx(Path(tmp) / "a.docx", ["标题", "", "正文"])
        assert read_all_paragraphs(path) == ["标题", "", "正文"]


def test_read_doc_meta_reports_name_and_count():
    skip_if_no_docx()
    with tempfile.TemporaryDirectory() as tmp:
        path = make_docx(Path(tmp) / "meta.docx", ["一", "二", ""])
        meta = read_doc_meta(path)
        assert meta["name"] == "meta.docx"
        assert meta["paragraph_count"] == 3


def test_validate_docx_path_accepts_real_file():
    skip_if_no_docx()
    with tempfile.TemporaryDirectory() as tmp:
        path = make_docx(Path(tmp) / "ok.docx", ["正文"])
        assert validate_docx_path(path) is None


def test_validate_docx_path_rejects_missing_file():
    with tempfile.TemporaryDirectory() as tmp:
        err = validate_docx_path(str(Path(tmp) / "不存在.docx"))
        assert err is not None and err["error_type"] == "file_not_found"


def test_validate_docx_path_rejects_directory():
    with tempfile.TemporaryDirectory() as tmp:
        err = validate_docx_path(tmp)
        assert err is not None and err["error_type"] == "invalid_format"


def test_validate_docx_path_rejects_wrong_suffix():
    with tempfile.TemporaryDirectory() as tmp:
        p = Path(tmp) / "note.txt"
        p.write_text("hello", encoding="utf-8")
        err = validate_docx_path(str(p))
        assert err is not None and err["error_type"] == "invalid_format"
        assert ".docx" in err["message"]
