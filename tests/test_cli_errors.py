# -*- coding: utf-8 -*-
"""第 5 层：异常与边界（对应任务书"安全与异常处理"得分点）。

断言错误返回的**统一契约**：
    {"status": "error", "error_type": "file_not_found|invalid_format|runtime_error", "message": "..."}
错误场景的退出码统一为 1，且必须是可解析的 JSON（而不是抛栈），
这样智能体才能把失败原因转述给用户。
"""
import tempfile
from pathlib import Path

from helpers import make_docx, run_cli, skip_if_no_docx

ERROR_KEYS = {"status", "error_type", "message"}


def _assert_error(result, code, expected_type):
    assert isinstance(result, dict), "错误场景也应输出可解析的 JSON"
    assert result.get("status") == "error"
    assert ERROR_KEYS <= set(result), f"错误返回缺少字段 {sorted(ERROR_KEYS - set(result))}"
    assert result["error_type"] == expected_type, f"error_type 应为 {expected_type}，实际 {result['error_type']!r}"
    assert result["message"], "错误信息不应为空（要能解释给用户听）"
    assert code == 1, f"错误场景退出码应为 1，实际 {code}"


def test_missing_file_returns_file_not_found():
    with tempfile.TemporaryDirectory() as tmp:
        code, result, _ = run_cli(
            [str(Path(tmp) / "没有这个文件.docx"), str(Path(tmp) / "另一个也不存在.docx")]
        )
    _assert_error(result, code, "file_not_found")


def test_directory_as_argument_returns_invalid_format():
    with tempfile.TemporaryDirectory() as tmp:
        code, result, _ = run_cli([tmp, tmp])
    _assert_error(result, code, "invalid_format")


def test_wrong_suffix_returns_invalid_format():
    with tempfile.TemporaryDirectory() as tmp:
        note = Path(tmp) / "笔记.txt"
        note.write_text("hello", encoding="utf-8")
        code, result, _ = run_cli([str(note), str(note)])
    _assert_error(result, code, "invalid_format")


def test_corrupted_docx_returns_runtime_error():
    """扩展名对、但内容不是合法 docx（例如改名来的文本文件）。"""
    skip_if_no_docx()
    with tempfile.TemporaryDirectory() as tmp:
        broken = Path(tmp) / "broken.docx"
        broken.write_text("这不是一个 zip 包，更不是 docx", encoding="utf-8")
        code, result, _ = run_cli([str(broken), str(broken)])
    _assert_error(result, code, "runtime_error")


def test_missing_argument_exits_with_code_2():
    """参数不足时由 argparse 直接报错退出（退出码 2），不应输出半截结果。"""
    with tempfile.TemporaryDirectory() as tmp:
        only_one = make_docx(Path(tmp) / "single.docx", ["正文"])
        code, result, out = run_cli([only_one])
    assert code == 2, f"参数不足应返回 2，实际 {code}"
    assert result is None


def test_same_file_twice_is_not_an_error():
    skip_if_no_docx()
    with tempfile.TemporaryDirectory() as tmp:
        path = make_docx(Path(tmp) / "a.docx", ["标题", "正文"])
        code, result, _ = run_cli([path, path, "--summary"])
    assert code == 0
    assert result["status"] == "success" and result["is_identical"] is True


def test_empty_documents_are_identical():
    """两份都没有任何段落的文档：正常结论"一致"，不是报错。"""
    skip_if_no_docx()
    with tempfile.TemporaryDirectory() as tmp:
        a = make_docx(Path(tmp) / "empty_a.docx", [])
        b = make_docx(Path(tmp) / "empty_b.docx", [])
        code, result, _ = run_cli([a, b, "--summary"])
    assert code == 0 and result["is_identical"] is True


def test_advanced_missing_file_returns_error():
    """进阶分析走同一套错误契约。"""
    from word_diff.advanced import analyze_advanced

    with tempfile.TemporaryDirectory() as tmp:
        result = analyze_advanced(
            str(Path(tmp) / "不存在1.docx"), str(Path(tmp) / "不存在2.docx")
        )
    assert result["status"] == "error"
    assert result["error_type"] == "file_not_found"
