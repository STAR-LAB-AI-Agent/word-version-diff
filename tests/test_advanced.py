# -*- coding: utf-8 -*-
"""第 4 层：进阶分析（格式 / 图片 / 批注 / 空段落），需要 python-docx。

这里的用例多数是**回归测试**，对应修复前的两个已知假阴性/假阳性：
* 段落按序号对齐 → 插入段落会让其后所有段落的格式比较被整段跳过（漏报）；
* run 数量不等就直接放弃属性比较（既漏报也误报）。
"""
import tempfile
from pathlib import Path

from helpers import make_docx, make_png, run_advanced, skip_if_no_docx


def _pair(tmp, blocks_a, blocks_b, picture_a=None, picture_b=None):
    a = make_docx(Path(tmp) / "a.docx", blocks_a, picture=picture_a)
    b = make_docx(Path(tmp) / "b.docx", blocks_b, picture=picture_b)
    return run_advanced(a, b)


# ------------------------------------------------------------------ 格式

def test_format_change_detected_with_char_range():
    """粗体/字号变化要被检出，并给出逐字符范围。"""
    skip_if_no_docx()
    with tempfile.TemporaryDirectory() as tmp:
        res = _pair(
            tmp,
            [{"runs": [("标题", {"bold": True, "size": 14})]}, "正文"],
            [{"runs": [("标题", {"bold": False, "size": 12})]}, "正文"],
        )
        assert res["status"] == "success"
        assert res["format"]["count"] >= 1
        change = res["format"]["format_changes"][0]
        assert change["char_range"] and len(change["char_range"]) == 2
        assert any("粗体" in d for d in change["detail"])


def test_alignment_survives_inserted_paragraph():
    """回归：在文档开头插入一段后，后续段落的格式差异**仍要**被检出。

    （按段落序号对齐的旧实现会在这里漏报，因为第 0 段文本已经不同。）
    """
    skip_if_no_docx()
    with tempfile.TemporaryDirectory() as tmp:
        res = _pair(
            tmp,
            [
                {"runs": [("标题", {"bold": True})]},
                "正文A",
                "正文B",
            ],
            [
                "新插入的段落",
                {"runs": [("标题", {"bold": False})]},
                "正文A",
                "正文B",
            ],
        )
        assert res["format"]["count"] == 1, "插入段落不应导致后续段落格式比较被跳过"


def test_run_split_does_not_create_format_change():
    """回归：run 切分不同（1 个 run vs 2 个 run）但渲染一致 → 不算格式变化。"""
    skip_if_no_docx()
    with tempfile.TemporaryDirectory() as tmp:
        res = _pair(
            tmp,
            [{"runs": [("AB", {"bold": True})]}],
            [{"runs": [("A", {"bold": True}), ("B", {"bold": True})]}],
        )
        assert res["format"]["count"] == 0
        assert res["has_any_advanced_change"] is False


def test_identical_documents_have_no_advanced_change():
    skip_if_no_docx()
    with tempfile.TemporaryDirectory() as tmp:
        blocks = [{"runs": [("标题", {"bold": True})]}, "正文"]
        res = _pair(tmp, blocks, blocks)
        assert res["has_any_advanced_change"] is False
        assert res["total_changes"] == 0


# ------------------------------------------------------------------ 空段落

def test_blank_paragraph_added_is_detected():
    skip_if_no_docx()
    with tempfile.TemporaryDirectory() as tmp:
        res = _pair(tmp, ["A", "B"], ["A", "", "B"])
        blank = res["blank_paragraphs"]
        assert blank["changed"] is True
        assert blank["doc1"]["count"] == 0
        assert blank["doc2"]["count"] >= 1
        assert blank["blank_diff"], "应给出具体是哪个段落被加了空行"


def test_whitespace_only_paragraph_reports_codepoints():
    """全角空格（U+3000）这类"看起来是空行"的段落要能被识别并给出码点。"""
    skip_if_no_docx()
    with tempfile.TemporaryDirectory() as tmp:
        res = _pair(tmp, ["A", "B"], ["A", "\u3000", "B"])
        doc2 = res["blank_paragraphs"]["doc2"]
        assert doc2["whitespace_only"] >= 1
        items = [i for i in doc2["items"] if i["kind"] == "whitespace_only"]
        assert items and items[0]["codepoints"], "应给出空白字符的 Unicode 码点"


# ------------------------------------------------------------------ 批注

def test_comment_added_is_detected():
    skip_if_no_docx()
    with tempfile.TemporaryDirectory() as tmp:
        res = _pair(
            tmp,
            ["标题", "正文"],
            [{"runs": [("标题", {})], "comment": "这里是批注"}, "正文"],
        )
        comments = res["comments"]
        assert comments["changed"] is True
        assert comments["doc1_count"] == 0
        assert comments["doc2_count"] == 1


# ------------------------------------------------------------------ 图片

def test_identical_image_is_not_reported_as_change():
    skip_if_no_docx()
    with tempfile.TemporaryDirectory() as tmp:
        png = make_png(Path(tmp) / "same.png", rgb=(255, 0, 0))
        res = _pair(tmp, ["正文"], ["正文"], picture_a=png, picture_b=png)
        assert res["images"]["changed"] is False
        assert res["images"]["doc1_count"] == res["images"]["doc2_count"] == 1


def test_added_image_is_detected():
    skip_if_no_docx()
    with tempfile.TemporaryDirectory() as tmp:
        png = make_png(Path(tmp) / "only_b.png", rgb=(0, 0, 255), size=3)
        res = _pair(tmp, ["正文"], ["正文"], picture_b=png)
        assert res["images"]["changed"] is True
        assert res["images"]["image_diff"], "应给出图片差异明细"


def test_same_count_different_image_content_is_detected():
    """回归：图片数量相同、但内容(MD5)不同也要报差异。

    对应修复前的 bug：InlineShape 没有 `.part` 属性，导致 MD5 恒为 None，
    “数量相同、内容不同”的图片被漏报。
    """
    skip_if_no_docx()
    with tempfile.TemporaryDirectory() as tmp:
        red = make_png(Path(tmp) / "red.png", rgb=(255, 0, 0), size=3)
        green = make_png(Path(tmp) / "green.png", rgb=(0, 255, 0), size=3)
        res = _pair(tmp, ["正文"], ["正文"], picture_a=red, picture_b=green)
        assert res["images"]["changed"] is True
        assert res["images"]["doc1_count"] == res["images"]["doc2_count"] == 1
        detail = " ".join(str(d) for item in res["images"]["image_diff"] for d in item.get("detail", []))
        assert "MD5" in detail, "应报出“图片内容变化（MD5 不同）”"
