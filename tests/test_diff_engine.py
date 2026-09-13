# -*- coding: utf-8 -*-
"""第 1 层：核心 diff 引擎单元测试（**不需要 python-docx，任何环境都能跑**）。

被测对象：`word_diff.diff_engine` + `word_diff.models`

覆盖：正常路径（新增 / 删除 / 修改）、边界（空文档、全空段落、超长段落、
重复段落、相似度阈值上下、位置保序）、回归（replace 块内部分保留部分替换）。
"""
from word_diff.diff_engine import (
    DEFAULT_SIMILARITY_THRESHOLD,
    _similarity,
    collect_paragraphs,
    compute_diff,
)
from word_diff.models import ParagraphEntry


def _entries(*texts):
    """便捷构造：index 即列表下标。"""
    return [ParagraphEntry(index=i, text=t) for i, t in enumerate(texts)]


# ---------------------------------------------------------------- 正常路径

def test_identical_documents():
    a = _entries("标题", "第一段", "第二段")
    b = _entries("标题", "第一段", "第二段")
    result = compute_diff(a, b)
    assert result.is_identical
    assert result.total == 0
    assert result.summary() == {"added": 0, "removed": 0, "modified": 0, "total": 0}


def test_pure_addition():
    a = _entries("标题", "内容A")
    b = _entries("标题", "内容A", "新增段落")
    result = compute_diff(a, b)
    assert len(result.added) == 1
    assert result.added[0].new_text == "新增段落"
    assert result.added[0].old_text is None
    assert result.removed == [] and result.modified == []


def test_pure_deletion():
    a = _entries("标题", "将被删除")
    b = _entries("标题")
    result = compute_diff(a, b)
    assert len(result.removed) == 1
    assert result.removed[0].old_text == "将被删除"
    assert result.removed[0].new_text is None
    assert result.added == []


def test_modification():
    a = _entries("标题", "这是第一段内容")
    b = _entries("标题", "这是第一段内容（已修改）")
    result = compute_diff(a, b)
    assert len(result.modified) == 1
    assert result.modified[0].old_text == "这是第一段内容"
    assert result.modified[0].new_text == "这是第一段内容（已修改）"


def test_modification_position_preserved():
    a = _entries("标题", "旧表述", "结尾")
    b = _entries("标题", "新表述", "结尾")
    result = compute_diff(a, b)
    assert len(result.modified) == 1
    assert result.modified[0].position == 1  # 第 2 段（0-based）


# ---------------------------------------------------------------- 边界

def test_dissimilar_replacement_splits_into_removed_and_added():
    """相似度低于阈值时，不应硬配对成"修改"，而应拆成 删除 + 新增。"""
    a = _entries("标题", "完全不同的句子")
    b = _entries("标题", "另一句毫不相干的话")
    result = compute_diff(a, b, similarity_threshold=0.99)
    assert len(result.removed) == 1
    assert len(result.added) == 1
    assert result.modified == []


def test_high_similarity_pairs_as_modified():
    a = _entries("标题", "第一段：原文内容")
    b = _entries("标题", "第一段：原文内容（已修改）")
    result = compute_diff(a, b)  # 默认阈值 0.6
    assert len(result.modified) == 1
    assert result.added == [] and result.removed == []


def test_empty_paragraphs_skipped_but_index_preserved():
    entries_a = collect_paragraphs(["标题", "", "正文"], skip_empty=True)
    entries_b = collect_paragraphs(["标题", "正文"], skip_empty=True)
    assert len(entries_a) == 2
    assert entries_a[1].index == 2  # 空段被过滤，但原段落号保留
    assert compute_diff(entries_a, entries_b).is_identical


def test_empty_paragraphs_kept_when_skip_disabled():
    entries = collect_paragraphs(["标题", "", "正文"], skip_empty=False)
    assert [e.index for e in entries] == [0, 1, 2]
    assert entries[1].text == ""


def test_all_empty_documents_are_identical():
    assert collect_paragraphs(["", "   ", "\t"], skip_empty=True) == []
    assert compute_diff([], []).is_identical


def test_long_paragraph_single_char_change():
    text = "字" * 1000
    a = _entries(text + "甲")
    b = _entries(text + "乙")
    result = compute_diff(a, b)
    assert len(result.modified) == 1
    assert result.modified[0].position == 0
    assert result.added == [] and result.removed == []


def test_duplicate_paragraphs_counted_once():
    a = _entries("A", "A", "B")
    b = _entries("A", "B")
    result = compute_diff(a, b)
    assert len(result.removed) == 1
    assert result.removed[0].position == 0
    assert result.total == 1


def test_items_sorted_by_position():
    """同时有"修改"和"新增"时，结果按段落号升序排列，便于阅读。"""
    a = _entries("第一段原文", "第二段原文", "第三段")
    b = _entries("第一段原文", "第二段原文（已改）", "第三段", "第四段新增")
    result = compute_diff(a, b)
    positions = [i.position for i in result.items]
    assert positions == sorted(positions)
    assert len(result.modified) == 1
    assert len(result.added) == 1
    assert result.removed == []


def test_summary_counts_are_consistent():
    a = _entries("1", "2", "3")
    b = _entries("1", "2b", "3", "4")
    s = compute_diff(a, b).summary()
    assert s["total"] == s["added"] + s["removed"] + s["modified"]
    assert s["total"] == sum(
        len(x) for x in compute_diff(a, b).to_dict().values()
    )


def test_similarity_boundary_values():
    assert _similarity("完全一样", "完全一样") == 1.0
    assert _similarity("", "abc") == 0.0
    assert _similarity("abc", "") == 0.0
    assert _similarity("abc", "abd") > 0.6


def test_default_similarity_threshold_is_documented_value():
    assert DEFAULT_SIMILARITY_THRESHOLD == 0.6


# ---------------------------------------------------------------- 回归

def test_replace_block_partial_keep_and_replace():
    """replace 块内"部分保留、部分被替换"时不应误判为整块改写。"""
    a = _entries("开头", "保留句A", "将被删句B", "结尾")
    b = _entries("开头", "保留句A", "全新句C", "结尾")
    result = compute_diff(a, b)
    kinds = [i.kind for i in result.items]
    assert "added" in kinds or "modified" in kinds
    assert len(result.items) == 2  # 一句改 + 一句被替换（保留句不计入）
