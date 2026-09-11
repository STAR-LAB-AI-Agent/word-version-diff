#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""word_diff.diff_engine —— 段落差异计算（纯逻辑核心）。

本模块不依赖 python-docx，只接受文本段落列表，因此可脱离 docx
单独单元测试 / 复用。docx 的读取与解析由 docx_reader 层负责。

设计要点：
* 使用 difflib.SequenceMatcher 得到粗粒度 opcodes；
* 对 replace 块进一步用“相似度 + 贪心配对”拆成 修改/删除/新增，
  避免原实现中“长度不等时误配 / 重复计数 / 漏算”的问题；
* 结果统一使用 DiffItem（kind + position + 新旧文本），便于程序消费。
"""
from __future__ import annotations

from collections.abc import Sequence

import difflib

from .models import DiffItem, DiffResult, ParagraphEntry

# 一个段落被视为“相似/配对”的相似度阈值
DEFAULT_SIMILARITY_THRESHOLD = 0.6


def _similarity(a: str, b: str) -> float:
    """计算两个段落文本的相似度（0.0 ~ 1.0）。"""
    if a == b:
        return 1.0
    if not a or not b:
        return 0.0
    return difflib.SequenceMatcher(None, a, b).ratio()


def _align_replace_block(
    old_entries: Sequence[ParagraphEntry],
    new_entries: Sequence[ParagraphEntry],
    threshold: float,
) -> list[DiffItem]:
    """对单个 replace 块进行细粒度配对，输出差异项列表。

    replace 表示“源文档这一段被目标文档的这一段替换”。我们借此找出
    该块内部哪些是真正的“修改”、哪些是“删除”、哪些是“新增”：
    * 先按最优相似度做贪心一对一配对（配对上的视为修改）；
    * 配对后，源文档未配对的段落视为删除；
    * 目标文档未配对的段落视为新增。
    """
    items: list[DiffItem] = []

    old_len = len(old_entries)
    new_len = len(new_entries)

    # ---- 构造所有 (相似度, 旧索引, 新索引) 候选对 ----
    candidates: list[tuple[float, int, int]] = []
    for oi, old in enumerate(old_entries):
        for ni, new in enumerate(new_entries):
            ratio = _similarity(old.text, new.text)
            if ratio >= threshold:
                candidates.append((ratio, oi, ni))

    # 相似度从高到低，贪心分配（每个旧段/新段最多被用一次）
    candidates.sort(key=lambda x: x[0], reverse=True)
    used_old: set[int] = set()
    used_new: set[int] = set()

    for ratio, oi, ni in candidates:
        if oi in used_old or ni in used_new:
            continue
        used_old.add(oi)
        used_new.add(ni)
        # ratio 为 1.0 说明文本完全一致，这属于 diff 引擎漏掉的“相同”，
        # 正常不会进入 replace 块；这里仍视为修改（内容未变，不产生噪声）。
        old_text = old_entries[oi].text
        new_text = new_entries[ni].text
        if ratio < 1.0 or old_text != new_text:
            items.append(
                DiffItem(
                    kind="modified",
                    position=old_entries[oi].index,
                    old_text=old_text,
                    new_text=new_text,
                )
            )

    # ---- 未配对的旧段落 → 删除 ----
    for oi, old in enumerate(old_entries):
        if oi not in used_old:
            items.append(
                DiffItem(kind="removed", position=old.index, old_text=old.text, new_text=None)
            )

    # ---- 未配对的新段落 → 新增 ----
    for ni, new in enumerate(new_entries):
        if ni not in used_new:
            items.append(
                DiffItem(kind="added", position=new.index, old_text=None, new_text=new.text)
            )

    return items


def compute_diff(
    source: Sequence[ParagraphEntry],
    target: Sequence[ParagraphEntry],
    similarity_threshold: float = DEFAULT_SIMILARITY_THRESHOLD,
) -> DiffResult:
    """比较两个段落列表，返回结构化 DiffResult。

    source: 源文档（文档1）段落
    target: 目标文档（文档2）段落
    """
    source_texts = [p.text for p in source]
    target_texts = [p.text for p in target]

    matcher = difflib.SequenceMatcher(None, source_texts, target_texts)
    result = DiffResult()

    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            continue
        elif tag == "delete":
            for p in source[i1:i2]:
                result.items.append(
                    DiffItem(kind="removed", position=p.index, old_text=p.text, new_text=None)
                )
        elif tag == "insert":
            for p in target[j1:j2]:
                result.items.append(
                    DiffItem(kind="added", position=p.index, old_text=None, new_text=p.text)
                )
        elif tag == "replace":
            result.items.extend(
                _align_replace_block(
                    source[i1:i2],
                    target[j1:j2],
                    similarity_threshold,
                )
            )

    # 按“源/目标段落位置”排序，便于阅读（先旧后新自然交错）。
    result.items.sort(key=lambda i: i.position)
    return result


def collect_paragraphs(
    texts: Sequence[str],
    skip_empty: bool = True,
) -> list[ParagraphEntry]:
    """把纯文本列表包装成 ParagraphEntry（index = 原列表下标）。

    skip_empty 为真时过滤掉空白段落，但仍保留其原始 index，
    以便 diff 位置号与原文档段落一致。
    """
    entries: list[ParagraphEntry] = []
    for idx, text in enumerate(texts):
        stripped = (text or "").strip()
        if skip_empty and not stripped:
            continue
        entries.append(ParagraphEntry(index=idx, text=stripped))
    return entries
