#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""word_diff.models —— 差异结果的数据模型。

本模块不依赖 python-docx，可被单独测试 / 复用。
使用 dataclass + 稳定的英文 key，便于程序化消费与序列化；
中文 label 由 reporting 层按需附加。
"""
from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class ParagraphEntry:
    """带原始段落号的一段文本。

    index 是段落号（0-based，对应 .docx 中 doc.paragraphs 的序号），
    供 diff 结果中 position 使用，从而在过滤空段落后仍能精确定位。
    text 为该段去除首尾空白后的内容。
    """
    index: int
    text: str


@dataclass
class DiffItem:
    """单个差异条目。

    kind 取值:
        added      —— 目标文档新增（old_text 为 None）
        removed    —— 源文档删除（new_text 为 None）
        modified   —— 内容发生变化（old_text / new_text 均非 None）
    position 是该条目所在段落号（0-based，对应原文档段落序号）。
    """
    kind: str
    position: int
    old_text: str | None = None
    new_text: str | None = None

    def to_dict(self) -> dict[str, object]:
        return {
            "kind": self.kind,
            "paragraph": self.position,
            "old_text": self.old_text,
            "new_text": self.new_text,
        }


@dataclass
class DiffResult:
    """一次文档比较的结构化结果。"""
    items: list[DiffItem] = field(default_factory=list)

    @property
    def added(self) -> list[DiffItem]:
        return [i for i in self.items if i.kind == "added"]

    @property
    def removed(self) -> list[DiffItem]:
        return [i for i in self.items if i.kind == "removed"]

    @property
    def modified(self) -> list[DiffItem]:
        return [i for i in self.items if i.kind == "modified"]

    @property
    def total(self) -> int:
        return len(self.items)

    @property
    def is_identical(self) -> bool:
        return self.total == 0

    def summary(self) -> dict[str, int]:
        return {
            "added": len(self.added),
            "removed": len(self.removed),
            "modified": len(self.modified),
            "total": self.total,
        }

    def to_dict(self) -> dict[str, object]:
        return {
            "added": [i.to_dict() for i in self.added],
            "removed": [i.to_dict() for i in self.removed],
            "modified": [i.to_dict() for i in self.modified],
        }

    def __iter__(self):
        # 兼容旧版按类别遍历的用法
        return iter(self.items)
