#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""word_diff.advanced —— 进阶分析：格式 / 图片 / 批注 / 空段落 (按需调用)。

本版本修复了格式比较的两个核心缺陷（旧版按“段落序号”对齐，插入/删除段落
后其后所有段落全部错位，导致格式差异被整体漏报）：

1. **对齐方式**：改为“按段落内容对齐”。复用 difflib.SequenceMatcher 得到
   equal / replace 块；equal 块天然给出 1:1 的 (源段号, 目标段号) 映射，
   replace 块再用“相似度 + 贪心配对”（与 diff_engine 同一口径）拆成
   modified 配对。因此插入、删除段落不会影响其后段落的格式比较。
2. **比较粒度**：从“按 run 下标 zip”改为“按字符比较格式”。旧版一旦
   `len(runs1) != len(runs2)` 就只报 structure_change 并 continue，拿不到
   “哪几个字符变了”；新版把每个字符映射到它的格式属性后逐字比较，
   run 切分变化不再产生假差异、也不再掩盖真差异。

其它要点：
* 格式属性同时保留“原始值”（run 上直接写的）与“生效值”（沿 字符样式 →
  段落样式 → docDefaults 链解析后的值）。判定“是否变化”用生效值，
  避免“显式 bold=False” 与“未设置（继承为不加粗）”被误报为差异；
  这类“写法不同、渲染一致”的情况单独放在 style_expression_changes 里。
* 新增空段落检查：区分“真空段落”与“仅含空白字符的段落”，并列出具体的
   Unicode 码点（例如 U+0020 空格 / U+3000 全角空格 / U+0009 制表符），
   TC 便能判断“多按的是回车还是空格”。
* 图片 / 批注比较逻辑保持原样（批注仍按 (作者, 内容) 逐条比，python-docx
  不暴露时间字段）。
* 与基础脚本共用 docx_reader 的校验/读取与 reporting 的错误结构。

限制：`paragraph.runs` 不含超链接内部的 run，因此超链接内的文字与格式不
参与比较（比较用的段落文本由 runs 拼接，两侧口径一致，不会错位）；
表格单元格内文字仍不参与。
"""
from __future__ import annotations

import difflib
import json
import sys
from typing import Any

from .diff_engine import DEFAULT_SIMILARITY_THRESHOLD
from .docx_reader import read_paragraph_entries, validate_docx_path
from .reporting import build_error

try:  # 与基础层共用同一相似度实现，保证“修改段落”的配对口径一致
    from .diff_engine import _similarity as _para_similarity
except ImportError:  # pragma: no cover - 仅防御性兜底
    def _para_similarity(a: str, b: str) -> float:
        if a == b:
            return 1.0
        if not a or not b:
            return 0.0
        return difflib.SequenceMatcher(None, a, b).ratio()


def _require_docx() -> None:
    """复用 docx_reader 的解析逻辑（含 vendor/ 兜底）。"""
    from .docx_reader import _require_docx as _req
    _req()


# ==================== 格式变化检测 ====================

# 参与比较的格式属性（键名与旧版保持一致，额外增加 strike）
_ATTRS = ("bold", "italic", "underline", "strike", "font_name", "font_size")
_BOOL_ATTRS = ("bold", "italic", "underline", "strike")
_LABELS = {
    "bold": "粗体",
    "italic": "斜体",
    "underline": "下划线",
    "strike": "删除线",
    "font_name": "字体",
    "font_size": "字号",
}


def extract_run_format_info(paragraph) -> list[dict[str, Any]]:
    """提取一个段落内所有 Run 的格式信息（保持旧版行为与字段不变）。"""
    _require_docx()
    runs_info: list[dict[str, Any]] = []
    for run in paragraph.runs:
        font = run.font
        runs_info.append({
            "text": run.text,
            "font_name": font.name if font.name else "默认",
            "font_size": font.size.pt if font.size else None,
            "bold": font.bold,
            "italic": font.italic,
            "underline": font.underline,
        })
    return runs_info


def _font_attr(font, attr: str) -> Any:
    """从 Font 对象取单个属性值；未直接设置时返回 None。"""
    if font is None:
        return None
    try:
        if attr == "font_size":
            size = font.size
            return round(size.pt, 2) if size is not None else None
        if attr == "font_name":
            return font.name or None
        return getattr(font, attr)
    except Exception:
        return None


def _style_chain_value(style, attr: str) -> tuple[Any, str | None]:
    """沿样式的 base_style 链查找第一个非 None 的属性值。"""
    seen: set[int] = set()
    while style is not None and id(style) not in seen:
        seen.add(id(style))
        try:
            font = style.font
        except Exception:
            font = None
        value = _font_attr(font, attr)
        if value is not None:
            name = getattr(style, "name", None) or "?"
            return value, f"style:{name}"
        style = getattr(style, "base_style", None)
    return None, None


def _doc_default_fonts(document) -> dict[str, Any]:
    """读取 w:docDefaults/w:rPrDefault/w:rPr（python-docx 未暴露，需走 XML）。

    读不到时返回空 dict，调用方会退化为“未解析”。
    """
    out: dict[str, Any] = {}
    try:
        styles_el = document.styles.element
        doc_defaults = styles_el.find(_qn("w:docDefaults"))
        if doc_defaults is None:
            return out
        rpr_default = doc_defaults.find(_qn("w:rPrDefault"))
        rpr = rpr_default.find(_qn("w:rPr")) if rpr_default is not None else None
        if rpr is None:
            return out

        for tag, attr in (("w:b", "bold"), ("w:i", "italic"), ("w:strike", "strike")):
            node = rpr.find(_qn(tag))
            if node is not None:
                val = node.get(_qn("w:val"))
                out[attr] = str(val).lower() not in ("0", "false", "off") if val is not None else True

        node = rpr.find(_qn("w:u"))
        if node is not None:
            val = node.get(_qn("w:val"))
            out["underline"] = (val is None) or (str(val).lower() != "none")

        fonts = rpr.find(_qn("w:rFonts"))
        if fonts is not None:
            name = fonts.get(_qn("w:ascii")) or fonts.get(_qn("w:eastAsia"))
            if name:
                out["font_name"] = name

        size = rpr.find(_qn("w:sz"))
        if size is not None:
            raw = size.get(_qn("w:val"))
            if raw:
                out["font_size"] = round(int(raw) / 2.0, 2)
    except Exception:
        return out
    return out


def _qn(tag: str):
    """延迟导入 qn，避免无 docx 环境下导入本模块即失败。"""
    from docx.oxml.ns import qn
    return qn(tag)


def _effective_run_format(run, doc_defaults: dict[str, Any]) -> tuple[dict, dict, dict]:
    """返回 (raw, effective, source) 三组属性字典。

    raw       run 上直接写的值（None 表示未直接设置）
    effective 沿样式链解析后的生效值；布尔属性最终兜底为 False，
              字体名/字号无法确定时保持 None（来自主题，不可解析）
    source    每个生效值的来源标记：run / style:xxx / docDefaults / unresolved
    """
    raw: dict[str, Any] = {}
    effective: dict[str, Any] = {}
    source: dict[str, str] = {}

    for attr in _ATTRS:
        direct = _font_attr(run.font, attr)
        raw[attr] = direct
        if direct is not None:
            effective[attr] = direct
            source[attr] = "run"
            continue

        value, src = _style_chain_value(getattr(run, "style", None), attr)
        if value is not None:
            effective[attr] = value
            source[attr] = src or "style"
            continue

        if attr in doc_defaults:
            effective[attr] = doc_defaults[attr]
            source[attr] = "docDefaults"
            continue

        # 兜底：布尔属性按 Word 默认（不加粗/不斜体/无下划线/无删除线）
        effective[attr] = False if attr in _BOOL_ATTRS else None
        source[attr] = "unresolved"

    return raw, effective, source


def _paragraph_char_formats(paragraph, doc_defaults: dict[str, Any]) -> tuple[str, list[dict]]:
    """把段落展开成“逐字符格式”：返回 (文本, 每个字符的格式信息)。

    文本由 paragraph.runs 拼接（与逐字符表严格等长），因此不存在长度错配；
    超链接内的 run 不在 paragraph.runs 中，故不参与（见模块 docstring）。
    """
    chars: list[dict] = []
    text_parts: list[str] = []
    for run in paragraph.runs:
        raw, effective, source = _effective_run_format(run, doc_defaults)
        text_parts.append(run.text)
        for _ch in run.text:
            chars.append({"raw": raw, "eff": effective, "src": source})
    return "".join(text_parts), chars


def _value_text(attr: str, value: Any) -> str:
    """把属性值渲染成摘要里可读的文本。"""
    if attr == "font_size":
        return f"{value}pt" if value is not None else "未确定"
    if value is None:
        return "未确定" if attr == "font_name" else "无"
    if isinstance(value, bool):
        return "True" if value else "False"
    return str(value)


def _char_change_spec(c1: dict, c2: dict) -> tuple:
    """字符级“生效值差异”签名；全部一致时返回空元组。"""
    spec = []
    for attr in _ATTRS:
        old, new = c1["eff"][attr], c2["eff"][attr]
        if old != new:
            spec.append((attr, old, new, c1["raw"][attr], c2["raw"][attr],
                         c1["src"][attr], c2["src"][attr]))
    return tuple(spec)


def _char_noise_spec(c1: dict, c2: dict) -> tuple:
    """字符级“写法不同但生效值一致”签名（raw 不同、effective 相同）。"""
    spec = []
    for attr in _ATTRS:
        if c1["raw"][attr] != c2["raw"][attr] and c1["eff"][attr] == c2["eff"][attr]:
            spec.append((attr, c1["raw"][attr], c2["raw"][attr], c1["src"][attr], c2["src"][attr]))
    return tuple(spec)


def _clip(text: str, limit: int = 60) -> str:
    text = (text or "").replace("\n", " ").replace("\t", " ")
    return text if len(text) <= limit else text[: limit - 1] + "…"


def _build_spec_item(spec: tuple, index1: int, index2: int, text1: str,
                     start: int, end: int, run_counts: tuple[int, int]) -> dict[str, Any]:
    """把一段连续字符的差异签名组装成结果条目。"""
    detail: list[str] = []
    props: dict[str, Any] = {}
    uncertain = False
    for attr, old, new, raw_old, raw_new, src_old, src_new in spec:
        detail.append(f"{_LABELS[attr]}: {_value_text(attr, old)} → {_value_text(attr, new)}")
        props[attr] = {
            "old": old,
            "new": new,
            "raw_old": raw_old,
            "raw_new": raw_new,
            "source_old": src_old,
            "source_new": src_new,
        }
        if attr in ("font_name", "font_size") and "unresolved" in (src_old, src_new):
            uncertain = True

    return {
        "paragraph": index1,
        "paragraph_new": index2,
        "kind": "property_change",
        "char_range": [start, end],           # 0-based，左闭右开
        "length": end - start,
        "text": _clip(text1[start:end]),
        "detail": detail,
        "props": props,
        "run_count": [run_counts[0], run_counts[1]],
        "certainty": "inheritance_unknown" if uncertain else "confirmed",
    }


def align_paragraphs(
    texts1: list[str],
    texts2: list[str],
    similarity_threshold: float = DEFAULT_SIMILARITY_THRESHOLD,
) -> tuple[list[tuple[int, int]], list[int], list[int]]:
    """按内容对齐两份文档的段落，返回 (配对, 仅文档1有的下标, 仅文档2有的下标)。

    与 diff_engine.compute_diff 同源：SequenceMatcher 拿粗粒度 opcodes，
    equal 块直接 1:1 配对，replace 块用“相似度 + 贪心”细化配对。
    """
    matcher = difflib.SequenceMatcher(None, list(texts1), list(texts2))
    pairs: list[tuple[int, int]] = []
    only1: list[int] = []
    only2: list[int] = []

    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            for offset in range(i2 - i1):
                pairs.append((i1 + offset, j1 + offset))
        elif tag == "replace":
            candidates: list[tuple[float, int, int]] = []
            for oi in range(i1, i2):
                for nj in range(j1, j2):
                    ratio = _para_similarity(texts1[oi], texts2[nj])
                    if ratio >= similarity_threshold:
                        candidates.append((ratio, oi, nj))
            candidates.sort(key=lambda x: x[0], reverse=True)
            used1: set[int] = set()
            used2: set[int] = set()
            for _ratio, oi, nj in candidates:
                if oi in used1 or nj in used2:
                    continue
                used1.add(oi)
                used2.add(nj)
                pairs.append((oi, nj))
            only1.extend(oi for oi in range(i1, i2) if oi not in used1)
            only2.extend(nj for nj in range(j1, j2) if nj not in used2)
        elif tag == "delete":
            only1.extend(range(i1, i2))
        else:  # insert
            only2.extend(range(j1, j2))

    pairs.sort()
    return pairs, only1, only2


def compare_format(
    doc1_path: str,
    doc2_path: str,
    similarity_threshold: float = DEFAULT_SIMILARITY_THRESHOLD,
    skip_empty: bool = False,
) -> dict[str, Any]:
    """比较两份文档的格式差异（按内容对齐段落 + 逐字符比较格式）。

    skip_empty=False（默认）表示空段落也参与对齐，保证段号映射稳定；
    空段落本身通常没有字符，不产生格式差异，不影响结论。
    """
    _require_docx()
    from docx import Document

    entries1 = read_paragraph_entries(doc1_path, skip_empty=skip_empty)
    entries2 = read_paragraph_entries(doc2_path, skip_empty=skip_empty)
    doc1 = Document(doc1_path)
    doc2 = Document(doc2_path)
    defaults1 = _doc_default_fonts(doc1)
    defaults2 = _doc_default_fonts(doc2)

    pairs, only1, only2 = align_paragraphs(
        [e.text for e in entries1], [e.text for e in entries2], similarity_threshold
    )

    changes: list[dict[str, Any]] = []
    style_expression: list[dict[str, Any]] = []
    structure: list[dict[str, Any]] = []
    text_changes: list[dict[str, Any]] = []

    for pos1, pos2 in pairs:
        index1 = entries1[pos1].index
        index2 = entries2[pos2].index
        para1 = doc1.paragraphs[index1]
        para2 = doc2.paragraphs[index2]

        text1, chars1 = _paragraph_char_formats(para1, defaults1)
        text2, chars2 = _paragraph_char_formats(para2, defaults2)
        run_counts = (len(para1.runs), len(para2.runs))

        if run_counts[0] != run_counts[1]:
            structure.append({
                "paragraph": index1,
                "paragraph_new": index2,
                "kind": "structure_change",
                "detail": f"格式片段数量由 {run_counts[0]} 变为 {run_counts[1]}"
                          "（已改为逐字符比较，此项仅作提示）",
                "run_count": [run_counts[0], run_counts[1]],
            })

        if text1 == text2:
            blocks = [("equal", 0, len(text1), 0, len(text2))]
        else:
            blocks = difflib.SequenceMatcher(None, text1, text2).get_opcodes()

        for tag, i1, i2, j1, j2 in blocks:
            if tag != "equal":
                text_changes.append({
                    "paragraph": index1,
                    "paragraph_new": index2,
                    "kind": "text_change",
                    "detail": "该区间文本不同，不做格式比较",
                    "old_text": _clip(text1[i1:i2]),
                    "new_text": _clip(text2[j1:j2]),
                })
                continue

            offset = 0
            while offset < i2 - i1:
                c1 = chars1[i1 + offset]
                c2 = chars2[j1 + offset]
                spec = _char_change_spec(c1, c2)
                noise = () if spec else _char_noise_spec(c1, c2)
                if not spec and not noise:
                    offset += 1
                    continue

                # 合并“差异签名完全一致”的连续字符
                length = 1
                while offset + length < i2 - i1:
                    n1 = chars1[i1 + offset + length]
                    n2 = chars2[j1 + offset + length]
                    if spec:
                        if _char_change_spec(n1, n2) != spec:
                            break
                    elif _char_noise_spec(n1, n2) != noise:
                        break
                    length += 1

                start = i1 + offset
                end = start + length
                if spec:
                    changes.append(_build_spec_item(
                        spec, index1, index2, text1, start, end, run_counts))
                else:
                    style_expression.append({
                        "paragraph": index1,
                        "paragraph_new": index2,
                        "kind": "style_expression",
                        "char_range": [start, end],
                        "length": length,
                        "text": _clip(text1[start:end]),
                        "detail": [
                            f"{_LABELS[attr]}: 直接设置 {raw_old} → {raw_new}"
                            f"（生效值一致，渲染无差别）"
                            for attr, raw_old, raw_new, _s1, _s2 in noise
                        ],
                    })
                offset += length

    return {
        "format_changes": changes,
        "count": len(changes),
        "style_expression_changes": style_expression,
        "style_expression_count": len(style_expression),
        "run_structure_changes": structure,
        "structure_count": len(structure),
        "text_changes": text_changes,
        "text_change_count": len(text_changes),
        "alignment": {
            "method": "content(SequenceMatcher equal) + similarity-greedy(replace)",
            "similarity_threshold": similarity_threshold,
            "paired_paragraphs": len(pairs),
            "doc1_only_paragraphs": [entries1[i].index for i in only1],
            "doc2_only_paragraphs": [entries2[j].index for j in only2],
        },
        "note": (
            "按段落内容对齐（不再按段落序号），插入/删除段落不会导致其后段落错位；"
            "格式按字符比较，run 切分变化不再产生假差异；"
            "判定差异使用沿样式链解析后的生效值，写法不同但渲染一致的差异归入 "
            "style_expression_changes；文本不同的区间不做格式比较，见 text_changes；"
            "超链接内文字与格式不在比较范围。"
        ),
    }


# ==================== 空段落 / 空白段落检测 ====================

_WS_NAMES = {
    "\u0020": "空格",
    "\u00a0": "不换行空格(NBSP)",
    "\u3000": "全角空格",
    "\u2002": "EN SPACE",
    "\u2003": "EM SPACE",
    "\u2009": "THIN SPACE",
    "\u200b": "零宽空格",
    "\t": "制表符",
    "\u000b": "垂直制表符",
    "\u000c": "分页符",
}


def _blank_kind(text: str) -> str | None:
    """把段落原始文本分类：empty（真空）/ whitespace_only（仅空白）/ None（有内容）。

    注意用原始文本而非 strip 后的文本，从而能区分“多按了一个回车”与
    “多按了一个空格”。
    """
    if text == "":
        return "empty"
    if text.strip() == "":
        return "whitespace_only"
    return None


def _whitespace_breakdown(text: str) -> list[dict[str, Any]]:
    """统计段落里的空白字符构成（码点 + 数量），便于人工判断。"""
    counts: dict[str, int] = {}
    for ch in text:
        counts[ch] = counts.get(ch, 0) + 1
    return [
        {
            "codepoint": f"U+{ord(ch):04X}",
            "name": _WS_NAMES.get(ch, "其它空白字符"),
            "count": n,
        }
        for ch, n in counts.items()
    ]


def _blank_desc(item: dict[str, Any]) -> str:
    if item["kind"] == "empty":
        return "真空段落（无任何字符，也没有文本片段）"
    parts = "、".join(f"{c['name']} ×{c['count']}" for c in item["codepoints"])
    return f"仅含空白字符的段落（{parts}，共 {item['length']} 个字符）"


def extract_blank_paragraphs(doc_path: str) -> list[dict[str, Any]]:
    """列出所有“真空段落”与“仅含空白字符的段落”（只读，不改动文档）。"""
    _require_docx()
    from docx import Document

    doc = Document(doc_path)
    items: list[dict[str, Any]] = []
    for idx, para in enumerate(doc.paragraphs):
        text = para.text or ""
        kind = _blank_kind(text)
        if kind is None:
            continue
        item = {
            "index": idx,
            "kind": kind,
            "length": len(text),
            "run_count": len(para.runs),
            "codepoints": _whitespace_breakdown(text),
            "raw": text,
        }
        item["desc"] = _blank_desc(item)
        items.append(item)
    return items


def _blank_run_indices(blanks: list[dict[str, Any]], index: int) -> list[int]:
    """返回包含 index 的“连续空段落”区段（按段落号连续）。"""
    indices = {item["index"] for item in blanks}
    if index not in indices:
        return [index]
    low = high = index
    while low - 1 in indices:
        low -= 1
    while high + 1 in indices:
        high += 1
    return list(range(low, high + 1))


def _neighbour_context(entries, pos: int) -> dict[str, Any]:
    """取该段落前后最近的“有内容段落”文本，帮助人工定位空段落位置。"""
    before = ""
    for i in range(pos - 1, -1, -1):
        if entries[i].text:
            before = _clip(entries[i].text, 40)
            break
    after = ""
    for i in range(pos + 1, len(entries)):
        if entries[i].text:
            after = _clip(entries[i].text, 40)
            break
    return {"before": before, "after": after}


def compare_blank_paragraphs(
    doc1_path: str,
    doc2_path: str,
    similarity_threshold: float = DEFAULT_SIMILARITY_THRESHOLD,
) -> dict[str, Any]:
    """比较空段落 / 仅空白段落（口径与 --keep-empty 的基础比较一致）。"""
    blanks1 = extract_blank_paragraphs(doc1_path)
    blanks2 = extract_blank_paragraphs(doc2_path)

    entries1 = read_paragraph_entries(doc1_path, skip_empty=False)
    entries2 = read_paragraph_entries(doc2_path, skip_empty=False)
    pairs, only1, only2 = align_paragraphs(
        [e.text for e in entries1], [e.text for e in entries2], similarity_threshold
    )

    by1 = {item["index"]: item for item in blanks1}
    by2 = {item["index"]: item for item in blanks2}

    diff: list[dict[str, Any]] = []

    for pos1, pos2 in pairs:
        index1 = entries1[pos1].index
        index2 = entries2[pos2].index
        a, b = by1.get(index1), by2.get(index2)
        if a is None or b is None:
            continue
        if a["kind"] != b["kind"]:
            diff.append({
                "kind": "changed",
                "paragraph": index1,
                "paragraph_new": index2,
                "detail": f"第{index1}段（doc1）{a['desc']} → 第{index2}段（doc2）{b['desc']}",
            })

    for pos in only1:
        index = entries1[pos].index
        item = by1.get(index)
        if item is None:
            continue
        entry: dict[str, Any] = {
            "kind": "removed",
            "paragraph": index,
            "detail": f"doc1 第{index}段被删除：{item['desc']}",
            "context": _neighbour_context(entries1, pos),
        }
        run = _blank_run_indices(blanks1, index)
        if len(run) > 1:
            entry["candidates"] = run
            entry["position_ambiguous"] = True
            entry["detail"] += f"（该处连续 {len(run)} 个空段落，具体是哪一段不唯一：{run}）"
        diff.append(entry)

    for pos in only2:
        index = entries2[pos].index
        item = by2.get(index)
        if item is None:
            continue
        entry = {
            "kind": "added",
            "paragraph": index,
            "detail": f"doc2 第{index}段为新增：{item['desc']}",
            "context": _neighbour_context(entries2, pos),
        }
        run = _blank_run_indices(blanks2, index)
        if len(run) > 1:
            entry["candidates"] = run
            entry["position_ambiguous"] = True
            entry["detail"] += f"（该处连续 {len(run)} 个空段落，具体是哪一段不唯一：{run}）"
        diff.append(entry)

    return {
        "doc1": {
            "count": len(blanks1),
            "empty": sum(1 for i in blanks1 if i["kind"] == "empty"),
            "whitespace_only": sum(1 for i in blanks1 if i["kind"] == "whitespace_only"),
            "indices": [i["index"] for i in blanks1],
            "items": blanks1,
        },
        "doc2": {
            "count": len(blanks2),
            "empty": sum(1 for i in blanks2 if i["kind"] == "empty"),
            "whitespace_only": sum(1 for i in blanks2 if i["kind"] == "whitespace_only"),
            "indices": [i["index"] for i in blanks2],
            "items": blanks2,
        },
        "blank_diff": diff,
        "changed": len(diff) > 0 or len(blanks1) != len(blanks2),
        "note": (
            "按原始段落文本判定：empty=真空段落，whitespace_only=仅含空白字符"
            "（并给出 Unicode 码点，用于区分回车与空格）；"
            "连续多个空段落之间的增删在位置上不唯一，此时会给出 position_ambiguous "
            "与 candidates（候选段号）以及前后最近的有内容段落作为上下文。"
        ),
    }


# ==================== 图片差异检测 ====================

def _image_blob(shape, document_part) -> bytes | None:
    """尽力从 inline_shape 提取图片二进制；读取不到返回 None。

    说明：python-docx 的 InlineShape 没有 `.part` 属性（只有
    _inline/height/type/width），图片部件需通过文档部件的
    `document_part.related_parts[rId]` 解析，因此显式传入文档部件。
    """
    try:
        blob = shape._inline.graphic.graphicData.pic.blipFill.blip
        # python-docx 不直接暴露图片二进制，这里通过 relationship 读取
        rId = blob.embed if hasattr(blob, "embed") else None
        if rId is None:
            return None
        image_part = document_part.related_parts[rId]
        return image_part.blob
    except Exception:
        return None


def extract_image_info(doc_path: str) -> list[dict[str, Any]]:
    _require_docx()
    import hashlib

    from docx import Document

    doc = Document(doc_path)
    images: list[dict[str, Any]] = []
    for idx, shape in enumerate(doc.inline_shapes):
        info: dict[str, Any] = {"index": idx}
        try:
            info["width_cm"] = round(shape.width.cm, 2)
            info["height_cm"] = round(shape.height.cm, 2)
        except Exception:
            info["width_cm"] = None
            info["height_cm"] = None

        blob = _image_blob(shape, doc.part)
        if blob:
            info["md5"] = hashlib.md5(blob).hexdigest()
            info["hashed"] = True
            info["bytes"] = len(blob)
        else:
            info["md5"] = None
            info["hashed"] = False
        images.append(info)
    return images


def compare_images(doc1_path: str, doc2_path: str) -> dict[str, Any]:
    """比较两份文档的图片数量与内容差异。"""
    images1 = extract_image_info(doc1_path)
    images2 = extract_image_info(doc2_path)

    diff = []
    if len(images1) != len(images2):
        diff.append({"kind": "count_change", "detail": f"图片数量 {len(images1)} → {len(images2)}"})

    for idx in range(min(len(images1), len(images2))):
        img1, img2 = images1[idx], images2[idx]
        detail = []

        # 尺寸
        w1, w2 = img1.get("width_cm"), img2.get("width_cm")
        h1, h2 = img1.get("height_cm"), img2.get("height_cm")
        if w1 and w2 and abs(w1 - w2) > 0.1:
            detail.append(f"宽度: {w1}cm → {w2}cm")
        if h1 and h2 and abs(h1 - h2) > 0.1:
            detail.append(f"高度: {h1}cm → {h2}cm")

        # 内容：仅当两边都成功哈希时才可能做比较
        m1, m2 = img1.get("md5"), img2.get("md5")
        if m1 and m2 and m1 != m2:
            detail.append("图片内容变化（MD5 不同）")

        if detail:
            diff.append({"image_index": idx, "detail": detail})

    return {
        "doc1_count": len(images1),
        "doc2_count": len(images2),
        "image_diff": diff,
        "changed": len(diff) > 0 or len(images1) != len(images2),
    }


# ==================== 批注差异检测 ====================

def extract_comments_info(doc_path: str) -> list[dict[str, Any]]:
    """提取文档中所有批注（作者、内容、缩写，保留顺序）。"""
    _require_docx()
    from docx import Document

    doc = Document(doc_path)
    comments: list[dict[str, Any]] = []
    for comment in doc.comments:
        text = comment.text or ""
        comments.append({
            "author": getattr(comment, "author", "未知"),
            "initials": getattr(comment, "initials", None),
            "text": text,
            "truncated": len(text) > 100,
        })
    return comments


def compare_comments(doc1_path: str, doc2_path: str) -> dict[str, Any]:
    """比较两份文档的批注差异（顺序与作者敏感，改用逐条对比而非 set）。"""
    comments1 = extract_comments_info(doc1_path)
    comments2 = extract_comments_info(doc2_path)

    c1 = [(c["author"], c["text"]) for c in comments1]
    c2 = [(c["author"], c["text"]) for c in comments2]

    diff: list[dict[str, Any]] = []
    if len(comments1) != len(comments2):
        diff.append({"kind": "count_change", "detail": f"批注数量 {len(comments1)} → {len(comments2)}"})

    # 简单的前缀/追加式比较：找公共前段与公共后段
    prefix = 0
    max_prefix = min(len(c1), len(c2))
    while prefix < max_prefix and c1[prefix] == c2[prefix]:
        prefix += 1

    suffix = 0
    max_suffix = min(len(c1), len(c2)) - prefix
    while suffix < max_suffix and c1[len(c1) - 1 - suffix] == c2[len(c2) - 1 - suffix]:
        suffix += 1

    added = c2[prefix : len(c2) - suffix if suffix else len(c2)]
    removed = c1[prefix : len(c1) - suffix if suffix else len(c1)]

    if removed:
        diff.append({"kind": "removed", "detail": [f"{a}: {t[:60]}" for a, t in removed]})
    if added:
        diff.append({"kind": "added", "detail": [f"{a}: {t[:60]}" for a, t in added]})

    return {
        "doc1_count": len(comments1),
        "doc2_count": len(comments2),
        "comment_diff": diff,
        "changed": len(diff) > 0 or len(comments1) != len(comments2),
        "note": "按(作者, 内容)逐条对比；批注时间信息在 python-docx 中不可用",
    }


# ==================== 主流程 ====================

def analyze_advanced(
    doc1_path: str,
    doc2_path: str,
    similarity_threshold: float = DEFAULT_SIMILARITY_THRESHOLD,
) -> dict[str, Any]:
    """整合格式 / 图片 / 批注 / 空段落四项进阶检测。"""
    for path in (doc1_path, doc2_path):
        err = validate_docx_path(path)
        if err is not None:
            return err

    try:
        format_result = compare_format(doc1_path, doc2_path, similarity_threshold)
        image_result = compare_images(doc1_path, doc2_path)
        comment_result = compare_comments(doc1_path, doc2_path)
        blank_result = compare_blank_paragraphs(doc1_path, doc2_path, similarity_threshold)

        total = (
            format_result["count"]
            + len(image_result["image_diff"])
            + len(comment_result["comment_diff"])
            + len(blank_result["blank_diff"])
        )
        has_any = (
            format_result["count"] > 0
            or image_result["changed"]
            or comment_result["changed"]
            or blank_result["changed"]
        )

        return {
            "status": "success",
            "has_any_advanced_change": has_any,
            "total_changes": total,
            "format": format_result,
            "images": image_result,
            "comments": comment_result,
            "blank_paragraphs": blank_result,
            "summary_text": (
                f"进阶分析：格式变化 {format_result['count']} 处，"
                f"图片差异 {len(image_result['image_diff'])} 处，"
                f"批注差异 {len(comment_result['comment_diff'])} 处，"
                f"空段落差异 {len(blank_result['blank_diff'])} 处"
            ),
        }
    except Exception as exc:
        return build_error("runtime_error", f"进阶分析失败: {exc}")


def main(argv=None) -> int:
    """进阶 CLI 入口（可独立运行测试），统一用 argparse 而非手动解析。"""
    import argparse

    parser = argparse.ArgumentParser(
        prog="word_diff_advanced",
        description="进阶比较两份 Word 文档的格式 / 图片 / 批注 / 空段落差异。",
        epilog="示例: python -m word_diff.advanced a.docx b.docx --pretty",
    )
    parser.add_argument("doc1", help="第一份文档路径")
    parser.add_argument("doc2", help="第二份文档路径")
    parser.add_argument("--pretty", action="store_true", help="美化输出")
    parser.add_argument(
        "--similarity",
        type=float,
        default=DEFAULT_SIMILARITY_THRESHOLD,
        help=f"修改段落配对（格式比较用）的相似度阈值，默认 {DEFAULT_SIMILARITY_THRESHOLD}",
    )

    args = parser.parse_args(argv)
    result = analyze_advanced(args.doc1, args.doc2, args.similarity)
    if args.pretty:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print(json.dumps(result, ensure_ascii=False))
    return 0 if result.get("status") == "success" else 1


if __name__ == "__main__":
    sys.exit(main())
