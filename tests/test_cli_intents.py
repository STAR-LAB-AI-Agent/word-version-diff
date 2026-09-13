# -*- coding: utf-8 -*-
"""第 3 层：自然语言意图 → 命令 → 结果契约（对应任务书"3 类核心意图"）。

用例数据写在 `nl_cases.py`（一句大白话 + 应映射的参数 + 断言名），本文件只负责执行与断言。

这个文件把验收过程分成两半：

* **工具侧**（本文件自动验证）：给定参数，工具是否返回正确的结构化结果；
* **智能体侧**（验收时人工验证）：教师说出同一句大白话时，智能体是否选中同一个命令。
  验证办法：把 `nl_cases.py` 里的 ``nl`` 原话发给我，检查输出的 JSON 是否满足同一组断言。
"""
import tempfile
from pathlib import Path

import nl_cases
from helpers import intent_doc_pair, make_docx, run_cli, run_script, skip_if_no_docx


# ------------------------------------------------------------------ 断言实现

def _status_ok(data, code, out):
    if not isinstance(data, dict):
        return False, f"未输出可解析的 JSON（退出码 {code}，stdout={out[:120]!r}）"
    if data.get("status") != "success":
        return False, f"status={data.get('status')!r}, message={data.get('message')!r}"
    return (code == 0), f"退出码应为 0，实际 {code}"


def _not_identical(data, code, out):
    return (data.get("is_identical") is False), "应判定两份文档存在差异"


def _differences_shape(data, code, out):
    diffs = data.get("differences")
    if not isinstance(diffs, dict):
        return False, "缺少 differences 对象"
    for key in ("added", "removed", "modified"):
        if not isinstance(diffs.get(key), list):
            return False, f"differences.{key} 缺失或不是列表"
    for item in [i for v in diffs.values() for i in v]:
        missing = {"kind", "paragraph", "old_text", "new_text"} - set(item)
        if missing:
            return False, f"差异项缺少字段 {sorted(missing)}：{item}"
    return True, ""


def _summary_counts_match(data, code, out):
    s = data.get("summary") or {}
    diffs = data.get("differences") or {}
    if s.get("total") != s.get("added", 0) + s.get("removed", 0) + s.get("modified", 0):
        return False, f"summary 自身不自洽：{s}"
    counted = sum(len(v) for v in diffs.values())
    if s.get("total") != counted:
        return False, f"summary.total={s.get('total')} 与 differences 条数={counted} 不一致"
    return True, ""


def _summary_text_cn(data, code, out):
    text = data.get("summary_text") or ""
    ok = isinstance(text, str) and any(k in text for k in ("变动", "一致", "差异"))
    return ok, f"summary_text 应为中文结论，实际 {text[:80]!r}"


def _has_kind(kind):
    def checker(data, code, out):
        items = (data.get("differences") or {}).get(kind) or []
        return bool(items), f"differences.{kind} 不应为空（该意图要求能回答这类追问）"
    return checker


def _summary_only_low_token(data, code, out):
    expected = {"status", "is_identical", "summary", "summary_text"}
    keys = set(data or {})
    if keys != expected:
        return False, f"--summary 应只返回 {sorted(expected)}，实际 {sorted(keys)}"
    return ("differences" not in keys), "低 Token 场景不应把完整差异塞回上下文"


def _pretty_multiline(data, code, out):
    return (out.count("\n") >= 3), "--pretty 应输出缩进的多行 JSON"


def _advanced_present(data, code, out):
    adv = data.get("advanced")
    if not isinstance(adv, dict):
        return False, "缺少 advanced 结果（可能落进了 advanced_warning）"
    return (adv.get("status") == "success"), f"advanced.status={adv.get('status')!r}"


def _advanced_keys(data, code, out):
    adv = data.get("advanced") or {}
    required = {"format", "images", "comments", "blank_paragraphs", "summary_text"}
    missing = required - set(adv)
    return (not missing), f"advanced 缺少字段 {sorted(missing)}"


def _format_change_detected(data, code, out):
    fmt = (data.get("advanced") or {}).get("format") or {}
    if fmt.get("count", 0) < 1:
        return False, f"应检出格式变化，实际 count={fmt.get('count')}"
    changes = fmt.get("format_changes") or []
    if not changes or not changes[0].get("char_range"):
        return False, "格式变化项缺少 char_range（逐字符范围）"
    return True, ""


def _blank_changed(data, code, out):
    blank = (data.get("advanced") or {}).get("blank_paragraphs") or {}
    return bool(blank.get("changed")), "应检出空段落差异"


def _comments_changed(data, code, out):
    comments = (data.get("advanced") or {}).get("comments") or {}
    return bool(comments.get("changed")), "应检出批注差异"


def _images_unchanged(data, code, out):
    images = (data.get("advanced") or {}).get("images") or {}
    return (images.get("changed") is False), "该文档对没有图片差异，不应误报"


CHECKS = {
    "status_ok": _status_ok,
    "not_identical": _not_identical,
    "differences_shape": _differences_shape,
    "summary_counts_match": _summary_counts_match,
    "summary_text_cn": _summary_text_cn,
    "has_added": _has_kind("added"),
    "has_removed": _has_kind("removed"),
    "has_modified": _has_kind("modified"),
    "summary_only_low_token": _summary_only_low_token,
    "pretty_multiline": _pretty_multiline,
    "advanced_present": _advanced_present,
    "advanced_keys": _advanced_keys,
    "format_change_detected": _format_change_detected,
    "blank_changed": _blank_changed,
    "comments_changed": _comments_changed,
    "images_unchanged": _images_unchanged,
}


# ------------------------------------------------------------------ 用例执行

def test_natural_language_intents():
    """逐条跑 nl_cases.py 里的"自然语言 → 命令 → 结果契约"。"""
    skip_if_no_docx()
    failures = []
    with tempfile.TemporaryDirectory(prefix="word_diff_nl_") as tmp:
        doc1, doc2 = intent_doc_pair(tmp)
        for case in nl_cases.NL_CASES:
            code, data, out = run_cli([doc1, doc2, *case["args"]])
            problems = []
            for name in case["checks"]:
                ok, msg = CHECKS[name](data, code, out)
                if not ok:
                    problems.append(f"{name}: {msg}")
            print(
                f"[{'PASS' if not problems else 'FAIL'}] {case['id']} {case['intent']} | "
                f"用户说：{case['nl']} | 参数：{case['args'] or '(默认)'}"
            )
            for problem in problems:
                print(f"        -> {problem}")
                failures.append(f"{case['id']}（{case['nl']}） {problem}")
    assert not failures, "自然语言用例失败：\n" + "\n".join(failures)


def test_identical_documents_reported_as_identical():
    """同一份文档比两次 → 必须明确回答"内容一致"。"""
    skip_if_no_docx()
    with tempfile.TemporaryDirectory() as tmp:
        path = make_docx(Path(tmp) / "same.docx", ["标题", "正文"])
        code, data, out = run_cli([path, path, "--summary"])
        assert code == 0 and data["status"] == "success"
        assert data["is_identical"] is True
        assert "一致" in data["summary_text"]


def test_keep_empty_flag_makes_blank_only_change_visible():
    """只多了一个空段落：默认结论是"一致"，加 --keep-empty 后能看出来。"""
    skip_if_no_docx()
    with tempfile.TemporaryDirectory() as tmp:
        a = make_docx(Path(tmp) / "a.docx", ["A", "B"])
        b = make_docx(Path(tmp) / "b.docx", ["A", "", "B"])
        _, default_result, _ = run_cli([a, b, "--summary"])
        assert default_result["is_identical"] is True  # 默认忽略空段落降噪

        _, strict_result, _ = run_cli([a, b, "--keep-empty", "--summary"])
        assert strict_result["is_identical"] is False
        assert strict_result["summary"]["total"] >= 1


def test_standalone_scripts_run_without_agent():
    """底层 Script/CLI 能脱离智能体独立运行（验收硬要求）。"""
    skip_if_no_docx()
    with tempfile.TemporaryDirectory() as tmp:
        doc1, doc2 = intent_doc_pair(tmp)

        code, data, out, err = run_script("word_diff.py", [doc1, doc2, "--summary"])
        assert code == 0, f"word_diff.py 退出码 {code}，stderr={err[:200]}"
        assert data and data["status"] == "success"

        code2, data2, out2, err2 = run_script("word_diff_advanced.py", [doc1, doc2])
        assert code2 == 0, f"word_diff_advanced.py 退出码 {code2}，stderr={err2[:200]}"
        assert data2 and data2["status"] == "success"
        assert {"format", "images", "comments", "blank_paragraphs"} <= set(data2)


def test_version_flag_prints_package_version():
    code, data, out = run_cli(["--version"])
    assert code == 0
    assert out.strip().startswith("word_diff"), f"--version 输出异常：{out!r}"
