# -*- coding: utf-8 -*-
"""自然语言用例表——"用户的自然语言 + 一份要对比的文档对 + 期望结果"。

每条用例三要素：

* ``nl``     ：用户会说的大白话（验收时教师随口说的就是这类话）
* ``args``   ：这句话**应映射到**的 CLI 参数（"自然语言 → 命令"的映射由智能体完成）
* ``checks`` ：对该工具返回结果的断言名（实现见 ``test_cli_intents.py`` 里的 ``CHECKS``）

新增用例只需在这张表里加一行，不用改测试代码。
可用断言名：
    status_ok / not_identical / differences_shape / summary_counts_match /
    summary_text_cn / has_added / has_removed / has_modified /
    summary_only_low_token / advanced_present / advanced_keys /
    format_change_detected / blank_changed / comments_changed /
    images_unchanged / pretty_multiline
"""

INTENT_1 = "意图1：对比差异（看改了什么）"
INTENT_2 = "意图2：生成修改摘要（结果要短、省 Token）"
INTENT_3 = "意图3：追问详情（进阶：格式 / 批注 / 空段落 / 图片）"

NL_CASES = [
    # ---------------- 意图 1：对比差异 ----------------
    {
        "id": "I1-001",
        "intent": INTENT_1,
        "nl": "帮我对比一下 v1.docx 和 v2.docx，看改了哪些内容。",
        "args": [],
        "checks": [
            "status_ok",
            "not_identical",
            "differences_shape",
            "summary_counts_match",
            "summary_text_cn",
        ],
    },
    {
        "id": "I1-002",
        "intent": INTENT_1,
        "nl": "这两份文档有哪些段落被删掉了？",
        "args": [],
        "checks": ["status_ok", "has_removed"],
    },
    {
        "id": "I1-003",
        "intent": INTENT_1,
        "nl": "v2 里新增了哪些段落？",
        "args": [],
        "checks": ["status_ok", "has_added"],
    },
    {
        "id": "I1-004",
        "intent": INTENT_1,
        "nl": "有哪几段内容被改写了？",
        "args": [],
        "checks": ["status_ok", "has_modified"],
    },
    # ---------------- 意图 2：生成修改摘要 ----------------
    {
        "id": "I2-001",
        "intent": INTENT_2,
        "nl": "总结一下 v1.docx 到 v2.docx 的主要改动。",
        "args": ["--summary"],
        "checks": [
            "status_ok",
            "not_identical",
            "summary_only_low_token",
            "summary_text_cn",
        ],
    },
    {
        "id": "I2-002",
        "intent": INTENT_2,
        "nl": "结果别太长，给我一句话结论就行。",
        "args": ["--summary"],
        "checks": ["status_ok", "summary_only_low_token"],
    },
    {
        "id": "I2-003",
        "intent": INTENT_2,
        "nl": "输出缩进一下，我好自己读。",
        "args": ["--pretty"],
        "checks": ["status_ok", "pretty_multiline"],
    },
    # ---------------- 意图 3：追问详情（进阶） ----------------
    {
        "id": "I3-001",
        "intent": INTENT_3,
        "nl": "v2 里哪几个字被加粗了？",
        "args": ["--advanced"],
        "checks": [
            "status_ok",
            "advanced_present",
            "advanced_keys",
            "format_change_detected",
        ],
    },
    {
        "id": "I3-002",
        "intent": INTENT_3,
        "nl": "v2 里是不是多插了空行/空段落？",
        "args": ["--advanced"],
        "checks": ["status_ok", "advanced_keys", "blank_changed"],
    },
    {
        "id": "I3-003",
        "intent": INTENT_3,
        "nl": "有没有新增的批注？",
        "args": ["--advanced"],
        "checks": ["status_ok", "advanced_keys", "comments_changed"],
    },
    {
        "id": "I3-004",
        "intent": INTENT_3,
        "nl": "两边的图片有变化吗？",
        "args": ["--advanced"],
        "checks": ["status_ok", "advanced_keys", "images_unchanged"],
    },
]
