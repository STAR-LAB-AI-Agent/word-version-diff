---
name: word-diff
description: 比较两份 Word(.docx) 文档的新增/删除/修改段落差异，返回结构化 JSON 与中文摘要。当用户需要“对比两份 Word 文档的改动”“看两个版本之间改了什么”“总结 Word 修改要点”时使用。
license: MIT
---

# Skill: Word 版本比较(word-diff)

## 使用场景

- 用户给出两份 `.docx` 文件，希望知道内容差异。
- 需要输出**新增 / 删除 / 修改**三个维度的结构化差异，或生成中文修改摘要。
- 需要进一步查看**格式、图片、批注、空段落**层面的差异（进阶分析）。

## 调用方式

脚本在项目根目录，`python-docx` 需已安装：

```bash
pip install python-docx
```

### 基础比较（结构化 JSON）

```bash
python word_diff.py <文档1.docx> <文档2.docx>
python word_diff.py <文档1.docx> <文档2.docx> --pretty
```

### 仅中文修改摘要（低 Token）

```bash
python word_diff.py <文档1.docx> <文档2.docx> --summary
```

### 进阶分析（格式 / 图片 / 批注 / 空段落）

```bash
python word_diff.py <文档1.docx> <文档2.docx> --advanced
python word_diff_advanced.py <文档1.docx> <文档2.docx> --pretty
```

### 可选参数

| 参数 | 说明 |
| --- | --- |
| `--pretty` | 缩进美化 JSON 输出 |
| `--summary` | 仅输出摘要与统计（推荐，低 Token） |
| `--advanced` | 追加格式 / 图片 / 批注 / 空段落进阶分析 |
| `--similarity <0~1>` | 修改/删除/新增 配对的相似度阈值，默认 `0.6`，越接近 1 越严格（进阶脚本 `word_diff_advanced.py` 同样支持，用于格式比较的段落配对） |
| `--keep-empty` | 保留空段落参与比较（默认忽略空段落，降低噪声） |

## 参数

- `doc1` / `doc2`：两个 `.docx` 文件的路径（必填）。需存在且以 `.docx`/`.DOCX` 结尾。
- 对比基于**段落文本快照**，不依赖 Word 修订记录。

## 结果格式

成功返回 `status=success`，核心在 `differences`（`added`/`removed`/`modified` 三组，每项含 `kind`、`paragraph`(0-based 段落号)、`old_text`、`new_text`）与 `summary`（`added`/`removed`/`modified`/`total`），以及 `summary_text`（中文摘要）。`--advanced` 时额外含 `advanced` 字段。失败返回 `status=error` 及 `error_type`、`message`。

### `--advanced` 的 `advanced` 字段

| 子字段 | 内容 |
| --- | --- |
| `format` | 格式差异。`format_changes[].char_range`（0-based 左闭右开）+ `props`（`old`/`new` 为生效值，`raw_old`/`raw_new` 为 run 上直接设置的原值，`source_*` 为 `run`/`style:名称`/`docDefaults`/`unresolved`）；`count` 为真实格式变化处数 |
| `format.run_structure_changes` | 仅 run 切分变化（渲染无差异）的提示项，单独计数 `structure_count` |
| `format.style_expression_changes` | 写法不同（如显式 `False` vs 未设置）但生效值一致，不算格式变化 |
| `format.text_changes` | 文本不同的区间，不做格式比较 |
| `format.alignment` | 段落对齐结果：配对数、仅文档1/文档2 有的段落号 |
| `images` | 图片数量与内容（MD5）/尺寸差异 |
| `comments` | 批注按 (作者, 内容) 逐条对比 |
| `blank_paragraphs` | 真空段落 / 仅含空白字符的段落统计与差异，含 `codepoints`（U+0020 空格、U+3000 全角空格、U+0009 制表符等） |

## 示例

**示例 1 — 基础比较：**

```bash
python word_diff.py v1.docx v2.docx --pretty
```

```json
{
  "status": "success",
  "is_identical": false,
  "summary": {"added": 1, "removed": 0, "modified": 1, "total": 2},
  "summary_text": "对比《v1.docx》与《v2.docx》共发现 2 处变动（新增 1，删除 0，修改 1）。\n\n- 修改（第2段）：…",
  "differences": {
    "added":    [{"kind": "added", "paragraph": 3, "old_text": null, "new_text": "新增章节"}],
    "removed":  [],
    "modified": [{"kind": "modified", "paragraph": 1, "old_text": "旧表述", "new_text": "新表述"}]
  }
}
```

**示例 2 — 仅摘要：**

```bash
python word_diff.py v1.docx v2.docx --summary
```

```json
{
  "status": "success",
  "is_identical": false,
  "summary": {"added": 1, "removed": 0, "modified": 1, "total": 2},
  "summary_text": "对比《v1.docx》与《v2.docx》共发现 2 处变动（新增 1，删除 0，修改 1）。"
}
```

**示例 3 — 进阶分析（追加格式 / 图片 / 批注 / 空段落）：**

```bash
python word_diff.py v1.docx v2.docx --advanced
```

```json
{
  "advanced": {
    "has_any_advanced_change": true,
    "total_changes": 4,
    "format": {
      "format_changes": [{"paragraph": 3, "paragraph_new": 4, "char_range": [8, 49],
                          "detail": ["粗体: False → True"], "certainty": "confirmed"}],
      "count": 1
    },
    "images": {"doc1_count": 0, "doc2_count": 0, "changed": false},
    "comments": {"doc1_count": 0, "doc2_count": 1, "changed": true},
    "blank_paragraphs": {"doc1": {"count": 2}, "doc2": {"count": 3}, "changed": true},
    "summary_text": "进阶分析：格式变化 1 处，图片差异 0 处，批注差异 2 处，空段落差异 1 处"
  }
}
```

## 注意事项 / 限制

- 只比较段落文本与（可选）格式/图片/批注/空段落，**不解析表格单元格内文字**。
- 格式比较**按段落内容对齐**（不是按段落序号），因此插入/删除段落不会导致其后段落的格式被漏报；比较粒度是**逐字符**，run 切分变化不再产生假差异。
- 判定“格式变化”用**生效值**（沿 字符样式 → 段落样式 → docDefaults 解析），`raw_*` 保留 run 上的原值；仅写法不同、渲染一致的情况归入 `style_expression_changes`。
- `paragraph.runs` **不含超链接内部的 run**，超链接内的文字与格式不参与比较。
- `format.count` 只统计真实格式变化；run 切分差异见 `run_structure_changes`（`structure_count`）。
- 空段落检测基于**原始段落文本**，可区分“真空段落”与“仅含空白字符的段落”；连续多个空段落时增删位置不唯一，会给 `position_ambiguous` + `candidates` + 前后上下文。
- 图片内容差异依赖内嵌关系部件读取，浮动图片可能无法读取（`hashed=false`）。
- 批注比较不含时间字段（`python-docx` 限制）。
- 结果中不输出完整文档内容，仅输出差异；建议结合 `--summary` 降低 Token。
