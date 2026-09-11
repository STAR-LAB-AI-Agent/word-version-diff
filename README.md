# AI Word 版本比较

比较两份 `.docx` 文档的新增 / 删除 / 修改内容，输出结构化 JSON 差异结果，
并生成人类可读的中文修改摘要。用于“用户自然语言 → 智能体 → Skill → Python Script/CLI → 开源项目 → 结果”链路中的确定性能力层。

## 环境要求

- **Python：3.10 ~ 3.13**（推荐 3.12，任务书建议范围 3.10~3.12；`pyproject.toml` 已声明 `requires-python = ">=3.10,<3.14"`）。
- 依赖：`python-docx`（见 `requirements.txt` / `pyproject.toml`）。

## 用户场景（支持的 3 类核心意图）

1. **对比差异**：“帮我对比 v1.docx 和 v2.docx，看改了哪些内容。”
2. **生成修改摘要**：“总结一下 v1.docx 到 v2.docx 的主要改动。”
3. **追问详情**：“v2 里新增了哪些段落？第 3 段改了什么？”（配合 `--advanced` 可查格式 / 图片 / 批注差异）

## 参考开源项目

- **python-docx**：读取 `.docx` 段落、格式、内嵌图片与批注。
- 许可证：MIT License（pydocx/python-docx 项目）。
- 实际使用方式：`docx.Document(path)` 读取段落文本与内联图片；`doc.comments` 读取批注；`shape.part.related_parts[rId]` 读取图片二进制以计算哈希。

## 安装

```bash
pip install -r requirements.txt
```

## 运行方法

```bash
# 基础比较（结构化 JSON）
python word_diff.py v1.docx v2.docx

# 缩进美化
python word_diff.py v1.docx v2.docx --pretty

# 仅输出中文修改摘要（低 Token 场景推荐）
python word_diff.py v1.docx v2.docx --summary

# 进阶分析：额外输出格式 / 图片 / 批注差异
python word_diff.py v1.docx v2.docx --advanced

# 或独立运行进阶脚本
python word_diff_advanced.py v1.docx v2.docx --pretty

# 单元测试（不依赖 python-docx，可直接运行）
python tests/test_diff_engine.py
```

## 项目结构

```
word_diff.py            基础比较入口（兼容旧用法）
word_diff_advanced.py   进阶比较入口（按需调用）
word_diff/
  __init__.py           包导出
  models.py             数据模型（ParagraphEntry / DiffItem / DiffResult）
  diff_engine.py        纯逻辑 diff 核心（不依赖 python-docx，可单测）
  docx_reader.py        读取 .docx 的薄适配层（唯一依赖 python-docx 处）
  reporting.py          结果组装 / 摘要 / 输出
  advanced.py           进阶分析（格式 / 图片 / 批注）
  cli.py                命令行入口
tests/
  test_diff_engine.py   核心逻辑单元测试
```

## 结果格式

`differences` 为三部分，每项含 `kind`（`added`/`removed`/`modified`）、`paragraph`（段落号，0-based）、`old_text`、`new_text`：

```json
{
  "status": "success",
  "doc1": {"path": "v1.docx", "name": "v1.docx", "paragraph_count": 12},
  "doc2": {"path": "v2.docx", "name": "v2.docx", "paragraph_count": 15},
  "is_identical": false,
  "summary": {"added": 3, "removed": 1, "modified": 2, "total": 6},
  "differences": {
    "added":    [{"kind": "added",    "paragraph": 10, "old_text": null, "new_text": "..."}],
    "removed":  [{"kind": "removed",  "paragraph": 4,  "old_text": "...", "new_text": null}],
    "modified": [{"kind": "modified", "paragraph": 2,  "old_text": "...", "new_text": "..."}]
  }
}
```

错误时统一返回 `{"status": "error", "error_type": "file_not_found|invalid_format|runtime_error", "message": "..."}`。

## 低 Token / 安全说明

- `--summary` 只输出摘要与统计（默认不开 `--advanced`），避免无条件把完整差异塞进模型上下文。
- 进阶分析独立成脚本/模块，默认关闭，按用户明确要求才调用。
- 日志：脚本不输出完整文档内容到日志，仅返回结构化结果；不涉及真实密钥/敏感凭据。

## 已知问题

- `python-docx` 不直接暴露段落级“删除标记（修订）”信息，本工具比较的是两份静态文档的快照差异，而非 Word 修订记录。
- 图片内容哈希依赖 `inline_shapes` 的关系部件读取，若图片以浮动/非内联方式插入可能无法提取，此时 `hashed=false`，不参与内容比较。
- 批注时间字段（创建/修改时间）在 `python-docx` 中不可用，批注比较仅基于作者与内容。
