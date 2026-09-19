# AI Word 版本比较

比较两份 `.docx` 文档的新增 / 删除 / 修改内容，输出结构化 JSON 差异结果，
并生成人类可读的中文修改摘要。用于“用户自然语言 → 智能体 → Skill → Python Script/CLI → 开源项目 → 结果”链路中的确定性能力层。

## 环境要求

- **Python：实验环境 3.12**（兼容 3.10 ~ 3.13；`pyproject.toml` 声明 `requires-python = ">=3.10,<3.14"`，任务书建议范围 3.10 ~ 3.12）。
- 依赖：`python-docx`（见 `requirements.txt` / `pyproject.toml`）。

## 用户场景（支持的 3 类核心意图）

1. **对比差异**：“帮我对比 v1.docx 和 v2.docx，看改了哪些内容。”
2. **生成修改摘要**：“总结一下 v1.docx 到 v2.docx 的主要改动。”
3. **追问详情**：“v2 里新增了哪些段落？第 3 段改了什么？哪几个字被加粗了？”
   （基础层回答段落增删改；`--advanced` 回答格式 / 图片 / 批注 / 空段落）

## 参考开源项目

- **python-docx**：读取 `.docx` 段落、格式、内嵌图片与批注。
- 许可证：MIT License（pydocx/python-docx 项目）。
- 实际使用方式：`docx.Document(path)` 读取段落文本与内联图片；`doc.comments` 读取批注；`shape.part.related_parts[rId]` 读取图片二进制以计算哈希；`document.styles.element` 直接读 `w:docDefaults` 以解析样式继承。

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

# 让空段落也参与基础比较（默认忽略空段落以降噪）
python word_diff.py v1.docx v2.docx --keep-empty --summary

# 进阶分析：额外输出格式 / 图片 / 批注 / 空段落差异
python word_diff.py v1.docx v2.docx --advanced

# 或独立运行进阶脚本（可调修改段落配对的相似度阈值）
python word_diff_advanced.py v1.docx v2.docx --pretty --similarity 0.6
```

> 测试方法见下一节「[测试](#测试)」。

## 项目结构

```
word_diff.py            基础比较入口（兼容旧用法）
word_diff_advanced.py   进阶比较入口（按需调用）
word_diff/
  __init__.py           包导出
  models.py             数据模型（ParagraphEntry / DiffItem / DiffResult）
  diff_engine.py        纯逻辑 diff 核心（不依赖 python-docx，可单测）
  docx_reader.py        读取 .docx 的薄适配层
  reporting.py          结果组装 / 摘要 / 输出
  advanced.py           进阶分析（格式 / 图片 / 批注 / 空段落）
  cli.py                命令行入口
tests/
  README.md               测试目录说明（跑法 / 覆盖范围 / 如何加用例）
  conftest.py             把项目根加入 sys.path（不写死机器路径）
  helpers.py              测试工具：造 .docx / 生成 PNG / 跑 CLI
  nl_cases.py             自然语言用例表（3 类意图，加用例只加一行）
  test_diff_engine.py     第1层 单元：diff 引擎（不依赖 python-docx）
  test_docx_reader.py     第2层 集成：读取层与路径校验
  test_cli_intents.py     第3层 意图：自然语言 → 命令 → 结果契约
  test_advanced.py        第4层 进阶：格式 / 图片 / 批注 / 空段落
  test_cli_errors.py      第5层 异常：统一错误契约与退出码
  test_real_documents.py  可选：用环境变量指定的真实文档回归
  run_tests.py            不依赖 pytest 的兜底运行器
```

## 测试

共 **50 个用例**，分 5 层；**不写死任何机器路径**，测试用的 .docx 与图片现场生成在系统临时目录，测完自动清理。

```bash
# 方式一（推荐，需先装 pytest）
pip install pytest
python -m pytest tests -q

# 方式二（不需要 pytest，只用标准库）
python tests/run_tests.py

# 方式三：只跑某一层
python tests/run_tests.py diff_engine    # 单元层（不依赖 python-docx）
python tests/run_tests.py cli_intents    # 自然语言意图层
```

用**你自己的真实文档**再跑一遍（路径放环境变量，不进代码；未设置时自动跳过，换设备不会失败）：

```powershell
$env:WORD_DIFF_REAL_DOC1="D:\...\v1.docx"
$env:WORD_DIFF_REAL_DOC2="D:\...\v2.docx"
python tests/run_tests.py real_documents
```

覆盖范围：

- **3 类自然语言意图**：对比差异 / 生成修改摘要 / 追问详情（格式 · 批注 · 空段落 · 图片）；
- **正常、边界、异常**三类输入：文件不存在、目录、非 .docx、损坏文件、参数不足、空文档等；
- **两条历史缺陷的回归**：段落被插入/删除后格式差异不再漏报；run 切分方式不同不再误报。

测试目录的完整说明见 [`tests/README.md`](tests/README.md)。

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

### `--advanced` 追加字段

`advanced` 下含 `format` / `images` / `comments` / `blank_paragraphs` / `summary_text`。
其中格式差异**按段落内容对齐**（不是按段落序号），并**逐字符**给出变化范围：

```json
{
  "format": {
    "format_changes": [
      {
        "paragraph": 3, "paragraph_new": 4, "kind": "property_change",
        "char_range": [8, 49], "length": 41,
        "text": "代工业化前，二战无线通信，1946年第一台计算机诞生，…",
        "detail": ["粗体: False → True"],
        "props": {"bold": {"old": false, "new": true, "raw_old": null, "raw_new": true,
                           "source_old": "unresolved", "source_new": "run"}},
        "run_count": [9, 12], "certainty": "confirmed"
      }
    ],
    "count": 1,
    "run_structure_changes": [{"paragraph": 3, "paragraph_new": 4,
      "detail": "格式片段数量由 9 变为 12（已改为逐字符比较，此项仅作提示）"}],
    "style_expression_changes": [],
    "text_changes": [],
    "alignment": {"method": "content(SequenceMatcher equal) + similarity-greedy(replace)",
                  "paired_paragraphs": 96, "doc1_only_paragraphs": [], "doc2_only_paragraphs": [1]}
  },
  "blank_paragraphs": {
    "doc1": {"count": 2, "empty": 2, "whitespace_only": 0, "indices": [1, 95]},
    "doc2": {"count": 3, "empty": 3, "whitespace_only": 0, "indices": [1, 2, 96]},
    "blank_diff": [{"kind": "added", "paragraph": 1,
                    "detail": "doc2 第1段为新增：真空段落（无任何字符，也没有文本片段）",
                    "candidates": [1, 2], "position_ambiguous": true,
                    "context": {"before": "三、简答题", "after": "简要说明信息技术的发展过程…"}}],
    "changed": true
  }
}
```

字段口径：

| 字段 | 含义 |
| --- | --- |
| `format.count` | **真实格式变化处数**（`property_change`），run 切分差异不计入 |
| `char_range` | 变化字符范围，0-based 左闭右开 |
| `props.*.old/new` | 沿样式链解析后的**生效值**（判定是否变化以此为准） |
| `props.*.raw_old/raw_new` | run 上**直接设置**的原值（`null` 表示未直接设置） |
| `props.*.source_*` | 生效值来源：`run` / `style:样式名` / `docDefaults` / `unresolved` |
| `certainty` | `confirmed`；当字体名/字号一侧来自主题无法解析时为 `inheritance_unknown` |
| `run_structure_changes` | 仅 run 切分变化（渲染无差异）的提示项，`structure_count` 单独计数 |
| `style_expression_changes` | 写法不同（如显式 `False` vs 未设置）但生效值一致，不算格式变化 |
| `text_changes` | 文本不同的区间，不做格式比较 |
| `blank_paragraphs.*.empty` | 真空段落（0 字符、通常 0 个 run） |
| `blank_paragraphs.*.whitespace_only` | 仅含空白字符的段落，`items[].codepoints` 给出 U+0020 空格 / U+3000 全角空格 / U+0009 制表符等 |
| `position_ambiguous` / `candidates` | 连续多个空段落时，增删位置不唯一，给出候选段号与前后上下文 |

## 低 Token / 安全说明

- `--summary` 只输出摘要与统计，避免无条件把完整差异塞进模型上下文。
- 进阶分析独立成脚本/模块，默认关闭，按用户明确要求才调用。
- 日志：脚本不输出完整文档内容到日志，仅返回结构化结果；不涉及真实密钥/敏感凭据。

## 已知问题

- `python-docx` 不直接暴露段落级“删除标记（修订）”信息，本工具比较的是两份静态文档的快照差异，而非 Word 修订记录。
- 图片通过 `doc.inline_shapes` 读取数量与尺寸；内联图片的二进制经文档部件 `related_parts` 按关系 ID 取出后计算 MD5。**浮动（非内联）图片不在 `inline_shapes` 中**，因此不参与数量/尺寸/内容比较。
- 批注时间字段（创建/修改时间）在 `python-docx` 中不可用，批注比较仅基于作者与内容。
- 格式比较基于 `paragraph.runs`，**超链接内部的 run 不在其中**，因此超链接内的文字与格式不参与比较。
- 段落文本相同、仅插入内联图片时，承载图片的 run（底层 XML 含 `<w:drawing>`）会被 `len(paragraph.runs)` 一并计入“格式片段数量”，导致 `run_structure_changes` 多报一条 `kind: "structure_change"`（`structure_count` +1）。这仅是提示项：图片 run 的 `text` 为空，逐字符格式比较与 `text_changes` 不受影响；彻底修复需在统计 run 数量时跳过含 `<w:drawing>` 的图片 run，计划后续迭代完成。
- 表格单元格内的文字不参与比较（只比正文段落）。
- 段落的字体名/字号若完全来自主题（`unresolved`），无法与显式设置的同值区分，此时会给出 `certainty = inheritance_unknown`；布尔属性（粗体/斜体/下划线/删除线）已按 Word 默认值兜底，不受此影响。
- 连续多个空段落之间的增删在位置上不唯一，`blank_diff` 会给出 `position_ambiguous` + `candidates`。
- `add_picture` 等产生的**纯图片段落没有文字**，在 `blank_paragraphs` 分析中会被计为“真空段落”（0 字符），因此在新增/删除图片时会额外出现“空段落差异”；阅读结果时需与真正的回车空行区分。
