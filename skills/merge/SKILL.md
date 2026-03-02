---
name: docx-merge
description: 合并多个 DOCX 文档为一个，保留各自格式。当用户提到"合并文档"、"整合章节"、"拼接文档"时自动触发。
---

# DOCX 文档合并

## 触发场景

- "合并这些文档"
- "把章节整合成一个文件"
- "拼接多个 Word 文件"
- 用户有多个需要合并的 .docx 文件

## 使用方式

### 基础合并

```python
from docxcompose.composer import Composer
from docx import Document

# 创建主文档
master = Document("chapter1.docx")
composer = Composer(master)

# 添加其他文档
composer.append(Document("chapter2.docx"))
composer.append(Document("chapter3.docx"))

# 保存
composer.save("merged.docx")
```

### 使用模板统一格式

```python
from scripts.core.merger import merge_with_template

# 使用模板的页面设置合并文档
merge_with_template(
    doc_paths=["ch1.docx", "ch2.docx", "ch3.docx"],
    template_path="template.docx",
    output_path="merged.docx"
)
```

## 命令

`/docx-toolkit:merge file1.docx file2.docx ... output.docx`

## 注意事项

1. 所有文档必须是有效的 .docx 格式
2. 合并后保留各自的格式和样式
3. 默认会在文档间添加分页符
4. 输出目录会自动创建
