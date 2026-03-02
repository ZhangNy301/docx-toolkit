---
name: docx-batch
description: 批量处理多个 DOCX 文档。当用户提到"批量"、"多个文档"、"批量处理"时触发。
---

# DOCX 批量操作

## 触发场景

- "批量生成文档"
- "处理多个文件"
- "批量转换"

## 支持的操作

### 批量模板渲染

```python
from scripts.core.template_engine import render_batch

data_list = [
    {"name": "张三", "title": "报告1"},
    {"name": "李四", "title": "报告2"},
    {"name": "王五", "title": "报告3"},
]

render_batch(
    template_path="template.docx",
    data_list=data_list,
    output_dir="output/",
    filename_pattern="{name}_{title}.docx"
)
```

### 批量格式转换

将目录下所有 DOCX 转换为 PDF。

### 批量内容提取

提取多个文档的文本或图片。

## 命令

`/docx-toolkit:batch-render template.docx data.json output/`
