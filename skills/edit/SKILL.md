---
name: docx-edit
description: 对 DOCX 文档进行精细编辑，包括插入图片、修改段落、添加表格等。当用户提到"插入图片"、"修改段落"、"添加表格"、"查找替换"时触发。
---

# DOCX 精细编辑

## 触发场景

- "插入图片到文档"
- "修改这段文字"
- "添加表格"
- "查找并替换"

## 支持的操作

### 图片操作
- 插入图片（支持指定位置和图注）
- 替换图片
- 调整图片大小

### 段落操作
- 插入段落
- 删除段落
- 查找替换文本

### 表格操作
- 创建表格
- 插入行列
- 填充数据

## 技术实现

使用 **unpack-edit-pack** 模式：

1. **Unpack**: 解压 DOCX 为 XML 文件
2. **Edit**: 直接编辑 XML
3. **Pack**: 重新打包为 DOCX

## 示例：插入图片

```python
from scripts.core.editor import insert_image

insert_image(
    docx_path="report.docx",
    image_path="figure.png",
    output_path="report_with_image.docx",
    position={"type": "after_text", "value": "1.2 研究意义"},
    caption="图1. 研究框架"
)
```

## 示例：查找替换

```python
from scripts.core.editor import find_and_replace

find_and_replace(
    docx_path="template.docx",
    replacements={
        "{{name}}": "张三",
        "{{date}}": "2024-03-01"
    },
    output_path="output.docx"
)
```

## 命令

`/docx-toolkit:insert-image document.docx image.png --position "after:1.2节" --caption "图1. 说明"`
