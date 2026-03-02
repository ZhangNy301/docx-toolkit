---
name: docx-template
description: 使用 Jinja2 模板渲染 Word 文档。当用户需要基于模板生成文档、填充数据到 Word 模板时自动触发。支持变量替换、条件渲染、循环渲染、动态图片等。
---

# DOCX 模板渲染

## 概述

基于 python-docx-template (docxtpl) 实现的 Word 模板渲染功能。核心理念：**用 Word 创建模板保留所有格式，只替换内容**。

## 触发场景

- 用户提到"模板"、"生成文档"、"填充数据"
- 用户有 .docx 模板文件和数据
- 用户需要批量生成相似文档

## 模板语法

### 基础变量

在 Word 模板中使用 `{{ variable }}` 作为占位符：

```
标题：{{ title }}
作者：{{ author }}
日期：{{ date }}
```

### 条件渲染

```jinja2
{% if show_header %}
这部分内容只在 show_header 为 True 时显示
{% endif %}
```

### 循环渲染

```jinja2
{% for item in items %}
- {{ item.name }}: {{ item.value }}
{% endfor %}
```

### 动态图片

在模板中放置 `{{ image_var }}`，渲染时传入图片路径。

## 使用方式

### 1. 准备模板

在 Word 中创建模板文件，插入 Jinja2 占位符，保存为 .docx。

### 2. 准备数据

数据可以是 JSON 文件或 Python 字典：

```json
{
  "title": "研究报告",
  "author": "张三",
  "items": [
    {"name": "项目A", "value": "100"},
    {"name": "项目B", "value": "200"}
  ]
}
```

### 3. 执行渲染

使用 `/docx-toolkit:render` 命令或直接描述需求。

## 实现脚本

当用户需要渲染模板时，生成以下 Python 脚本：

```python
from docxtpl import DocxTemplate

# 加载模板
doc = DocxTemplate("template.docx")

# 准备数据
context = {
    "title": "研究报告",
    "author": "张三",
    # ... 其他变量
}

# 渲染
doc.render(context)

# 保存
doc.save("output.docx")
```

## 图片处理

如需插入图片：

```python
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm

doc = DocxTemplate("template.docx")

# 创建图片对象
logo = InlineImage(doc, "logo.png", width=Mm(50))

context = {
    "title": "报告标题",
    "logo": logo
}

doc.render(context)
doc.save("output.docx")
```

## 批量渲染

```python
import json
from docxtpl import DocxTemplate

# 加载数据
with open("data.json") as f:
    data_list = json.load(f)

# 批量渲染
for i, data in enumerate(data_list):
    doc = DocxTemplate("template.docx")
    doc.render(data)
    doc.save(f"output_{i}.docx")
```

## 注意事项

1. 模板必须是有效的 .docx 文件
2. 占位符语法必须正确
3. 图片需要使用 InlineImage 对象
4. 复杂格式建议在模板中预设好
