---
description: 基于模板渲染 Word 文档
---

# DOCX 模板渲染

渲染 Word 模板，替换占位符为实际数据。

**用法**: `/docx-toolkit:render <模板路径> <数据文件> <输出路径>`

**参数**:
- 模板路径: 包含 Jinja2 占位符的 .docx 文件
- 数据文件: JSON 格式的数据文件
- 输出路径: 生成的文档路径

**示例**:
```
/docx-toolkit:render template.docx data.json output.docx
```

## 实现步骤

1. 读取模板和数据
2. 生成 Python 脚本:

```python
import json
from docxtpl import DocxTemplate

# 解析参数
args = "$ARGUMENTS".split()
template_path = args[0]
data_path = args[1]
output_path = args[2]

# 加载数据
with open(data_path, encoding='utf-8') as f:
    context = json.load(f)

# 渲染模板
doc = DocxTemplate(template_path)
doc.render(context)
doc.save(output_path)

print(f"✓ 文档已生成: {output_path}")
```

3. 执行脚本
4. 验证输出

## 数据格式

JSON 文件示例:
```json
{
  "title": "研究报告",
  "author": "张三",
  "date": "2024-03-01",
  "items": [
    {"name": "项目A", "value": 100},
    {"name": "项目B", "value": 200}
  ]
}
```
