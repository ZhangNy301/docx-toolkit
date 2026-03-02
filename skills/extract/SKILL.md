---
name: docx-extract
description: 从 DOCX 文档中提取内容，包括文本、图片、表格等。当用户提到"提取"、"获取内容"、"导出"时触发。
---

# DOCX 内容提取

## 触发场景

- "提取所有文本"
- "导出图片"
- "获取表格数据"

## 支持的提取

### 提取文本

```python
from docx import Document

doc = Document("document.docx")
text = "\n".join([para.text for para in doc.paragraphs])
```

### 提取图片

```python
from scripts.core.utils import get_media_files
import zipfile

with zipfile.ZipFile("document.docx", 'r') as z:
    for name in z.namelist():
        if name.startswith('word/media/'):
            z.extract(name, "output/")
```

### 提取表格

```python
from docx import Document

doc = Document("document.docx")
for table in doc.tables:
    for row in table.rows:
        row_data = [cell.text for cell in row.cells]
        print(row_data)
```

## 命令

`/docx-toolkit:extract document.docx --type text --output content.txt`
