---
description: 合并多个 DOCX 文档
---

# DOCX 文档合并

合并多个 Word 文档为一个，保留各自格式。

**用法**: `/docx-toolkit:merge <文件1> <文件2> ... <输出文件>`

**参数**:
- 文件1, 文件2, ...: 要合并的 DOCX 文件列表
- 输出文件: 合并后的文件路径

**示例**:
```
/docx-toolkit:merge chapter1.docx chapter2.docx chapter3.docx merged.docx
```

## 实现步骤

1. 解析参数
2. 生成 Python 脚本:

```python
from docxcompose.composer import Composer
from docx import Document

# 解析参数
args = "$ARGUMENTS".split()
doc_paths = args[:-1]
output_path = args[-1]

# 创建主文档
master = Document(doc_paths[0])
composer = Composer(master)

# 添加其他文档
for path in doc_paths[1:]:
    composer.append(Document(path))

# 保存
composer.save(output_path)

print(f"✓ 已合并 {len(doc_paths)} 个文档到: {output_path}")
```

3. 执行脚本
4. 验证输出

## 注意事项

- 所有文件必须是有效的 .docx 格式
- 合并后保留各自的格式
- 默认在文档间添加分页符
