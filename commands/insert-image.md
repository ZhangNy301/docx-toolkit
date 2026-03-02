---
description: 在文档中插入图片
---

# 插入图片

在 Word 文档的指定位置插入图片，支持添加图注。

**用法**: `/docx-toolkit:insert-image <文档路径> <图片路径> [选项]`

**选项**:
- `--position`: 插入位置 (如 "after:1.2节")
- `--caption`: 图注文本
- `--width`: 图片宽度（毫米）

**示例**:
```
/docx-toolkit:insert-image report.docx figure.png --caption "图1. 研究框架"
```

## 实现步骤

1. 解析参数和选项
2. 生成 Python 脚本:

```python
from scripts.core.editor import insert_image

# 解析参数
args = "$ARGUMENTS".split()
docx_path = args[0]
image_path = args[1]
output_path = docx_path.replace('.docx', '_with_image.docx')

# 解析选项
caption = None
position = None
width_mm = 150

i = 2
while i < len(args):
    if args[i] == '--caption' and i + 1 < len(args):
        caption = args[i + 1]
        i += 2
    elif args[i] == '--position' and i + 1 < len(args):
        pos_value = args[i + 1]
        if pos_value.startswith('after:'):
            position = {'type': 'after_text', 'value': pos_value[6:]}
        elif pos_value.startswith('before:'):
            position = {'type': 'before_text', 'value': pos_value[7:]}
        i += 2
    elif args[i] == '--width' and i + 1 < len(args):
        width_mm = float(args[i + 1])
        i += 2
    else:
        i += 1

# 插入图片
insert_image(docx_path, image_path, output_path, position, caption, width_mm)

print(f"✓ 图片已插入到: {output_path}")
```

3. 执行脚本
4. 验证输出

## 位置说明

- `after:文本` - 在指定文本之后插入
- `before:文本` - 在指定文本之前插入
- 如果不指定位置，默认在文档末尾插入
