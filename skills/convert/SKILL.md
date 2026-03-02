---
name: docx-convert
description: 格式转换，包括 DOCX 转 PDF、Markdown 等。当用户提到"转换"、"PDF"、"Markdown"时触发。
---

# DOCX 格式转换

## 触发场景

- "转换为 PDF"
- "转成 Markdown"
- "导出为 HTML"

## 支持的转换

| 源格式 | 目标格式 | 工具 |
|--------|----------|------|
| DOCX | PDF | LibreOffice |
| DOCX | Markdown | pandoc |
| Markdown | DOCX | pandoc |
| DOCX | HTML | pandoc |

## 使用方式

### DOCX 转 PDF

```bash
# 使用 LibreOffice
soffice --headless --convert-to pdf document.docx
```

### DOCX 转 Markdown

```bash
# 使用 pandoc
pandoc document.docx -o document.md
```

### Markdown 转 DOCX

```bash
pandoc document.md -o document.docx
```

## 命令

`/docx-toolkit:convert input.docx output.pdf --format pdf`
