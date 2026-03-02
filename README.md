# DOCX Toolkit

专业级 DOCX 文档处理 Claude Code Plugin。

## 功能

- **模板渲染**: 基于 Jinja2 的 Word 模板渲染
- **文档合并**: 合并多个 DOCX 文件，保留格式
- **精细编辑**: 插入图片、修改段落、添加表格等
- **批量操作**: 批量渲染、转换、提取
- **格式转换**: DOCX ↔ PDF ↔ Markdown
- **内容提取**: 提取文本、图片、表格

## 安装

```bash
claude --plugin-dir ./docx-toolkit
```

## 使用

- `/docx-toolkit:render` - 渲染模板
- `/docx-toolkit:merge` - 合并文档
- `/docx-toolkit:insert-image` - 插入图片
