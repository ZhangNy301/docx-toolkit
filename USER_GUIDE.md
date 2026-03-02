# docx-toolkit 用户使用指南

## 简介

**docx-toolkit** 是一个专业的 DOCX 文档处理 Claude Code Plugin，提供模板渲染、文档合并、精细编辑、批量操作等完整功能。

**核心理念**：用 Word 模板保留所有格式，只替换内容。

---

## 安装与加载

### 1. 安装依赖

```bash
cd /Users/ningyizhang/tmp/rPPG-Fund/docx-toolkit
.venv/bin/pip install -r requirements.txt
```

依赖包：
- `docxtpl>=0.16.0` - Jinja2 模板渲染
- `python-docx>=0.8.11` - DOCX 基础操作
- `docxcompose>=1.4.0` - 文档合并
- `Jinja2>=3.0.0` - 模板引擎
- `Pillow>=9.0.0` - 图片处理

### 2. 加载 Plugin

**重要**：必须在**源码目录之外**启动 Claude Code 才能正确加载 Plugin：

```bash
# 错误方式（在源码目录内会优先加载内置 skill）
cd /Users/ningyizhang/tmp/rPPG-Fund/docx-toolkit
claude

# 正确方式（在任意其他目录）
cd ~/Desktop
claude --plugin-dir /Users/ningyizhang/tmp/rPPG-Fund/docx-toolkit
```

加载成功后，你会看到类似提示：
```
Loaded plugin: docx-toolkit v1.0.0
```

---

## 功能一：模板渲染

### 功能说明

基于 Jinja2 语法在 Word 模板中填充数据，保留所有格式设置。

### 支持的 Jinja2 语法

| 语法 | 说明 | 示例 |
|------|------|------|
| `{{ var }}` | 变量替换 | `{{ title }}` |
| `{% if %}` | 条件渲染 | `{% if show %}...{% endif %}` |
| `{% for %}` | 循环渲染 | `{% for item in items %}...{% endfor %}` |
| `{{ img }}` | 动态图片 | 传入 InlineImage 对象 |

### Prompt 示例

#### 基础渲染

```
我有一个 Word 模板需要填充数据。

模板路径：/Users/ningyizhang/tmp/rPPG-Fund/docx-toolkit/examples/template_proper.docx
数据路径：/Users/ningyizhang/tmp/rPPG-Fund/docx-toolkit/examples/data.json
输出路径：/Users/ningyizhang/tmp/rPPG-Fund/docx-toolkit/examples/output.docx

请使用 docx-toolkit 的模板渲染功能生成文档。
要求：
1. 使用 docxtpl 进行 Jinja2 渲染
2. 保留模板中的标题样式和字体
3. 确保中文字体正确显示
```

#### 带图片的渲染

```
我需要生成一份带封面的报告。

模板：report_template.docx（包含 {{ logo }} 图片占位符）
数据：{
  "title": "年度报告",
  "author": "张三",
  "date": "2024-03-01"
}
Logo图片：company_logo.png

请渲染模板，将 logo 图片插入到封面，输出到 annual_report.docx。
```

#### 批量渲染

```
我有 10 份员工数据需要生成合同。

模板：contract_template.docx
数据文件：employees.json（包含 10 个员工对象的数组）
输出目录：contracts/
命名规则：合同_{name}_{date}.docx

请批量渲染生成 10 份合同文档。
```

### Python API

```python
from scripts.core.template_engine import render_template, render_batch

# 单文档渲染
render_template(
    template_path="template.docx",
    context={"title": "报告", "author": "张三"},
    output_path="output.docx",
    image_mappings={"logo": "logo.png"}  # 可选：图片映射
)

# 批量渲染
data_list = [
    {"name": "张三", "title": "合同1"},
    {"name": "李四", "title": "合同2"}
]
render_batch(
    template_path="template.docx",
    data_list=data_list,
    output_dir="output/",
    filename_pattern="{name}_{title}.docx"
)
```

---

## 功能二：文档合并

### 功能说明

将多个 DOCX 文档合并为一个，保留各自格式，自动添加分页符。

### Prompt 示例

#### 基础合并

```
我有三个章节需要合并成完整报告。

章节文件：
- /Users/ningyizhang/tmp/rPPG-Fund/docx-toolkit/examples/chapter1.docx
- /Users/ningyizhang/tmp/rPPG-Fund/docx-toolkit/examples/chapter2.docx
- /Users/ningyizhang/tmp/rPPG-Fund/docx-toolkit/examples/chapter3.docx

输出：/Users/ningyizhang/tmp/rPPG-Fund/docx-toolkit/examples/merged.docx

请使用 docx-toolkit 的合并功能，要求：
1. 保留各章节的原有格式
2. 章节之间添加分页符
3. 使用 docxcompose 确保格式不丢失
```

#### 使用模板统一格式

```
我需要合并多个章节，并统一应用格式模板。

章节：chapter1.docx, chapter2.docx, chapter3.docx
格式模板：format_template.docx（包含统一的页面设置、页边距）
输出：unified_report.docx

请使用 merge_with_template 功能，应用模板的页面设置到合并后的文档。
```

### Python API

```python
from scripts.core.merger import merge_documents, merge_with_template

# 基础合并
merge_documents(
    doc_paths=["ch1.docx", "ch2.docx", "ch3.docx"],
    output_path="merged.docx",
    page_break=True,      # 章节间添加分页符
    preserve_format=True  # 保留格式
)

# 带模板格式的合并
merge_with_template(
    doc_paths=["ch1.docx", "ch2.docx"],
    template_path="template.docx",
    output_path="unified.docx"
)
```

---

## 功能三：精细编辑

### 功能说明

在 Word 文档中插入图片、查找替换文本，使用 unpack-edit-pack 模式确保文档结构完整。

### Prompt 示例

#### 插入图片

```
我需要在报告中插入一张示意图。

源文档：/Users/ningyizhang/tmp/rPPG-Fund/docx-toolkit/examples/output_rendered.docx
图片：/Users/ningyizhang/tmp/rPPG-Fund/docx-toolkit/examples/figure.png
输出：/Users/ningyizhang/tmp/rPPG-Fund/docx-toolkit/examples/with_figure.docx

要求：
1. 在"摘要"段落之后插入图片
2. 图片居中显示
3. 添加图注："图1. 研究框架示意图"
4. 图片宽度约 150mm
5. 确保文档结构完整，Word 能正常打开
```

#### 查找替换

```
我需要批量替换文档中的占位符。

文档：template.docx
替换规则：
- "{{company}}" → "科技有限公司"
- "{{date}}" → "2024年3月"
- "{{version}}" → "v2.0"

输出：filled.docx
```

### Python API

```python
from scripts.core.editor import insert_image, find_and_replace

# 插入图片
insert_image(
    docx_path="report.docx",
    image_path="figure.png",
    output_path="with_image.docx",
    position={"type": "after_text", "value": "摘要"},
    caption="图1. 研究框架",
    width_mm=150
)

# 查找替换
find_and_replace(
    docx_path="template.docx",
    replacements={
        "{{name}}": "张三",
        "{{date}}": "2024-03-01"
    },
    output_path="filled.docx"
)
```

---

## 功能四：内容提取

### 功能说明

从 DOCX 文档中提取文本、图片、表格等内容。

### Prompt 示例

```
我需要从文档中提取所有内容。

文档：report.docx
提取类型：
1. 所有文本内容，保存为 report_text.txt
2. 所有图片，保存到 images/ 目录
3. 所有表格数据，保存为表格.csv

请使用 docx-toolkit 的提取功能完成。
```

### Python API

```python
from docx import Document
from scripts.core.utils import get_media_files
import zipfile

# 提取文本
doc = Document("report.docx")
text = "\n".join([p.text for p in doc.paragraphs])

# 提取图片
with zipfile.ZipFile("report.docx", 'r') as z:
    for name in z.namelist():
        if name.startswith('word/media/'):
            z.extract(name, "output/")

# 提取表格
doc = Document("report.docx")
for table in doc.tables:
    for row in table.rows:
        row_data = [cell.text for cell in row.cells]
        print(row_data)
```

---

## 功能五：格式转换

### 功能说明

将 DOCX 转换为其他格式（PDF、Markdown、HTML）或反向转换。

### Prompt 示例

```
我需要将文档转换为不同格式。

源文档：report.docx
转换目标：
1. PDF 版本：report.pdf
2. Markdown：report.md
3. HTML：report.html

请使用适当的工具（pandoc 或 LibreOffice）完成转换。
```

### 命令行工具

```bash
# DOCX → PDF
soffice --headless --convert-to pdf document.docx

# DOCX → Markdown
pandoc document.docx -o document.md

# Markdown → DOCX
pandoc document.md -o document.docx --reference-doc=template.docx

# DOCX → HTML
pandoc document.docx -o document.html
```

---

## 功能六：批量操作

### 功能说明

批量处理多个文档，包括批量渲染、批量转换、批量提取等。

### Prompt 示例

```
我需要批量处理目录中的所有文档。

输入目录：/Users/ningyizhang/Documents/reports/
操作：
1. 将所有 chapter_*.docx 文件合并成一个 complete.docx
2. 生成 PDF 版本 complete.pdf
3. 提取所有图片到 images/ 目录

输出目录：/Users/ningyizhang/Documents/output/
```

---

## Slash 命令速查

| 命令 | 功能 | 示例 |
|------|------|------|
| `/docx-toolkit:render` | 模板渲染 | `/docx-toolkit:render template.docx data.json output.docx` |
| `/docx-toolkit:merge` | 文档合并 | `/docx-toolkit:merge ch1.docx ch2.docx merged.docx` |
| `/docx-toolkit:insert-image` | 插入图片 | `/docx-toolkit:insert-image doc.docx img.png --caption "图1"` |

---

## 模板制作指南

### 1. 创建模板文档

在 Word 中新建文档，按以下原则设计：

**格式预设**：
- 标题使用「标题1」「标题2」等内置样式
- 正文使用「正文」样式
- 设置好字体（建议中文字体使用宋体、黑体）
- 设置好段落间距、行距

**插入占位符**：
```
标题：{{ title }}
作者：{{ author }}
日期：{{ date }}

{% for section in sections %}
{{ section.title }}
{{ section.content }}
{% endfor %}
```

**插入图片占位符**：
```
{{ logo }}
```

### 2. 准备数据文件

JSON 格式示例：
```json
{
  "title": "人工智能研究报告",
  "author": "张三",
  "date": "2024-03-01",
  "sections": [
    {"title": "引言", "content": "人工智能正在改变..."},
    {"title": "方法", "content": "我们采用深度学习..."}
  ]
}
```

### 3. 测试模板

保存模板后，先用简单数据测试渲染，确保格式正确。

---

## 常见问题

### Q1: Word 打开生成的文档提示"发现无法读取的内容"？

**原因**：XML 结构被破坏（如图片插到了 `<w:sectPr>` 之后）

**解决**：确保使用 `scripts/core/editor.py` 的 `insert_image` 函数，它会正确处理文档结构。

### Q2: 中文字体显示异常？

**原因**：python-docx 默认字体不支持中文

**解决**：在模板中显式设置中文字体：
```python
def set_chinese_font(run, font_name='宋体', size=12):
    run.font.name = font_name
    run._element.rPr.rFonts.set(docx.oxml.ns.qn('w:eastAsia'), font_name)
    run.font.size = Pt(size)
```

### Q3: Plugin 没有生效？

**原因**：在源码目录内启动会优先加载内置 skill

**解决**：在任意其他目录启动 Claude Code：
```bash
cd ~/Desktop
claude --plugin-dir /Users/ningyizhang/tmp/rPPG-Fund/docx-toolkit
```

### Q4: 合并后格式不一致？

**原因**：各文档使用不同的页面设置

**解决**：使用 `merge_with_template` 统一应用模板的页面设置。

---

## 完整示例工作流

### 场景：生成项目报告

**步骤 1：创建模板**
```python
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import docx.oxml.ns

doc = Document()

# 封面
title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title.add_run('{{ project_name }}')
run.font.size = Pt(24)
run.font.bold = True
run.font.name = '黑体'
run._element.rPr.rFonts.set(docx.oxml.ns.qn('w:eastAsia'), '黑体')

# 章节循环
doc.add_paragraph('{% for section in sections %}')

heading = doc.add_paragraph()
heading.style = doc.styles['Heading 1']
run = heading.add_run('{{ section.title }}')

content = doc.add_paragraph('{{ section.content }}')

doc.add_paragraph('{% endfor %}')

doc.save('project_template.docx')
```

**步骤 2：准备数据**
```json
{
  "project_name": "智能监控系统",
  "sections": [
    {"title": "项目背景", "content": "随着城市化进程..."},
    {"title": "技术方案", "content": "采用深度学习算法..."},
    {"title": "预期成果", "content": "实现 95% 以上准确率..."}
  ]
}
```

**步骤 3：渲染模板**
```
请渲染模板：
- 模板：project_template.docx
- 数据：project_data.json
- 输出：project_report.docx
```

**步骤 4：插入图片**
```
在报告的"技术方案"章节后插入系统架构图。
- 文档：project_report.docx
- 图片：architecture.png
- 图注："图1. 系统架构图"
- 输出：project_report_final.docx
```

---

## 参考资源

- [docxtpl 文档](https://docxtpl.readthedocs.io/)
- [python-docx 文档](https://python-docx.readthedocs.io/)
- [Jinja2 模板语法](https://jinja.palletsprojects.com/)
- [Claude Code Plugin 开发指南](https://claude.ai/docs/plugin-development)
