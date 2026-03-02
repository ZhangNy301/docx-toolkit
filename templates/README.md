# 内置模板

本目录包含预设的 Word 模板，可直接使用。

## 目录结构

```
templates/
├── academic/         # 学术模板
│   ├── paper.docx
│   └── grant-proposal.docx
└── business/         # 商业模板
    ├── contract.docx
    └── report.docx
```

## 学术模板

### paper.docx
学术论文模板，包含：
- 标题页（标题、作者、机构）
- 摘要
- 关键词
- 正文（带章节标题）
- 参考文献

**占位符**:
- `{{title}}` - 论文标题
- `{{author}}` - 作者
- `{{abstract}}` - 摘要
- `{{keywords}}` - 关键词

### grant-proposal.docx
基金申请书模板，包含：
- 立项依据
- 研究内容
- 研究方案
- 研究基础

## 商业模板

### contract.docx
合同模板

### report.docx
报告模板

## 使用方式

### 1. 使用内置模板

```python
from docxtpl import DocxTemplate

# 加载内置模板
template = DocxTemplate("templates/academic/paper.docx")

# 准备数据
context = {
    "title": "我的论文标题",
    "author": "张三",
    "abstract": "这是摘要...",
    "keywords": "关键词1, 关键词2"
}

# 渲染
template.render(context)
template.save("output.docx")
```

### 2. 创建自定义模板

1. 在 Word 中创建模板
2. 插入 Jinja2 占位符（如 `{{variable}}`）
3. 保存为 .docx 格式
4. 放入对应的模板目录

## 模板语法

### 变量
```
{{title}}
{{author}}
```

### 条件
```jinja2
{% if show_appendix %}
附录内容
{% endif %}
```

### 循环
```jinja2
{% for section in sections %}
{{section.title}}
{{section.content}}
{% endfor %}
```

## 贡献模板

如果您有好的模板想要贡献：
1. 确保模板使用标准占位符
2. 添加适当的文档注释
3. 提交到对应的目录

## 注意事项

1. 模板必须是有效的 .docx 格式
2. 复杂格式建议在模板中预设
3. 图片使用 InlineImage 对象
