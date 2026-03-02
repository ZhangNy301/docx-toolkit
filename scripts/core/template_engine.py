"""
DOCX 模板渲染引擎

基于 docxtpl 实现 Jinja2 模板渲染
"""
import os
from typing import Dict, Any, Optional

from docxtpl import DocxTemplate
from docxtpl import InlineImage
from docx.shared import Mm

from .utils import validate_docx


def render_template(
    template_path: str,
    context: Dict[str, Any],
    output_path: str,
    image_mappings: Optional[Dict[str, str]] = None
) -> str:
    """
    渲染 Word 模板

    Args:
        template_path: 模板文件路径
        context: 模板变量字典
        output_path: 输出文件路径
        image_mappings: 图片映射 {变量名: 图片路径}

    Returns:
        输出文件路径
    """
    # 验证模板
    if not validate_docx(template_path):
        raise ValueError(f"无效的 DOCX 模板: {template_path}")

    # 确保输出目录存在
    output_dir = os.path.dirname(output_path)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)

    # 加载模板
    doc = DocxTemplate(template_path)

    # 处理图片映射
    if image_mappings:
        for var_name, img_path in image_mappings.items():
            if os.path.exists(img_path):
                context[var_name] = InlineImage(
                    doc,
                    img_path,
                    width=Mm(150)
                )

    # 渲染
    doc.render(context)

    # 保存
    doc.save(output_path)

    return output_path


def create_inline_image(
    doc: DocxTemplate,
    image_path: str,
    width_mm: float = 150,
    height_mm: float = None
) -> InlineImage:
    """创建内联图片对象"""
    if height_mm:
        return InlineImage(
            doc,
            image_path,
            width=Mm(width_mm),
            height=Mm(height_mm)
        )
    else:
        return InlineImage(doc, image_path, width=Mm(width_mm))


def render_batch(
    template_path: str,
    data_list: list,
    output_dir: str,
    filename_pattern: str = "output_{index}.docx"
) -> list:
    """批量渲染模板"""
    os.makedirs(output_dir, exist_ok=True)
    output_paths = []

    for index, context in enumerate(data_list):
        filename = filename_pattern.format(index=index, **context)
        output_path = os.path.join(output_dir, filename)
        render_template(template_path, context, output_path)
        output_paths.append(output_path)

    return output_paths
