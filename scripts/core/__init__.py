"""
DOCX Toolkit Core Module
"""

from .utils import (
    ensure_dir,
    unpack_docx,
    pack_docx,
    get_media_files,
    get_next_image_name,
    validate_docx,
)

from .template_engine import (
    render_template,
    create_inline_image,
    render_batch,
)

from .editor import (
    insert_image,
    find_and_replace,
    insert_paragraph,
    delete_paragraph,
    get_document_text,
    find_text,
)

__all__ = [
    # Utils
    'ensure_dir',
    'unpack_docx',
    'pack_docx',
    'get_media_files',
    'get_next_image_name',
    'validate_docx',
    # Template Engine
    'render_template',
    'create_inline_image',
    'render_batch',
    # Editor
    'insert_image',
    'find_and_replace',
    'insert_paragraph',
    'delete_paragraph',
    'get_document_text',
    'find_text',
]
