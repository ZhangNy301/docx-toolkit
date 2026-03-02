"""
DOCX Toolkit 通用工具函数
"""
import os
import zipfile
import shutil
import tempfile
from pathlib import Path


def ensure_dir(path: str) -> str:
    """确保目录存在"""
    Path(path).mkdir(parents=True, exist_ok=True)
    return path


def unpack_docx(docx_path: str, output_dir: str = None) -> str:
    """
    解压 DOCX 文件

    Args:
        docx_path: DOCX 文件路径
        output_dir: 输出目录，默认为临时目录

    Returns:
        解压后的目录路径
    """
    if output_dir is None:
        output_dir = tempfile.mkdtemp(prefix='docx_unpack_')

    with zipfile.ZipFile(docx_path, 'r') as z:
        z.extractall(output_dir)

    return output_dir


def pack_docx(input_dir: str, output_path: str) -> str:
    """
    打包为 DOCX 文件

    Args:
        input_dir: 解压后的目录
        output_path: 输出文件路径

    Returns:
        输出文件路径
    """
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for root, dirs, files in os.walk(input_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, input_dir)
                z.write(file_path, arcname)

    return output_path


def get_media_files(docx_path: str) -> list:
    """
    获取 DOCX 中的媒体文件列表

    Args:
        docx_path: DOCX 文件路径

    Returns:
        媒体文件路径列表
    """
    media_files = []
    with zipfile.ZipFile(docx_path, 'r') as z:
        for name in z.namelist():
            if name.startswith('word/media/'):
                media_files.append(name)
    return media_files


def get_next_image_name(docx_path) -> tuple:
    """
    获取下一个图片名称和编号

    Args:
        docx_path: DOCX 文件路径或解压目录

    Returns:
        (新图片名称, 最大编号)
    """
    # 判断是文件还是目录
    if os.path.isfile(docx_path):
        # 解压到临时目录
        temp_dir = tempfile.mkdtemp(prefix='docx_media_')
        with zipfile.ZipFile(docx_path, 'r') as z:
            media_files = [f for f in z.namelist() if f.startswith('word/media/')]
            for f in media_files:
                z.extract(f, temp_dir)
        media_dir = os.path.join(temp_dir, 'word', 'media')
    else:
        media_dir = os.path.join(docx_path, 'word', 'media')

    if not os.path.exists(media_dir):
        return 'image1.png', 0

    existing = [f for f in os.listdir(media_dir) if f.startswith('image')]
    max_num = 0
    for img in existing:
        num = ''.join(filter(str.isdigit, os.path.splitext(img)[0]))
        if num:
            max_num = max(max_num, int(num))

    # 清理临时目录
    if os.path.isfile(docx_path) and os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)

    return f'image{max_num + 1}.png', max_num + 1


def validate_docx(path: str) -> bool:
    """
    验证文件是否为有效的 DOCX

    Args:
        path: 文件路径

    Returns:
        是否有效
    """
    if not os.path.exists(path):
        return False

    try:
        with zipfile.ZipFile(path, 'r') as z:
            return 'word/document.xml' in z.namelist()
    except:
        return False
