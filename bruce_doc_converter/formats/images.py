"""Images conversion helpers."""
from .common import (
    IMAGE_OUTPUT_DIR_NAME,
    MAX_ASPECT_RATIO,
    MIN_IMAGE_DATA_BYTES,
    MIN_IMAGE_DIMENSION_PX,
    OOXML_IMAGE_NAMESPACES,
    _IMAGE_SIGNATURES,
)
import os
import struct
import tempfile
import logging

logger = logging.getLogger(__name__)


def _detect_image_format(data):
    """通过文件头魔数识别图片格式，返回扩展名（不含点）或 None"""
    if not data or len(data) < 8:
        return None
    for signature, fmt in _IMAGE_SIGNATURES.items():
        if data[:len(signature)] == signature:
            # EMF 需要额外验证：前 4 字节为 \x01\x00\x00\x00 且偏移 40 处有 ' EMF' 标记
            if fmt == 'emf' and len(data) >= 44:
                if data[40:44] != b' EMF':
                    continue
            return fmt
    return None


def _get_image_dimensions(data):
    """
    从图片二进制数据中解析宽高（像素）。
    不依赖 PIL，仅通过文件头解析。
    对于不支持解析的格式返回 (None, None)。
    """
    if not data or len(data) < 8:
        return None, None

    fmt = _detect_image_format(data)
    if fmt is None:
        return None, None

    try:
        if fmt == 'png':
            # PNG: IHDR chunk 位于文件头之后，偏移 16 处为宽高各 4 字节大端
            if len(data) >= 24:
                width = struct.unpack('>I', data[16:20])[0]
                height = struct.unpack('>I', data[20:24])[0]
                return width, height

        elif fmt == 'jpeg':
            # JPEG: 扫描 SOF marker (0xFF 0xC0-0xCF, 排除 0xC4/0xC8/0xCC)
            offset = 2
            while offset < len(data) - 9:
                if data[offset] != 0xFF:
                    break
                marker = data[offset + 1]
                if marker == 0xD9:  # EOI
                    break
                if marker == 0xDA:  # SOS - 数据流开始，停止扫描
                    break
                length = struct.unpack('>H', data[offset + 2:offset + 4])[0]
                # SOF markers: 0xC0-0xCF 但排除 DHT(0xC4)、JPG(0xC8)、DAC(0xCC)
                if 0xC0 <= marker <= 0xCF and marker not in (0xC4, 0xC8, 0xCC):
                    if offset + 9 <= len(data):
                        height = struct.unpack('>H', data[offset + 5:offset + 7])[0]
                        width = struct.unpack('>H', data[offset + 7:offset + 9])[0]
                        return width, height
                offset += 2 + length

        elif fmt == 'gif':
            # GIF: 宽高位于偏移 6 处，各 2 字节小端
            if len(data) >= 10:
                width = struct.unpack('<H', data[6:8])[0]
                height = struct.unpack('<H', data[8:10])[0]
                return width, height

        elif fmt == 'bmp':
            # BMP: 宽高位于 DIB header 中，偏移 18 处各 4 字节小端（有符号）
            if len(data) >= 26:
                width = struct.unpack('<i', data[18:22])[0]
                height = abs(struct.unpack('<i', data[22:26])[0])
                return width, height

        elif fmt == 'tiff':
            # TIFF 解析复杂，跳过尺寸检测
            pass

    except (struct.error, IndexError):
        pass

    return None, None


def _is_decorative_image(data, width=None, height=None, is_decorative_flag=False,
                         is_pptx_background=False):
    """
    综合判定图片是否为装饰性/无意义图片。

    Args:
        data: 图片二进制数据
        width: 图片宽度（像素），None 时自动检测
        height: 图片高度（像素），None 时自动检测
        is_decorative_flag: Office 文档中 adec:decorative 标记
        is_pptx_background: PowerPoint 中覆盖整个幻灯片的背景图

    Returns:
        True 表示应过滤掉此图片
    """
    # 1. Office 自身的装饰性标记（最可靠）
    if is_decorative_flag:
        return True

    # 2. PowerPoint 全屏背景图
    if is_pptx_background:
        return True

    if not data:
        return True

    # 3. 数据量极小（纯色/透明占位）
    if len(data) <= MIN_IMAGE_DATA_BYTES:
        return True

    # 自动检测尺寸
    if width is None or height is None:
        width, height = _get_image_dimensions(data)

    # 4. 尺寸过小（项目符号图标、边框像素等）
    if width is not None and height is not None:
        if width <= MIN_IMAGE_DIMENSION_PX and height <= MIN_IMAGE_DIMENSION_PX:
            return True

        # 5. 尺寸过窄/过扁（分隔线、装饰条）
        if width > 0 and height > 0:
            ratio = max(width / height, height / width)
            if ratio >= MAX_ASPECT_RATIO:
                return True

    return False


def _setup_image_output_dir(markdown_output_path):
    """
    在 Markdown 输出文件旁创建 images/ 子目录。

    Args:
        markdown_output_path: Markdown 输出文件的绝对路径

    Returns:
        (image_save_dir, image_rel_dir) 绝对路径和相对路径
    """
    md_dir = os.path.dirname(markdown_output_path)
    image_root = os.path.join(md_dir, IMAGE_OUTPUT_DIR_NAME)
    os.makedirs(image_root, exist_ok=True)
    stem = os.path.splitext(os.path.basename(markdown_output_path))[0]
    image_save_dir = tempfile.mkdtemp(prefix=stem[:80] + '-', dir=image_root)
    return image_save_dir, os.path.relpath(image_save_dir, md_dir).replace(os.sep, '/')


def _image_warning(diagnostics, location, error, **coordinates):
    diagnostics.append({
        'code': 'IMAGE_EXTRACTION_FAILED', 'severity': 'warning',
        'message': f'{location}: 图片提取失败: {error}', **coordinates,
    })


def _save_extracted_image(data, image_save_dir, image_rel_dir, base_name, image_counter):
    """
    保存图片到 images/ 目录。

    Args:
        data: 图片二进制数据
        image_save_dir: images/ 目录绝对路径
        image_rel_dir: images/ 相对路径（用于 Markdown 引用）
        base_name: 文档基础名称（不含扩展名）
        image_counter: 图片序号

    Returns:
        相对路径字符串，如 'images/report_img_001.png'，失败返回 None
    """
    fmt = _detect_image_format(data)
    if fmt is None:
        # 尝试猜测，默认保存为 png
        fmt = 'png'

    # 对 wmf/emf 保持原格式
    ext = fmt
    filename = f"{base_name}_img_{image_counter:03d}.{ext}"
    abs_path = os.path.join(image_save_dir, filename)

    try:
        with open(abs_path, 'wb') as f:
            f.write(data)
        # 使用正斜杠确保 Markdown 跨平台兼容
        return f"{image_rel_dir}/{filename}"
    except OSError:
        logger.debug("Failed to save extracted image: %s", abs_path, exc_info=True)
        return None


def _make_image_markdown(rel_path, alt_text=None):
    """生成 Markdown 图片语法"""
    alt = alt_text.strip() if alt_text else "image"
    # 转义 alt 文本中的 Markdown 特殊字符
    alt = alt.replace('[', '\\[').replace(']', '\\]')
    return f"![{alt}]({rel_path})"


def _check_ooxml_decorative_flag(element, namespaces=None):
    """
    检查 OOXML 元素是否标记为装饰性（adec:decorative val="1"）。
    同时获取元素的 alt text（descr 属性）。

    Args:
        element: lxml/ET 元素（通常是 docPr 或 cNvPr）
        namespaces: 命名空间字典

    Returns:
        (is_decorative, alt_text) 元组
    """
    if element is None:
        return False, ""

    ns = namespaces or OOXML_IMAGE_NAMESPACES
    is_decorative = False
    alt_text = ""

    # 尝试从 docPr / cNvPr 读取 descr 属性
    alt_text = element.get('descr', '') or ''

    # 检查 adec:decorative 子元素
    adec_ns = ns.get('adec', 'http://schemas.microsoft.com/office/drawing/2017/decorative')
    for child in element:
        tag = child.tag
        # 处理带命名空间和不带命名空间两种情况
        if tag == f'{{{adec_ns}}}decorative' or tag.endswith('}decorative'):
            if child.get('val', '0') == '1':
                is_decorative = True
                break

    return is_decorative, alt_text
