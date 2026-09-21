"""Common conversion helpers."""
import re


DOCX_XML_NAMESPACES = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}


DOCX_W_NS = DOCX_XML_NAMESPACES['w']


IMAGE_OUTPUT_DIR_NAME = "images"


MIN_IMAGE_DIMENSION_PX = 20          # 小于此像素的图片视为装饰性


MAX_ASPECT_RATIO = 10.0              # 宽高比超过此值视为装饰线条


MIN_IMAGE_DATA_BYTES = 500           # 数据量低于此值视为纯色/透明占位


PPTX_BACKGROUND_COVERAGE_RATIO = 0.9 # 覆盖幻灯片面积超过此比例视为背景图


OOXML_IMAGE_NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'adec': 'http://schemas.microsoft.com/office/drawing/2017/decorative',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
}


_IMAGE_SIGNATURES = {
    b'\x89PNG\r\n\x1a\n': 'png',
    b'\xff\xd8\xff': 'jpeg',
    b'GIF87a': 'gif',
    b'GIF89a': 'gif',
    b'BM': 'bmp',
    b'II\x2a\x00': 'tiff',
    b'MM\x00\x2a': 'tiff',
    b'\xd7\xcd\xc6\x9a': 'wmf',
    b'\x01\x00\x00\x00': 'emf',
}


_RE_COLLAPSE_WHITESPACE = re.compile(r"\s+")


_RE_COLLAPSE_EXTRA_BLANK_LINES = re.compile(r"\n{3,}")


_RE_ESCAPE_MARKDOWN_LEADING = re.compile(r"^([>#\-\+\*])")


_RE_ESCAPE_MARKDOWN_ORDERED_LIST = re.compile(r"^(\d+)\.\s")


_RE_WRAP_INLINE_MARKDOWN = re.compile(r"^(\s*)(.*?)(\s*)$", re.DOTALL)


def _normalize_text(value, preserve_newlines=False):
    """规范化提取出来的文本，减少空白和换行噪声"""
    if value is None:
        return ""

    text = str(value).replace("\r\n", "\n").replace("\r", "\n")
    if preserve_newlines:
        lines = [_RE_COLLAPSE_WHITESPACE.sub(" ", line).strip() for line in text.split("\n")]
        text = "\n".join(line for line in lines if line)
        return _RE_COLLAPSE_EXTRA_BLANK_LINES.sub("\n\n", text).strip()

    text = _RE_COLLAPSE_WHITESPACE.sub(" ", text.replace("\n", " "))
    return text.strip()


def _escape_plain_markdown_text(text):
    """避免普通文本误触发 Markdown 标题、列表、引用等语法"""
    if not text:
        return ""

    escaped = text
    escaped = _RE_ESCAPE_MARKDOWN_LEADING.sub(r"\\\1", escaped)
    escaped = _RE_ESCAPE_MARKDOWN_ORDERED_LIST.sub(r"\\\1. ", escaped)
    return escaped


def _format_inline_markdown(text, *, bold=False, italic=False):
    """保留两侧空白后再包裹 Markdown 强调标记，避免单词粘连"""
    if not text:
        return ""

    if not (bold or italic):
        return text

    match = _RE_WRAP_INLINE_MARKDOWN.match(text)
    if not match:
        return text

    leading, core, trailing = match.groups()
    if not core:
        return text

    if bold and italic:
        wrapped = f"***{core}***"
    elif bold:
        wrapped = f"**{core}**"
    else:
        wrapped = f"*{core}*"
    return f"{leading}{wrapped}{trailing}"


def _compose_inline_markdown(groups):
    """将分组后的富文本片段转换为 Markdown，并避免误转义已生成的强调语法"""
    formatted_parts = []
    for index, ((bold, italic), text) in enumerate(groups):
        current_text = text
        if index == 0 and not (bold or italic):
            current_text = _escape_plain_markdown_text(current_text)
        formatted_parts.append(_format_inline_markdown(current_text, bold=bold, italic=italic))
    return "".join(formatted_parts)


def _normalize_table_cell(value):
    """将单元格内容规范化为安全的 Markdown 表格单元格文本"""
    text = _normalize_text(value)
    # Markdown 表格分隔符转义
    text = text.replace("|", "\\|")
    return text


def _table_position_has_content(value, occupied=False):
    """判断表格位置是否应保留，用于保留合并单元格占位"""
    return occupied or (value is not None and str(value).strip() != "")
