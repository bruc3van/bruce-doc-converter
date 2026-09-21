"""Pdf conversion helpers."""
from .common import (
    _escape_plain_markdown_text,
    _normalize_table_cell,
    _normalize_text,
)
import re


def _render_pdf_table(table_obj):
    """将 pdfplumber 表格对象渲染为 Markdown 表格字符串"""
    table = table_obj.extract()
    if not table or len(table) == 0:
        return ""
    filtered_table = []
    for row in table:
        if row and any(cell is not None and str(cell).strip() for cell in row):
            cleaned_row = [_normalize_table_cell(cell) for cell in row]
            filtered_table.append(cleaned_row)
    if not filtered_table:
        return ""
    max_cols = max(len(r) for r in filtered_table)
    normalized = [r + [""] * (max_cols - len(r)) for r in filtered_table]
    result = ""
    for idx, row in enumerate(normalized):
        result += "| " + " | ".join(row) + " |\n"
        if idx == 0:
            result += "| " + " | ".join(["---"] * len(row)) + " |\n"
    result += "\n"
    return result


def _group_words_into_lines(words, y_tolerance=3):
    """将词元按 top 坐标分组成行列表，每行按 x0 排序"""
    if not words:
        return []
    sorted_words = sorted(words, key=lambda w: (w['top'], w['x0']))
    lines = []
    cur_line = [sorted_words[0]]
    cur_top = sorted_words[0]['top']
    for word in sorted_words[1:]:
        if abs(word['top'] - cur_top) <= y_tolerance:
            cur_line.append(word)
        else:
            lines.append(sorted(cur_line, key=lambda w: w['x0']))
            cur_line = [word]
            cur_top = word['top']
    lines.append(sorted(cur_line, key=lambda w: w['x0']))
    return lines


def _reconstruct_line_text(line_words):
    """从同一行的词元重建文本，词间用空格分隔"""
    if not line_words:
        return ""
    return ' '.join(w['text'] for w in line_words)


def _get_body_font_size(chars):
    """计算正文字体大小（出现频率最高的字体大小，0.5pt 精度）"""
    from collections import Counter
    sizes = [round(c.get('size', 0) * 2) / 2 for c in chars if c.get('size', 0) > 0]
    if not sizes:
        return 10.0
    return float(Counter(sizes).most_common(1)[0][0])


def _get_line_avg_font_size(line_words, page_chars):
    """通过 page.chars 计算某行的平均字体大小"""
    if not line_words or not page_chars:
        return 0.0
    line_top = min(w['top'] for w in line_words)
    line_bottom = max(w['bottom'] for w in line_words)
    margin = max((line_bottom - line_top) * 0.5, 1.0)
    relevant = [
        c for c in page_chars
        if c.get('top', 0) >= line_top - margin and c.get('bottom', 0) <= line_bottom + margin
    ]
    sizes = [c.get('size', 0) for c in relevant if c.get('size', 0) > 0]
    return sum(sizes) / len(sizes) if sizes else 0.0


def _detect_column_split(page_width, words, min_side_words=15):
    """
    检测双栏布局。检查页面中央是否存在明显空白带，是则返回分割线 x 坐标，否则返回 None。
    适用于学术论文的双栏 PDF。
    """
    if not words or len(words) < min_side_words * 2:
        return None
    mid = page_width / 2
    band = page_width * 0.07  # 中央空白带宽度的一半

    center = sum(1 for w in words if w['x0'] > mid - band and w['x1'] < mid + band)
    sides = sum(1 for w in words if w['x1'] <= mid - band or w['x0'] >= mid + band)
    total = center + sides
    if total == 0:
        return None
    # 中央词占比 < 6% 且两侧词数充足，判定为双栏
    if sides >= min_side_words * 2 and center / total < 0.06:
        return mid
    return None


def _split_pdf_words_by_columns(words, page_chars, col_split, gap=5):
    """按双栏分割词元，额外保留跨栏词元，避免标题等内容丢失"""
    left_words, right_words, spanning_words = [], [], []
    left_chars, right_chars, spanning_chars = [], [], []

    left_limit = col_split + gap
    right_limit = col_split - gap

    for word in words:
        if word['x1'] <= left_limit:
            left_words.append(word)
        elif word['x0'] >= right_limit:
            right_words.append(word)
        else:
            spanning_words.append(word)

    for char in page_chars:
        char_x0 = char.get('x0', 0)
        char_x1 = char.get('x1', 0)
        if char_x1 <= left_limit:
            left_chars.append(char)
        elif char_x0 >= right_limit:
            right_chars.append(char)
        else:
            spanning_chars.append(char)

    return left_words, right_words, spanning_words, left_chars, right_chars, spanning_chars


def _split_markdown_blocks(content):
    return [block.strip() for block in re.split(r"\n\s*\n", content or "") if block.strip()]


def _parse_pdf_academic_section_block(block):
    """识别论文常见章节标题，支持标题独占或同段内联写法"""
    raw = (block or '').strip()
    plain = re.sub(r"^#{1,6}\s*", "", raw).strip()

    appendix_match = re.match(r"^(appendix|appendices|附录)\s*([A-Za-z0-9一二三四五六七八九十]*)\s*[:：\-]?\s*(.*)$", plain, re.IGNORECASE)
    if appendix_match:
        _, suffix, inline_body = appendix_match.groups()
        heading = "Appendix"
        if suffix:
            heading = f"{heading} {suffix.strip()}"
        return {"section": "appendix", "heading": heading, "inline_body": inline_body.strip()}

    patterns = [
        ("abstract", "Abstract", [r"abstract", r"摘要"]),
        ("keywords", "Keywords", [r"keywords?", r"index terms?", r"关键词"]),
        ("references", "References", [r"references?", r"bibliography", r"参考文献"]),
    ]

    for section_key, heading, aliases in patterns:
        alias_pattern = "|".join(f"(?:{alias})" for alias in aliases)
        match = re.match(rf"^(?:{alias_pattern})\s*[:：]?\s*(.*)$", plain, re.IGNORECASE)
        if not match:
            continue
        inline_body = match.group(1).strip()
        if inline_body and re.match(r"^[A-Za-z]+$", inline_body) and inline_body.lower() == plain.lower():
            inline_body = ""
        return {"section": section_key, "heading": heading, "inline_body": inline_body}

    return None


def _is_markdown_heading_block(block):
    first_line = (block or '').strip().splitlines()[0] if (block or '').strip() else ''
    return bool(re.match(r"^#{1,6}\s+\S", first_line))


def _format_pdf_keywords_block(body_blocks):
    text = " ".join(_normalize_text(block) for block in body_blocks if block.strip())
    text = re.sub(r"^(?:keywords?|index terms?|关键词)\s*[:：]\s*", "", text, flags=re.IGNORECASE)
    if not text:
        return "## Keywords"

    parts = [part.strip() for part in re.split(r"[;,；，、]\s*", text) if part.strip()]
    if not parts:
        return "## Keywords"
    return "## Keywords\n\n" + "\n".join(f"- {part}" for part in parts)


def _format_pdf_references_block(body_blocks):
    items = []
    for block in body_blocks:
        cleaned = _normalize_text(block, preserve_newlines=True)
        for line in cleaned.splitlines():
            text = re.sub(r"^(?:\[\d+\]|\d+[.)])\s*", "", line.strip())
            if text:
                items.append(text)

    if not items:
        return "## References"
    return "## References\n\n" + "\n".join(f"1. {item}" for item in items)


def _format_pdf_academic_section(section_key, heading, body_blocks):
    if section_key == "keywords":
        return _format_pdf_keywords_block(body_blocks)
    if section_key == "references":
        return _format_pdf_references_block(body_blocks)

    body = "\n\n".join(block.strip() for block in body_blocks if block.strip()).strip()
    rendered_heading = f"## {heading}"
    return f"{rendered_heading}\n\n{body}".strip()


def _postprocess_pdf_academic_sections(content):
    """将论文常见章节归一成固定 Markdown 结构"""
    blocks = _split_markdown_blocks(content)
    if not blocks:
        return content

    result_blocks = []
    current_section = None
    current_heading = None
    current_body = []

    def _flush_current():
        nonlocal current_section, current_heading, current_body
        if current_section is None:
            return
        result_blocks.append(_format_pdf_academic_section(current_section, current_heading, current_body))
        current_section = None
        current_heading = None
        current_body = []

    for block in blocks:
        parsed = _parse_pdf_academic_section_block(block)
        if parsed:
            _flush_current()
            current_section = parsed["section"]
            current_heading = parsed["heading"]
            current_body = [parsed["inline_body"]] if parsed["inline_body"] else []
            continue

        if current_section is not None and (_is_markdown_heading_block(block) or re.match(r"^##\s+Page\s+\d+\s*$", block)):
            _flush_current()
            result_blocks.append(block)
            continue

        if current_section is not None:
            current_body.append(block)
            continue

        result_blocks.append(block)

    _flush_current()
    return "\n\n".join(block for block in result_blocks if block.strip()).strip()


def _lines_to_markdown_blocks(lines, page_chars, body_size):
    """
    将行列表（每行为词元列表）转换为 (top, markdown_text) 块列表。
    - 字体明显大于正文的行识别为标题，连续标题行合并为一个标题
    - 连续普通文本行合并成段落，行间距过大时另起段落
    """
    blocks = []
    para_lines = []
    para_top = None
    prev_bottom = None
    prev_line_height = None
    heading_lines = []   # 连续标题行暂存
    heading_top = None

    def _flush_para():
        if not para_lines:
            return
        text = _normalize_text(' '.join(para_lines))
        if text:
            blocks.append((para_top, _escape_plain_markdown_text(text) + "\n\n"))

    def _flush_heading():
        if not heading_lines:
            return
        text = ' '.join(heading_lines)
        blocks.append((heading_top, f"### {text}\n\n"))

    for line_words in lines:
        if not line_words:
            continue
        line_top = min(w['top'] for w in line_words)
        line_bottom = max(w['bottom'] for w in line_words)
        line_height = max(line_bottom - line_top, 1.0)
        line_text = _reconstruct_line_text(line_words).strip()
        if not line_text:
            continue

        # 字体大小检测：比正文大 15% 以上视为标题
        line_size = _get_line_avg_font_size(line_words, page_chars)
        is_heading = line_size > 0 and body_size > 0 and line_size >= body_size * 1.15

        # 行间距检测：使用较小行高作为参考，避免大字号标题的阈值过大
        if prev_bottom is not None:
            gap = line_top - prev_bottom
            ref_height = min(line_height, prev_line_height) if prev_line_height else line_height
            large_gap = gap > ref_height * 0.8
        else:
            large_gap = False

        if is_heading:
            _flush_para()
            para_lines = []
            para_top = None
            # 连续标题行合并：若与上一标题行无大间距则追加
            if heading_lines and not large_gap:
                heading_lines.append(line_text)
            else:
                _flush_heading()
                heading_lines = [line_text]
                heading_top = line_top
        else:
            _flush_heading()
            heading_lines = []
            heading_top = None
            if large_gap:
                _flush_para()
                para_lines = []
                para_top = None
            if para_top is None:
                para_top = line_top
            para_lines.append(line_text)

        prev_bottom = line_bottom
        prev_line_height = line_height

    _flush_heading()
    _flush_para()
    return blocks


def _extract_pdf_page_blocks(page, tables):
    """
    提取单页 PDF 的文本和表格块。
    改进：
    - 使用 extract_words() 重建文本，修复 LaTeX PDF 词间空格丢失问题
    - 基于字体大小识别标题行
    - 检测双栏布局（学术论文常见），分栏提取后顺序拼接
    """
    table_bboxes = [t.bbox for t in tables] if tables else []
    blocks = []

    # 收集表格块
    for table_obj in tables:
        md = _render_pdf_table(table_obj)
        if md:
            blocks.append((table_obj.bbox[1], md))

    # 过滤掉表格区域内的对象
    filtered_page = page
    for bbox in table_bboxes:
        filtered_page = filtered_page.filter(
            lambda obj, b=bbox: not (
                obj.get("top", 0) >= b[1] and
                obj.get("bottom", 0) <= b[3] and
                obj.get("x0", 0) >= b[0] and
                obj.get("x1", 0) <= b[2]
            )
        )

    # 使用词元提取（修复空格丢失）
    # x_tolerance=2: 学术 PDF（LaTeX）词间距约 2.7pt，默认值 3 会把相邻词粘连
    try:
        words = filtered_page.extract_words(x_tolerance=2, y_tolerance=3, keep_blank_chars=False)
    except TypeError:
        # 旧版 pdfplumber 不支持 keep_blank_chars 参数
        words = filtered_page.extract_words(x_tolerance=2, y_tolerance=3)

    # 过滤旋转文字（如 arXiv 水印），保留 upright（正向）词元
    words = [w for w in words if w.get('upright', 1)]

    page_chars = filtered_page.chars

    if not words:
        # 回退到 extract_text()
        text = _normalize_text(filtered_page.extract_text(), preserve_newlines=True)
        if text:
            blocks.append((0.0, "\n".join(_escape_plain_markdown_text(line) for line in text.splitlines()) + "\n\n"))
        blocks.sort(key=lambda b: b[0])
        return blocks

    body_size = _get_body_font_size(page_chars) if page_chars else 10.0

    # 检测双栏布局
    col_split = _detect_column_split(page.width, words)

    if col_split is not None:
        # 双栏：分别处理左右两栏，并单独保留跨栏标题/摘要等元素
        left_words, right_words, spanning_words, left_chars, right_chars, spanning_chars = _split_pdf_words_by_columns(
            words,
            page_chars,
            col_split,
        )

        left_lines = _group_words_into_lines(left_words)
        right_lines = _group_words_into_lines(right_words)
        spanning_lines = _group_words_into_lines(spanning_words)

        left_blocks = _lines_to_markdown_blocks(left_lines, left_chars, body_size)
        right_blocks = _lines_to_markdown_blocks(right_lines, right_chars, body_size)
        spanning_blocks = _lines_to_markdown_blocks(spanning_lines, spanning_chars, body_size)

        left_max = max((top for top, _ in left_blocks), default=0.0)
        offset = left_max + page.height
        right_shifted = [(top + offset, content) for top, content in right_blocks]

        blocks.extend(spanning_blocks)
        blocks.extend(left_blocks)
        blocks.extend(right_shifted)
    else:
        lines = _group_words_into_lines(words)
        text_blocks = _lines_to_markdown_blocks(lines, page_chars, body_size)
        blocks.extend(text_blocks)

    blocks.sort(key=lambda b: b[0])
    return blocks


def convert_pdf(file_path, diagnostics=None):
    """转换 PDF 文件，支持文本和表格提取，按页面位置交错输出"""
    import pdfplumber

    content_parts = []
    page_errors = []
    diagnostics = diagnostics if diagnostics is not None else []
    with pdfplumber.open(file_path) as pdf:
        for page_number, page in enumerate(pdf.pages, 1):
            status, detail = 'extracted', ''
            try:
                tables = page.find_tables()
                blocks = _extract_pdf_page_blocks(page, tables)
            except Exception as exc:
                status, detail = 'fallback', str(exc)
                page_errors.append((page_number, str(exc)))
                try:
                    fallback_text = _normalize_text(page.extract_text(), preserve_newlines=True)
                except Exception:
                    fallback_text = ""
                if not fallback_text:
                    diagnostics.append({'code': 'PDF_PAGE_FAILED', 'severity': 'warning',
                                        'page': page_number, 'status': 'failed',
                                        'message': f'PDF 第 {page_number} 页提取失败: {detail}'})
                    continue
                blocks = [(0.0, "\n".join(_escape_plain_markdown_text(line) for line in fallback_text.splitlines()) + "\n\n")]

            if not blocks:
                diagnostics.append({'code': 'PDF_PAGE_EMPTY', 'severity': 'warning',
                                    'page': page_number, 'status': 'empty',
                                    'message': f'PDF 第 {page_number} 页无可提取内容，可能为空白页或需要 OCR。'})
                continue

            page_content = "".join(block_content for _, block_content in blocks).strip()
            if not page_content:
                diagnostics.append({'code': 'PDF_PAGE_EMPTY', 'severity': 'warning',
                                    'page': page_number, 'status': 'empty',
                                    'message': f'PDF 第 {page_number} 页无可提取内容。'})
                continue

            diagnostics.append({'code': 'PDF_PAGE_FALLBACK' if status == 'fallback' else 'PDF_PAGE_EXTRACTED',
                                'severity': 'warning' if status == 'fallback' else 'info',
                                'page': page_number, 'status': status,
                                'message': f'PDF 第 {page_number} 页已降级为纯文本提取: {detail}' if detail else f'PDF 第 {page_number} 页提取成功。'})

            if len(pdf.pages) > 1:
                content_parts.append(f"## Page {page_number}\n\n{page_content}")
            else:
                content_parts.append(page_content)

    content = "\n\n".join(content_parts).strip()
    if not content and page_errors:
        page_numbers = ", ".join(str(page_number) for page_number, _ in page_errors)
        raise ValueError(f"PDF 解析失败，无法提取任何内容。异常页码: {page_numbers}")
    return _postprocess_pdf_academic_sections(content)
