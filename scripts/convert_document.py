#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
独立的文档转换脚本
- 将 Word、Excel、PowerPoint 和 PDF 文件转换为 Markdown 格式
- 将 Markdown 文件转换为 Word (.docx) 格式

依赖：
- python-docx: pip install python-docx
- openpyxl: pip install openpyxl
- python-pptx: pip install python-pptx
- pdfplumber: pip install pdfplumber
- Node.js: 用于 Markdown 转 DOCX（需要单独安装）
"""

import sys
import os
import json
import importlib
import re
import subprocess
import shutil
import io

SUPPORTED_EXTENSIONS = ['.docx', '.xlsx', '.pptx', '.pdf', '.md']
MAX_FILE_SIZE_BYTES = 100 * 1024 * 1024
NODE_CONVERT_TIMEOUT_SECONDS = 120
NODE_SHARED_HOME_ENV = "DOCUGENIUS_NODE_HOME"
GENERATED_OUTPUT_DIR_NAMES = {"Markdown", "Word"}

def _configure_windows_stdio():
    """
    Windows 控制台经常使用 GBK/CP936，直接输出某些符号（如 ✓/✗）会触发 UnicodeEncodeError。

    策略：
    - 交互式控制台（isatty=True）：不强制切换编码，尽量保持用户终端显示正常；仅将 errors 调整为 replace，避免崩溃。
    - 非交互（被管道/测试框架捕获）：优先输出 UTF-8，保证机器可读；同样使用 errors=replace 兜底。
    """
    if sys.platform != "win32":
        return

    def _safe_reconfigure(stream, *, encoding=None, errors=None):
        if not hasattr(stream, "reconfigure"):
            return False
        try:
            kwargs = {}
            if encoding is not None:
                kwargs["encoding"] = encoding
            if errors is not None:
                kwargs["errors"] = errors
            stream.reconfigure(**kwargs)
            return True
        except Exception:
            return False

    def _safe_wrap(stream, *, encoding, errors):
        buffer = None
        if hasattr(stream, "detach"):
            try:
                buffer = stream.detach()
            except Exception:
                buffer = None
        if buffer is None:
            buffer = getattr(stream, "buffer", None)
        if buffer is None:
            return False
        try:
            wrapped = io.TextIOWrapper(buffer, encoding=encoding, errors=errors, line_buffering=True)
            if stream is sys.stdout:
                sys.stdout = wrapped
            elif stream is sys.stderr:
                sys.stderr = wrapped
            return True
        except Exception:
            return False

    is_tty = bool(getattr(sys.stdout, "isatty", lambda: False)())
    errors = "replace"

    if is_tty:
        if not _safe_reconfigure(sys.stdout, errors=errors):
            current_encoding = getattr(sys.stdout, "encoding", None) or "utf-8"
            _safe_wrap(sys.stdout, encoding=current_encoding, errors=errors)
        if not _safe_reconfigure(sys.stderr, errors=errors):
            current_encoding = getattr(sys.stderr, "encoding", None) or "utf-8"
            _safe_wrap(sys.stderr, encoding=current_encoding, errors=errors)
        return

    target_encoding = "utf-8"
    if not _safe_reconfigure(sys.stdout, encoding=target_encoding, errors=errors):
        _safe_wrap(sys.stdout, encoding=target_encoding, errors=errors)
    if not _safe_reconfigure(sys.stderr, encoding=target_encoding, errors=errors):
        _safe_wrap(sys.stderr, encoding=target_encoding, errors=errors)


_configure_windows_stdio()

# ==================== 依赖配置 ====================

_DEPENDENCIES_BY_EXT = {
    '.docx': [('docx', 'python-docx')],
    '.xlsx': [('openpyxl', 'openpyxl')],
    '.pptx': [('pptx', 'python-pptx')],
    '.pdf': [('pdfplumber', 'pdfplumber')],
    '.md': [],  # Markdown 转 DOCX 使用 Node.js，无 Python 依赖
}

# ==================== Node.js 共享依赖目录 ====================

def _get_node_shared_root():
    override = os.environ.get(NODE_SHARED_HOME_ENV)
    if override:
        return override

    if sys.platform == "win32":
        base = os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA")
        if not base:
            base = os.path.join(os.path.expanduser("~"), "AppData", "Local")
        return os.path.join(base, "DocuGenius", "node")

    return os.path.join(os.path.expanduser("~"), ".docugenius", "node")

def _sync_shared_package_files(source_dir, target_dir):
    for filename in ("package.json", "package-lock.json"):
        src = os.path.join(source_dir, filename)
        if not os.path.exists(src):
            continue
        dst = os.path.join(target_dir, filename)
        try:
            shutil.copy2(src, dst)
        except Exception as e:
            return False, f"无法复制 {filename} 到共享目录: {str(e)}"
    return True, None

def _ensure_shared_node_modules(shared_dir, source_dir):
    npm_cmd = shutil.which('npm')
    if not npm_cmd:
        return False, "未找到 npm。请安装 Node.js（自带 npm）后重试。"

    try:
        os.makedirs(shared_dir, exist_ok=True)
    except Exception as e:
        return False, f"无法创建共享依赖目录: {str(e)}"

    ok, err = _sync_shared_package_files(source_dir, shared_dir)
    if not ok:
        return False, err

    cmd = [npm_cmd, "install", "--no-fund", "--no-audit"]
    try:
        print("[DocuGenius] 正在安装 Node.js 依赖到用户共享目录...", file=sys.stderr)
        result = subprocess.run(
            cmd,
            cwd=shared_dir,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        if result.returncode != 0:
            error_output = result.stderr.strip() or result.stdout.strip()
            return False, f"Node.js 依赖安装失败: {error_output}"
    except Exception as e:
        return False, f"Node.js 依赖安装失败: {str(e)}"

    return True, None

def _find_mmdc_binary(node_modules_dir):
    if not node_modules_dir:
        return None

    ext = ".cmd" if sys.platform == "win32" else ""
    candidate = os.path.join(node_modules_dir, ".bin", f"mmdc{ext}")
    if os.path.exists(candidate):
        return candidate
    return None

# ==================== 依赖安装函数 ====================

def install_dependencies(pip_packages):
    """
    自动安装缺失的依赖到用户目录（使用 --user 标志）

    使用 --user 安装的好处：
    - 不受 PEP 668 系统保护限制（适用于 macOS 从系统 Python 安装）
    - 无需虚拟环境
    - 安装到 ~/.local/ (Linux/macOS) 或 %APPDATA% (Windows)
    - 所有项目共享

    Args:
        pip_packages: 要安装的包名列表

    Returns:
        (success: bool, error_message: str or None)
    """
    if not pip_packages:
        return True, None

    # 获取当前 Python 可执行文件路径
    python_exe = sys.executable or 'python'

    # 构建安装命令
    # 首先尝试 --user 安装到用户目录，避免 PEP 668 限制
    # 如果失败，回退到 --break-system-packages（适用于使用系统 Python 的情况）
    install_methods = [
        ('--user', '用户目录'),
        ('--break-system-packages', '系统目录（绕过 PEP 668 保护）'),
    ]

    for install_flag, location_desc in install_methods:
        cmd = [python_exe, '-m', 'pip', 'install', install_flag] + pip_packages

        try:
            # 显示安装提示（仅在首次安装时）
            print(f"[DocuGenius] 正在安装缺失的依赖: {', '.join(pip_packages)}", file=sys.stderr)
            print(f"[DocuGenius] 安装位置: {location_desc}", file=sys.stderr)

            # 运行安装命令
            result = subprocess.run(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding="utf-8",
                errors="replace",
            )

            if result.returncode == 0:
                print(f"[DocuGenius] 依赖安装成功！", file=sys.stderr)
                return True, None

            # 检查错误类型，决定是否尝试下一种方法
            error_output = result.stderr.strip() or result.stdout.strip()

            # 如果是 PEP 668 相关错误，尝试下一种方法
            if 'externally-managed-environment' in error_output or 'PEP 668' in error_output:
                continue

            # 其他错误不再重试
            # 检查是否是权限问题
            if "Permission denied" in error_output or "Access denied" in error_output:
                return False, f"权限不足。请尝试: pip install {install_flag} {' '.join(pip_packages)}"

            # 检查是否是 pip 不存在
            if "No module named pip" in error_output:
                return False, "pip 未安装。请先安装 pip: python -m ensurepip --upgrade"

            # 其他错误
            return False, f"依赖安装失败: {error_output}"

        except FileNotFoundError:
            return False, f"找不到 Python 解释器: {python_exe}"
        except Exception as e:
            return False, f"依赖安装时发生错误: {str(e)}"

    # 所有方法都失败
    return False, "依赖安装失败：所有安装方法都失败（包括 --user 和 --break-system-packages）"


# ==================== 依赖检查函数 ====================

def check_dependencies(file_ext=None, auto_install=True):
    """
    检查必需的依赖是否已安装（默认检查全部；传入 file_ext 时仅检查该格式所需）

    Args:
        file_ext: 文件扩展名（如 '.docx'），仅检查该格式所需的依赖
        auto_install: 是否自动安装缺失的依赖（默认 True）

    Returns:
        (success: bool, error_message: str or None)
    """
    # 确定需要检查的依赖
    deps = _DEPENDENCIES_BY_EXT.get(file_ext) if file_ext else [
        ('docx', 'python-docx'),
        ('openpyxl', 'openpyxl'),
        ('pptx', 'python-pptx'),
        ('pdfplumber', 'pdfplumber'),
    ]
    if deps is None:
        deps = [
            ('docx', 'python-docx'),
            ('openpyxl', 'openpyxl'),
            ('pptx', 'python-pptx'),
            ('pdfplumber', 'pdfplumber'),
        ]

    # 检查依赖是否已安装
    missing = []
    missing_pip_names = []
    for module_name, pip_name in deps:
        try:
            importlib.import_module(module_name)
        except ImportError:
            missing.append(module_name)
            missing_pip_names.append(pip_name)

    # 如果有缺失依赖且启用了自动安装
    if missing and auto_install:
        success, error = install_dependencies(missing_pip_names)
        if success:
            # 安装成功，重新检查
            still_missing = []
            for module_name, pip_name in deps:
                try:
                    importlib.import_module(module_name)
                except ImportError:
                    still_missing.append(pip_name)

            if still_missing:
                return False, f"依赖安装后仍无法加载: {', '.join(still_missing)}。请手动安装并检查 Python 环境。"

            return True, None
        else:
            # 安装失败
            return False, error

    # 有缺失但未启用自动安装
    if missing:
        return False, f"缺少依赖库: {', '.join(missing_pip_names)}。请运行: pip install --user {' '.join(missing_pip_names)}"

    return True, None

def _normalize_text(value, preserve_newlines=False):
    """规范化提取出来的文本，减少空白和换行噪声"""
    if value is None:
        return ""

    text = str(value).replace("\r\n", "\n").replace("\r", "\n")
    if preserve_newlines:
        lines = [re.sub(r"\s+", " ", line).strip() for line in text.split("\n")]
        text = "\n".join(line for line in lines if line)
        return re.sub(r"\n{3,}", "\n\n", text).strip()

    text = re.sub(r"\s+", " ", text.replace("\n", " "))
    return text.strip()

def _escape_plain_markdown_text(text):
    """避免普通文本误触发 Markdown 标题、列表、引用等语法"""
    if not text:
        return ""

    escaped = text
    escaped = re.sub(r"^([>#\-\+\*])", r"\\\1", escaped)
    escaped = re.sub(r"^(\d+)\.\s", r"\\\1. ", escaped)
    return escaped

def _format_inline_markdown(text, *, bold=False, italic=False):
    """保留两侧空白后再包裹 Markdown 强调标记，避免单词粘连"""
    if not text:
        return ""

    if not (bold or italic):
        return text

    match = re.match(r"^(\s*)(.*?)(\s*)$", text, re.DOTALL)
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

def _validate_input_file(file_path):
    """校验并规范化输入文件路径"""
    if file_path is None:
        return None, "文件路径不能为空"

    normalized = os.path.abspath(os.path.normpath(os.path.expanduser(str(file_path))))
    if not os.path.exists(normalized):
        return None, f'文件不存在: {normalized}'
    if not os.path.isfile(normalized):
        return None, f'输入路径不是文件: {normalized}'
    return normalized, None

def _resolve_markdown_output_path(file_path, output_dir=None):
    """生成 Markdown 输出路径，并确保输出目录可用"""
    if output_dir:
        target_dir = os.path.abspath(os.path.normpath(os.path.expanduser(str(output_dir))))
    else:
        file_dir = os.path.dirname(file_path) or '.'
        target_dir = os.path.join(file_dir, 'Markdown')

    if os.path.exists(target_dir) and not os.path.isdir(target_dir):
        raise NotADirectoryError(f'输出路径不是目录: {target_dir}')

    os.makedirs(target_dir, exist_ok=True)
    output_filename = os.path.splitext(os.path.basename(file_path))[0] + '.md'
    return os.path.join(target_dir, output_filename)

def _iter_batch_input_files(directory, recursive=True, output_dir=None):
    """遍历批量转换输入文件，跳过已生成输出目录，避免重复处理"""
    normalized_directory = os.path.abspath(os.path.normpath(os.path.expanduser(str(directory))))
    custom_output_dir = None
    if output_dir:
        custom_output_dir = os.path.abspath(os.path.normpath(os.path.expanduser(str(output_dir))))

    def _should_skip_dir(dir_path):
        if custom_output_dir and os.path.normcase(dir_path) == os.path.normcase(custom_output_dir):
            return True
        return os.path.basename(dir_path) in GENERATED_OUTPUT_DIR_NAMES

    if recursive:
        for root, dirs, files in os.walk(normalized_directory):
            dirs[:] = [dir_name for dir_name in dirs if not _should_skip_dir(os.path.join(root, dir_name))]
            for file in files:
                if os.path.splitext(file)[1].lower() in SUPPORTED_EXTENSIONS:
                    yield os.path.join(root, file)
        return

    for file in os.listdir(normalized_directory):
        file_path = os.path.join(normalized_directory, file)
        if os.path.isfile(file_path) and os.path.splitext(file)[1].lower() in SUPPORTED_EXTENSIONS:
            yield file_path

def _normalize_table_cell(value):
    """将单元格内容规范化为安全的 Markdown 表格单元格文本"""
    text = _normalize_text(value)
    # Markdown 表格分隔符转义
    text = text.replace("|", "\\|")
    return text

def convert_docx(file_path):
    """转换 Word 文档，支持标题、格式和列表（含编号/层级）"""
    import docx

    doc = docx.Document(file_path)
    content = ""

    def get_numbering_info(para):
        """
        尝试从段落的 numPr / numbering.xml 解析列表信息

        Returns:
            None 或 {'level': int, 'ordered': bool}
        """
        try:
            p = getattr(para, "_p", None)
            ppr = getattr(p, "pPr", None) if p is not None else None
            num_pr = getattr(ppr, "numPr", None) if ppr is not None else None
            if num_pr is None:
                return None

            num_id_el = getattr(num_pr, "numId", None)
            ilvl_el = getattr(num_pr, "ilvl", None)
            level = int(ilvl_el.val) if ilvl_el is not None and getattr(ilvl_el, "val", None) is not None else 0

            num_fmt = None
            num_id = None
            if num_id_el is not None and getattr(num_id_el, "val", None) is not None:
                try:
                    num_id = int(num_id_el.val)
                except Exception:
                    num_id = None

            if num_id is not None:
                try:
                    numbering = doc.part.numbering_part.element
                    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
                    num_nodes = numbering.xpath(f'.//w:num[@w:numId="{num_id}"]', namespaces=ns)
                    if num_nodes:
                        abstract_id_nodes = num_nodes[0].xpath('./w:abstractNumId', namespaces=ns)
                        if abstract_id_nodes:
                            abstract_id = abstract_id_nodes[0].get(f'{{{ns["w"]}}}val')
                            abstract_nodes = numbering.xpath(
                                f'.//w:abstractNum[@w:abstractNumId="{abstract_id}"]',
                                namespaces=ns,
                            )
                            if abstract_nodes:
                                lvl_nodes = abstract_nodes[0].xpath(f'./w:lvl[@w:ilvl="{level}"]', namespaces=ns)
                                if not lvl_nodes:
                                    lvl_nodes = abstract_nodes[0].xpath('./w:lvl', namespaces=ns)
                                if lvl_nodes:
                                    fmt_nodes = lvl_nodes[0].xpath('./w:numFmt', namespaces=ns)
                                    if fmt_nodes:
                                        num_fmt = fmt_nodes[0].get(f'{{{ns["w"]}}}val')
                except Exception:
                    num_fmt = None

            style = getattr(para, "style", None)
            style_name = getattr(style, "name", "") if style is not None else ""
            style_id = getattr(style, "style_id", "") if style is not None else ""
            style_hint = f"{style_name} {style_id}".lower()

            if (num_fmt or "").lower() == "bullet":
                ordered = False
            elif (num_fmt or ""):
                ordered = True
            elif "bullet" in style_hint or "项目符号" in style_hint or "符号" in style_hint:
                ordered = False
            elif "number" in style_hint or "编号" in style_hint:
                ordered = True
            else:
                ordered = True

            return {'level': max(level, 0), 'ordered': ordered}
        except Exception:
            return None

    def process_paragraph(para):
        """处理单个段落，识别标题、列表和格式"""
        if not para.text.strip():
            return ""

        style = para.style if hasattr(para, "style") else None
        style_name = style.name if style else ""
        style_id = getattr(style, "style_id", "") if style else ""

        # 先拼接富文本（列表项也需要保留粗体/斜体）
        # 将相邻同格式的 run 合并后再添加 Markdown 标记，避免 **text1****text2** 碎片
        groups = []
        for run in para.runs:
            text = run.text
            if not text:
                continue
            fmt = (bool(run.bold), bool(run.italic))
            if groups and groups[-1][0] == fmt:
                groups[-1] = (fmt, groups[-1][1] + text)
            else:
                groups.append((fmt, text))
        formatted_text = ""
        for (bold, italic), text in groups:
            formatted_text += _format_inline_markdown(text, bold=bold, italic=italic)
        text_value = _normalize_text(formatted_text.strip() or para.text.strip())
        if not text_value:
            return ""

        # 识别标题层级
        heading_level = None
        if isinstance(style_id, str) and style_id:
            m = re.match(r"(?i)^heading(\d+)$", style_id.strip())
            if m:
                heading_level = int(m.group(1))
        if heading_level is None and isinstance(style_name, str) and style_name:
            # 兼容中文 Word 默认标题样式（例如“标题 1”）
            m = re.match(r"^(?:Heading|标题)\s*(\d+)$", style_name.strip())
            if m:
                heading_level = int(m.group(1))

        if heading_level is not None:
            heading_prefix = "#" * min(heading_level, 6)  # Markdown最多支持6级标题
            return f"{heading_prefix} {text_value}\n\n"

        # 检查是否是列表项
        # 优先使用 numPr + numbering.xml 解析列表编号格式与层级
        numbering_info = get_numbering_info(para)
        if numbering_info:
            indent = "    " * numbering_info["level"]
            marker = "1." if numbering_info["ordered"] else "-"
            return f"{indent}{marker} {text_value}\n"

        # 注意：python-docx对列表的支持有限，这里做基本处理（按样式兜底）
        style_name_str = style_name or ""
        style_id_str = style_id or ""
        if (isinstance(style_id_str, str) and style_id_str.startswith("List")) or (
            isinstance(style_name_str, str) and style_name_str.startswith("List")
        ):
            level = 0
            m = re.search(r"(\d+)$", style_name_str.strip())
            if m:
                level = max(int(m.group(1)) - 1, 0)
            indent = "    " * level
            if "Bullet" in style_id_str or "Bullet" in style_name_str or style_name_str == "List Bullet":
                return f"{indent}- {text_value}\n"
            if "Number" in style_id_str or "Number" in style_name_str or style_name_str == "List Number":
                return f"{indent}1. {text_value}\n"

        return _escape_plain_markdown_text(text_value) + "\n\n"

    # 处理文档中的所有元素（段落和表格）
    # 需要按照它们在文档中的顺序处理
    paragraphs_iter = iter(doc.paragraphs)
    tables_iter = iter(doc.tables)
    for element in doc.element.body:
        # 处理段落
        if element.tag.endswith('p'):
            para = next(paragraphs_iter, None)
            if para is not None:
                content += process_paragraph(para)

        # 处理表格
        elif element.tag.endswith('tbl'):
            table = next(tables_iter, None)
            if table is not None:
                # 两遍处理：先收集去重后的行数据，再统一列数输出
                all_rows_data = []
                for row in table.rows:
                    # 去重合并单元格：python-docx 对合并单元格会返回重复的 _tc 引用
                    seen_tcs = set()
                    unique_cells = []
                    for cell in row.cells:
                        tc_id = id(cell._tc)
                        if tc_id not in seen_tcs:
                            seen_tcs.add(tc_id)
                            unique_cells.append(cell)
                    all_rows_data.append([_normalize_table_cell(cell.text) for cell in unique_cells])

                max_cols = max((len(r) for r in all_rows_data), default=0)
                if max_cols == 0:
                    continue
                for i, row_data in enumerate(all_rows_data):
                    padded = row_data + [""] * (max_cols - len(row_data))
                    content += "| " + " | ".join(padded) + " |\n"
                    if i == 0:
                        content += "| " + " | ".join(["---"] * max_cols) + " |\n"
                content += "\n"

    return content.strip()

def convert_xlsx(file_path):
    """转换 Excel 文件"""
    import openpyxl

    workbook = openpyxl.load_workbook(file_path, data_only=True)
    content = ""

    def _build_merge_map(worksheet):
        """构建合并单元格映射：(row, col) → 左上角单元格的值"""
        merge_map = {}
        for merge_range in worksheet.merged_cells.ranges:
            top_left_value = worksheet.cell(merge_range.min_row, merge_range.min_col).value
            for row in range(merge_range.min_row, merge_range.max_row + 1):
                for col in range(merge_range.min_col, merge_range.max_col + 1):
                    if row == merge_range.min_row and col == merge_range.min_col:
                        continue  # 左上角本身不需要映射
                    merge_map[(row, col)] = top_left_value
        return merge_map

    try:
        for sheet_name in workbook.sheetnames:
            if len(workbook.sheetnames) > 1:
                content += f"## {_normalize_text(sheet_name)}\n\n"

            worksheet = workbook[sheet_name]
            merge_map = _build_merge_map(worksheet)
            rows = list(worksheet.iter_rows(values_only=False))

            if rows:
                # 过滤空行（合并单元格的值也要参与判断）
                non_empty_rows = []
                for row in rows:
                    values = []
                    for cell in row:
                        val = cell.value
                        if val is None:
                            val = merge_map.get((cell.row, cell.column))
                        values.append(val)
                    if any(v is not None and str(v).strip() for v in values):
                        non_empty_rows.append((row, values))

                if non_empty_rows:
                    def last_non_empty_col_count(values):
                        for idx in range(len(values) - 1, -1, -1):
                            if values[idx] is not None and str(values[idx]).strip():
                                return idx + 1
                        return 0

                    max_cols = max(last_non_empty_col_count(v) for _, v in non_empty_rows)

                    if max_cols == 0:
                        continue

                    for i, (_, values) in enumerate(non_empty_rows):
                        row_data = []
                        for j in range(max_cols):
                            if j < len(values) and values[j] is not None:
                                row_data.append(_normalize_table_cell(values[j]))
                            else:
                                row_data.append("")

                        if i == 0:  # 表头
                            content += "| " + " | ".join(row_data) + " |\n"
                            content += "| " + " | ".join(["---"] * len(row_data)) + " |\n"
                        else:
                            content += "| " + " | ".join(row_data) + " |\n"
                    content += "\n"
    finally:
        workbook.close()

    return content.strip()

def convert_pptx(file_path):
    """转换 PowerPoint 文件，提取标题、列表、格式、表格和备注"""
    import pptx
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    presentation = pptx.Presentation(file_path)
    content = ""

    def _process_text_frame(text_frame, is_title=False):
        """处理文本框，保留段落层级和格式"""
        result = ""
        for para in text_frame.paragraphs:
            if not para.text.strip():
                continue

            # 将相邻同格式的 run 合并后再添加 Markdown 标记，避免 **text1****text2** 碎片
            groups = []
            for run in para.runs:
                text = run.text
                if not text:
                    continue
                fmt = (bool(run.font.bold), bool(run.font.italic))
                if groups and groups[-1][0] == fmt:
                    groups[-1] = (fmt, groups[-1][1] + text)
                else:
                    groups.append((fmt, text))
            formatted = ""
            for (bold, italic), text in groups:
                formatted += _format_inline_markdown(text, bold=bold, italic=italic)
            text_value = _normalize_text(formatted.strip() or para.text.strip())

            if is_title:
                result += f"### {text_value}\n\n"
                continue

            # 检查列表层级
            level = para.level if para.level else 0
            if level > 0:
                indent = "  " * level
                result += f"{indent}- {text_value}\n"
            elif hasattr(para, '_pPr') and para._pPr is not None and para._pPr.find(
                './/{http://schemas.openxmlformats.org/drawingml/2006/main}buChar') is not None:
                result += f"- {text_value}\n"
            elif hasattr(para, '_pPr') and para._pPr is not None and para._pPr.find(
                './/{http://schemas.openxmlformats.org/drawingml/2006/main}buAutoNum') is not None:
                result += f"1. {text_value}\n"
            else:
                result += _escape_plain_markdown_text(text_value) + "\n\n"

        return result

    def _iter_shapes(shapes):
        for shape in shapes:
            yield shape
            if getattr(shape, "shape_type", None) == MSO_SHAPE_TYPE.GROUP:
                yield from _iter_shapes(shape.shapes)

    def _shape_is_title(shape):
        if not getattr(shape, "is_placeholder", False):
            return False

        try:
            return shape.placeholder_format.idx in (0, 1)
        except Exception:
            return False

    for i, slide in enumerate(presentation.slides, 1):
        if len(presentation.slides) > 1:
            content += f"## Slide {i}\n\n"

        # 按占位符类型分类处理
        for shape in _iter_shapes(slide.shapes):
            # 处理表格
            if shape.has_table:
                table = shape.table
                all_rows_data = []
                for row in table.rows:
                    row_data = []
                    seen_cells = set()
                    for cell in row.cells:
                        cell_id = id(cell)
                        if cell_id in seen_cells:
                            continue
                        seen_cells.add(cell_id)
                        # 去重合并单元格：跳过非左上角的水平合并续格
                        if hasattr(cell, 'is_merge_origin'):
                            if cell.is_merge_origin:
                                row_data.append(_normalize_table_cell(cell.text))
                            else:
                                row_data.append(_normalize_table_cell(cell.text))
                        else:
                            row_data.append(_normalize_table_cell(cell.text))
                    all_rows_data.append(row_data)
                max_cols = max((len(r) for r in all_rows_data), default=0)
                if max_cols == 0:
                    continue
                for idx, row_data in enumerate(all_rows_data):
                    padded = row_data + [""] * (max_cols - len(row_data))
                    content += "| " + " | ".join(padded) + " |\n"
                    if idx == 0:
                        content += "| " + " | ".join(["---"] * max_cols) + " |\n"
                content += "\n"
                continue

            # 处理文本框
            if shape.has_text_frame and shape.text.strip():
                content += _process_text_frame(shape.text_frame, is_title=_shape_is_title(shape))

        # 提取备注
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
            notes_text = _normalize_text(slide.notes_slide.notes_text_frame.text, preserve_newlines=True)
            if notes_text:
                quoted_lines = "\n".join(f"> {line}" for line in notes_text.splitlines())
                content += f"\n> **备注**:\n{quoted_lines}\n\n"

        if i < len(presentation.slides):
            content += "---\n\n"

    return content.strip()

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

def _extract_pdf_page_blocks(page, tables):
    """提取单页 PDF 的文本和表格块，尽量保持上下顺序"""
    table_bboxes = [t.bbox for t in tables] if tables else []
    blocks = []  # (top_position, content_string)

    for table_obj in tables:
        md = _render_pdf_table(table_obj)
        if md:
            blocks.append((table_obj.bbox[1], md))

    if table_bboxes:
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

        text_chars = filtered_page.chars
        if text_chars:
            split_tops = sorted(set([0.0] + [b[1] for b in table_bboxes] + [b[3] for b in table_bboxes]))
            for seg_idx, seg_top in enumerate(split_tops):
                seg_bottom = split_tops[seg_idx + 1] if seg_idx + 1 < len(split_tops) else float('inf')
                seg_page = filtered_page.filter(
                    lambda obj, t=seg_top, b=seg_bottom: obj.get("top", 0) >= t and obj.get("top", 0) < b
                )
                seg_text = _normalize_text(seg_page.extract_text(), preserve_newlines=True)
                if seg_text:
                    blocks.append((seg_top, "\n".join(_escape_plain_markdown_text(line) for line in seg_text.splitlines()) + "\n\n"))
        else:
            text = _normalize_text(filtered_page.extract_text(), preserve_newlines=True)
            if text:
                blocks.append((0.0, "\n".join(_escape_plain_markdown_text(line) for line in text.splitlines()) + "\n\n"))
    else:
        text = _normalize_text(page.extract_text(), preserve_newlines=True)
        if text:
            blocks.append((0.0, "\n".join(_escape_plain_markdown_text(line) for line in text.splitlines()) + "\n\n"))

    blocks.sort(key=lambda block: block[0])
    return blocks

def convert_pdf(file_path):
    """转换 PDF 文件，支持文本和表格提取，按页面位置交错输出"""
    import pdfplumber

    content_parts = []
    page_errors = []
    with pdfplumber.open(file_path) as pdf:
        for page_number, page in enumerate(pdf.pages, 1):
            try:
                tables = page.find_tables()
                blocks = _extract_pdf_page_blocks(page, tables)
            except Exception as exc:
                page_errors.append((page_number, str(exc)))
                try:
                    fallback_text = _normalize_text(page.extract_text(), preserve_newlines=True)
                except Exception:
                    fallback_text = ""
                if not fallback_text:
                    continue
                blocks = [(0.0, "\n".join(_escape_plain_markdown_text(line) for line in fallback_text.splitlines()) + "\n\n")]

            if not blocks:
                continue

            page_content = "".join(block_content for _, block_content in blocks).strip()
            if not page_content:
                continue

            if len(pdf.pages) > 1:
                content_parts.append(f"## Page {page_number}\n\n{page_content}")
            else:
                content_parts.append(page_content)

    content = "\n\n".join(content_parts).strip()
    if not content and page_errors:
        page_numbers = ", ".join(str(page_number) for page_number, _ in page_errors)
        raise ValueError(f"PDF 解析失败，无法提取任何内容。异常页码: {page_numbers}")
    return content

def convert_md(file_path, output_dir=None):
    """
    将 Markdown 文件转换为 DOCX 格式（通过 Node.js 脚本）

    Args:
        file_path: Markdown 文件路径
        output_dir: 可选的输出目录

    Returns:
        包含 'success'、'output_path' 和可选 'error' 的字典
    """
    # 检查 Node.js 是否可用
    node_cmd = shutil.which('node')
    if not node_cmd:
        return {
            'success': False,
            'error': '未找到 Node.js。Markdown 转 DOCX 需要 Node.js 环境。请安装 Node.js: https://nodejs.org/'
        }

    # 获取 Node.js 脚本路径
    script_dir = os.path.dirname(os.path.abspath(__file__))
    node_script = os.path.join(script_dir, 'md_to_docx', 'index.js')

    if not os.path.exists(node_script):
        return {
            'success': False,
            'error': f'Node.js 转换脚本不存在: {node_script}。请运行 npm install 安装依赖。'
        }

    source_dir = os.path.join(script_dir, 'md_to_docx')
    local_node_modules = os.path.join(source_dir, 'node_modules')
    shared_root = _get_node_shared_root()
    shared_dir = os.path.join(shared_root, 'md_to_docx')
    shared_node_modules = os.path.join(shared_dir, 'node_modules')

    local_mmdc = _find_mmdc_binary(local_node_modules)
    shared_mmdc = _find_mmdc_binary(shared_node_modules)

    use_shared = False
    need_shared = (not os.path.exists(local_node_modules)) or (local_mmdc is None)
    if need_shared and shared_mmdc is None:
        ok, err = _ensure_shared_node_modules(shared_dir, source_dir)
        if not ok:
            return {
                'success': False,
                'error': (
                    f"{err}\n"
                    f"可手动安装：\n"
                    f"  1) 本地安装：cd {source_dir} && npm install\n"
                    f"  2) 共享安装：cd {shared_dir} && npm install\n"
                    f"可通过环境变量 {NODE_SHARED_HOME_ENV} 指定共享目录。"
                )
            }
        shared_mmdc = _find_mmdc_binary(shared_node_modules)

    if need_shared and shared_mmdc:
        use_shared = True

    try:
        # 调用 Node.js 脚本
        cmd = [node_cmd, node_script, file_path]
        if output_dir:
            cmd.append(output_dir)

        env = os.environ.copy()
        if use_shared:
            existing = env.get("NODE_PATH")
            if existing:
                env["NODE_PATH"] = os.pathsep.join([shared_node_modules, existing])
            else:
                env["NODE_PATH"] = shared_node_modules

        mmdc_binary = local_mmdc or shared_mmdc
        if mmdc_binary:
            env["DOCUGENIUS_MMDC_PATH"] = mmdc_binary

        result = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=NODE_CONVERT_TIMEOUT_SECONDS,
            env=env,
        )

        # 解析输出
        try:
            stdout_text = result.stdout or ""
            output = json.loads(stdout_text)
            return output
        except json.JSONDecodeError:
            if result.returncode == 0:
                return {
                    'success': True,
                    'output_path': (result.stdout or "").strip(),
                    'message': '转换成功'
                }
            else:
                stdout_text = (result.stdout or "").strip()
                stderr_text = (result.stderr or "").strip()
                return {
                    'success': False,
                    'error': f'Node.js 脚本输出解析失败: {stdout_text}\n{stderr_text}'
                }

    except subprocess.TimeoutExpired:
        return {
            'success': False,
            'error': '转换超时（超过2分钟）'
        }
    except Exception as e:
        return {
            'success': False,
            'error': f'调用 Node.js 脚本失败: {str(e)}'
        }

def convert_document(file_path, extract_images=True, output_dir=None):
    """
    将文档转换为 Markdown 格式

    Args:
        file_path: 文档文件路径
        extract_images: 是否提取图片（当前版本暂不支持）
        output_dir: 可选的输出目录（默认为同目录下的 Markdown/ 子目录）

    Returns:
        包含 'success'、'markdown_content'、'output_path' 和可选 'error' 的字典
    """
    # 验证输入文件
    file_path, input_error = _validate_input_file(file_path)
    if input_error:
        return {
            'success': False,
            'error': input_error
        }

    # 检查文件大小（限制为100MB）
    try:
        file_size = os.path.getsize(file_path)
        if file_size > MAX_FILE_SIZE_BYTES:
            return {
                'success': False,
                'error': f'文件过大: {file_size / (1024*1024):.2f}MB，超过限制 {MAX_FILE_SIZE_BYTES / (1024*1024):.0f}MB'
            }
    except OSError as e:
        return {
            'success': False,
            'error': f'无法读取文件大小: {str(e)}'
        }

    # 检查文件扩展名
    file_ext = os.path.splitext(file_path)[1].lower()
    if file_ext not in SUPPORTED_EXTENSIONS:
        return {
            'success': False,
            'error': f'不支持的文件格式: {file_ext}。支持的格式: {", ".join(SUPPORTED_EXTENSIONS)}'
        }

    # Markdown 转 DOCX 使用单独的处理流程
    if file_ext == '.md':
        return convert_md(file_path, output_dir)

    # 检查依赖（按格式按需检查，避免无关依赖阻塞）
    deps_ok, error_msg = check_dependencies(file_ext)
    if not deps_ok:
        return {
            'success': False,
            'error': error_msg
        }

    try:
        # 根据文件类型转换
        if file_ext == '.docx':
            markdown_content = convert_docx(file_path)
        elif file_ext == '.xlsx':
            markdown_content = convert_xlsx(file_path)
        elif file_ext == '.pptx':
            markdown_content = convert_pptx(file_path)
        elif file_ext == '.pdf':
            markdown_content = convert_pdf(file_path)
        else:
            return {
                'success': False,
                'error': f'不支持的文件类型: {file_ext}'
            }

        warning = None
        if not markdown_content.strip():
            if file_ext == '.pdf':
                return {
                    'success': False,
                    'error': 'PDF 未提取到任何文本或表格，文件可能是扫描件、受保护文档，或仅包含图片。请先进行 OCR 或解除保护后再试。'
                }
            warning = '未提取到任何可写入的内容，原文档可能为空，或仅包含当前版本暂不支持的对象。'

        # 确定输出路径
        output_path = _resolve_markdown_output_path(file_path, output_dir)

        # 保存 Markdown 文件
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(markdown_content)

        result = {
            'success': True,
            'markdown_content': markdown_content,
            'output_path': output_path
        }
        if warning:
            result['warning'] = warning
        return result

    except PermissionError as e:
        return {
            'success': False,
            'error': f'权限不足: 无法读取文件或写入输出目录 - {str(e)}'
        }
    except MemoryError:
        return {
            'success': False,
            'error': '内存不足: 文件可能过大，请尝试处理较小的文件'
        }
    except FileNotFoundError as e:
        return {
            'success': False,
            'error': f'文件未找到: {str(e)}'
        }
    except OSError as e:
        return {
            'success': False,
            'error': f'系统错误: {str(e)}'
        }
    except Exception as e:
        return {
            'success': False,
            'error': f'转换错误 ({type(e).__name__}): {str(e)}'
        }

def batch_convert(directory, recursive=True, extract_images=True, output_dir=None):
    """
    批量转换目录中的所有支持的文档

    Args:
        directory: 要扫描的目录
        recursive: 是否递归扫描子目录
        extract_images: 是否提取图片
        output_dir: 可选的输出目录

    Returns:
        转换结果列表
    """
    results = []

    normalized_directory = os.path.abspath(os.path.normpath(os.path.expanduser(str(directory))))
    if not os.path.exists(normalized_directory):
        return [{
            'file': normalized_directory,
            'result': {
                'success': False,
                'error': f'目录不存在: {normalized_directory}'
            }
        }]
    if not os.path.isdir(normalized_directory):
        return [{
            'file': normalized_directory,
            'result': {
                'success': False,
                'error': f'输入路径不是目录: {normalized_directory}'
            }
        }]

    for file_path in _iter_batch_input_files(normalized_directory, recursive=recursive, output_dir=output_dir):
        result = convert_document(file_path, extract_images, output_dir)
        results.append({
            'file': file_path,
            'result': result
        })

    return results

def main():
    if len(sys.argv) < 2:
        print('用法: python convert_document.py <file_path> [extract_images] [output_dir]')
        print('  file_path: 文档文件路径')
        print('  extract_images: true/false (默认: true，当前版本暂不支持)')
        print('  output_dir: 可选的输出目录')
        print('')
        print('支持的格式:')
        print('  - Office/PDF 转 Markdown: .docx, .xlsx, .pptx, .pdf')
        print('  - Markdown 转 Word: .md')
        print('')
        print('批量转换: python convert_document.py --batch <directory> [recursive]')
        print('  directory: 要扫描的目录')
        print('  recursive: true/false (默认: true)')
        sys.exit(1)

    # 批量转换模式
    if sys.argv[1] == '--batch':
        if len(sys.argv) < 3:
            print('错误: 批量转换需要指定目录')
            sys.exit(1)

        directory = sys.argv[2]
        recursive = sys.argv[3].lower() == 'true' if len(sys.argv) > 3 else True

        results = batch_convert(directory, recursive)

        # 输出结果统计
        success_count = sum(1 for r in results if r['result']['success'])
        total_count = len(results)

        print(json.dumps({
            'total': total_count,
            'success': success_count,
            'failed': total_count - success_count,
            'results': results
        }, ensure_ascii=False, indent=2))

        sys.exit(0 if success_count == total_count else 1)

    # 单文件转换模式
    file_path = sys.argv[1]
    extract_images = sys.argv[2].lower() == 'true' if len(sys.argv) > 2 else True
    output_dir = sys.argv[3] if len(sys.argv) > 3 else None

    result = convert_document(file_path, extract_images, output_dir)

    # 输出结果为 JSON
    print(json.dumps(result, ensure_ascii=False, indent=2))

    sys.exit(0 if result['success'] else 1)

if __name__ == '__main__':
    main()
