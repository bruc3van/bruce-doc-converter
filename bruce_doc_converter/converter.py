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
from bruce_doc_converter.output import atomic_write, reserve_output
from bruce_doc_converter.formats.common import (
    DOCX_XML_NAMESPACES,
    DOCX_W_NS,
    IMAGE_OUTPUT_DIR_NAME,
    MIN_IMAGE_DIMENSION_PX,
    MAX_ASPECT_RATIO,
    MIN_IMAGE_DATA_BYTES,
    PPTX_BACKGROUND_COVERAGE_RATIO,
    OOXML_IMAGE_NAMESPACES,
    _IMAGE_SIGNATURES,
    _RE_COLLAPSE_WHITESPACE,
    _RE_COLLAPSE_EXTRA_BLANK_LINES,
    _RE_ESCAPE_MARKDOWN_LEADING,
    _RE_ESCAPE_MARKDOWN_ORDERED_LIST,
    _RE_WRAP_INLINE_MARKDOWN,
    _normalize_text,
    _escape_plain_markdown_text,
    _format_inline_markdown,
    _compose_inline_markdown,
    _normalize_table_cell,
    _table_position_has_content,
)
from bruce_doc_converter.formats.images import (
    _detect_image_format,
    _get_image_dimensions,
    _is_decorative_image,
    _setup_image_output_dir,
    _save_extracted_image,
    _make_image_markdown,
    _check_ooxml_decorative_flag,
)
from bruce_doc_converter.formats.docx import (
    _resolve_docx_style_font_flag,
    _get_docx_heading_level,
    _resolve_docx_run_font_flag,
    _get_docx_grid_span,
    _is_docx_vertical_merge_continuation,
    _extract_docx_table_cell_text,
    _docx_attr,
    _build_docx_numbering_index,
    _get_docx_style_numpr,
    _get_docx_paragraph_numpr,
    _to_roman,
    _to_alpha,
    _to_chinese_counting,
    _to_circled_number,
    _format_docx_number_value,
    _render_docx_list_marker,
    _is_docx_toc_paragraph,
    convert_docx,
)
from bruce_doc_converter.formats.xlsx import (
    convert_xlsx,
)
from bruce_doc_converter.formats.pptx import (
    convert_pptx,
)
from bruce_doc_converter.formats.pdf import (
    _render_pdf_table,
    _group_words_into_lines,
    _reconstruct_line_text,
    _get_body_font_size,
    _get_line_avg_font_size,
    _detect_column_split,
    _split_pdf_words_by_columns,
    _split_markdown_blocks,
    _parse_pdf_academic_section_block,
    _is_markdown_heading_block,
    _format_pdf_keywords_block,
    _format_pdf_references_block,
    _format_pdf_academic_section,
    _postprocess_pdf_academic_sections,
    _lines_to_markdown_blocks,
    _extract_pdf_page_blocks,
    convert_pdf,
)

SUPPORTED_EXTENSIONS = ['.docx', '.xlsx', '.pptx', '.pdf', '.md']
MAX_FILE_SIZE_BYTES = 100 * 1024 * 1024
NODE_CONVERT_TIMEOUT_SECONDS = 120
NODE_INSTALL_TIMEOUT_SECONDS = 300
NODE_SHARED_HOME_ENV = "BRUCE_DOC_CONVERTER_NODE_HOME"
BROWSER_PATH_ENV = "BRUCE_DOC_CONVERTER_CHROME_PATH"
GENERATED_OUTPUT_DIR_NAMES = {"Markdown", "Word"}
# Markdown 转 DOCX 运行时必需的 Node.js 包（Mermaid CLI 另行检查）
NODE_RUNTIME_PACKAGES = ("docx", "jsdom", "markdown-it", "markdown-it-footnote", "temml", "jszip", "saxes")



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
        return os.path.join(base, "BruceDocConverter", "node")

    return os.path.join(os.path.expanduser("~"), ".bruce-doc-converter", "node")

def _candidate_browser_paths():
    paths = []
    override = os.environ.get(BROWSER_PATH_ENV)
    if override:
        paths.append(override)

    if sys.platform == "win32":
        program_files = [
            os.environ.get("PROGRAMFILES"),
            os.environ.get("PROGRAMFILES(X86)"),
            os.environ.get("LOCALAPPDATA"),
        ]
        for base in [p for p in program_files if p]:
            paths.extend([
                os.path.join(base, "Google", "Chrome", "Application", "chrome.exe"),
                os.path.join(base, "Microsoft", "Edge", "Application", "msedge.exe"),
                os.path.join(base, "Chromium", "Application", "chrome.exe"),
            ])
    elif sys.platform == "darwin":
        home = os.path.expanduser("~")
        paths.extend([
            "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
            "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
            "/Applications/Chromium.app/Contents/MacOS/Chromium",
            os.path.join(home, "Applications", "Google Chrome.app", "Contents", "MacOS", "Google Chrome"),
            os.path.join(home, "Applications", "Microsoft Edge.app", "Contents", "MacOS", "Microsoft Edge"),
            os.path.join(home, "Applications", "Chromium.app", "Contents", "MacOS", "Chromium"),
        ])

    for command in ("google-chrome", "chrome", "chromium", "chromium-browser", "msedge", "microsoft-edge"):
        found = shutil.which(command)
        if found:
            paths.append(found)

    return paths

def _find_local_browser():
    seen = set()
    for candidate in _candidate_browser_paths():
        if not candidate:
            continue
        normalized = os.path.abspath(os.path.normpath(os.path.expanduser(str(candidate))))
        key = normalized.lower() if sys.platform == "win32" else normalized
        if key in seen:
            continue
        seen.add(key)
        if os.path.isfile(normalized):
            return normalized
    return None

def _markdown_contains_mermaid(file_path):
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            content = f.read()
    except UnicodeDecodeError:
        with open(file_path, "r", encoding="utf-8-sig", errors="replace") as f:
            content = f.read()
    return bool(re.search(r"(^|\n)[ \t>]*(?:`{3,}|~{3,})[ \t]*mermaid\b", content, re.IGNORECASE))

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

def _ensure_shared_node_modules(shared_dir, source_dir, allow_scripts=False):
    version_error = _node_version_error()
    if version_error:
        return False, version_error
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

    install_action = "ci" if os.path.exists(os.path.join(shared_dir, "package-lock.json")) else "install"
    cmd = [npm_cmd, install_action, "--no-fund", "--no-audit"]
    if not allow_scripts:
        cmd.append("--ignore-scripts")
    try:
        print("[BruceDocConverter] 正在安装 Node.js 依赖到用户共享目录...", file=sys.stderr)
        result = subprocess.run(
            cmd,
            cwd=shared_dir,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=NODE_INSTALL_TIMEOUT_SECONDS,
        )
        if result.returncode != 0:
            error_output = result.stderr.strip() or result.stdout.strip()
            return False, f"Node.js 依赖安装失败: {error_output}"
    except subprocess.TimeoutExpired:
        return False, f"Node.js 依赖安装超时（超过{NODE_INSTALL_TIMEOUT_SECONDS}秒）"
    except Exception as e:
        return False, f"Node.js 依赖安装失败: {str(e)}"

    return True, None


def _node_version_error():
    failure = _node_runtime_failure()
    return failure['error'] if failure else None


def _node_runtime_failure():
    node = shutil.which('node')
    if not node:
        return _error_result('NODE_NOT_FOUND', '未找到 Node.js。需要 Node.js >=22.0。')
    try:
        result = subprocess.run([node, '--version'], capture_output=True, text=True, timeout=10)
        version = tuple(int(part) for part in result.stdout.strip().lstrip('v').split('.')[:2])
        if result.returncode != 0 or len(version) != 2 or version < (22, 0):
            return _error_result('NODE_VERSION_UNSUPPORTED', f'需要 Node.js >=22.0，当前版本: {result.stdout.strip() or "unknown"}')
    except (OSError, ValueError, subprocess.TimeoutExpired) as exc:
        return _error_result('NODE_VERSION_CHECK_FAILED', f'无法检查 Node.js 版本: {exc}')
    return None

def _ensure_puppeteer_browser(shared_dir):
    npm_cmd = shutil.which('npm')
    if not npm_cmd:
        return False, "未找到 npm。请安装 Node.js（自带 npm）后重试。"

    cmd = [
        npm_cmd,
        "exec",
        "puppeteer",
        "--",
        "browsers",
        "install",
        "chrome-headless-shell",
    ]
    try:
        print("[BruceDocConverter] 正在安装 Mermaid 渲染所需的 Chromium 浏览器...", file=sys.stderr)
        result = subprocess.run(
            cmd,
            cwd=shared_dir,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=NODE_INSTALL_TIMEOUT_SECONDS,
        )
        if result.returncode != 0:
            error_output = result.stderr.strip() or result.stdout.strip()
            return False, f"Chromium 浏览器安装失败: {error_output}"
    except subprocess.TimeoutExpired:
        return False, f"Chromium 浏览器安装超时（超过{NODE_INSTALL_TIMEOUT_SECONDS}秒）"
    except Exception as e:
        return False, f"Chromium 浏览器安装失败: {str(e)}"

    return True, None

def _find_mmdc_binary(node_modules_dir):
    if not node_modules_dir:
        return None

    ext = ".cmd" if sys.platform == "win32" else ""
    candidate = os.path.join(node_modules_dir, ".bin", f"mmdc{ext}")
    if os.path.exists(candidate):
        return candidate
    return None

def _files_have_same_content(left, right):
    if not os.path.exists(left) or not os.path.exists(right):
        return False
    try:
        with open(left, "rb") as left_file, open(right, "rb") as right_file:
            return left_file.read() == right_file.read()
    except Exception:
        return False

def _shared_node_dependencies_ready(shared_dir, source_dir):
    for filename in ("package.json", "package-lock.json"):
        if not _files_have_same_content(
            os.path.join(source_dir, filename),
            os.path.join(shared_dir, filename),
        ):
            return False

    node_modules_dir = os.path.join(shared_dir, "node_modules")
    required_paths = [os.path.join(node_modules_dir, package) for package in NODE_RUNTIME_PACKAGES]
    required_paths.append(os.path.join(node_modules_dir, "@mermaid-js", "mermaid-cli"))
    if any(not os.path.exists(path) for path in required_paths):
        return False

    return _find_mmdc_binary(node_modules_dir) is not None

# ==================== 依赖检查函数 ====================

def check_dependencies(file_ext=None, auto_install=False):
    """
    检查必需的依赖是否已安装（默认检查全部；传入 file_ext 时仅检查该格式所需）

    Args:
        file_ext: 文件扩展名（如 '.docx'），仅检查该格式所需的依赖
        auto_install: 保留的兼容参数。发布包不会在运行时自动安装 Python 依赖。

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

    if missing:
        return False, (
            f"缺少依赖库: {', '.join(missing_pip_names)}。"
            "请重新安装发布包以恢复依赖: pipx reinstall bruce-doc-converter"
        )

    return True, None


def _validate_input_file(file_path):
    """校验并规范化输入文件路径"""
    if file_path is None:
        return None, "文件路径不能为空", "USAGE_ERROR"

    normalized = os.path.abspath(os.path.normpath(os.path.expanduser(str(file_path))))
    if not os.path.exists(normalized):
        return None, f'文件不存在: {normalized}', "FILE_NOT_FOUND"
    if not os.path.isfile(normalized):
        return None, f'输入路径不是文件: {normalized}', "NOT_A_FILE"
    return normalized, None, None

def _error_result(error_code, error):
    return {
        'success': False,
        'error_code': error_code,
        'error': error,
    }

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
    base_name, source_ext = os.path.splitext(os.path.basename(file_path))
    output_filename = base_name + '.md'
    output_path = os.path.join(target_dir, output_filename)
    if not os.path.exists(output_path):
        return output_path

    source_suffix = source_ext.lower() or ".input"
    output_path = os.path.join(target_dir, f"{base_name}{source_suffix}.md")
    if not os.path.exists(output_path):
        return output_path

    counter = 2
    while True:
        output_path = os.path.join(target_dir, f"{base_name}{source_suffix}.{counter}.md")
        if not os.path.exists(output_path):
            return output_path
        counter += 1

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


# ==================== 图片提取公共基础设施 ====================


# ---------- PDF 词元级文本重建辅助函数 ----------


def _normalize_mermaid_scale(mermaid_scale):
    if mermaid_scale is None:
        return None
    try:
        scale = float(mermaid_scale)
    except (TypeError, ValueError):
        raise ValueError("Mermaid scale 必须是正数")
    if not scale > 0:
        raise ValueError("Mermaid scale 必须大于 0")
    return scale

def _format_mermaid_scale(scale):
    if float(scale).is_integer():
        return str(int(scale))
    return str(scale)

def convert_md(file_path, output_dir=None, mermaid_scale=None, strict=False):
    """
    将 Markdown 文件转换为 DOCX 格式（通过 Node.js 脚本）

    Args:
        file_path: Markdown 文件路径
        output_dir: 可选的输出目录
        mermaid_scale: Mermaid PNG 渲染倍率，默认由 Node 渲染器决定

    Returns:
        包含 'success'、'output_path' 和可选 'error' 的字典
    """
    # 检查 Node.js 是否可用
    node_cmd = shutil.which('node')
    if not node_cmd:
        return _error_result(
            'NODE_NOT_FOUND',
            '未找到 Node.js。Markdown 转 DOCX 需要 Node.js 环境。请安装 Node.js: https://nodejs.org/'
        )

    # 获取 Node.js 脚本路径
    script_dir = os.path.dirname(os.path.abspath(__file__))
    node_script = os.path.join(script_dir, 'md_to_docx', 'index.js')

    if not os.path.exists(node_script):
        return _error_result(
            'NODE_CONVERSION_FAILED',
            f'Node.js 转换脚本不存在: {node_script}。请重新安装发布包: pipx reinstall bruce-doc-converter'
        )

    source_dir = os.path.join(script_dir, 'md_to_docx')
    local_node_modules = os.path.join(source_dir, 'node_modules')
    shared_root = _get_node_shared_root()
    shared_dir = os.path.join(shared_root, 'md_to_docx')
    shared_node_modules = os.path.join(shared_dir, 'node_modules')

    local_mmdc = _find_mmdc_binary(local_node_modules)
    shared_mmdc = _find_mmdc_binary(shared_node_modules)

    use_shared = False
    local_packages_ready = all(os.path.exists(os.path.join(local_node_modules, package))
                               for package in NODE_RUNTIME_PACKAGES)
    need_shared = not local_packages_ready or local_mmdc is None
    if need_shared and (shared_mmdc is None or not _shared_node_dependencies_ready(shared_dir, source_dir)):
        return _error_result(
            'DEPENDENCY_INSTALL_REQUIRED',
            (
                "Markdown 转 Word 需要先显式安装 Node.js 依赖。"
                "请运行: bdc setup-node"
            )
        )

    if need_shared and shared_mmdc:
        use_shared = True

    try:
        # 调用 Node.js 脚本
        cmd = [node_cmd, node_script, file_path]
        if output_dir:
            cmd.append(output_dir)

        env = os.environ.copy()
        env['BRUCE_DOC_CONVERTER_STRICT'] = '1' if strict else '0'
        if use_shared:
            existing = env.get("NODE_PATH")
            if existing:
                env["NODE_PATH"] = os.pathsep.join([shared_node_modules, existing])
            else:
                env["NODE_PATH"] = shared_node_modules

        mmdc_binary = local_mmdc or shared_mmdc
        if mmdc_binary:
            env["BRUCE_DOC_CONVERTER_MMDC_PATH"] = mmdc_binary
        normalized_scale = _normalize_mermaid_scale(mermaid_scale)
        if normalized_scale is not None:
            env["BRUCE_DOC_CONVERTER_MMDC_SCALE"] = _format_mermaid_scale(normalized_scale)
        if _markdown_contains_mermaid(file_path):
            browser_path = _find_local_browser()
            if browser_path:
                env[BROWSER_PATH_ENV] = browser_path

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
            if not output.get('success') and not output.get('error_code'):
                output['error_code'] = 'NODE_CONVERSION_FAILED'
            # 规范化 Node 侧 warnings，便于 CLI 统一透传
            if output.get('success') and 'warnings' in output and not isinstance(output.get('warnings'), list):
                output['warnings'] = [str(output['warnings'])]
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
                return _error_result(
                    'NODE_CONVERSION_FAILED',
                    f'Node.js 脚本输出解析失败: {stdout_text}\n{stderr_text}'
                )

    except subprocess.TimeoutExpired:
        return _error_result('CONVERSION_TIMEOUT', '转换超时（超过2分钟）')
    except Exception as e:
        return _error_result('NODE_CONVERSION_FAILED', f'调用 Node.js 脚本失败: {str(e)}')

def setup_node_dependencies(allow_scripts=False, install_browser=False):
    """显式安装 Markdown -> DOCX 所需的 Node.js 依赖到用户共享目录。"""
    runtime_failure = _node_runtime_failure()
    if runtime_failure:
        return runtime_failure
    source_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'md_to_docx')
    shared_root = _get_node_shared_root()
    shared_dir = os.path.join(shared_root, 'md_to_docx')
    if _shared_node_dependencies_ready(shared_dir, source_dir):
        browser_install_action = 'not_requested'
        if install_browser:
            ok, err = _ensure_puppeteer_browser(shared_dir)
            if not ok:
                return _error_result('DEPENDENCY_INSTALL_FAILED', err)
            browser_install_action = 'installed_or_verified'
        return {
            'success': True,
            'node_home': shared_dir,
            'allow_scripts': bool(allow_scripts),
            'already_installed': True,
            'install_action': 'skipped',
            'browser_install_action': browser_install_action,
            'detected_browser_path': _find_local_browser(),
        }

    ok, err = _ensure_shared_node_modules(shared_dir, source_dir, allow_scripts=allow_scripts)
    if not ok:
        return _error_result('DEPENDENCY_INSTALL_FAILED', err)
    browser_install_action = 'not_requested'
    if install_browser:
        ok, err = _ensure_puppeteer_browser(shared_dir)
        if not ok:
            return _error_result('DEPENDENCY_INSTALL_FAILED', err)
        browser_install_action = 'installed_or_verified'
    return {
        'success': True,
        'node_home': shared_dir,
        'allow_scripts': bool(allow_scripts),
        'already_installed': False,
        'install_action': 'installed',
        'browser_install_action': browser_install_action,
        'detected_browser_path': _find_local_browser(),
    }

def convert_document(file_path, extract_images=True, output_dir=None, mermaid_scale=None, strict=False):
    """
    将文档转换为 Markdown 格式

    Args:
        file_path: 文档文件路径
        extract_images: 是否提取图片（默认 True，支持 Word/Excel/PowerPoint）
        output_dir: 可选的输出目录（默认为同目录下的 Markdown/ 子目录）
        mermaid_scale: Markdown 转 Word 时 Mermaid PNG 渲染倍率

    Returns:
        包含 'success'、'markdown_content'、'output_path'、可选 'extracted_images' 和 'error' 的字典
    """
    # 验证输入文件
    file_path, input_error, input_error_code = _validate_input_file(file_path)
    if input_error:
        return _error_result(input_error_code, input_error)

    # 检查文件大小（限制为100MB）
    try:
        file_size = os.path.getsize(file_path)
        if file_size > MAX_FILE_SIZE_BYTES:
            return _error_result(
                'FILE_TOO_LARGE',
                f'文件过大: {file_size / (1024*1024):.2f}MB，超过限制 {MAX_FILE_SIZE_BYTES / (1024*1024):.0f}MB'
            )
    except OSError as e:
        return _error_result('OS_ERROR', f'无法读取文件大小: {str(e)}')

    # 检查文件扩展名
    file_ext = os.path.splitext(file_path)[1].lower()
    if file_ext not in SUPPORTED_EXTENSIONS:
        return _error_result(
            'UNSUPPORTED_FORMAT',
            f'不支持的文件格式: {file_ext}。支持的格式: {", ".join(SUPPORTED_EXTENSIONS)}'
        )

    # Markdown 转 DOCX 使用单独的处理流程
    if file_ext == '.md':
        try:
            _normalize_mermaid_scale(mermaid_scale)
        except ValueError as e:
            return _error_result('USAGE_ERROR', str(e))
        return convert_md(file_path, output_dir, mermaid_scale=mermaid_scale, strict=strict)

    # 检查依赖（按格式按需检查，避免无关依赖阻塞）
    deps_ok, error_msg = check_dependencies(file_ext)
    if not deps_ok:
        return _error_result('DEPENDENCY_INSTALL_FAILED', error_msg)

    output_path, image_save_dir = None, None
    committed = False
    diagnostics = []
    try:
        # 预先确定输出路径，以便设置图片目录
        output_path = reserve_output(lambda: _resolve_markdown_output_path(file_path, output_dir))

        # 设置图片提取目录
        image_save_dir = None
        image_rel_dir = None
        if extract_images and file_ext in ('.docx', '.xlsx', '.pptx'):
            image_save_dir, image_rel_dir = _setup_image_output_dir(output_path)

        # 根据文件类型转换
        extracted_images = []
        if file_ext == '.docx':
            markdown_content, extracted_images = convert_docx(
                file_path, image_save_dir=image_save_dir, image_rel_dir=image_rel_dir, diagnostics=diagnostics
            )
        elif file_ext == '.xlsx':
            markdown_content, extracted_images = convert_xlsx(
                file_path, image_save_dir=image_save_dir, image_rel_dir=image_rel_dir, diagnostics=diagnostics
            )
        elif file_ext == '.pptx':
            markdown_content, extracted_images = convert_pptx(
                file_path, image_save_dir=image_save_dir, image_rel_dir=image_rel_dir, diagnostics=diagnostics
            )
        elif file_ext == '.pdf':
            markdown_content = convert_pdf(file_path, diagnostics=diagnostics)
        else:
            return _error_result('UNSUPPORTED_FORMAT', f'不支持的文件类型: {file_ext}')

        warning = None
        if not markdown_content.strip():
            if file_ext == '.pdf':
                return {**_error_result(
                    'EMPTY_PDF_CONTENT',
                    'PDF 未提取到任何文本或表格，文件可能是扫描件、受保护文档，或仅包含图片。请先进行 OCR 或解除保护后再试。'
                ), 'diagnostics': diagnostics, 'warnings': [d['message'] for d in diagnostics if d['severity'] == 'warning']}
            warning = '未提取到任何可写入的内容，原文档可能为空，或仅包含当前版本暂不支持的对象。'

        warnings = [d['message'] for d in diagnostics if d['severity'] == 'warning']
        if warning:
            warnings.append(warning)
        if strict and warnings:
            return {**_error_result('CONTENT_INCOMPLETE', '严格模式拒绝包含内容缺失或降级警告的转换。'),
                    'warnings': warnings, 'diagnostics': diagnostics}
        atomic_write(output_path, markdown_content)
        committed = True

        result = {
            'success': True,
            'markdown_content': markdown_content,
            'output_path': output_path
        }
        if extracted_images:
            result['extracted_images'] = extracted_images
        if warning:
            result['warning'] = warning
        if warnings:
            result['warnings'] = warnings
        if diagnostics:
            result['diagnostics'] = diagnostics
        return result

    except PermissionError as e:
        return _error_result('PERMISSION_DENIED', f'权限不足: 无法读取文件或写入输出目录 - {str(e)}')
    except MemoryError:
        return _error_result('OUT_OF_MEMORY', '内存不足: 文件可能过大，请尝试处理较小的文件')
    except FileNotFoundError as e:
        return _error_result('FILE_NOT_FOUND', f'文件未找到: {str(e)}')
    except OSError as e:
        return _error_result('OS_ERROR', f'系统错误: {str(e)}')
    except Exception as e:
        return {**_error_result('CONVERSION_ERROR', f'转换错误 ({type(e).__name__}): {str(e)}'),
                'diagnostics': diagnostics, 'warnings': [d['message'] for d in diagnostics if d['severity'] == 'warning']}
    finally:
        if not committed:
            if output_path and os.path.exists(output_path):
                os.unlink(output_path)
            if image_save_dir and os.path.isdir(image_save_dir):
                shutil.rmtree(image_save_dir)

def iter_batch_convert(directory, recursive=True, extract_images=True, output_dir=None, mermaid_scale=None, strict=False):
    """
    批量转换目录中的所有支持的文档

    Args:
        directory: 要扫描的目录
        recursive: 是否递归扫描子目录
        extract_images: 是否提取图片
        output_dir: 可选的输出目录
        mermaid_scale: Markdown 转 Word时 Mermaid PNG 渲染倍率

    Returns:
        转换结果列表
    """
    normalized_directory = os.path.abspath(os.path.normpath(os.path.expanduser(str(directory))))
    if not os.path.exists(normalized_directory):
        yield {
            'file': normalized_directory,
            'result': _error_result('FILE_NOT_FOUND', f'目录不存在: {normalized_directory}')
        }
        return
    if not os.path.isdir(normalized_directory):
        yield {
            'file': normalized_directory,
            'result': _error_result('NOT_A_DIRECTORY', f'输入路径不是目录: {normalized_directory}')
        }
        return

    normalized_output_dir = None
    if output_dir:
        normalized_output_dir = os.path.abspath(os.path.normpath(os.path.expanduser(str(output_dir))))

    for file_path in _iter_batch_input_files(normalized_directory, recursive=recursive, output_dir=output_dir):
        file_output_dir = normalized_output_dir
        if normalized_output_dir and recursive:
            relative_parent = os.path.relpath(os.path.dirname(file_path), normalized_directory)
            if relative_parent != '.':
                file_output_dir = os.path.join(normalized_output_dir, relative_parent)

        result = convert_document(
            file_path,
            extract_images=extract_images,
            output_dir=file_output_dir,
            mermaid_scale=mermaid_scale,
            strict=strict,
        )
        yield {
            'file': file_path,
            'result': result
        }


def batch_convert(directory, recursive=True, extract_images=True, output_dir=None, mermaid_scale=None, strict=False):
    """Compatibility API; streaming consumers should use iter_batch_convert."""
    return list(iter_batch_convert(directory, recursive, extract_images, output_dir, mermaid_scale, strict))
