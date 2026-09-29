import argparse
import json
import os
import sys
import tempfile
from bruce_doc_converter import __version__
from bruce_doc_converter.converter import (
    SUPPORTED_EXTENSIONS,
    iter_batch_convert,
    convert_document,
    setup_node_dependencies,
)

SCHEMA_VERSION = "1.0"

SUGGESTIONS = {
    "UNSUPPORTED_FORMAT": "请先转换为 .docx/.xlsx/.pptx 后再重试。",
    "NODE_NOT_FOUND": "请安装 Node.js 后重试 Markdown 到 Word 转换。",
    "NODE_VERSION_UNSUPPORTED": "请升级到 Node.js >=22.0 后重试。",
    "DEPENDENCY_INSTALL_REQUIRED": "请先运行 bdc setup-node 安装 Markdown 到 Word 所需的 Node.js 依赖。",
    "EMPTY_PDF_CONTENT": "请先对扫描件执行 OCR，或解除 PDF 保护后重试。",
}

NEXT_COMMANDS = {
    "DEPENDENCY_INSTALL_REQUIRED": "bdc setup-node",
}

RETRYABLE_ERRORS = {
    "DEPENDENCY_INSTALL_REQUIRED",
}


def _emit(payload, exit_code):
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    return exit_code


def _usage_error(message):
    return {
        "schema_version": SCHEMA_VERSION,
        "success": False,
        "error_code": "USAGE_ERROR",
        "error": message,
        "retryable": False,
    }


class _JsonArgumentParser(argparse.ArgumentParser):
    """ArgumentParser that emits JSON on error instead of human-readable text."""

    def error(self, message):
        _emit(_usage_error(message), 1)
        sys.exit(1)


def _format_of(path):
    ext = os.path.splitext(str(path))[1].lower()
    if not ext:
        return "unknown"
    return ext[1:] if ext.startswith(".") else ext


def _output_format(input_format):
    return "docx" if input_format == "md" else "markdown"


def _classify_error(error):
    text = error or ""
    if "文件不存在" in text or "目录不存在" in text or "文件未找到" in text:
        return "FILE_NOT_FOUND"
    if "输入路径不是文件" in text:
        return "NOT_A_FILE"
    if "输入路径不是目录" in text or "输出路径不是目录" in text:
        return "NOT_A_DIRECTORY"
    if "文件过大" in text:
        return "FILE_TOO_LARGE"
    if "不支持的文件格式" in text or "不支持的文件类型" in text:
        return "UNSUPPORTED_FORMAT"
    if "未找到 Node.js" in text:
        return "NODE_NOT_FOUND"
    if "需要先显式安装 Node.js 依赖" in text:
        return "DEPENDENCY_INSTALL_REQUIRED"
    if "Node.js 依赖安装失败" in text or "依赖安装失败" in text:
        return "DEPENDENCY_INSTALL_FAILED"
    if "Node.js 脚本输出解析失败" in text or "调用 Node.js 脚本失败" in text:
        return "NODE_CONVERSION_FAILED"
    if "转换超时" in text:
        return "CONVERSION_TIMEOUT"
    if "PDF 未提取到任何文本或表格" in text:
        return "EMPTY_PDF_CONTENT"
    if "权限不足" in text:
        return "PERMISSION_DENIED"
    if "内存不足" in text:
        return "OUT_OF_MEMORY"
    if "系统错误" in text:
        return "OS_ERROR"
    return "CONVERSION_ERROR"


def _normalize_single_result(input_path, result, content_mode='full', preview_chars=4000):
    normalized_input = os.path.realpath(os.path.expanduser(str(input_path)))
    input_format = _format_of(normalized_input)

    if result.get("success"):
        raw_output = result.get("output_path")
        output_path = os.path.realpath(raw_output) if raw_output else None
        payload = {
            "schema_version": SCHEMA_VERSION,
            "success": True,
            "input_path": normalized_input,
            "input_format": input_format,
            "output_format": _output_format(input_format),
            "output_path": output_path,
            "warnings": [],
        }
        if input_format == "md":
            payload["message"] = result.get("message", "")
        else:
            content = result.get("markdown_content", "")
            payload['markdown_chars'] = len(content)
            payload['content_mode'] = content_mode
            if content_mode != 'none':
                payload["markdown_content"] = content if content_mode == 'full' else content[:preview_chars]
                payload['markdown_truncated'] = len(payload['markdown_content']) < len(content)
            payload["extracted_images"] = result.get("extracted_images", [])
        if result.get("warning"):
            payload["warnings"].append(result["warning"])
        extra_warnings = result.get("warnings")
        if isinstance(extra_warnings, list):
            for item in extra_warnings:
                if item is None:
                    continue
                text = item if isinstance(item, str) else str(item)
                if text and text not in payload["warnings"]:
                    payload["warnings"].append(text)
        if result.get('diagnostics'):
            payload['diagnostics'] = result['diagnostics']
        return payload

    error = result.get("error", "转换失败")
    error_code = result.get("error_code") or _classify_error(error)
    payload = {
        "schema_version": SCHEMA_VERSION,
        "success": False,
        "input_path": normalized_input,
        "input_format": input_format,
        "error_code": error_code,
        "error": error,
        "retryable": error_code in RETRYABLE_ERRORS,
    }
    if error_code in SUGGESTIONS:
        payload["suggestion"] = SUGGESTIONS[error_code]
    if error_code in NEXT_COMMANDS:
        payload["next_command"] = NEXT_COMMANDS[error_code]
    for key in ('warnings', 'diagnostics'):
        if result.get(key):
            payload[key] = result[key]
    return payload


PACKAGE_NAME = "bruce-doc-converter"


def _install_info(executable=None, prefix=None, base_prefix=None, module_file=None, user_site=None):
    """Describe how this CLI was installed so agents can pick a safe upgrade path.

    auto_upgrade is true only for isolated tool installs owned by the user
    (pipx, uv tool, pip --user). Project venvs, system Python and source
    checkouts are left to the user.
    """
    executable = executable or sys.executable
    prefix = prefix if prefix is not None else sys.prefix
    base_prefix = base_prefix if base_prefix is not None else getattr(sys, "base_prefix", sys.prefix)
    module_file = module_file or os.path.abspath(__file__)
    if user_site is None:
        try:
            import site
            user_site = site.getusersitepackages()
        except Exception:
            user_site = ""
    module_path = module_file.replace("\\", "/").lower()
    quoted_python = f'"{executable}"'

    # pipx and uv mark their tool environments, which also covers custom PIPX_HOME / UV_TOOL_DIR.
    if "/site-packages/" not in module_path and "/dist-packages/" not in module_path:
        source, command = "source", None
    elif os.path.isfile(os.path.join(prefix, "pipx_metadata.json")):
        source, command = "pipx", f"pipx upgrade {PACKAGE_NAME}"
    elif os.path.isfile(os.path.join(prefix, "uv-receipt.toml")):
        source, command = "uv", f"uv tool upgrade {PACKAGE_NAME}"
    elif os.path.normcase(os.path.realpath(prefix)) != os.path.normcase(os.path.realpath(base_prefix)):
        source, command = "venv", f"{quoted_python} -m pip install --upgrade {PACKAGE_NAME}"
    elif user_site and module_path.startswith(user_site.replace("\\", "/").lower().rstrip("/") + "/"):
        source, command = "user", f"{quoted_python} -m pip install --user --upgrade {PACKAGE_NAME}"
    else:
        source, command = "system", None
    return {
        "source": source,
        "python": executable,
        "upgrade_command": command,
        "auto_upgrade": source in ("pipx", "uv", "user"),
    }


def _help_payload():
    return {
        "schema_version": SCHEMA_VERSION,
        "success": True,
        "cli_version": __version__,
        "install": _install_info(),
        "commands": {
            "convert": "Convert one .docx/.xlsx/.pptx/.pdf file to Markdown, or one .md file to DOCX.",
            "batch": "Convert supported files in a directory.",
            "setup-node": "Install Node.js dependencies required for Markdown to DOCX conversion.",
        },
        "supported_extensions": SUPPORTED_EXTENSIONS,
        "options": {
            "content": {"choices": ["none", "preview", "full"], "default": "full"},
            "preview_chars": {"default": 4000},
            "strict": "Reject conversions with content-loss or fallback warnings.",
            "batch": {"jsonl": "One result per line followed by a summary.",
                      "manifest": "Exclusive JSONL manifest path, or auto; flushed after each result.",
                      "max_results": "Cap response entries; requires --manifest. Manifest retains all entries."},
        },
    }


def _build_parser():
    # add_help=False: -h would print human text, breaking the JSON-only stdout contract.
    # Use --help-json for machine-readable help instead.
    parser = _JsonArgumentParser(prog="bdc", add_help=False)
    parser.add_argument("--help-json", action="store_true")
    parser.add_argument("--version", action="store_true")
    subparsers = parser.add_subparsers(dest="command")

    # Subparsers inherit _JsonArgumentParser because type(parser) is _JsonArgumentParser
    convert_parser = subparsers.add_parser("convert", add_help=False)
    convert_parser.add_argument("file")
    convert_parser.add_argument("--output-dir")
    convert_parser.add_argument("--extract-images", choices=["true", "false"], default="false")
    convert_parser.add_argument("--mermaid-scale", type=float, default=4.0)

    batch_parser = subparsers.add_parser("batch", add_help=False)
    batch_parser.add_argument("directory")
    batch_parser.add_argument("--output-dir")
    batch_parser.add_argument("--recursive", choices=["true", "false"], default="true")
    batch_parser.add_argument("--extract-images", choices=["true", "false"], default="false")
    batch_parser.add_argument("--mermaid-scale", type=float, default=4.0)
    for command_parser in (convert_parser, batch_parser):
        command_parser.add_argument('--content', choices=['none', 'preview', 'full'], default='full')
        command_parser.add_argument('--preview-chars', type=_nonnegative_int, default=4000)
        command_parser.add_argument('--strict', action='store_true')
    batch_parser.add_argument('--manifest', metavar='PATH_OR_AUTO')
    batch_parser.add_argument('--max-results', type=_nonnegative_int)
    batch_parser.add_argument('--jsonl', action='store_true')

    setup_node_parser = subparsers.add_parser("setup-node", add_help=False)
    setup_node_parser.add_argument("--allow-scripts", action="store_true")
    setup_node_parser.add_argument("--install-browser", action="store_true")
    return parser


def _nonnegative_int(value):
    number = int(value)
    if number < 0:
        raise argparse.ArgumentTypeError('must be nonnegative')
    return number


def _run_batch(namespace):
    if namespace.max_results is not None and not namespace.manifest:
        return _emit(_usage_error('--max-results requires --manifest to retain all outcomes'), 1)
    if namespace.jsonl and namespace.max_results is not None:
        return _emit(_usage_error('--jsonl cannot be combined with --max-results'), 1)
    output_dir = os.path.realpath(os.path.expanduser(namespace.output_dir)) if namespace.output_dir else None
    manifest, manifest_path = None, None
    try:
        if namespace.manifest == 'auto':
            directory = os.path.realpath(os.path.expanduser(namespace.directory))
            if not os.path.exists(directory):
                return _emit({**_usage_error(f'目录不存在: {directory}'), 'error_code': 'FILE_NOT_FOUND'}, 1)
            if not os.path.isdir(directory):
                return _emit({**_usage_error('Batch input must be an existing directory'), 'error_code': 'NOT_A_DIRECTORY'}, 1)
            destination = output_dir or os.path.join(directory, 'Markdown')
            os.makedirs(destination, exist_ok=True)
            fd, manifest_path = tempfile.mkstemp(prefix='bdc-manifest-', suffix='.jsonl', dir=destination)
            manifest = os.fdopen(fd, 'w', encoding='utf-8')
        elif namespace.manifest:
            manifest_path = os.path.realpath(os.path.expanduser(namespace.manifest))
            # Never overwrite a document, input, or previous manifest.
            manifest = open(manifest_path, 'x', encoding='utf-8')

        def record(value):
            line = json.dumps(value, ensure_ascii=False)
            if manifest:
                manifest.write(line + '\n')
                manifest.flush()
            if namespace.jsonl:
                print(line, flush=True)

        results, total, succeeded = [], 0, 0
        for item in iter_batch_convert(namespace.directory, recursive=namespace.recursive == 'true',
                                       extract_images=namespace.extract_images == 'true', output_dir=output_dir,
                                       mermaid_scale=namespace.mermaid_scale, strict=namespace.strict):
            result = _normalize_single_result(item['file'], item['result'], namespace.content, namespace.preview_chars)
            entry = {'input_path': result['input_path'], 'result': result}
            total += 1
            succeeded += int(result['success'])
            record({'type': 'result', **entry})
            if not namespace.jsonl and (namespace.max_results is None or len(results) < namespace.max_results):
                results.append(entry)
        summary = {'schema_version': SCHEMA_VERSION, 'success': total == succeeded,
                   'total': total, 'succeeded': succeeded, 'failed': total - succeeded}
        if manifest_path:
            summary['manifest_path'] = manifest_path
        record({'type': 'summary', **summary})
        if namespace.jsonl:
            return 0 if summary['success'] else 1
        summary['results'] = results
        if total > len(results):
            summary['omitted'] = total - len(results)
        return _emit(summary, 0 if summary['success'] else 1)
    except OSError as exc:
        payload = {'schema_version': SCHEMA_VERSION, 'success': False,
                   'error_code': 'BATCH_IO_ERROR', 'error': str(exc), 'retryable': False}
        if manifest and manifest_path:
            payload['manifest_path'] = manifest_path
        if namespace.jsonl:
            print(json.dumps({'type': 'error', **payload}, ensure_ascii=False), flush=True)
            return 1
        return _emit(payload, 1)
    finally:
        if manifest:
            manifest.close()


def main(argv=None):
    args = list(sys.argv[1:] if argv is None else argv)
    if not args:
        return _emit(_usage_error("缺少命令。可用命令: convert, batch, setup-node"), 1)

    parser = _build_parser()
    namespace = parser.parse_args(args)

    if namespace.help_json:
        return _emit(_help_payload(), 0)

    if namespace.version:
        # A single line, like other CLIs, for quick version checks.
        print(__version__)
        return 0

    if namespace.command == "convert":
        output_dir = os.path.realpath(os.path.expanduser(namespace.output_dir)) if namespace.output_dir else None
        result = convert_document(
            namespace.file,
            extract_images=namespace.extract_images == "true",
            output_dir=output_dir,
            mermaid_scale=namespace.mermaid_scale,
            strict=namespace.strict,
        )
        payload = _normalize_single_result(namespace.file, result, namespace.content, namespace.preview_chars)
        return _emit(payload, 0 if payload["success"] else 1)

    if namespace.command == "batch":
        return _run_batch(namespace)

    if namespace.command == "setup-node":
        payload = setup_node_dependencies(
            allow_scripts=namespace.allow_scripts,
            install_browser=namespace.install_browser,
        )
        payload = {"schema_version": SCHEMA_VERSION, **payload}
        if not payload["success"]:
            payload.setdefault("retryable", False)
            if payload.get('error_code') in SUGGESTIONS:
                payload.setdefault('suggestion', SUGGESTIONS[payload['error_code']])
        return _emit(payload, 0 if payload["success"] else 1)

    return _emit(_usage_error("缺少命令。可用命令: convert, batch, setup-node"), 1)


if __name__ == "__main__":
    sys.exit(main())
