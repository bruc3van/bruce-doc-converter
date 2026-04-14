---
name: bruce-doc-converter
description: 双向文档转换工具，将 Word (.docx)、Excel (.xlsx)、PowerPoint (.pptx) 和 PDF (.pdf) 转换为 AI 友好的 Markdown 格式，或将 Markdown (.md) 转换为 Word (.docx) 格式。当用户请求以下操作时使用：(1) 明确请求文档转换，包括任何包含"转换"、"转为"、"转成"、"convert"、"导出"、"export"等词汇的请求（例如："转换文档"、"把这个文件转为docx"、"convert to markdown"、"导出为Word"）；(2) 需要 AI 理解文档内容（"帮我分析这个 Word 文件"、"读取这个 PDF"、"总结这个 Excel"）；(3) 上传文档文件并询问内容（"这是什么"、"帮我看看"）；(4) 任何涉及 .docx、.xlsx、.pptx、.pdf、.md 文件格式转换的请求。
---
# Bruce Doc Converter

双向文档转换工具，将 Word (.docx)、Excel (.xlsx)、PowerPoint (.pptx) 和 PDF (.pdf) 转换为 AI 友好的 Markdown 格式，或将 Markdown (.md) 转换为 Word (.docx) 格式。

## Quick Reference

| 操作                   | Linux/macOS                       | Windows                                               | 输出位置             |
| ---------------------- | --------------------------------- | ----------------------------------------------------- | -------------------- |
| Office/PDF → Markdown | `bash convert.sh <file>`        | `powershell.exe -Command "Set-Location '<skill-dir>'; .\convert.ps1 '<file>'"` | 同目录 `Markdown/` |
| Markdown → Word       | `bash convert.sh <file.md>`     | 同上（传入 .md 文件）                                  | 同目录 `Word/`     |
| 批量转换               | `bash convert.sh --batch <dir>` | —                                                    | 同上                 |

> **Windows 说明**：路径含空格时用单引号包裹；用 `Set-Location` 而非 `cd`；`<skill-dir>` 替换为实际 skill 目录路径。

## 工作流程

```
用户请求转换 → 直接运行 bash convert.sh → 解析 JSON 输出 → 处理结果
```

**关键原则：不要预先检查任何依赖**（Python 库、Node.js 等）。直接执行转换命令，只在转换失败（`success: false`）时才根据错误信息处理。

## 命令示例

```bash
# 单文件转换（依赖自动安装）
bash convert.sh /path/to/document.docx

# 自定义输出目录
bash convert.sh /path/to/file.pdf true /custom/output

# 批量转换
bash convert.sh --batch /path/to/documents
```

## 解析输出

脚本返回 JSON，关键字段：

```json
{
  "success": true,
  "output_path": "/path/to/output.md",
  "markdown_content": "# 转换后的内容..."
}
```

- `success`: 转换是否成功
- `output_path`: 输出文件路径
- `markdown_content`: Markdown 内容（方便直接分析）
- `error`: 错误信息（失败时）

## 错误处理

**仅在转换失败时（返回 `success: false`）才处理错误**：

| 错误类型                   | 处理方法                                                              |
| -------------------------- | --------------------------------------------------------------------- |
| Python 依赖缺失            | 脚本会自动安装；如失败则运行 `pip install --user xxx`               |
| `未找到 Node.js`         | 仅 MD→DOCX 失败且报此错时，提示安装 Node.js                          |
| `Node.js 依赖未安装`     | 脚本自动安装到共享目录；失败时在 `scripts/md_to_docx` 运行 `npm install` |
| `文件不存在`             | 提示用户验证文件路径                                                  |
| `不支持的文件格式: .doc` | 提示用户先转换为 .docx                                                |
| `文件过大`               | 提示超过 100MB 限制                                                   |

## 支持的格式

详见 [references/supported-formats.md](references/supported-formats.md)。简要汇总：

| 格式  | 转换方向 | 质量            |
| ----- | -------- | --------------- |
| .docx | ↔       | 优秀            |
| .xlsx | →       | 优秀            |
| .pptx | →       | 良好            |
| .pdf  | →       | 取决于 PDF 类型 |
| .md   | ↔       | 优秀            |

依赖自动安装：Python 依赖安装到用户目录，Node.js 依赖安装到用户级共享目录（可用 `BRUCE_DOC_CONVERTER_NODE_HOME` 指定）。旧版 .doc/.xls/.ppt 需先转换为对应新格式。
