---
name: bruce-doc-converter
description: 双向文档转换：将 Word (.docx)、Excel (.xlsx)、PowerPoint (.pptx) 和 PDF (.pdf) 转为 AI 友好的 Markdown 以便读取和分析；或将 Markdown (.md) 导出为 Word (.docx)，支持中文排版、Mermaid 图表、可编辑数学公式和原生脚注。当用户请求转换、读取、分析 Office/PDF 文件，上传这些格式并询问内容，或要求交付 DOCX 时使用；不用于修改已有 Word 文档内容、生成 PDF 或套用自定义 Word 模板。
---

# Bruce Doc Converter

Run the local `bdc` CLI through a command execution tool. It prints JSON on stdout; progress logs may appear on stderr.

## Prepare the CLI (once per session)

Check installation and updates the first time you convert in a session; later conversions reuse the same entry point without checking again.

1. Run `bdc --help-json`. If the command is missing, look for an existing virtual environment (`.venv/bin/bdc`, or `.venv\Scripts\bdc` on Windows) before concluding it is not installed; otherwise install it following [Install and update](references/install.md).
2. Read `cli_version` and `install` (`source`, `auto_upgrade`, `upgrade_command`). CLIs before 0.2.2 have no `install` field; identify the installation as described in the reference.
3. When the network is available, find the latest stable release: `python -m pip index versions bruce-doc-converter` (first line). Compare semantic versions, not strings; never downgrade or choose a pre-release.
4. If a newer release exists:
   - `auto_upgrade: true` (pipx, uv tool, `pip --user`): run `upgrade_command`, re-run `bdc --help-json` to confirm, and tell the user the old and new versions.
   - `venv`: this may be a project dependency. Tell the user an update exists and run `upgrade_command` only after they agree.
   - `system` or `source` (a development checkout): tell the user; do not modify it.
   - A version the user or host pinned stays as is; only mention the update.
5. Offline, a failed query, or a host policy that forbids installs: keep using the installed version and say that updates were not checked. Follow the host's install approval and audit rules; do not bypass them to get a newer version.

After an upgrade, the next Markdown to Word conversion may return `DEPENDENCY_INSTALL_REQUIRED`; run its `next_command` (`bdc setup-node`) and retry.

## Convert

**Office/PDF to Markdown**

```bash
bdc convert "<file>"
bdc convert "<file>" --content preview      # large files: return a preview, full text is still written
bdc convert "<file>" --extract-images true  # also save embedded images
bdc batch "<directory>"
```

Read `markdown_content` directly for analysis. Output goes to a `Markdown/` directory beside the source unless `--output-dir` is given.

**Markdown to Word**

1. Use the user's existing `.md` file. When writing a new report, save it as UTF-8 Markdown first so the source stays editable.
2. Markdown to Word needs Node.js >=22.0 and a one-time `bdc setup-node`. Office/PDF conversion does not need Node.js.
3. Export deliverables in strict mode, quoting paths with spaces or CJK characters:

   ```bash
   bdc convert "docs/报告.md" --strict
   ```

   The file goes to a `Word/` directory beside the source unless `--output-dir` is given. Existing files are never overwritten; a numbered name such as `报告.2.docx` is used instead, so read the real path from `output_path`.

## Check the result

1. Check the exit code and parse stdout. On success, confirm `output_path` exists and is not empty, then report that exact path. Files live on the machine that ran the command; attach them only when the host provides attachments, and never invent download links.
2. Read `warnings` and `diagnostics`. Each diagnostic has `code`, `severity`, `message`, and a location when known (`line` for Markdown, `page` for PDF, `sheet`/`cell` for Excel). `warning` items mean missing or degraded content and fail `--strict`; `info` items are advisory.
3. On `CONTENT_INCOMPLETE`, nothing was written. Fix the source at the reported lines and rerun the same strict command. Do not drop `--strict` for a final deliverable unless the user accepts degraded output, and then say exactly what is missing. If the same failure repeats with nothing new to fix, stop and explain what material or environment is needed.
4. For other failures, use `error_code`, `retryable`, `next_command` and `suggestion`. See [Diagnostics and errors](references/diagnostics.md) for every code.

A strict success only means no content loss was detected. It does not prove the facts are right or the layout was checked; if you did not open or render the document, say that pagination and visual layout are unchecked.

## Markdown to Word: what converts

- Fixed Chinese report layout on A4: SimSun body with a two-character first-line indent, SimHei headings, 1.5 line spacing. No custom templates, table of contents, headers, footers or page numbers.
- Headings, paragraphs, lists (up to five levels), tables, quotes, code blocks and links stay editable. Tables keep Markdown column alignment and get content-based column widths. `[text](#heading)` links jump to the matching heading.
- LaTeX `$...$`, `\(...\)`, `$$...$$` and `\[...\]` become editable Word equations. Custom macros, `\label`/`\ref`/`\tag`, colors and manual spacing are not supported and keep their source with `MATH_NOT_CONVERTED`. Escape currency as `\$` when a line could be read as a formula.
- `[^label]` footnotes become Word footnotes; repeated references reuse the first number. Inline `^[...]` footnotes stay literal.
- Images: PNG, JPEG, GIF, BMP and SVG (SVG needs Word 2016 or later; older readers show a blank placeholder), from paths relative to the Markdown file's directory (never outside it) or `data:` URIs. Remote URLs, absolute paths and `file:` links are not embedded.
- Mermaid blocks render to PNG through the local Chrome, Edge or Chromium (`--mermaid-scale`, default 4). Unsupported syntax keeps the source with a visible notice and `MERMAID_NOT_RENDERED`.
- Raw HTML and task-list checkboxes stay as plain text. Ordered lists always start at 1; a start number such as `3.` is not preserved.
