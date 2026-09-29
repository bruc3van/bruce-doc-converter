# Diagnostics and errors

Every result is JSON on stdout. Successful conversions may carry `warnings` (messages) and `diagnostics` (structured items):

```json
{"code": "MATH_NOT_CONVERTED", "severity": "warning", "message": "...", "line": 12}
```

`severity: warning` means content is missing or degraded and fails `--strict`; `severity: info` does not. Locations are `line` (Markdown source), `page` (PDF) or `sheet`/`cell` (Excel). Treat unknown codes by their severity and message; never report a result with warnings as clean.

A strict failure returns `error_code: CONTENT_INCOMPLETE` with the same `warnings` and `diagnostics`, and writes no output file. Fix the source using those items and rerun the same strict command.

## Markdown to Word diagnostics

| Code | Severity | What to do |
| --- | --- | --- |
| `IMAGE_UNAVAILABLE` | warning | Check the image exists, is PNG/JPEG/GIF/BMP/SVG, and sits inside the Markdown file's directory. Remote URLs, absolute paths and `file:` links are never embedded; download or copy the image next to the Markdown and use a relative path. |
| `MATH_NOT_CONVERTED` | warning | Check delimiters and syntax at the reported line. Rewrite unsupported constructs (custom macros, `\label`, `\ref`, `\tag`, colors, manual spacing, font commands such as `\textbf` or `\textrm`; plain `\text{}` works) while keeping the mathematical meaning. The source is kept as text, which is not an editable equation. |
| `MERMAID_NOT_RENDERED` | warning | Check the diagram syntax at the reported line. If the browser failed to start, see the Markdown to Word runtime section of [install.md](install.md). The source is kept as a code block with a notice. |
| `LINK_UNAVAILABLE` | warning | A `#fragment` link has no matching heading. Match the heading text, or remove the link. |
| `FOOTNOTE_UNDEFINED` | warning | Add the missing `[^label]:` definition; the reference is kept as literal text. |
| `FOOTNOTE_DUPLICATE` / `FOOTNOTE_NESTED` | warning | Merge duplicate definitions or remove footnote references inside footnotes. Do not delete footnote text just to pass strict mode. |
| `FOOTNOTE_TABLE_FLATTENED` | warning | A table inside a footnote was split into paragraphs; move the table into the body if it matters. |
| `FOOTNOTE_UNUSED` | info | A definition has no reference and was kept in the body. Check whether a reference is missing. |
| `LIST_DEPTH_REDUCED` | info | Lists deeper than five levels were laid out at level five; reduce nesting. |
| `DIAGNOSTICS_TRUNCATED` | as the omitted items | More than 100 diagnostics. Fix the listed ones and convert again to see the rest; do not claim all problems were listed. |

## Office/PDF to Markdown diagnostics

| Code | Severity | What to do |
| --- | --- | --- |
| `FORMULA_CACHE_MISSING` | warning | An Excel formula has no cached value (the workbook was never recalculated and saved). The formula text is kept; ask the user to open and save the file in Excel if values are needed. Formulas are never calculated. |
| `PDF_PAGE_FALLBACK` | warning | The page used simpler text extraction; layout or table structure may be lost. |
| `PDF_PAGE_EMPTY` | warning | No text on the page: blank, scanned or image-only. OCR the PDF if content is expected. |
| `PDF_PAGE_FAILED` | warning | The page could not be extracted. |
| `PDF_PAGE_EXTRACTED` | info | Per-page success record. |
| `IMAGE_EXTRACTION_FAILED` | warning | An embedded image could not be saved with `--extract-images true`. |

## Error codes

| Code | What to do |
| --- | --- |
| `USAGE_ERROR` | Fix the command arguments; run `bdc --help-json` for the options. |
| `FILE_NOT_FOUND` / `NOT_A_FILE` / `NOT_A_DIRECTORY` | Check the path relative to the current working directory and quote paths with spaces or CJK characters. |
| `UNSUPPORTED_FORMAT` | Only `.docx`, `.xlsx`, `.pptx`, `.pdf` and `.md` are supported. Legacy `.doc`/`.xls`/`.ppt` must be saved in the new formats first. |
| `FILE_TOO_LARGE` / `OUT_OF_MEMORY` | Split the document or reduce embedded content; do not retry unchanged. |
| `EMPTY_PDF_CONTENT` | No text or tables in the whole PDF; it is likely scanned or protected. OCR or unlock it first. |
| `CONTENT_INCOMPLETE` | Strict mode rejected degraded content; see the sections above. |
| `DEPENDENCY_INSTALL_REQUIRED` | Run `next_command` (`bdc setup-node`) and retry. This is expected after installing or upgrading. |
| `DEPENDENCY_INSTALL_FAILED` | Read the npm error; handle network, proxy or permission causes as in [install.md](install.md). |
| `NODE_NOT_FOUND` / `NODE_VERSION_UNSUPPORTED` / `NODE_VERSION_CHECK_FAILED` | Markdown to Word needs Node.js >=22.0. Tell the user what to install; rerunning `setup-node` cannot fix the runtime. |
| `CONVERSION_TIMEOUT` | Markdown to Word exceeded 2 minutes, usually from many or large Mermaid diagrams. Split the document or simplify diagrams. |
| `NODE_CONVERSION_FAILED` | Keep the real error message. It includes generated files that fail the DOCX integrity check and documents over 1000 formulas or footnotes (split those). Report it rather than inventing an output path. |
| `PERMISSION_DENIED` / `OS_ERROR` / `BATCH_IO_ERROR` | Check read access to the input and write access to the output directory, and disk space. |
| `CONVERSION_ERROR` | Unexpected failure; report the message. The input file may be damaged. |
