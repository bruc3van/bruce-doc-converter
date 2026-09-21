# Bruce Doc Converter

> Bi-directional document conversion for Claude Code / OpenClaw

[![Agent Skill](https://img.shields.io/badge/Agent_Skill-Skill-purple.svg)](https://github.com/anthropics/claude-code)
[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)

**Bruce Doc Converter** is an agent-facing document converter CLI that gives **Claude Code / OpenClaw** bi-directional document conversion capabilities:

- **Office/PDF → Markdown**: Convert Word, Excel, PowerPoint, and PDF into AI-friendly Markdown
- **Markdown → Word**: Export Markdown as a professionally formatted Word document with automatic Mermaid diagram rendering

## Installation

First, check if `bdc` is already available:

```bash
command -v bdc        # macOS / Linux
where bdc             # Windows
```

If not installed, try the following in order — stop at the first that succeeds:

```bash
# 1. pipx (preferred — isolated, bdc lands in PATH)
pipx install bruce-doc-converter

# 2. uv (if available — fast, isolated, bdc lands in PATH)
uv tool install bruce-doc-converter

# 3. pip --user (most universally available, bdc lands in PATH)
pip3 install --user bruce-doc-converter  # macOS/Linux
pip install --user bruce-doc-converter   # Windows
# or universally: python3 -m pip install --user bruce-doc-converter (use `python` on Windows)

# 4. venv fallback (works everywhere, but bdc will NOT be in PATH)
python3 -m venv .venv
.venv/bin/pip install bruce-doc-converter
# Windows: .venv\Scripts\pip install bruce-doc-converter
```

> **venv note:** If you used the venv fallback, replace every `bdc` command below with `.venv/bin/bdc` (macOS/Linux) or `.venv\Scripts\bdc` (Windows).

> **Windows note:** Use `python` instead of `python3` if the former is not recognized.

## Agent CLI Usage

```bash
bdc convert /path/to/document.docx
bdc convert /path/to/notes.md
bdc convert /path/to/notes.md --mermaid-scale 4
bdc batch /path/to/documents
```

The CLI writes JSON to stdout by default. Progress logs go to stderr.

Markdown to Word requires Node.js dependencies. Initialize them explicitly before first use:

```bash
bdc setup-node
```

By default this runs `npm ci --ignore-scripts` against the locked dependency set, avoiding third-party npm lifecycle scripts; it does not download a browser by default. When Markdown contains Mermaid diagrams, conversion automatically detects and uses local Chrome / Edge / Chromium in headless mode with a temporary browser profile, avoiding visible windows, the user's real profile, and first-run/default-browser prompts. Mermaid PNG rendering defaults to scale `4`; override it with `--mermaid-scale`:

```bash
bdc convert /path/to/notes.md --mermaid-scale 5
bdc batch /path/to/documents --mermaid-scale 5
```

If the target machine has no usable local browser and you want Puppeteer to download its dedicated `chrome-headless-shell`, run this explicitly:

```bash
bdc setup-node --install-browser
```

If your environment really needs npm lifecycle scripts, combine both flags:

```bash
bdc setup-node --allow-scripts --install-browser
```

`bdc setup-node` is idempotent: if the shared dependency directory already matches the installed package, it skips reinstalling Node dependencies. Recoverable failures include `retryable` and `next_command` JSON fields; agents should prefer those machine-readable fields for remediation.

Get help:

```bash
bdc --help-json
```

### Success output (single file)

```json
{
  "schema_version": "1.0",
  "success": true,
  "input_path": "/absolute/input.docx",
  "input_format": "docx",
  "output_format": "markdown",
  "output_path": "/absolute/Markdown/input.md",
  "markdown_content": "# content...",
  "extracted_images": [],
  "warnings": []
}
```

### Failure output

```json
{
  "schema_version": "1.0",
  "success": false,
  "input_path": "/absolute/input.doc",
  "input_format": "doc",
  "error_code": "UNSUPPORTED_FORMAT",
  "error": "Unsupported format: .doc. Supported: .docx, .xlsx, .pptx, .pdf, .md",
  "suggestion": "Convert the file to .docx/.xlsx/.pptx first."
}
```

### Batch output

For batch conversion, `success` means every file converted successfully. If only some files fail, `success` is `false`, while `succeeded`, `failed`, and `results` still report per-file details.

```json
{
  "schema_version": "1.0",
  "success": true,
  "total": 1,
  "succeeded": 1,
  "failed": 0,
  "results": [
    {
      "input_path": "/absolute/input.docx",
      "result": {
        "schema_version": "1.0",
        "success": true,
        "input_path": "/absolute/input.docx",
        "input_format": "docx",
        "output_format": "markdown",
        "output_path": "/absolute/Markdown/input.md",
        "markdown_content": "# content...",
        "extracted_images": [],
        "warnings": []
      }
    }
  ]
}
```

## Features

- **Heading detection**: Automatically recognizes Word heading levels (Heading 1–6) and Chinese heading styles
- **Format preservation**: Retains bold, italic, and other inline formatting
- **Table conversion**: Converts tables to clean Markdown
- **List support**: Ordered, unordered, and nested lists
- **Mermaid diagrams**: Renders Mermaid code blocks via `mmdc` and embeds them as PNG images in Word
- **Image extraction**: Extracts embedded images from Word/Excel/PowerPoint; PDF image extraction is not supported

## Supported Formats

| Format | Input | Output | Quality |
| --- | --- | --- | --- |
| Word (.docx) | ✅ | ✅ | Excellent |
| Excel (.xlsx) | ✅ | ❌ | Excellent |
| PowerPoint (.pptx) | ✅ | ❌ | Good |
| PDF (.pdf) | ✅ | ❌ | Depends on file type |
| Markdown (.md) | ✅ | ✅ | Excellent |

> **Note**: Legacy formats (.doc, .xls, .ppt) are not supported. Convert them to modern formats first.

## Requirements

- **Python 3.8+** (required)
- **Node.js >=22.0** (optional, only needed for Markdown → Word; checked before installation/conversion)

The standalone CLI minimum follows the locked dependencies (`chevrotain` requires
Node.js >=22.0.0) and has been tested on Node 22.0.0. The optional DSH plugin keeps
its independent `^22.19 || >=24` requirement; it does not raise the CLI requirement
for other agents using the Skill.

## FAQ

### Installation troubleshooting

| Error | Cause | Fix |
| --- | --- | --- |
| `SOCKS support` / proxy connection error | `all_proxy` or `http_proxy` env vars set | Run `unset all_proxy http_proxy https_proxy` (macOS/Linux) or `set all_proxy=` (Windows CMD), then retry |
| `command not found: pipx` | pipx not installed | Try `uv tool install` or `pip install --user` instead |
| `externally-managed-environment` | Python 3.11+ system Python forbids global pip | Use `pipx`, `uv tool install`, or the venv fallback |
| Permission denied | No write access to install location | Add `--user` flag, or use venv fallback |
| `bdc: command not found` after venv install | venv bin not in PATH | Use full path: `.venv/bin/bdc` (macOS/Linux) or `.venv\Scripts\bdc` (Windows) |

### File too large?

Current limit is 100 MB. Consider splitting the file or reducing content.

### Markdown → Word failing?

Node.js is required, and the Node.js dependencies must be installed explicitly:

```bash
bdc setup-node
```

When Markdown contains Mermaid, conversion prefers local Chrome / Edge / Chromium and launches it with a temporary headless profile. Set `BRUCE_DOC_CONVERTER_CHROME_PATH` to force a browser path; if no local browser is available, run `bdc setup-node --install-browser` to download Puppeteer's dedicated browser.

On Linux, Chromium is not launched with `--no-sandbox` by default. If you understand the risk and your environment requires it, set `BRUCE_DOC_CONVERTER_ALLOW_CHROMIUM_NO_SANDBOX=1` before conversion.

### PDF returns no content?

Scanned PDFs require OCR first. Protected PDFs must be unlocked before conversion.

## Best Practices

1. **Use modern Office formats** (.docx, .xlsx, .pptx)
2. **Prefer text-based PDFs** over scanned images; run OCR first if needed
3. **Keep files under 50 MB** for best performance

## DeepSeek Harness plugin

For [DeepSeek Harness](https://github.com/deepseek-ai/deepseek-harness) (`dsh`), install [`dsh-plugin/`](dsh-plugin/README.en.md) to expose `bdc` as three model tools (`doc_convert`, `doc_batch`, `doc_setup`):

```sh
dsh plugin --profile web add bruce-doc-converter-dsh
```

The plugin builds the argv and runs it through `ctx.shell`, so sandboxing, approval, timeouts, and output caps apply; the CLI's JSON becomes a typed tool result, and install paths move into validated `cordis.yml` fields instead of prompt text. See the [plugin README](dsh-plugin/README.en.md).

## Project Structure

```
bruce-doc-converter/
├── bruce-doc-converter-skill/
│   └── SKILL.md                  # Agent Skill definition
├── dsh-plugin/                   # DeepSeek Harness plugin (npm: bruce-doc-converter-dsh)
│   ├── src/                      # Tools, runner, envelope validation, embedded skill
│   ├── tests/                    # Unit tests plus real-CLI end-to-end tests
│   └── cordis.patch.yml          # Bundle patch layer
├── pyproject.toml                # Python package metadata
├── requirements.txt              # Local development dependencies
├── bruce_doc_converter/
│   ├── __init__.py
│   ├── cli.py                    # bdc CLI entry point
│   ├── converter.py              # Core conversion logic
│   └── md_to_docx/              # Node.js module for Markdown → Word
└── tests/
    ├── test_cli.py
    ├── test_convert_document.py
    └── md_to_docx.test.js
```

## Integrity and bounded batch output

```bash
bdc convert budget.xlsx --strict
bdc convert report.docx --content preview --preview-chars 4000
bdc batch ./documents --content none --manifest auto --max-results 200
bdc batch ./documents --content none --jsonl --manifest ./results.jsonl
```

`--content full` remains the default. `preview` returns a prefix and `none` omits
the body; the saved document stays complete. `markdown_chars` counts Unicode
code points in the full text and `markdown_truncated` identifies previews.
`--max-results` caps response entries and requires a manifest; `omitted` reports
the omitted count. It cannot be combined with `--jsonl`.

The JSONL manifest contains every file outcome (`type: result`) and a final
`type: summary` record on normal completion. Body inclusion follows `--content`.
Records flush after each file so completed work remains inspectable after an
interruption. Absence of the summary means the batch did not finish normally.
`--manifest auto` creates a unique file under the output directory or the input
root's `Markdown/`. Explicit manifest paths require an existing parent directory
and never overwrite an existing file.

Output names are exclusively reserved and files are written atomically. Images
use a separate directory per conversion so repeated/concurrent runs preserve old
results. Uncached Excel formulas remain visible and produce
`FORMULA_CACHE_MISSING` diagnostics with worksheet/cell coordinates; formulas
are not calculated. PDF diagnostics report each page as `extracted`, `fallback`,
`failed`, or `empty`. Empty and scanned pages cannot be reliably distinguished
solely by lack of text, so both warn. Strict mode rejects content-loss/fallback
warnings, including missing formula caches, empty pages and failed image/Mermaid
embedding, without writing that conversion's output.

## Development checks

Format extractors live in `bruce_doc_converter/formats/`; `converter.py` retains
the existing Python entry points and orchestration. `output.py` owns exclusive
reservation and atomic writes. Markdown uses markdown-it CommonMark parsing with
tables/strikethrough, literal raw HTML and Mermaid fence rendering.

```bash
python -m pip install -e . build
npm ci --ignore-scripts --prefix bruce_doc_converter/md_to_docx
python -m unittest discover -s tests -p "test_*.py"
node --test tests/md_to_docx.test.js
pnpm --dir dsh-plugin install --frozen-lockfile --ignore-scripts
pnpm --dir dsh-plugin typecheck
pnpm --dir dsh-plugin test
python -m build
python scripts/check_wheel.py
```

`tests/fixtures/` is a reproducible synthetic corpus. Tests inspect saved DOCX
XML, cached/uncached formulas, concurrent output isolation and JSONL contracts.
CLI CI covers Windows/Linux/macOS with Python 3.8/3.12 (macOS: 3.12 only) and Node
22.0, plus Node 24 on Linux. The DSH plugin uses a separate Node 22.19 job. Missing
real CLI integration prerequisites fail CI instead of skipping.

## License

MIT License
