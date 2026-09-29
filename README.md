# Bruce Doc Converter

> 为 Claude Code / OpenClaw 添加双向文档转换能力

[![Agent Skill](https://img.shields.io/badge/Agent_Skill-Skill-purple.svg)](https://github.com/anthropics/claude-code)
[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)

**Bruce Doc Converter** 是一个面向 Agent 的文档转换 CLI，为 **Claude Code / OpenClaw** 添加**双向文档转换**能力：

- **Office/PDF → Markdown**：将 Word、Excel、PowerPoint、PDF 转换为 AI 友好的 Markdown 格式
- **Markdown → Word**：将 Markdown 导出为排版精美的 Word 文档，自动渲染 Mermaid 图表

## 安装

首先检查 `bdc` 是否已安装：

```bash
command -v bdc        # macOS / Linux
where bdc             # Windows
```

若未安装，依次尝试以下方式（成功即止）：

```bash
# 1. pipx（首选，独立环境，bdc 直接可用）
pipx install bruce-doc-converter

# 2. uv（快速、独立，bdc 直接可用）
uv tool install bruce-doc-converter

# 3. pip --user（最通用，bdc 直接可用）
pip3 install --user bruce-doc-converter   # macOS/Linux
pip install --user bruce-doc-converter    # Windows
# 或通用写法：python3 -m pip install --user bruce-doc-converter（Windows 用 `python`）

# 4. venv 兜底（处处可用，但 bdc 不在 PATH 中）
python3 -m venv .venv
.venv/bin/pip install bruce-doc-converter
# Windows: .venv\Scripts\pip install bruce-doc-converter
```

> **venv 提示**：使用 venv 方式安装后，下文所有 `bdc` 命令需替换为 `.venv/bin/bdc`（macOS/Linux）或 `.venv\Scripts\bdc`（Windows）。
>
> **Windows 提示**：若 `python3` 未识别，改用 `python`。

## Agent CLI 用法

```bash
bdc convert /path/to/document.docx
bdc convert /path/to/notes.md
bdc convert /path/to/notes.md --mermaid-scale 4
bdc batch /path/to/documents
```

CLI 默认向 stdout 输出 JSON，stderr 仅用于进度日志。

### 内容完整性与批量结果

```bash
# 有内容缺失或降级警告时返回失败，不生成该文件的转换结果
bdc convert budget.xlsx --strict
# 仅返回前 4000 个字符；完整正文仍保存在输出文件中
bdc convert report.docx --content preview --preview-chars 4000
# 不将正文装入 JSON；完整逐文件明细写入独立 JSONL 清单
bdc batch ./documents --content none --manifest auto --max-results 200
# 每完成一个文件即输出一行 JSON，最后输出汇总
bdc batch ./documents --content none --jsonl --manifest ./results.jsonl
```

`--content` 支持 `full`（默认，兼容原调用）、`preview`、`none`。
`markdown_chars` 是完整正文的 Unicode 字符数；预览被截断时 `markdown_truncated` 为 `true`。
`--max-results` 只限制 stdout 返回的条数，必须同时指定 `--manifest`；`omitted` 表示未返回的条数。
`--jsonl` 与 `--max-results` 不能同时使用。

清单使用 JSONL：逐文件记录的 `type` 为 `result`，正常结束后附加 `summary`。
清单逐条刷新；若进程中断，可读取已经完成的记录，没有最终 `summary` 就表示批次未正常结束。
清单保留所有文件的结果明细，正文是否包含仍遵循 `--content`。
`--manifest auto` 在指定输出目录或输入根目录的 `Markdown/` 中创建唯一清单；指定文件路径时父目录必须存在，已有文件不会被覆盖。

输出文件采用独占命名和原子写入；每次转换的图片放在独立子目录，重复或并发转换不会改写旧结果。
Excel 有缓存值的公式使用缓存值；无缓存时保留公式，返回含工作表和单元格位置的 `FORMULA_CACHE_MISSING` 诊断，工具不会计算公式。
PDF 返回逐页 `diagnostics`，区分 `extracted`、`fallback`、`failed`、`empty`。
空页和扫描页无法仅凭无文本可靠区分，均会告警；`--strict` 拒绝内容缺失或降级警告，包括空页、公式无缓存、图片无法嵌入、Mermaid 渲染失败。
Markdown 转 Word 同样返回 `diagnostics`（`code`、`severity`、`message`，可定位时附 `line` 源文件行号），如 `MATH_NOT_CONVERTED`、`FOOTNOTE_UNDEFINED`、`LINK_UNAVAILABLE`、`IMAGE_UNAVAILABLE`、`MERMAID_NOT_RENDERED`。`severity` 为 `warning` 的诊断同时出现在 `warnings` 中并被 `--strict` 拒绝；`info`（如未被引用的脚注）只作提示。生成的 DOCX 在写出前会校验压缩包与 XML 完整性。

Markdown 转 Word 需要 Node.js 依赖。首次使用前请显式初始化：

```bash
bdc setup-node
```

默认初始化会使用 `npm ci --ignore-scripts` 安装锁定依赖，避免运行第三方 npm 生命周期脚本；不会默认下载浏览器。Markdown 中包含 Mermaid 图表时，转换阶段会自动探测并使用本机 Chrome / Edge / Chromium，并以 headless 模式、临时浏览器 profile 启动，避免打开窗口、使用用户真实 profile 或触发默认浏览器/首次启动检查。Mermaid PNG 默认渲染倍率为 `4`，可通过 `--mermaid-scale` 调整：

```bash
bdc convert /path/to/notes.md --mermaid-scale 5
bdc batch /path/to/documents --mermaid-scale 5
```

如果目标机器没有可用的本地浏览器，并且需要让 Puppeteer 下载专用的 `chrome-headless-shell`，显式运行：

```bash
bdc setup-node --install-browser
```

如果你的环境确实需要运行 npm 生命周期脚本，可同时使用：

```bash
bdc setup-node --allow-scripts --install-browser
```

`bdc setup-node` 是幂等命令：如果共享依赖目录已经和当前发布包匹配，会跳过 Node 依赖重装。可恢复失败会在 JSON 中提供 `retryable` 和 `next_command` 字段，智能体应优先使用这些机器字段决定下一步。

查看帮助与版本：

```bash
bdc --help-json
bdc --version
```

`--help-json` 中的 `install` 字段说明 CLI 的安装方式（`pipx`、`uv`、`user`、`venv`、`system` 或源码 `source`）及对应的 `upgrade_command`；`auto_upgrade` 为 `true` 时，Skill 会引导智能体在发现新版本后自动升级并告知版本变化，项目 venv、系统 Python 和源码目录只提示、不擅自升级。

### 输出示例（单文件成功）

```json
{
  "schema_version": "1.0",
  "success": true,
  "input_path": "/absolute/input.docx",
  "input_format": "docx",
  "output_format": "markdown",
  "output_path": "/absolute/Markdown/input.md",
  "markdown_content": "# 内容...",
  "extracted_images": [],
  "warnings": []
}
```

### 输出示例（失败）

```json
{
  "schema_version": "1.0",
  "success": false,
  "input_path": "/absolute/input.doc",
  "input_format": "doc",
  "error_code": "UNSUPPORTED_FORMAT",
  "error": "不支持的文件格式: .doc。支持的格式: .docx, .xlsx, .pptx, .pdf, .md",
  "suggestion": "请先转换为 .docx/.xlsx/.pptx 后再重试。"
}
```

### 输出示例（批量转换）

批量转换的 `success` 表示是否所有文件都转换成功；部分失败时 `success` 为 `false`，但 `succeeded`、`failed` 和 `results` 会保留每个文件的明细。

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
        "markdown_content": "# 内容...",
        "extracted_images": [],
        "warnings": []
      }
    }
  ]
}
```

## 功能特性

- **标题识别**：自动识别 Word 标题层级（Heading 1-6）及中文标题样式
- **格式保留**：保留粗体、斜体等文本格式
- **表格转换**：智能转换表格为 Markdown 格式
- **列表支持**：有序列表、无序列表及多级嵌套
- **Mermaid 图表**：支持通过 `mmdc` 渲染 Mermaid 代码块，嵌入 Word 为 PNG 图片
- **数学公式**：Markdown 转 Word 时，`$...$`、`\(...\)` 行内公式和 `$$...$$`、`\[...\]` 块公式转换为可编辑的 Word 原生公式；不支持的公式保留源码并告警
- **脚注与文档内跳转**：`[^名称]` 生成 Word 原生脚注；`[文字](#标题)` 生成跳转到对应标题的内部链接
- **中文排版**：中文源文件中的换行不再插入多余空格；表格按内容分配固定列宽并保留 Markdown 对齐方式；图片按自然尺寸显示并限制在版心内
- **图片提取**：Word/Excel/PowerPoint 转 Markdown 时可提取内嵌图片；暂不支持 PDF 图片提取

## 支持的格式

| 格式               | 输入 | 输出 | 质量       |
| ------------------ | ---- | ---- | ---------- |
| Word (.docx)       | ✅   | ✅   | 优秀       |
| Excel (.xlsx)      | ✅   | ❌   | 优秀       |
| PowerPoint (.pptx) | ✅   | ❌   | 良好       |
| PDF (.pdf)         | ✅   | ❌   | 取决于类型 |
| Markdown (.md)     | ✅   | ✅   | 优秀       |

> **注意**：不支持旧版格式（.doc, .xls, .ppt），请先转换为新格式。

## 环境要求

- **Python 3.8+**（必需）
- **Node.js >=22.0**（可选，仅 Markdown → Word 需要；安装及转换时检查版本）

独立 CLI 的最低版本取决于锁定依赖（其中 `chevrotain` 要求 Node.js >=22.0.0），已在 Node 22.0.0 验证。
DSH 插件独立遵循其 `^22.19 || >=24` 环境要求，不影响其他 Agent 通过 Skill 调用 CLI。

## 常见问题

### 安装故障排查

| 错误 | 原因 | 解决方案 |
| --- | --- | --- |
| `SOCKS support` / 代理连接错误 | `all_proxy` 或 `http_proxy` 环境变量已设置 | 运行 `unset all_proxy http_proxy https_proxy`（macOS/Linux）或 `set all_proxy=`（Windows CMD），然后重试 |
| `command not found: pipx` | 未安装 pipx | 改用 `uv tool install` 或 `pip install --user` |
| `externally-managed-environment` | Python 3.11+ 系统 Python 禁止全局 pip 安装 | 使用 `pipx`、`uv tool install` 或 venv 兜底 |
| Permission denied | 无安装目录写权限 | 添加 `--user` 标志，或使用 venv 兜底 |
| 安装 venv 后 `bdc: command not found` | venv bin 未加入 PATH | 使用完整路径：`.venv/bin/bdc`（macOS/Linux）或 `.venv\Scripts\bdc`（Windows） |

### 文件过大怎么办？

当前限制为 100MB，建议分割文件或压缩内容。

### Markdown 转 Word 失败？

需要安装 Node.js，并先显式安装 Node.js 依赖：

```bash
bdc setup-node
```

如果 Markdown 中包含 Mermaid，转换时会优先使用本地 Chrome / Edge / Chromium，并以无窗口、临时 profile 模式启动。可通过 `BRUCE_DOC_CONVERTER_CHROME_PATH` 指定浏览器路径；没有本地浏览器时再运行 `bdc setup-node --install-browser` 下载 Puppeteer 专用浏览器。

Linux 下默认不会为 Chromium 传入 `--no-sandbox`。如果你理解风险且运行环境确实需要，可设置 `BRUCE_DOC_CONVERTER_ALLOW_CHROMIUM_NO_SANDBOX=1` 后再转换。

### PDF 提取不到内容？

扫描型 PDF 需先执行 OCR，或解除 PDF 保护后重试。

## 最佳实践

1. **使用新版 Office 格式**（.docx, .xlsx, .pptx）
2. **PDF 优先使用文本型**，扫描型建议先 OCR
3. **文件大小建议 < 50MB**

## DeepSeek Harness 插件

如果你用 [DeepSeek Harness](https://github.com/deepseek-ai/deepseek-harness)（`dsh`），可以额外安装 [`dsh-plugin/`](dsh-plugin/README.md)，把 `bdc` 变成三个模型工具（`doc_convert` / `doc_batch` / `doc_setup`）：

```sh
dsh plugin --profile web add bruce-doc-converter-dsh
```

工具由插件拼好 argv 后经 `ctx.shell` 执行，因此沙箱、审批、超时和输出上限照常生效；CLI 的 JSON 变成类型化结果，安装方式和路径改由 `cordis.yml` 配置项决定，不再占用提示词。详见 [插件 README](dsh-plugin/README.md)。

## 项目结构

```
bruce-doc-converter/
├── bruce-doc-converter-skill/
│   ├── SKILL.md                  # Agent Skill 定义
│   └── references/               # 安装与更新、诊断与错误码
├── dsh-plugin/                   # DeepSeek Harness 插件（npm: bruce-doc-converter-dsh）
│   ├── src/                      # 工具、runner、信封校验、内置技能
│   ├── tests/                    # 单元测试 + 真实 CLI 端到端测试
│   └── cordis.patch.yml          # bundle 补丁层
├── pyproject.toml                # Python 包元数据
├── requirements.txt              # 本地开发依赖
├── bruce_doc_converter/
│   ├── __init__.py
│   ├── cli.py                    # bdc CLI 入口
│   ├── converter.py              # 转换核心逻辑
│   └── md_to_docx/              # Markdown → Word 的 Node.js 模块
└── tests/
    ├── test_cli.py
    ├── test_convert_document.py
    └── md_to_docx.test.js
```

## 开发验证

转换核心按格式拆分到 `bruce_doc_converter/formats/`；`converter.py` 保留调度与既有 Python 入口，`output.py` 管理输出占位和原子写入。
Markdown 使用 markdown-it 的 CommonMark 解析与表格/删除线扩展，原始 HTML 保持为文本，Mermaid 作为 fence token 渲染。
公式经 Temml 转为 MathML 后映射为 Word OMML，脚注使用 markdown-it-footnote；这部分实现及表格列宽、嵌套布局参考自同源项目 [bruce-md2word](https://github.com/bruc3van/bruce-md2word)。

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

`tests/fixtures/` 包含可重建的合成 Office/PDF 文档和 Markdown 样例。
测试检查真实 DOCX XML、公式缓存、图片隔离、并发命名及 JSONL 协议。
CLI 的 CI 配置覆盖 Windows、Linux、macOS，以及 Python 3.8/3.12（macOS 仅 3.12），验证 Node.js 22.0，并在 Linux 额外验证 Node 24。
DSH 插件单独使用 Node 22.19 测试；真实 CLI 集成测试在 CI 中不允许静默跳过。

## 许可证

MIT License
