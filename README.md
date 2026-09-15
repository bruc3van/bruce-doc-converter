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

查看帮助：

```bash
bdc --help-json
```

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
- **图片提取**：Office/PDF 转 Markdown 时可提取内嵌图片

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
- **Node.js 14+**（可选，仅 Markdown → Word 需要）

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
│   └── SKILL.md                  # Agent Skill 定义
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

## 许可证

MIT License
