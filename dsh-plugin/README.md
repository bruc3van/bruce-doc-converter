# bruce-doc-converter-dsh

[English](README.en.md) | 中文

把 [Bruce Doc Converter](../README.md)（`bdc` CLI）接入 **DeepSeek Harness** 的官方 bundle。安装后，智能体通过三个模型工具完成文档转换，不再需要自己在 shell 里拼命令、记安装方式、解析 stdout JSON。

## 它解决什么

`bdc` 本身已经是一个面向 Agent 的 CLI：固定参数、stdout 只输出 JSON、失败时给出 `error_code` / `retryable` / `next_command`。这个插件的价值在于把这份契约搬到 Harness 的工具层：

- **执行留在 Harness 的边界内。** 命令由插件拼好（argv 逐项引用转义），经 `ctx.shell` 运行，因此沙箱、审批、超时、输出上限、遥测都照常生效；模型不再自己拼 shell 命令。
- **配置取代提示词。** 可执行文件路径、超时、返回上限都是 `cordis.yml` 里经过校验的配置项，装错就报错，而不是让模型去猜 `pipx` / `uv` / `pip --user` / `venv` 哪一种装法可用。
- **结果结构化。** CLI 的 JSON 变成类型化的工具结果：`success`、`outputPath`、`markdown`、`errorCode`、`retryable`、`nextCommand` 都是字段，模型不必解析散文。
- **配套技能。** 插件还会注册一个很短的 skill（`bruce-doc-converter`），只放工具 schema 装不下的判断性内容：扫描件要先 OCR、旧格式要转新格式、什么时候该跑 `doc_setup`。

## 环境要求

- DeepSeek Harness（`dsh`）`0.1.5-rc.1` 及以上（`0.1.6-alpha.x` 亦可）。
- 已安装 `bdc` CLI：`pipx install bruce-doc-converter`（推荐）或 `uv tool install bruce-doc-converter`。
- Markdown → Word 需要 Node.js 依赖，首次使用调用一次 `doc_setup` 工具即可（等价于 `bdc setup-node`）。

## 安装

```sh
dsh plugin --profile web add bruce-doc-converter-dsh
```

把 `web` 换成你自己的 profile 名即可。该包声明了 `dsh.bundle`，因此 `dsh plugin` 会把它记录进 profile 的 `dsh.profile.bundles`，并在启动时应用它带来的补丁层。

安装后确认插件行已进入有效配置：

```sh
dsh --profile web --dump-config | grep -A 2 bruce-doc-converter
```

`web` profile 支持热重载，重启 Harness 后工具即可用。

### 覆盖配置

插件行由 bundle 插入，id 为 `bruce-doc-converter`。在你的 profile 自己的 `cordis.patch.yml`（它在本 bundle 之后应用）里按 id 覆盖即可：

```yaml
- id: bruce-doc-converter
  name: bruce-doc-converter-dsh
  config:
    bdcPath: /Users/you/.local/pipx/venvs/bruce-doc-converter/bin/bdc
    convertTimeoutMs: 300000
```

> 补丁是按行整体替换 `config` 的，不需要的键请一并写上或依赖默认值。

## 模型工具

| 工具 | 作用 | 说明 |
|---|---|---|
| `doc_convert` | 转换单个文件 | `.docx` / `.xlsx` / `.pptx` / `.pdf` → Markdown，并把正文返回给模型；`.md` → `.docx`，返回生成文件路径 |
| `doc_batch` | 转换整个目录 | 只返回每个文件的成功/失败与输出路径，**不返回正文**；需要正文时再对单个文件调用 `doc_convert` |
| `doc_setup` | 安装 Markdown → Word 所需的 Node.js 依赖 | 等价于 `bdc setup-node`，幂等；`installBrowser` 才会额外下载 Puppeteer 浏览器，`allowScripts` 才会允许 npm 生命周期脚本 |

模型返回的 Markdown 会被截断到 `maxMarkdownChars` 个字符，完整内容始终写在 `outputPath` 指向的文件里；`doc_batch` 的逐文件明细最多 `maxBatchEntries` 条，超出部分以 `omitted` 计数。

## 配置项

| 字段 | 默认值 | 含义 |
|---|---|---|
| `bdcPath` | `bdc` | `bdc` 可执行文件名或绝对路径。用 venv 安装时指向 `.venv/bin/bdc` |
| `convertTimeoutMs` | `180000` | 单次转换/批量的协作式超时（毫秒），同时作为工具超时预算 |
| `setupTimeoutMs` | `600000` | `doc_setup` 的超时（毫秒） |
| `maxMarkdownChars` | `200000` | 返回给模型的 Markdown 字符上限，完整内容仍在文件里 |
| `maxBatchEntries` | `200` | `doc_batch` 结果里保留的逐文件明细条数上限 |
| `maxMermaidScale` | `16` | 允许的 Mermaid PNG 最大倍率 |
| `stdoutMaxBytes` | `33554432` | 前台 stdout 采集上限；必须能装下完整 JSON 信封 |
| `shellDialect` | `win32` 为 `powershell`，其余为 `posix` | 拼命令行时的引用方言。**shell 跑在远端执行世界时请显式设置** |
| `convert` / `batch` / `setup` | `true` | 分别控制三个工具是否注册 |
| `skill` | `true` | 是否注册内置技能（组合里没有 skill 注册表时自动不注册） |

## 失败分类

插件自己产生的错误码（CLI 自己的错误码原样透传，例如 `UNSUPPORTED_FORMAT`、`EMPTY_PDF_CONTENT`、`DEPENDENCY_INSTALL_REQUIRED`）：

| 错误码 | 含义 |
|---|---|
| `BDC_NOT_FOUND` | 找不到或无法执行配置的可执行文件 |
| `BDC_CLI_INCOMPATIBLE` | 已安装的 `bdc` 太旧，不认这个调用（通常是缺少 `setup-node`），提示升级 |
| `BDC_TIMEOUT` | 超出 `convertTimeoutMs`，可通过配置放宽 |
| `BDC_SANDBOX_DENIED` | 沙箱拒绝了文件操作，属于策略拒绝而不是命令失败 |
| `BDC_OUTPUT_TRUNCATED` | stdout 超过 `stdoutMaxBytes` 采集上限，JSON 不完整；`doc_batch` 出现时请改用 `doc_convert` 逐文件处理 |
| `BDC_PROTOCOL_ERROR` | stdout 不是可用的 JSON 信封，错误信息里会带上 stderr 尾部 |

## 开发

```sh
pnpm install
pnpm run build      # tsc 输出到 lib/
pnpm run typecheck  # 源码 + 测试
pnpm test           # 单元测试；本机有 bdc 时还会跑真实 CLI 端到端
```

`tests/integration.spec.ts` 会在 `command -v bdc` 失败时自动跳过，因此无 `bdc` 的环境也能跑通。

发布：

```sh
pnpm run build
npm publish --registry https://registry.npmjs.org
```

> 本机 npm 默认指向 `registry.npmmirror.com`（只读镜像，不能发布），且需要先 `npm login`。发布后镜像同步有延迟，想立刻安装请显式指定 `--registry https://registry.npmjs.org`。

## 已知限制

- 插件不缓存 CLI 版本：`doc_setup` 是否可用由 `bdc` 自己决定，过旧时以 `BDC_CLI_INCOMPATIBLE` 报告而不是启动时探测。
- `doc_batch` 的 stdout 包含每个文件的 Markdown 正文，批量很大时可能触达 `stdoutMaxBytes`；此时会报 `BDC_OUTPUT_TRUNCATED`，请改用 `doc_convert`。
- 远端执行世界（E2B / SSH 等）下 `ctx.shell` 的平台可能与本机不同，`shellDialect` 需要手动指定。

## 许可证

MIT
