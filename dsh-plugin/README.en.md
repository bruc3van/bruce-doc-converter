# bruce-doc-converter-dsh

English | [中文](README.md)

The official **DeepSeek Harness** bundle for [Bruce Doc Converter](../README.md) (the `bdc` CLI). Once installed, the agent converts documents through three model tools instead of assembling shell commands, remembering install methods, and parsing stdout JSON by hand.

## What it adds

`bdc` is already agent-facing: fixed arguments, JSON-only stdout, and `error_code` / `retryable` / `next_command` on failure. This bundle lifts that contract into the harness tool layer:

- **Execution stays inside harness boundaries.** The plugin builds the argv (quoting every element) and runs it through `ctx.shell`, so sandboxing, approval, timeouts, output caps, and telemetry all apply. The model no longer composes shell commands.
- **Configuration replaces prompt text.** The executable path, timeouts, and output caps are validated `cordis.yml` fields that fail loud, instead of a model guessing whether `pipx`, `uv`, `pip --user`, or a venv is available.
- **Structured results.** The CLI's JSON becomes a typed tool value: `success`, `outputPath`, `markdown`, `errorCode`, `retryable`, and `nextCommand` are fields, not prose to parse.
- **A short companion skill.** The plugin also registers `bruce-doc-converter`, a skill carrying only the judgment a tool schema cannot express: OCR a scanned PDF first, convert legacy formats first, and when to call `doc_setup`.

## Requirements

- DeepSeek Harness (`dsh`), `0.1.5-rc.1` or newer (`0.1.6-alpha.x` included).
- The `bdc` CLI installed: `pipx install bruce-doc-converter` (recommended) or `uv tool install bruce-doc-converter`.
- Markdown-to-Word needs Node.js dependencies; call the `doc_setup` tool once per machine (the `bdc setup-node` equivalent).

## Install

```sh
dsh plugin --profile web add bruce-doc-converter-dsh
```

Replace `web` with your profile name. The package declares `dsh.bundle`, so `dsh plugin` records it in the profile's `dsh.profile.bundles` and applies its patch layer at boot.

Confirm the row reached the effective configuration:

```sh
dsh --profile web --dump-config | grep -A 2 bruce-doc-converter
```

The `web` profile reloads live; restart the harness and the tools are available.

### Overriding configuration

The bundle inserts one row with id `bruce-doc-converter`. Override it by id from your profile's own `cordis.patch.yml` (applied after this bundle):

```yaml
- id: bruce-doc-converter
  name: bruce-doc-converter-dsh
  config:
    bdcPath: /Users/you/.local/pipx/venvs/bruce-doc-converter/bin/bdc
    convertTimeoutMs: 300000
```

> A patch replaces a row's entire `config`; restate every key the row needs, or rely on the schema defaults.

## Tools the model gets

| Tool | Purpose | Notes |
|---|---|---|
| `doc_convert` | Convert one file | `.docx` / `.xlsx` / `.pptx` / `.pdf` → Markdown with the text returned to the model; `.md` → `.docx` with the written path returned |
| `doc_batch` | Convert a directory | Returns per-file success and output paths only, **never Markdown bodies**; call `doc_convert` for one file's text |
| `doc_setup` | Install the Node.js dependencies for Markdown to Word | The idempotent `bdc setup-node` equivalent; only `installBrowser` downloads a Puppeteer browser, and only `allowScripts` permits npm lifecycle scripts |

Returned Markdown is capped at `maxMarkdownChars`; the complete text always stays in the file at `outputPath`. `doc_batch` keeps at most `maxBatchEntries` per-file entries and reports the rest as `omitted`.

Single-file calls request `--content preview` from the CLI. Batch calls use
`--content none --manifest auto --max-results`, preventing full bodies from
entering stdout before truncation. `manifestPath` points to every outcome in a
JSONL manifest. Both conversion tools accept `strict: true` and retain warnings
and page/cell diagnostics. A matching CLI is required; older CLIs report
`BDC_CLI_INCOMPATIBLE` with upgrade guidance.

## Configuration

| Field | Default | Meaning |
|---|---|---|
| `bdcPath` | `bdc` | Executable name or absolute path. Point it at `.venv/bin/bdc` for a venv install |
| `convertTimeoutMs` | `180000` | Cooperative timeout (ms) for one conversion or batch; also the tool timeout budget |
| `setupTimeoutMs` | `600000` | `doc_setup` timeout (ms) |
| `maxMarkdownChars` | `200000` | Cap on Markdown characters returned to the model; the file keeps everything |
| `maxBatchEntries` | `200` | Cap on per-file entries kept in a `doc_batch` result |
| `maxMermaidScale` | `16` | Largest accepted Mermaid PNG scale factor |
| `stdoutMaxBytes` | `33554432` | Foreground stdout capture budget; must hold the complete JSON envelope |
| `shellDialect` | `powershell` on win32, `posix` elsewhere | Quoting dialect. **Set it explicitly when `ctx.shell` runs in a remote execution world** |
| `convert` / `batch` / `setup` | `true` | Register each tool |
| `skill` | `true` | Register the embedded skill (skipped automatically when the composition has no skill registry) |

## Failure codes

Plugin-owned codes (CLI codes such as `UNSUPPORTED_FORMAT`, `EMPTY_PDF_CONTENT`, and `DEPENDENCY_INSTALL_REQUIRED` pass through unchanged):

| Code | Meaning |
|---|---|
| `BDC_NOT_FOUND` | The configured executable is missing or could not be run |
| `BDC_CLI_INCOMPATIBLE` | The installed `bdc` is too old for this invocation (usually no `setup-node`); the result says to upgrade |
| `BDC_TIMEOUT` | The run exceeded `convertTimeoutMs`; raise it in configuration |
| `BDC_SANDBOX_DENIED` | The sandbox denied a file operation — a policy denial, not a command failure |
| `BDC_OUTPUT_TRUNCATED` | Stdout exceeded the `stdoutMaxBytes` capture budget, so the JSON is incomplete; convert files one at a time with `doc_convert` when a batch hits it |
| `BDC_PROTOCOL_ERROR` | Stdout was not a usable JSON envelope; the message carries the stderr tail |

## Development

```sh
pnpm install
pnpm run build      # tsc emits lib/
pnpm run typecheck  # source and tests
pnpm test           # unit tests, plus real-CLI end-to-end when bdc is installed
```

`tests/integration.spec.ts` invokes the real CLI through Bash or Windows PowerShell. It skips locally when the CLI is absent; `BDC_REQUIRE_INTEGRATION=1` makes missing prerequisites fail the suite, as configured in CI.

Publishing:

```sh
pnpm run build
npm publish --registry https://registry.npmjs.org
```

> The default registry on some machines is `registry.npmmirror.com`, a read-only mirror that cannot publish, and `npm login` comes first. Mirrors also sync with a delay; pass `--registry https://registry.npmjs.org` to install immediately after publishing.

## Known limitations

- The plugin never probes the CLI version: `doc_setup` availability is decided by `bdc` itself and reported as `BDC_CLI_INCOMPATIBLE` rather than checked at mount time.
- A `doc_batch` run makes the CLI print every file's Markdown to stdout, so a large batch can reach `stdoutMaxBytes` and report `BDC_OUTPUT_TRUNCATED`; use `doc_convert` instead.
- In a remote execution world (E2B, SSH) the `ctx.shell` platform can differ from the host's, so `shellDialect` must be set explicitly.

## License

MIT
