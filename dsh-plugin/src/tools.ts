/**
 * The model-facing document tools. Each tool validates its own argument
 * constraints, builds plugin-owned argv, runs the CLI through `ctx.shell`, and
 * returns a canonical value whose renderer is the model-facing text. Execution
 * outcomes are values; only invalid arguments, cancellation, and infrastructure
 * failures reject.
 * @module bruce-doc-converter-dsh/tools
 */

import type { Context } from '@deepseek-ai/cordis'
import { defineTool } from '@deepseek-ai/dsh-tools'
import { parseCliJson, readBatchEnvelope, readConvertEnvelope, readSetupEnvelope } from './envelope.ts'
import type { BdcRunner } from './runner.ts'
import type { BdcExecution } from './types.ts'
import {
  BATCH_VALUE_SCHEMA,
  BDC_NOT_FOUND,
  BDC_OUTPUT_TRUNCATED,
  BDC_PROTOCOL_ERROR,
  BDC_SANDBOX_DENIED,
  BDC_TIMEOUT,
  CONVERT_VALUE_SCHEMA,
  SETUP_VALUE_SCHEMA,
  batchFailureValue,
  batchValue,
  cliFailureValue,
  failureValue,
  formatBatch,
  formatConvert,
  formatSetup,
  inputFormatOf,
  setupFailureValue,
  setupValue,
  successValue,
} from './values.ts'
import type { BatchValue, ConvertValue, SetupValue } from './values.ts'

/** Everything the tools need from the plugin configuration. */
export interface DocToolOptions {
  /** CLI invocation owner over `ctx.shell`. */
  readonly runner: BdcRunner
  /** Configured executable identity, named in the missing-executable diagnostic. */
  readonly bdcPath: string
  /** Cooperative timeout budget for one conversion or batch (ms). */
  readonly convertTimeoutMs: number
  /** Cooperative timeout budget for dependency setup (ms). */
  readonly setupTimeoutMs: number
  /** Cap on Markdown characters returned to the model. */
  readonly maxMarkdownChars: number
  /** Foreground stdout capture budget the runner requests, quoted when it overflows. */
  readonly stdoutMaxBytes: number
  /** Cap on per-file batch outcomes carried in the canonical value. */
  readonly maxBatchEntries: number
  /** Largest accepted Mermaid PNG scale factor. */
  readonly maxMermaidScale: number
}

/** Cancellation is reported as an abort so the tool registry classifies it as such. */
function abortError(): Error {
  const error = new Error('tool call aborted')
  error.name = 'AbortError'
  return error
}

function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}

/** The tail of stderr, bounded, for a diagnostic that fits in one tool result. */
function stderrTail(execution: BdcExecution, maxChars = 400): string {
  const text = execution.stderr.trim()
  if (text.length === 0) return ''
  return text.length <= maxChars ? text : `…${text.slice(-maxChars)}`
}

/** The shell could not start the configured executable at all. */
function missingExecutable(execution: BdcExecution): boolean {
  if (execution.exitCode === 127) return true
  return /command not found|not recognized as the name of a cmdlet|not recognized as an internal or external command/i.test(execution.stderr)
}

function validateNonEmpty(name: string, value: string): void {
  if (value.trim().length === 0) {
    throw new Error(`invalid ${name}: expected a non-empty string`)
  }
}

function validateMermaidScale(value: number, max: number): void {
  if (!Number.isFinite(value) || value <= 0) {
    throw new Error(`invalid mermaidScale: expected a positive number, got ${JSON.stringify(value)}`)
  }
  if (value > max) {
    throw new Error(`invalid mermaidScale: ${value} exceeds the configured maximum ${max}`)
  }
}

/**
 * How one tool's run names its own limits. Each tool quotes the budget that
 * actually expired and the configuration key that raises it, so a diagnostic
 * never points at another tool's field.
 */
interface Invocation {
  /** Cooperative timeout this run used (ms). */
  readonly timeoutMs: number
  /** Recovery guidance when that timeout expires. */
  readonly timeoutSuggestion: string
  /** Recovery guidance when the stdout capture budget overflows. */
  readonly truncationSuggestion: string
}

/** The limits of the run behind `doc_convert`. */
function convertLimits(options: DocToolOptions): Invocation {
  return {
    timeoutMs: options.convertTimeoutMs,
    timeoutSuggestion: 'Raise `convertTimeoutMs` in the plugin config, or convert a smaller document.',
    truncationSuggestion: 'Raise `stdoutMaxBytes` in the plugin config, or convert a smaller document.',
  }
}

/** The limits of the run behind `doc_batch`. */
function batchLimits(options: DocToolOptions): Invocation {
  return {
    timeoutMs: options.convertTimeoutMs,
    timeoutSuggestion: 'Raise `convertTimeoutMs` in the plugin config, or scan a smaller directory.',
    truncationSuggestion: 'Convert the files one at a time with doc_convert: one batch keeps every file\'s output in a single capture budget.',
  }
}

/** The limits of the run behind `doc_setup`. */
function setupLimits(options: DocToolOptions): Invocation {
  return {
    timeoutMs: options.setupTimeoutMs,
    timeoutSuggestion: 'Raise `setupTimeoutMs` in the plugin config, or install the dependencies with `bdc setup-node` outside the harness.',
    truncationSuggestion: 'Raise `stdoutMaxBytes` in the plugin config.',
  }
}

/**
 * Classify a run that cannot yield a usable envelope, or `undefined` when the
 * caller should parse stdout instead. `onFailure` builds the caller's canonical
 * value shape for the classified code.
 */
function classifyWithoutPayload<T>(
  execution: BdcExecution,
  options: DocToolOptions,
  limits: Invocation,
  onFailure: (errorCode: string, error: string, extras?: { suggestion?: string; retryable?: boolean }) => T,
): T | undefined {
  if (execution.timedOut) {
    return onFailure(BDC_TIMEOUT, `bdc did not finish within ${limits.timeoutMs} ms`, {
      suggestion: limits.timeoutSuggestion,
      retryable: true,
    })
  }
  if (execution.sandboxDenied) {
    const mode = execution.sandboxMode === undefined ? '' : ` under ${execution.sandboxMode} mode`
    return onFailure(BDC_SANDBOX_DENIED, `the sandbox denied an operation${mode}`, {
      suggestion: 'Retry through an approved escalation, or grant the converter access to the input and output paths.',
    })
  }
  if (missingExecutable(execution)) {
    return onFailure(BDC_NOT_FOUND, `the bdc executable "${options.bdcPath}" could not be run`, {
      suggestion: 'Install it (for example `pipx install bruce-doc-converter`), or point the plugin config key `bdcPath` at the executable.',
    })
  }
  if (execution.stdoutTruncated) {
    return onFailure(BDC_OUTPUT_TRUNCATED, `the CLI output exceeded the ${options.stdoutMaxBytes}-byte stdout capture budget`, {
      suggestion: limits.truncationSuggestion,
    })
  }
  return undefined
}

/** Parse and classify one single-file run. */
async function convertExecution(
  inputPath: string,
  execution: BdcExecution,
  options: DocToolOptions,
): Promise<ConvertValue> {
  const classified = classifyWithoutPayload(execution, options, convertLimits(options), (errorCode, error, extras) =>
    failureValue(inputPath, errorCode, error, extras ?? {}))
  if (classified !== undefined) return classified as ConvertValue
  let payload: unknown
  try {
    payload = parseCliJson(execution.stdout)
  } catch (error) {
    const tail = stderrTail(execution)
    return failureValue(inputPath, BDC_PROTOCOL_ERROR, `${messageOf(error)}${tail === '' ? '' : `; stderr: ${tail}`}`, {
      retryable: false,
    })
  }
  try {
    const envelope = readConvertEnvelope(payload)
    return envelope.success ? successValue(envelope, options.maxMarkdownChars) : cliFailureValue(inputPath, envelope)
  } catch (error) {
    return failureValue(inputPath, BDC_PROTOCOL_ERROR, messageOf(error), { retryable: false })
  }
}

/** Parse and classify one batch run. */
async function batchExecution(
  directory: string,
  execution: BdcExecution,
  options: DocToolOptions,
): Promise<BatchValue> {
  const classified = classifyWithoutPayload(execution, options, batchLimits(options), (errorCode, error, extras) =>
    batchFailureValue(directory, errorCode, error, extras ?? {}))
  if (classified !== undefined) return classified as BatchValue
  let payload: unknown
  try {
    payload = parseCliJson(execution.stdout)
  } catch (error) {
    const tail = stderrTail(execution)
    return batchFailureValue(directory, BDC_PROTOCOL_ERROR, `${messageOf(error)}${tail === '' ? '' : `; stderr: ${tail}`}`)
  }
  try {
    return batchValue(directory, readBatchEnvelope(payload), options.maxBatchEntries)
  } catch (error) {
    return batchFailureValue(directory, BDC_PROTOCOL_ERROR, messageOf(error))
  }
}

/** Parse and classify one dependency-setup run. */
async function setupExecution(execution: BdcExecution, options: DocToolOptions): Promise<SetupValue> {
  const classified = classifyWithoutPayload(execution, options, setupLimits(options), (errorCode, error, extras) =>
    setupFailureValue(errorCode, error, extras ?? {}))
  if (classified !== undefined) return classified as SetupValue
  let payload: unknown
  try {
    payload = parseCliJson(execution.stdout)
  } catch (error) {
    const tail = stderrTail(execution)
    return setupFailureValue(BDC_PROTOCOL_ERROR, `${messageOf(error)}${tail === '' ? '' : `; stderr: ${tail}`}`)
  }
  try {
    return setupValue(readSetupEnvelope(payload))
  } catch (error) {
    return setupFailureValue(BDC_PROTOCOL_ERROR, messageOf(error))
  }
}

/** The `doc_convert` description, stating the configured Markdown cap. */
function convertDescription(options: DocToolOptions): string {
  return 'Convert one document with the local `bdc` CLI. A .docx, .xlsx, .pptx, or .pdf input produces Markdown and returns its text; a .md input produces a .docx file. '
    + 'Mermaid diagrams in Markdown are rendered as PNG images when a local browser is available. The path of the written file is always returned. '
    + `Returned Markdown is capped at ${options.maxMarkdownChars} characters; the complete text stays in the written file. `
    + 'Only .docx, .xlsx, .pptx, .pdf, and .md are accepted; legacy .doc/.xls/.ppt must be converted to the modern format first.'
}

/** The `doc_batch` description, stating the configured per-file cap. */
function batchDescription(options: DocToolOptions): string {
  return 'Convert every supported document in a directory with the local `bdc` CLI. '
    + `Per-file outcomes are capped at ${options.maxBatchEntries} entries and never include Markdown text; call doc_convert for one file's text. `
    + 'Use `recursive: false` to leave subdirectories alone.'
}

/** The `doc_setup` description. */
function setupDescription(): string {
  return 'Install the Node.js dependencies that Markdown-to-Word conversion needs, once per machine. '
    + 'The command is idempotent: it reports `alreadyInstalled: true` when the shared dependency directory already matches the installed package. '
    + 'Set `installBrowser: true` to also provision a Puppeteer browser for Mermaid rendering when no local Chrome, Edge, or Chromium is available; '
    + '`allowScripts: true` additionally permits npm lifecycle scripts during installation.'
}

/**
 * Register the document-conversion tool on the tool registry.
 * @param ctx - plugin context whose `ctx.tools` receives the definition.
 * @param options - resolved plugin configuration.
 */
export function registerConvertTool(ctx: Context, options: DocToolOptions): void {
  ctx.tools.register(defineTool({
    name: 'doc_convert',
    description: convertDescription(options),
    parameters: {
      path: { type: 'string', required: true, description: 'Path of the .docx, .xlsx, .pptx, .pdf, or .md file to convert.' },
      outputDir: { type: 'string', description: 'Directory for the generated file. Defaults to a `Markdown/` directory beside the input.' },
      extractImages: { type: 'boolean', description: 'Also write images embedded in a .docx, .xlsx, .pptx, or .pdf input next to the Markdown.' },
      mermaidScale: { type: 'number', description: 'PNG scale factor for Mermaid diagrams when converting Markdown to Word. Defaults to 4.' },
    },
    timeoutMs: options.convertTimeoutMs,
    isConcurrencySafe: () => true,
    output: {
      schema: CONVERT_VALUE_SCHEMA,
      render: (_args, value) => [{ type: 'text', text: formatConvert(value) }],
    },
    presentCall: (args) => ({
      card: 'generic',
      title: args.path,
      kind: inputFormatOf(args.path) === 'md' ? 'edit' : 'read',
      locations: [{ path: args.path }],
    }),
    async execute(args, exec) {
      validateNonEmpty('path', args.path)
      if (args.outputDir !== undefined) validateNonEmpty('outputDir', args.outputDir)
      if (args.mermaidScale !== undefined) validateMermaidScale(args.mermaidScale, options.maxMermaidScale)
      const argv = ['convert', args.path]
      if (args.outputDir !== undefined) argv.push('--output-dir', args.outputDir)
      if (args.extractImages === true) argv.push('--extract-images', 'true')
      if (args.mermaidScale !== undefined) argv.push('--mermaid-scale', String(args.mermaidScale))
      const execution = await options.runner.run(argv, exec.signal, options.convertTimeoutMs)
      if (execution.aborted) throw abortError()
      return await convertExecution(args.path, execution, options)
    },
  }))
}

/**
 * Register the directory-conversion tool on the tool registry.
 * @param ctx - plugin context whose `ctx.tools` receives the definition.
 * @param options - resolved plugin configuration.
 */
export function registerBatchTool(ctx: Context, options: DocToolOptions): void {
  ctx.tools.register(defineTool({
    name: 'doc_batch',
    description: batchDescription(options),
    parameters: {
      path: { type: 'string', required: true, description: 'Directory holding the documents to convert.' },
      outputDir: { type: 'string', description: 'Directory for the generated files. Defaults to a `Markdown/` directory inside the scanned directory.' },
      recursive: { type: 'boolean', description: 'Also convert supported files in subdirectories. Defaults to true.' },
      extractImages: { type: 'boolean', description: 'Also write images embedded in each Office/PDF input next to its Markdown.' },
      mermaidScale: { type: 'number', description: 'PNG scale factor for Mermaid diagrams when converting Markdown to Word. Defaults to 4.' },
    },
    timeoutMs: options.convertTimeoutMs,
    output: {
      schema: BATCH_VALUE_SCHEMA,
      render: (_args, value) => [{ type: 'text', text: formatBatch(value) }],
    },
    presentCall: (args) => ({ card: 'generic', title: args.path, kind: 'read', locations: [{ path: args.path }] }),
    async execute(args, exec) {
      validateNonEmpty('path', args.path)
      if (args.outputDir !== undefined) validateNonEmpty('outputDir', args.outputDir)
      if (args.mermaidScale !== undefined) validateMermaidScale(args.mermaidScale, options.maxMermaidScale)
      const argv = ['batch', args.path]
      if (args.outputDir !== undefined) argv.push('--output-dir', args.outputDir)
      if (args.recursive === false) argv.push('--recursive', 'false')
      if (args.extractImages === true) argv.push('--extract-images', 'true')
      if (args.mermaidScale !== undefined) argv.push('--mermaid-scale', String(args.mermaidScale))
      const execution = await options.runner.run(argv, exec.signal, options.convertTimeoutMs)
      if (execution.aborted) throw abortError()
      return await batchExecution(args.path, execution, options)
    },
  }))
}

/**
 * Register the dependency-setup tool on the tool registry.
 * @param ctx - plugin context whose `ctx.tools` receives the definition.
 * @param options - resolved plugin configuration.
 */
export function registerSetupTool(ctx: Context, options: DocToolOptions): void {
  ctx.tools.register(defineTool({
    name: 'doc_setup',
    description: setupDescription(),
    parameters: {
      installBrowser: { type: 'boolean', description: 'Also install the Puppeteer browser used for Mermaid rendering when no local browser is detected.' },
      allowScripts: { type: 'boolean', description: 'Permit npm lifecycle scripts while installing. Only set this for a source you trust.' },
    },
    timeoutMs: options.setupTimeoutMs,
    output: {
      schema: SETUP_VALUE_SCHEMA,
      render: (_args, value) => [{ type: 'text', text: formatSetup(value) }],
    },
    presentCall: () => ({ card: 'generic', title: 'Install Markdown-to-Word dependencies', kind: 'execute' }),
    async execute(args, exec) {
      const argv = ['setup-node']
      if (args.allowScripts === true) argv.push('--allow-scripts')
      if (args.installBrowser === true) argv.push('--install-browser')
      const execution = await options.runner.run(argv, exec.signal, options.setupTimeoutMs)
      if (execution.aborted) throw abortError()
      return await setupExecution(execution, options)
    },
  }))
}
