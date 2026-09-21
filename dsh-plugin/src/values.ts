/**
 * Canonical tool values and their model-facing text. One schema per tool
 * derives the value type, so a builder cannot drift from what the registry
 * validates; rendering stays a pure function of the validated value.
 * @module bruce-doc-converter-dsh/values
 */

import type { InferValue, ValueSchemaSpec } from '@deepseek-ai/dsh-tools'
import type { CliBatchEnvelope, CliConvertSuccess, CliFailure, CliSetupEnvelope } from './types.ts'

/** Plugin-owned failure codes; CLI-owned codes pass through unchanged. */
export const BDC_NOT_FOUND = 'BDC_NOT_FOUND'
export const BDC_TIMEOUT = 'BDC_TIMEOUT'
export const BDC_OUTPUT_TRUNCATED = 'BDC_OUTPUT_TRUNCATED'
export const BDC_PROTOCOL_ERROR = 'BDC_PROTOCOL_ERROR'
export const BDC_SANDBOX_DENIED = 'BDC_SANDBOX_DENIED'
export const BDC_CLI_INCOMPATIBLE = 'BDC_CLI_INCOMPATIBLE'

/** The CLI's own code for an argv it cannot parse, including an unknown subcommand. */
const USAGE_ERROR = 'USAGE_ERROR'

/** Recovery guidance for an installed CLI that predates this plugin's invocations. */
const UPGRADE_SUGGESTION =
  'The installed bdc CLI does not support this invocation, which means it predates the installed plugin. Upgrade it (for example `pipx upgrade bruce-doc-converter`) and retry.'

/** Extension of one path, lowercased without the dot, mirroring the CLI's `input_format`. */
export function inputFormatOf(path: string): string {
  const base = path.replaceAll('\\', '/').split('/').pop() ?? path
  const dot = base.lastIndexOf('.')
  if (dot <= 0 || dot === base.length - 1) return 'unknown'
  return base.slice(dot + 1).toLowerCase()
}

/** One conversion outcome, successful or not. */
const DIAGNOSTIC_SCHEMA = {
  type: 'object', additionalProperties: false,
  properties: {
    code: { type: 'string', required: true }, severity: { type: 'string', required: true },
    message: { type: 'string', required: true }, page: { type: 'number' },
    sheet: { type: 'string' }, cell: { type: 'string' }, status: { type: 'string' },
  },
} as const

export const CONVERT_VALUE_SCHEMA = {
  type: 'object',
  additionalProperties: false,
  properties: {
    success: { type: 'boolean', required: true, description: 'Whether the conversion produced its output.' },
    inputPath: { type: 'string', required: true, description: 'Absolute path of the converted input.' },
    inputFormat: { type: 'string', required: true, description: 'Input extension without the dot.' },
    outputFormat: { type: 'string', description: '`markdown` or `docx`.' },
    outputPath: { type: 'string', description: 'Path of the file the CLI wrote; absent when nothing was written.' },
    markdown: { type: 'string', description: 'Markdown text for Office/PDF input, capped by configuration.' },
    markdownChars: { type: 'number', description: 'Complete Markdown length in characters, before any cap.' },
    markdownTruncated: { type: 'boolean', description: 'Whether `markdown` holds only a prefix of the document.' },
    extractedImages: { type: 'array', items: { type: 'string' }, description: 'Paths of images extracted beside the Markdown.' },
    warnings: { type: 'array', items: { type: 'string' }, required: true, description: 'Non-fatal conversion warnings.' },
    diagnostics: { type: 'array', items: DIAGNOSTIC_SCHEMA },
    errorCode: { type: 'string', description: 'Machine-readable failure code.' },
    error: { type: 'string', description: 'Human-readable failure detail.' },
    suggestion: { type: 'string', description: 'Recovery guidance from the converter.' },
    retryable: { type: 'boolean', description: 'Whether retrying after the suggested action can succeed.' },
    nextCommand: { type: 'string', description: 'The CLI command that resolves the failure, when one exists.' },
  },
} satisfies ValueSchemaSpec

/** Validated value of one `doc_convert` call. */
export type ConvertValue = InferValue<typeof CONVERT_VALUE_SCHEMA>

/** One batch outcome with bounded per-file detail. */
export const BATCH_VALUE_SCHEMA = {
  type: 'object',
  additionalProperties: false,
  properties: {
    success: { type: 'boolean', required: true, description: 'Whether every file converted.' },
    directory: { type: 'string', required: true, description: 'Directory that was scanned.' },
    manifestPath: { type: 'string', description: 'JSONL manifest containing every outcome, flushed per file.' },
    total: { type: 'number', description: 'Files the CLI attempted.' },
    succeeded: { type: 'number', description: 'Files that converted.' },
    failed: { type: 'number', description: 'Files that did not convert.' },
    results: {
      type: 'array',
      description: 'Per-file outcomes, capped by configuration; no Markdown bodies are included.',
      items: {
        type: 'object',
        additionalProperties: false,
        properties: {
          inputPath: { type: 'string', required: true },
          success: { type: 'boolean', required: true },
          outputPath: { type: 'string' },
          errorCode: { type: 'string' },
          error: { type: 'string' },
          warnings: { type: 'array', items: { type: 'string' } },
          diagnostics: { type: 'array', items: DIAGNOSTIC_SCHEMA },
        },
      },
    },
    omitted: { type: 'number', description: 'Per-file outcomes dropped from `results` by the configured cap.' },
    errorCode: { type: 'string', description: 'Machine-readable failure code when the batch never ran.' },
    error: { type: 'string', description: 'Human-readable failure detail when the batch never ran.' },
    suggestion: { type: 'string', description: 'Recovery guidance when the batch never ran.' },
    retryable: { type: 'boolean', description: 'Whether retrying can succeed.' },
  },
} satisfies ValueSchemaSpec

/** Validated value of one `doc_batch` call. */
export type BatchValue = InferValue<typeof BATCH_VALUE_SCHEMA>

/** One Node.js dependency setup outcome. */
export const SETUP_VALUE_SCHEMA = {
  type: 'object',
  additionalProperties: false,
  properties: {
    success: { type: 'boolean', required: true, description: 'Whether the dependencies are ready.' },
    nodeHome: { type: 'string', description: 'Shared directory holding the installed Node.js dependencies.' },
    alreadyInstalled: { type: 'boolean', description: 'Whether the run found matching dependencies already present.' },
    installAction: { type: 'string', description: '`installed` or `skipped`.' },
    browserInstallAction: { type: 'string', description: 'Browser provisioning outcome for Mermaid rendering.' },
    detectedBrowserPath: { type: 'string', description: 'Local Chrome/Edge/Chromium path the converter will use.' },
    errorCode: { type: 'string', description: 'Machine-readable failure code.' },
    error: { type: 'string', description: 'Human-readable failure detail.' },
    suggestion: { type: 'string', description: 'Recovery guidance.' },
    retryable: { type: 'boolean', description: 'Whether retrying can succeed.' },
  },
} satisfies ValueSchemaSpec

/** Validated value of one `doc_setup` call. */
export type SetupValue = InferValue<typeof SETUP_VALUE_SCHEMA>

/** Optional failure facts shared by every builder. */
export interface FailureExtras {
  readonly suggestion?: string
  readonly retryable?: boolean
  readonly nextCommand?: string
}

/** Build one failed conversion value for a path the caller asked about. */
export function failureValue(inputPath: string, errorCode: string, error: string, extras: FailureExtras = {}): ConvertValue {
  return {
    success: false,
    inputPath,
    inputFormat: inputFormatOf(inputPath),
    warnings: [],
    errorCode,
    error,
    ...extras.suggestion === undefined ? {} : { suggestion: extras.suggestion },
    ...extras.retryable === undefined ? {} : { retryable: extras.retryable },
    ...extras.nextCommand === undefined ? {} : { nextCommand: extras.nextCommand },
  }
}

/** Map one CLI success envelope onto the canonical value, capping the Markdown. */
export function successValue(envelope: CliConvertSuccess, maxMarkdownChars: number): ConvertValue {
  const markdown = envelope.markdownContent
  const characters = markdown === undefined ? undefined : Array.from(markdown)
  const truncated = envelope.markdownTruncated === true || (characters !== undefined && characters.length > maxMarkdownChars)
  return {
    success: true,
    inputPath: envelope.inputPath,
    inputFormat: envelope.inputFormat,
    outputFormat: envelope.outputFormat,
    ...envelope.outputPath === null ? {} : { outputPath: envelope.outputPath },
    warnings: [...envelope.warnings],
    ...envelope.diagnostics === undefined ? {} : { diagnostics: [...envelope.diagnostics] },
    ...characters === undefined ? {} : { markdown: characters.slice(0, maxMarkdownChars).join('') },
    ...characters === undefined ? {} : { markdownChars: envelope.markdownChars ?? characters.length },
    ...truncated ? { markdownTruncated: true } : {},
    ...envelope.extractedImages === undefined ? {} : { extractedImages: [...envelope.extractedImages] },
  }
}

/**
 * Map one CLI failure envelope onto the canonical value. An argv the CLI
 * rejects becomes the plugin's own incompatible-CLI code, because every argv
 * here is plugin-owned: a usage error means the executable is too old for this
 * invocation, not that the model asked for something wrong.
 */
export function cliFailureValue(requestedPath: string, envelope: CliFailure): ConvertValue {
  const incompatible = envelope.errorCode === USAGE_ERROR
  return {
    success: false,
    inputPath: envelope.inputPath ?? requestedPath,
    inputFormat: envelope.inputFormat ?? inputFormatOf(requestedPath),
    warnings: [...(envelope.warnings ?? [])],
    ...envelope.diagnostics === undefined ? {} : { diagnostics: [...envelope.diagnostics] },
    errorCode: incompatible ? BDC_CLI_INCOMPATIBLE : envelope.errorCode,
    error: envelope.error,
    ...incompatible
      ? { suggestion: UPGRADE_SUGGESTION }
      : envelope.suggestion === undefined ? {} : { suggestion: envelope.suggestion },
    ...envelope.retryable === undefined ? {} : { retryable: envelope.retryable },
    ...envelope.nextCommand === undefined ? {} : { nextCommand: envelope.nextCommand },
  }
}

/** Map one CLI batch envelope onto the canonical value, capping the per-file detail. */
export function batchValue(directory: string, envelope: CliBatchEnvelope, maxEntries: number): BatchValue {
  const included = envelope.results.slice(0, maxEntries)
  const omitted = (envelope.omitted ?? 0) + envelope.results.length - included.length
  return {
    success: envelope.success,
    directory,
    ...envelope.manifestPath === undefined ? {} : { manifestPath: envelope.manifestPath },
    total: envelope.total,
    succeeded: envelope.succeeded,
    failed: envelope.failed,
    results: included.map((entry) => entry.result.success
      ? {
        inputPath: entry.inputPath,
        success: true,
        warnings: [...entry.result.warnings],
        ...entry.result.diagnostics === undefined ? {} : { diagnostics: [...entry.result.diagnostics] },
        ...entry.result.outputPath === null ? {} : { outputPath: entry.result.outputPath },
      }
      : { inputPath: entry.inputPath, success: false, errorCode: entry.result.errorCode, error: entry.result.error,
          warnings: [...(entry.result.warnings ?? [])],
          ...entry.result.diagnostics === undefined ? {} : { diagnostics: [...entry.result.diagnostics] } }),
    ...omitted > 0 ? { omitted } : {},
  }
}

/** Build one batch value for a failure that produced no batch envelope. */
export function batchFailureValue(
  directory: string,
  errorCode: string,
  error: string,
  extras: FailureExtras = {},
): BatchValue {
  return {
    success: false,
    directory,
    errorCode,
    error,
    ...extras.suggestion === undefined ? {} : { suggestion: extras.suggestion },
    ...extras.retryable === undefined ? {} : { retryable: extras.retryable },
  }
}

/** Map one CLI setup envelope onto the canonical value. */
export function setupValue(envelope: CliSetupEnvelope): SetupValue {
  if (envelope.success) {
    return {
      success: true,
      ...envelope.nodeHome === undefined ? {} : { nodeHome: envelope.nodeHome },
      ...envelope.alreadyInstalled === undefined ? {} : { alreadyInstalled: envelope.alreadyInstalled },
      ...envelope.installAction === undefined ? {} : { installAction: envelope.installAction },
      ...envelope.browserInstallAction === undefined ? {} : { browserInstallAction: envelope.browserInstallAction },
      ...envelope.detectedBrowserPath === undefined || envelope.detectedBrowserPath === null
        ? {}
        : { detectedBrowserPath: envelope.detectedBrowserPath },
    }
  }
  const incompatible = envelope.errorCode === USAGE_ERROR
  return {
    success: false,
    ...incompatible
      ? { errorCode: BDC_CLI_INCOMPATIBLE }
      : envelope.errorCode === undefined ? {} : { errorCode: envelope.errorCode },
    ...envelope.error === undefined ? {} : { error: envelope.error },
    ...incompatible ? { suggestion: UPGRADE_SUGGESTION }
      : envelope.suggestion === undefined ? {} : { suggestion: envelope.suggestion },
    ...envelope.retryable === undefined ? {} : { retryable: envelope.retryable },
  }
}

/** Build one failed setup value for a failure the plugin itself detected. */
export function setupFailureValue(errorCode: string, error: string, extras: FailureExtras = {}): SetupValue {
  return {
    success: false,
    errorCode,
    error,
    ...extras.suggestion === undefined ? {} : { suggestion: extras.suggestion },
    ...extras.retryable === undefined ? {} : { retryable: extras.retryable },
  }
}

/** Render one failed conversion value as model-facing text. */
export function formatFailure(value: ConvertValue): string {
  const lines = [`Conversion failed (${value.errorCode ?? 'UNKNOWN'}): ${value.error ?? 'no detail'}`]
  if (value.suggestion !== undefined) lines.push(`Suggestion: ${value.suggestion}`)
  if (value.errorCode === 'DEPENDENCY_INSTALL_REQUIRED') {
    lines.push('Call the doc_setup tool, then retry this call.')
  } else if (value.nextCommand !== undefined) {
    lines.push(`Next command: ${value.nextCommand}`)
  }
  return lines.join('\n')
}

/** Render one successful conversion value as model-facing text. */
export function formatSuccess(value: ConvertValue): string {
  const target = value.outputPath === undefined ? '(no file written)' : value.outputPath
  const lines = [`Converted ${value.inputPath} → ${target} (${value.outputFormat ?? 'unknown'}).`]
  if (value.extractedImages !== undefined && value.extractedImages.length > 0) {
    lines.push(`Extracted images: ${value.extractedImages.join(', ')}`)
  }
  if (value.warnings.length > 0) lines.push(`Warnings: ${value.warnings.join('; ')}`)
  if (value.markdown !== undefined) {
    lines.push('', value.markdown)
    if (value.markdownTruncated === true) {
      lines.push('', `[Markdown truncated after ${Array.from(value.markdown ?? '').length} of ${value.markdownChars ?? 0} characters; the complete document is at ${target}]`)
    }
  } else if (value.outputPath !== undefined) {
    lines.push('Open the written file to inspect the result.')
  }
  return lines.join('\n')
}

/** Render one conversion value as model-facing text. */
export function formatConvert(value: ConvertValue): string {
  return value.success ? formatSuccess(value) : formatFailure(value)
}

/** Render one batch value as model-facing text, one line per included file. */
export function formatBatch(value: BatchValue): string {
  if (value.total === undefined || value.succeeded === undefined || value.results === undefined) {
    const lines = [`Batch failed (${value.errorCode ?? 'UNKNOWN'}): ${value.error ?? 'no detail'}`]
    if (value.suggestion !== undefined) lines.push(`Suggestion: ${value.suggestion}`)
    return lines.join('\n')
  }
  const lines = [
    `Batch ${value.success ? 'completed' : 'completed with failures'}: ${value.succeeded}/${value.total} converted in ${value.directory}.`,
  ]
  if (value.omitted !== undefined && value.omitted > 0) {
    lines.push(`Per-file detail omits ${value.omitted} outcome(s); the counts above are complete.`)
  }
  for (const entry of value.results) {
    lines.push(entry.success
      ? `- ok: ${entry.inputPath} → ${entry.outputPath ?? '(no file written)'}`
      : `- failed: ${entry.inputPath} (${entry.errorCode ?? 'UNKNOWN'}) ${entry.error ?? ''}`.trimEnd())
    if (entry.warnings?.length) lines.push(`  Warnings: ${entry.warnings.join('; ')}`)
  }
  if (value.manifestPath !== undefined) lines.push(`Complete manifest: ${value.manifestPath}`)
  lines.push('Markdown bodies are not included in a batch result; call doc_convert for the text of one file.')
  return lines.join('\n')
}

/** Render one setup value as model-facing text. */
export function formatSetup(value: SetupValue): string {
  if (!value.success) {
    const lines = [`Node.js dependency setup failed (${value.errorCode ?? 'UNKNOWN'}): ${value.error ?? 'no detail'}`]
    if (value.suggestion !== undefined) lines.push(`Suggestion: ${value.suggestion}`)
    lines.push('Markdown to Word conversion stays unavailable until this succeeds.')
    return lines.join('\n')
  }
  const lines = [value.alreadyInstalled === true
    ? 'Node.js dependencies were already installed and match the current package.'
    : 'Node.js dependencies installed.']
  if (value.nodeHome !== undefined) lines.push(`Dependency directory: ${value.nodeHome}`)
  if (value.detectedBrowserPath !== undefined) lines.push(`Mermaid browser: ${value.detectedBrowserPath}`)
  else lines.push('No local Chrome/Edge/Chromium detected; Mermaid rendering needs doc_setup with installBrowser.')
  return lines.join('\n')
}
