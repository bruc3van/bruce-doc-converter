/**
 * Validation of the JSON envelopes the `bdc` CLI writes to stdout. This is a
 * process boundary, so every field the plugin consumes is checked at runtime;
 * additive fields the CLI may grow later are ignored rather than rejected.
 * @module bruce-doc-converter-dsh/envelope
 */

import type { CliBatchEnvelope, CliConvertEnvelope, CliConvertSuccess, CliFailure, CliSetupEnvelope } from './types.ts'

/** One stdout payload that is not a usable `bdc` envelope. */
export class CliProtocolError extends Error {
  /**
   * @param message - what was expected and what was found.
   */
  constructor(message: string) {
    super(message)
    this.name = 'CliProtocolError'
  }
}

/** Parse stdout as JSON, naming the failure when the payload is not JSON. */
export function parseCliJson(stdout: string): unknown {
  const text = stdout.trim()
  if (text.length === 0) {
    throw new CliProtocolError('the CLI wrote no stdout payload')
  }
  try {
    return JSON.parse(text) as unknown
  } catch {
    // JSON.parse reports position only; the caller adds the stderr tail for context.
    throw new CliProtocolError('the CLI stdout is not JSON')
  }
}

/** Copy one present value onto a payload under its canonical key. */
function present<T>(key: string, value: T | undefined): Record<string, T> {
  return value === undefined ? {} : { [key]: value }
}

function asRecord(value: unknown, path: string): Record<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    throw new CliProtocolError(`${path} must be a JSON object`)
  }
  return value as Record<string, unknown>
}

function readString(record: Record<string, unknown>, key: string, path: string): string {
  const value = record[key]
  if (typeof value !== 'string') {
    throw new CliProtocolError(`${path}.${key} must be a string`)
  }
  return value
}

function readOptionalString(record: Record<string, unknown>, key: string, path: string): string | undefined {
  const value = record[key]
  if (value === undefined || value === null) return undefined
  if (typeof value !== 'string') {
    throw new CliProtocolError(`${path}.${key} must be a string when present`)
  }
  return value
}

function readBoolean(record: Record<string, unknown>, key: string, path: string): boolean {
  const value = record[key]
  if (typeof value !== 'boolean') {
    throw new CliProtocolError(`${path}.${key} must be a boolean`)
  }
  return value
}

function readOptionalBoolean(record: Record<string, unknown>, key: string, path: string): boolean | undefined {
  const value = record[key]
  if (value === undefined || value === null) return undefined
  if (typeof value !== 'boolean') {
    throw new CliProtocolError(`${path}.${key} must be a boolean when present`)
  }
  return value
}

function readNumber(record: Record<string, unknown>, key: string, path: string): number {
  const value = record[key]
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new CliProtocolError(`${path}.${key} must be a finite number`)
  }
  return value
}

function readStringArray(record: Record<string, unknown>, key: string, path: string): readonly string[] {
  const value = record[key]
  if (!Array.isArray(value)) {
    throw new CliProtocolError(`${path}.${key} must be an array`)
  }
  return value.map((item, index) => {
    if (typeof item !== 'string') {
      throw new CliProtocolError(`${path}.${key}[${index}] must be a string`)
    }
    return item
  })
}

function readOptionalStringArray(record: Record<string, unknown>, key: string, path: string): readonly string[] | undefined {
  if (record[key] === undefined || record[key] === null) return undefined
  return readStringArray(record, key, path)
}

function readFailure(record: Record<string, unknown>, path: string): CliFailure {
  return {
    schemaVersion: readOptionalString(record, 'schema_version', path) ?? null,
    success: false,
    ...present('inputPath', readOptionalString(record, 'input_path', path)),
    ...present('inputFormat', readOptionalString(record, 'input_format', path)),
    errorCode: readString(record, 'error_code', path),
    error: readString(record, 'error', path),
    ...present('suggestion', readOptionalString(record, 'suggestion', path)),
    ...present('retryable', readOptionalBoolean(record, 'retryable', path)),
    ...present('nextCommand', readOptionalString(record, 'next_command', path)),
  }
}

function readSuccess(record: Record<string, unknown>, path: string): CliConvertSuccess {
  const outputPath = record['output_path']
  if (outputPath !== undefined && outputPath !== null && typeof outputPath !== 'string') {
    throw new CliProtocolError(`${path}.output_path must be a string or null`)
  }
  return {
    schemaVersion: readString(record, 'schema_version', path),
    success: true,
    inputPath: readString(record, 'input_path', path),
    inputFormat: readString(record, 'input_format', path),
    outputFormat: readString(record, 'output_format', path),
    outputPath: outputPath ?? null,
    ...present('markdownContent', readOptionalString(record, 'markdown_content', path)),
    ...present('extractedImages', readOptionalStringArray(record, 'extracted_images', path)),
    warnings: record['warnings'] === undefined ? [] : readStringArray(record, 'warnings', path),
    ...present('message', readOptionalString(record, 'message', path)),
  }
}

/**
 * Validate one `bdc convert` payload.
 * @param value - the parsed stdout value.
 * @returns the typed single-file envelope.
 * @throws CliProtocolError when a consumed field is missing or mistyped.
 */
export function readConvertEnvelope(value: unknown): CliConvertEnvelope {
  const record = asRecord(value, 'convert payload')
  return readBoolean(record, 'success', 'convert payload')
    ? readSuccess(record, 'convert payload')
    : readFailure(record, 'convert payload')
}

/**
 * Validate one `bdc batch` payload.
 * @param value - the parsed stdout value.
 * @returns the typed batch envelope.
 * @throws CliProtocolError when a consumed field is missing or mistyped.
 */
export function readBatchEnvelope(value: unknown): CliBatchEnvelope {
  const record = asRecord(value, 'batch payload')
  const rawResults = record['results']
  if (!Array.isArray(rawResults)) {
    throw new CliProtocolError('batch payload.results must be an array')
  }
  const results = rawResults.map((item, index) => {
    const path = `batch payload.results[${index}]`
    const entry = asRecord(item, path)
    return { inputPath: readString(entry, 'input_path', path), result: readConvertEnvelope(entry['result']) }
  })
  return {
    success: readBoolean(record, 'success', 'batch payload'),
    total: readNumber(record, 'total', 'batch payload'),
    succeeded: readNumber(record, 'succeeded', 'batch payload'),
    failed: readNumber(record, 'failed', 'batch payload'),
    results,
  }
}

/**
 * Validate one `bdc setup-node` payload.
 * @param value - the parsed stdout value.
 * @returns the typed setup envelope.
 * @throws CliProtocolError when a consumed field is missing or mistyped.
 */
export function readSetupEnvelope(value: unknown): CliSetupEnvelope {
  const record = asRecord(value, 'setup payload')
  if (!readBoolean(record, 'success', 'setup payload')) {
    return {
      success: false,
      ...present('errorCode', readOptionalString(record, 'error_code', 'setup payload')),
      ...present('error', readOptionalString(record, 'error', 'setup payload')),
      ...present('retryable', readOptionalBoolean(record, 'retryable', 'setup payload')),
    }
  }
  return {
    success: true,
    ...present('nodeHome', readOptionalString(record, 'node_home', 'setup payload')),
    ...present('allowScripts', readOptionalBoolean(record, 'allow_scripts', 'setup payload')),
    ...present('alreadyInstalled', readOptionalBoolean(record, 'already_installed', 'setup payload')),
    ...present('installAction', readOptionalString(record, 'install_action', 'setup payload')),
    ...present('browserInstallAction', readOptionalString(record, 'browser_install_action', 'setup payload')),
    ...present('detectedBrowserPath', readOptionalString(record, 'detected_browser_path', 'setup payload')),
  }
}
