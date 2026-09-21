/**
 * Vocabulary shared by the bdc-backed document tools: the JSON envelopes the
 * `bdc` CLI emits, one normalized execution result, and the shell dialect used
 * to quote argv before it reaches `ctx.shell`.
 * @module bruce-doc-converter-dsh/types
 */

/** Shell dialect used to quote argv for the mounted `ctx.shell` provider. */
export type ShellDialect = 'posix' | 'powershell'

export interface CliDiagnostic {
  readonly code: string
  readonly severity: string
  readonly message: string
  readonly page?: number
  readonly sheet?: string
  readonly cell?: string
  readonly status?: string
}

/**
 * One successful single-file conversion as `bdc convert` reports it. Markdown
 * input produces the DOCX path and `message`; Office/PDF input produces the
 * Markdown text and any extracted image paths.
 */
export interface CliConvertSuccess {
  readonly schemaVersion: string
  readonly success: true
  readonly inputPath: string
  readonly inputFormat: string
  readonly outputFormat: string
  readonly outputPath: string | null
  readonly markdownContent?: string
  readonly markdownChars?: number
  readonly markdownTruncated?: boolean
  readonly diagnostics?: readonly CliDiagnostic[]
  readonly extractedImages?: readonly string[]
  readonly warnings: readonly string[]
  readonly message?: string
}

/** One failed CLI outcome; the CLI classifies its own errors into `errorCode`. */
export interface CliFailure {
  readonly warnings?: readonly string[]
  readonly diagnostics?: readonly CliDiagnostic[]
  readonly schemaVersion: string | null
  readonly success: false
  readonly inputPath?: string
  readonly inputFormat?: string
  readonly errorCode: string
  readonly error: string
  readonly suggestion?: string
  readonly retryable?: boolean
  readonly nextCommand?: string
}

/** One `bdc convert` payload. */
export type CliConvertEnvelope = CliConvertSuccess | CliFailure

/** One entry of a `bdc batch` payload. */
export interface CliBatchEntry {
  readonly inputPath: string
  readonly result: CliConvertEnvelope
}

/** One `bdc batch` payload. */
export interface CliBatchEnvelope {
  readonly manifestPath?: string
  readonly omitted?: number
  readonly success: boolean
  readonly total: number
  readonly succeeded: number
  readonly failed: number
  readonly results: readonly CliBatchEntry[]
}

/** One `bdc setup-node` payload. */
export interface CliSetupEnvelope {
  readonly suggestion?: string
  readonly success: boolean
  readonly nodeHome?: string
  readonly allowScripts?: boolean
  readonly alreadyInstalled?: boolean
  readonly installAction?: string
  readonly browserInstallAction?: string
  readonly detectedBrowserPath?: string | null
  readonly errorCode?: string
  readonly error?: string
  readonly retryable?: boolean
}

/** One normalized foreground run of the CLI through `ctx.shell`. */
export interface BdcExecution {
  /** The exact command line handed to the shell provider. */
  readonly command: string
  readonly exitCode: number | null
  readonly signal: NodeJS.Signals | null
  /** The executor's own timeout was the first cause to stop the command. */
  readonly timedOut: boolean
  /** The caller's abort signal was the first cause to stop the command. */
  readonly aborted: boolean
  readonly stdout: string
  /** Stdout overflowed the capture budget, so `stdout` is a tail and not parseable JSON. */
  readonly stdoutTruncated: boolean
  readonly stderr: string
  readonly stderrTruncated: boolean
  /** The sandbox denied an operation during this run. */
  readonly sandboxDenied: boolean
  /** The mode the command actually ran under, when a sandboxing executor handled it. */
  readonly sandboxMode?: string
}
