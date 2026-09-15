/**
 * Foreground runs of the `bdc` CLI through the `ctx.shell` capability seam.
 * Confinement, approval, timeouts, and output collection belong to the mounted
 * shell provider; this module only builds plugin-owned argv and normalizes the
 * result the tool layer consumes.
 * @module bruce-doc-converter-dsh/runner
 */

import type { ShellExecRequest, ShellExecutor } from '@deepseek-ai/dsh-shell'
import { buildCommandLine } from './quoting.ts'
import type { BdcExecution, ShellDialect } from './types.ts'

/** Executable identity and capture budget for every CLI run. */
export interface BdcRunnerOptions {
  /** PATH name or absolute path of the `bdc` executable. */
  readonly bdcPath: string
  /** Dialect the mounted shell provider interprets. */
  readonly dialect: ShellDialect
  /** Foreground stdout capture budget; the JSON envelope must survive intact. */
  readonly stdoutMaxBytes: number
}

/** One CLI invocation owner. */
export class BdcRunner {
  /**
   * @param shell - the mounted shell provider that runs the CLI.
   * @param options - executable identity, dialect, and capture budget.
   */
  constructor(private readonly shell: ShellExecutor, private readonly options: BdcRunnerOptions) {}

  /**
   * Build the exact command line for a set of CLI arguments.
   * @param args - CLI arguments after the executable, in order.
   * @returns the quoted command line handed to `ctx.shell`.
   */
  buildCommand(args: readonly string[]): string {
    return buildCommandLine(this.options.dialect, [this.options.bdcPath, ...args])
  }

  /**
   * Run one foreground CLI invocation and normalize its outcome. Nonzero exits,
   * timeout kills, and abort kills resolve with their facts; only preparation
   * failure or cancellation before process publication rejects.
   * @param args - CLI arguments after the executable, in order.
   * @param signal - caller lifetime; the executor kills the command when it fires.
   * @param timeoutMs - cooperative timeout budget for this invocation.
   * @returns the normalized execution facts.
   */
  async run(args: readonly string[], signal: AbortSignal, timeoutMs: number): Promise<BdcExecution> {
    const command = this.buildCommand(args)
    const request: ShellExecRequest = {
      command,
      timeoutMs,
      stdoutMaxBytes: this.options.stdoutMaxBytes,
      signal,
    }
    const result = await this.shell.run(this.shell.resolve(request))
    return {
      command,
      exitCode: result.exitCode,
      signal: result.signal,
      timedOut: result.timedOut,
      aborted: result.aborted,
      stdout: result.stdout.text,
      stdoutTruncated: result.stdout.truncated,
      stderr: result.stderr.text,
      stderrTruncated: result.stderr.truncated,
      sandboxDenied: result.sandbox?.denied === true,
      ...result.sandbox === undefined ? {} : { sandboxMode: result.sandbox.mode as string },
    }
  }
}
