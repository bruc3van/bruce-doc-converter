/**
 * Shared test harness: mounts the plugin over a duck-typed shell and tool
 * registry, and drives a registered tool the way the registry does.
 * @module bruce-doc-converter-dsh/tests/harness
 */

import type { Context } from '@deepseek-ai/cordis'
import type { ShellExecRequest, ShellExecSpec, ShellExecutor, ShellRunResult } from '@deepseek-ai/dsh-shell'
import type { ToolDefinition, ToolRunContext } from '@deepseek-ai/dsh-tools'
import { apply } from '../src/index.ts'
import type { Config } from '../src/index.ts'

/** Per-run facts a test can override; unset fields take the settled-success defaults. */
export type Outcome = Partial<Omit<ShellRunResult, 'stdout' | 'stderr' | 'sandbox'>> & {
  readonly stdoutText?: string
  readonly stdoutTruncated?: boolean
  readonly stderrText?: string
  readonly sandbox?: ShellRunResult['sandbox']
}

/** A shell stand-in plus the command lines it was asked to run. */
export interface ShellProbe {
  readonly shell: ShellExecutor
  readonly commands: string[]
}

/** A mounted plugin plus the registrations it produced. */
export interface Harness {
  readonly tools: Map<string, ToolDefinition>
  readonly commands: string[]
  readonly injected: string[][]
  readonly skills: unknown[]
}

/**
 * Build a shell stand-in that records command lines and answers with one fixed
 * outcome, or with an outcome derived from the command line.
 * @param outcome - the settled facts for every run.
 * @returns the shell and its recorded command lines.
 */
export function stubShell(outcome: Outcome | ((command: string) => Outcome) = {}): ShellProbe {
  const commands: string[] = []
  const shell = {
    resolve(request: ShellExecRequest): ShellExecSpec {
      return {
        command: request.command,
        workdir: request.workdir ?? '/workspace',
        timeoutMs: request.timeoutMs ?? 0,
        stdoutMaxBytes: request.stdoutMaxBytes ?? 0,
        sandboxPolicy: undefined,
        ...(request.signal === undefined ? {} : { signal: request.signal }),
      }
    },
    async run(spec: ShellExecSpec): Promise<ShellRunResult> {
      commands.push(spec.command)
      const current = typeof outcome === 'function' ? outcome(spec.command) : outcome
      const { stdoutText, stdoutTruncated, stderrText, ...rest } = current
      return {
        exitCode: 0,
        signal: null,
        timedOut: false,
        aborted: false,
        timeoutMs: spec.timeoutMs,
        stdout: { text: stdoutText ?? '', truncated: stdoutTruncated ?? false },
        stderr: { text: stderrText ?? '', truncated: false },
        ...rest,
      }
    },
  }
  return { shell: shell as unknown as ShellExecutor, commands }
}

/**
 * Mount the plugin over one shell and a recording tool registry.
 * @param shell - the shell provider the plugin runs the CLI through.
 * @param config - plugin configuration; omitted fields take their defaults.
 * @returns the mounted registrations.
 */
export function mountWith(shell: ShellExecutor, config: Config = {}): Harness {
  const injected: string[][] = []
  const skills: unknown[] = []
  const tools = new Map<string, ToolDefinition>()
  const ctx = {
    tools: {
      register(definition: ToolDefinition): () => void {
        tools.set(definition.name, definition)
        return () => {}
      },
    },
    shell,
    effect<T>(execute: () => T): T {
      return execute()
    },
    inject(deps: string[], callback: (inner: unknown) => void): void {
      injected.push(deps)
      if (deps.includes('skills')) {
        callback({
          skills: {
            register(skill: unknown): () => void {
              skills.push(skill)
              return () => {}
            },
          },
          effect<T>(execute: () => T): T {
            return execute()
          },
        })
      }
    },
  }
  apply(ctx as unknown as Context, config)
  return { tools, commands: [], injected, skills }
}

/**
 * Mount the plugin over a shell stand-in.
 * @param config - plugin configuration; omitted fields take their defaults.
 * @param outcome - the settled facts for every run.
 * @returns the mounted registrations.
 */
export function mount(config: Config = {}, outcome: Outcome | ((command: string) => Outcome) = {}): Harness {
  const probe = stubShell(outcome)
  return { ...mountWith(probe.shell, config), commands: probe.commands }
}

/**
 * Execute one registered tool the way the registry does, then render its value.
 * @param harness - the mounted plugin.
 * @param name - tool name.
 * @param args - validated arguments.
 * @param signal - caller lifetime.
 * @returns the canonical value and its model-facing text.
 */
export async function callTool(
  harness: Harness,
  name: string,
  args: unknown,
  signal: AbortSignal = new AbortController().signal,
): Promise<{ value: Record<string, unknown>; text: string }> {
  const definition = harness.tools.get(name)
  if (definition === undefined) throw new Error(`tool ${name} is not registered`)
  const exec = {
    callId: 'call-1',
    name,
    arguments: args,
    rootCallId: 'call-1',
    token: 'token-1',
    signal,
    deferContext(): void {},
    concludeTurn(): void {},
  } as unknown as ToolRunContext
  const value = await definition.execute(args, exec)
  const render = definition.output.render as unknown as (args: unknown, value: unknown) => { type: string; text: string }[]
  return { value: value as Record<string, unknown>, text: render(args, value).map((block) => block.text).join('\n') }
}
