/**
 * DeepSeek Harness bundle that exposes the `bdc` document converter to the
 * agent. The bundle is the whole distribution: a profile installs it with
 * `dsh plugin --profile <name> add bruce-doc-converter-dsh`, and this patch row
 * mounts the plugin, which registers the document tools on `ctx.tools` and its
 * guidance skill on `ctx.skills` when that registry is present.
 * @module bruce-doc-converter-dsh
 */

import type { Context } from '@deepseek-ai/cordis'
import z from '@deepseek-ai/schemastery'
import type {} from '@deepseek-ai/dsh-shell'
import type {} from '@deepseek-ai/dsh-tools'
import { BdcRunner } from './runner.ts'
import { registerGuidanceSkill } from './skill.ts'
import { registerBatchTool, registerConvertTool, registerSetupTool } from './tools.ts'
import type { DocToolOptions } from './tools.ts'
import type { ShellDialect } from './types.ts'

/** Cordis plugin name used by loader diagnostics. */
export const name = 'bruce-doc-converter'

/** The tool registry and the shell seam this plugin runs the CLI through. */
export const inject = ['tools', 'shell']

/** Default executable name; a pipx or pip `--user` install puts `bdc` on PATH. */
export const DEFAULT_BDC_PATH = 'bdc'
/** Default cooperative timeout for one conversion or batch (ms). */
export const DEFAULT_CONVERT_TIMEOUT_MS = 180_000
/** Default cooperative timeout for dependency setup (ms). */
export const DEFAULT_SETUP_TIMEOUT_MS = 600_000
/** Default cap on Markdown characters returned to the model. */
export const DEFAULT_MAX_MARKDOWN_CHARS = 200_000
/** Default cap on per-file batch outcomes carried in the canonical value. */
export const DEFAULT_MAX_BATCH_ENTRIES = 200
/** Default largest accepted Mermaid PNG scale factor. */
export const DEFAULT_MAX_MERMAID_SCALE = 16
/** Default foreground stdout capture budget; the JSON envelope must survive intact. */
export const DEFAULT_STDOUT_MAX_BYTES = 32 * 1024 * 1024

/** Default dialect for the local shell providers: pwsh on win32, POSIX elsewhere. */
export const DEFAULT_SHELL_DIALECT: ShellDialect = process.platform === 'win32' ? 'powershell' : 'posix'

/** Plugin configuration. */
export interface Config {
  /** PATH name or absolute path of the `bdc` executable. Defaults to `bdc`. */
  bdcPath?: string
  /** Cooperative timeout for one conversion or batch (ms). Defaults to 180000. */
  convertTimeoutMs?: number
  /** Cooperative timeout for `doc_setup` (ms). Defaults to 600000. */
  setupTimeoutMs?: number
  /** Cap on Markdown characters returned to the model. Defaults to 200000. */
  maxMarkdownChars?: number
  /** Cap on per-file batch outcomes in the canonical value. Defaults to 200. */
  maxBatchEntries?: number
  /** Largest accepted Mermaid PNG scale factor. Defaults to 16. */
  maxMermaidScale?: number
  /** Foreground stdout capture budget in bytes. Defaults to 33554432. */
  stdoutMaxBytes?: number
  /** Dialect the mounted shell provider interprets. Defaults to the host platform's. */
  shellDialect?: ShellDialect
  /** Register `doc_convert`. Defaults to true. */
  convert?: boolean
  /** Register `doc_batch`. Defaults to true. */
  batch?: boolean
  /** Register `doc_setup`. Defaults to true. */
  setup?: boolean
  /** Register the embedded guidance skill when a skill registry is mounted. Defaults to true. */
  skill?: boolean
}

/** Runtime configuration schema for the bundle. */
export const Config: z<Config> = z.object({
  bdcPath: z.string().default(DEFAULT_BDC_PATH),
  convertTimeoutMs: z.number().default(DEFAULT_CONVERT_TIMEOUT_MS),
  setupTimeoutMs: z.number().default(DEFAULT_SETUP_TIMEOUT_MS),
  maxMarkdownChars: z.number().default(DEFAULT_MAX_MARKDOWN_CHARS),
  maxBatchEntries: z.number().default(DEFAULT_MAX_BATCH_ENTRIES),
  maxMermaidScale: z.number().default(DEFAULT_MAX_MERMAID_SCALE),
  stdoutMaxBytes: z.number().default(DEFAULT_STDOUT_MAX_BYTES),
  shellDialect: z.union([z.const('posix'), z.const('powershell')]).default(DEFAULT_SHELL_DIALECT),
  convert: z.boolean().default(true),
  batch: z.boolean().default(true),
  setup: z.boolean().default(true),
  skill: z.boolean().default(true),
})

/** Complete config after every default is applied. */
type ResolvedConfig = Required<Config>

/** Configured limits must be positive integers. */
function assertPositiveInteger(field: string, value: number): void {
  if (!Number.isInteger(value) || value < 1) {
    throw new Error(`bruce-doc-converter: ${field} must be a positive integer`)
  }
}

/** Resolve raw config into the fully defaulted values the tools consume. */
function resolveConfig(config: Config): ResolvedConfig {
  const resolved: ResolvedConfig = {
    bdcPath: config.bdcPath ?? DEFAULT_BDC_PATH,
    convertTimeoutMs: config.convertTimeoutMs ?? DEFAULT_CONVERT_TIMEOUT_MS,
    setupTimeoutMs: config.setupTimeoutMs ?? DEFAULT_SETUP_TIMEOUT_MS,
    maxMarkdownChars: config.maxMarkdownChars ?? DEFAULT_MAX_MARKDOWN_CHARS,
    maxBatchEntries: config.maxBatchEntries ?? DEFAULT_MAX_BATCH_ENTRIES,
    maxMermaidScale: config.maxMermaidScale ?? DEFAULT_MAX_MERMAID_SCALE,
    stdoutMaxBytes: config.stdoutMaxBytes ?? DEFAULT_STDOUT_MAX_BYTES,
    shellDialect: config.shellDialect ?? DEFAULT_SHELL_DIALECT,
    convert: config.convert ?? true,
    batch: config.batch ?? true,
    setup: config.setup ?? true,
    skill: config.skill ?? true,
  }
  if (resolved.bdcPath.trim().length === 0) {
    throw new Error('bruce-doc-converter: bdcPath must be a non-empty string')
  }
  assertPositiveInteger('convertTimeoutMs', resolved.convertTimeoutMs)
  assertPositiveInteger('setupTimeoutMs', resolved.setupTimeoutMs)
  assertPositiveInteger('maxMarkdownChars', resolved.maxMarkdownChars)
  assertPositiveInteger('maxBatchEntries', resolved.maxBatchEntries)
  assertPositiveInteger('maxMermaidScale', resolved.maxMermaidScale)
  assertPositiveInteger('stdoutMaxBytes', resolved.stdoutMaxBytes)
  return resolved
}

/**
 * Mount the bundle: register the enabled document tools and the guidance skill.
 * Every registration is a Cordis effect, so unloading the plugin removes the
 * tools and restores the prior registries.
 * @param ctx - plugin context carrying `ctx.tools` and `ctx.shell`.
 * @param config - plugin configuration; omitted fields take their defaults.
 */
export function apply(ctx: Context, config: Config = {}): void {
  const resolved = resolveConfig(config)
  const options: DocToolOptions = {
    runner: new BdcRunner(ctx.shell, {
      bdcPath: resolved.bdcPath,
      dialect: resolved.shellDialect,
      stdoutMaxBytes: resolved.stdoutMaxBytes,
    }),
    bdcPath: resolved.bdcPath,
    convertTimeoutMs: resolved.convertTimeoutMs,
    setupTimeoutMs: resolved.setupTimeoutMs,
    maxMarkdownChars: resolved.maxMarkdownChars,
    stdoutMaxBytes: resolved.stdoutMaxBytes,
    maxBatchEntries: resolved.maxBatchEntries,
    maxMermaidScale: resolved.maxMermaidScale,
  }
  if (resolved.convert) registerConvertTool(ctx, options)
  if (resolved.batch) registerBatchTool(ctx, options)
  if (resolved.setup) registerSetupTool(ctx, options)
  if (resolved.skill) registerGuidanceSkill(ctx)
}
