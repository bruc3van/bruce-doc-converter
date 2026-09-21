/**
 * End-to-end coverage against the real `bdc` CLI through a `bash -c` stand-in
 * that mirrors the local shell provider's argv. The suite skips when `bdc` is
 * not installed, so it stays keyless and dependency-free elsewhere.
 * @module bruce-doc-converter-dsh/tests/integration
 */

import { execFile } from 'node:child_process'
import { copyFile, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import type { ShellExecRequest, ShellExecSpec, ShellExecutor, ShellRunResult } from '@deepseek-ai/dsh-shell'
import { callTool, mountWith } from './harness.ts'

const FIXTURES = join(dirname(fileURLToPath(import.meta.url)), 'fixtures')

/** Run one command line the way the local bash provider does. */
const windows = process.platform === 'win32'
function execBash(command: string, timeoutMs = 120_000): Promise<{ code: number | null; stdout: string; stderr: string }> {
  return new Promise((resolve) => {
    const shell = windows ? 'powershell.exe' : 'bash'
    const argv = windows ? ['-NoProfile', '-NonInteractive', '-Command', `& ${command}; exit $LASTEXITCODE`] : ['-c', command]
    execFile(shell, argv, { encoding: 'utf8', maxBuffer: 64 * 1024 * 1024, timeout: timeoutMs, windowsHide: true }, (error, stdout, stderr) => {
      const failure = error as (Error & { code?: unknown }) | null
      const code = failure === null ? 0 : typeof failure.code === 'number' ? failure.code : 1
      resolve({ code, stdout: String(stdout), stderr: String(stderr) })
    })
  })
}

/** A shell provider with the local provider's contract and no sandbox. */
function realShell(): ShellExecutor {
  return {
    resolve(request: ShellExecRequest): ShellExecSpec {
      return {
        command: request.command,
        workdir: request.workdir ?? process.cwd(),
        timeoutMs: request.timeoutMs ?? 120_000,
        stdoutMaxBytes: request.stdoutMaxBytes ?? 64_000,
        sandboxPolicy: undefined,
        ...(request.signal === undefined ? {} : { signal: request.signal }),
      }
    },
    async run(spec: ShellExecSpec): Promise<ShellRunResult> {
      const result = await execBash(spec.command, spec.timeoutMs)
      return {
        exitCode: result.code,
        signal: null,
        timedOut: false,
        aborted: false,
        timeoutMs: spec.timeoutMs,
        stdout: { text: result.stdout, truncated: Buffer.byteLength(result.stdout) > spec.stdoutMaxBytes },
        stderr: { text: result.stderr, truncated: false },
      }
    },
  } as unknown as ShellExecutor
}

/** A fresh working directory holding the fixture document. */
async function fixtureDir(): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), 'bdc-dsh-'))
  await copyFile(join(FIXTURES, 'sample.docx'), join(dir, 'sample.docx'))
  return dir
}

const installed = (await execBash("'bdc' '--help-json'")).code === 0
if (!installed && process.env.BDC_REQUIRE_INTEGRATION === '1') throw new Error('bdc must be installed for integration tests')

function mountReal(config: Parameters<typeof mountWith>[1] = {}) {
  return mountWith(realShell(), { shellDialect: windows ? 'powershell' : 'posix', ...config })
}

describe.skipIf(!installed)('bdc end to end', () => {
  it('converts a real .docx and returns its Markdown', async () => {
    const dir = await fixtureDir()
    try {
      const harness = mountReal({ bdcPath: 'bdc' })
      const { value, text } = await callTool(harness, 'doc_convert', { path: join(dir, 'sample.docx') })
      expect(value['success']).toBe(true)
      expect(String(value['markdown'])).toContain('Quarterly Report')
      expect(String(value['markdown'])).toContain('12 percent')
      expect(String(value['outputPath'])).toMatch(/Markdown[/\\]sample\.md$/)
      expect(text).toContain('Converted')
      const written = await readFile(String(value['outputPath']), 'utf8')
      expect(written).toContain('Second bullet')
    } finally {
      await rm(dir, { recursive: true, force: true })
    }
  })

  it('honors an explicit output directory and image extraction', async () => {
    const dir = await fixtureDir()
    try {
      const target = join(dir, 'exported')
      const harness = mountReal({ bdcPath: 'bdc' })
      const { value } = await callTool(harness, 'doc_convert', { path: join(dir, 'sample.docx'), outputDir: target, extractImages: true })
      expect(value['success']).toBe(true)
      expect(String(value['outputPath'])).toMatch(/exported[/\\]sample\.md$/)
      expect(await readFile(String(value['outputPath']), 'utf8')).toContain('Quarterly Report')
    } finally {
      await rm(dir, { recursive: true, force: true })
    }
  })

  it('caps the returned Markdown while the written file keeps the document', async () => {
    const dir = await fixtureDir()
    try {
      const harness = mountReal({ bdcPath: 'bdc', maxMarkdownChars: 12 })
      const { value, text } = await callTool(harness, 'doc_convert', { path: join(dir, 'sample.docx') })
      expect(value['markdownTruncated']).toBe(true)
      expect(String(value['markdown'])).toHaveLength(12)
      expect(text).toContain('truncated after 12 of')
      expect(await readFile(String(value['outputPath']), 'utf8')).toContain('Second bullet')
    } finally {
      await rm(dir, { recursive: true, force: true })
    }
  })

  it('summarizes a real batch without Markdown bodies', async () => {
    const dir = await fixtureDir()
    try {
      const harness = mountReal({ bdcPath: 'bdc' })
      const { value, text } = await callTool(harness, 'doc_batch', { path: dir })
      expect(value).toMatchObject({ success: true, total: 1, succeeded: 1, failed: 0 })
      expect(JSON.stringify(value)).not.toContain('Quarterly Report')
      expect(text).toContain('1/1 converted')
      const manifest = await readFile(String(value['manifestPath']), 'utf8')
      expect(manifest).toContain('"type": "summary"')
      expect(manifest).not.toContain('markdown_content')
    } finally {
      await rm(dir, { recursive: true, force: true })
    }
  })

  it('reports a real unsupported-format failure', async () => {
    const dir = await fixtureDir()
    try {
      await writeFile(join(dir, 'legacy.doc'), 'not a real document')
      const harness = mountReal({ bdcPath: 'bdc' })
      const { value, text } = await callTool(harness, 'doc_convert', { path: join(dir, 'legacy.doc') })
      expect(value).toMatchObject({ success: false, errorCode: 'UNSUPPORTED_FORMAT', inputFormat: 'doc' })
      expect(text).toContain('Conversion failed (UNSUPPORTED_FORMAT)')
    } finally {
      await rm(dir, { recursive: true, force: true })
    }
  })

  it('reports a real missing-file failure', async () => {
    const dir = await fixtureDir()
    try {
      const harness = mountReal({ bdcPath: 'bdc' })
      const { value } = await callTool(harness, 'doc_convert', { path: join(dir, 'absent.docx') })
      expect(value).toMatchObject({ success: false, errorCode: 'FILE_NOT_FOUND' })
    } finally {
      await rm(dir, { recursive: true, force: true })
    }
  })

  it('reports a missing executable when bdcPath is wrong', async () => {
    const harness = mountReal({ bdcPath: 'bdc-does-not-exist' })
    const { value } = await callTool(harness, 'doc_convert', { path: '/tmp/whatever.docx' })
    expect(value).toMatchObject({ success: false, errorCode: 'BDC_NOT_FOUND' })
  })
})
