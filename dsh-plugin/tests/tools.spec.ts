import { describe, expect, it } from 'vitest'
import { callTool, mount } from './harness.ts'

const CONVERT_SUCCESS = JSON.stringify({
  schema_version: '1.0',
  success: true,
  input_path: '/docs/report.docx',
  input_format: 'docx',
  output_format: 'markdown',
  output_path: '/docs/Markdown/report.md',
  markdown_content: '# Report\n\nBody text.',
  extracted_images: ['/docs/Markdown/images/a.png'],
  warnings: ['heading level 7 mapped to 6'],
})

describe('registration', () => {
  it('registers all three tools and the guidance skill by default', () => {
    const harness = mount()
    expect([...harness.tools.keys()].sort()).toEqual(['doc_batch', 'doc_convert', 'doc_setup'])
    expect(harness.skills).toHaveLength(1)
    expect(harness.injected).toContainEqual(['skills'])
  })

  it('honors the enable switches', () => {
    const harness = mount({ convert: false, batch: false, setup: false, skill: false })
    expect([...harness.tools.keys()]).toEqual([])
    expect(harness.skills).toHaveLength(0)
  })

  it('rejects a non-positive limit at mount time', () => {
    expect(() => mount({ maxMarkdownChars: 0 })).toThrow(/must be a positive integer/)
    expect(() => mount({ bdcPath: '  ' })).toThrow(/bdcPath must be a non-empty string/)
  })
})

describe('doc_convert', () => {
  it('preserves CLI preview length and Unicode characters', async () => {
    const payload = { ...JSON.parse(CONVERT_SUCCESS), markdown_content: '😀文', markdown_chars: 100, markdown_truncated: true }
    const harness = mount({ maxMarkdownChars: 2 }, { stdoutText: JSON.stringify(payload) })
    const { value, text } = await callTool(harness, 'doc_convert', { path: '/docs/report.docx', strict: true })
    expect(value).toMatchObject({ markdown: '😀文', markdownChars: 100, markdownTruncated: true })
    expect(text).toContain('truncated after 2 of 100 characters')
    expect(harness.commands[0]).toContain("'--strict'")
  })

  it('preserves diagnostics on a strict content failure', async () => {
    const diagnostic = { code: 'FORMULA_CACHE_MISSING', severity: 'warning', message: 'uncached', sheet: 'Budget', cell: 'B4' }
    const harness = mount({}, { exitCode: 1, stdoutText: JSON.stringify({ success: false, error_code: 'CONTENT_INCOMPLETE', error: 'incomplete', warnings: ['uncached'], diagnostics: [diagnostic] }) })
    const { value } = await callTool(harness, 'doc_convert', { path: '/docs/budget.xlsx', strict: true })
    expect(value).toMatchObject({ success: false, warnings: ['uncached'], diagnostics: [diagnostic] })
  })
  it('returns the Markdown text and the written path', async () => {
    const harness = mount({}, { stdoutText: CONVERT_SUCCESS })
    const { value, text } = await callTool(harness, 'doc_convert', { path: '/docs/report.docx' })
    expect(value).toMatchObject({
      success: true,
      inputPath: '/docs/report.docx',
      outputPath: '/docs/Markdown/report.md',
      markdown: '# Report\n\nBody text.',
      extractedImages: ['/docs/Markdown/images/a.png'],
      warnings: ['heading level 7 mapped to 6'],
    })
    expect(text).toContain('Converted /docs/report.docx → /docs/Markdown/report.md (markdown).')
    expect(text).toContain('# Report')
  })

  it('caps the Markdown and reports the complete length', async () => {
    const harness = mount({ maxMarkdownChars: 10 }, { stdoutText: CONVERT_SUCCESS })
    const { value, text } = await callTool(harness, 'doc_convert', { path: '/docs/report.docx' })
    expect(value['markdown']).toBe('# Report\n\n')
    expect(value['markdownTruncated']).toBe(true)
    expect(value['markdownChars']).toBe(20)
    expect(text).toContain('truncated after 10 of 20 characters')
  })

  it('builds a quoted argv with the configured executable', async () => {
    const harness = mount({ bdcPath: '/opt/my tools/bdc' }, { stdoutText: CONVERT_SUCCESS })
    await callTool(harness, 'doc_convert', { path: '/docs/my report.docx', outputDir: '/out dir', extractImages: true, mermaidScale: 5 })
    expect(harness.commands[0]).toBe(
      '\'/opt/my tools/bdc\' \'convert\' \'/docs/my report.docx\' \'--content\' \'preview\' \'--preview-chars\' \'200000\' \'--output-dir\' \'/out dir\' \'--extract-images\' \'true\' \'--mermaid-scale\' \'5\'',
    )
  })

  it('quotes for PowerShell when configured', async () => {
    const harness = mount({ shellDialect: 'powershell' }, { stdoutText: CONVERT_SUCCESS })
    await callTool(harness, 'doc_convert', { path: '/docs/it\'s.docx' })
    expect(harness.commands[0]).toBe('\'bdc\' \'convert\' \'/docs/it\'\'s.docx\' \'--content\' \'preview\' \'--preview-chars\' \'200000\'')
  })

  it('honors an explicit output dir argument', async () => {
    const harness = mount({}, { stdoutText: CONVERT_SUCCESS })
    await callTool(harness, 'doc_convert', { path: '/docs/report.docx', outputDir: '/tmp' })
    expect(harness.commands[0]).toContain('\'--output-dir\' \'/tmp\'')
  })

  it('surfaces a CLI domain failure with its own code', async () => {
    const harness = mount({}, {
      exitCode: 1,
      stdoutText: JSON.stringify({
        schema_version: '1.0',
        success: false,
        input_path: '/docs/old.doc',
        input_format: 'doc',
        error_code: 'UNSUPPORTED_FORMAT',
        error: 'unsupported format: .doc',
        retryable: false,
        suggestion: 'Convert to .docx first.',
      }),
    })
    const { value, text } = await callTool(harness, 'doc_convert', { path: '/docs/old.doc' })
    expect(value).toMatchObject({ success: false, errorCode: 'UNSUPPORTED_FORMAT', inputPath: '/docs/old.doc', inputFormat: 'doc' })
    expect(text).toContain('Conversion failed (UNSUPPORTED_FORMAT): unsupported format: .doc')
    expect(text).toContain('Suggestion: Convert to .docx first.')
  })

  it('points at doc_setup for a missing dependency', async () => {
    const harness = mount({}, {
      exitCode: 1,
      stdoutText: JSON.stringify({
        schema_version: '1.0',
        success: false,
        input_path: '/docs/notes.md',
        input_format: 'md',
        error_code: 'DEPENDENCY_INSTALL_REQUIRED',
        error: 'Node.js dependencies are not installed',
        retryable: true,
        next_command: 'bdc setup-node',
      }),
    })
    const { text } = await callTool(harness, 'doc_convert', { path: '/docs/notes.md' })
    expect(text).toContain('Call the doc_setup tool, then retry this call.')
    expect(text).not.toContain('Next command')
  })

  it('reports a missing executable distinctly from a CLI failure', async () => {
    const harness = mount({ bdcPath: '/missing/bdc' }, { exitCode: 127, stderrText: '/missing/bdc: command not found' })
    const { value, text } = await callTool(harness, 'doc_convert', { path: '/docs/report.docx' })
    expect(value).toMatchObject({ success: false, errorCode: 'BDC_NOT_FOUND' })
    expect(text).toContain('bdcPath')
  })

  it('reports a timeout with the budget that expired and a retryable suggestion', async () => {
    const harness = mount({ convertTimeoutMs: 5000 }, { timedOut: true, exitCode: null })
    const { value, text } = await callTool(harness, 'doc_convert', { path: '/docs/huge.pdf' })
    expect(value).toMatchObject({ success: false, errorCode: 'BDC_TIMEOUT', retryable: true })
    expect(String(value['error'])).toContain('did not finish within 5000 ms')
    expect(text).toContain('Raise `convertTimeoutMs` in the plugin config')
  })

  it('reports a sandbox denial with the executed mode', async () => {
    const harness = mount({}, {
      exitCode: 1,
      sandbox: { mode: 'workspace-write', denied: true },
    })
    const { value } = await callTool(harness, 'doc_convert', { path: '/docs/report.docx' })
    expect(value).toMatchObject({ success: false, errorCode: 'BDC_SANDBOX_DENIED' })
  })

  it('reports a truncated capture instead of parsing a partial envelope', async () => {
    const harness = mount({ stdoutMaxBytes: 4096 }, { stdoutText: '...tail...', stdoutTruncated: true })
    const { value, text } = await callTool(harness, 'doc_convert', { path: '/docs/report.docx' })
    expect(value).toMatchObject({ success: false, errorCode: 'BDC_OUTPUT_TRUNCATED' })
    expect(String(value['error'])).toContain('4096-byte stdout capture budget')
    expect(text).toContain('Raise `stdoutMaxBytes` in the plugin config')
  })

  it('reports unparseable stdout with the stderr tail', async () => {
    const harness = mount({}, { exitCode: 1, stdoutText: 'Traceback (most recent call last)', stderrText: 'ModuleNotFoundError: docx' })
    const { value } = await callTool(harness, 'doc_convert', { path: '/docs/report.docx' })
    expect(value).toMatchObject({ success: false, errorCode: 'BDC_PROTOCOL_ERROR' })
    expect(String(value['error'])).toContain('ModuleNotFoundError: docx')
  })

  it('rejects invalid arguments before spawning', async () => {
    const harness = mount({ maxMermaidScale: 8 })
    await expect(callTool(harness, 'doc_convert', { path: '   ' })).rejects.toThrow(/invalid path/)
    await expect(callTool(harness, 'doc_convert', { path: '/a.md', mermaidScale: 9 })).rejects.toThrow(/exceeds the configured maximum/)
    await expect(callTool(harness, 'doc_convert', { path: '/a.md', mermaidScale: 0 })).rejects.toThrow(/positive number/)
    expect(harness.commands).toHaveLength(0)
  })

  it('rejects a missing required path at the schema boundary', async () => {
    const harness = mount()
    await expect(callTool(harness, 'doc_convert', {})).rejects.toThrow(/path/)
    expect(harness.commands).toHaveLength(0)
  })

  it('reports a CLI too old for this invocation', async () => {
    const harness = mount({}, {
      exitCode: 1,
      stdoutText: JSON.stringify({
        schema_version: '1.0',
        success: false,
        error_code: 'USAGE_ERROR',
        error: "argument command: invalid choice: 'setup-node' (choose from convert, batch)",
        retryable: false,
      }),
    })
    const { value, text } = await callTool(harness, 'doc_convert', { path: '/docs/report.docx' })
    expect(value).toMatchObject({ success: false, errorCode: 'BDC_CLI_INCOMPATIBLE' })
    expect(text).toContain('pipx upgrade bruce-doc-converter')
  })

  it('throws an abort error when the caller cancels', async () => {
    const harness = mount({}, { aborted: true, exitCode: null })
    await expect(callTool(harness, 'doc_convert', { path: '/docs/a.docx' })).rejects.toMatchObject({ name: 'AbortError' })
  })
})

describe('doc_batch', () => {
  it('recognizes an older CLI rejecting batch flags', async () => {
    const harness = mount({}, { exitCode: 1, stdoutText: JSON.stringify({ success: false, error_code: 'USAGE_ERROR', error: 'unrecognized arguments' }) })
    const { value } = await callTool(harness, 'doc_batch', { path: '/d' })
    expect(value).toMatchObject({ success: false, errorCode: 'BDC_CLI_INCOMPATIBLE' })
  })

  it('keeps the complete manifest and CLI omission count', async () => {
    const harness = mount({}, { stdoutText: JSON.stringify({ success: true, total: 5, succeeded: 5, failed: 0, results: [], omitted: 5, manifest_path: '/out/results.jsonl' }) })
    const { value, text } = await callTool(harness, 'doc_batch', { path: '/d' })
    expect(value).toMatchObject({ manifestPath: '/out/results.jsonl', omitted: 5 })
    expect(text).toContain('Complete manifest: /out/results.jsonl')
  })
  const batchPayload = JSON.stringify({
    schema_version: '1.0',
    success: false,
    total: 3,
    succeeded: 2,
    failed: 1,
    results: [
      { input_path: '/d/a.docx', result: { schema_version: '1.0', success: true, input_path: '/d/a.docx', input_format: 'docx', output_format: 'markdown', output_path: '/d/Markdown/a.md', markdown_content: 'AAA', warnings: [] } },
      { input_path: '/d/b.md', result: { schema_version: '1.0', success: true, input_path: '/d/b.md', input_format: 'md', output_format: 'docx', output_path: '/d/b.docx', warnings: [] } },
      { input_path: '/d/c.pdf', result: { schema_version: '1.0', success: false, input_path: '/d/c.pdf', input_format: 'pdf', error_code: 'EMPTY_PDF_CONTENT', error: 'no text' } },
    ],
  })

  it('summarizes per-file outcomes without Markdown bodies', async () => {
    const harness = mount({}, { exitCode: 1, stdoutText: batchPayload })
    const { value, text } = await callTool(harness, 'doc_batch', { path: '/d' })
    expect(value).toMatchObject({ success: false, directory: '/d', total: 3, succeeded: 2, failed: 1 })
    expect(JSON.stringify(value)).not.toContain('AAA')
    expect(text).toContain('2/3 converted in /d')
    expect(text).toContain('- failed: /d/c.pdf (EMPTY_PDF_CONTENT) no text')
    expect(text).toContain('call doc_convert for the text of one file')
  })

  it('caps the per-file detail and reports the omission', async () => {
    const harness = mount({ maxBatchEntries: 2 }, { stdoutText: batchPayload })
    const { value, text } = await callTool(harness, 'doc_batch', { path: '/d' })
    expect((value['results'] as unknown[])).toHaveLength(2)
    expect(value['omitted']).toBe(1)
    expect(text).toContain('omits 1 outcome(s)')
  })

  it('passes recursive false explicitly', async () => {
    const harness = mount({}, { stdoutText: batchPayload })
    await callTool(harness, 'doc_batch', { path: '/d', recursive: false, mermaidScale: 3 })
    expect(harness.commands[0]).toBe('\'bdc\' \'batch\' \'/d\' \'--content\' \'none\' \'--manifest\' \'auto\' \'--max-results\' \'200\' \'--recursive\' \'false\' \'--mermaid-scale\' \'3\'')
  })

  it('advises per-file conversion when a batch overflows the capture budget', async () => {
    const harness = mount({}, { stdoutText: '...tail...', stdoutTruncated: true })
    const { value, text } = await callTool(harness, 'doc_batch', { path: '/d' })
    expect(value).toMatchObject({ success: false, errorCode: 'BDC_OUTPUT_TRUNCATED' })
    expect(text).toContain('Convert the files one at a time with doc_convert')
  })

  it('reports a failure that produced no batch envelope', async () => {
    const harness = mount({ convertTimeoutMs: 5000 }, { timedOut: true, exitCode: null })
    const { value, text } = await callTool(harness, 'doc_batch', { path: '/d' })
    expect(value).toMatchObject({ success: false, errorCode: 'BDC_TIMEOUT', retryable: true })
    expect(text).toContain('Batch failed (BDC_TIMEOUT)')
    expect(text).toContain('did not finish within 5000 ms')
    expect(text).toContain('Raise `convertTimeoutMs` in the plugin config, or scan a smaller directory')
  })
})

describe('doc_setup', () => {
  it('passes through Node runtime upgrade guidance', async () => {
    const harness = mount({}, { exitCode: 1, stdoutText: JSON.stringify({ success: false,
      error_code: 'NODE_VERSION_UNSUPPORTED', error: 'Node 20 is unsupported', suggestion: 'Install Node.js >=22.19' }) })
    const { value, text } = await callTool(harness, 'doc_setup', {})
    expect(value).toMatchObject({ success: false, errorCode: 'NODE_VERSION_UNSUPPORTED', suggestion: 'Install Node.js >=22.19' })
    expect(text).toContain('Install Node.js >=22.19')
  })
  it('maps an idempotent success', async () => {
    const harness = mount({}, {
      stdoutText: JSON.stringify({
        schema_version: '1.0',
        success: true,
        node_home: '/home/u/.cache/bdc/node',
        allow_scripts: false,
        already_installed: true,
        install_action: 'skipped',
        browser_install_action: 'not_requested',
        detected_browser_path: '/usr/bin/chromium',
      }),
    })
    const { value, text } = await callTool(harness, 'doc_setup', {})
    expect(value).toMatchObject({ success: true, alreadyInstalled: true, nodeHome: '/home/u/.cache/bdc/node' })
    expect(text).toContain('already installed')
    expect(text).toContain('/usr/bin/chromium')
  })

  it('forwards the explicit opt-ins', async () => {
    const harness = mount({}, { stdoutText: JSON.stringify({ schema_version: '1.0', success: true }) })
    await callTool(harness, 'doc_setup', { installBrowser: true, allowScripts: true })
    expect(harness.commands[0]).toBe('\'bdc\' \'setup-node\' \'--allow-scripts\' \'--install-browser\'')
  })

  it('reports the CLI dependency failure', async () => {
    const harness = mount({}, {
      exitCode: 1,
      stdoutText: JSON.stringify({ schema_version: '1.0', success: false, error_code: 'DEPENDENCY_INSTALL_FAILED', error: 'npm ci failed', retryable: false }),
    })
    const { value, text } = await callTool(harness, 'doc_setup', {})
    expect(value).toMatchObject({ success: false, errorCode: 'DEPENDENCY_INSTALL_FAILED' })
    expect(text).toContain('Markdown to Word conversion stays unavailable')
  })

  it('names the setup budget, not the conversion budget, when setup times out', async () => {
    const harness = mount({ setupTimeoutMs: 1234, convertTimeoutMs: 5000 }, { timedOut: true, exitCode: null })
    const { value, text } = await callTool(harness, 'doc_setup', {})
    expect(value).toMatchObject({ success: false, errorCode: 'BDC_TIMEOUT' })
    expect(String(value['error'])).toContain('did not finish within 1234 ms')
    expect(text).toContain('Raise `setupTimeoutMs` in the plugin config')
    expect(text).not.toContain('convertTimeoutMs')
  })

  it('tells the user to upgrade a CLI without setup-node', async () => {
    const harness = mount({}, {
      exitCode: 1,
      stdoutText: JSON.stringify({
        schema_version: '1.0',
        success: false,
        error_code: 'USAGE_ERROR',
        error: "argument command: invalid choice: 'setup-node' (choose from convert, batch)",
        retryable: false,
      }),
    })
    const { value, text } = await callTool(harness, 'doc_setup', {})
    expect(value).toMatchObject({ success: false, errorCode: 'BDC_CLI_INCOMPATIBLE' })
    expect(text).toContain('pipx upgrade bruce-doc-converter')
  })
})

describe('guidance skill', () => {
  it('registers one runtime skill with the tool guidance', async () => {
    const harness = mount()
    expect(harness.skills[0]).toMatchObject({ name: 'bruce-doc-converter', source: 'runtime' })
    expect(String((harness.skills[0] as { content: string }).content)).toContain('DEPENDENCY_INSTALL_REQUIRED')
  })
})
