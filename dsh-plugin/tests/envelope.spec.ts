import { describe, expect, it } from 'vitest'
import { CliProtocolError, parseCliJson, readBatchEnvelope, readConvertEnvelope, readSetupEnvelope } from '../src/envelope.ts'

describe('parseCliJson', () => {
  it('rejects empty stdout', () => {
    expect(() => parseCliJson('   ')).toThrow(CliProtocolError)
  })

  it('rejects non-JSON stdout', () => {
    expect(() => parseCliJson('bdc: command not found')).toThrow(/not JSON/)
  })
})

describe('readConvertEnvelope', () => {
  it('reads a success payload', () => {
    const envelope = readConvertEnvelope({
      schema_version: '1.0',
      success: true,
      input_path: '/a/in.docx',
      input_format: 'docx',
      output_format: 'markdown',
      output_path: '/a/Markdown/in.md',
      markdown_content: '# hello',
      extracted_images: ['/a/img1.png'],
      warnings: ['w1'],
    })
    expect(envelope).toEqual({
      schemaVersion: '1.0',
      success: true,
      inputPath: '/a/in.docx',
      inputFormat: 'docx',
      outputFormat: 'markdown',
      outputPath: '/a/Markdown/in.md',
      markdownContent: '# hello',
      extractedImages: ['/a/img1.png'],
      warnings: ['w1'],
    })
  })

  it('reads a failure payload and keeps the CLI classification', () => {
    const envelope = readConvertEnvelope({
      schema_version: '1.0',
      success: false,
      input_path: '/a/in.doc',
      input_format: 'doc',
      error_code: 'UNSUPPORTED_FORMAT',
      error: 'unsupported',
      retryable: false,
      suggestion: 'convert first',
      next_command: 'bdc setup-node',
    })
    expect(envelope.success).toBe(false)
    if (envelope.success) throw new Error('unreachable')
    expect(envelope.errorCode).toBe('UNSUPPORTED_FORMAT')
    expect(envelope.nextCommand).toBe('bdc setup-node')
  })

  it('rejects a mistyped consumed field', () => {
    expect(() => readConvertEnvelope({ success: 'yes' })).toThrow(/success must be a boolean/)
    expect(() => readConvertEnvelope({ success: false, error_code: 'X' })).toThrow(/error must be a string/)
    expect(() => readConvertEnvelope({
      schema_version: '1.0',
      success: true,
      input_path: '/a',
      input_format: 'docx',
      output_format: 'markdown',
      output_path: '/b',
      warnings: 'none',
    })).toThrow(/warnings must be an array/)
  })

  it('defaults an absent warnings list to empty', () => {
    const envelope = readConvertEnvelope({
      schema_version: '1.0',
      success: true,
      input_path: '/a/in.md',
      input_format: 'md',
      output_format: 'docx',
      output_path: '/a/in.docx',
    })
    expect(envelope.success && envelope.warnings).toEqual([])
  })

  it('ignores additive fields the CLI may grow later', () => {
    const envelope = readConvertEnvelope({
      schema_version: '2.0',
      success: true,
      input_path: '/a/in.md',
      input_format: 'md',
      output_format: 'docx',
      output_path: '/a/in.docx',
      warnings: [],
      future_field: { nested: true },
    })
    expect(envelope.success).toBe(true)
  })
})

describe('readBatchEnvelope', () => {
  it('reads per-file results', () => {
    const envelope = readBatchEnvelope({
      schema_version: '1.0',
      success: false,
      total: 2,
      succeeded: 1,
      failed: 1,
      results: [
        { input_path: '/d/a.docx', result: { schema_version: '1.0', success: true, input_path: '/d/a.docx', input_format: 'docx', output_format: 'markdown', output_path: '/d/Markdown/a.md', warnings: [] } },
        { input_path: '/d/b.pdf', result: { schema_version: '1.0', success: false, input_path: '/d/b.pdf', input_format: 'pdf', error_code: 'EMPTY_PDF_CONTENT', error: 'no text' } },
      ],
    })
    expect(envelope.total).toBe(2)
    expect(envelope.results[1]?.result.success).toBe(false)
  })

  it('rejects a missing results array', () => {
    expect(() => readBatchEnvelope({ success: true, total: 0, succeeded: 0, failed: 0 })).toThrow(/results must be an array/)
  })
})

describe('readSetupEnvelope', () => {
  it('reads a success payload', () => {
    expect(readSetupEnvelope({
      schema_version: '1.0',
      success: true,
      node_home: '/home/u/.bdc/node',
      already_installed: true,
      install_action: 'skipped',
      browser_install_action: 'not_requested',
      detected_browser_path: '/usr/bin/chromium',
    })).toEqual({
      success: true,
      nodeHome: '/home/u/.bdc/node',
      alreadyInstalled: true,
      installAction: 'skipped',
      browserInstallAction: 'not_requested',
      detectedBrowserPath: '/usr/bin/chromium',
    })
  })

  it('reads a failure payload', () => {
    expect(readSetupEnvelope({ schema_version: '1.0', success: false, error_code: 'DEPENDENCY_INSTALL_FAILED', error: 'npm failed', retryable: false }))
      .toEqual({ success: false, errorCode: 'DEPENDENCY_INSTALL_FAILED', error: 'npm failed', retryable: false })
  })
})
