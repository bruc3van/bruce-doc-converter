import { describe, expect, it } from 'vitest'
import { buildCommandLine, quoteArgument } from '../src/quoting.ts'

describe('quoteArgument', () => {
  it('quotes a POSIX element unconditionally', () => {
    expect(quoteArgument('posix', 'bdc')).toBe('\'bdc\'')
    expect(quoteArgument('posix', '/tmp/my notes/a.docx')).toBe('\'/tmp/my notes/a.docx\'')
  })

  it('escapes both quote characters differently per dialect', () => {
    expect(quoteArgument('posix', "it's")).toBe('\'it\'\\\'\'s\'')
    expect(quoteArgument('powershell', "it's")).toBe('\'it\'\'s\'')
  })

  it('keeps non-ASCII paths intact', () => {
    expect(quoteArgument('posix', '/tmp/文档/报告.docx')).toBe('\'/tmp/文档/报告.docx\'')
  })

  it('rejects a NUL byte', () => {
    expect(() => quoteArgument('posix', 'a\0b')).toThrow(/NUL/)
  })
})

describe('buildCommandLine', () => {
  it('joins every element quoted, in order', () => {
    expect(buildCommandLine('posix', ['bdc', 'convert', '/a b/c.docx', '--mermaid-scale', '4']))
      .toBe('\'bdc\' \'convert\' \'/a b/c.docx\' \'--mermaid-scale\' \'4\'')
  })

  it('never lets an argument change the command structure', () => {
    expect(buildCommandLine('posix', ['bdc', 'convert', '; rm -rf /']))
      .toBe('\'bdc\' \'convert\' \'; rm -rf /\'')
  })
})
