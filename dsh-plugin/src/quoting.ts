/**
 * Shell quoting for argv elements. `ctx.shell` takes one command string per
 * call, so every plugin-owned argv element is quoted for the dialect that the
 * mounted provider interprets; nothing here is model-authored text.
 * @module bruce-doc-converter-dsh/quoting
 */

import type { ShellDialect } from './types.ts'

/**
 * Quote one argv element for the given shell dialect.
 *
 * POSIX shells use single quotes, where a literal single quote is spelled by
 * closing the quote, escaping the quote, and reopening it. PowerShell uses
 * single-quoted strings, where a literal single quote is doubled. Every element
 * is quoted unconditionally, so a path never changes the command structure.
 * @param dialect - the dialect the mounted shell provider interprets.
 * @param value - the argv element to quote.
 * @returns the quoted argv element.
 * @throws when the value contains a NUL byte, which no shell can carry.
 */
export function quoteArgument(dialect: ShellDialect, value: string): string {
  if (value.includes('\0')) {
    throw new Error('cannot quote an argument containing a NUL byte')
  }
  const escaped = dialect === 'posix'
    ? value.replaceAll('\'', '\'\\\'\'')
    : value.replaceAll('\'', '\'\'')
  return `'${escaped}'`
}

/**
 * Join an argv into the single command string `ctx.shell` accepts.
 * @param dialect - the dialect the mounted shell provider interprets.
 * @param argv - executable and arguments, `argv[0]` first.
 * @returns the quoted command line.
 */
export function buildCommandLine(dialect: ShellDialect, argv: readonly string[]): string {
  return argv.map((part) => quoteArgument(dialect, part)).join(' ')
}
