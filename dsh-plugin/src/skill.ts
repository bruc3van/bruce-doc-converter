/**
 * The optional embedded skill. The plugin's tools own the mechanics of a
 * conversion; this skill carries only the judgment a model needs before
 * choosing them, and it appears whenever a skill registry is mounted.
 * @module bruce-doc-converter-dsh/skill
 */

import type { Context } from '@deepseek-ai/cordis'
import type {} from '@deepseek-ai/dsh-skill'

/** Skill name used for model and `/name` invocation. */
export const GUIDANCE_SKILL_NAME = 'bruce-doc-converter'

/** Routing description shown in the session skill catalog. */
export const GUIDANCE_SKILL_DESCRIPTION =
  'Convert or export documents with the doc_convert, doc_batch, and doc_setup tools (.docx, .xlsx, .pptx, .pdf to Markdown; Markdown to Word). '
  + 'Use it when a task involves reading, converting, or exporting an Office, PDF, or Markdown document.'

/** Task-level guidance that no tool schema can carry. */
export const GUIDANCE_SKILL_CONTENT = `# Bruce Doc Converter

The \`doc_convert\`, \`doc_batch\`, and \`doc_setup\` tools wrap the local \`bdc\` CLI. Use this guidance to choose and sequence them.

## Choosing the call

- One document: \`doc_convert\`.
- A directory of documents: \`doc_batch\`. It returns per-file paths and errors, never Markdown bodies, so follow up with \`doc_convert\` for the text of a file you need to read.
- Office and PDF input (\`.docx\`, \`.xlsx\`, \`.pptx\`, \`.pdf\`) produces Markdown and returns its text directly: read the tool result instead of opening the generated file.
- Markdown input produces a Word file. Inspect the result's \`outputPath\` and hand that path to the user.

## Before converting

- Only modern formats are accepted. Legacy \`.doc\`, \`.xls\`, and \`.ppt\` files must be converted to the modern format first; the tool reports \`UNSUPPORTED_FORMAT\` otherwise.
- Scanned PDFs carry no extractable text. The tool reports \`EMPTY_PDF_CONTENT\`; run OCR on the file before retrying, and do not repeat the same call.
- Inputs above 100 MB are rejected as \`FILE_TOO_LARGE\`. Prefer files under 50 MB; split or compress larger ones.

## When Markdown to Word fails

A fresh machine has no Node.js dependencies for Markdown-to-Word conversion. When a conversion reports \`DEPENDENCY_INSTALL_REQUIRED\`, call \`doc_setup\` once and retry the same conversion. \`doc_setup\` is idempotent and reports \`alreadyInstalled\` when the shared directory already matches.

Mermaid diagrams become PNG images. The CLI uses a local Chrome, Edge, or Chromium headlessly with a temporary profile. Call \`doc_setup\` with \`installBrowser: true\` only when no local browser exists, and set \`allowScripts\` only for a source you trust.

Expect quality to vary with the input: \`.docx\` and \`.xlsx\` convert well, \`.pptx\` reasonably, and PDF quality depends on whether it is text-based.
`

/**
 * Register the guidance skill when the composition has a skill registry. A
 * composition without `ctx.skills` simply never receives it.
 * @param ctx - plugin context whose optional `ctx.skills` receives the skill.
 */
export function registerGuidanceSkill(ctx: Context): void {
  ctx.inject(['skills'], (skillsCtx) => {
    skillsCtx.effect(
      () => skillsCtx.skills.register({
        name: GUIDANCE_SKILL_NAME,
        description: GUIDANCE_SKILL_DESCRIPTION,
        source: 'runtime',
        content: GUIDANCE_SKILL_CONTENT,
      }),
      'bruce-doc-converter guidance skill',
    )
  })
}
