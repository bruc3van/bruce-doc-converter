/** CommonMark parsing with table/strikethrough extensions and async Mermaid fences. */
const MarkdownIt = require('markdown-it');
const mermaid = require('./mermaid-renderer');

const parser = new MarkdownIt({ html: false, linkify: false, typographer: false });
const escapeHTML = (text) => parser.utils.escapeHtml(String(text ?? ''));

// Remove only the fence terminator's newline, retaining intentional blank lines.
parser.renderer.rules.fence = (tokens, index) => {
  const token = tokens[index];
  if (token.meta?.mermaidHTML) return token.meta.mermaidHTML;
  const language = token.info.trim().split(/\s+/)[0].replace(/[^a-zA-Z0-9_-]/g, '');
  const attribute = language ? ` class="language-${language}"` : '';
  return `<pre><code${attribute}>${escapeHTML(token.content.replace(/\n$/, ''))}</code></pre>\n`;
};

async function markdownToHTML(markdown) {
  const warnings = [];
  if (!markdown || typeof markdown !== 'string') return { html: '', warnings };
  const env = {};
  const tokens = parser.parse(markdown, env);
  for (const token of tokens) {
    if (token.type !== 'fence' || token.info.trim().toLowerCase() !== 'mermaid') continue;
    const rendered = await mermaid.renderMermaidToDataUrl(token.content.replace(/\n$/, ''));
    if (!rendered.success) {
      warnings.push(`Mermaid 渲染失败: ${rendered.error || '未知错误'}`);
      continue;
    }
    const width = Number.isFinite(rendered.width) && rendered.width > 0 ? ` width="${rendered.width}"` : '';
    const height = Number.isFinite(rendered.height) && rendered.height > 0 ? ` height="${rendered.height}"` : '';
    token.meta = { mermaidHTML: `<div><img src="${escapeHTML(rendered.dataUrl)}" alt="Mermaid Diagram"${width}${height}></div>\n` };
  }
  return { html: parser.renderer.render(tokens, parser.options, env).trimEnd(), warnings };
}

module.exports = { markdownToHTML, escapeHTML };
