/**
 * CommonMark parsing with table/strikethrough, math, footnotes and async Mermaid fences.
 * Math, footnote and CJK line rules are ported from bruce-md2word (src/core/markdown.ts,
 * math-markdown.ts, document-rules.ts).
 */
const MarkdownIt = require('markdown-it');
const footnote = require('markdown-it-footnote');
const mermaid = require('./mermaid-renderer');
const { Diagnostics } = require('./diagnostics');

const MAX_FORMULAS = 1000;
const MAX_FOOTNOTES = 1000;
const escapeHTML = (text) => new MarkdownIt().utils.escapeHtml(String(text ?? ''));

// Hangul is excluded: Korean separates words with spaces, so a line break stays a space.
const CJK = /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Bopomofo}　-〿＀-￯]/u;
const CJK_PUNCTUATION = /[　-〿！-／：-＠［-｀｛-･]/u;
const lastChar = (text) => (/[\uD800-\uDBFF][\uDC00-\uDFFF]$/.test(text) ? text.slice(-2) : text.slice(-1));
const firstChar = (text) => String.fromCodePoint(text.codePointAt(0));

/** The adjacent character across emphasis/link boundaries, or undefined at any other inline token. */
function neighbor(tokens, index, step) {
  for (let i = index + step; i >= 0 && i < tokens.length; i += step) {
    const token = tokens[i];
    if (/_(?:open|close)$/.test(token.type)) continue;
    if (token.type !== 'text' && token.type !== 'code_inline') return undefined;
    if (!token.content) continue;
    return step < 0 ? lastChar(token.content) : firstChar(token.content);
  }
  return undefined;
}

/** A source line break between CJK characters is not a word space (CSS Text segment-break rules). */
function joinCjkLines(tokens) {
  for (const token of tokens) {
    if (token.type !== 'inline' || !token.children) continue;
    token.children.forEach((child, i, children) => {
      if (child.type !== 'softbreak') return;
      const before = neighbor(children, i, -1);
      const after = neighbor(children, i, 1);
      if (before && after && ((CJK.test(before) && CJK.test(after)) || CJK_PUNCTUATION.test(before) || CJK_PUNCTUATION.test(after))) {
        child.type = 'text';
        child.content = '';
      }
    });
  }
}

/** Math is tokenized before Markdown escapes/emphasis, never inside code tokens. */
function installMathRules(parser) {
  const escaped = (source, pos) => {
    let slashes = 0;
    while (pos > 0 && source[--pos] === '\\') slashes++;
    return slashes % 2 === 1;
  };
  parser.inline.ruler.before('escape', 'math_inline', (state, silent) => {
    const start = state.pos;
    const source = state.src;
    const opener = source.startsWith('\\(', start) ? '\\('
      : source[start] === '$' && source[start + 1] !== '$' && source[start - 1] !== '$' ? '$' : '';
    if (!opener) return false;
    const dollar = opener === '$';
    if (dollar && (/\s/.test(source[start + 1] ?? ' ') || /\d/.test(source[start - 1] ?? ''))) return false;
    const closer = dollar ? '$' : '\\)';
    let end = start + opener.length;
    for (; end < state.posMax; end++) {
      if (source[end] === '\n' || (dollar && source[end] === '`')) break;
      if (source.startsWith(closer, end) && !escaped(source, end)) break;
    }
    const closed = end < state.posMax && source.startsWith(closer, end)
      && (!dollar || (!/\s/.test(source[end - 1]) && !/[\d$]/.test(source[end + 1] ?? '') && source[end - 1] !== '$'));
    // An unmatched dollar is ordinary currency/prose. Explicit \( is unambiguous.
    if (!closed && dollar) return false;
    const content = source.slice(start + opener.length, end);
    // Avoid treating "$5 and $10" or "$5 to 10$" as mathematical markup.
    if (dollar && /^\d[\d.,]*\s+(?:and|to|至|和)\s*\d[\d.,]*$/.test(content)) return false;
    if (!silent) {
      const token = state.push('math_inline', 'span', 0);
      token.content = content;
      token.meta = {
        raw: source.slice(start, end + (closed ? closer.length : 0)),
        unclosed: !closed,
        lineOffset: source.slice(0, start).split('\n').length - 1
      };
    }
    state.pos = end + (closed ? closer.length : 0);
    return true;
  });
  parser.block.ruler.before('fence', 'math_block', (state, start, end, silent) => {
    if (state.sCount[start] - state.blkIndent >= 4) return false;
    const first = state.src.slice(state.bMarks[start] + state.tShift[start], state.eMarks[start]);
    const opener = first.startsWith('$$') ? '$$' : first.startsWith('\\[') ? '\\[' : '';
    if (!opener) return false;
    if (silent) return true;
    const closer = opener === '$$' ? '$$' : '\\]';
    const lines = [];
    let closed = false;
    let next = start;
    for (; next < end; next++) {
      if (next > start && state.sCount[next] < state.blkIndent && !state.isEmpty(next)) break;
      const line = state.src.slice(state.bMarks[next] + state.tShift[next], state.eMarks[next]);
      const body = next === start ? line.slice(opener.length) : line;
      let at = body.indexOf(closer);
      while (at >= 0 && (escaped(body, at) || body.slice(at + closer.length).trim())) at = body.indexOf(closer, at + closer.length);
      if (at >= 0) {
        lines.push(body.slice(0, at));
        closed = true;
        next++;
        break;
      }
      lines.push(body);
    }
    const token = state.push('math_block', 'p', 0);
    token.block = true;
    token.map = [start, next];
    token.content = lines.join('\n');
    token.meta = { raw: opener + token.content + (closed ? closer : ''), unclosed: !closed };
    state.line = next;
    return true;
  }, { alt: ['paragraph', 'reference', 'blockquote', 'list'] });
  const render = (tokens, index) => {
    const token = tokens[index];
    const id = token.meta.mathId;
    return token.type === 'math_block'
      ? `<p data-math-block="${id}"><span data-math="${id}"></span></p>\n`
      : `<span data-math="${id}"></span>`;
  };
  parser.renderer.rules.math_inline = render;
  parser.renderer.rules.math_block = render;
}

/** Named footnotes become Word footnotes; undefined, duplicate and nested ones are reported. */
function installFootnoteRules(md, diagnostics) {
  md.use(footnote);
  // Named definitions only. Inline ^[...] remains literal Markdown.
  md.inline.ruler.disable('footnote_inline');
  const def = md.block.ruler.getRules('').find(rule => rule.name === 'footnote_def');
  const reference = md.block.ruler.getRules('').find(rule => rule.name === 'reference');
  const labels = new Set();
  md.block.ruler.at('footnote_def', (state, start, end, silent) => {
    if (state.sCount[start] - state.blkIndent >= 4) return false;
    const text = state.src.slice(state.bMarks[start] + state.tShift[start], state.eMarks[start]);
    const label = (/^\[\^([^\]\s]+)\]:/.exec(text) || [])[1];
    if (!label) return false;
    const nested = String(state.parentType) === 'footnote';
    if (labels.has(label) || nested) {
      if (!silent) {
        diagnostics.add(nested ? 'FOOTNOTE_NESTED' : 'FOOTNOTE_DUPLICATE', '重复或嵌套的脚注定义已作为正文保留，已有定义不变。', 'warning', start + 1);
        const token = state.push('footnote_duplicate', '', 0);
        token.content = text;
        token.map = [start, start + 1];
        state.line = start + 1;
      }
      return true;
    }
    if (!silent && labels.size >= MAX_FOOTNOTES) throw new Error(`脚注数量超过 ${MAX_FOOTNOTES} 条上限`);
    const index = state.tokens.length;
    if (!silent) labels.add(label);
    const accepted = def(state, start, end, silent);
    if (accepted && !silent) state.tokens[index].map = [start, state.line];
    return accepted;
  }, { alt: ['paragraph', 'reference'] });
  md.block.ruler.at('reference', (state, start, end, silent) => {
    const text = state.src.slice(state.bMarks[start] + state.tShift[start], state.eMarks[start]);
    return /^\[\^/.test(text) ? false : reference(state, start, end, silent);
  });
  md.inline.ruler.before('footnote_ref', 'missing_footnote', (state, silent) => {
    const match = /^\[\^([^\]\s]+)\]/.exec(state.src.slice(state.pos));
    if (!match || labels.has(match[1])) return false;
    if (!silent) {
      const token = state.push('footnote_missing', '', 0);
      token.content = match[0];
      token.meta = { lineOffset: state.src.slice(0, state.pos).split('\n').length - 1 };
    }
    state.pos += match[0].length;
    return true;
  });
  md.core.ruler.before('footnote_tail', 'preserve_footnote_content', (state) => {
    let definition = false;
    let orphan = false;
    let line;
    for (const token of state.tokens) {
      if (token.type === 'footnote_reference_open') {
        definition = true;
        line = token.map ? token.map[0] + 1 : undefined;
        const refs = state.env.footnotes && state.env.footnotes.refs;
        orphan = !!refs && refs[`:${token.meta.label}`] === -1;
        if (orphan) {
          diagnostics.add('FOOTNOTE_UNUSED', '脚注定义未被正文引用，已作为正文保留。', 'info', line);
          token.type = 'footnote_unused_open';
          token.tag = 'div';
        }
      } else if (token.type === 'footnote_reference_close') {
        if (orphan) {
          token.type = 'footnote_unused_close';
          token.tag = 'div';
        }
        definition = false;
        orphan = false;
      }
      if (definition) {
        for (const child of token.children || []) {
          if (child.type === 'footnote_ref') {
            child.type = 'text';
            child.content = `[^${child.meta.label}]`;
            diagnostics.add('FOOTNOTE_NESTED', '脚注内的脚注引用不受支持，已保留为文本。', 'warning', line);
          }
        }
      }
    }
  });
  md.renderer.rules.footnote_missing = (tokens, i) => md.utils.escapeHtml(tokens[i].content);
  md.renderer.rules.footnote_duplicate = (tokens, i) => `<p>${md.utils.escapeHtml(tokens[i].content)}</p>\n`;
  md.renderer.rules.footnote_unused_open = (tokens, i) => `<div><p>${md.utils.escapeHtml(`[^${tokens[i].meta.label}]:`)}</p>`;
  md.renderer.rules.footnote_unused_close = () => '</div>\n';
  md.renderer.rules.footnote_ref = (tokens, i) => `<span data-footnote-ref="${tokens[i].meta.id + 1}"></span>`;
  md.renderer.rules.footnote_block_open = () => '<section data-footnotes="true">\n';
  md.renderer.rules.footnote_block_close = () => '</section>\n';
  md.renderer.rules.footnote_open = (tokens, i) => `<div data-footnote-id="${tokens[i].meta.id + 1}">`;
  md.renderer.rules.footnote_close = () => '</div>\n';
  md.renderer.rules.footnote_anchor = () => '';
}

function createParser(diagnostics) {
  const parser = new MarkdownIt({ html: false, linkify: false, typographer: false });
  installMathRules(parser);
  installFootnoteRules(parser, diagnostics);
  // Runs after footnote_tail so footnote bodies are joined too.
  parser.core.ruler.push('cjk_line_join', state => joinCjkLines(state.tokens));
  // Remove only the fence terminator's newline, retaining intentional blank lines.
  parser.renderer.rules.fence = (tokens, index) => {
    const token = tokens[index];
    if (token.meta && token.meta.mermaidHTML) return token.meta.mermaidHTML;
    const language = token.info.trim().split(/\s+/)[0].replace(/[^a-zA-Z0-9_-]/g, '');
    const attribute = language ? ` class="language-${language}"` : '';
    const code = `<pre><code${attribute}>${parser.utils.escapeHtml(token.content.replace(/\n$/, ''))}</code></pre>\n`;
    if (token.meta && token.meta.mermaidFailed) {
      const where = token.meta.line ? `（源文件第 ${token.meta.line} 行）` : '';
      return `<p data-mermaid-notice="true">Mermaid 图表未渲染${where}，以下保留原始代码。</p>\n${code}`;
    }
    return code;
  };
  return parser;
}

/**
 * @param {string} markdown
 * @param {Diagnostics} [diagnostics] shared with the HTML -> DOCX stage
 * @returns {Promise<{html: string, warnings: string[], diagnostics: Diagnostics, formulas: object[]}>}
 */
async function markdownToHTML(markdown, diagnostics = new Diagnostics()) {
  const formulas = [];
  if (!markdown || typeof markdown !== 'string') return { html: '', warnings: [], diagnostics, formulas };
  const parser = createParser(diagnostics);
  const env = {};
  const tokens = parser.parse(markdown, env);

  const walk = (items, inheritedLine) => {
    let currentLine = inheritedLine;
    for (const token of items) {
      if (token.map) currentLine = token.map[0] + 1;
      const offset = (token.meta && token.meta.lineOffset) || 0;
      const line = currentLine === undefined ? undefined : currentLine + offset;
      if (token.type === 'footnote_missing') diagnostics.add('FOOTNOTE_UNDEFINED', `脚注 ${token.content} 未定义，已保留原文。`, 'warning', line);
      if (token.type === 'math_inline' || token.type === 'math_block') {
        if (formulas.length >= MAX_FORMULAS) throw new Error(`公式数量超过 ${MAX_FORMULAS} 个上限`);
        const id = `math-${formulas.length}`;
        formulas.push({ id, source: token.content, raw: token.meta.raw, display: token.type === 'math_block', unclosed: token.meta.unclosed, line });
        token.meta.mathId = id;
      }
      if (token.type === 'image') {
        // Alt text is plain description: markdown-it's renderInlineAsText drops custom tokens.
        for (const child of token.children || []) {
          if (child.type === 'math_inline') {
            child.type = 'text';
            child.content = child.meta.raw;
          } else if (child.type === 'footnote_missing') {
            diagnostics.add('FOOTNOTE_UNDEFINED', `图片说明中的脚注 ${child.content} 未定义，已保留原文。`, 'warning', line);
            child.type = 'text';
          } else if (child.type === 'footnote_ref') {
            child.type = 'text';
            child.content = `[^${child.meta.label}]`;
          }
        }
        if (line !== undefined) token.attrSet('data-line', String(line));
      }
      // Only in-document links can fail later (missing heading); record where they come from.
      if (token.type === 'link_open' && line !== undefined && (token.attrGet('href') || '').startsWith('#')) token.attrSet('data-line', String(line));
      if (token.children && token.type !== 'image') walk(token.children, line);
    }
  };
  walk(tokens);

  for (const token of tokens) {
    if (token.type !== 'fence' || token.info.trim().toLowerCase() !== 'mermaid') continue;
    const line = token.map ? token.map[0] + 1 : undefined;
    const rendered = await mermaid.renderMermaidToDataUrl(token.content.replace(/\n$/, ''));
    if (!rendered.success) {
      diagnostics.add('MERMAID_NOT_RENDERED', `Mermaid 渲染失败: ${rendered.error || '未知错误'}`, 'warning', line);
      token.meta = { mermaidFailed: true, line };
      continue;
    }
    const width = Number.isFinite(rendered.width) && rendered.width > 0 ? ` width="${rendered.width}"` : '';
    const height = Number.isFinite(rendered.height) && rendered.height > 0 ? ` height="${rendered.height}"` : '';
    const lineAttr = line ? ` data-line="${line}"` : '';
    token.meta = { mermaidHTML: `<p><img src="${escapeHTML(rendered.dataUrl)}" alt="Mermaid Diagram"${width}${height}${lineAttr}></p>\n` };
  }

  const html = parser.renderer.render(tokens, parser.options, env).trimEnd();
  return { html, warnings: diagnostics.warnings(), diagnostics, formulas };
}

module.exports = { markdownToHTML, escapeHTML };
