/**
 * Markdown 到 HTML 转换器
 * 简化版 Markdown 解析器，支持常见语法
 */

const { renderMermaidToDataUrl } = require('./mermaid-renderer');

/**
 * 将 Markdown 文本转换为 HTML
 * @param {string} markdown - Markdown 文本
 * @returns {Promise<string>} HTML 字符串
 */
async function markdownToHTML(markdown) {
  if (!markdown || typeof markdown !== 'string') {
    return '';
  }

  // 统一换行符，避免不同平台导致的列表解析失败
  let html = markdown.replace(/\r\n?/g, '\n');

  // 1. 首先保护代码块
  const codeBlocks = [];
  html = html.replace(/```([^\n`]*)[ \t]*\n([\s\S]*?)```/g, (match, lang, code) => {
    const placeholder = `\x00CODEBLOCK${codeBlocks.length}\x00`;
    codeBlocks.push({ lang: lang || '', code: normalizeCodeBlockContent(code) });
    return placeholder;
  });

  // 2. 保护行内代码
  const inlineCodes = [];
  html = html.replace(/`([^`]+)`/g, (match, code) => {
    const placeholder = `\x00INLINECODE${inlineCodes.length}\x00`;
    inlineCodes.push(code);
    return placeholder;
  });

  // 3. 保护图片和链接语法（避免路径中的下划线/星号被误处理为粗体斜体）
  const protectedLinks = protectMarkdownLinksAndImages(html);
  html = protectedLinks.text;
  const images = protectedLinks.images;
  const links = protectedLinks.links;

  // 处理标题
  html = html.replace(/^######\s+(.+)$/gm, '<h6>$1</h6>');
  html = html.replace(/^#####\s+(.+)$/gm, '<h5>$1</h5>');
  html = html.replace(/^####\s+(.+)$/gm, '<h4>$1</h4>');
  html = html.replace(/^###\s+(.+)$/gm, '<h3>$1</h3>');
  html = html.replace(/^##\s+(.+)$/gm, '<h2>$1</h2>');
  html = html.replace(/^#\s+(.+)$/gm, '<h1>$1</h1>');

  // 处理粗体和斜体
  html = html.replace(/\*\*\*(.+?)\*\*\*/g, '<strong><em>$1</em></strong>');
  html = html.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
  html = html.replace(/\*(.+?)\*/g, '<em>$1</em>');
  html = html.replace(/___(.+?)___/g, '<strong><em>$1</em></strong>');
  html = html.replace(/__(.+?)__/g, '<strong>$1</strong>');
  html = html.replace(/_(.+?)_/g, '<em>$1</em>');

  // 处理删除线
  html = html.replace(/~~(.+?)~~/g, '<del>$1</del>');

  // 处理引用块（将连续的 > 行合并为一个 blockquote，每行内容包裹在 <p> 中）
  html = processBlockquotes(html);

  // 处理列表
  html = processLists(html);

  // 处理表格
  html = processTables(html);

  // 处理水平线
  html = html.replace(/^(-{3,}|_{3,}|\*{3,})$/gm, '<hr>');

  // 恢复图片（alt 文本中的粗体斜体已在前面处理）
  images.forEach((img, i) => {
    html = html.replace(`\x00IMAGE${i}\x00`, () => `<img src="${img.url}" alt="${img.alt}">`);
  });

  // 恢复链接
  links.forEach((link, i) => {
    html = html.replace(`\x00LINK${i}\x00`, () => `<a href="${link.url}">${link.text}</a>`);
  });

  // 恢复行内代码
  inlineCodes.forEach((code, i) => {
    html = html.replace(`\x00INLINECODE${i}\x00`, () => `<code>${escapeHTML(code)}</code>`);
  });

  // 恢复代码块
  for (let i = 0; i < codeBlocks.length; i++) {
    const block = codeBlocks[i];
    const language = (block.lang || '').trim().toLowerCase();

    if (language === 'mermaid') {
      const rendered = await renderMermaidToDataUrl(block.code);
      if (rendered.success) {
        const widthAttr = Number.isFinite(rendered.width) && rendered.width > 0 ? ` width="${rendered.width}"` : '';
        const heightAttr = Number.isFinite(rendered.height) && rendered.height > 0 ? ` height="${rendered.height}"` : '';
        html = html.replace(
          `\x00CODEBLOCK${i}\x00`,
          () => `<div><img src="${rendered.dataUrl}" alt="Mermaid Diagram"${widthAttr}${heightAttr}></div>`
        );
      } else {
        html = html.replace(
          `\x00CODEBLOCK${i}\x00`,
          () => `<pre><code class="language-mermaid">${escapeHTML(block.code)}</code></pre>`
        );
      }
      continue;
    }

    const safeLang = (block.lang || '').replace(/[^a-zA-Z0-9_-]/g, '');
    html = html.replace(
      `\x00CODEBLOCK${i}\x00`,
      () => `<pre><code class="language-${safeLang}">${escapeHTML(block.code)}</code></pre>`
    );
  }

  // 处理段落（放在最后，避免把代码块/图表包进段落）
  html = processParagraphs(html);

  return html;
}

/**
 * 转义 HTML 特殊字符
 */
function escapeHTML(text) {
  const map = {
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#039;'
  };
  return text.replace(/[&<>"']/g, char => map[char]);
}

function normalizeCodeBlockContent(code) {
  if (typeof code !== 'string') {
    return '';
  }
  return code.replace(/\r\n?/g, '\n').replace(/\n$/, '');
}

function protectMarkdownLinksAndImages(markdown) {
  const images = [];
  const links = [];
  let result = '';

  for (let i = 0; i < markdown.length;) {
    const imageToken = markdown[i] === '!' ? parseMarkdownLinkToken(markdown, i, true) : null;
    if (imageToken) {
      const placeholder = `\x00IMAGE${images.length}\x00`;
      images.push({ alt: imageToken.label, url: imageToken.url });
      result += placeholder;
      i = imageToken.end + 1;
      continue;
    }

    const linkToken = markdown[i] === '[' ? parseMarkdownLinkToken(markdown, i, false) : null;
    if (linkToken) {
      const placeholder = `\x00LINK${links.length}\x00`;
      links.push({ text: linkToken.label, url: linkToken.url });
      result += placeholder;
      i = linkToken.end + 1;
      continue;
    }

    result += markdown[i];
    i += 1;
  }

  return { text: result, images, links };
}

function parseMarkdownLinkToken(source, startIndex, isImage) {
  let cursor = startIndex;
  if (isImage) {
    if (source[cursor] !== '!' || source[cursor + 1] !== '[') {
      return null;
    }
    cursor += 1;
  }

  if (source[cursor] !== '[') {
    return null;
  }

  const labelSection = readBalancedSection(source, cursor, '[', ']');
  if (!labelSection) {
    return null;
  }

  cursor = labelSection.end + 1;
  while (cursor < source.length && /\s/.test(source[cursor]) && source[cursor] !== '\n') {
    cursor += 1;
  }

  if (source[cursor] !== '(') {
    return null;
  }

  const destinationSection = readBalancedSection(source, cursor, '(', ')');
  if (!destinationSection) {
    return null;
  }

  const url = extractLinkDestination(destinationSection.value);
  if (!url) {
    return null;
  }

  return {
    label: labelSection.value,
    url,
    end: destinationSection.end
  };
}

function readBalancedSection(source, startIndex, openChar, closeChar) {
  if (source[startIndex] !== openChar) {
    return null;
  }

  let depth = 0;
  let value = '';
  let escaped = false;

  for (let i = startIndex + 1; i < source.length; i++) {
    const char = source[i];

    if (escaped) {
      value += char;
      escaped = false;
      continue;
    }

    if (char === '\\') {
      escaped = true;
      value += char;
      continue;
    }

    if (char === openChar) {
      depth += 1;
      value += char;
      continue;
    }

    if (char === closeChar) {
      if (depth === 0) {
        return { value, end: i };
      }
      depth -= 1;
      value += char;
      continue;
    }

    value += char;
  }

  return null;
}

function extractLinkDestination(rawValue) {
  const trimmed = (rawValue || '').trim();
  if (!trimmed) {
    return '';
  }

  if (trimmed.startsWith('<')) {
    const endIndex = trimmed.indexOf('>');
    if (endIndex > 1) {
      return trimmed.slice(1, endIndex);
    }
  }

  let destination = '';
  let escaped = false;
  let nestedParens = 0;

  for (let i = 0; i < trimmed.length; i++) {
    const char = trimmed[i];

    if (escaped) {
      destination += char;
      escaped = false;
      continue;
    }

    if (char === '\\') {
      escaped = true;
      destination += char;
      continue;
    }

    if (/\s/.test(char) && nestedParens === 0) {
      break;
    }

    if (char === '(') {
      nestedParens += 1;
    } else if (char === ')' && nestedParens > 0) {
      nestedParens -= 1;
    }

    destination += char;
  }

  return destination.trim();
}

/**
 * 处理列表
 */
function processLists(html) {
  const lines = html.split('\n');
  const result = [];
  const stack = [];

  const countLeadingSpaces = (s) => {
    let count = 0;
    for (let i = 0; i < s.length; i++) {
      if (s[i] === ' ') count++;
      else if (s[i] === '\t') count += 4;
      else break;
    }
    return count;
  };

  // 将原始空格数规范化为离散层级（每2个空格算一级），
  // 避免2空格/3空格/4空格混用导致层级判断错误
  const normalizeIndent = (spaces) => Math.floor(spaces / 2);

  const closeToIndent = (targetIndent) => {
    while (stack.length > 0 && targetIndent < stack[stack.length - 1].indent) {
      const top = stack[stack.length - 1];
      if (top.hasOpenLi) {
        result.push('</li>');
        top.hasOpenLi = false;
      }
      result.push(`</${top.type}>`);
      stack.pop();
    }
  };

  const closeAll = () => closeToIndent(-1);

  const ensureList = (type, indent) => {
    if (stack.length === 0) {
      result.push(`<${type}>`);
      stack.push({ type, indent, hasOpenLi: false });
      return;
    }

    const top = stack[stack.length - 1];
    if (indent > top.indent) {
      result.push(`<${type}>`);
      stack.push({ type, indent, hasOpenLi: false });
      return;
    }

    if (indent === top.indent && type !== top.type) {
      if (top.hasOpenLi) {
        result.push('</li>');
        top.hasOpenLi = false;
      }
      result.push(`</${top.type}>`);
      stack.pop();
      result.push(`<${type}>`);
      stack.push({ type, indent, hasOpenLi: false });
    }
  };

  for (const line of lines) {
    const ulMatch = line.match(/^(\s*)[-*+]\s+(.+)$/);
    const olMatch = line.match(/^(\s*)\d+\.\s+(.+)$/);

    if (!ulMatch && !olMatch) {
      if (stack.length > 0) {
        const trimmedLine = line.trim();
        if (trimmedLine === '') {
          // 列表内空行：保持列表不关闭
          continue;
        }

        const lineIndent = normalizeIndent(countLeadingSpaces(line));

        // 如果是从更深层级缩进回退，先关闭嵌套列表
        while (stack.length > 1 && lineIndent <= stack[stack.length - 1].indent) {
          const top = stack[stack.length - 1];
          if (top.hasOpenLi) {
            result.push('</li>');
            top.hasOpenLi = false;
          }
          result.push(`</${top.type}>`);
          stack.pop();
        }

        if (stack.length > 0) {
          const current = stack[stack.length - 1];
          if (lineIndent > current.indent && current.hasOpenLi) {
            // 视为当前列表项的续行，避免打断嵌套结构
            result.push(`<br>${trimmedLine}`);
            continue;
          }
        }
      }

      closeAll();
      result.push(line);
      continue;
    }

    const indentStr = (ulMatch || olMatch)[1] || '';
    const indent = normalizeIndent(countLeadingSpaces(indentStr));
    const type = ulMatch ? 'ul' : 'ol';
    const itemText = (ulMatch || olMatch)[2];

    closeToIndent(indent);
    ensureList(type, indent);

    const current = stack[stack.length - 1];
    if (current.hasOpenLi) {
      result.push('</li>');
      current.hasOpenLi = false;
    }

    result.push(`<li>${itemText}`);
    current.hasOpenLi = true;
  }

  closeAll();
  return result.join('\n');
}

/**
 * 处理表格
 */
function processTables(html) {
  const lines = html.split('\n');
  const result = [];
  let inTable = false;
  let tableRows = [];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();

    if (line.startsWith('|') && line.endsWith('|')) {
      if (/^\|[\s:|-]+\|$/.test(line) && line.includes('-')) {
        continue;
      }

      if (!inTable) {
        inTable = true;
      }

      tableRows.push(splitMarkdownTableRow(line));
    } else {
      if (inTable) {
        result.push(buildTable(tableRows));
        tableRows = [];
        inTable = false;
      }
      result.push(line);
    }
  }

  if (inTable) {
    result.push(buildTable(tableRows));
  }

  return result.join('\n');
}

function buildTable(rows) {
  if (rows.length === 0) return '';

  const maxCols = Math.max(...rows.map(row => row.length));
  let html = '<table>\n';

  if (rows.length > 0) {
    html += '<thead>\n<tr>\n';
    const headerRow = rows[0];
    for (let i = 0; i < maxCols; i++) {
      const cell = i < headerRow.length ? headerRow[i] : '';
      html += `<th>${cell}</th>\n`;
    }
    html += '</tr>\n</thead>\n';
  }

  if (rows.length > 1) {
    html += '<tbody>\n';
    for (let i = 1; i < rows.length; i++) {
      html += '<tr>\n';
      const bodyRow = rows[i];
      for (let j = 0; j < maxCols; j++) {
        const cell = j < bodyRow.length ? bodyRow[j] : '';
        html += `<td>${cell}</td>\n`;
      }
      html += '</tr>\n';
    }
    html += '</tbody>\n';
  }

  html += '</table>';
  return html;
}

function splitMarkdownTableRow(line) {
  const content = line.slice(1, -1);
  const cells = [];
  let current = '';

  for (let i = 0; i < content.length; i++) {
    const char = content[i];
    const next = content[i + 1];

    if (char === '\\' && (next === '|' || next === '\\')) {
      current += next;
      i += 1;
      continue;
    }

    if (char === '|') {
      cells.push(current.trim());
      current = '';
      continue;
    }

    current += char;
  }

  cells.push(current.trim());
  return cells;
}

/**
 * 处理段落
 */
function processParagraphs(html) {
  const lines = html.split('\n');
  const result = [];
  let paragraph = [];
  let inPreBlock = false;

  const blockElements = ['<h1', '<h2', '<h3', '<h4', '<h5', '<h6', '<ul', '<ol', '<li', '<table', '<thead', '<tbody', '<tr', '<th', '<td', '<blockquote', '<hr', '<div', '</li', '</ul', '</ol', '</table', '</thead', '</tbody', '</tr', '</blockquote', '</div'];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const trimmedLine = line.trim();
    const startsPre = trimmedLine.startsWith('<pre');
    const endsPre = trimmedLine.includes('</pre>');

    if (startsPre) {
      if (paragraph.length > 0) {
        result.push(`<p>${paragraph.join(' ')}</p>`);
        paragraph = [];
      }
      inPreBlock = !endsPre;
      result.push(line);
      continue;
    }

    if (inPreBlock) {
      if (endsPre) {
        inPreBlock = false;
      }
      result.push(line);
      continue;
    }

    const isBlock = blockElements.some(tag => trimmedLine.startsWith(tag));

    if (isBlock || trimmedLine === '') {
      if (paragraph.length > 0) {
        result.push(`<p>${paragraph.join(' ')}</p>`);
        paragraph = [];
      }
      if (trimmedLine !== '') {
        result.push(line);
      }
    } else {
      paragraph.push(trimmedLine);
    }
  }

  if (paragraph.length > 0) {
    result.push(`<p>${paragraph.join(' ')}</p>`);
  }

  return result.join('\n');
}

/**
 * 处理引用块：将连续的 > 行合并为一个 <blockquote>，每行内容包裹在 <p> 中
 */
function processBlockquotes(html) {
  const lines = html.split('\n');
  const result = [];
  let blockquoteLines = [];

  const flushBlockquote = () => {
    if (blockquoteLines.length > 0) {
      const content = blockquoteLines
        .map(line => {
          const trimmed = line.trim();
          return trimmed ? `<p>${trimmed}</p>` : '';
        })
        .filter(Boolean)
        .join('');
      result.push(`<blockquote>${content}</blockquote>`);
      blockquoteLines = [];
    }
  };

  for (const line of lines) {
    const match = line.match(/^>\s?(.*)/);
    if (match) {
      blockquoteLines.push(match[1]);
    } else {
      flushBlockquote();
      result.push(line);
    }
  }

  flushBlockquote();
  return result.join('\n');
}

module.exports = {
  markdownToHTML
};
