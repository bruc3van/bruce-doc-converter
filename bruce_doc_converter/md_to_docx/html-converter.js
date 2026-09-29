/**
 * HTML 到 DOCX 转换器
 * 将 markdown-it 输出的 HTML 转换为 docx.js 组件结构（Node.js + jsdom）。
 * 布局、表格、脚注与公式处理参考 bruce-md2word 的 src/core/html-to-docx.ts。
 * 所有状态都限定在单次转换内，可安全重入。
 */

const { JSDOM } = require('jsdom');
const fs = require('fs');
const path = require('path');
const {
  Paragraph, TextRun, ImageRun, Table, TableRow, TableCell, WidthType, BorderStyle, AlignmentType,
  VerticalAlign, ExternalHyperlink, InternalHyperlink, Bookmark, FootnoteReferenceRun, SimpleField,
  TableLayoutType, HeadingLevel
} = require('docx');
const { charsToTwips, PAGE_WIDTH, PAGE_HEIGHT, MARGIN, MAX_IMAGE_HEIGHT } = require('./styles');
const { columnWidths, estimatedLines } = require('./layout');
const { Diagnostics } = require('./diagnostics');
const { latexToWordMath } = require('./math');

const MAX_LIST_LEVEL = 4;
// 编号层级的悬挂缩进（twips），与 styles.js 中的 numbering 定义一致
const LIST_HANGING = { ordered: [480, 720, 420, 420, 420], bullet: [360, 360, 360, 360, 360] };
const LIST_INDENT_STEP = 720;
const CONTENT_WIDTH = PAGE_WIDTH - MARGIN * 2;
const CONTENT_HEIGHT = PAGE_HEIGHT - MARGIN * 2;
const BLOCK_TAGS = /^(UL|OL|P|PRE|TABLE|BLOCKQUOTE|HR|H[1-6])$/;
const ROOT_LAYOUT = Object.freeze({ left: 0, right: 0, quote: false });

// Emoji 正则：匹配常见 Emoji 字符（包括组合 Emoji）
const EMOJI_SPLIT = /(\p{Emoji_Presentation}|\p{Extended_Pictographic}(?:\u{FE0F}|\u{200D}\p{Extended_Pictographic})*)/gu;
const EMOJI_TEST = /\p{Emoji_Presentation}|\p{Extended_Pictographic}/u;

const SVG_FALLBACK_PIXEL = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO6N6t0AAAAASUVORK5CYII=',
  'base64'
);
const MIME_TO_IMAGE_TYPE = {
  'image/png': 'png', 'image/jpg': 'jpg', 'image/jpeg': 'jpg', 'image/gif': 'gif', 'image/bmp': 'bmp', 'image/svg+xml': 'svg'
};
const EXT_TO_IMAGE_TYPE = {
  '.png': 'png', '.jpg': 'jpg', '.jpeg': 'jpg', '.gif': 'gif', '.bmp': 'bmp', '.svg': 'svg'
};

/** 将文本拆分为普通文本和 Emoji 片段，对 Emoji 使用 Segoe UI Emoji 字体 */
function textRuns(text, style = {}) {
  return text.split(EMOJI_SPLIT).filter(Boolean).map(piece => new TextRun({
    ...(EMOJI_TEST.test(piece) ? { font: 'Segoe UI Emoji' } : {}), ...style, text: piece
  }));
}

const headingSlug = (text) => text.toLowerCase().trim().replace(/[^\p{L}\p{N}\p{M}_\-\s]/gu, '').replace(/\s/g, '-');
const sourceLine = (el) => Number(el && el.getAttribute && el.getAttribute('data-line')) || undefined;

/**
 * 将 HTML 字符串转换为 docx 组件
 * @param {string} htmlString
 * @param {string} [basePath] Markdown 所在目录，用于解析相对图片
 * @param {{formulas?: object[], diagnostics?: Diagnostics}} [options]
 * @returns {{children: Array, footnotes: object, updateFields: boolean, warnings: string[], diagnostics: Diagnostics}}
 */
function convertHTMLToDocx(htmlString, basePath, options = {}) {
  const diagnostics = options.diagnostics || new Diagnostics();
  const images = createImageLoader(typeof basePath === 'string' && basePath ? basePath : null);
  const formulas = convertFormulas(options.formulas || [], diagnostics);
  const dom = new JSDOM(`<body>${htmlString}</body>`);
  const document = dom.window.document;
  const body = document.body;

  let orderedListInstance = 0;
  let updateFields = false;
  const footnotes = {};
  const referencedNotes = new Set();
  const noteNumbers = new Map();
  const noteElements = new Map();
  for (const el of Array.from(document.querySelectorAll('[data-footnote-id]'))) {
    noteElements.set(el.getAttribute('data-footnote-id'), el);
  }

  // 标题书签：支持 [文字](#标题) 形式的文档内跳转
  const anchors = new Map();
  const bookmarks = new Map();
  for (const heading of Array.from(document.querySelectorAll('h1,h2,h3,h4,h5,h6'))) {
    const base = headingSlug(heading.textContent || '') || 'section';
    let slug = base;
    let suffix = 0;
    while (anchors.has(slug)) slug = `${base}-${++suffix}`;
    const name = `heading_${bookmarks.size + 1}`;
    anchors.set(slug, name);
    bookmarks.set(heading, name);
  }

  const availablePixels = (layout) => Math.max(1, (CONTENT_WIDTH - layout.left - layout.right) / 15);

  function imageRun(el, maxWidth, style) {
    const src = (el.getAttribute('src') || '').trim();
    const alt = el.getAttribute('alt') || '';
    const run = images.createRun(el, { maxWidth, maxHeight: Math.min(MAX_IMAGE_HEIGHT, CONTENT_HEIGHT / 15 - 80) });
    if (run) return [run];
    diagnostics.add('IMAGE_UNAVAILABLE', imageFailureMessage(src, alt), 'warning', sourceLine(el));
    const label = `[图片: ${alt || '图片'}]${src && !src.startsWith('data:') ? ` (${src})` : ''}`;
    return textRuns(label, { ...style, italics: true, color: '6B7280' });
  }

  // Inline code is 11 pt in body text; codeSize 0 inherits the paragraph size (headings).
  function inline(nodes, style = {}, maxWidth = CONTENT_WIDTH / 15, codeSize = 22) {
    const runs = [];
    for (const node of nodes) {
      if (node.nodeType === 3) {
        // markdown-it adds a formatting newline after <br>; the run already holds the break.
        const raw = node.previousSibling && node.previousSibling.nodeName === 'BR'
          ? (node.textContent || '').replace(/^\r?\n/, '') : node.textContent || '';
        const text = raw.replace(/[ \t\r\n\f]+/g, ' ');
        if (text) runs.push(...textRuns(text, style));
        continue;
      }
      if (node.nodeType !== 1) continue;
      const el = node;
      const tag = el.tagName;
      if (el.hasAttribute('data-footnote-ref')) {
        const id = el.getAttribute('data-footnote-ref');
        if (referencedNotes.has(id)) {
          // Repeated references point at the first one's number and stay in sync after edits.
          const field = new SimpleField(`NOTEREF note_${id} \\h \\f`);
          field.addChildElement(new TextRun({ text: String(noteNumbers.get(id)), style: 'FootnoteReference', superScript: true }));
          runs.push(field);
          updateFields = true;
        } else {
          referencedNotes.add(id);
          noteNumbers.set(id, referencedNotes.size);
          runs.push(new Bookmark({ id: `note_${id}`, children: [new FootnoteReferenceRun(Number(id))] }));
        }
        continue;
      }
      if (el.hasAttribute('data-math')) {
        const formula = formulas.get(el.getAttribute('data-math'));
        if (formula) runs.push(...formula);
        continue;
      }
      if (tag === 'IMG') runs.push(...imageRun(el, maxWidth, style));
      else if (tag === 'BR') runs.push(new TextRun({ text: '', break: 1 }));
      else if (tag === 'CODE') {
        const size = style.size || codeSize;
        runs.push(...textRuns(el.textContent || '', { ...style, font: 'Consolas', color: 'DC2626', ...(size ? { size } : {}) }));
      } else if (tag === 'A') {
        const href = (el.getAttribute('href') || '').trim();
        const children = inline(el.childNodes, { ...style, color: '2563EB', underline: {} }, maxWidth, codeSize);
        if (href.startsWith('#')) {
          let target;
          try { target = anchors.get(decodeURIComponent(href.slice(1))); } catch (error) { /* invalid fragment */ }
          if (target) runs.push(new InternalHyperlink({ anchor: target, children }));
          else {
            runs.push(...children);
            diagnostics.add('LINK_UNAVAILABLE', `文档内链接 ${safeDecodeURIComponent(href)} 没有对应标题，已保留链接文字。`, 'warning', sourceLine(el));
          }
        } else if (href && children.length) runs.push(new ExternalHyperlink({ link: href, children }));
        else runs.push(...children);
      } else {
        runs.push(...inline(el.childNodes, {
          ...style,
          ...(['STRONG', 'B'].includes(tag) ? { bold: true } : {}),
          ...(['EM', 'I'].includes(tag) ? { italics: true } : {}),
          ...(['DEL', 'S'].includes(tag) ? { strike: true } : {})
        }, maxWidth, codeSize));
      }
    }
    return runs;
  }

  /**
   * 列表：独立的有序列表（顶层或挂在无序列表下的新 OL）分配新 instance，Word 才会从 1 重新编号；
   * 嵌套 OL 继承父有序列表的 instance。缩进按所在容器（引用、父列表）叠加。
   */
  function list(el, level, layout, parentInstance = null) {
    if (level > MAX_LIST_LEVEL) diagnostics.add('LIST_DEPTH_REDUCED', '超过五级的列表嵌套已按第五级排版。', 'info');
    const safeLevel = Math.min(level, MAX_LIST_LEVEL);
    const ordered = el.tagName === 'OL';
    const reference = ordered ? 'numbered-list' : 'bullet-list';
    const instance = ordered ? (parentInstance == null ? orderedListInstance++ : parentInstance) : null;
    const hanging = LIST_HANGING[ordered ? 'ordered' : 'bullet'][safeLevel];
    const textLeft = Math.min(layout.left + (level > MAX_LIST_LEVEL ? 0 : LIST_INDENT_STEP), CONTENT_WIDTH - layout.right - LIST_INDENT_STEP);
    const itemLayout = { ...layout, left: textLeft };
    const result = [];
    for (const li of Array.from(el.children).filter(child => child.tagName === 'LI')) {
      let numbered = false;
      let pending = [];
      const flush = (force = false) => {
        if (!pending.length && !force) return;
        const runs = inline(pending, {}, availablePixels(itemLayout));
        result.push(new Paragraph({
          children: runs.length ? runs : [new TextRun('')],
          ...(layout.quote ? { style: 'Quote' } : {}),
          ...(!numbered
            ? { numbering: { reference, level: safeLevel, ...(ordered ? { instance } : {}) }, indent: { left: textLeft, hanging, right: layout.right } }
            // 列表项续行：与列表文本对齐，不带编号
            : { indent: { left: textLeft, firstLine: 0, hanging: 0, right: layout.right } })
        }));
        numbered = true;
        pending = [];
      };
      for (const child of li.childNodes) {
        // Keep soft breaks between inline siblings, but discard list HTML layout whitespace.
        if (child.nodeType === 3 && !(child.textContent || '').trim()
          && (!pending.length || !child.nextSibling || BLOCK_TAGS.test(child.nextSibling.nodeName))) continue;
        const tag = child.nodeName;
        if (tag === 'UL' || tag === 'OL') {
          flush(!numbered);
          result.push(...list(child, level + 1, itemLayout, tag === 'OL' && ordered ? instance : null));
        } else if (tag === 'P' && child.hasAttribute('data-math-block')) {
          flush(!numbered);
          result.push(...block(child, itemLayout));
        } else if (tag === 'P') {
          flush();
          pending.push(...child.childNodes);
          flush(!numbered);
        } else if (BLOCK_TAGS.test(tag)) {
          flush(!numbered);
          result.push(...block(child, itemLayout));
        } else pending.push(child);
      }
      flush(!numbered);
    }
    return result;
  }

  function table(el, layout) {
    const trs = Array.from(el.querySelectorAll('tr'));
    const count = trs.reduce((max, tr) => Math.max(max, tr.children.length), 1);
    const tableWidth = Math.max(1, CONTENT_WIDTH - layout.left - layout.right);
    const widths = columnWidths(trs, count, tableWidth);
    const outer = { style: BorderStyle.SINGLE, size: 6, color: '9CA3AF' };
    const inner = { style: BorderStyle.SINGLE, size: 4, color: 'D1D5DB' };
    const alignments = { left: AlignmentType.LEFT, center: AlignmentType.CENTER, right: AlignmentType.RIGHT };
    return new Table({
      layout: TableLayoutType.FIXED,
      indent: { size: layout.left, type: WidthType.DXA },
      width: { size: tableWidth, type: WidthType.DXA },
      columnWidths: widths,
      borders: { top: outer, bottom: outer, left: outer, right: outer, insideHorizontal: inner, insideVertical: inner },
      rows: trs.map(tr => {
        const header = tr.parentElement && tr.parentElement.tagName === 'THEAD';
        const cells = Array.from(tr.children);
        // Short rows stay on one page; tall rows (images, formulas, long text) may split.
        const cantSplit = cells.every((cell, i) => !cell.querySelector('img,[data-math]')
          && estimatedLines(cell.textContent || '', widths[i] - 300) <= 8);
        return new TableRow({
          cantSplit,
          tableHeader: header || undefined,
          children: cells.map((cell, i) => {
            const runs = inline(cell.childNodes, header ? { bold: true, size: 24 } : {}, Math.max(1, (widths[i] - 300) / 15));
            const align = alignments[cell.style.textAlign]
              || (header || cell.tagName === 'TH' ? AlignmentType.CENTER : AlignmentType.LEFT);
            return new TableCell({
              width: { size: widths[i], type: WidthType.DXA },
              ...(header ? { shading: { fill: 'E5E7EB' } } : {}),
              verticalAlign: VerticalAlign.CENTER,
              margins: { top: header ? 120 : 100, bottom: header ? 120 : 100, left: 150, right: 150 },
              children: [new Paragraph({ children: runs.length ? runs : [new TextRun('')], alignment: align, indent: { firstLine: 0 } })]
            });
          })
        });
      })
    });
  }

  function block(node, layout = ROOT_LAYOUT, quoteAfter) {
    if (node.nodeType === 3) {
      return (node.textContent || '').trim()
        ? [new Paragraph({ children: inline([node], {}, availablePixels(layout)), style: 'BodyText', indent: { left: layout.left, right: layout.right } })]
        : [];
    }
    if (node.nodeType !== 1) return [];
    const el = node;
    const tag = el.tagName;
    if (el.hasAttribute('data-footnotes')) return [];
    const width = availablePixels(layout);
    const indent = { left: layout.left, right: layout.right, firstLine: 0 };

    if (/^H[1-6]$/.test(tag)) {
      const runs = inline(el.childNodes, {}, width, 0);
      return [new Paragraph({
        children: [new Bookmark({ id: bookmarks.get(el), children: runs.length ? runs : [new TextRun('')] })],
        heading: HeadingLevel[`HEADING_${tag[1]}`],
        ...(layout.left || layout.right ? { indent } : {}),
        keepNext: true,
        keepLines: true
      })];
    }
    if (tag === 'P' && el.hasAttribute('data-math-block')) {
      return [new Paragraph({ children: inline(el.childNodes, {}, width), alignment: AlignmentType.CENTER, indent, spacing: { before: 160, after: 160 } })];
    }
    if (tag === 'P' && el.hasAttribute('data-mermaid-notice')) {
      return [new Paragraph({ children: [new TextRun({ text: el.textContent || '', color: '92400E', size: 20 })], indent, spacing: { before: 160, after: 80 }, keepNext: true })];
    }
    if (tag === 'P') {
      const meaningful = Array.from(el.childNodes).filter(child => child.nodeType !== 3 || (child.textContent || '').trim());
      if (meaningful.length === 1 && meaningful[0].nodeName === 'IMG') return block(meaningful[0], layout);
      const runs = inline(el.childNodes, {}, width);
      return [new Paragraph({
        children: runs.length ? runs : [new TextRun('')],
        style: layout.quote ? 'Quote' : 'BodyText',
        widowControl: true,
        ...(layout.quote || layout.left || layout.right ? { indent } : {}),
        ...(layout.quote && quoteAfter !== undefined ? { spacing: { after: quoteAfter } } : {})
      })];
    }
    if (tag === 'PRE') {
      const text = el.textContent || '';
      const lines = text.split('\n');
      return [new Paragraph({
        style: 'CodeBlock',
        // Short code blocks stay on one page; long ones may split to avoid large gaps.
        keepLines: estimatedLines(text, width * 15 - 480) <= 12,
        widowControl: true,
        indent: { ...indent, left: layout.left + 240, right: layout.right + 240 },
        children: lines.flatMap((line, index) => [
          ...(index ? [new TextRun({ text: '', break: 1 })] : []),
          new TextRun({ text: line || ' ', font: 'Consolas', size: 22, color: '1F2937' })
        ])
      })];
    }
    if (tag === 'HR') {
      return [new Paragraph({ indent, spacing: { before: 200, after: 200 }, border: { bottom: { style: BorderStyle.SINGLE, color: '9CA3AF', size: 12, space: 1 } }, children: [] })];
    }
    if (tag === 'UL' || tag === 'OL') return list(el, 0, layout);
    if (tag === 'TABLE') return [table(el, layout)];
    if (tag === 'IMG') {
      const runs = inline([el], {}, width);
      const embedded = runs.some(run => run instanceof ImageRun);
      return [new Paragraph({ children: runs, alignment: embedded ? AlignmentType.CENTER : undefined, indent, spacing: { before: 200, after: 200 } })];
    }
    if (tag === 'BLOCKQUOTE') {
      const children = Array.from(el.childNodes).filter(child => child.nodeType !== 3 || (child.textContent || '').trim());
      const quoteLayout = { ...layout, left: layout.left + charsToTwips(2), quote: true };
      // 多段落引用块：只有最后一段保留 after 间距，避免段间距叠加
      return children.flatMap((child, index) => block(child, quoteLayout, index === children.length - 1 ? undefined : 0));
    }
    return Array.from(el.childNodes).flatMap(child => block(child, layout));
  }

  function footnoteParagraphs(note) {
    const paragraphs = [];
    const visit = (el) => {
      if (el.tagName === 'TABLE') {
        diagnostics.add('FOOTNOTE_TABLE_FLATTENED', '脚注中的表格已按单元格顺序拆为段落。', 'warning');
        for (const cell of Array.from(el.querySelectorAll('th,td'))) {
          paragraphs.push(new Paragraph({ style: 'FootnoteText', children: inline(cell.childNodes) }));
        }
      } else if (el.tagName === 'P' && !el.hasAttribute('data-math-block')) {
        paragraphs.push(new Paragraph({ style: 'FootnoteText', children: inline(el.childNodes), widowControl: true }));
      } else if (el.querySelector('table')) {
        for (const child of Array.from(el.children)) visit(child);
      } else {
        for (const converted of block(el)) if (converted instanceof Paragraph) paragraphs.push(converted);
      }
    };
    for (const child of Array.from(note.children)) visit(child);
    return paragraphs;
  }

  try {
    const children = [];
    // Walk siblings rather than body.childNodes: a live body-level NodeList makes
    // jsdom's window.close() teardown quadratic in the number of top-level blocks.
    for (let node = body.firstChild; node; node = node.nextSibling) {
      for (const item of block(node)) {
        // Word merges adjacent tables even when separate table XML is emitted.
        if (item instanceof Table && children[children.length - 1] instanceof Table) {
          children.push(new Paragraph({ spacing: { before: 0, after: 80, line: 20 }, children: [new TextRun({ text: '', size: 2 })] }));
        }
        children.push(item);
      }
    }
    for (const [id, note] of noteElements) {
      const paragraphs = footnoteParagraphs(note);
      if (paragraphs.length) paragraphs[0].addRunToFront(new TextRun(' '));
      if (referencedNotes.has(id)) footnotes[id] = { children: paragraphs.length ? paragraphs : [new Paragraph('')] };
      else children.push(...paragraphs);
    }
    return {
      children: children.length ? children : [new Paragraph({ text: '' })],
      footnotes,
      updateFields,
      warnings: diagnostics.warnings(),
      diagnostics
    };
  } finally {
    dom.window.close();
  }
}

/** LaTeX -> Word 公式；失败时保留完整源码并报告诊断 */
function convertFormulas(formulas, diagnostics) {
  const converted = new Map();
  for (const formula of formulas) {
    try {
      if (formula.unclosed) throw new Error('Unclosed formula delimiter');
      converted.set(formula.id, [latexToWordMath(formula.source, formula.display)]);
    } catch (error) {
      diagnostics.add('MATH_NOT_CONVERTED', '公式未转换：语法无效、超出处理范围或不支持该结构，已保留完整源码。', 'warning', formula.line);
      converted.set(formula.id, [
        new TextRun({ text: '[公式未转换] ', color: '92400E' }),
        ...formula.raw.split('\n').map((text, i) => new TextRun({ text, ...(i ? { break: 1 } : {}), font: 'Consolas' }))
      ]);
    }
  }
  return converted;
}

function imageFailureMessage(src, alt) {
  const label = src || alt || '(空 src)';
  if (/^https?:\/\//i.test(src)) return `图片无法嵌入（不支持远程 URL）: ${label}`;
  if (src.startsWith('file://')) return `图片无法嵌入（不支持 file://）: ${label}`;
  if (src && path.isAbsolute(src)) return `图片无法嵌入（禁止绝对路径）: ${label}`;
  return `图片无法嵌入: ${label}`;
}

/** 只读取 data URI 或 Markdown 目录内的相对路径图片 */
function createImageLoader(basePath) {
  function parseDataImage(src) {
    const match = src.match(/^data:(image\/[a-zA-Z0-9.+-]+)(;[^,]*)?,([\s\S]+)$/i);
    if (!match) return null;
    const imageType = MIME_TO_IMAGE_TYPE[match[1].toLowerCase()];
    if (!imageType) return null;
    let buffer;
    if ((match[2] || '').toLowerCase().includes(';base64')) buffer = Buffer.from(match[3], 'base64');
    else {
      try { buffer = Buffer.from(decodeURIComponent(match[3]), 'utf-8'); } catch (error) { return null; }
    }
    return buffer.length ? describe(imageType, buffer) : null;
  }

  function parseLocalImage(src) {
    if (/^https?:\/\//i.test(src) || src.startsWith('data:') || src.startsWith('file://') || !basePath) return null;
    let decodedPath = safeDecodeURIComponent(src);
    // 处理 /C:/path 格式（前导斜杠 + Windows 盘符）
    if (process.platform === 'win32' && /^\/[a-zA-Z]:[\\/]/.test(decodedPath)) decodedPath = decodedPath.slice(1);
    if (path.isAbsolute(decodedPath)) return null;
    const resolvedPath = path.resolve(basePath, decodedPath);
    if (!fs.existsSync(resolvedPath) || !fs.statSync(resolvedPath).isFile()) return null;
    let baseRealPath;
    let imageRealPath;
    try {
      baseRealPath = fs.realpathSync(basePath);
      imageRealPath = fs.realpathSync(resolvedPath);
    } catch (error) {
      return null;
    }
    if (!isPathInside(imageRealPath, baseRealPath)) return null;
    const imageType = EXT_TO_IMAGE_TYPE[path.extname(resolvedPath).toLowerCase()];
    return imageType ? describe(imageType, fs.readFileSync(resolvedPath)) : null;
  }

  function describe(type, data) {
    const size = type === 'svg' ? parseSvgMeta(data.toString('utf-8')) : parseBitmapSize(data, type);
    return { type, data, size };
  }

  /** Natural size (or explicit width/height attributes), capped by the placement bounds. */
  function createRun(imgElement, bounds) {
    const src = (imgElement.getAttribute('src') || '').trim();
    if (!src) return null;
    const image = parseDataImage(src) || parseLocalImage(src);
    if (!image) return null;
    const widthAttr = Number.parseInt(imgElement.getAttribute('width') || '', 10);
    const heightAttr = Number.parseInt(imgElement.getAttribute('height') || '', 10);
    const natural = image.size && image.size.width > 0 && image.size.height > 0 ? image.size : null;
    const ratio = widthAttr > 0 && heightAttr > 0 ? widthAttr / heightAttr : natural ? natural.width / natural.height : 16 / 9;
    const preferred = widthAttr > 0 ? widthAttr : heightAttr > 0 ? heightAttr * ratio : natural ? natural.width : bounds.maxWidth;
    const width = Math.max(1, Math.round(Math.min(preferred, bounds.maxWidth, bounds.maxHeight * ratio)));
    const transformation = { width, height: Math.max(1, Math.round(width / ratio)) };
    const alt = imgElement.getAttribute('alt') || '';
    const altText = { title: alt, description: alt, name: 'Image' };
    if (image.type === 'svg') {
      return new ImageRun({ type: 'svg', data: image.data, fallback: { type: 'png', data: SVG_FALLBACK_PIXEL }, transformation, altText });
    }
    return new ImageRun({ type: image.type, data: image.data, transformation, altText });
  }

  return { createRun };
}

/** 从位图 Buffer 中解析实际像素尺寸（PNG、JPEG、GIF、BMP） */
function parseBitmapSize(buffer, type) {
  if (!Buffer.isBuffer(buffer) || buffer.length < 10) return null;
  if (type === 'png') {
    return buffer.length >= 24 && buffer.toString('hex', 0, 8) === '89504e470d0a1a0a'
      ? { width: buffer.readUInt32BE(16), height: buffer.readUInt32BE(20) } : null;
  }
  if (type === 'jpg') {
    // 查找 SOF0/SOF2 marker
    let offset = 2;
    while (offset + 9 < buffer.length) {
      if (buffer[offset] !== 0xFF) break;
      const marker = buffer[offset + 1];
      if (marker === 0xC0 || marker === 0xC2) return { width: buffer.readUInt16BE(offset + 7), height: buffer.readUInt16BE(offset + 5) };
      offset += 2 + buffer.readUInt16BE(offset + 2);
    }
    return null;
  }
  if (type === 'gif') {
    return buffer.toString('ascii', 0, 3) === 'GIF' ? { width: buffer.readUInt16LE(6), height: buffer.readUInt16LE(8) } : null;
  }
  if (type === 'bmp') {
    return buffer.length >= 26 && buffer.readUInt16LE(0) === 0x4D42
      ? { width: Math.abs(buffer.readInt32LE(18)), height: Math.abs(buffer.readInt32LE(22)) } : null;
  }
  return null;
}

function parseSvgMeta(svgText) {
  const viewBoxMatch = svgText.match(/viewBox="([^"]+)"/i);
  if (viewBoxMatch) {
    const values = viewBoxMatch[1].trim().split(/[\s,]+/).map(value => Number.parseFloat(value));
    if (values.length === 4 && values[2] > 0 && values[3] > 0) return { width: values[2], height: values[3] };
  }
  const widthMatch = svgText.match(/\bwidth="([\d.]+)(px)?"/i);
  const heightMatch = svgText.match(/\bheight="([\d.]+)(px)?"/i);
  const width = widthMatch ? Number.parseFloat(widthMatch[1]) : NaN;
  const height = heightMatch ? Number.parseFloat(heightMatch[1]) : NaN;
  return width > 0 && height > 0 ? { width, height } : null;
}

function safeDecodeURIComponent(value) {
  try {
    return decodeURIComponent(value);
  } catch (error) {
    return value;
  }
}

function isPathInside(childPath, parentPath) {
  const relative = path.relative(parentPath, childPath);
  return relative === '' || (!relative.startsWith('..') && !path.isAbsolute(relative));
}

module.exports = {
  convertHTMLToDocx
};
