/**
 * LaTeX -> Temml MathML -> Word 原生公式（OMML）。
 * 移植自 bruce-md2word 的 src/core/math.ts。只输出能理解的结构，未知 MathML 直接抛错，
 * 由调用方保留源码并报告诊断，绝不静默降级为纯文本。
 */
const temml = require('temml');
const { JSDOM } = require('jsdom');
const { Math: WordMath, XmlComponent, XmlAttributeComponent } = require('docx');

class Attributes extends XmlAttributeComponent {}
class Element extends XmlComponent {
  constructor(tag, children = [], attrs) {
    super(tag);
    if (attrs) this.root.push(new Attributes(attrs));
    this.root.push(...children);
  }
}

const m = (tag, children = []) => new Element(`m:${tag}`, children);
const val = (tag, value) => new Element(`m:${tag}`, [], { 'm:val': value });
const run = (text, plain = false, bold = false, italic = false) => m('r', [
  ...(plain || bold || italic ? [m('rPr', [val('sty', bold ? (italic ? 'bi' : 'b') : italic ? 'i' : 'p')])] : []),
  new Element('w:rPr', [new Element('w:rFonts', [], { 'w:ascii': 'Cambria Math', 'w:hAnsi': 'Cambria Math' })]),
  new Element('m:t', [text], { 'xml:space': 'preserve' })
]);
const NARY_CHARS = /^[∑∏∐⋂⋃⋀⋁∫∬∭∮∯∰]$/u;
// These constructs need document-wide state or layout the mapper does not represent;
// Temml would otherwise drop labels, tags or styling and report a complete conversion.
const UNSUPPORTED_COMMANDS = /\\(?:textbf|textit|textsf|texttt|textrm|textnormal|label|ref|eqref|tag|notag|nonumber|def|gdef|edef|xdef|newcommand|renewcommand|providecommand|let|global|includegraphics|href|url|html\w*|class|style|color|textcolor|definecolor|phantom|hphantom|vphantom|smash|raisebox|rule|kern|mkern|hspace|vspace|fontsize|tiny|small|large|Large|huge|Huge)\b/;

function latexToWordMath(source, display) {
  if (!source.trim() || Buffer.byteLength(source) > 50000) throw new Error('Formula is empty or too large');
  if (UNSUPPORTED_COMMANDS.test(source)) throw new Error('Unsupported formula command');
  const mathml = temml.renderToString(source, {
    displayMode: display, throwOnError: true, xml: true, trust: false, maxExpand: 1000, maxSize: [20, 200], macros: {}
  });
  if (Buffer.byteLength(mathml) > 1000000) throw new Error('Expanded formula is too large');
  const dom = new JSDOM(mathml, { contentType: 'text/xml' });
  let count = 0;

  const unwrap = (node) => {
    while (node.localName === 'mrow' && node.children.length === 1) node = node.children[0];
    return node;
  };
  const children = (node, depth) => sequence(Array.from(node.children), depth);

  function sequence(nodes, depth) {
    const result = [];
    for (let i = 0; i < nodes.length; i++) {
      const current = unwrap(nodes[i]);
      const base = current.children[0] && unwrap(current.children[0]);
      const next = nodes[i + 1];
      // An n-ary operator (sum, integral) owns the following operand.
      if (/^(msub|msup|msubsup|munder|mover|munderover)$/.test(current.localName)
        && base && base.localName === 'mo' && NARY_CHARS.test(base.textContent || '')
        && next && unwrap(next).localName !== 'mo' && unwrap(next).localName !== 'mspace') {
        result.push(...convert(current, depth + 1, next));
        i++;
      } else {
        result.push(...convert(nodes[i], depth + 1));
      }
    }
    return result;
  }

  function convert(node, depth, operand) {
    if (++count > 10000 || depth > 80) throw new Error('Formula complexity exceeded');
    const tag = node.localName;
    const parts = Array.from(node.children);
    const at = (i) => {
      if (!parts[i]) throw new Error('Missing MathML argument');
      return convert(parts[i], depth + 1);
    };
    // Temml layout classes/padding are not copied; semantic styles must never silently vanish.
    if (node.hasAttribute('href') || node.hasAttribute('mathcolor') || node.hasAttribute('mathbackground')
      || /(?:color|background|visibility|transform):/.test(node.getAttribute('style') || '')) {
      throw new Error('Unsupported MathML style');
    }
    switch (tag) {
      case 'mstyle': {
        if (Array.from(node.attributes).some(attr => !['displaystyle', 'scriptlevel'].includes(attr.name))
          || !['', '0'].includes(node.getAttribute('scriptlevel') || '')) throw new Error('Unsupported math style');
        return children(node, depth);
      }
      case 'math':
      case 'mrow': {
        const last = parts[parts.length - 1];
        if (parts.length >= 2 && parts[0].localName === 'mo' && parts[0].getAttribute('fence') === 'true'
          && last.localName === 'mo' && last.getAttribute('fence') === 'true') {
          return [m('d', [
            m('dPr', [val('begChr', parts[0].textContent || ''), val('endChr', last.textContent || ''), val('grow', '1')]),
            m('e', sequence(parts.slice(1, -1), depth))
          ])];
        }
        return children(node, depth);
      }
      case 'mi': case 'mn': case 'mo': case 'mtext': {
        if (parts.length) throw new Error('Unexpected token children');
        const variant = node.getAttribute('mathvariant');
        if (variant && variant !== 'normal') throw new Error('Unsupported math variant');
        const style = node.getAttribute('style') || '';
        if (style.includes('font-family:')) throw new Error('Unsupported math font');
        const text = node.textContent || '';
        return [run(text, tag === 'mtext' || tag === 'mo' || tag === 'mn' || variant === 'normal' || text.length > 1,
          /font-weight:\s*bold/.test(style), /font-style:\s*italic/.test(style))];
      }
      case 'mspace': {
        const width = node.getAttribute('width') || '0em';
        if (/^0(?:em|pt|px)?$/.test(width)) return [];
        const match = /^(-?[\d.]+)em$/.exec(width);
        if (!match || !Number.isFinite(Number(match[1])) || Math.abs(Number(match[1])) > 2) throw new Error('Unsupported math spacing');
        return Number(match[1]) > 0 ? [run(Number(match[1]) >= 1 ? ' ' : ' ', true)] : [];
      }
      case 'mfrac': {
        const thickness = node.getAttribute('linethickness');
        if (thickness && !/^0(?:px|pt|em)?$/.test(thickness)) throw new Error('Unsupported fraction thickness');
        return [m('f', [...(thickness ? [m('fPr', [val('type', 'noBar')])] : []), m('num', at(0)), m('den', at(1))])];
      }
      case 'msqrt': return [m('rad', [m('radPr', [val('degHide', '1')]), m('deg'), m('e', children(node, depth))])];
      case 'mroot': return [m('rad', [m('radPr', [val('degHide', '0')]), m('deg', at(1)), m('e', at(0))])];
      case 'msub': case 'msup': case 'msubsup':
      case 'munder': case 'mover': case 'munderover': {
        const base = unwrap(parts[0]);
        const glyph = base.textContent || '';
        const lower = ['msub', 'msubsup', 'munder', 'munderover'].includes(tag);
        const upper = ['msup', 'msubsup', 'mover', 'munderover'].includes(tag);
        if (base.localName === 'mo' && NARY_CHARS.test(glyph)) {
          return [m('nary', [
            m('naryPr', [val('chr', glyph), val('limLoc', tag.includes('under') || tag.includes('over') ? 'undOvr' : 'subSup'),
              val('subHide', lower ? '0' : '1'), val('supHide', upper ? '0' : '1')]),
            m('sub', lower ? at(1) : []), m('sup', upper ? at(lower ? 2 : 1) : []),
            m('e', operand ? convert(operand, depth + 1) : [])
          ])];
        }
        if (tag === 'mover' && parts[1].localName === 'mo' && Array.from(parts[1].textContent || '').length === 1) {
          const accent = parts[1].textContent;
          if ('→←↔ˆ^˜~˙¨‾¯˘ˇ´`'.includes(accent)) return [m('acc', [m('accPr', [val('chr', accent)]), m('e', at(0))])];
          if (accent === '⏞') return [m('groupChr', [m('groupChrPr', [val('chr', accent), val('pos', 'top'), val('vertJc', 'bot')]), m('e', at(0))])];
        }
        if (tag === 'munder' && parts[1].textContent === '⏟') {
          return [m('groupChr', [m('groupChrPr', [val('chr', '⏟'), val('pos', 'bot'), val('vertJc', 'top')]), m('e', at(0))])];
        }
        if (tag === 'msubsup') return [m('sSubSup', [m('e', at(0)), m('sub', at(1)), m('sup', at(2))])];
        if (tag === 'msub' || tag === 'msup') return [m(tag === 'msub' ? 'sSub' : 'sSup', [m('e', at(0)), m(tag === 'msub' ? 'sub' : 'sup', at(1))])];
        if (tag === 'munderover') return [m('limUpp', [m('e', [m('limLow', [m('e', at(0)), m('lim', at(1))])]), m('lim', at(2))])];
        return [m(tag === 'munder' ? 'limLow' : 'limUpp', [m('e', at(0)), m('lim', at(1))])];
      }
      case 'menclose': {
        const notation = node.getAttribute('notation');
        if (notation === 'top' || notation === 'bottom') {
          return [m('bar', [m('barPr', [val('pos', notation === 'top' ? 'top' : 'bot')]), m('e', children(node, depth))])];
        }
        throw new Error('Unsupported enclosure');
      }
      case 'mtable': {
        if (node.hasAttribute('columnlines') || node.hasAttribute('rowlines') || node.hasAttribute('frame')
          || parts.some(row => row.localName !== 'mtr' || Array.from(row.children).some(cell => cell.localName !== 'mtd'))) {
          throw new Error('Unsupported equation table');
        }
        const width = Math.max(0, ...parts.map(row => row.children.length));
        if (!width || width > 32 || parts.length > 100) throw new Error('Matrix dimensions exceeded');
        const columns = Array.from({ length: width }, (_, i) => {
          const cell = parts[0].children[i];
          const align = cell && cell.classList.contains('tml-left') ? 'left' : cell && cell.classList.contains('tml-right') ? 'right' : 'center';
          return m('mc', [m('mcPr', [val('count', '1'), val('mcJc', align)])]);
        });
        return [m('m', [
          m('mPr', [m('mcs', columns)]),
          ...parts.map(row => m('mr', Array.from({ length: width }, (_, i) => m('e', row.children[i] ? children(row.children[i], depth + 1) : []))))
        ])];
      }
      default:
        throw new Error(`Unsupported MathML element: ${tag}`);
    }
  }

  try {
    // Validate every node, including rows/cells and fences handled by their parent.
    const allNodes = dom.window.document.querySelectorAll('*');
    if (allNodes.length > 10000) throw new Error('Formula node budget exceeded');
    for (const node of allNodes) {
      const style = node.getAttribute('style') || '';
      if (/(?:border|color|background|visibility|transform):/.test(style) || /border-/.test(style)
        || ['rowspan', 'columnspan', 'rowlines', 'columnlines', 'frame', 'href', 'mathcolor', 'mathbackground'].some(attr => node.hasAttribute(attr))) {
        throw new Error('Unsupported MathML layout');
      }
    }
    const math = new WordMath({ children: [] });
    for (const child of convert(dom.window.document.documentElement, 0)) math.addChildElement(child);
    return display ? m('oMathPara', [m('oMathParaPr', [val('jc', 'center')]), math]) : math;
  } finally {
    dom.window.close();
  }
}

module.exports = { latexToWordMath };
