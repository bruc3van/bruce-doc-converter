const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

const { markdownToHTML } = require('../bruce_doc_converter/md_to_docx/markdown-converter');
const { convertHTMLToDocx } = require('../bruce_doc_converter/md_to_docx/html-converter');
const { buildPuppeteerConfig } = require('../bruce_doc_converter/md_to_docx/mermaid-renderer');
const { resolveUniqueDocxPath } = require('../bruce_doc_converter/md_to_docx/index');

test('CommonMark keeps underscores, escapes and tilde code fences', async () => {
  const { html, warnings } = await markdownToHTML('foo_bar_baz \\*literal\\*\n\n~~~js\nconst x = "<tag>";\n~~~');
  assert.match(html, /foo_bar_baz \*literal\*/);
  assert.match(html, /<pre><code class="language-js">const x = &quot;&lt;tag&gt;&quot;;<\/code>/);
  assert.deepEqual(warnings, []);
});

test('CommonMark preserves inline formatting in links and rejects unsafe links', async () => {
  const { html } = await markdownToHTML('[**bold**](https://example.com) [bad](javascript:alert%281%29)');
  assert.match(html, /<a href="https:\/\/example.com"><strong>bold<\/strong><\/a>/);
  assert.doesNotMatch(html, /href="javascript:/);
});

test('Mermaid tokens support tilde fences and preserve source on renderer failure', async () => {
  const renderer = require('../bruce_doc_converter/md_to_docx/mermaid-renderer');
  const { mock } = require('node:test');
  try {
    mock.method(renderer, 'renderMermaidToDataUrl', async () => ({ success: false, error: 'fixture failure' }));
    const result = await markdownToHTML('~~~mermaid\ngraph TD; A-->B;\n~~~');
    assert.match(result.html, /language-mermaid/);
    assert.match(result.html, /A--&gt;B/);
    assert.match(result.warnings[0], /fixture failure/);
  } finally {
    mock.restoreAll();
  }
});

test('Node runtime gate rejects unsupported versions', () => {
  const { supportedNodeVersion } = require('../bruce_doc_converter/md_to_docx/index');
  assert.equal(supportedNodeVersion('21.7.3'), false);
  assert.equal(supportedNodeVersion('20.20.0'), false);
  assert.equal(supportedNodeVersion('22.0.0'), true);
  assert.equal(supportedNodeVersion('22.18.0'), true);
  assert.equal(supportedNodeVersion('24.0.0'), true);
});

function collectDocxText(value) {
  if (typeof value === 'string') return value;
  if (!value || typeof value !== 'object') return '';
  if (Array.isArray(value)) return value.map(collectDocxText).join('');
  return collectDocxText(value.root);
}

function getNumberingRefs(paragraph) {
  return (paragraph.properties && paragraph.properties.numberingReferences) || [];
}

async function mdToChildren(markdown) {
  const { html, warnings } = await markdownToHTML(markdown);
  const { children, warnings: htmlWarnings } = convertHTMLToDocx(html, process.cwd());
  return { children, warnings: [...warnings, ...htmlWarnings], html };
}

test('markdown 表格支持转义管道符', async () => {
  const markdown = [
    '| col1 | col2 |',
    '| --- | --- |',
    '| a\\|b | c |'
  ].join('\n');

  const { html } = await markdownToHTML(markdown);

  assert.match(html, /<th>col1<\/th>/);
  assert.match(html, /<th>col2<\/th>/);
  assert.match(html, /<td>a\|b<\/td>/);
  assert.equal((html.match(/<th>/g) || []).length, 2);
  assert.equal((html.match(/<td>/g) || []).length, 2);
});

test('markdown 链接和图片支持带圆括号的 URL', async () => {
  const markdown = [
    '[link](https://example.com/a_(b).png)',
    '',
    '![img](https://example.com/a_(b).png)'
  ].join('\n');

  const { html } = await markdownToHTML(markdown);

  assert.match(html, /href="https:\/\/example\.com\/a_\(b\)\.png"/);
  assert.match(html, /src="https:\/\/example\.com\/a_\(b\)\.png"/);
});

test('HTML 转 DOCX 保留混合内联样式和超链接目标', () => {
  const { children } = convertHTMLToDocx(
    '<p><strong><em>混合格式</em></strong> <a href="https://example.com">链接</a></p>',
    process.cwd()
  );

  assert.equal(children.length, 1);

  const paragraph = children[0];
  const firstRun = paragraph.root.find(child => child && child.rootKey === 'w:r');
  const hyperlink = paragraph.root.find(child => child && child.rootKey === 'w:externalHyperlink');

  assert.ok(firstRun, '应生成首个文本 run');
  assert.ok(hyperlink, '应保留超链接节点');
  assert.equal(hyperlink.options.link, 'https://example.com');

  const styleKeys = firstRun.properties.root.map(item => item.rootKey);
  assert.ok(styleKeys.includes('w:b'));
  assert.ok(styleKeys.includes('w:i'));
});

test('fenced code block 保留首个空行，只移除 fence 结尾带来的一个换行', async () => {
  const markdown = '```js\n\nconst x = 1;\n\n```';
  const { html } = await markdownToHTML(markdown);

  assert.equal(html, '<pre><code class="language-js">\nconst x = 1;\n</code></pre>');
});

test('Mermaid Puppeteer 配置使用 headless 临时 profile 并禁用首次启动提示', () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'bdc-puppeteer-config-'));
  try {
    const config = buildPuppeteerConfig(tmpDir);

    assert.equal(config.headless, true);
    assert.equal(config.userDataDir, path.join(tmpDir, 'browser-profile'));
    assert.ok(config.args.includes('--no-first-run'));
    assert.ok(config.args.includes('--no-default-browser-check'));
    assert.ok(config.args.includes('--disable-extensions'));
    assert.ok(config.args.includes('--use-mock-keychain'));
  } finally {
    fs.rmSync(tmpDir, { recursive: true, force: true });
  }
});

test('分离的有序列表在标题后重新编号（使用不同 numbering instance）', async () => {
  const markdown = [
    '1. 第一段第一项',
    '2. 第一段第二项',
    '',
    '## 下一节',
    '',
    '1. 第二段第一项',
    '2. 第二段第二项',
    '',
    '普通段落',
    '',
    '1. 第三段第一项'
  ].join('\n');

  const { children } = await mdToChildren(markdown);

  const orderedParas = children.filter((p) =>
    getNumberingRefs(p).some((ref) => ref.reference === 'numbered-list')
  );

  assert.equal(orderedParas.length, 5, '应有 5 个有序列表项段落');

  const instances = orderedParas.map((p) => getNumberingRefs(p)[0].instance);
  assert.equal(instances[0], instances[1], '同一有序列表内 instance 应相同');
  assert.notEqual(instances[1], instances[2], '标题后新有序列表应使用新 instance');
  assert.equal(instances[2], instances[3], '第二组列表内部 instance 应相同');
  assert.notEqual(instances[3], instances[4], '段落后新有序列表应使用新 instance');
});

test('嵌套有序列表共享同一 numbering instance', () => {
  const html = [
    '<ol>',
    '  <li>外层1',
    '    <ol>',
    '      <li>内层1</li>',
    '      <li>内层2</li>',
    '    </ol>',
    '  </li>',
    '  <li>外层2</li>',
    '</ol>'
  ].join('');

  const { children } = convertHTMLToDocx(html, process.cwd());
  const orderedParas = children.filter((p) =>
    getNumberingRefs(p).some((ref) => ref.reference === 'numbered-list')
  );

  assert.equal(orderedParas.length, 4);
  const instances = orderedParas.map((p) => getNumberingRefs(p)[0].instance);
  assert.equal(new Set(instances).size, 1, '嵌套有序列表应共享同一 instance');
});

test('标题保留 inline 加粗格式', () => {
  const { children } = convertHTMLToDocx(
    '<h2>含 <strong>重点</strong> 的标题</h2>',
    process.cwd()
  );

  assert.equal(children.length, 1);
  const styleNode = children[0].properties.root.find(item => item && item.rootKey === 'w:pStyle');
  assert.ok(styleNode, '标题应带段落样式');
  const runs = children[0].root.filter(child => child && child.rootKey === 'w:r');
  assert.ok(runs.length >= 2, '标题应拆成多个 run');
  const boldRun = runs.find(run =>
    run.properties && run.properties.root.some(item => item && item.rootKey === 'w:b')
  );
  assert.ok(boldRun, '标题内加粗应保留');
});

test('HR 生成带底边框的分隔段落', () => {
  const { children } = convertHTMLToDocx('<p>上</p><hr><p>下</p>', process.cwd());
  assert.equal(children.length, 3);
  const hr = children[1];
  const hasBottomBorder = hr.properties && hr.properties.root && hr.properties.root.some(item => {
    if (!item || item.rootKey !== 'w:pBdr') return false;
    return (item.root || []).some(border => border && border.rootKey === 'w:bottom');
  });
  assert.ok(hasBottomBorder, 'HR 应生成底边框');
});

test('原文 HTML 特殊字符被转义，不进入原始标签', async () => {
  const { html } = await markdownToHTML('段落含 <script>alert(1)</script> 文本');
  assert.match(html, /&lt;script&gt;/);
  assert.doesNotMatch(html, /<script>/);
});

test('链接属性中的引号被 URL 编码', async () => {
  const { html } = await markdownToHTML('[x](https://example.com/a"b)');
  assert.match(html, /href="https:\/\/example\.com\/a%22b"/);
});

test('远程图片无法嵌入时产生 warning', () => {
  const { children, warnings } = convertHTMLToDocx(
    '<p><img src="https://example.com/a.png" alt="remote"></p>',
    process.cwd()
  );
  assert.equal(children.length, 1);
  assert.ok(warnings.some(w => w.includes('远程') || w.includes('无法嵌入')));
  assert.match(collectDocxText(children), /\[图片: remote\]/);
});

test('docx 输出路径在重名时递增后缀', () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'bdc-docx-unique-'));
  try {
    fs.writeFileSync(path.join(tmpDir, 'report.docx'), 'x');
    const first = resolveUniqueDocxPath(tmpDir, 'report');
    assert.equal(path.basename(first), 'report.2.docx');
    fs.writeFileSync(first, 'y');
    const second = resolveUniqueDocxPath(tmpDir, 'report');
    assert.equal(path.basename(second), 'report.3.docx');
  } finally {
    fs.rmSync(tmpDir, { recursive: true, force: true });
  }
});

test('HTML 转 DOCX 只允许读取 Markdown 目录内的相对图片', () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'bdc-image-scope-'));
  try {
    const insideDir = path.join(tmpDir, 'inside');
    const outsideDir = path.join(tmpDir, 'outside');
    fs.mkdirSync(insideDir);
    fs.mkdirSync(outsideDir);

    const png = Buffer.from(
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO6N6t0AAAAASUVORK5CYII=',
      'base64'
    );
    const allowedImage = path.join(insideDir, 'allowed.png');
    const outsideImage = path.join(outsideDir, 'secret.png');
    fs.writeFileSync(allowedImage, png);
    fs.writeFileSync(outsideImage, png);

    const allowed = convertHTMLToDocx('<p><img src="allowed.png" alt="allowed"></p>', insideDir);
    const absolute = convertHTMLToDocx(`<p><img src="${outsideImage}" alt="absolute"></p>`, insideDir);
    const fileUrl = convertHTMLToDocx(`<p><img src="file://${outsideImage}" alt="file"></p>`, insideDir);
    const traversal = convertHTMLToDocx('<p><img src="../outside/secret.png" alt="traversal"></p>', insideDir);

    assert.ok(
      allowed.children[0].root.some(child => child && child.rootKey === 'w:r'),
      '目录内相对图片应正常生成 run'
    );
    assert.match(collectDocxText(absolute.children), /\[图片: absolute\]/);
    assert.match(collectDocxText(fileUrl.children), /\[图片: file\]/);
    assert.match(collectDocxText(traversal.children), /\[图片: traversal\]/);
    assert.ok(absolute.warnings.length > 0);
    assert.ok(fileUrl.warnings.length > 0);
    assert.ok(traversal.warnings.length > 0);
  } finally {
    fs.rmSync(tmpDir, { recursive: true, force: true });
  }
});

// ---- 端到端：Markdown -> 打包后的 DOCX XML ----

// Resolve packages exactly as the converter does, so docx classes are the same instances.
const converterRequire = require('node:module').createRequire(require.resolve('../bruce_doc_converter/md_to_docx/index'));
const JSZip = converterRequire('jszip');
const { Document, Packer } = converterRequire('docx');
const { createStyles, createNumbering, createMargins } = require('../bruce_doc_converter/md_to_docx/styles');
const { Diagnostics } = require('../bruce_doc_converter/md_to_docx/diagnostics');
const { validateDocx } = require('../bruce_doc_converter/md_to_docx/docx-validate');
const { convertMarkdownToDocx } = require('../bruce_doc_converter/md_to_docx/index');

async function mdToXml(markdown, basePath = process.cwd()) {
  const diagnostics = new Diagnostics();
  const { html, formulas } = await markdownToHTML(markdown, diagnostics);
  const { children, footnotes } = convertHTMLToDocx(html, basePath, { formulas, diagnostics });
  const doc = new Document({
    styles: createStyles(), numbering: createNumbering(), footnotes,
    sections: [{ properties: { page: { margin: createMargins() } }, children }]
  });
  const buffer = await Packer.toBuffer(doc);
  const zip = await JSZip.loadAsync(buffer);
  const read = async (name) => (zip.file(name) ? zip.file(name).async('string') : '');
  return { xml: await read('word/document.xml'), footnotesXml: await read('word/footnotes.xml'), diagnostics: diagnostics.items, buffer };
}
const plainText = (xml) => xml.replace(/<\/w:p>/g, '\n').replace(/<[^>]+>/g, '');
const codes = (items) => items.map(item => item.code);

test('中文源文件换行不插入空格，英文换行保留词间空格', async () => {
  const { xml } = await mdToXml('第一行中文\n第二行中文。English line\ncontinues here.');
  assert.match(plainText(xml), /第一行中文第二行中文。English line continues here\./);
});

test('LaTeX 公式转换为 Word 原生公式，金额不误判为公式', async () => {
  const { xml, diagnostics } = await mdToXml([
    '公式 $E=mc^2$，价格 $5 and $10。',
    '',
    '$$',
    String.raw`\begin{pmatrix} 1 & 0 \\ 0 & 1 \end{pmatrix}`,
    '$$'
  ].join('\n'));
  assert.equal((xml.match(/<m:oMath>/g) || []).length, 2);
  assert.equal((xml.match(/<m:oMathPara>/g) || []).length, 1);
  assert.equal((xml.match(/<m:mr>/g) || []).length, 2, '矩阵应保留两行');
  assert.match(plainText(xml), /价格 \$5 and \$10/);
  assert.deepEqual(diagnostics, []);
});

test('不支持的公式保留源码并报告带行号的诊断', async () => {
  const { xml, diagnostics } = await mdToXml(String.raw`第一行

坏公式 $\color{red}{x}$`);
  assert.ok(plainText(xml).includes(String.raw`[公式未转换] $\color{red}{x}$`));
  assert.deepEqual(diagnostics.map(d => [d.code, d.severity, d.line]), [['MATH_NOT_CONVERTED', 'warning', 3]]);
});

test('脚注生成 Word 原生脚注，重复引用使用 NOTEREF，定义内容不丢失', async () => {
  const { xml, footnotesXml, diagnostics } = await mdToXml('正文[^n]，再次[^n]。\n\n[^n]: 脚注 **内容**。\n\n    第二段。');
  assert.equal((xml.match(/<w:footnoteReference /g) || []).length, 1);
  assert.match(xml, /NOTEREF note_1/);
  assert.match(plainText(footnotesXml), /脚注 内容。/);
  assert.match(plainText(footnotesXml), /第二段。/);
  assert.doesNotMatch(plainText(xml), /\^n/);
  assert.deepEqual(diagnostics, []);
});

test('未定义与未引用的脚注保留原文并分级报告', async () => {
  const { xml, diagnostics } = await mdToXml('正文[^missing]。\n\n[^orphan]: 未引用。');
  assert.match(plainText(xml), /正文\[\^missing\]。/);
  assert.match(plainText(xml), /未引用。/);
  assert.deepEqual(diagnostics.map(d => [d.code, d.severity]).sort(), [['FOOTNOTE_UNDEFINED', 'warning'], ['FOOTNOTE_UNUSED', 'info']]);
});

test('文档内链接跳转到标题书签，缺失目标保留文字并报告', async () => {
  const { xml, diagnostics } = await mdToXml('# 概述\n\n见[概述](#概述)与[缺失](#不存在)。');
  assert.match(xml, /<w:bookmarkStart w:name="heading_1" w:id="\d+"\/>/);
  assert.match(xml, /<w:hyperlink [^>]*w:anchor="heading_1"/);
  assert.deepEqual(diagnostics.map(d => [d.code, d.line]), [['LINK_UNAVAILABLE', 3]]);
  assert.match(diagnostics[0].message, /#不存在/);
});

test('引用块内的列表保持列表结构和引用样式', async () => {
  const { xml } = await mdToXml('> 引用段落\n>\n> - 列表一\n> - 列表二');
  const paragraphs = xml.split('</w:p>').filter(p => /列表[一二]/.test(p));
  assert.equal(paragraphs.length, 2, '列表项不应被合并为一段');
  for (const p of paragraphs) {
    assert.match(p, /<w:pStyle w:val="Quote"\/>/);
    assert.match(p, /<w:numPr>/);
  }
});

test('表格使用按内容估算的固定列宽并保留对齐方式，相邻表格不合并', async () => {
  const { xml } = await mdToXml([
    '| 编号 | 状态 | 说明 |', '| :--- | :---: | ---: |', '| A-01 | 完成 | 这是一段相当长的说明文字，用于测试列宽分配 |',
    '', '| 第二张表 | x |', '| --- | --- |', '| 1 | 2 |'
  ].join('\n'));
  const first = xml.split('</w:tbl>')[0];
  const widths = [...first.matchAll(/<w:gridCol w:w="(\d+)"\/>/g)].map(m => Number(m[1]));
  assert.equal(widths.reduce((a, b) => a + b, 0), 11906 - 1417 * 2);
  assert.ok(widths[2] > widths[0] && widths[2] > widths[1], '长说明列应更宽');
  assert.match(first, /<w:tblLayout w:type="fixed"\/>/);
  const bodyRow = first.split('<w:tr>').pop();
  assert.deepEqual([...bodyRow.matchAll(/<w:jc w:val="(\w+)"\/>/g)].map(m => m[1]), ['left', 'center', 'right']);
  assert.match(xml, /<\/w:tbl><w:p>((?!<w:tbl>).)*<\/w:p><w:tbl>/s, '相邻表格之间应有分隔段落');
});

test('正文首行缩进只作用于正文段落，列表和标题不继承', async () => {
  const { xml } = await mdToXml('# 标题\n\n正文段落。\n\n- 列表项');
  const body = xml.split('</w:p>').find(p => p.includes('正文段落'));
  const item = xml.split('</w:p>').find(p => p.includes('列表项'));
  assert.match(body, /<w:pStyle w:val="BodyText"\/>/);
  assert.doesNotMatch(item, /BodyText|w:firstLine="[1-9]/);
});

test('图片按自然尺寸显示，不放大小图，超高图片按页面高度缩小', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'bdc-image-size-'));
  try {
    const png = (width, height) => {
      const header = Buffer.alloc(33);
      Buffer.from('89504e470d0a1a0a0000000d49484452', 'hex').copy(header);
      header.writeUInt32BE(width, 16);
      header.writeUInt32BE(height, 20);
      return header;
    };
    fs.writeFileSync(path.join(tmpDir, 'small.png'), png(32, 16));
    fs.writeFileSync(path.join(tmpDir, 'tall.png'), png(400, 4000));
    const { xml } = await mdToXml('![s](small.png)\n\n![t](tall.png)', tmpDir);
    // 96 DPI 像素 -> EMU
    const extents = [...xml.matchAll(/<wp:extent cx="(\d+)" cy="(\d+)"\/>/g)].map(m => [Number(m[1]) / 9525, Number(m[2]) / 9525]);
    assert.deepEqual(extents[0], [32, 16]);
    assert.ok(extents[1][1] <= 740, `超高图片应按页面高度缩小: ${extents[1]}`);
    assert.equal(extents[1][0], Math.round(extents[1][1] / 10));
  } finally {
    fs.rmSync(tmpDir, { recursive: true, force: true });
  }
});

test('DOCX 完整性校验接受正常输出并拒绝损坏的包', async () => {
  const { buffer } = await mdToXml('# 标题\n\n正文[^a]\n\n[^a]: 注。');
  await validateDocx(buffer);
  const zip = await JSZip.loadAsync(buffer);
  zip.file('word/document.xml', '<w:document><broken></w:document>');
  await assert.rejects(validateDocx(await zip.generateAsync({ type: 'nodebuffer' })));
  zip.remove('word/document.xml');
  await assert.rejects(validateDocx(await zip.generateAsync({ type: 'nodebuffer' })), /document\.xml/);
});

test('严格模式只拒绝 warning 级诊断，info 级不阻止导出', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'bdc-strict-'));
  try {
    const infoOnly = path.join(tmpDir, 'info.md');
    const degraded = path.join(tmpDir, 'degraded.md');
    fs.writeFileSync(infoOnly, '正文。\n\n[^orphan]: 未引用。');
    fs.writeFileSync(degraded, String.raw`坏公式 $\color{red}{x}$`);
    const ok = await convertMarkdownToDocx(infoOnly, tmpDir, { strict: true });
    assert.equal(ok.success, true, JSON.stringify(ok));
    assert.equal(ok.warnings, undefined);
    assert.deepEqual(codes(ok.diagnostics), ['FOOTNOTE_UNUSED']);
    const rejected = await convertMarkdownToDocx(degraded, tmpDir, { strict: true });
    assert.equal(rejected.error_code, 'CONTENT_INCOMPLETE');
    assert.deepEqual(codes(rejected.diagnostics), ['MATH_NOT_CONVERTED']);
    assert.match(rejected.warnings[0], /第 1 行/);
    assert.equal(fs.readdirSync(tmpDir).filter(name => name.startsWith('degraded')).length, 1, '只应存在源 Markdown');
  } finally {
    fs.rmSync(tmpDir, { recursive: true, force: true });
  }
});

test('诊断数量有上限且截断时保留阻断级别', () => {
  const diagnostics = new Diagnostics(2);
  diagnostics.add('A', 'a', 'info');
  diagnostics.add('B', 'b', 'info');
  diagnostics.add('C', 'c', 'warning');
  assert.equal(diagnostics.items.length, 2);
  assert.deepEqual(diagnostics.items[1], { code: 'DIAGNOSTICS_TRUNCATED', severity: 'warning', message: '另有 2 条诊断被省略。' });
});
