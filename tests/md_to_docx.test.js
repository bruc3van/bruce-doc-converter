const test = require('node:test');
const assert = require('node:assert/strict');

const { markdownToHTML } = require('../scripts/md_to_docx/markdown-converter');
const { convertHTMLToDocx } = require('../scripts/md_to_docx/html-converter');

test('markdown 表格支持转义管道符', async () => {
  const markdown = [
    '| col1 | col2 |',
    '| --- | --- |',
    '| a\\|b | c |'
  ].join('\n');

  const html = await markdownToHTML(markdown);

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

  const html = await markdownToHTML(markdown);

  assert.match(html, /href="https:\/\/example\.com\/a_\(b\)\.png"/);
  assert.match(html, /src="https:\/\/example\.com\/a_\(b\)\.png"/);
});

test('HTML 转 DOCX 保留混合内联样式和超链接目标', () => {
  const children = convertHTMLToDocx(
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
  const html = await markdownToHTML(markdown);

  assert.equal(html, '<pre><code class="language-js">\nconst x = 1;\n</code></pre>');
});
