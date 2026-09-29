/**
 * 写出前校验生成的 DOCX：压缩包结构、各 XML 部件格式以及内部关系引用。
 * 移植自 bruce-md2word 的 src/runtime/artifact.ts。
 */
const path = require('path');
const JSZip = require('jszip');
const { SaxesParser } = require('saxes');

const REQUIRED_PARTS = ['[Content_Types].xml', '_rels/.rels', 'word/document.xml', 'word/_rels/document.xml.rels'];

/** Namespace-aware well-formedness check; returns each Relationship's attributes. */
function parseXml(xml) {
  const relationships = [];
  const parser = new SaxesParser({ xmlns: true, position: false });
  parser.on('error', (error) => { throw error; });
  parser.on('opentag', (tag) => {
    if (tag.name === 'Relationship') {
      relationships.push(Object.fromEntries(Object.entries(tag.attributes).map(([name, attr]) => [name, attr.value])));
    }
  });
  parser.write(xml).close();
  return relationships;
}

/** @throws {Error} when the package is damaged or references a missing part */
async function validateDocx(data) {
  const zip = await JSZip.loadAsync(data, { checkCRC32: true });
  for (const file of REQUIRED_PARTS) {
    if (!zip.file(file)) throw new Error(`缺少 DOCX 部件 ${file}`);
  }
  const relationships = new Map();
  const documents = new Map();
  for (const entry of Object.values(zip.files)) {
    if (entry.dir || !/\.(xml|rels)$/.test(entry.name)) continue;
    const xml = await entry.async('string');
    const rels = parseXml(xml);
    if (!entry.name.endsWith('.rels')) {
      documents.set(entry.name, xml);
      continue;
    }
    const root = entry.name === '_rels/.rels';
    const base = root ? '' : path.posix.dirname(path.posix.dirname(entry.name));
    const owner = root ? '' : path.posix.join(base, path.posix.basename(entry.name, '.rels'));
    const ids = new Set();
    for (const rel of rels) {
      if (!rel.Id || !rel.Target || ids.has(rel.Id)) throw new Error(`${entry.name} 中存在无效关系`);
      ids.add(rel.Id);
      if (rel.TargetMode !== 'External') {
        const resolved = path.posix.normalize(path.posix.join(base, rel.Target));
        if (resolved.startsWith('../') || !zip.file(resolved)) throw new Error(`${entry.name} 引用了不存在的部件 ${rel.Target}`);
      }
    }
    relationships.set(owner, ids);
  }
  for (const [name, xml] of documents) {
    for (const match of xml.matchAll(/\br:(?:embed|id|link)="([^"]+)"/g)) {
      const ids = relationships.get(name);
      if (!ids || !ids.has(match[1])) throw new Error(`${name} 引用了不存在的关系 ${match[1]}`);
    }
  }
}

module.exports = { validateDocx };
