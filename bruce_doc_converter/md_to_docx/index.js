#!/usr/bin/env node
/**
 * Markdown 转 DOCX 命令行工具
 * 用法: node index.js <input.md> [output_dir]
 */

const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

function supportedNodeVersion(version = process.versions.node) {
  const major = Number(version.split('.')[0]);
  return Number.isInteger(major) && major >= 22;
}

function writeUniqueDocx(outputDir, baseName, buffer) {
  // Reserve before writing so concurrent conversions never replace each other.
  let outputPath;
  while (true) {
    outputPath = resolveUniqueDocxPath(outputDir, baseName);
    try {
      fs.closeSync(fs.openSync(outputPath, 'wx'));
      break;
    } catch (error) {
      if (error.code !== 'EEXIST') throw error;
    }
  }
  const temporary = path.join(outputDir, `.bdc-${crypto.randomUUID()}.tmp`);
  try {
    fs.writeFileSync(temporary, buffer, { flag: 'wx' });
    fs.renameSync(temporary, outputPath);
    return outputPath;
  } catch (error) {
    fs.rmSync(outputPath, { force: true });
    throw error;
  } finally {
    fs.rmSync(temporary, { force: true });
  }
}

/**
 * 生成不覆盖已有文件的输出路径
 * @param {string} outputDir
 * @param {string} baseName
 * @returns {string}
 */
function resolveUniqueDocxPath(outputDir, baseName) {
  let outputPath = path.join(outputDir, `${baseName}.docx`);
  if (!fs.existsSync(outputPath)) {
    return outputPath;
  }

  let counter = 2;
  while (true) {
    outputPath = path.join(outputDir, `${baseName}.${counter}.docx`);
    if (!fs.existsSync(outputPath)) {
      return outputPath;
    }
    counter += 1;
  }
}

/**
 * 将 Markdown 文件转换为 DOCX
 * @param {string} inputPath - 输入的 Markdown 文件路径
 * @param {string} outputDir - 输出目录（可选）
 * @returns {Object} 转换结果
 */
async function convertMarkdownToDocx(inputPath, outputDir, options = {}) {
  try {
    if (!supportedNodeVersion()) {
      return { success: false, error_code: 'NODE_VERSION_UNSUPPORTED', error: 'Markdown conversion requires Node.js >=22.0.' };
    }
    const { Document, Packer } = require('docx');
    const { markdownToHTML } = require('./markdown-converter');
    const { convertHTMLToDocx } = require('./html-converter');
    const { createStyles, createNumbering, createMargins } = require('./styles');
    const { Diagnostics } = require('./diagnostics');
    const { validateDocx } = require('./docx-validate');

    // 验证输入文件
    if (!fs.existsSync(inputPath)) {
      return { success: false, error: `文件不存在: ${inputPath}` };
    }

    // 读取 Markdown 文件（容忍 UTF-8 BOM）
    let markdown = fs.readFileSync(inputPath, 'utf-8');
    if (markdown.charCodeAt(0) === 0xFEFF) {
      markdown = markdown.slice(1);
    }

    // Markdown -> HTML，两阶段共享同一份诊断
    const diagnostics = new Diagnostics();
    const { html, formulas } = await markdownToHTML(markdown, diagnostics);

    // HTML -> DOCX 组件（传入 Markdown 文件所在目录，用于解析相对路径图片）
    const mdDir = path.dirname(path.resolve(inputPath));
    const { children: docxChildren, footnotes, updateFields } = convertHTMLToDocx(html, mdDir, { formulas, diagnostics });
    // 只有 warning 级诊断（内容缺失或降级）进入 warnings 并阻止严格导出；info 仅作提示
    const warnings = diagnostics.warnings();
    const diagnosticItems = diagnostics.items;
    if (options.strict && warnings.length) {
      return { success: false, error_code: 'CONTENT_INCOMPLETE', error: '严格模式拒绝包含内容缺失或降级警告的转换。', warnings, diagnostics: diagnosticItems };
    }

    // 创建文档
    const doc = new Document({
      styles: createStyles(),
      numbering: createNumbering(),
      footnotes,
      // 重复引用同一脚注时使用 NOTEREF 域，打开时请求更新
      ...(updateFields ? { features: { updateFields: true } } : {}),
      sections: [{
        properties: { page: { margin: createMargins() } },
        children: docxChildren
      }]
    });

    // 确定输出路径
    const inputDir = path.dirname(inputPath);
    const baseName = path.basename(inputPath, path.extname(inputPath));

    let finalOutputDir;
    if (outputDir) {
      finalOutputDir = outputDir;
    } else {
      finalOutputDir = path.join(inputDir, 'Word');
    }

    // 创建输出目录
    if (!fs.existsSync(finalOutputDir)) {
      fs.mkdirSync(finalOutputDir, { recursive: true });
    }

    // 生成并保存文档
    const buffer = await Packer.toBuffer(doc);
    try {
      await validateDocx(buffer);
    } catch (error) {
      return { success: false, error_code: 'NODE_CONVERSION_FAILED', error: `生成的 DOCX 未通过完整性校验: ${error.message}` };
    }
    const outputPath = writeUniqueDocx(finalOutputDir, baseName, buffer);

    const result = {
      success: true,
      output_path: outputPath,
      message: `转换成功: ${outputPath}`
    };
    if (warnings.length > 0) {
      result.warnings = warnings;
    }
    if (diagnosticItems.length > 0) {
      result.diagnostics = diagnosticItems;
    }
    return result;

  } catch (error) {
    return {
      success: false,
      error: `转换错误: ${error.message}`
    };
  }
}

// 命令行入口
async function main() {
  const args = process.argv.slice(2);

  if (args.includes('-h') || args.includes('--help')) {
    console.log(JSON.stringify({
      success: true,
      usage: 'node index.js <input.md> [output_dir]'
    }, null, 2));
    process.exit(0);
  }

  if (args.length < 1) {
    console.log(JSON.stringify({
      success: false,
      error: '用法: node index.js <input.md> [output_dir]'
    }));
    process.exit(1);
  }

  const inputPath = args[0];
  const outputDir = args[1] || '';

  const result = await convertMarkdownToDocx(inputPath, outputDir || null, { strict: process.env.BRUCE_DOC_CONVERTER_STRICT === '1' });
  console.log(JSON.stringify(result, null, 2));

  process.exit(result.success ? 0 : 1);
}

if (require.main === module) {
  main();
}

module.exports = {
  convertMarkdownToDocx,
  resolveUniqueDocxPath,
  supportedNodeVersion
};
