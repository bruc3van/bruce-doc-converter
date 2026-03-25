/**
 * Mermaid 渲染器
 * 使用 mmdc (mermaid-cli) 在 Node.js 环境渲染为图片 data URL
 */

const fs = require('fs');
const os = require('os');
const path = require('path');
const { spawn } = require('child_process');

const MMDC_TIMEOUT_MS = Number.parseInt(process.env.BRUCE_DOC_CONVERTER_MMDC_TIMEOUT_MS || '90000', 10);
const MMDC_FORMAT = (process.env.BRUCE_DOC_CONVERTER_MMDC_FORMAT || 'png').toLowerCase();
const MMDC_SCALE = Number.parseFloat(process.env.BRUCE_DOC_CONVERTER_MMDC_SCALE || '2');
const TEMP_DIR_PREFIX = 'bruce-doc-converter-mmdc-';

const MERMAID_CONFIG = {
  startOnLoad: false,
  securityLevel: 'strict',
  theme: 'default',
  fontFamily: 'Arial',
  flowchart: {
    htmlLabels: false
  }
};

const LINUX_PUPPETEER_ARGS = ['--no-sandbox', '--disable-setuid-sandbox'];

/**
 * 将 Mermaid 源码渲染为图片 data URL
 * 默认输出 png，兼容 Word 的渲染能力
 * @param {string} mermaidCode Mermaid 源码
 * @returns {Promise<{success: boolean, dataUrl?: string, width?: number, height?: number, error?: string}>}
 */
async function renderMermaidToDataUrl(mermaidCode) {
  if (!mermaidCode || typeof mermaidCode !== 'string') {
    return { success: false, error: 'Mermaid 代码为空' };
  }

  const outputFormat = MMDC_FORMAT === 'svg' ? 'svg' : 'png';
  const mmdcCommand = resolveMMDCCommand();
  let tempDir = '';

  try {
    tempDir = await fs.promises.mkdtemp(path.join(os.tmpdir(), TEMP_DIR_PREFIX));

    const inputPath = path.join(tempDir, 'diagram.mmd');
    const outputPath = path.join(tempDir, `diagram.${outputFormat}`);
    const configPath = path.join(tempDir, 'mermaid-config.json');
    const puppeteerConfigPath = path.join(tempDir, 'puppeteer-config.json');

    await fs.promises.writeFile(inputPath, mermaidCode, 'utf-8');
    await fs.promises.writeFile(configPath, JSON.stringify(MERMAID_CONFIG), 'utf-8');
    await fs.promises.writeFile(puppeteerConfigPath, JSON.stringify(buildPuppeteerConfig()), 'utf-8');

    const args = [
      '-i', inputPath,
      '-o', outputPath,
      '-e', outputFormat,
      '-c', configPath,
      '-p', puppeteerConfigPath,
      '-b', 'transparent',
      '-s', Number.isFinite(MMDC_SCALE) && MMDC_SCALE > 0 ? String(MMDC_SCALE) : '2',
      '-q'
    ];

    const result = await runCommand(mmdcCommand, args, tempDir, MMDC_TIMEOUT_MS);
    if (result.timedOut) {
      return { success: false, error: `mmdc 渲染超时（>${MMDC_TIMEOUT_MS}ms）` };
    }

    if (result.errorCode === 'ENOENT') {
      return {
        success: false,
        error: '未找到 mmdc 可执行文件，请在 scripts/md_to_docx 目录执行 npm install'
      };
    }

    if (result.code !== 0) {
      const detail = (result.stderr || result.stdout || '').trim();
      return {
        success: false,
        error: detail ? `mmdc 渲染失败: ${detail}` : 'mmdc 渲染失败（未知错误）'
      };
    }

    if (outputFormat === 'svg') {
      const svg = await fs.promises.readFile(outputPath, 'utf-8');
      if (!svg.includes('<svg')) {
        return { success: false, error: 'mmdc 未生成有效 SVG 输出' };
      }

      const viewBoxSize = parseSvgViewBoxSize(svg);
      return {
        success: true,
        dataUrl: `data:image/svg+xml;base64,${Buffer.from(svg, 'utf-8').toString('base64')}`,
        width: viewBoxSize.width,
        height: viewBoxSize.height
      };
    }

    const pngBuffer = await fs.promises.readFile(outputPath);
    if (!pngBuffer.length) {
      return { success: false, error: 'mmdc 未生成有效 PNG 输出' };
    }

    const pngSize = parsePngSize(pngBuffer);
    return {
      success: true,
      dataUrl: `data:image/png;base64,${pngBuffer.toString('base64')}`,
      width: pngSize.width,
      height: pngSize.height
    };
  } catch (error) {
    return {
      success: false,
      error: error instanceof Error ? error.message : String(error)
    };
  } finally {
    if (tempDir) {
      await fs.promises.rm(tempDir, { recursive: true, force: true }).catch(() => {});
    }
  }
}

function resolveMMDCCommand() {
  const candidates = [];
  const ext = process.platform === 'win32' ? '.cmd' : '';

  if (process.env.BRUCE_DOC_CONVERTER_MMDC_PATH) {
    candidates.push(process.env.BRUCE_DOC_CONVERTER_MMDC_PATH);
  }

  candidates.push(path.join(__dirname, 'node_modules', '.bin', `mmdc${ext}`));
  candidates.push(path.join(process.cwd(), 'node_modules', '.bin', `mmdc${ext}`));

  const sharedRoot = resolveSharedNodeRoot();
  if (sharedRoot) {
    candidates.push(path.join(sharedRoot, 'md_to_docx', 'node_modules', '.bin', `mmdc${ext}`));
  }

  const nodePath = process.env.NODE_PATH || '';
  if (nodePath) {
    for (const entry of nodePath.split(path.delimiter)) {
      const normalized = (entry || '').trim();
      if (!normalized) continue;
      candidates.push(path.join(normalized, '.bin', `mmdc${ext}`));
    }
  }

  try {
    const mermaidPkg = require.resolve('@mermaid-js/mermaid-cli/package.json');
    const pkgDir = path.dirname(mermaidPkg);
    const nodeModulesDir = path.resolve(pkgDir, '..', '..', '..');
    candidates.push(path.join(nodeModulesDir, '.bin', `mmdc${ext}`));
  } catch (error) {
    // noop: 依赖可能尚未安装
  }

  for (const command of candidates) {
    if (command && fs.existsSync(command)) {
      return command;
    }
  }

  return 'mmdc';
}

function resolveSharedNodeRoot() {
  if (process.env.BRUCE_DOC_CONVERTER_NODE_HOME) {
    return process.env.BRUCE_DOC_CONVERTER_NODE_HOME;
  }

  if (process.platform === 'win32') {
    const base = process.env.LOCALAPPDATA || process.env.APPDATA || path.join(os.homedir(), 'AppData', 'Local');
    return path.join(base, 'BruceDocConverter', 'node');
  }

  return path.join(os.homedir(), '.bruce-doc-converter', 'node');
}

function buildPuppeteerConfig() {
  if (process.platform === 'linux') {
    return { args: LINUX_PUPPETEER_ARGS };
  }
  return {};
}

function runCommand(command, args, cwd, timeoutMs) {
  return new Promise((resolve) => {
    let finalCommand = command;
    let finalArgs = args;

    if (process.platform === 'win32' && /\.(cmd|bat)$/i.test(command)) {
      finalCommand = process.env.ComSpec || 'cmd.exe';
      finalArgs = ['/d', '/s', '/c', command, ...args];
    }

    const child = spawn(finalCommand, finalArgs, {
      cwd,
      stdio: ['ignore', 'pipe', 'pipe'],
      windowsHide: true,
      shell: false
    });

    let stdout = '';
    let stderr = '';
    let timedOut = false;
    let finished = false;

    const timer = setTimeout(() => {
      timedOut = true;
      child.kill('SIGKILL');
    }, timeoutMs);

    child.stdout.on('data', (chunk) => {
      stdout += chunk.toString();
    });

    child.stderr.on('data', (chunk) => {
      stderr += chunk.toString();
    });

    child.on('error', (error) => {
      if (finished) return;
      finished = true;
      clearTimeout(timer);
      resolve({
        code: null,
        stdout,
        stderr,
        timedOut,
        errorCode: error && error.code ? error.code : '',
      });
    });

    child.on('close', (code) => {
      if (finished) return;
      finished = true;
      clearTimeout(timer);
      resolve({
        code,
        stdout,
        stderr,
        timedOut,
        errorCode: ''
      });
    });
  });
}

function parsePngSize(buffer) {
  if (!Buffer.isBuffer(buffer) || buffer.length < 24) {
    return { width: 0, height: 0 };
  }

  const signature = buffer.toString('hex', 0, 8);
  if (signature !== '89504e470d0a1a0a') {
    return { width: 0, height: 0 };
  }

  return {
    width: buffer.readUInt32BE(16),
    height: buffer.readUInt32BE(20)
  };
}

function parseSvgViewBoxSize(svgText) {
  const match = svgText.match(/viewBox="([^"]+)"/i);
  if (!match) return { width: 0, height: 0 };

  const values = match[1]
    .trim()
    .split(/[\s,]+/)
    .map(value => Number.parseFloat(value));

  if (values.length !== 4 || values[2] <= 0 || values[3] <= 0 || values.some(value => !Number.isFinite(value))) {
    return { width: 0, height: 0 };
  }

  return {
    width: Math.round(values[2]),
    height: Math.round(values[3])
  };
}

module.exports = {
  renderMermaidToDataUrl
};
