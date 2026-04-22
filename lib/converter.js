'use strict';

const { spawn } = require('child_process');
const path = require('path');
const fs = require('fs');
const os = require('os');
const { log } = require('./logger');
const { buildFontConfig } = require('./fonts');
const { buildLoProfile } = require('./loFontSetup');
const { prepareDocx } = require('./wpsCompat');
const { convertFileNative } = require('./nativeEngine');

function summarizeProcessOutput(output, limit = 280) {
  if (!output) return '';
  const singleLine = output.replace(/\s+/g, ' ').trim();
  if (!singleLine) return '';
  return singleLine.length <= limit
    ? singleLine
    : `${singleLine.slice(0, limit)}...`;
}

function _formatSofficeFailure({ code, signal, stdout, stderr }) {
  const reason = signal
    ? `terminated by signal ${signal}`
    : typeof code === 'number'
      ? `exited with code ${code}`
      : 'exited unexpectedly';

  const stderrSummary = summarizeProcessOutput(stderr);
  const stdoutSummary = summarizeProcessOutput(stdout);
  const outputSummary = stderrSummary || stdoutSummary;
  const detail = outputSummary ? `; output: ${outputSummary}` : '';

  return new Error(`soffice ${reason}${detail}`);
}

function _resolvePreparedInput(inputPath, tmpDir, opts = {}) {
  const {
    wpsCompat = true,
    prepareDocxImpl = prepareDocx,
  } = opts;

  if (!wpsCompat) return inputPath;
  return prepareDocxImpl(inputPath, tmpDir);
}

/**
 * Convert a single .docx file to PDF.
 */
async function convertFile(inputPath, outputPath, opts = {}) {
  const {
    engine = 'libreoffice',
    sobinary = 'soffice',
    fontDir = null,
    timeoutMs = 60_000,
    overwrite = false,
    silent = false,
    wpsCompat = true,
  } = opts;

  if (engine === 'native') {
    return convertFileNative(inputPath, outputPath, {
      overwrite,
      silent,
      timeoutMs,
      chromeBinary: opts.chromeBinary || null,
      nativeLayout: opts.nativeLayout || 'structured',
    });
  }

  if (!overwrite && fs.existsSync(outputPath)) {
    if (!silent) log.skip(outputPath);
    return true;
  }

  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'docx2pdf-'));

  try {
    // 1. 初始化隔离 Profile，避免历史配置影响排版
    const profileDir = path.join(tmpDir, 'lo-profile');
    buildLoProfile(profileDir, sobinary);

    // 2. WPS 与 XML 预处理 (字体替换 + 行距释放)
    const preparedInput = _resolvePreparedInput(inputPath, tmpDir, { wpsCompat });

    // 3. 构建临时 FontConfig (作为额外的双重保障)
    const fontConfPath = buildFontConfig(tmpDir, fontDir);

    if (!silent) log.converting(path.basename(inputPath));

    // 4. 调用 LibreOffice (带上所有核心修复参数)
    await runLibreOffice({
      sobinary,
      inputPath: preparedInput,
      outDir: tmpDir,
      profileDir,
      fontConfPath,
      timeoutMs,
    });

    const actualExpectedPdf = path.join(tmpDir, path.basename(preparedInput, '.docx') + '.pdf');
    // console.log('Expecting PDF at:', actualExpectedPdf);

    if (!fs.existsSync(actualExpectedPdf)) {
      console.log('Files in tmpDir:', fs.readdirSync(tmpDir));
      throw new Error('LibreOffice finished but no PDF was produced.');
    }

    fs.mkdirSync(path.dirname(outputPath), { recursive: true });
    fs.renameSync(actualExpectedPdf, outputPath);

    if (!silent) log.success(outputPath);
    return true;
  } catch (err) {
    log.error(`${path.basename(inputPath)}: ${err.message}`);
    return false;
  } finally {
    fs.rmSync(tmpDir, { recursive: true, force: true });
  }
}

async function convertDir(inputDir, outputDir, opts = {}) {
  const files = fs.readdirSync(inputDir)
    .filter(f => f.toLowerCase().endsWith('.docx'))
    .map(f => path.join(inputDir, f));

  if (files.length === 0) {
    log.warn(`No .docx files found in ${inputDir}`);
    return;
  }

  log.info(`Found ${files.length} file(s) → ${outputDir}`);
  fs.mkdirSync(outputDir, { recursive: true });

  let ok = 0;
  let fail = 0;
  for (const file of files) {
    const outFile = path.join(outputDir, path.basename(file, '.docx') + '.pdf');
    const success = await convertFile(file, outFile, opts);
    success ? ok++ : fail++;
  }
  log.summary(ok, fail);
}

function runLibreOffice({ sobinary, inputPath, outDir, profileDir, fontConfPath, timeoutMs }) {
  return new Promise((resolve, reject) => {
    const args = [
      '--headless',
      '--norestore',
      '--nofirststartwizard',
      `-env:UserInstallation=file://${profileDir}`,
      '--convert-to', 'pdf',
      '--outdir', outDir,
      inputPath,
    ];

    const env = {
      ...process.env,
      FONTCONFIG_FILE: fontConfPath,
      // Headless 环境隔离
      LIBO_HEADLESS: '1',
      JAVA_TOOL_OPTIONS: '',
    };

    const proc = spawn(sobinary, args, { env, stdio: ['ignore', 'pipe', 'pipe'] });

    let stdout = '';
    let stderr = '';
    proc.stdout.on('data', d => { stdout += d.toString(); });
    proc.stderr.on('data', d => { stderr += d.toString(); });

    const timer = setTimeout(() => {
      proc.kill('SIGKILL');
      reject(new Error(`Timed out after ${timeoutMs / 1000}s`));
    }, timeoutMs);

    proc.on('close', (code, signal) => {
      clearTimeout(timer);
      if (code === 0) resolve();
      else reject(_formatSofficeFailure({ code, signal, stdout, stderr }));
    });

    proc.on('error', err => {
      clearTimeout(timer);
      reject(new Error(`Failed to start: ${err.message}`));
    });
  });
}

module.exports = { convertFile, convertDir, _formatSofficeFailure, _resolvePreparedInput };
