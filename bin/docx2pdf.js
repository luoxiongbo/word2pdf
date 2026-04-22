#!/usr/bin/env node

'use strict';

const { program } = require('commander');
const path = require('path');
const fs = require('fs');
const { convertFile, convertDir } = require('../lib/converter');
const { checkLibreOffice } = require('../lib/preflight');
const { log } = require('../lib/logger');
const pkg = require('../package.json');

program
  .name('docx2pdf')
  .description('Convert .docx files to PDF using LibreOffice — no cloud, no account.')
  .version(pkg.version)
  .argument('<input>', '.docx file or directory of .docx files')
  .option('--engine <engine>', 'conversion engine: libreoffice|native', 'libreoffice')
  .option('--native-layout <mode>', 'native layout mode: structured|simple|rich', 'structured')
  .option('-o, --output <path>', 'output file or directory (default: same as input)')
  .option('-f, --font-dir <path>', 'extra font directory to register with LibreOffice')
  .option('--chrome <path>', 'custom Chrome/Chromium binary path for native engine')
  .option('--libreoffice <path>', 'custom path to LibreOffice/soffice binary')
  .option('--timeout <seconds>', 'per-file timeout in seconds', '60')
  .option('--overwrite', 'overwrite existing PDFs', false)
  .option('--no-wps-compat', 'disable WPS compatibility preprocessing (not recommended)')
  .option('--silent', 'suppress all output except errors', false)
  .action(async (input, opts) => {
    if (!['libreoffice', 'native'].includes(opts.engine)) {
      log.error(`Unsupported engine: ${opts.engine}. Use "libreoffice" or "native".`);
      process.exit(1);
    }
    if (!['structured', 'simple', 'rich'].includes(opts.nativeLayout)) {
      log.error(`Unsupported native layout: ${opts.nativeLayout}. Use "structured", "simple" or "rich".`);
      process.exit(1);
    }

    // ── 0. Pre-flight ─────────────────────────────────────────────────────────
    let sobinary = null;
    if (opts.engine === 'libreoffice') {
      sobinary = await checkLibreOffice(opts.libreoffice);
      if (!sobinary) {
        log.error(
          'LibreOffice not found.\n' +
          '  macOS :  brew install --cask libreoffice\n' +
          '  Linux :  sudo apt install libreoffice  |  sudo dnf install libreoffice\n' +
          '  Manual:  https://www.libreoffice.org/download/'
        );
        process.exit(1);
      }
    }

    const absInput = path.resolve(input);

    if (!fs.existsSync(absInput)) {
      log.error(`Input not found: ${absInput}`);
      process.exit(1);
    }

    const isDir = fs.statSync(absInput).isDirectory();
    const timeoutMs = parseInt(opts.timeout, 10) * 1000;
    const shared = {
      engine: opts.engine,
      sobinary,
      fontDir: opts.fontDir ? path.resolve(opts.fontDir) : null,
      chromeBinary: opts.chrome ? path.resolve(opts.chrome) : null,
      nativeLayout: opts.nativeLayout,
      timeoutMs,
      overwrite: opts.overwrite,
      wpsCompat: opts.wpsCompat,
      silent: opts.silent,
    };

    // ── 1. Dispatch ───────────────────────────────────────────────────────────
    if (isDir) {
      const outputDir = opts.output ? path.resolve(opts.output) : absInput;
      await convertDir(absInput, outputDir, shared);
    } else {
      if (!absInput.toLowerCase().endsWith('.docx')) {
        log.error('Input file must be a .docx file.');
        process.exit(1);
      }
      let outputPath;
      if (opts.output) {
        outputPath = path.resolve(opts.output);
        // If user passed a directory, put the PDF inside it
        if (fs.existsSync(outputPath) && fs.statSync(outputPath).isDirectory()) {
          outputPath = path.join(outputPath, path.basename(absInput, '.docx') + '.pdf');
        }
      } else {
        outputPath = absInput.replace(/\.docx$/i, '.pdf');
      }
      const ok = await convertFile(absInput, outputPath, shared);
      process.exit(ok ? 0 : 1);
    }
  });

program.parse();
