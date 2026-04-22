'use strict';

const { execFileSync } = require('child_process');
const path = require('path');
const fs = require('fs');

function _normalizeXmlContent(xml) {
  let out = xml;

  // Fixed line-height often causes overlap in WPS-origin text boxes.
  out = out.replace(/lineRule="exact"/g, 'lineRule="atLeast"');
  return out;
}

function collectXmlFiles(rootDir) {
  const result = [];

  function walk(dir) {
    const entries = fs.readdirSync(dir, { withFileTypes: true });
    for (const entry of entries) {
      const fullPath = path.join(dir, entry.name);
      if (entry.isDirectory()) {
        walk(fullPath);
      } else if (entry.isFile() && fullPath.toLowerCase().endsWith('.xml')) {
        result.push(fullPath);
      }
    }
  }

  walk(rootDir);
  return result;
}

/**
 * 探测并清理 WPS 风格文档。
 */
function prepareDocx(inputPath, tmpDir) {
  const normalizedDocx = path.join(tmpDir, 'prepared.docx');
  const contentsDir = path.join(tmpDir, 'contents');
  fs.copyFileSync(inputPath, normalizedDocx);

  try {
    fs.mkdirSync(contentsDir, { recursive: true });
    execFileSync('unzip', ['-q', normalizedDocx, '-d', contentsDir], { stdio: 'ignore' });

    const xmlFiles = collectXmlFiles(contentsDir);
    for (const xmlFile of xmlFiles) {
      const original = fs.readFileSync(xmlFile, 'utf8');
      const normalized = _normalizeXmlContent(original);
      if (normalized !== original) {
        fs.writeFileSync(xmlFile, normalized, 'utf8');
      }
    }

    fs.rmSync(normalizedDocx, { force: true });
    execFileSync('zip', ['-qr', normalizedDocx, '.'], { cwd: contentsDir, stdio: 'ignore' });
    execFileSync('unzip', ['-t', normalizedDocx], { stdio: 'ignore' });
    fs.rmSync(contentsDir, { recursive: true, force: true });
    return normalizedDocx;
  } catch (err) {
    console.error('Normalization error:', err);
    fs.rmSync(contentsDir, { recursive: true, force: true });
    return inputPath;
  }
}

module.exports = { prepareDocx, _normalizeXmlContent };
