'use strict';

const { execFileSync } = require('child_process');
const fs = require('fs');

// Candidate binary paths per platform
const CANDIDATES = {
  darwin: [
    '/Applications/LibreOffice.app/Contents/MacOS/soffice',
    '/Applications/LibreOffice.app/Contents/MacOS/soffice.bin',
  ],
  linux: [
    '/usr/bin/soffice',
    '/usr/bin/libreoffice',
    '/usr/local/bin/soffice',
    '/opt/libreoffice/program/soffice',
  ],
};

/**
 * Resolve the soffice binary path.
 * @param {string|null} userPath  - path from --libreoffice CLI flag
 * @returns {Promise<string|null>}
 */
async function checkLibreOffice(userPath = null) {
  const candidates = [];

  if (userPath) candidates.push(userPath);

  // Platform-specific well-known paths
  const platformPaths = CANDIDATES[process.platform] || [];
  candidates.push(...platformPaths);

  // Also try whatever is on PATH
  candidates.push('soffice', 'libreoffice');

  for (const candidate of candidates) {
    if (await isUsable(candidate)) return candidate;
  }

  return null;
}

async function isUsable(bin) {
  // For absolute paths, check existence first (fast)
  if (bin.startsWith('/') && !fs.existsSync(bin)) return false;

  try {
    execFileSync(bin, ['--version'], { stdio: 'pipe', timeout: 8000 });
    return true;
  } catch {
    return false;
  }
}

module.exports = { checkLibreOffice };
