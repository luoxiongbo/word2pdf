'use strict';

/**
 * fonts.js
 *
 * Generates a custom fontconfig XML file that:
 *  1. Includes all the platform's standard font directories
 *  2. Optionally adds a user-supplied extra font directory
 *  3. Registers explicit font-family aliases so LibreOffice maps
 *     Windows/Office font names (Microsoft YaHei, SimSun, Calibri…)
 *     to whatever is actually installed on the system.
 *
 * This is the single biggest lever for improving CJK fidelity under
 * LibreOffice headless — mismatched font metrics cause most of the
 * text overflow and table misalignment issues.
 */

const fs = require('fs');
const path = require('path');
const os = require('os');

// ─── Platform font directories ─────────────────────────────────────────────

const PLATFORM_FONT_DIRS = (() => {
  switch (process.platform) {
    case 'darwin':
      return [
        '/Library/Fonts',
        '/System/Library/Fonts',
        '/System/Library/Fonts/Supplemental',
        '/System/Library/AssetsV2/com_apple_MobileAsset_Font8',
        path.join(os.homedir(), 'Library/Fonts'),
        // LibreOffice bundles its own fonts here
        '/Applications/LibreOffice.app/Contents/Resources/fonts/truetype',
      ];
    case 'linux':
      return [
        '/usr/share/fonts',
        '/usr/local/share/fonts',
        path.join(os.homedir(), '.fonts'),
        path.join(os.homedir(), '.local/share/fonts'),
      ];
    default:
      return [];
  }
})();

// ─── Font-family alias mappings ────────────────────────────────────────────
//
// Format: { from: 'Windows font name', to: ['preferred', 'fallback', ...] }
//
// LibreOffice will use the first installed font in each 'to' array.

const FONT_ALIASES = [
  // ── CJK — Simplified Chinese ──────────────────────────────────────────────
  {
    from: 'Microsoft YaHei',
    to: ['PingFang SC', 'Heiti SC', 'STHeitiSC-Light', 'Noto Sans CJK SC', 'WenQuanYi Micro Hei', 'Source Han Sans SC'],
  },
  {
    from: 'Microsoft YaHei UI',
    to: ['PingFang SC', 'Heiti SC', 'STHeitiSC-Light', 'Noto Sans CJK SC', 'WenQuanYi Micro Hei'],
  },
  {
    from: 'SimSun',
    to: ['Songti SC', 'STSongti-SC-Regular', 'Noto Serif CJK SC', 'AR PL UMing CN', 'WenQuanYi Bitmap Song'],
  },
  {
    from: 'NSimSun',
    to: ['Songti SC', 'STSongti-SC-Regular', 'Noto Serif CJK SC'],
  },
  {
    from: 'SimHei',
    to: ['Heiti SC', 'STHeitiSC-Medium', 'Noto Sans CJK SC', 'WenQuanYi Micro Hei'],
  },
  {
    from: 'FangSong',
    to: ['STFangsong', 'FangSong', 'Songti SC', 'Noto Serif CJK SC'],
  },
  {
    from: 'KaiTi',
    to: ['Kaiti SC', 'STKaiti', 'STKaitiSC-Regular', 'Noto Serif CJK SC'],
  },
  // ── CJK — Traditional Chinese ─────────────────────────────────────────────
  {
    from: 'Microsoft JhengHei',
    to: ['PingFang TC', 'Heiti TC', 'Noto Sans CJK TC'],
  },
  {
    from: 'MingLiU',
    to: ['Songti TC', 'STSongti-TC-Regular', 'Noto Serif CJK TC'],
  },
  // ── CJK — Japanese ────────────────────────────────────────────────────────
  {
    from: 'Meiryo',
    to: ['Meiryo', 'Hiragino Kaku Gothic Pro', 'Noto Sans CJK JP'],
  },
  {
    from: 'MS Gothic',
    to: ['MS Gothic', 'Hiragino Kaku Gothic Pro', 'Noto Sans CJK JP'],
  },
  // ── Latin — Office defaults ───────────────────────────────────────────────
  {
    from: 'Calibri',
    to: ['Calibri', 'Carlito', 'DejaVu Sans', 'Liberation Sans'],
  },
  {
    from: 'Cambria',
    to: ['Cambria', 'Caladea', 'Georgia', 'Liberation Serif'],
  },
  {
    from: 'Times New Roman',
    to: ['Times New Roman', 'Liberation Serif', 'FreeSerif', 'DejaVu Serif'],
  },
  {
    from: 'Arial',
    to: ['Arial', 'Liberation Sans', 'FreeSans', 'DejaVu Sans'],
  },
  {
    from: 'Helvetica',
    to: ['Helvetica', 'Arial', 'Liberation Sans', 'DejaVu Sans'],
  },
  {
    from: 'Courier New',
    to: ['Courier New', 'Liberation Mono', 'FreeMono', 'DejaVu Sans Mono'],
  },
];

// ─── XML helpers ──────────────────────────────────────────────────────────

function xmlEscape(str) {
  return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function dirElement(dirPath) {
  if (!fs.existsSync(dirPath)) return '';
  return `    <dir>${xmlEscape(dirPath)}</dir>`;
}

function aliasBlock({ from, to }) {
  const accepts = to.map(f => `      <family>${xmlEscape(f)}</family>`).join('\n');
  return `
  <alias>
    <family>${xmlEscape(from)}</family>
    <accept>
${accepts}
    </accept>
  </alias>`;
}

// ─── Public API ────────────────────────────────────────────────────────────

/**
 * Write a fontconfig XML to tmpDir and return the path.
 * @param {string} tmpDir     - temp working directory
 * @param {string|null} extraFontDir - optional user-supplied font directory
 */
function buildFontConfig(tmpDir, extraFontDir = null) {
  const allDirs = [...PLATFORM_FONT_DIRS];
  if (extraFontDir && fs.existsSync(extraFontDir)) {
    allDirs.push(extraFontDir);
  }

  const dirElements = allDirs.map(dirElement).filter(Boolean).join('\n');
  const aliasBlocks = FONT_ALIASES.map(aliasBlock).join('');

  const xml = `<?xml version="1.0"?>
<!DOCTYPE fontconfig SYSTEM "fonts.dtd">
<!--
  Auto-generated by docx2pdf-cli.
  Maps common Windows/Office font names to installed system fonts,
  improving CJK rendering fidelity in LibreOffice headless.
-->
<fontconfig>
  <!-- Font search directories -->
${dirElements}

  <!-- Scan subdirectories -->
  <selectfont>
    <rejectfont>
      <pattern><patelt name="scalable"><bool>false</bool></patelt></pattern>
    </rejectfont>
  </selectfont>

  <!-- Font-family alias mappings -->
${aliasBlocks}

  <!-- Prefer TrueType/OpenType over bitmap -->
  <match target="pattern">
    <test name="outline"><bool>true</bool></test>
    <edit name="outline" mode="assign"><bool>true</bool></edit>
  </match>

  <!-- Enable sub-pixel rendering for CJK sharpness -->
  <match target="font">
    <edit name="rgba" mode="assign"><const>rgb</const></edit>
    <edit name="hinting" mode="assign"><bool>true</bool></edit>
    <edit name="hintstyle" mode="assign"><const>hintslight</const></edit>
    <edit name="antialias" mode="assign"><bool>true</bool></edit>
  </match>
</fontconfig>
`;

  const confPath = path.join(tmpDir, 'fonts.conf');
  fs.writeFileSync(confPath, xml, 'utf8');
  return confPath;
}

module.exports = { buildFontConfig, FONT_ALIASES, PLATFORM_FONT_DIRS };
