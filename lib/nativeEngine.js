'use strict';

const fs = require('fs');
const path = require('path');
const os = require('os');
const { execFileSync, spawn, spawnSync } = require('child_process');
const { log } = require('./logger');

function _decodeXmlEntities(text) {
  if (!text) return '';
  return text
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&nbsp;/g, ' ')
    .replace(/&#xA;/gi, '\n')
    .replace(/&#10;/g, '\n');
}

function _extractTextFromParagraphXml(xml) {
  if (!xml) return '';
  const tokenRe = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>|<w:tab(?:\s[^>]*)?\/>|<w:br(?:\s[^>]*)?\/>|<w:cr(?:\s[^>]*)?\/>/g;
  const out = [];
  let match;

  while ((match = tokenRe.exec(xml)) !== null) {
    if (typeof match[1] === 'string') {
      out.push(_decodeXmlEntities(match[1]));
      continue;
    }
    const token = match[0];
    if (token.startsWith('<w:tab')) out.push('\t');
    if (token.startsWith('<w:br') || token.startsWith('<w:cr')) out.push('\n');
  }

  return out.join('')
    .replace(/\r/g, '')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function _extractTextFromDocumentXml(xml) {
  if (!xml) return '';
  const tokenRe = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>|<w:tab(?:\s[^>]*)?\/>|<w:br(?:\s[^>]*)?\/>|<w:cr(?:\s[^>]*)?\/>|<\/w:p>|<\/w:tr>|<\/w:tc>/g;
  const out = [];
  let match;

  while ((match = tokenRe.exec(xml)) !== null) {
    if (typeof match[1] === 'string') {
      out.push(_decodeXmlEntities(match[1]));
      continue;
    }

    const token = match[0];
    if (token.startsWith('<w:tab')) {
      out.push('\t');
    } else if (token.startsWith('<w:br') || token.startsWith('<w:cr') || token === '</w:tr>') {
      out.push('\n');
    } else if (token === '</w:p>') {
      const ahead = xml.slice(tokenRe.lastIndex, tokenRe.lastIndex + 16);
      if (!ahead.startsWith('</w:tc>')) out.push('\n');
    } else if (token === '</w:tc>') {
      out.push('\t');
    }
  }

  return out.join('')
    .replace(/\r/g, '')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\t+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function _charVisualWidth(ch, fontSize) {
  if (ch === '\t') return fontSize * 2;
  const code = ch.codePointAt(0);
  if (code <= 0x7f) return fontSize * 0.55;
  return fontSize;
}

function _wrapLineByVisualWidth(line, maxWidth, fontSize) {
  if (!line) return [''];
  if (!maxWidth || maxWidth <= 0) return [line];

  const wrapped = [];
  let current = '';
  let currentWidth = 0;

  for (const ch of line) {
    const cw = _charVisualWidth(ch, fontSize);
    if (current && currentWidth + cw > maxWidth) {
      wrapped.push(current);
      current = ch;
      currentWidth = cw;
    } else {
      current += ch;
      currentWidth += cw;
    }
  }

  if (current || wrapped.length === 0) wrapped.push(current);
  return wrapped;
}

function _toUtf16BeHex(text) {
  const utf16le = Buffer.from(text, 'utf16le');
  for (let i = 0; i < utf16le.length; i += 2) {
    const a = utf16le[i];
    utf16le[i] = utf16le[i + 1];
    utf16le[i + 1] = a;
  }
  return utf16le.toString('hex').toUpperCase();
}

function _num(n) {
  return Number.isInteger(n) ? String(n) : n.toFixed(2);
}

function _escapeHtml(text) {
  if (!text) return '';
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function _buildHtmlDocument(lines, options = {}) {
  const title = _escapeHtml(options.title || 'docx2pdf-native');
  const body = lines.map(line => _escapeHtml(line)).join('\n');
  return [
    '<!doctype html>',
    '<html>',
    '<head>',
    '<meta charset="utf-8">',
    `<title>${title}</title>`,
    '<style>',
    '@page { size: A4; margin: 14mm 12mm; }',
    'html, body { margin: 0; padding: 0; }',
    'body {',
    '  font-family: "PingFang SC", "Hiragino Sans GB", "Microsoft YaHei", "STHeiti", "Songti SC", sans-serif;',
    '  font-size: 14px;',
    '  line-height: 1.55;',
    '  white-space: pre-wrap;',
    '  word-break: break-word;',
    '}',
    '</style>',
    '</head>',
    `<body>${body}</body>`,
    '</html>',
    '',
  ].join('\n');
}

function _extractRelationshipMap(relsXml) {
  const map = {};
  if (!relsXml) return map;
  const re = /<Relationship\b[^>]*\bId="([^"]+)"[^>]*\bTarget="([^"]+)"[^>]*\/>/g;
  let match;
  while ((match = re.exec(relsXml)) !== null) {
    map[match[1]] = match[2];
  }
  return map;
}

function _mimeFromEntryName(entryName) {
  const ext = path.extname(entryName || '').toLowerCase();
  if (ext === '.png') return 'image/png';
  if (ext === '.jpg' || ext === '.jpeg') return 'image/jpeg';
  if (ext === '.gif') return 'image/gif';
  if (ext === '.bmp') return 'image/bmp';
  if (ext === '.svg') return 'image/svg+xml';
  return 'application/octet-stream';
}

function _safeReadZipText(docxPath, entryName) {
  try {
    return execFileSync('unzip', ['-p', docxPath, entryName], {
      encoding: 'utf8',
      maxBuffer: 64 * 1024 * 1024,
    });
  } catch (_err) {
    return '';
  }
}

function _safeReadZipBinary(docxPath, entryName) {
  try {
    return execFileSync('unzip', ['-p', docxPath, entryName], {
      encoding: null,
      maxBuffer: 128 * 1024 * 1024,
    });
  } catch (_err) {
    return null;
  }
}

function _resolveWordZipEntry(target) {
  if (!target) return null;
  const cleaned = target.replace(/^\/+/, '');
  let entry = path.posix.normalize(path.posix.join('word', cleaned));
  if (entry.startsWith('../')) return null;
  if (!entry.startsWith('word/')) entry = `word/${cleaned}`;
  return entry;
}

function _buildMediaResolver(docxPath) {
  const relsXml = _safeReadZipText(docxPath, 'word/_rels/document.xml.rels');
  const relMap = _extractRelationshipMap(relsXml);
  const cache = {};

  return (rid) => {
    if (!rid) return null;
    if (cache[rid]) return cache[rid];
    const target = relMap[rid];
    const entry = _resolveWordZipEntry(target);
    if (!entry) return null;
    const bin = _safeReadZipBinary(docxPath, entry);
    if (!bin) return null;
    const mime = _mimeFromEntryName(entry);
    const dataUri = `data:${mime};base64,${bin.toString('base64')}`;
    const data = { rid, entry, mime, dataUri };
    cache[rid] = data;
    return data;
  };
}

function _stripFallback(xml) {
  if (!xml) return '';
  return xml
    .replace(/<mc:Fallback>[\s\S]*?<\/mc:Fallback>/g, '')
    .replace(/<w:pict>[\s\S]*?<\/w:pict>/g, '');
}

function _extractBodyXml(xml) {
  const match = xml.match(/<w:body>([\s\S]*?)<w:sectPr[\s\S]*?<\/w:sectPr><\/w:body>/);
  if (!match) return xml;
  return match[1];
}

function _extractTopLevelBodyBlocks(bodyXml) {
  const blocks = [];
  if (!bodyXml) return blocks;

  let cursor = 0;
  const startRe = /<w:(p|tbl)\b/g;

  while (cursor < bodyXml.length) {
    startRe.lastIndex = cursor;
    const startMatch = startRe.exec(bodyXml);
    if (!startMatch) break;

    const type = startMatch[1];
    const start = startMatch.index;
    let end = -1;

    if (type === 'p') {
      const close = bodyXml.indexOf('</w:p>', start);
      if (close !== -1) end = close + '</w:p>'.length;
    } else {
      const tblTokenRe = /<\/?w:tbl\b[^>]*>/g;
      tblTokenRe.lastIndex = start;
      let depth = 0;
      let token;
      while ((token = tblTokenRe.exec(bodyXml)) !== null) {
        const text = token[0];
        const isClose = text.startsWith('</');
        const isSelfClose = text.endsWith('/>');
        if (!isClose && !isSelfClose) depth += 1;
        if (isClose) depth -= 1;
        if (depth === 0) {
          end = token.index + text.length;
          break;
        }
      }
    }

    if (end <= start) break;
    blocks.push({ type, xml: bodyXml.slice(start, end) });
    cursor = end;
  }

  return blocks;
}

function _extractTableRows(tableXml) {
  const rows = [];
  if (!tableXml) return rows;
  const trRe = /<w:tr\b[\s\S]*?<\/w:tr>/g;
  let trMatch;

  while ((trMatch = trRe.exec(tableXml)) !== null) {
    const trXml = trMatch[0];
    const row = [];
    const tcRe = /<w:tc\b[\s\S]*?<\/w:tc>/g;
    let tcMatch;

    while ((tcMatch = tcRe.exec(trXml)) !== null) {
      const tcXml = tcMatch[0];
      const pRe = /<w:p\b[\s\S]*?<\/w:p>/g;
      let pMatch;
      const cellLines = [];

      while ((pMatch = pRe.exec(tcXml)) !== null) {
        const line = _extractTextFromParagraphXml(pMatch[0]);
        if (line) cellLines.push(line);
      }

      row.push(cellLines.join('\n').trim());
    }

    if (row.length > 0) rows.push(row);
  }

  return rows;
}

function _extractEmbeddedImageRids(paragraphXml) {
  const rids = [];
  if (!paragraphXml) return rids;
  const re = /<a:blip\b[^>]*\br:embed="([^"]+)"/g;
  let match;
  while ((match = re.exec(paragraphXml)) !== null) {
    rids.push(match[1]);
  }
  return [...new Set(rids)];
}

function _extractParagraphSegments(snippet) {
  const out = [];
  if (!snippet) return out;
  const pRe = /<w:p\b[\s\S]*?<\/w:p>/g;
  let match;
  while ((match = pRe.exec(snippet)) !== null) {
    const paragraphXml = match[0];
    const text = _extractTextFromParagraphXml(paragraphXml);
    if (!text) continue;
    const sizeMatch = paragraphXml.match(/<w:sz\s+[^>]*w:val="(\d+)"/);
    const size = sizeMatch ? parseInt(sizeMatch[1], 10) : null;
    const bold = /<w:b(?:\s*\/>|>)/.test(paragraphXml) || /<w:bCs(?:\s*\/>|>)/.test(paragraphXml);
    out.push({ type: 'paragraph', text, size, bold });
  }
  return out;
}

function _extractRenderSegmentsFromDocumentXml(xml) {
  const cleaned = _stripFallback(xml);
  const body = _extractBodyXml(cleaned);
  const out = [];
  const re = /<w:txbxContent>([\s\S]*?)<\/w:txbxContent>/g;

  let cursor = 0;
  let match;
  while ((match = re.exec(body)) !== null) {
    const before = body.slice(cursor, match.index);
    out.push(..._extractParagraphSegments(before));

    const textboxXml = match[1];
    const lines = _extractParagraphSegments(textboxXml).map(s => s.text).filter(Boolean);
    if (lines.length > 0) {
      out.push({
        type: 'textbox',
        lines,
        text: lines.join('\n'),
      });
    }

    cursor = match.index + match[0].length;
  }

  out.push(..._extractParagraphSegments(body.slice(cursor)));

  return out
    .map(seg => {
      if (seg.type === 'textbox') {
        return {
          ...seg,
          lines: seg.lines.map(line => line.trim()).filter(Boolean),
          text: seg.lines.join('\n').trim(),
        };
      }
      return {
        ...seg,
        text: seg.text.trim(),
      };
    })
    .filter(seg => (seg.text || '').trim().length > 0);
}

function _looksLikeHeading(text) {
  const t = (text || '').trim();
  if (!t) return false;
  if (t.length > 8) return false;
  if (/^[-•]/.test(t)) return false;
  if (/\d/.test(t)) return false;
  return true;
}

function _looksLikeHeadingByStyle(seg) {
  if (!seg || seg.type !== 'paragraph') return false;
  if (typeof seg.size === 'number' && seg.size >= 28) return true;
  if (seg.bold && (seg.text || '').trim().length <= 10) return true;
  return false;
}

function _isDateRangeText(text) {
  if (!text) return false;
  const t = text.replace(/\s+/g, '');
  const re = /^\d{4}[./-]\d{1,2}(?:[./-]\d{1,2})?[-~—]\d{4}[./-]\d{1,2}(?:[./-]\d{1,2})?$/;
  return re.test(t);
}

function _segmentsToStructuredBlocks(segments) {
  const blocks = [];
  for (let i = 0; i < segments.length; i++) {
    const seg = segments[i];

    if (seg.type === 'paragraph') {
      const text = seg.text.trim();
      if (!text) continue;

      if (_looksLikeHeadingByStyle(seg) || _looksLikeHeading(text)) {
        blocks.push({ type: 'heading', text });
        continue;
      }

      if (/^[-•]/.test(text)) {
        const bulletLines = text
          .split('\n')
          .map(line => line.trim())
          .filter(Boolean);
        for (const line of bulletLines) {
          blocks.push({ type: 'bullet', text: line.replace(/^[-•]\s*/, '') });
        }
      } else {
        blocks.push({ type: 'paragraph', text });
      }
      continue;
    }

    if (seg.type === 'textbox') {
      const next = segments[i + 1];
      if (next && next.type === 'textbox') {
        const left = seg.lines;
        const right = next.lines;
        const canPair = left.length <= 5 && right.length <= 5;
        if (canPair) {
          blocks.push({ type: 'row', left, right });
          i += 1;
          continue;
        }
      }
      const maybeDateOnly = seg.lines.length === 1 && _isDateRangeText(seg.lines[0]);
      if (maybeDateOnly && blocks.length > 0) {
        const prev = blocks[blocks.length - 1];
        if (prev.type === 'row') {
          prev.right = [...prev.right, seg.lines[0]];
          continue;
        }
      }

      blocks.push({ type: 'textbox', lines: seg.lines });
    }
  }

  return blocks;
}

function _extractRichBlocksFromDocx(docxPath, xml) {
  const cleaned = _stripFallback(xml);
  const bodyXml = _extractBodyXml(cleaned);
  const topBlocks = _extractTopLevelBodyBlocks(bodyXml);
  const resolveMedia = _buildMediaResolver(docxPath);
  const blocks = [];

  for (const block of topBlocks) {
    if (block.type === 'tbl') {
      const rows = _extractTableRows(block.xml);
      if (rows.length > 0) blocks.push({ type: 'table', rows });
      continue;
    }

    const paragraphXml = block.xml;
    const text = _extractTextFromParagraphXml(paragraphXml).trim();

    const imageRids = _extractEmbeddedImageRids(paragraphXml);
    for (const rid of imageRids) {
      const media = resolveMedia(rid);
      if (!media) continue;
      blocks.push({ type: 'image', dataUri: media.dataUri, mime: media.mime, rid });
    }

    if (!text) continue;

    if (_looksLikeHeading(text)) {
      blocks.push({ type: 'heading', text });
      continue;
    }

    if (/^[-•]/.test(text)) {
      const bulletLines = text
        .split('\n')
        .map(line => line.trim())
        .filter(Boolean);
      for (const line of bulletLines) {
        blocks.push({ type: 'bullet', text: line.replace(/^[-•]\s*/, '') });
      }
      continue;
    }

    blocks.push({ type: 'paragraph', text });
  }

  return blocks;
}

function _buildStructuredHtmlDocument(blocks, options = {}) {
  const title = _escapeHtml(options.title || 'docx2pdf-native');
  const html = [];

  html.push('<!doctype html>');
  html.push('<html>');
  html.push('<head>');
  html.push('<meta charset="utf-8">');
  html.push(`<title>${title}</title>`);
  html.push('<style>');
  html.push('@page { size: A4; margin: 10mm 12mm; }');
  html.push('html, body { margin: 0; padding: 0; }');
  html.push('body { font-family: "PingFang SC", "Hiragino Sans GB", "Microsoft YaHei", "STHeiti", "Songti SC", sans-serif; color: #111; font-size: 14px; line-height: 1.55; }');
  html.push('.page { width: 100%; }');
  html.push('.heading { font-weight: 700; font-size: 20px; margin: 10px 0 8px; padding-bottom: 4px; border-bottom: 2px solid #222; }');
  html.push('.row { display: grid; grid-template-columns: 1fr 1fr; column-gap: 28px; margin: 6px 0; }');
  html.push('.cell { white-space: pre-wrap; word-break: break-word; min-height: 24px; }');
  html.push('.paragraph { margin: 5px 0; white-space: pre-wrap; word-break: break-word; }');
  html.push('.bullet { margin: 4px 0; white-space: pre-wrap; word-break: break-word; text-indent: -1em; padding-left: 1em; }');
  html.push('.doc-table { width: 100%; border-collapse: collapse; margin: 8px 0 14px; font-size: 13px; }');
  html.push('.doc-table td, .doc-table th { border: 1px solid #4b5563; padding: 6px 8px; vertical-align: top; }');
  html.push('.image-wrap { margin: 10px 0 12px; text-align: center; }');
  html.push('.image-wrap img { max-width: 100%; height: auto; border: 1px solid #d1d5db; }');
  html.push('</style>');
  html.push('</head>');
  html.push('<body><main class="page">');

  for (const block of blocks) {
    if (block.type === 'heading') {
      html.push(`<section class="heading">${_escapeHtml(block.text)}</section>`);
      continue;
    }
    if (block.type === 'row') {
      const left = _escapeHtml(block.left.join('\n'));
      const right = _escapeHtml(block.right.join('\n'));
      html.push('<section class="row">');
      html.push(`<div class="cell">${left}</div>`);
      html.push(`<div class="cell">${right}</div>`);
      html.push('</section>');
      continue;
    }
    if (block.type === 'bullet') {
      html.push(`<p class="bullet">- ${_escapeHtml(block.text)}</p>`);
      continue;
    }
    if (block.type === 'textbox') {
      html.push(`<p class="paragraph">${_escapeHtml(block.lines.join('\n'))}</p>`);
      continue;
    }
    if (block.type === 'table') {
      html.push('<table class="doc-table">');
      for (const row of block.rows) {
        html.push('<tr>');
        for (const cell of row) {
          html.push(`<td>${_escapeHtml(cell || '').replace(/\n/g, '<br>')}</td>`);
        }
        html.push('</tr>');
      }
      html.push('</table>');
      continue;
    }
    if (block.type === 'image') {
      html.push(`<div class="image-wrap"><img src="${block.dataUri}" alt="embedded-image"></div>`);
      continue;
    }
    html.push(`<p class="paragraph">${_escapeHtml(block.text || '')}</p>`);
  }

  html.push('</main></body>');
  html.push('</html>');
  html.push('');
  return html.join('\n');
}

function _findChromeBinary(customPath) {
  const candidates = [];
  if (customPath) candidates.push(customPath);
  if (process.env.CHROME_PATH) candidates.push(process.env.CHROME_PATH);
  candidates.push(
    '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
    '/Applications/Chromium.app/Contents/MacOS/Chromium',
    'google-chrome',
    'chromium',
    'chromium-browser'
  );

  for (const candidate of candidates) {
    if (!candidate) continue;
    if (candidate.startsWith('/')) {
      if (fs.existsSync(candidate)) return candidate;
      continue;
    }
    const res = spawnSync('which', [candidate], { encoding: 'utf8' });
    if (res.status === 0) return candidate;
  }
  return null;
}

function _runChromePrintToPdf({
  chromeBinary,
  htmlPath,
  outputPath,
  timeoutMs = 45_000,
}) {
  return new Promise((resolve, reject) => {
    const args = [
      '--headless=new',
      '--disable-gpu',
      '--no-first-run',
      '--no-default-browser-check',
      `--print-to-pdf=${outputPath}`,
      `file://${htmlPath}`,
    ];
    const proc = spawn(chromeBinary, args, { stdio: ['ignore', 'pipe', 'pipe'] });
    let stderr = '';
    let stdout = '';

    proc.stdout.on('data', chunk => { stdout += chunk.toString(); });
    proc.stderr.on('data', chunk => { stderr += chunk.toString(); });

    const timer = setTimeout(() => {
      proc.kill('SIGKILL');
      reject(new Error(`chrome print timed out after ${Math.floor(timeoutMs / 1000)}s`));
    }, timeoutMs);

    proc.on('close', code => {
      clearTimeout(timer);
      if (code === 0 && fs.existsSync(outputPath)) return resolve();
      const detail = (stderr || stdout || '').replace(/\s+/g, ' ').trim();
      reject(new Error(`chrome print failed with code ${code}; ${detail}`));
    });

    proc.on('error', err => {
      clearTimeout(timer);
      reject(new Error(`failed to launch chrome: ${err.message}`));
    });
  });
}

function _buildPageStream(lines, opts) {
  const {
    fontSize,
    lineHeight,
    marginLeft,
    marginTop,
    pageHeight,
  } = opts;

  const chunks = [
    'BT',
    `/F1 ${_num(fontSize)} Tf`,
  ];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (!line) continue;
    const y = pageHeight - marginTop - fontSize - i * lineHeight;
    chunks.push(`1 0 0 1 ${_num(marginLeft)} ${_num(y)} Tm`);
    chunks.push(`<${_toUtf16BeHex(line)}> Tj`);
  }

  chunks.push('ET');
  return `${chunks.join('\n')}\n`;
}

function _buildSimplePdf(pages, options = {}) {
  const pageWidth = options.pageWidth || 595.28;
  const pageHeight = options.pageHeight || 841.89;
  const marginLeft = options.marginLeft || 50;
  const marginTop = options.marginTop || 50;
  const fontSize = options.fontSize || 11;
  const lineHeight = options.lineHeight || 16;
  const normalizedPages = pages && pages.length > 0 ? pages : [['']];

  const pageCount = normalizedPages.length;
  const descId = 1;
  const fontId = 2;
  const contentStartId = 3;
  const pageStartId = contentStartId + pageCount;
  const pagesId = pageStartId + pageCount;
  const catalogId = pagesId + 1;
  const maxId = catalogId;
  const objects = new Array(maxId + 1);

  objects[descId] = '<< /Type /Font /Subtype /CIDFontType0 /BaseFont /STSong-Light /CIDSystemInfo << /Registry (Adobe) /Ordering (GB1) /Supplement 4 >> /DW 1000 >>';
  objects[fontId] = `<< /Type /Font /Subtype /Type0 /BaseFont /STSong-Light /Encoding /UniGB-UCS2-H /DescendantFonts [${descId} 0 R] >>`;

  for (let i = 0; i < pageCount; i++) {
    const contentId = contentStartId + i;
    const pageId = pageStartId + i;
    const stream = _buildPageStream(normalizedPages[i], {
      pageHeight,
      marginLeft,
      marginTop,
      fontSize,
      lineHeight,
    });
    const streamLen = Buffer.byteLength(stream, 'utf8');
    objects[contentId] = `<< /Length ${streamLen} >>\nstream\n${stream}endstream`;
    objects[pageId] = `<< /Type /Page /Parent ${pagesId} 0 R /MediaBox [0 0 ${_num(pageWidth)} ${_num(pageHeight)}] /Resources << /Font << /F1 ${fontId} 0 R >> >> /Contents ${contentId} 0 R >>`;
  }

  const kids = Array.from({ length: pageCount }, (_, idx) => `${pageStartId + idx} 0 R`).join(' ');
  objects[pagesId] = `<< /Type /Pages /Count ${pageCount} /Kids [${kids}] >>`;
  objects[catalogId] = `<< /Type /Catalog /Pages ${pagesId} 0 R >>`;

  let body = '%PDF-1.4\n%\xC2\xC3\xC4\xC5\n';
  const offsets = new Array(maxId + 1).fill(0);

  for (let id = 1; id <= maxId; id++) {
    offsets[id] = Buffer.byteLength(body, 'utf8');
    body += `${id} 0 obj\n${objects[id]}\nendobj\n`;
  }

  const xrefOffset = Buffer.byteLength(body, 'utf8');
  body += `xref\n0 ${maxId + 1}\n`;
  body += '0000000000 65535 f \n';
  for (let id = 1; id <= maxId; id++) {
    body += `${String(offsets[id]).padStart(10, '0')} 00000 n \n`;
  }
  body += `trailer\n<< /Size ${maxId + 1} /Root ${catalogId} 0 R >>\nstartxref\n${xrefOffset}\n%%EOF\n`;

  return Buffer.from(body, 'utf8');
}

function _extractDocumentXml(docxPath) {
  try {
    return execFileSync('unzip', ['-p', docxPath, 'word/document.xml'], {
      encoding: 'utf8',
      maxBuffer: 64 * 1024 * 1024,
    });
  } catch (err) {
    const detail = err.stderr ? String(err.stderr).trim() : err.message;
    throw new Error(`failed to read DOCX XML: ${detail}`);
  }
}

function _paginateLines(lines, opts = {}) {
  const pageHeight = opts.pageHeight || 841.89;
  const marginTop = opts.marginTop || 50;
  const marginBottom = opts.marginBottom || 50;
  const lineHeight = opts.lineHeight || 16;

  const usableHeight = pageHeight - marginTop - marginBottom;
  const linesPerPage = Math.max(1, Math.floor(usableHeight / lineHeight));
  const pages = [];

  for (let i = 0; i < lines.length; i += linesPerPage) {
    pages.push(lines.slice(i, i + linesPerPage));
  }

  return pages.length > 0 ? pages : [['']];
}

async function convertFileNative(inputPath, outputPath, opts = {}) {
  const {
    overwrite = false,
    silent = false,
    timeoutMs = 45_000,
    chromeBinary = null,
    nativeLayout = 'structured',
  } = opts;

  if (!overwrite && fs.existsSync(outputPath)) {
    if (!silent) log.skip(outputPath);
    return true;
  }

  try {
    if (!silent) log.converting(path.basename(inputPath));
    const xml = _extractDocumentXml(inputPath);
    const chrome = _findChromeBinary(chromeBinary);

    fs.mkdirSync(path.dirname(outputPath), { recursive: true });

    if (chrome) {
      const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'docx2pdf-native-'));
      try {
        const htmlPath = path.join(tmpDir, 'render.html');
        let html;

        if (nativeLayout === 'simple') {
          const text = _extractTextFromDocumentXml(xml);
          const normalizedLines = text.split('\n').map(line => line.replace(/\t/g, '    '));
          html = _buildHtmlDocument(normalizedLines, { title: path.basename(inputPath) });
        } else if (nativeLayout === 'rich') {
          const richBlocks = _extractRichBlocksFromDocx(inputPath, xml);
          html = _buildStructuredHtmlDocument(richBlocks, { title: path.basename(inputPath) });
        } else {
          const hasTextBoxes = /<w:txbxContent>/.test(xml);
          if (hasTextBoxes) {
            const segments = _extractRenderSegmentsFromDocumentXml(xml);
            const blocks = _segmentsToStructuredBlocks(segments);
            html = _buildStructuredHtmlDocument(blocks, { title: path.basename(inputPath) });
          } else {
            const richBlocks = _extractRichBlocksFromDocx(inputPath, xml);
            html = _buildStructuredHtmlDocument(richBlocks, { title: path.basename(inputPath) });
          }
        }

        fs.writeFileSync(htmlPath, html, 'utf8');
        await _runChromePrintToPdf({
          chromeBinary: chrome,
          htmlPath,
          outputPath,
          timeoutMs,
        });
      } finally {
        fs.rmSync(tmpDir, { recursive: true, force: true });
      }
    } else {
      const text = _extractTextFromDocumentXml(xml);
      const normalizedLines = text.split('\n').map(line => line.replace(/\t/g, '    '));
      const baseLines = normalizedLines.flatMap(line => _wrapLineByVisualWidth(line, 495, 11));
      const pages = _paginateLines(baseLines, { lineHeight: 16 });
      const pdf = _buildSimplePdf(pages, {
        fontSize: 11,
        lineHeight: 16,
        marginLeft: 50,
        marginTop: 50,
        marginBottom: 50,
      });
      fs.writeFileSync(outputPath, pdf);
    }

    if (!silent) log.success(outputPath);
    return true;
  } catch (err) {
    log.error(`${path.basename(inputPath)}: ${err.message}`);
    return false;
  }
}

module.exports = {
  convertFileNative,
  _decodeXmlEntities,
  _extractTextFromParagraphXml,
  _extractTextFromDocumentXml,
  _wrapLineByVisualWidth,
  _buildSimplePdf,
  _buildHtmlDocument,
  _buildStructuredHtmlDocument,
  _extractRenderSegmentsFromDocumentXml,
  _segmentsToStructuredBlocks,
  _extractRelationshipMap,
  _extractTopLevelBodyBlocks,
  _extractTableRows,
  _extractRichBlocksFromDocx,
  _findChromeBinary,
};
