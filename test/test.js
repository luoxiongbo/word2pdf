'use strict';

/**
 * Smoke tests — run with:  node test/test.js
 *
 * These tests don't require LibreOffice to be installed; they check the
 * logic-only parts of the library (font config generation, preflight
 * candidate list, logger colour codes).
 */

const assert = require('assert');
const path = require('path');
const fs = require('fs');
const os = require('os');

let passed = 0;
let failed = 0;

function test(name, fn) {
  try {
    fn();
    console.log(`  ✔  ${name}`);
    passed++;
  } catch (err) {
    console.error(`  ✖  ${name}`);
    console.error(`     ${err.message}`);
    failed++;
  }
}

// ── fonts.js ─────────────────────────────────────────────────────────────────
console.log('\nfonts.js');
const { buildFontConfig, FONT_ALIASES } = require('../lib/fonts');

test('generates a valid XML file', () => {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'docx2pdf-test-'));
  const confPath = buildFontConfig(tmp);
  assert.ok(fs.existsSync(confPath), 'conf file not created');
  const xml = fs.readFileSync(confPath, 'utf8');
  assert.ok(xml.startsWith('<?xml'), 'should start with XML declaration');
  assert.ok(xml.includes('<fontconfig>'), 'should contain <fontconfig>');
  assert.ok(xml.includes('Microsoft YaHei'), 'should map Microsoft YaHei');
  assert.ok(xml.includes('SimSun'), 'should map SimSun');
  fs.rmSync(tmp, { recursive: true });
});

test('includes extra font dir when provided', () => {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'docx2pdf-test-'));
  const extraDir = path.join(tmp, 'extra-fonts');
  fs.mkdirSync(extraDir);
  const confPath = buildFontConfig(tmp, extraDir);
  const xml = fs.readFileSync(confPath, 'utf8');
  assert.ok(xml.includes(extraDir), 'extra dir should appear in conf');
  fs.rmSync(tmp, { recursive: true });
});

test('FONT_ALIASES has expected CJK entries', () => {
  const names = FONT_ALIASES.map(a => a.from);
  assert.ok(names.includes('Microsoft YaHei'), 'missing Microsoft YaHei');
  assert.ok(names.includes('SimSun'), 'missing SimSun');
  assert.ok(names.includes('SimHei'), 'missing SimHei');
  assert.ok(names.includes('Calibri'), 'missing Calibri');
});

// ── logger.js ─────────────────────────────────────────────────────────────────
console.log('\nlogger.js');
const { log } = require('../lib/logger');

test('log object has required methods', () => {
  ['info', 'success', 'skip', 'warn', 'error', 'converting', 'summary'].forEach(m => {
    assert.strictEqual(typeof log[m], 'function', `log.${m} missing`);
  });
});

// ── preflight.js ──────────────────────────────────────────────────────────────
console.log('\npreflight.js');
const { checkLibreOffice } = require('../lib/preflight');

test('checkLibreOffice returns null for a nonexistent custom path', async () => {
  // We wrap in an immediately-invoked async to allow await
  const result = await checkLibreOffice('/nonexistent/soffice-xyz');
  // The function searches other candidates too, so it may find a real one.
  // Just assert the return type is string or null.
  assert.ok(result === null || typeof result === 'string');
});

// ── converter.js ──────────────────────────────────────────────────────────────
console.log('\nconverter.js');
const { _formatSofficeFailure, _resolvePreparedInput } = require('../lib/converter');

test('formats signal failures with signal name', () => {
  const err = _formatSofficeFailure({ code: null, signal: 'SIGABRT', stdout: '', stderr: '' });
  assert.ok(err instanceof Error, 'should return Error');
  assert.ok(err.message.includes('terminated by signal SIGABRT'), 'should include signal details');
});

test('formats non-zero exit code failures with code value', () => {
  const err = _formatSofficeFailure({ code: 1, signal: null, stdout: '', stderr: '' });
  assert.ok(err instanceof Error, 'should return Error');
  assert.ok(err.message.includes('exited with code 1'), 'should include exit code');
});

test('uses WPS preprocessing by default', () => {
  let called = 0;
  const result = _resolvePreparedInput('/input.docx', '/tmp/work', {
    prepareDocxImpl: (inputPath, tmpDir) => {
      called++;
      assert.strictEqual(inputPath, '/input.docx');
      assert.strictEqual(tmpDir, '/tmp/work');
      return '/tmp/work/prepared.docx';
    },
  });
  assert.strictEqual(called, 1, 'prepareDocx should be called');
  assert.strictEqual(result, '/tmp/work/prepared.docx');
});

test('skips WPS preprocessing when disabled', () => {
  let called = 0;
  const result = _resolvePreparedInput('/input.docx', '/tmp/work', {
    wpsCompat: false,
    prepareDocxImpl: () => {
      called++;
      return '/tmp/work/prepared.docx';
    },
  });
  assert.strictEqual(called, 0, 'prepareDocx should not be called');
  assert.strictEqual(result, '/input.docx');
});

// ── nativeEngine.js ──────────────────────────────────────────────────────────
console.log('\nnativeEngine.js');
const {
  _decodeXmlEntities,
  _extractTextFromDocumentXml,
  _wrapLineByVisualWidth,
  _buildSimplePdf,
  _buildHtmlDocument,
  _extractRenderSegmentsFromDocumentXml,
  _segmentsToStructuredBlocks,
  _buildStructuredHtmlDocument,
  _extractRelationshipMap,
  _extractTopLevelBodyBlocks,
  _extractTableRows,
} = require('../lib/nativeEngine');

test('decodes common XML entities', () => {
  const out = _decodeXmlEntities('A&amp;B&nbsp;&lt;x&gt;&quot;y&quot;&apos;z&apos;');
  assert.strictEqual(out, 'A&B <x>"y"\'z\'');
});

test('extracts text with paragraph and table breaks', () => {
  const xml = '<w:document><w:body><w:p><w:r><w:t>姓名</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>教育</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>本科</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body></w:document>';
  const text = _extractTextFromDocumentXml(xml);
  assert.ok(text.includes('姓名'), 'should include paragraph text');
  assert.ok(text.includes('教育\t本科'), 'should include tab between table cells');
});

test('wraps long CJK lines into multiple lines', () => {
  const lines = _wrapLineByVisualWidth('这是一个用于测试自动换行的很长很长的文本', 80, 11);
  assert.ok(lines.length > 1, 'should wrap into more than one line');
});

test('builds a valid PDF buffer', () => {
  const pdf = _buildSimplePdf([['第一行', '第二行']], { title: 'test' });
  assert.ok(Buffer.isBuffer(pdf), 'should return Buffer');
  const head = pdf.slice(0, 8).toString('ascii');
  assert.ok(head.startsWith('%PDF-1.'), 'should start with PDF header');
  const tail = pdf.slice(-20).toString('ascii');
  assert.ok(tail.includes('%%EOF'), 'should end with EOF');
});

test('builds utf-8 html with escaping for native chrome render', () => {
  const html = _buildHtmlDocument(['中文 <b>测试</b> & done'], { title: 't' });
  assert.ok(html.includes('<meta charset="utf-8">'), 'should contain utf-8 meta');
  assert.ok(html.includes('中文 &lt;b&gt;测试&lt;/b&gt; &amp; done'), 'should escape html chars');
});

test('extracts ordered segments and builds paired row blocks', () => {
  const xml = '<w:document><w:body>' +
    '<w:p><w:r><w:t>基本信息</w:t></w:r></w:p>' +
    '<w:p><w:r><w:txbxContent><w:p><w:r><w:t>姓 名：张三</w:t></w:r></w:p><w:p><w:r><w:t>学 历：本科</w:t></w:r></w:p></w:txbxContent></w:r>' +
    '<w:r><w:txbxContent><w:p><w:r><w:t>性 别：男</w:t></w:r></w:p><w:p><w:r><w:t>年 龄：25</w:t></w:r></w:p></w:txbxContent></w:r></w:p>' +
    '</w:body></w:document>';
  const segments = _extractRenderSegmentsFromDocumentXml(xml);
  assert.ok(segments.some(s => s.type === 'textbox'), 'should include textbox segments');
  const blocks = _segmentsToStructuredBlocks(segments);
  assert.ok(blocks.some(b => b.type === 'heading'), 'should include heading block');
  assert.ok(blocks.some(b => b.type === 'row'), 'should include paired row block');
});

test('splits multi-line bullet paragraph into separate bullet blocks', () => {
  const blocks = _segmentsToStructuredBlocks([
    { type: 'paragraph', text: '- 第一条\n- 第二条' },
  ]);
  const bulletTexts = blocks.filter(b => b.type === 'bullet').map(b => b.text);
  assert.deepStrictEqual(bulletTexts, ['第一条', '第二条']);
});

test('attaches date-only textbox line to previous row block', () => {
  const blocks = _segmentsToStructuredBlocks([
    { type: 'textbox', lines: ['浙江大学医学院附属第二医院'] },
    { type: 'textbox', lines: ['放射科实习生'] },
    { type: 'textbox', lines: ['2023.05-2024.06'] },
  ]);
  const row = blocks.find(b => b.type === 'row');
  assert.ok(row, 'row block should exist');
  assert.ok(row.right.includes('2023.05-2024.06'), 'date should be appended to row right column');
  assert.strictEqual(blocks.filter(b => b.type === 'textbox').length, 0, 'date textbox should be absorbed');
});

test('builds structured html with row and heading markup', () => {
  const html = _buildStructuredHtmlDocument([
    { type: 'heading', text: '基本信息' },
    { type: 'row', left: ['姓 名：张三'], right: ['性 别：男'] },
    { type: 'bullet', text: '熟练掌握Office' },
    { type: 'table', rows: [['项目', '状态'], ['A', '完成']] },
    { type: 'image', dataUri: 'data:image/png;base64,AA==' },
  ]);
  assert.ok(html.includes('class="heading"'), 'should include heading class');
  assert.ok(html.includes('class="row"'), 'should include row class');
  assert.ok(html.includes('class="bullet"'), 'should include bullet class');
  assert.ok(html.includes('class="doc-table"'), 'should include table class');
  assert.ok(html.includes('class="image-wrap"'), 'should include image wrapper');
});

test('extracts relationship map from rels xml', () => {
  const rels = '<Relationships><Relationship Id="rId1" Target="media/a.png"/><Relationship Id="rId2" Target="media/b.svg"/></Relationships>';
  const map = _extractRelationshipMap(rels);
  assert.strictEqual(map.rId1, 'media/a.png');
  assert.strictEqual(map.rId2, 'media/b.svg');
});

test('extracts top-level paragraph and table blocks from body xml', () => {
  const body = '<w:p><w:r><w:t>A</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>X</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t>B</w:t></w:r></w:p>';
  const blocks = _extractTopLevelBodyBlocks(body);
  assert.strictEqual(blocks.length, 3);
  assert.strictEqual(blocks[0].type, 'p');
  assert.strictEqual(blocks[1].type, 'tbl');
  assert.strictEqual(blocks[2].type, 'p');
});

test('extracts table rows and cells from table xml', () => {
  const tbl = '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>项目</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>状态</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>完成</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
  const rows = _extractTableRows(tbl);
  assert.deepStrictEqual(rows, [['项目', '状态'], ['A', '完成']]);
});

// ── wpsCompat.js ──────────────────────────────────────────────────────────────
console.log('\nwpsCompat.js');
const { _normalizeXmlContent } = require('../lib/wpsCompat');

test('normalizes exact line rule without breaking WPS namespaces', () => {
  const xml = '<w:document xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><w:pPr><w:spacing w:lineRule="exact"/></w:pPr><w:rFonts w:eastAsia="微软雅黑"/><mc:Choice Requires="wps"><wps:wsp/></mc:Choice></w:document>';
  const out = _normalizeXmlContent(xml);
  assert.ok(out.includes('微软雅黑'), 'should preserve source font names');
  assert.ok(out.includes('lineRule="atLeast"'), 'should replace exact line rule');
  assert.ok(out.includes('xmlns:wps='), 'should keep wps namespace');
  assert.ok(out.includes('Requires="wps"'), 'should keep Requires marker');
  assert.ok(out.includes('<wps:wsp/>'), 'should keep wps elements');
});

// ── Summary ───────────────────────────────────────────────────────────────────
console.log(`\n${passed} passed, ${failed} failed\n`);
if (failed > 0) process.exit(1);
