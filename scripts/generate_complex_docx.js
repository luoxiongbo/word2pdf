#!/usr/bin/env node
'use strict';

const fs = require('fs');
const path = require('path');
const {
  AlignmentType,
  Document,
  HeadingLevel,
  ImageRun,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} = require('docx');

const outputPath = process.argv[2] || path.resolve(process.cwd(), 'complex_test.docx');

const tinyPngBase64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO7Z7S8AAAAASUVORK5CYII=';

const chartSvg = `
<svg xmlns="http://www.w3.org/2000/svg" width="1200" height="520" viewBox="0 0 1200 520">
  <rect width="1200" height="520" fill="#ffffff"/>
  <text x="600" y="56" text-anchor="middle" font-size="36" fill="#1f2937">季度销售图表（示例）</text>
  <line x1="130" y1="440" x2="1080" y2="440" stroke="#374151" stroke-width="3"/>
  <line x1="130" y1="120" x2="130" y2="440" stroke="#374151" stroke-width="3"/>
  <rect x="220" y="270" width="140" height="170" fill="#60a5fa"/>
  <rect x="420" y="220" width="140" height="220" fill="#34d399"/>
  <rect x="620" y="180" width="140" height="260" fill="#fbbf24"/>
  <rect x="820" y="140" width="140" height="300" fill="#f87171"/>
  <text x="290" y="468" text-anchor="middle" font-size="26" fill="#111827">Q1</text>
  <text x="490" y="468" text-anchor="middle" font-size="26" fill="#111827">Q2</text>
  <text x="690" y="468" text-anchor="middle" font-size="26" fill="#111827">Q3</text>
  <text x="890" y="468" text-anchor="middle" font-size="26" fill="#111827">Q4</text>
  <text x="290" y="255" text-anchor="middle" font-size="24" fill="#111827">85</text>
  <text x="490" y="205" text-anchor="middle" font-size="24" fill="#111827">110</text>
  <text x="690" y="165" text-anchor="middle" font-size="24" fill="#111827">130</text>
  <text x="890" y="125" text-anchor="middle" font-size="24" fill="#111827">150</text>
</svg>
`;

const photoSvg = `
<svg xmlns="http://www.w3.org/2000/svg" width="720" height="420" viewBox="0 0 720 420">
  <defs>
    <linearGradient id="g1" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#dbeafe"/>
      <stop offset="100%" stop-color="#bfdbfe"/>
    </linearGradient>
  </defs>
  <rect width="720" height="420" fill="url(#g1)"/>
  <circle cx="160" cy="130" r="60" fill="#93c5fd"/>
  <rect x="70" y="220" width="180" height="120" rx="18" fill="#60a5fa"/>
  <path d="M300 320 L410 190 L500 290 L560 230 L670 320 Z" fill="#3b82f6"/>
  <text x="360" y="70" text-anchor="middle" font-size="34" fill="#1e3a8a">示例图片区域</text>
  <text x="360" y="110" text-anchor="middle" font-size="24" fill="#1e40af">用于测试复杂文档中的图片渲染</text>
</svg>
`;

const imageFallback = {
  type: 'png',
  data: Buffer.from(tinyPngBase64, 'base64'),
  transformation: { width: 640, height: 240 },
};

const table = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({
      children: [
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: '项目', bold: true })] })] }),
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: '负责人', bold: true })] })] }),
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: '状态', bold: true })] })] }),
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: '完成度', bold: true })] })] }),
      ],
    }),
    new TableRow({
      children: [
        new TableCell({ children: [new Paragraph('数据中台升级')] }),
        new TableCell({ children: [new Paragraph('张三')] }),
        new TableCell({ children: [new Paragraph('进行中')] }),
        new TableCell({ children: [new Paragraph('78%')] }),
      ],
    }),
    new TableRow({
      children: [
        new TableCell({ children: [new Paragraph('客服流程优化')] }),
        new TableCell({ children: [new Paragraph('李四')] }),
        new TableCell({ children: [new Paragraph('已完成')] }),
        new TableCell({ children: [new Paragraph('100%')] }),
      ],
    }),
    new TableRow({
      children: [
        new TableCell({ children: [new Paragraph('供应链可视化')] }),
        new TableCell({ children: [new Paragraph('王五')] }),
        new TableCell({ children: [new Paragraph('规划中')] }),
        new TableCell({ children: [new Paragraph('35%')] }),
      ],
    }),
  ],
});

const doc = new Document({
  sections: [
    {
      children: [
        new Paragraph({
          text: '复杂文档转换测试样本',
          heading: HeadingLevel.HEADING_1,
          alignment: AlignmentType.CENTER,
        }),
        new Paragraph({
          children: [
            new TextRun('本样本包含：'),
            new TextRun({ text: '表格', bold: true }),
            new TextRun('、'),
            new TextRun({ text: '图表', bold: true }),
            new TextRun('、'),
            new TextRun({ text: '图片', bold: true }),
            new TextRun('，用于验证 DOCX 转 PDF 复杂场景效果。'),
          ],
        }),
        new Paragraph(''),
        new Paragraph({ text: '1) 业务表格', heading: HeadingLevel.HEADING_2 }),
        table,
        new Paragraph(''),
        new Paragraph({ text: '2) 图表（SVG）', heading: HeadingLevel.HEADING_2 }),
        new Paragraph({
          children: [
            new ImageRun({
              type: 'svg',
              data: Buffer.from(chartSvg, 'utf8'),
              fallback: imageFallback,
              transformation: { width: 620, height: 260 },
            }),
          ],
        }),
        new Paragraph({ text: '3) 图片（SVG）', heading: HeadingLevel.HEADING_2 }),
        new Paragraph({
          children: [
            new ImageRun({
              type: 'svg',
              data: Buffer.from(photoSvg, 'utf8'),
              fallback: imageFallback,
              transformation: { width: 540, height: 300 },
            }),
          ],
        }),
        new Paragraph({ text: '4) 文本要点', heading: HeadingLevel.HEADING_2 }),
        new Paragraph({ text: '转换后检查表格边框、标题层级、图像是否完整保留。', bullet: { level: 0 } }),
        new Paragraph({ text: '重点关注分页、行距、中文字体是否稳定。', bullet: { level: 0 } }),
      ],
    },
  ],
});

Packer.toBuffer(doc)
  .then((buffer) => {
    fs.mkdirSync(path.dirname(outputPath), { recursive: true });
    fs.writeFileSync(outputPath, buffer);
    console.log(`Generated DOCX: ${outputPath}`);
  })
  .catch((err) => {
    console.error(err);
    process.exit(1);
  });
