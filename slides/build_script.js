const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, AlignmentType, BorderStyle, Footer, PageNumber, ShadingType } = require('docx');
const NOTES = require('./notes.json');

const FONT = 'Malgun Gothic';
const SAGE = '5E8C6A', MUTE = '6E7B73', INK = '33413A';
// 말하기 속도: 공백 제외 약 350자/분
const secs = t => Math.round(t.replace(/\s/g, '').length / 350 * 60 / 10) * 10;
const fmt = s => (s >= 60 ? `${Math.floor(s / 60)}분 ` : '') + (s % 60 ? `${s % 60}초` : '').trim();
const total = NOTES.reduce((a, n) => a + secs(n.script), 0);

const run = (text, o = {}) => new TextRun({ text, font: FONT, color: INK, size: 26, ...o });
const children = [
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [run('영유아교육과정의 이해', { bold: true, size: 40, color: SAGE })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 }, children: [run('발표 대본 (연습용)', { size: 24, color: MUTE })] }),
  new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { after: 360 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: 'C9D6CC', space: 8 } },
    children: [run(`슬라이드 6장  ·  예상 시간 약 ${fmt(total)}  ·  연습 체크  □ □ □ □ □`, { size: 20, color: MUTE })],
  }),
];

NOTES.forEach((n, i) => {
  children.push(new Paragraph({
    keepNext: true, spacing: { before: i ? 360 : 0, after: 160 },
    shading: { type: ShadingType.CLEAR, color: 'auto', fill: 'DCEBDD' },
    children: [
      run(`  슬라이드 ${i + 1}  `, { bold: true, size: 24, color: 'FFFFFF', shading: { type: ShadingType.CLEAR, color: 'auto', fill: SAGE } }),
      run(`  ${n.title}`, { bold: true, size: 26 }),
      run(`   (약 ${fmt(secs(n.script))})`, { size: 20, color: MUTE }),
    ],
  }));
  n.script.split(/\n+/).filter(Boolean).forEach(line => {
    children.push(new Paragraph({ spacing: { after: 140, line: 400 }, children: [run(line)] }));
  });
});

const doc = new Document({
  styles: { default: { document: { run: { font: FONT } } } },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1300, bottom: 1300, left: 1300, right: 1300 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 18, color: MUTE })] })] }) },
    children,
  }],
});
Packer.toBuffer(doc).then(b => fs.writeFileSync('영유아교육과정의_이해_발표대본.docx', b));
