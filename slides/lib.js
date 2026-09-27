const React = require('react');
const ReactDOMServer = require('react-dom/server');
const sharp = require('sharp');
const fa = require('react-icons/fa');

const FONT = 'Malgun Gothic';
const C = {
  sage: '5E8C6A', sageL: 'DCEBDD', sageM: 'A9CBB0',
  peach: 'F2A27E', peachL: 'FCE4D8',
  butter: 'F5D07A', butterL: 'FDF3D6',
  rose: 'E8A1A8', roseL: 'F9E1E3',
  ink: '33413A', mute: '6E7B73', white: 'FFFFFF',
};

async function icon(name, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(fa[name], { color: '#' + color, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return 'image/png;base64,' + buf.toString('base64');
}

function T(slide, text, o) {
  slide.addText(text, Object.assign({ fontFace: FONT, color: C.ink, isTextBox: true, margin: 0 }, o));
}
// 합본에서는 state.pres를 지정하면 슬라이드 번호가 전체 쪽 번호로 매겨진다
const state = { pres: null };
function title(slide, text, n) {
  if (state.pres) n = state.pres.slides.length;
  slide.addShape('ellipse', { x: 0.5, y: 0.42, w: 0.55, h: 0.55, fill: { color: C.sage } });
  T(slide, String(n).padStart(2, '0'), { x: 0.5, y: 0.42, w: 0.55, h: 0.55, fontSize: 14, bold: true, color: C.white, align: 'center', valign: 'middle' });
  T(slide, text, { x: 1.2, y: 0.35, w: 8.3, h: 0.7, fontSize: 28, bold: true, valign: 'middle' });
}
function rings(slide, cx, cy, r, colors) {
  colors.forEach((col, i) => {
    const rr = r * (1 - i / colors.length);
    slide.addShape('ellipse', { x: cx - rr, y: cy - rr, w: rr * 2, h: rr * 2, fill: { color: col } });
  });
}

// 표지. part가 있으면 합본용 PART 구분 슬라이드(발표자 칸 없음)
function cover(pres, { title: t, sub, part, titleSize = 44, twoLine = false }) {
  const s = pres.addSlide();
  s.background = { color: C.sage };
  rings(s, 8.6, 1.3, 1.6, ['7FA88A', 'A9CBB0', 'DCEBDD', 'F5D07A']);
  rings(s, 9.4, 4.7, 0.8, ['F2A27E', 'FCE4D8', 'F2A27E']);
  T(s, part ? `PART ${part}` : '05  교육과정의 계획과 운영', { x: 0.7, y: 1.35, w: 6, h: 0.4, fontSize: 16, bold: !!part, color: C.butterL });
  T(s, t, { x: 0.7, y: 1.8, w: 7, h: twoLine ? 1.5 : 1.0, fontSize: titleSize, bold: true, color: C.white, ...(twoLine ? { valign: 'top' } : {}) });
  const subY = twoLine ? 3.4 : 2.85;
  T(s, sub, { x: 0.7, y: subY, w: 7, h: 0.45, fontSize: 18, color: C.white });
  if (!part) T(s, '과목명  |  학번  |  발표자 이름', { x: 0.7, y: twoLine ? 4.5 : 4.3, w: 6, h: 0.4, fontSize: 14, color: C.sageL });
  return s;
}

module.exports = { FONT, C, icon, T, title, rings, cover, state };
