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
function title(slide, text, n) {
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

module.exports = { FONT, C, icon, T, title, rings };
