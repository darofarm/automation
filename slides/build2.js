const pptxgen = require('pptxgenjs');
const { C, icon, T, title, rings, cover } = require('./lib');
const DEFAULT_NOTES = require('./notes2.json');

const card = (s, x, y, w, h, fill) => s.addShape('roundRect', { x, y, w, h, fill: { color: fill }, rectRadius: 0.15 });
async function dot(s, x, y, d, fill, ic) {
  s.addShape('ellipse', { x, y, w: d, h: d, fill: { color: fill } });
  s.addImage({ data: await icon(ic, C.white), x: x + d * 0.25, y: y + d * 0.25, w: d * 0.5, h: d * 0.5 });
}

async function build(pres, opts = {}) {
  const NOTES = opts.notes || DEFAULT_NOTES;
  let s = cover(pres, { title: '영유아교육과정의 구성', sub: '구성 요소와 구성의 네 가지 특징', part: opts.part, titleSize: 44, twoLine: false });
  s.addNotes(NOTES[0].script);

  // ---------- 2. 구성 요소: 타일러의 네 가지 질문 ----------
  s = pres.addSlide();
  s.background = { color: C.white };
  title(s, '교육과정의 구성 요소', 2);
  T(s, '타일러(Tyler)가 제시한 네 가지 질문', { x: 1.2, y: 1.0, w: 8.3, h: 0.3, fontSize: 13, color: C.mute });
  const el = [
    { k: '교육목표', ic: 'FaBullseye', c: C.peach, f: C.peachL, q: '학교가 성취해야 할\n교육목표는 무엇인가?' },
    { k: '교육내용', ic: 'FaBookOpen', c: C.butter, f: C.butterL, q: '목표 달성을 위해 어떤\n교육적 경험을 제공할까?' },
    { k: '교수-학습 방법', ic: 'FaChalkboardTeacher', c: C.sage, f: C.sageL, q: '교육적 경험을 어떻게\n효과적으로 조직할까?' },
    { k: '평가', ic: 'FaClipboardCheck', c: C.rose, f: C.roseL, q: '목표가 달성되었는지\n어떻게 알 수 있을까?' },
  ];
  for (let i = 0; i < 4; i++) {
    const x = 0.5 + i * 2.3, e = el[i];
    card(s, x, 1.5, 2.1, 3.1, e.f);
    await dot(s, x + 0.7, 1.75, 0.7, e.c, e.ic);
    T(s, e.q, { x: x + 0.15, y: 2.65, w: 1.8, h: 0.9, fontSize: 12, color: C.mute, align: 'center', valign: 'top' });
    T(s, '▼', { x: x + 0.15, y: 3.5, w: 1.8, h: 0.3, fontSize: 11, color: e.c === C.butter ? 'B08A2E' : e.c, align: 'center' });
    T(s, e.k, { x: x + 0.1, y: 3.85, w: 1.9, h: 0.5, fontSize: 17, bold: true, align: 'center', valign: 'middle' });
  }
  T(s, '→ 우리나라 영유아교육과정에도 이 네 가지 요소가 포함된다', { x: 0.5, y: 4.85, w: 9.0, h: 0.35, fontSize: 14, bold: true, color: C.sage });
  s.addNotes(NOTES[1].script);

  // ---------- 3. 실행 순서 (순환) ----------
  s = pres.addSlide();
  s.background = { color: C.white };
  title(s, '구성 요소 = 교육과정의 실행 순서', 3);
  const cx = 3.05, cy = 3.25, Rx = 1.75, Ry = 1.45, bw = 1.7;
  s.addShape('ellipse', { x: cx - Rx, y: cy - Ry, w: Rx * 2, h: Ry * 2, line: { color: 'C9D6CC', width: 3, dashType: 'dash' } });
  T(s, '계속\n순환', { x: cx - 0.6, y: cy - 0.35, w: 1.2, h: 0.7, fontSize: 14, bold: true, color: C.mute, align: 'center', valign: 'middle' });
  const steps = [
    { t: '교육목표 설정', c: C.peach },
    { t: '교육내용 선정', c: C.butter },
    { t: '내용 조직·\n교수방법 결정', c: C.sage },
    { t: '평가', c: C.rose },
    { t: '다음 계획에\n반영', c: C.sageM },
  ];
  steps.forEach((e, i) => {
    const a = (-90 + i * 72) * Math.PI / 180;
    const x = cx + Rx * Math.cos(a), y = cy + Ry * Math.sin(a);
    s.addShape('roundRect', { x: x - bw / 2, y: y - 0.33, w: bw, h: 0.66, fill: { color: 'FFFFFF' }, line: { color: e.c, width: 2 }, rectRadius: 0.15 });
    s.addShape('ellipse', { x: x - bw / 2 + 0.1, y: y - 0.16, w: 0.32, h: 0.32, fill: { color: e.c } });
    T(s, String(i + 1), { x: x - bw / 2 + 0.1, y: y - 0.16, w: 0.32, h: 0.32, fontSize: 11, bold: true, color: C.white, align: 'center', valign: 'middle' });
    T(s, e.t, { x: x - bw / 2 + 0.5, y: y - 0.33, w: bw - 0.6, h: 0.66, fontSize: 11, bold: true, valign: 'middle' });
  });
  card(s, 5.9, 1.4, 3.6, 1.6, C.sageL);
  T(s, '각 단계의 기준', { x: 6.1, y: 1.5, w: 3.2, h: 0.35, fontSize: 14, bold: true, color: C.sage });
  T(s, [
    { text: '목표: 기관의 철학에 부합', options: { bullet: true, breakLine: true } },
    { text: '내용: 목표 달성에 꼭 필요한 것', options: { bullet: true, breakLine: true } },
    { text: '방법: 영유아의 발달과 흥미에 적합', options: { bullet: true } },
  ], { x: 6.1, y: 1.9, w: 3.3, h: 1.0, fontSize: 12, paraSpaceAfter: 4 });
  card(s, 5.9, 3.2, 3.6, 1.55, C.butterL);
  T(s, '타일러 모형의 영향', { x: 6.1, y: 3.3, w: 3.2, h: 0.35, fontSize: 14, bold: true, color: 'B08A2E' });
  T(s, '비판을 받으면서도 우리나라 초·중·고 교육과정의 기본 틀이 되었고, 교육과정 개념화에도 큰 영향을 미침', { x: 6.1, y: 3.7, w: 3.25, h: 0.95, fontSize: 12, valign: 'top' });
  s.addNotes(NOTES[2].script);

  // ---------- 4. 특징 ① 통합 ② 놀이 중심 ----------
  s = pres.addSlide();
  s.background = { color: C.white };
  title(s, '구성의 특징 ①  ②', 4);
  card(s, 0.5, 1.2, 4.4, 3.95, C.peachL);
  T(s, '① 교육내용 간의 통합', { x: 0.7, y: 1.3, w: 4.0, h: 0.4, fontSize: 16, bold: true });
  T(s, '분리된 교과별 학습보다 통합된 경험으로 더 잘 배운다', { x: 0.7, y: 1.7, w: 4.0, h: 0.3, fontSize: 11, color: C.mute });
  // '가을' 주제망
  const hx = 2.7, hy = 3.4;
  const sat = [{ t: '가을 과일\n만들기', x: 1.4, y: 2.45 }, { t: '과일가게\n놀이', x: 4.0, y: 2.45 }, { t: '간판·가격표\n쓰기', x: 2.7, y: 4.55 }];
  sat.forEach(p => s.addShape('line', { x: Math.min(hx, p.x), y: Math.min(hy, p.y), w: Math.abs(hx - p.x), h: Math.abs(hy - p.y), flipH: (p.x < hx) !== (p.y < hy), line: { color: C.peach, width: 2 } }));
  s.addShape('ellipse', { x: hx - 0.6, y: hy - 0.45, w: 1.2, h: 0.9, fill: { color: C.peach } });
  T(s, "주제\n'가을'", { x: hx - 0.6, y: hy - 0.45, w: 1.2, h: 0.9, fontSize: 13, bold: true, color: C.white, align: 'center', valign: 'middle' });
  sat.forEach(p => {
    s.addShape('roundRect', { x: p.x - 0.7, y: p.y - 0.3, w: 1.4, h: 0.6, fill: { color: C.white }, line: { color: C.peach, width: 1.5 }, rectRadius: 0.12 });
    T(s, p.t, { x: p.x - 0.7, y: p.y - 0.3, w: 1.4, h: 0.6, fontSize: 11, align: 'center', valign: 'middle' });
  });
  card(s, 5.1, 1.2, 4.4, 3.95, C.sageL);
  T(s, '② 놀이 중심 교육과정으로 운영', { x: 5.3, y: 1.3, w: 4.0, h: 0.4, fontSize: 16, bold: true });
  T(s, '유아와 놀이를 최우선으로 존중', { x: 5.3, y: 1.7, w: 4.0, h: 0.3, fontSize: 11, color: C.mute });
  const play = [
    ['FaClock', '충분한 놀이시간 확보'],
    ['FaChild', '유아 스스로 놀이하며 배움'],
    ['FaTree', '바깥놀이·자연물 놀이 장려'],
    ['FaHandsHelping', '교사 = 놀이지원자\n(관찰·지원·상호작용)'],
  ];
  for (let i = 0; i < play.length; i++) {
    const y = 2.2 + i * 0.7;
    await dot(s, 5.35, y, 0.5, C.sage, play[i][0]);
    T(s, play[i][1], { x: 6.0, y: y - 0.05, w: 3.35, h: 0.6, fontSize: 13, valign: 'middle', bold: i === 3 });
  }
  s.addNotes(NOTES[3].script);

  // ---------- 5. 특징 ③ 흥미·요구 ④ 다양성 ----------
  s = pres.addSlide();
  s.background = { color: C.white };
  title(s, '구성의 특징 ③  ④', 5);
  card(s, 0.5, 1.2, 4.4, 3.95, C.butterL);
  T(s, '③ 흥미와 요구를 반영한 교육적 경험', { x: 0.7, y: 1.3, w: 4.1, h: 0.4, fontSize: 16, bold: true });
  s.addShape('roundRect', { x: 0.7, y: 1.8, w: 4.0, h: 0.45, fill: { color: C.white }, line: { color: C.rose, width: 1.5 }, rectRadius: 0.2 });
  T(s, "주의! '좋아하는 것 = 좋은 교육과정'은 아님", { x: 0.7, y: 1.8, w: 4.0, h: 0.45, fontSize: 12, bold: true, color: 'B45A64', align: 'center', valign: 'middle' });
  const flow = ['영유아의 흥미 존중 + 발달적 요구', '교사의 주의 깊은 관찰·평가', '의미 있는 주제·놀이 선별', '교육적 활동으로 연결'];
  flow.forEach((t, i) => {
    const y = 2.45 + i * 0.66;
    s.addShape('roundRect', { x: 0.9, y, w: 3.6, h: 0.46, fill: { color: i === 3 ? C.butter : C.white }, rectRadius: 0.12 });
    T(s, t, { x: 0.9, y, w: 3.6, h: 0.46, fontSize: 12, bold: i === 3, align: 'center', valign: 'middle' });
    if (i < 3) T(s, '▼', { x: 0.9, y: y + 0.46, w: 3.6, h: 0.2, fontSize: 8, color: 'B08A2E', align: 'center', valign: 'middle' });
  });
  card(s, 5.1, 1.2, 4.4, 3.95, C.roseL);
  T(s, '④ 개인차와 문화적 특성 등 다양성 반영', { x: 5.3, y: 1.3, w: 4.1, h: 0.4, fontSize: 16, bold: true });
  ['발달 특성', '개인차', '기질', '사회문화적 배경'].forEach((t, i) => {
    const w = [0.95, 0.75, 0.6, 1.45][i], x = 5.3 + [0, 1.05, 1.9, 2.6][i];
    s.addShape('roundRect', { x, y: 1.85, w, h: 0.38, fill: { color: C.white }, rectRadius: 0.19 });
    T(s, t, { x, y: 1.85, w, h: 0.38, fontSize: 11, align: 'center', valign: 'middle' });
  });
  const div = [
    ['FaBaby', '영아', '개인별 보육계획 수립'],
    ['FaLayerGroup', '누리과정', '개인별 평가 기반 수준별 계획'],
    ['FaGlobeAsia', '다문화 사회', '다문화교육 계획'],
  ];
  for (let i = 0; i < div.length; i++) {
    const y = 2.5 + i * 0.85;
    await dot(s, 5.35, y + 0.05, 0.5, C.rose, div[i][0]);
    T(s, div[i][1], { x: 6.0, y, w: 3.3, h: 0.3, fontSize: 11, color: C.mute });
    T(s, div[i][2], { x: 6.0, y: y + 0.28, w: 3.35, h: 0.35, fontSize: 13, bold: true });
  }
  s.addNotes(NOTES[4].script);

  // ---------- 6. 한눈에 정리 ----------
  s = pres.addSlide();
  s.background = { color: C.sage };
  T(s, '한눈에 정리', { x: 0.5, y: 0.4, w: 9, h: 0.7, fontSize: 30, bold: true, color: C.white });
  const rowItems = (label, items, y, fill) => {
    T(s, label, { x: 0.5, y, w: 9, h: 0.35, fontSize: 14, bold: true, color: C.butterL });
    items.forEach((t, i) => {
      const x = 0.5 + i * 2.3;
      s.addShape('roundRect', { x, y: y + 0.45, w: 2.0, h: 0.7, fill: { color: fill }, rectRadius: 0.15 });
      T(s, t, { x, y: y + 0.45, w: 2.0, h: 0.7, fontSize: 15, bold: true, align: 'center', valign: 'middle' });
      if (i < items.length - 1) T(s, '→', { x: x + 2.0, y: y + 0.45, w: 0.3, h: 0.7, fontSize: 14, bold: true, color: C.white, align: 'center', valign: 'middle' });
    });
  };
  rowItems('구성 요소 (= 실행 순서)', ['교육목표', '교육내용', '교수-학습 방법', '평가'], 1.25, C.sageL);
  const feats = ['통합', '놀이 중심', '흥미·요구 반영', '다양성 반영'];
  T(s, '구성의 특징', { x: 0.5, y: 2.7, w: 9, h: 0.35, fontSize: 14, bold: true, color: C.butterL });
  feats.forEach((t, i) => {
    const x = 0.5 + i * 2.3;
    s.addShape('roundRect', { x, y: 3.15, w: 2.0, h: 0.7, fill: { color: C.butterL }, rectRadius: 0.15 });
    T(s, t, { x, y: 3.15, w: 2.0, h: 0.7, fontSize: 15, bold: true, align: 'center', valign: 'middle' });
  });
  T(s, '아이가 놀이 속에서 자연스럽게 배우도록 구성하는 것이 핵심', { x: 0.5, y: 4.4, w: 9, h: 0.5, fontSize: 17, bold: true, color: C.white });
  s.addNotes(NOTES[5].script);

}

module.exports = { build };

if (require.main === module) (async () => {
  const pres = new pptxgen();
  pres.layout = 'LAYOUT_16x9';
  pres.title = '영유아교육과정의 구성';
  await build(pres);
  await pres.writeFile({ fileName: '영유아교육과정의_구성.pptx' });
})();
