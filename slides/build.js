const pptxgen = require('pptxgenjs');
const { FONT, C, icon, T, title, rings, cover } = require('./lib');
const DEFAULT_NOTES = require('./notes.json');

async function build(pres, opts = {}) {
  const NOTES = opts.notes || DEFAULT_NOTES;
  let s = cover(pres, { title: '영유아교육과정의 이해', sub: '교육과정의 정의부터 누리과정·표준보육과정까지', part: opts.part, titleSize: 44, twoLine: false });
  s.addNotes(NOTES[0].script);

  // ---------- 2. 교육과정이란? ----------
  s = pres.addSlide();
  s.background = { color: C.white };
  title(s, '교육과정이란?', 2);
  const teach = await icon('FaChalkboardTeacher', C.white);
  const learn = await icon('FaChild', C.white);
  const cards = [
    { y: 1.3, fill: C.peachL, dot: C.peach, img: teach, h: '교사 중심으로 보면', b: '가르치는 내용 · 교과 · 의도적 계획' },
    { y: 2.55, fill: C.sageL, dot: C.sage, img: learn, h: '학생 중심으로 보면', b: '학습경험 · 학습 결과' },
  ];
  for (const c of cards) {
    s.addShape('roundRect', { x: 0.5, y: c.y, w: 4.6, h: 1.05, fill: { color: c.fill }, rectRadius: 0.15 });
    s.addShape('ellipse', { x: 0.7, y: c.y + 0.2, w: 0.65, h: 0.65, fill: { color: c.dot } });
    s.addImage({ data: c.img, x: 0.85, y: c.y + 0.35, w: 0.35, h: 0.35 });
    T(s, c.h, { x: 1.55, y: c.y + 0.15, w: 3.4, h: 0.35, fontSize: 13, color: C.mute });
    T(s, c.b, { x: 1.55, y: c.y + 0.5, w: 3.45, h: 0.4, fontSize: 14, bold: true });
  }
  T(s, '(이귀윤, 1990)', { x: 0.5, y: 3.68, w: 4.6, h: 0.25, fontSize: 10, color: C.mute, align: 'right' });
  // 넓은/좁은 의미 동심원
  s.addShape('ellipse', { x: 5.55, y: 1.15, w: 3.9, h: 2.9, fill: { color: C.butterL }, line: { color: C.butter, width: 1.5 } });
  T(s, '넓은 의미', { x: 5.55, y: 1.35, w: 3.9, h: 0.3, fontSize: 13, bold: true, color: 'B08A2E', align: 'center' });
  T(s, '"학생이 경험하는 모든 것"', { x: 5.55, y: 1.65, w: 3.9, h: 0.3, fontSize: 13, align: 'center' });
  s.addShape('ellipse', { x: 6.45, y: 2.2, w: 2.1, h: 1.55, fill: { color: C.roseL }, line: { color: C.rose, width: 1.5 } });
  T(s, '좁은 의미', { x: 6.45, y: 2.45, w: 2.1, h: 0.3, fontSize: 13, bold: true, color: 'B45A64', align: 'center' });
  T(s, '계획된 교과목\n또는 교수요목', { x: 6.45, y: 2.8, w: 2.1, h: 0.6, fontSize: 12, align: 'center' });
  // 정의 인용
  s.addShape('roundRect', { x: 0.5, y: 4.05, w: 9.0, h: 1.15, fill: { color: 'F4F7F4' }, rectRadius: 0.12 });
  T(s, [
    { text: '"학습자의 전인적 성장과 발달을 돕기 위해서 학교가 의미하는 교육적 상황을 마련하고 모든 종류의 교수-학습 경험을 체계적으로 구성한 ', options: {} },
    { text: '설계도', options: { bold: true, color: C.sage } },
    { text: '"', options: {} },
  ], { x: 0.75, y: 4.12, w: 8.5, h: 0.7, fontSize: 13, valign: 'middle' });
  T(s, '방인옥 등(1999)', { x: 0.75, y: 4.85, w: 8.5, h: 0.28, fontSize: 11, color: C.mute });
  s.addNotes(NOTES[1].script);

  // ---------- 3. 유치원교육과정 → 누리과정 ----------
  s = pres.addSlide();
  s.background = { color: C.white };
  title(s, '유치원교육과정에서 누리과정으로', 3);
  const tl = [
    { y: '1969', t: '유치원교육과정\n처음 제정', c: C.sageM },
    { y: '~2007', t: '여섯 차례\n개정', c: C.butter },
    { y: '2013', t: '만 3~5세\n누리과정 시행', c: C.peach },
    { y: '2020', t: '개정 누리과정 적용\n(유아중심·놀이중심)', c: C.sage },
  ];
  const lineY = 1.85;
  s.addShape('line', { x: 1.3, y: lineY, w: 7.4, h: 0, line: { color: 'C9D6CC', width: 3 } });
  tl.forEach((e, i) => {
    const cx = 1.3 + i * (7.4 / 3);
    s.addShape('ellipse', { x: cx - 0.2, y: lineY - 0.2, w: 0.4, h: 0.4, fill: { color: e.c }, line: { color: C.white, width: 3 } });
    T(s, e.y, { x: cx - 1.1, y: 1.2, w: 2.2, h: 0.4, fontSize: 20, bold: true, color: e.c === C.butter ? 'B08A2E' : e.c, align: 'center' });
    T(s, e.t, { x: cx - 1.1, y: 2.2, w: 2.2, h: 0.7, fontSize: 13, align: 'center', valign: 'top' });
  });
  // 일원화 도식
  s.addShape('roundRect', { x: 0.5, y: 3.25, w: 2.6, h: 0.75, fill: { color: C.peachL }, rectRadius: 0.12 });
  T(s, '유치원교육과정', { x: 0.5, y: 3.25, w: 2.6, h: 0.75, fontSize: 15, bold: true, align: 'center', valign: 'middle' });
  s.addShape('roundRect', { x: 0.5, y: 4.2, w: 2.6, h: 0.75, fill: { color: C.sageL }, rectRadius: 0.12 });
  T(s, '어린이집 표준보육과정', { x: 0.5, y: 4.2, w: 2.6, h: 0.75, fontSize: 15, bold: true, align: 'center', valign: 'middle' });
  s.addShape('rightArrow', { x: 3.3, y: 3.8, w: 0.8, h: 0.6, fill: { color: C.butter } });
  T(s, '일원화', { x: 3.1, y: 3.4, w: 1.2, h: 0.35, fontSize: 12, bold: true, color: 'B08A2E', align: 'center' });
  s.addShape('roundRect', { x: 4.3, y: 3.25, w: 5.2, h: 1.7, fill: { color: C.butterL }, rectRadius: 0.15 });
  T(s, '누리과정 (만 3~5세 공통교육과정)', { x: 4.55, y: 3.4, w: 4.8, h: 0.4, fontSize: 17, bold: true, color: C.sage });
  T(s, [
    { text: '취학 전 교육의 질 제고', options: { bullet: true, breakLine: true } },
    { text: '생애초기 출발점 평등 보장', options: { bullet: true } },
  ], { x: 4.55, y: 3.9, w: 4.8, h: 0.8, fontSize: 15, paraSpaceAfter: 6 });
  s.addNotes(NOTES[2].script);

  // ---------- 4. 어린이집 표준보육과정 ----------
  s = pres.addSlide();
  s.background = { color: C.white };
  title(s, '어린이집 표준보육과정', 4);
  T(s, '표준보육과정 개발의 배경', { x: 0.5, y: 1.2, w: 4.2, h: 0.35, fontSize: 15, bold: true, color: C.sage });
  const bg = [
    { i: 'FaHome', c: C.peach, t: '핵가족화' },
    { i: 'FaBriefcase', c: C.butter, t: '취업모의 지속적인 증가' },
    { i: 'FaSeedling', c: C.sage, t: '영유아기 교육의 중요성 부각' },
  ];
  for (let k = 0; k < bg.length; k++) {
    const y = 1.7 + k * 0.62;
    s.addShape('ellipse', { x: 0.5, y, w: 0.48, h: 0.48, fill: { color: bg[k].c } });
    s.addImage({ data: await icon(bg[k].i, C.white), x: 0.62, y: y + 0.12, w: 0.24, h: 0.24 });
    T(s, bg[k].t, { x: 1.15, y, w: 3.4, h: 0.48, fontSize: 15, valign: 'middle' });
  }
  s.addShape('roundRect', { x: 0.5, y: 3.65, w: 4.1, h: 1.35, fill: { color: C.sageL }, rectRadius: 0.12 });
  T(s, [
    { text: '영유아가 어린이집에 다니는 것이 보편화', options: { bold: true, breakLine: true } },
    { text: '→ 언제, 무엇을 목표로 어떤 경험을 제공할 것인가에 대한 ', options: {} },
    { text: '국가수준의 지침', options: { bold: true, color: C.sage } },
    { text: ' 필요', options: {} },
  ], { x: 0.7, y: 3.75, w: 3.8, h: 1.15, fontSize: 13, valign: 'middle', paraSpaceAfter: 4 });
  // 세로 연표
  const vt = [
    { y: '2004', t: '영유아보육법 전면 개정\n제29조 제2항에 개발·보급 명시' },
    { y: '2007', t: '표준보육과정 공포' },
    { y: '2013', t: '제3차 어린이집 표준보육과정으로\n개정' },
    { y: '2020', t: '제4차 어린이집 표준보육과정 고시' },
    { y: '2025', t: '2024 개정 표준보육과정(0~2세) 고시' },
  ];
  s.addShape('line', { x: 5.55, y: 1.35, w: 0, h: 3.55, line: { color: 'C9D6CC', width: 2.5 } });
  vt.forEach((e, k) => {
    const y = 1.2 + k * 0.78;
    s.addShape('ellipse', { x: 5.43, y: y + 0.08, w: 0.24, h: 0.24, fill: { color: k === 4 ? C.peach : C.sage } });
    T(s, e.y, { x: 5.85, y, w: 0.8, h: 0.4, fontSize: 16, bold: true, color: k === 4 ? 'D07A55' : C.sage, valign: 'middle' });
    T(s, e.t, { x: 6.7, y: y + 0.02, w: 2.9, h: 0.7, fontSize: 12, valign: 'top' });
  });
  s.addNotes(NOTES[3].script);

  // ---------- 5. 2019 개정 누리과정 ----------
  s = pres.addSlide();
  s.background = { color: C.white };
  title(s, '2019 개정 누리과정의 혁신 방향', 5);
  T(s, "출발선 평등 실현을 위한 '유아교육 혁신방안'에 의거 (교육부, 2017)", { x: 1.2, y: 1.0, w: 8.3, h: 0.3, fontSize: 12, color: C.mute });
  const inn = [
    { n: '첫째', i: 'FaChild', c: C.peach, f: C.peachL, from: '초등학교 준비 위주의 학습', to: '개별 유아에 대한 다양한\n특성을 고려한 내용' },
    { n: '둘째', i: 'FaPuzzlePiece', c: C.butter, f: C.butterL, from: '교사계획서 위주, 학습 위주의 교육', to: '유아의 자유놀이 권장\n관찰과 기록 등 유아와의\n상호작용 강조' },
    { n: '셋째', i: 'FaLeaf', c: C.sage, f: C.sageL, from: '5개 영역, 주제-소주제 등\n구성체계가 지나치게 세부적', to: '현장의 자율성 존중을 위해\n세부내용 삭제' },
  ];
  for (let k = 0; k < 3; k++) {
    const x = 0.5 + k * 3.1, e = inn[k];
    s.addShape('roundRect', { x, y: 1.5, w: 2.8, h: 2.75, fill: { color: e.f }, rectRadius: 0.15 });
    s.addShape('ellipse', { x: x + 0.2, y: 1.7, w: 0.6, h: 0.6, fill: { color: e.c } });
    s.addImage({ data: await icon(e.i, C.white), x: x + 0.35, y: 1.85, w: 0.3, h: 0.3 });
    T(s, e.n, { x: x + 0.95, y: 1.7, w: 1.6, h: 0.6, fontSize: 18, bold: true, valign: 'middle' });
    T(s, e.from, { x: x + 0.2, y: 2.45, w: 2.4, h: 0.55, fontSize: 11, color: C.mute, valign: 'top' });
    T(s, '▼', { x: x + 0.2, y: 3.0, w: 2.4, h: 0.25, fontSize: 10, color: e.c === C.butter ? 'B08A2E' : e.c });
    T(s, e.to, { x: x + 0.2, y: 3.28, w: 2.45, h: 0.95, fontSize: 13, bold: true, valign: 'top' });
  }
  s.addShape('roundRect', { x: 0.5, y: 4.45, w: 9.0, h: 0.7, fill: { color: C.sage }, rectRadius: 0.35 });
  T(s, "'유아 중심·놀이 중심 교육과정'  →  2020년부터 유치원과 어린이집 공통교육과정으로 적용", { x: 0.5, y: 4.45, w: 9.0, h: 0.7, fontSize: 14, bold: true, color: C.white, align: 'center', valign: 'middle' });
  s.addNotes(NOTES[4].script);

  // ---------- 6. 한눈에 정리 ----------
  s = pres.addSlide();
  s.background = { color: C.sage };
  T(s, '한눈에 정리', { x: 0.5, y: 0.4, w: 9, h: 0.7, fontSize: 30, bold: true, color: C.white });
  // 연령 막대
  s.addShape('roundRect', { x: 0.5, y: 1.65, w: 3.6, h: 1.1, fill: { color: C.sageL }, rectRadius: 0.15 });
  T(s, '표준보육과정', { x: 0.7, y: 1.75, w: 3.2, h: 0.45, fontSize: 18, bold: true, color: C.sage });
  T(s, '0~2세 · 어린이집', { x: 0.7, y: 2.2, w: 3.2, h: 0.4, fontSize: 13 });
  s.addShape('roundRect', { x: 4.3, y: 1.65, w: 5.2, h: 1.1, fill: { color: C.butterL }, rectRadius: 0.15 });
  T(s, '누리과정', { x: 4.5, y: 1.75, w: 4.8, h: 0.45, fontSize: 18, bold: true, color: 'B08A2E' });
  T(s, '3~5세 · 유치원과 어린이집 공통', { x: 4.5, y: 2.2, w: 4.8, h: 0.4, fontSize: 13 });
  const sum = [
    ['교육과정', '학습자의 전인적 성장과 발달을 돕기 위한 교수-학습 경험의 설계도'],
    ['누리과정', '유치원교육과정과 표준보육과정 일원화, 생애초기 출발점 평등 보장'],
    ['2019 개정', "'유아 중심·놀이 중심 교육과정'으로 개정"],
  ];
  sum.forEach((r, k) => {
    const y = 3.1 + k * 0.62;
    s.addShape('ellipse', { x: 0.5, y: y + 0.05, w: 0.4, h: 0.4, fill: { color: C.butter } });
    T(s, String(k + 1), { x: 0.5, y: y + 0.05, w: 0.4, h: 0.4, fontSize: 13, bold: true, align: 'center', valign: 'middle' });
    T(s, r[0], { x: 1.1, y, w: 1.6, h: 0.5, fontSize: 15, bold: true, color: C.white, valign: 'middle' });
    T(s, r[1], { x: 2.7, y, w: 6.8, h: 0.5, fontSize: 14, color: C.white, valign: 'middle' });
  });
  s.addNotes(NOTES[5].script);

}

module.exports = { build };

if (require.main === module) (async () => {
  const pres = new pptxgen();
  pres.layout = 'LAYOUT_16x9';
  pres.title = '영유아교육과정의 이해';
  await build(pres);
  await pres.writeFile({ fileName: '영유아교육과정의_이해.pptx' });
})();
