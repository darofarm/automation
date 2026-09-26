const pptxgen = require('pptxgenjs');
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

(async () => {
  const pres = new pptxgen();
  pres.layout = 'LAYOUT_16x9';
  pres.title = '영유아교육과정의 이해';

  // ---------- 1. 표지 ----------
  let s = pres.addSlide();
  s.background = { color: C.sage };
  rings(s, 8.6, 1.3, 1.6, ['7FA88A', 'A9CBB0', 'DCEBDD', 'F5D07A']);
  rings(s, 9.4, 4.7, 0.8, ['F2A27E', 'FCE4D8', 'F2A27E']);
  T(s, '05  교육과정의 계획과 운영', { x: 0.7, y: 1.35, w: 6, h: 0.4, fontSize: 16, color: C.butterL });
  T(s, '영유아교육과정의 이해', { x: 0.7, y: 1.8, w: 7, h: 1.0, fontSize: 44, bold: true, color: C.white });
  T(s, '교육과정의 정의부터 누리과정·표준보육과정까지', { x: 0.7, y: 2.85, w: 7, h: 0.45, fontSize: 18, color: C.white });
  T(s, '과목명  |  학번  |  발표자 이름', { x: 0.7, y: 4.3, w: 6, h: 0.4, fontSize: 14, color: C.sageL });
  s.addNotes("안녕하세요. 오늘 제가 발표할 내용은 '영유아교육과정의 이해'입니다.\n\n쉽게 말하면, 유치원과 어린이집에서 아이들에게 '무엇을, 어떻게 가르칠지' 정해 놓은 계획에 대한 이야기예요.\n\n먼저 교육과정이 무슨 뜻인지 알아보고, 유치원과 어린이집의 교육과정이 어떻게 바뀌어 왔는지 살펴본 뒤, 마지막으로 요즘 누리과정이 어떤 방향인지 말씀드리겠습니다.");

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
  s.addShape('roundRect', { x: 0.5, y: 4.2, w: 9.0, h: 0.95, fill: { color: 'F4F7F4' }, rectRadius: 0.12 });
  T(s, [
    { text: '교육과정 = ', options: { bold: true, color: C.sage } },
    { text: '전인적 성장·발달을 돕는 교수-학습 경험의 체계적 ', options: {} },
    { text: '설계도', options: { bold: true, color: C.sage } },
  ], { x: 0.75, y: 4.3, w: 8.5, h: 0.45, fontSize: 16 });
  T(s, '방인옥 등(1999)', { x: 0.75, y: 4.75, w: 8.5, h: 0.3, fontSize: 11, color: C.mute });
  s.addNotes("교육과정이라는 말, 많이 들어 보셨죠? 간단히 말하면 '무엇을, 어떤 순서로 가르치고 배울지 정리한 것'입니다.\n\n그런데 누구 입장에서 보느냐에 따라 뜻이 조금 달라져요. 선생님 입장에서 보면 '가르칠 내용과 계획'이고, 학생 입장에서 보면 '실제로 배우고 경험한 것'입니다(이귀윤, 1990).\n\n오른쪽 그림을 보시면, 넓게 보면 학생이 경험하는 모든 것이 교육과정이고, 좁게 보면 미리 짜 놓은 과목이나 수업 계획만을 말합니다.\n\n방인옥 등(1999)은 교육과정을 '아이가 몸과 마음 모두 골고루 자라도록 돕기 위해, 가르치고 배우는 경험을 잘 짜 놓은 설계도'라고 했습니다. 집을 지을 때 설계도가 필요하듯이, 교육에도 설계도가 필요하다는 뜻이에요.");

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
  s.addNotes("그럼 유치원 교육과정은 어떻게 바뀌어 왔을까요?\n\n우리나라 유치원 교육과정은 1969년에 처음 만들어졌고, 2007년까지 모두 여섯 번 고쳐졌습니다.\n\n그런데 예전에는 유치원은 유치원 교육과정, 어린이집은 표준보육과정, 이렇게 따로따로였어요. 같은 나이의 아이인데도 다니는 곳에 따라 배우는 내용이 달랐던 거죠.\n\n그래서 둘을 하나로 합친 것이 바로 '누리과정'입니다. 목적은 두 가지예요. 첫째, 초등학교에 들어가기 전 교육의 질을 높이는 것. 둘째, 어디에 다니든 모든 아이가 같은 출발선에서 시작할 수 있게 하는 것입니다.\n\n누리과정은 2013년부터 만 3~5세 아이들에게 적용되었고, 2020년부터는 '아이 중심, 놀이 중심'으로 바뀐 개정 누리과정을 쓰고 있습니다.");

  // ---------- 4. 어린이집 표준보육과정 ----------
  s = pres.addSlide();
  s.background = { color: C.white };
  title(s, '어린이집 표준보육과정', 4);
  T(s, '왜 필요해졌을까?', { x: 0.5, y: 1.2, w: 4.2, h: 0.35, fontSize: 15, bold: true, color: C.sage });
  const bg = [
    { i: 'FaHome', c: C.peach, t: '핵가족화' },
    { i: 'FaBriefcase', c: C.butter, t: '취업모의 지속적 증가' },
    { i: 'FaSeedling', c: C.sage, t: '조기 경험의 중요성 부각' },
  ];
  for (let k = 0; k < bg.length; k++) {
    const y = 1.7 + k * 0.62;
    s.addShape('ellipse', { x: 0.5, y, w: 0.48, h: 0.48, fill: { color: bg[k].c } });
    s.addImage({ data: await icon(bg[k].i, C.white), x: 0.62, y: y + 0.12, w: 0.24, h: 0.24 });
    T(s, bg[k].t, { x: 1.15, y, w: 3.4, h: 0.48, fontSize: 15, valign: 'middle' });
  }
  s.addShape('roundRect', { x: 0.5, y: 3.65, w: 4.1, h: 1.35, fill: { color: C.sageL }, rectRadius: 0.12 });
  T(s, [
    { text: '어린이집 이용의 보편화', options: { bold: true, breakLine: true } },
    { text: '언제, 무엇을 목표로, 어떤 경험을 줄지', options: { breakLine: true } },
    { text: '→ ', options: {} },
    { text: '국가수준의 지침', options: { bold: true, color: C.sage } },
    { text: '이 필요해짐', options: {} },
  ], { x: 0.7, y: 3.75, w: 3.8, h: 1.15, fontSize: 13, valign: 'middle', paraSpaceAfter: 4 });
  // 세로 연표
  const vt = [
    { y: '2004', t: '영유아보육법 제29조 제2항 개정\n표준보육과정 개발·보급 명시' },
    { y: '2007', t: '표준보육과정 공포' },
    { y: '2013', t: '제3차 표준보육과정 개정' },
    { y: '2020', t: '제4차 표준보육과정 고시' },
    { y: '2025', t: '2024 개정 표준보육과정(0~2세) 고시' },
  ];
  s.addShape('line', { x: 5.55, y: 1.35, w: 0, h: 3.55, line: { color: 'C9D6CC', width: 2.5 } });
  vt.forEach((e, k) => {
    const y = 1.2 + k * 0.78;
    s.addShape('ellipse', { x: 5.43, y: y + 0.08, w: 0.24, h: 0.24, fill: { color: k === 4 ? C.peach : C.sage } });
    T(s, e.y, { x: 5.85, y, w: 0.8, h: 0.4, fontSize: 16, bold: true, color: k === 4 ? 'D07A55' : C.sage, valign: 'middle' });
    T(s, e.t, { x: 6.7, y: y + 0.02, w: 2.9, h: 0.7, fontSize: 12, valign: 'top' });
  });
  s.addNotes("이번에는 어린이집의 표준보육과정입니다. 표준보육과정은 어린이집에서 아이들이 배우고 경험해야 할 것을 나라에서 정해 놓은 기준이에요.\n\n왜 필요해졌을까요? 1990년대 후반부터 어린이집이 크게 늘었습니다. 핵가족이 많아지고, 일하는 엄마들도 계속 늘었고, 어릴 때의 경험이 아이 발달에 아주 중요하다는 생각도 커졌기 때문이에요.\n\n이제 어린이집은 일부 가정만 보내는 곳이 아니라, 누구나 다니는 곳이 되었습니다. 그러다 보니 '어린이집마다 제각각 하지 말고, 나라에서 공통 기준을 정해 주자'는 필요가 생긴 거죠.\n\n그래서 2004년 영유아보육법에 '표준보육과정을 만들어 보급해야 한다'는 내용이 들어갔고, 2007년에 처음 발표되었습니다. 이후 2013년에 3차, 2020년에 4차로 바뀌었고, 2025년에는 0~2세 아이들을 위한 2024 개정 표준보육과정이 나왔습니다.");

  // ---------- 5. 2019 개정 누리과정 ----------
  s = pres.addSlide();
  s.background = { color: C.white };
  title(s, '2019 개정 누리과정의 혁신 방향', 5);
  T(s, "출발선 평등 실현을 위한 '유아교육 혁신방안'에 의거 (교육부, 2017)", { x: 1.2, y: 1.0, w: 8.3, h: 0.3, fontSize: 12, color: C.mute });
  const inn = [
    { n: '첫째', i: 'FaChild', c: C.peach, f: C.peachL, from: '초등학교 준비 위주 학습', to: '개별 유아의 다양한 특성 고려' },
    { n: '둘째', i: 'FaPuzzlePiece', c: C.butter, f: C.butterL, from: '교사계획서·학습 위주 교육', to: '자유놀이 권장\n관찰·기록 등 상호작용 강조' },
    { n: '셋째', i: 'FaLeaf', c: C.sage, f: C.sageL, from: '5개 영역, 주제-소주제의\n지나치게 세부적인 구성', to: '세부내용 삭제로\n현장의 자율성 존중' },
  ];
  for (let k = 0; k < 3; k++) {
    const x = 0.5 + k * 3.1, e = inn[k];
    s.addShape('roundRect', { x, y: 1.5, w: 2.8, h: 2.75, fill: { color: e.f }, rectRadius: 0.15 });
    s.addShape('ellipse', { x: x + 0.2, y: 1.7, w: 0.6, h: 0.6, fill: { color: e.c } });
    s.addImage({ data: await icon(e.i, C.white), x: x + 0.35, y: 1.85, w: 0.3, h: 0.3 });
    T(s, e.n, { x: x + 0.95, y: 1.7, w: 1.6, h: 0.6, fontSize: 18, bold: true, valign: 'middle' });
    T(s, e.from, { x: x + 0.2, y: 2.45, w: 2.4, h: 0.55, fontSize: 11, color: C.mute, valign: 'top' });
    T(s, '▼', { x: x + 0.2, y: 3.0, w: 2.4, h: 0.25, fontSize: 10, color: e.c === C.butter ? 'B08A2E' : e.c });
    T(s, e.to, { x: x + 0.2, y: 3.3, w: 2.4, h: 0.85, fontSize: 14, bold: true, valign: 'top' });
  }
  s.addShape('roundRect', { x: 0.5, y: 4.45, w: 9.0, h: 0.7, fill: { color: C.sage }, rectRadius: 0.35 });
  T(s, '유아 중심 · 놀이 중심 교육과정  →  2020년부터 유치원·어린이집 공통 적용', { x: 0.5, y: 4.45, w: 9.0, h: 0.7, fontSize: 17, bold: true, color: C.white, align: 'center', valign: 'middle' });
  s.addNotes("그럼 지금 쓰고 있는 2019 개정 누리과정은 무엇이 달라졌을까요? 모든 아이가 같은 출발선에서 시작하게 하자는 교육부의 계획(교육부, 2017)에 따라 바뀌었고, 크게 세 가지가 달라졌습니다.\n\n첫째, 초등학교 공부를 미리 준비시키는 대신, 아이 한 명 한 명의 특성을 살피도록 했습니다.\n\n둘째, 선생님이 짜 놓은 계획대로 가르치기보다 아이들이 자유롭게 놀게 하고, 선생님은 옆에서 잘 지켜보고 기록하면서 함께하도록 했습니다.\n\n셋째, 예전에는 5개 영역, 주제, 소주제처럼 너무 자세하게 정해져 있었는데, 이것을 줄여서 현장 선생님들이 상황에 맞게 판단할 수 있게 했습니다.\n\n한마디로 '아이 중심, 놀이 중심'이고, 2020년부터 유치원과 어린이집 모두에서 쓰이고 있습니다.");

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
    ['교육과정', '학습자의 전인적 성장·발달을 위한 교수-학습 경험의 설계도'],
    ['누리과정', '이원화된 과정을 일원화해 교육의 질과 출발점 평등을 보장'],
    ['2019 개정', '유아 중심·놀이 중심, 현장의 자율성 존중'],
  ];
  sum.forEach((r, k) => {
    const y = 3.1 + k * 0.62;
    s.addShape('ellipse', { x: 0.5, y: y + 0.05, w: 0.4, h: 0.4, fill: { color: C.butter } });
    T(s, String(k + 1), { x: 0.5, y: y + 0.05, w: 0.4, h: 0.4, fontSize: 13, bold: true, align: 'center', valign: 'middle' });
    T(s, r[0], { x: 1.1, y, w: 1.6, h: 0.5, fontSize: 15, bold: true, color: C.white, valign: 'middle' });
    T(s, r[1], { x: 2.7, y, w: 6.8, h: 0.5, fontSize: 14, color: C.white, valign: 'middle' });
  });
  s.addNotes("마지막으로 정리하겠습니다. 지금 0~2세 아이들은 어린이집 표준보육과정을, 3~5세 아이들은 유치원이든 어린이집이든 똑같이 누리과정을 따릅니다.\n\n세 가지만 기억해 주세요.\n하나, 교육과정은 아이가 골고루 자라도록 돕는 교육의 설계도입니다.\n둘, 누리과정은 따로 있던 유치원과 어린이집의 교육과정을 하나로 합친 것입니다.\n셋, 2019 개정 누리과정의 핵심은 아이 중심, 놀이 중심입니다.\n\n이상으로 발표를 마치겠습니다. 들어 주셔서 감사합니다.");

  await pres.writeFile({ fileName: '영유아교육과정의_이해.pptx' });
})();
