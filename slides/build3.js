const pptxgen = require('pptxgenjs');
const { FONT, C, icon, T, title, rings, cover } = require('./lib');
const DEFAULT_NOTES = require('./notes3.json');

const card = (s, x, y, w, h, fill) => s.addShape('roundRect', { x, y, w, h, fill: { color: fill }, rectRadius: 0.15 });
async function dot(s, x, y, d, fill, ic) {
  s.addShape('ellipse', { x, y, w: d, h: d, fill: { color: fill } });
  s.addImage({ data: await icon(ic, C.white), x: x + d * 0.25, y: y + d * 0.25, w: d * 0.5, h: d * 0.5 });
}
const cell = (text, o = {}) => ({ text, options: { fontFace: FONT, fontSize: 11, color: C.ink, valign: 'middle', align: 'center', margin: 0.05, ...o } });
const head = (text, o = {}) => cell(text, { bold: true, fill: { color: C.sageL }, ...o });

async function build(pres, opts = {}) {
  const NOTES = opts.notes || DEFAULT_NOTES;
  let s = cover(pres, { title: '영유아교육과정의\n계획 및 운영', sub: '표준보육과정과 관련 지침 살펴보기', part: opts.part, titleSize: 40, twoLine: true });
  s.addNotes(NOTES[0].script);

  // ---------- 2. 계획·운영의 기준 ----------
  s = pres.addSlide();
  s.background = { color: C.white };
  title(s, '교육과정 계획·운영의 지침', 2);
  T(s, '전문적 지식 + 다음을 지침으로 운영', { x: 0.5, y: 1.2, w: 4.2, h: 0.35, fontSize: 14, bold: true, color: C.sage });
  const guide = [
    ['FaSchool', C.peach, '각 원이 추구하는 교육 방향'],
    ['FaFlag', C.sage, '국가수준의 교육과정'],
    ['FaMapMarkerAlt', C.butter, '소속 지역의 유아교육 방향'],
    ['FaFileAlt', C.rose, '해당 연도 유치원 교육과정\n편성·운영 지침'],
    ['FaBook', C.sageM, '어린이집 보육사업안내'],
  ];
  for (let i = 0; i < guide.length; i++) {
    const y = 1.7 + i * 0.66;
    await dot(s, 0.5, y, 0.48, guide[i][1], guide[i][0]);
    T(s, guide[i][2], { x: 1.15, y: y - 0.06, w: 3.5, h: 0.6, fontSize: 13, valign: 'middle' });
  }
  T(s, '〈표 5-1〉 어린이집과 유치원의 국가수준 교육과정', { x: 5.0, y: 1.2, w: 4.5, h: 0.35, fontSize: 12, bold: true, color: C.mute });
  s.addTable([
    [head('구분'), head('0~2세'), head('3~5세')],
    [head('어린이집', { fill: { color: 'F4F7F4' } }), cell('표준보육과정\n(0~1세, 2세 보육과정)'), cell('표준보육과정\n(3~5세 = 누리과정)', { fill: { color: C.butterL }, bold: true })],
    [head('유치원', { fill: { color: 'F4F7F4' } }), cell('–', { color: C.mute }), cell('유치원 교육과정\n(누리과정)', { fill: { color: C.butterL }, bold: true })],
  ], { x: 5.0, y: 1.6, w: 4.5, colW: [0.9, 1.8, 1.8], rowH: [0.4, 0.75, 0.75], border: { type: 'solid', pt: 1, color: 'C9D6CC' } });
  card(s, 5.0, 3.75, 4.5, 1.25, C.butterL);
  T(s, [
    { text: '3~5세는 공통교육과정(누리과정)으로 내용이 동일', options: { bold: true, breakLine: true } },
    { text: '→ 교원이 함께 연구하고 각 기관의 교육에 어떻게 반영할지 협의', options: { breakLine: true } },
    { text: '※ 제4차 표준보육과정(보건복지부 고시 제2020-75호) 중심', options: { fontSize: 10.5, color: C.mute } },
  ], { x: 5.2, y: 3.85, w: 4.1, h: 1.05, fontSize: 12, valign: 'middle', paraSpaceAfter: 4 });
  s.addNotes(NOTES[1].script);

  // ---------- 3. 인간상·목적·목표 ----------
  s = pres.addSlide();
  s.background = { color: C.white };
  title(s, '추구하는 인간상과 목표', 3);
  T(s, '추구하는 인간상', { x: 0.5, y: 1.1, w: 4, h: 0.3, fontSize: 13, bold: true, color: C.sage });
  const hum = [['건강한 사람', C.peachL], ['자주적인 사람', C.butterL], ['창의적인 사람', C.sageL], ['감성이 풍부한 사람', C.roseL], ['더불어 사는 사람', 'EDE6F3']];
  hum.forEach(([t, f], i) => {
    const x = 0.5 + i * 1.84;
    s.addShape('roundRect', { x, y: 1.45, w: 1.66, h: 0.45, fill: { color: f }, rectRadius: 0.22 });
    T(s, t, { x, y: 1.45, w: 1.66, h: 0.45, fontSize: 12, bold: true, align: 'center', valign: 'middle' });
  });
  card(s, 0.5, 2.05, 9.0, 0.55, 'F4F7F4');
  T(s, [
    { text: '목적  ', options: { bold: true, color: C.sage } },
    { text: '영유아가 놀이를 통해 심신의 건강과 조화로운 발달을 이루고, 바른 인성과 민주 시민의 기초를 형성', options: {} },
  ], { x: 0.7, y: 2.05, w: 8.6, h: 0.55, fontSize: 12, valign: 'middle' });
  const goals = [
    ['0~1세 보육과정 및 2세 보육과정 목표', C.peachL, 'C0643E', ['자신의 소중함을 알고, 건강하고 안전한 환경에서 즐겁게 생활한다.', '자신의 일을 스스로 하고자 한다.', '호기심을 가지고 탐색하며 상상력을 기른다.', '일상에서 아름다움에 관심을 가지고 감성을 기른다.', '사람과 자연을 존중하고 소통하는 데 관심을 가진다.']],
    ['3~5세 보육과정 목표', C.butterL, 'B08A2E', ['자신의 소중함을 알고, 건강하고 안전한 생활습관을 기른다.', '자신의 일을 스스로 해결하는 기초능력을 기른다.', '호기심과 탐구심을 가지고 상상력과 창의력을 기른다.', '일상에서 아름다움을 느끼고 문화적 감수성을 기른다.', '사람과 자연을 존중하고 배려하며 소통하는 태도를 기른다.']],
  ];
  goals.forEach(([h, f, dark, items], k) => {
    const x = 0.5 + k * 4.6;
    card(s, x, 2.75, 4.4, 2.45, f);
    T(s, h, { x: x + 0.2, y: 2.83, w: 4.0, h: 0.35, fontSize: 12.5, bold: true, color: dark });
    T(s, items.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < items.length - 1 } })),
      { x: x + 0.2, y: 3.2, w: 4.05, h: 1.95, fontSize: 10.5, valign: 'top', paraSpaceAfter: 3 });
  });
  s.addNotes(NOTES[2].script);

  // ---------- 4. 구성의 중점 ----------
  s = pres.addSlide();
  s.background = { color: C.white };
  title(s, '표준보육과정 구성의 중점', 4);
  T(s, '전제: 영유아는 개별적인 특성을 지닌 고유한 존재', { x: 1.2, y: 1.0, w: 8.3, h: 0.3, fontSize: 13, color: C.mute });
  const focus = [
    ['FaUsers', C.peach, C.peachL, '0~5세 모든 영유아에게\n적용할 수 있도록'],
    ['FaStar', C.butter, C.butterL, '인간상 구현을 위한\n지식, 기능, 태도\n및 가치 반영'],
    ['FaShapes', C.sage, C.sageL, '5개 영역 중심'],
    ['FaSeedling', C.rose, C.roseL, '0~5세 영유아가\n경험해야 할 내용'],
    ['FaSchool', C.sageM, 'EEF4EF', '초등학교 교육과정\n과의 연계성 고려'],
  ];
  for (let i = 0; i < 5; i++) {
    const x = 0.5 + i * 1.84, f = focus[i];
    card(s, x, 1.55, 1.66, 2.1, f[2]);
    await dot(s, x + 0.53, 1.8, 0.6, f[1], f[0]);
    T(s, f[3], { x: x + 0.08, y: 2.55, w: 1.5, h: 0.9, fontSize: 11, bold: true, align: 'center', valign: 'top' });
  }
  T(s, '5개 영역', { x: 0.5, y: 3.95, w: 9, h: 0.3, fontSize: 13, bold: true, color: C.sage });
  const dom = ['신체운동·건강', '의사소통', '사회관계', '예술경험', '자연탐구'];
  const domC = [C.peach, C.butter, C.sage, C.rose, C.sageM];
  dom.forEach((d, i) => {
    const x = 0.5 + i * 1.84;
    s.addShape('roundRect', { x, y: 4.35, w: 1.66, h: 0.6, fill: { color: domC[i] }, rectRadius: 0.3 });
    T(s, d, { x, y: 4.35, w: 1.66, h: 0.6, fontSize: 13, bold: true, color: i === 1 ? C.ink : C.white, align: 'center', valign: 'middle' });
  });
  s.addNotes(NOTES[3].script);

  // ---------- 5. 운영 ----------
  s = pres.addSlide();
  s.background = { color: C.white };
  title(s, '표준보육과정의 운영', 5);
  const ops = [
    { h: '① 편성·운영', ic: 'FaCalendarAlt', c: C.peach, f: C.peachL, items: ['각 기관의 운영 시간에 맞추어 편성', '기관의 실정에 적합한 계획 수립', '바깥놀이를 포함하여 놀이가 충분히 이루어지도록', '성, 신체적 특성, 장애, 종교, 가족 및 문화적 배경 등에 따른 차별이 없도록', '발달과 장애 정도에 따라 조정', '가정과 지역사회와의 협력과 참여에 기반', '교사 연수를 통해 운영 개선'] },
    { h: '② 교수·학습', ic: 'FaPuzzlePiece', c: C.sage, f: C.sageL, items: ['흥미와 관심에 따라 놀이에 자유롭게 참여하고 즐기도록', '놀이를 통해 배우도록', '다양한 놀이와 활동을 경험하도록 실내외 환경 구성', '영유아와 영유아·교사·환경 사이 능동적 상호작용', '5개 영역의 내용이 통합적으로 경험과 연계', '개별 요구에 따라 휴식과 일상생활이 원활히', '연령, 발달, 장애, 배경 등을 고려해 개별 특성에 적합한 방식으로'] },
    { h: '③ 평가', ic: 'FaClipboardCheck', c: C.butter, f: C.butterL, items: ['운영의 질을 진단하고 개선하기 위해 평가를 계획하고 실시', '영유아의 특성 및 변화 정도와 표준보육과정의 운영을 평가', '평가의 목적에 따라 적합한 방법 사용', '결과는 영유아에 대한 이해와 운영 개선을 위한 자료로 활용'] },
  ];
  for (let k = 0; k < 3; k++) {
    const x = 0.5 + k * 3.07, o = ops[k];
    card(s, x, 1.2, 2.86, 3.95, o.f);
    await dot(s, x + 0.2, 1.35, 0.5, o.c, o.ic);
    T(s, o.h, { x: x + 0.8, y: 1.35, w: 1.9, h: 0.5, fontSize: 16, bold: true, valign: 'middle' });
    T(s, o.items.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < o.items.length - 1 } })),
      { x: x + 0.15, y: 1.95, w: 2.6, h: 3.15, fontSize: 10.5, valign: 'top', paraSpaceAfter: 4 });
  }
  s.addNotes(NOTES[4].script);

  // ---------- 6. 연령별 목표 및 내용 ----------
  s = pres.addSlide();
  s.background = { color: C.white };
  title(s, '연령별 목표 및 내용: 5개 영역과 내용범주', 6);
  T(s, '0~1세 · 2세 · 3~5세 모두 같은 5개 영역 (교육부고시 제2024-23호)', { x: 1.2, y: 1.0, w: 8.3, h: 0.3, fontSize: 12, color: C.mute });
  const cats = [
    ['신체운동·건강', ['신체활동 즐기기', '건강하게 생활하기', '안전하게 생활하기'], []],
    ['의사소통', ['듣기와 말하기', '읽기와 쓰기에\n관심 가지기', '책과 이야기 즐기기'], []],
    ['사회관계', ['나를 알고 존중하기', '더불어 생활하기'], ['사회에 관심 가지기']],
    ['예술경험', ['아름다움 찾아보기', '창의적으로 표현하기'], ['예술 감상하기']],
    ['자연탐구', ['탐구 과정 즐기기', '생활 속에서 탐구하기', '자연과 더불어 살기'], []],
  ];
  cats.forEach((c, i) => {
    const x = 0.5 + i * 1.84;
    s.addShape('roundRect', { x, y: 1.45, w: 1.66, h: 0.5, fill: { color: domC[i] }, rectRadius: 0.12 });
    T(s, c[0], { x, y: 1.45, w: 1.66, h: 0.5, fontSize: 13, bold: true, color: i === 1 ? C.ink : C.white, align: 'center', valign: 'middle' });
    const all = c[1].map(t => [t, false]).concat(c[2].map(t => [t, true]));
    all.forEach(([t, only35], j) => {
      const y = 2.05 + j * 0.56;
      s.addShape('roundRect', { x, y, w: 1.66, h: 0.48, fill: { color: only35 ? C.white : 'F4F7F4' }, line: only35 ? { color: C.peach, width: 1.25, dashType: 'dash' } : undefined, rectRadius: 0.1 });
      T(s, t, { x: x + 0.05, y, w: 1.56, h: 0.48, fontSize: 10.5, align: 'center', valign: 'middle', color: only35 ? 'C0643E' : C.ink });
    });
  });
  T(s, '점선 = 3~5세에 추가되는 내용범주', { x: 0.5, y: 3.75, w: 9.0, h: 0.25, fontSize: 10, color: 'C0643E', align: 'right' });
  // 연령별 심화 예시
  card(s, 0.5, 4.05, 9.0, 1.15, C.butterL);
  T(s, "예) '안전하게 생활하기'", { x: 0.7, y: 4.12, w: 3, h: 0.3, fontSize: 12, bold: true, color: 'B08A2E' });
  const ex = [['0~1세', '위험하다는 말에 주의한다'], ['2세', '위험한 상황에 대처하는\n방법을 경험한다'], ['3~5세', '안전사고·화재·재난·학대·유괴 등에\n대처하는 방법을 경험한다']];
  ex.forEach((e, i) => {
    const x = 0.7 + i * 2.95;
    T(s, e[0], { x, y: 4.42, w: 2.5, h: 0.26, fontSize: 12, bold: true, color: C.sage });
    T(s, e[1], { x, y: 4.7, w: 2.55, h: 0.45, fontSize: 10.5, valign: 'top' });
    if (i < 2) T(s, '→', { x: x + 2.55, y: 4.5, w: 0.35, h: 0.5, fontSize: 16, bold: true, color: 'B08A2E', align: 'center', valign: 'middle' });
  });
  s.addNotes(NOTES[5].script);

  // ---------- 7. 관련 법규와 지침 ----------
  s = pres.addSlide();
  s.background = { color: C.white };
  title(s, '관련 법규와 지침의 이해', 7);
  T(s, [
    { text: '사전에 교육청의 지침 파악, 필수 교육 내용은 교육과정에 반영', options: { bullet: true, breakLine: true } },
    { text: '교육청이 정하는 내용은 교육청마다 다를 수 있음 → 소속 교육청의 계획서와 지침 참조', options: { bullet: true, breakLine: true } },
    { text: '운영계획서 작성, 연간 학사일정 결정 시에도 소속청의 지침 숙지·반영', options: { bullet: true } },
  ], { x: 0.5, y: 1.2, w: 4.5, h: 1.3, fontSize: 11.5, paraSpaceAfter: 4, valign: 'top' });
  card(s, 0.5, 2.65, 2.15, 1.55, C.peachL);
  T(s, '유치원', { x: 0.65, y: 2.75, w: 1.9, h: 0.35, fontSize: 14, bold: true, color: 'C0643E' });
  T(s, '교육청별 유치원 교육과정 편성·운영 지침\n유아교육계획', { x: 0.65, y: 3.1, w: 1.9, h: 1.0, fontSize: 11, valign: 'top' });
  card(s, 2.8, 2.65, 2.15, 1.55, C.sageL);
  T(s, '어린이집', { x: 2.95, y: 2.75, w: 1.9, h: 0.35, fontSize: 14, bold: true, color: C.sage });
  T(s, '보육사업안내서\n업무별 매뉴얼', { x: 2.95, y: 3.1, w: 1.9, h: 1.0, fontSize: 11, valign: 'top' });
  card(s, 5.2, 1.2, 4.3, 3.0, 'F4F7F4');
  T(s, '〈표 5-5〉 2025학년도 경기도 유치원\n교육과정 편성·운영 지침 구성', { x: 5.4, y: 1.3, w: 3.9, h: 0.6, fontSize: 12, bold: true, color: C.mute });
  const parts = ['경기교육의 방향', '경기유아교육의 방향', '교육과정 편성·운영', '영역별 목표 및 내용', '교육과정 편성·운영 지원'];
  parts.forEach((p, i) => {
    const y = 2.0 + i * 0.42;
    s.addShape('roundRect', { x: 5.4, y, w: 0.75, h: 0.32, fill: { color: C.sage }, rectRadius: 0.16 });
    T(s, `제${i + 1}부`, { x: 5.4, y, w: 0.75, h: 0.32, fontSize: 10.5, bold: true, color: C.white, align: 'center', valign: 'middle' });
    T(s, p, { x: 6.3, y, w: 3.1, h: 0.32, fontSize: 12, valign: 'middle' });
  });
  s.addShape('roundRect', { x: 0.5, y: 4.45, w: 9.0, h: 0.7, fill: { color: C.sage }, rectRadius: 0.35 });
  T(s, '반드시 소속 교육청의 계획서와 지침을 참조', { x: 0.5, y: 4.45, w: 9.0, h: 0.7, fontSize: 16, bold: true, color: C.white, align: 'center', valign: 'middle' });
  s.addNotes(NOTES[6].script);

}

module.exports = { build };

if (require.main === module) (async () => {
  const pres = new pptxgen();
  pres.layout = 'LAYOUT_16x9';
  pres.title = '영유아교육과정의 계획 및 운영';
  await build(pres);
  await pres.writeFile({ fileName: '영유아교육과정의_계획및운영.pptx' });
})();
