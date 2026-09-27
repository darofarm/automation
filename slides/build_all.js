// 세 발표를 하나로 합친 PPT와 대본 notes_all.json을 만든다
const fs = require('fs');
const pptxgen = require('pptxgenjs');
const { C, icon, T, title, cover, state } = require('./lib');
const parts = [
  { mod: require('./build'), notes: require('./notes.json'), name: '영유아교육과정의 이해', keys: ['교육과정의 정의', '누리과정·표준보육과정의 변천', '2019 개정 누리과정'] },
  { mod: require('./build2'), notes: require('./notes2.json'), name: '영유아교육과정의 구성', keys: ['구성 요소(목표·내용·방법·평가)', '실행 순서', '구성의 네 가지 특징'] },
  { mod: require('./build3'), notes: require('./notes3.json'), name: '영유아교육과정의 계획 및 운영', keys: ['계획·운영의 기준', '표준보육과정 총론과 운영', '관련 법규와 지침'] },
];

// 합본용 대본: 각 PART 표지와 마지막 장의 인사말을 이어지는 말로 바꾼다
const ENDING = /\s+이상으로 발표를 마치겠습니다\.( 들어 주셔서)? 감사합니다\.$/;
const coverNotes = [
  "먼저 첫 번째 부분, '영유아교육과정의 이해'입니다.\n\n교육과정의 정의, 유치원교육과정과 누리과정, 어린이집 표준보육과정, 그리고 2019 개정 누리과정의 혁신 방향 순서로 살펴보겠습니다.",
  "두 번째 부분은 '영유아교육과정의 구성'입니다.\n\n교육과정의 구성 요소와 교육과정 구성의 특징 순서로 살펴보겠습니다.",
  "마지막 세 번째 부분은 '영유아교육과정의 계획 및 운영'입니다.\n\n교육과정 계획·운영의 지침, 표준보육과정의 총론과 운영, 연령별 목표 및 내용, 그리고 관련 법규와 지침의 이해 순서로 살펴보겠습니다.",
];
const bridges = [
  "\n\n다음으로 두 번째 부분, '영유아교육과정의 구성'을 살펴보겠습니다.",
  "\n\n다음으로 세 번째 부분, '영유아교육과정의 계획 및 운영'을 살펴보겠습니다.",
  "\n\n이상으로 5장 '교육과정의 계획과 운영' 발표를 마치겠습니다. 들어 주셔서 감사합니다.",
];
parts.forEach((p, i) => {
  const n = p.notes.map(x => ({ ...x }));
  n[0] = { title: `PART ${i + 1} ${p.name}`, script: coverNotes[i] };
  const last = n[n.length - 1];
  if (!ENDING.test(last.script)) throw new Error(`PART ${i + 1} 마지막 대본의 끝인사를 찾지 못함`);
  last.script = last.script.replace(ENDING, '') + bridges[i];
  p.notes = n;
});

(async () => {
  const pres = new pptxgen();
  pres.layout = 'LAYOUT_16x9';
  pres.title = '교육과정의 계획과 운영';
  state.pres = pres;
  const all = [];

  let s = cover(pres, { title: '교육과정의 계획과 운영', sub: '영유아교육과정의 이해 · 구성 · 계획 및 운영' });
  all.push({ title: '표지', script: "안녕하세요. 오늘은 5장 '교육과정의 계획과 운영'을 세 부분으로 나눠 발표하겠습니다." });
  s.addNotes(all[0].script);

  s = pres.addSlide();
  s.background = { color: C.white };
  title(s, '목차');
  const col = [[C.peach, C.peachL, 'C0643E'], [C.butter, C.butterL, 'B08A2E'], [C.sage, C.sageL, C.sage]];
  const ics = ['FaBookOpen', 'FaPuzzlePiece', 'FaClipboardList'];
  for (let i = 0; i < 3; i++) {
    const x = 0.5 + i * 3.07, [c, f, dark] = col[i];
    s.addShape('roundRect', { x, y: 1.3, w: 2.86, h: 3.8, fill: { color: f }, rectRadius: 0.15 });
    s.addShape('ellipse', { x: x + 0.25, y: 1.55, w: 0.7, h: 0.7, fill: { color: c } });
    s.addImage({ data: await icon(ics[i], C.white), x: x + 0.425, y: 1.725, w: 0.35, h: 0.35 });
    T(s, `PART ${i + 1}`, { x: x + 1.1, y: 1.6, w: 1.6, h: 0.6, fontSize: 16, bold: true, color: dark, valign: 'middle' });
    T(s, parts[i].name, { x: x + 0.25, y: 2.45, w: 2.4, h: 0.85, fontSize: 17, bold: true, valign: 'top' });
    T(s, parts[i].keys.map((k, j) => ({ text: k, options: { bullet: true, breakLine: j < 2 } })),
      { x: x + 0.25, y: 3.4, w: 2.45, h: 1.5, fontSize: 12, color: '4A5750', valign: 'top', paraSpaceAfter: 6 });
  }
  all.push({ title: '목차', script: "오늘 발표 순서입니다.\n\n첫째, '영유아교육과정의 이해'.\n둘째, '영유아교육과정의 구성'.\n셋째, '영유아교육과정의 계획 및 운영'입니다." });
  s.addNotes(all[1].script);

  for (let i = 0; i < 3; i++) {
    await parts[i].mod.build(pres, { part: i + 1, notes: parts[i].notes.map(n => n) });
    all.push(...parts[i].notes);
  }
  if (pres.slides.length !== all.length) throw new Error('슬라이드 수와 대본 수가 다름');
  fs.writeFileSync('notes_all.json', JSON.stringify(all, null, 2) + '\n');
  await pres.writeFile({ fileName: '교육과정의_계획과_운영_합본.pptx' });
})();
