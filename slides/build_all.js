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
  "먼저 첫 번째 부분, '영유아교육과정의 이해'입니다.\n\n유치원과 어린이집에서 아이들에게 '무엇을, 어떻게 가르칠지' 정해 놓은 계획, 즉 교육과정이 무슨 뜻인지 알아보고, 어떻게 바뀌어 왔는지, 요즘 누리과정은 어떤 방향인지 살펴보겠습니다.",
  "두 번째 부분은 '영유아교육과정의 구성'입니다.\n\n교육과정이 무엇으로 이루어져 있는지, 그리고 영유아를 위한 교육과정은 어떤 특징을 가져야 하는지 말씀드리겠습니다.",
  "마지막 세 번째 부분은 '영유아교육과정의 계획 및 운영'입니다.\n\n어린이집과 유치원이 교육과정을 계획하고 운영할 때 무엇을 기준으로 삼는지, 그리고 나라에서 정한 표준보육과정에는 어떤 내용이 담겨 있는지 살펴보겠습니다.",
];
const bridges = [
  '\n\n다음으로, 이런 교육과정이 무엇으로 이루어져 있는지 살펴보겠습니다.',
  '\n\n그럼 마지막으로, 실제로 교육과정을 어떻게 계획하고 운영하는지 살펴보겠습니다.',
  '\n\n오늘 발표를 한 줄로 정리하면, 교육과정은 아이가 골고루 자라도록 돕는 설계도이고, 놀이를 중심으로 구성하며, 나라와 지역의 기준을 바탕으로 우리 기관에 맞게 운영한다는 것입니다.\n\n이상으로 발표를 마치겠습니다. 들어 주셔서 감사합니다.',
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
  all.push({ title: '목차', script: "오늘 발표 순서입니다.\n\n첫째, 교육과정이 무엇이고 어떻게 바뀌어 왔는지 알아보는 '영유아교육과정의 이해'.\n둘째, 교육과정이 무엇으로 이루어져 있고 어떤 특징을 가져야 하는지 보는 '영유아교육과정의 구성'.\n셋째, 실제로 어떻게 계획하고 운영하는지 보는 '영유아교육과정의 계획 및 운영'입니다." });
  s.addNotes(all[1].script);

  for (let i = 0; i < 3; i++) {
    await parts[i].mod.build(pres, { part: i + 1, notes: parts[i].notes.map(n => n) });
    all.push(...parts[i].notes);
  }
  if (pres.slides.length !== all.length) throw new Error('슬라이드 수와 대본 수가 다름');
  fs.writeFileSync('notes_all.json', JSON.stringify(all, null, 2) + '\n');
  await pres.writeFile({ fileName: '교육과정의_계획과_운영_합본.pptx' });
})();
