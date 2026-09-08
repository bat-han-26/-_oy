const pptxgen = require('pptxgenjs');
const p = new pptxgen();
p.defineLayout({ name:'W16x9', width:13.333, height:7.5 });
p.layout = 'W16x9';
const BG='000000', CARD='1C1C1C', GREEN='A3E635', CYAN='38BDF8', WHITE='FFFFFF', GRAY='9E9E9E', DGREEN='0F573E';
const F='맑은 고딕';

function shell(title, part, page){
  const s = p.addSlide();
  s.background = { color: BG };
  // 상단 바
  s.addShape(p.ShapeType.rect, { x:0, y:0, w:13.333, h:0.30, fill:{color:DGREEN} });
  s.addText(part, { x:0.10, y:0.02, w:4.0, h:0.26, fontFace:F, fontSize:9, color:WHITE, bold:true, valign:'middle' });
  s.addShape(p.ShapeType.roundRect, { x:12.42, y:0.06, w:0.80, h:0.52, fill:{color:'22C55E'}, rectRadius:0.05 });
  s.addText('완료', { x:12.42, y:0.06, w:0.80, h:0.52, fontFace:F, fontSize:11, color:WHITE, bold:true, align:'center', valign:'middle' });
  s.addText('© 2026 BAT', { x:0.30, y:7.10, w:2.0, h:0.25, fontFace:F, fontSize:8, color:GRAY, valign:'middle' });
  s.addText(String(page), { x:12.60, y:7.10, w:0.50, h:0.25, fontFace:F, fontSize:9, color:GRAY, align:'right', valign:'middle' });
  return s;
}

function card(s, x, y, w, h, accent, label, value, basis, refs){
  s.addShape(p.ShapeType.roundRect, { x, y, w, h, fill:{color:CARD}, rectRadius:0.03 });
  s.addText(label, { x, y:y+0.16, w, h:0.28, fontFace:F, fontSize:11, color:WHITE, bold:true, align:'center', valign:'middle', underline:true });
  s.addText(value, { x, y:y+0.70, w, h:0.90, fontFace:F, fontSize:34, color:accent, bold:true, align:'center', valign:'middle' });
  s.addText(basis, { x:x+0.08, y:y+1.80, w:w-0.16, h:0.62, fontFace:F, fontSize:11, color:WHITE, align:'center', valign:'top', lineSpacingMultiple:1.2 });
  s.addText(refs, { x:x+0.12, y:y+2.62, w:w-0.24, h:0.95, fontFace:F, fontSize:8, color:GRAY, align:'center', valign:'top', lineSpacingMultiple:1.25 });
}

// ══════════════════════════════ 52p
let s = shell('', 'UGC 누적 데이터 기재', 52);
s.addText([
  { text:'BAT가 OYPB와 수행한 25-26년 2,165건의 ', options:{ color:WHITE } },
  { text:'UGC 협업 단가', options:{ color:GREEN } },
  { text:'는,', options:{ color:WHITE } },
], { x:0.6, y:0.48, w:12.13, h:0.40, fontFace:F, fontSize:21, bold:true, align:'center', valign:'middle' });
s.addText([
  { text:'2차 라이선스를 포함해 ', options:{ color:GREEN } },
  { text:'30만원 수준', options:{ color:GREEN } },
  { text:'으로 가성비와 합리성을 동시에 추구합니다.', options:{ color:WHITE } },
], { x:0.6, y:0.88, w:12.13, h:0.40, fontFace:F, fontSize:21, bold:true, align:'center', valign:'middle' });
s.addText('원고료 딜메이킹 전략은 고정 단가로 진행되는 일반적인 인플루언서 협업 단가 보다 효율·효과적입니다.',
  { x:0.6, y:1.38, w:12.13, h:0.30, fontFace:F, fontSize:12, color:'D9D9D9', align:'center', valign:'middle' });

const C5 = [
  [GREEN,'건당 단가','30만원','원고료 딜메이킹 전략',
   '※ 참고 : 정합 표본 1,770건\n실단가 290,531원\nUGC 협업 최다 아이디얼포맨 737건'],
  [GREEN,'10만뷰 이상','8.1%','172건 / 2,112건',
   '※ 참고 : 조회수 분포\n1) 50만뷰 이상 : 9건\n2) 10만뷰 이상 : 172건\n3) 5만뷰 이상 : 271건'],
  [GREEN,'최대 조회수','295.4만뷰','바이오힐보 US (25.08.25)',
   '※ 참고 : 50만뷰 이상 9건\n2위 브링그린 84.6만뷰 (26.02.10)\n3위 바이오힐보 US 79.2만뷰'],
  [GREEN,'최저 CPV','0.33원','아이디얼포맨 (26.01.30)',
   '※ 참고 : 단일 콘텐츠 기준\n조회수 746,820 / 원고료 250,000\n브랜드 평균 CPV 최저 브링그린 4.27원'],
  [GREEN,'평균 CPV','15원','OYPB 11개 브랜드',
   '※ 참고 : 세금계산서 발행 기준 15.61원\n건당 평균 조회수 최고 필리밀리 50,126뷰\n전체 건당 조회수 27,312뷰'],
];
{
  const n=5, gap=0.20, x0=0.60, tot=13.333-x0*2, w=(tot-gap*(n-1))/n;
  C5.forEach((c,i)=> card(s, x0+i*(w+gap), 1.85, w, 3.85, c[0], c[1], c[2], c[3], c[4]));
}

// ══════════════════════════════ 53p
s = shell('', '파트', 53);
s.addText([
  { text:'100% 전량의 2차 라이선스', options:{ color:CYAN } },
  { text:'를 확보한 UGC는', options:{ color:WHITE } },
], { x:0.6, y:0.48, w:12.13, h:0.40, fontFace:F, fontSize:21, bold:true, align:'center', valign:'middle' });
s.addText([
  { text:'편집본·클린본 모두 퍼포먼스 광고의 ', options:{ color:WHITE } },
  { text:"'장작 및 연료'", options:{ color:CYAN } },
  { text:'로 사용되고 있습니다.', options:{ color:WHITE } },
], { x:0.6, y:0.88, w:12.13, h:0.40, fontFace:F, fontSize:21, bold:true, align:'center', valign:'middle' });
s.addText('평균 조회당 비용 15.6원의 UGC는 평균 ROAS 527%의 고효율과 최대 1,297%(광고비 100만원 이상 소진 기준)의 수익율을 마크하고 있습니다.',
  { x:0.9, y:1.32, w:11.53, h:0.42, fontFace:F, fontSize:12, color:'D9D9D9', align:'center', valign:'middle' });

const C4 = [
  [CYAN,'평균 CPE','2,297원','OYPB 11개 브랜드',
   '※ 참고 : 세금계산서 발행 기준\n브랜드 최고 참여율 라운드어라운드 1.73%\nRFP 최고 아이디얼포맨 1.08%'],
  [CYAN,'평균 ROAS','527%','OYPB 11개 브랜드',
   '※ 참고 : 파트너십 광고 단독 539%\n브랜드 최고 ROAS 필리밀리 866%\n총 매출 14.9억원'],
  [CYAN,'최대 ROAS','1,297%','단일 소재 100만원 이상\n파트너십 광고비 소진 기준',
   '※ 참고 : 브링그린 (26.07.30)\n광고비 100만원 이상 78건 중 1위\n10만원 이상 기준 최대 3,681%'],
  [CYAN,'평균 CPA','3,269원','총 구매 86,740건\n광고 집행 331건',
   '※ 참고 : 최저 CPA\n필리밀리 1,129원\n소재 활용 647건 (파트너십 478 · 영상가공 169)'],
];
{
  const n=4, gap=0.24, x0=0.90, tot=13.333-x0*2, w=(tot-gap*(n-1))/n;
  C4.forEach((c,i)=> card(s, x0+i*(w+gap), 1.90, w, 3.80, c[0], c[1], c[2], c[3], c[4]));
}

p.writeFile({ fileName:'/home/user/-_oy/reports/ppt/OYPB_UGC_협업단가_광고성과_2P_HAN_260907.pptx' })
 .then(f=>console.log('saved', f));
