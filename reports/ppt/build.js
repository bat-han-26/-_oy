const pptxgen = require('pptxgenjs');
const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';           // 13.333 x 7.5
pres.author = 'BAT'; pres.title = '올리브영 PB UGC 시딩 성과';

const BG='000000', CARD='1E1E1E', CARD3='3F3F3F', GREEN='A3E635', CYAN='22D3EE',
      WHITE='FFFFFF', GRAY='9E9E9E', DIM='7A7A7A', BAR='14532D', TAG='F5C518', KO='맑은 고딕';

function base(slide, page, tag){
  slide.background = { color: BG };
  slide.addShape(pres.ShapeType.rect, { x:0, y:0, w:9.0, h:0.34, fill:{color:BAR} });
  slide.addText('파트', { x:0.18, y:0, w:1.2, h:0.34, isTextBox:true, margin:0,
    fontFace:KO, fontSize:12, color:WHITE, valign:'middle' });
  slide.addShape(pres.ShapeType.rect, { x:9.0, y:0, w:4.333, h:0.34, fill:{color:TAG} });
  slide.addText(tag, { x:9.0, y:0, w:4.333, h:0.34, isTextBox:true, margin:0,
    fontFace:KO, fontSize:11, bold:true, color:'1A1A1A', align:'center', valign:'middle' });
  slide.addText('© 2026 BAT', { x:0.35, y:6.95, w:2, h:0.3, isTextBox:true, margin:0,
    fontFace:KO, fontSize:10, color:GRAY, valign:'middle' });
  slide.addText(String(page), { x:12.4, y:6.95, w:0.6, h:0.3, isTextBox:true, margin:0,
    fontFace:KO, fontSize:10, color:GRAY, align:'right', valign:'middle' });
}
function heading(slide, title, sub){
  slide.addText(title, { x:0.6, y:0.48, w:12.133, h:1.02, isTextBox:true, margin:0,
    fontFace:KO, fontSize:26, color:WHITE, align:'center', valign:'middle', lineSpacingMultiple:1.1 });
  slide.addText(sub, { x:0.6, y:1.56, w:12.133, h:0.42, isTextBox:true, margin:0,
    fontFace:KO, fontSize:13.5, color:'C8C8C8', align:'center', valign:'middle' });
}
// KPI 카드
function kpi(slide, x, w, o, accent){
  const y=2.15, h=4.5;
  slide.addShape(pres.ShapeType.roundRect, { x, y, w, h, fill:{color:CARD}, rectRadius:0.06,
    line:{color:'2E2E2E', width:0.75} });
  slide.addText(o.label, { x:x+0.08, y:y+0.22, w:w-0.16, h:0.3, isTextBox:true, margin:0,
    fontFace:KO, fontSize:12.5, bold:true, color:WHITE, align:'center', valign:'middle', underline:{style:'sng'} });
  slide.addText(o.value, { x:x+0.05, y:y+0.85, w:w-0.10, h:0.95, isTextBox:true, margin:0,
    fontFace:KO, fontSize:o.vsize||33, bold:true, color:accent, align:'center', valign:'middle' });
  slide.addText(o.main, { x:x+0.12, y:y+2.02, w:w-0.24, h:0.75, isTextBox:true, margin:0,
    fontFace:KO, fontSize:12.5, bold:true, color:WHITE, align:'center', valign:'top', lineSpacingMultiple:1.15 });
  slide.addText(o.note, { x:x+0.14, y:y+3.05, w:w-0.28, h:1.25, isTextBox:true, margin:0,
    fontFace:KO, fontSize:9, color:DIM, align:'center', valign:'top', lineSpacingMultiple:1.25 });
}
// 분포 카드
function dist(slide, x, w, title, rows){
  const y=2.15, h=4.5;
  slide.addShape(pres.ShapeType.roundRect, { x, y, w, h, fill:{color:CARD3}, rectRadius:0.10,
    line:{color:'5A5A5A', width:0.75} });
  slide.addText(title, { x:x+0.08, y:y+0.28, w:w-0.16, h:0.36, isTextBox:true, margin:0,
    fontFace:KO, fontSize:15, bold:true, color:CYAN, align:'center', valign:'middle' });
  let ry = y+0.95;
  rows.forEach((r,i)=>{
    const last = i===rows.length-1;
    slide.addText(r[0], { x:x+0.20, y:ry, w:(w-0.40)*0.58, h:0.42, isTextBox:true, margin:0,
      fontFace:KO, fontSize:11.5, bold:last, color: last?WHITE:'E4E4E4', align:'center', valign:'middle' });
    slide.addText(r[1], { x:x+0.20+(w-0.40)*0.58, y:ry, w:(w-0.40)*0.42, h:0.42, isTextBox:true, margin:0,
      fontFace:KO, fontSize:11.5, bold:last, color: last?CYAN:'E4E4E4', align:'center', valign:'middle' });
    ry += last ? 0.42 : 0.615;
  });
}

const TAGTXT = 'IMD 26.09.04 기준 · 전체 2,139건';

/* ───────── Slide 1 ───────── */
let s = pres.addSlide(); base(s, 53, TAGTXT);
heading(s, '시딩 자체 성과만으로도 유의미한 도달을 확보하고 있습니다.',
        '원고료 딜메이킹 기반의 효율적 단가와 히어로 콘텐츠 발생률 확인 가능.');
const G=0.20, W1=(12.133-G*4)/5;
[
 { label:'평균 원고료', value:'24.7만원', main:'원고료 딜메이킹 전략',
   note:'※ 참고 : 총 원고료 5.28억원\n전체 2,139건 기준\nUGC 협업 최다 아이디얼포맨 737건' },
 { label:'1만뷰 이상', value:'34.7%', main:'743건 / 2,139건',
   note:'※ 참고 : 조회수 분포\n1) 50만뷰 이상 : 6건\n2) 10만뷰 이상 : 167건\n3) 5만뷰 이상 : 255건' },
 { label:'최대 조회수', value:'84.6만뷰', main:'브링그린 (26.02.10)',
   note:'※ 참고 : 50만뷰 이상 6건\n2위 바이오힐보 76.5만뷰\n(26.05.19)' },
 { label:'최저 CPV', value:'0.33원', main:'아이디얼포맨 (26.01.30)',
   note:'※ 참고 : 단일 콘텐츠 기준\n브랜드 평균 CPV 최저\n브링그린 4.27원' },
 { label:'평균 CPV', value:'10.2원', main:'OYPB 12개 브랜드',
   note:'※ 참고 : 건당 평균 조회수 최고\n필리밀리 50,126뷰\n(성과보유 2,050건 기준 9.42원)' },
].forEach((o,i)=> kpi(s, 0.6+i*(W1+G), W1, o, GREEN));
s.addNotes('시딩 단독 성과. 평균 원고료 24.7만원(총 5.28억÷2,139건), 1만뷰 이상 743건(34.7%), 최대 조회수 84.6만뷰(브링그린 26.02.10), 단일 최저 CPV 0.33원(아이디얼포맨 26.01.30), 평균 CPV 10.2원.');

/* ───────── Slide 2 ───────── */
s = pres.addSlide(); base(s, 54, TAGTXT);
heading(s, '파트너십 광고와 연결했을 때는 보다 극대화된 시너지를 창출하고 있습니다.',
        '참여 지수를 포함한 ROAS, CPA 등 우수한 세일즈 스코어 확인 가능.');
const W2=(12.133-0.30*3)/4;
[
 { label:'평균 CPE', value:'1,441원', main:'OYPB 12개 브랜드\n원고료 ÷ 총 참여 366,236',
   note:'※ 참고 : 브랜드 최고 참여율\n라운드어라운드 1.73%\nRFP 최고 아이디얼포맨 1.08%' },
 { label:'평균 ROAS', value:'530%', main:'OYPB 12개 브랜드\n총 광고비 2.68억원',
   note:'※ 참고 : 브랜드 최고 ROAS\n필리밀리 866%\n총 매출 14.2억원' },
 { label:'최대 ROAS', value:'1,297%', main:'광고비 100만원 이상 단일 소재\n브링그린 (26.07.30)',
   note:'※ 참고 : 10만원 이상 소진 기준\n최대 ROAS 브링그린 3,681%\n(26.07.28)' },
 { label:'평균 CPA', value:'3,409원', main:'총 구매 78,515건\n광고 집행 331건',
   note:'※ 참고 : 최저 CPA\n필리밀리 1,129원' },
].forEach((o,i)=> kpi(s, 0.6+i*(W2+0.30), W2, o, CYAN));
s.addNotes('파트너십 광고 연결 성과. 평균 CPE 1,441원, 평균 ROAS 530%(광고비 2.68억÷매출 14.2억), 단일 소재 최대 ROAS 1,297%(브링그린 26.07.30, 광고비 100만원 이상 기준), 평균 CPA 3,409원.');

/* ───────── Slide 3 ───────── */
s = pres.addSlide(); base(s, 55, TAGTXT);
heading(s, '최우수 성과를 한 눈에 보면 다음과 같습니다.',
        '전체 2,139건 기준 · 성과 데이터 보유 2,050건 (글로벌 89건 성과 데이터 없음)');
const W3=(12.133-0.30*3)/4;
[
 ['히어로 콘텐츠', [['조회수','건수'],['50만 이상','6건'],['30만 이상','17건'],['10만 이상','167건'],['5만 이상','255건'],['1만 이상','743건'],['총 합','2,139건']]],
 ['ROAS 기준',   [['ROAS','건수'],['1,000% 이상','25건'],['800% 이상','49건'],['600% 이상','86건'],['500% 이상','118건'],['300% 이상','172건'],['총 합','2,139건']]],
 ['매출액 기준',  [['매출액','건수'],['5천만원 이상','2건'],['3천만원 이상','6건'],['1천만원 이상','42건'],['5백만원 이상','82건'],['1백만원 이상','185건'],['총 합','2,139건']]],
 ['글로벌 시딩',  [['국가','건수'],['국내','1,972건'],['일본','99건'],['미국','68건'],['',''],['',''],['총 합','2,139건']]],
].forEach((c,i)=> dist(s, 0.6+i*(W3+0.30), W3, c[0], c[1]));
s.addText('※ ROAS 기준은 광고비 10만원 이상 집행 건(242건)에서만 카운트 · 매출액은 광고 집행 건에서만 발생 · 글로벌 시딩은 일본 99건(고정비 4,070만원) + 미국 68건(3,000만원)',
  { x:0.6, y:6.72, w:11.5, h:0.28, isTextBox:true, margin:0,
    fontFace:KO, fontSize:9, color:DIM, align:'left', valign:'middle' });
s.addNotes('분포 요약. 모수는 전체 2,139건. 히어로: 5만뷰 이상 255건(11.9%)이 전체 조회수의 70.4%. 매출: 상위 2건이 전체 매출의 9.0%. 글로벌: 일본 99 + 미국 68 = 167건.');

pres.writeFile({ fileName: '올리브영_PB_UGC_시딩_성과_3P.pptx' }).then(f=>console.log('saved', f));
