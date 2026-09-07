import openpyxl
from copy import copy
SRC='/root/.claude/uploads/ea1a909c-9825-524c-af08-aa69686f7bb5/8bba4283-__________.xlsx'
OUT='/home/user/-_oy/reports/OYPB_브랜드별_시딩_비용_광고성과_260907 HAN (최종).xlsx'
wb=openpyxl.load_workbook(SRC); ws=wb['Sheet1']
DEC='_-* #,##0.0_-;\\-* #,##0.0_-;_-* "-"_-;_-@_-'
tpl ={c:{k:copy(getattr(ws.cell(8,c),k)) for k in ('font','fill','border','alignment')} for c in range(2,15)}
tpls={c:{k:copy(getattr(ws.cell(19,c),k)) for k in ('font','fill','border','alignment')} for c in range(2,15)}
nf ={c:ws.cell(8,c).number_format for c in range(2,15)}
nfs={c:ws.cell(19,c).number_format for c in range(2,15)}

# 브랜드 | KR | 글로벌 | 조회수 | 참여수(보정) | 세금계산서 안분 | 광고비 | 매출 | 설명
BD=[
 ('아이디얼포맨',737,0,17330588,231083,325861024,33768439,102736712,
  '25년 4Q UGC 협업 첫 시작을 바탕으로 \n250%대 ROAS는 300% 이상으로 올라 섬'),
 ('브링그린',313,0,15535558,59372,115718190,124304394,729266288,
  '전체 퍼포먼스 광고비 44% 집행 · 전환 매출 49% 창출\nUGC&협력광고 중심 티트리 라인 스케일업 필요'),
 ('라운드어라운드',281,0,2413068,47169,107427671,133848,348140,'PR키트 진행 등 참여율 1위(1.95%)\n광고는 281건 중 1건만 집행'),
 ('바이오힐보 (국내)',195,0,5440064,29506,66149614,56333229,172219558,
  '퍼포먼스 광고 성과는 전량 국내에서 발생\n최대 조회수 764,503뷰'),
 ('바이오힐보 (해외·JP)',0,7,44255,156,2907729,0,0,
  '26.06 NAD 크림 나노 시딩 테스트\n조회수·참여 모두 저조'),
 ('바이오힐보 (해외·US)',0,68,4739117,24134,30000000,0,0,
  '68건 중 36건 성과 확보 · 건당 131,642뷰(국내의 4.7배)\n최대 2,954,129뷰 = 전체 1위'),
 ('식물나라',140,0,2788821,25032,56707149,14985030,118255596,
  '빅&스몰웨이브의 시작(25.05)\n마이크로 IMC 성과 가장 잘 워킹하는 브랜드'),
 ('웨이크메이크',70,60,3143931,19540,67815249,16409851,113788050,
  '26.04 행사 및 올영픽 등 미존재로\nUGC 협업 수량 증대해도 좋을 것으로 판단'),
 ('루테카',124,0,1655285,14046,39270879,0,0,'퍼포먼스 광고 미집행\n(브랜드 fade out)'),
 ('컬러그램',71,32,2653321,14526,67542290,23514230,134432958,
  'IMD 26.09.07로 103건 확정(v8 77건 → +26)\n조회수 +68% · 광고비 3.1배 · 매출 2.2배 · ROAS 572% 5위'),
 ('케어플러스',32,0,188488,1839,6457878,0,0,
  '26.07 UGC 첫 협업 · 퍼포먼스 광고 미집행\n원고료 마이크로 5건×20만 + 나노 27건×10만'),
 ('필리밀리',27,0,1353401,2409,10995846,14086284,121998372,'전체 브랜드 중 ROAS 1위(866%)\n최저 CPA 1,121원'),
 ('올더베러',8,0,396952,3440,3490745,0,0,'퍼포먼스 광고 미집행'),
]
r0=8
def style(r,c,val,sub=False):
    cell=ws.cell(r,c); cell.value=val
    src=(tpls if sub else tpl); fmt=(nfs if sub else nf)
    base = src[c] if c<=14 else src[14]
    for k in ('font','fill','border','alignment'): setattr(cell,k,copy(base[k]))
    cell.number_format = DEC if c==9 else (fmt[c] if c<=14 else fmt[6])
for i,(nm,kr,gl,vw,eng,inv,ad,rev,memo) in enumerate(BD):
    r=r0+i
    vals={2:nm,3:kr,4:gl,5:f'=SUM(C{r}:D{r})',6:vw,7:f'=F{r}/E{r}',8:eng,
          9:f'=O{r}/F{r}',10:f'=ROUND(O{r}/H{r},0)',11:ad,12:rev,
          13:(f'=L{r}/K{r}' if ad else '-'),14:memo,15:inv}
    for c in range(2,16): style(r,c,vals[c])
    ws.row_dimensions[r].height = 39.6 if '\n' in memo else 26.4
r=r0+len(BD)
vals={2:'소계',3:f'=SUM(C{r0}:C{r-1})',4:f'=SUM(D{r0}:D{r-1})',5:f'=SUM(E{r0}:E{r-1})',
      6:f'=SUM(F{r0}:F{r-1})',7:f'=F{r}/E{r}',8:f'=SUM(H{r0}:H{r-1})',
      9:f'=O{r}/F{r}',10:f'=ROUND(O{r}/H{r},0)',11:f'=SUM(K{r0}:K{r-1})',
      12:f'=SUM(L{r0}:L{r-1})',13:f'=L{r}/K{r}',14:'-',15:f'=SUM(O{r0}:O{r-1})'}
for c in range(2,16): style(r,c,vals[c],sub=True)
ws.row_dimensions[r].height=26.4
sub=r
# 헤더
ws['O6']='비용'; ws['O7']='세금계산서'
for c,src in ((15,6),(15,7)):
    d=ws.cell(src,15); t=ws.cell(src,13)
    for k in ('font','fill','border','alignment'): setattr(d,k,copy(getattr(t,k)))
ws.column_dimensions['O'].width=15
ws['I7']='평균 CPV'; ws['J7']='평균 CPE'
ws['B2']='OYPB 브랜드별 시딩·비용·광고성과'
nfont=copy(ws['B3'].font)
ws['B3']='※ 통계 : 25.04 ~ 26.09 · 국내외 전체 2,165건(글로벌 167건 = JP 99 + US 68) · 26.09.07 기준'
ws['B4']='※ RD : 표준 기준 ⑤ — 수량·비용 전체 2,165건 / 조회수·참여 성과 집계분 2,112건 · 모든 단가는 세금계산서 발행 비용 기준'
for k,t in enumerate([
 '※ 비용 = 세금계산서 발행 확정 900,344,264원 (26.02~08 확정 840,514,264 + 25년 발행 59,830,000 · 바이오힐보 US 30,000,000 포함). + JP 미발행 8,904,000 → 진행 기준 909,248,264원',
 '※ 브랜드별 세금계산서는 발행 총액을 원고료 비율로 안분한 값이다(바이오힐보 US는 실제 청구 30,000,000원 직접 계상). 발행 내역상 브랜드 귀속 미확정 조정액이 있어 안분 처리했다.',
 '※ 평균 CPV = 세금계산서 ÷ 조회수 / 평균 CPE = 세금계산서 ÷ 참여수 / 참여수 = 좋아요 + 댓글 → 소계 CPV 15.61원 · CPE 1,907원 · 건당 청구액 415,863원',
 '※ 참여수는 좋아요 데이터 공란 519건(IMD 2,076건의 25.0%)을 브랜드·국가별 좋아요율과 좋아요/댓글비의 중앙값으로 추정해 반영한 값이다 — 좋아요 실측 359,877 + 추정 80,292 = 440,169, 댓글 32,083(전량 실측)',
 '※ 미집계 53건(바이오힐보 US 32 + 웨이크메이크JP 26.09 21)은 조회수·참여 데이터가 없어 성과 분모에서 제외했다. 수량·비용에는 포함되어 있다.',
 '※ (참고) 크리에이터 원고료 498,658,205원 = 청구액의 55.4%. 마크업·VAT·솔루션·지급대행 제외 금액으로 위 단가 산정에는 사용하지 않았다.',
 '※ 컬러그램은 IMD 26.09.07 최신본으로 103건 확정(v8 77건은 부분집합, URL 대조 확인). 다른 브랜드도 동일 누락 가능성이 있어 최신 IMD 확인이 필요하다.',
]):
    c=ws.cell(sub+1+k,2); c.value=t; c.font=copy(nfont)
wb.calculation.fullCalcOnLoad=True
wb.save(OUT); print('saved', OUT, '| 소계행', sub)
