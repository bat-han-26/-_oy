import openpyxl
from copy import copy
SRC='/root/.claude/uploads/ea1a909c-9825-524c-af08-aa69686f7bb5/8bba4283-__________.xlsx'
OUT='/home/user/-_oy/reports/OYPB_브랜드별_시딩_비용_광고성과_260907 HAN.xlsx'
wb=openpyxl.load_workbook(SRC); ws=wb['Sheet1']

DEC='_-* #,##0.0_-;\\-* #,##0.0_-;_-* "-"_-;_-@_-'
# 스타일 템플릿 (원본 8행=일반, 19행=소계)
tpl={c:{k:copy(getattr(ws.cell(8,c),k)) for k in ('font','fill','border','alignment','number_format')}
     for c in range(2,15)}
tpl_s={c:{k:copy(getattr(ws.cell(19,c),k)) for k in ('font','fill','border','alignment','number_format')}
       for c in range(2,15)}
for d in (tpl,tpl_s):
    for c in d: d[c]['number_format']=ws.cell(8 if d is tpl else 19,c).number_format

# 브랜드 · KR · 글로벌 · 조회수 · 참여수 · 원고료(또는 청구) · 광고비 · 매출 · 설명
BD=[
 ('아이디얼포맨',737,0,17330588,186345,186700000,33768439,102736712,
  '25년 4Q UGC 협업 첫 시작을 바탕으로 \n250%대 ROAS는 300% 이상으로 올라 섬'),
 ('브링그린',313,0,15535558,49592,66300000,124304394,729266288,
  '11개 브랜드 전체 퍼포먼스 광고비 중 44% 집행 → UGC&협력광고 중심 티트리 라인 스케일업 필요'),
 ('라운드어라운드',281,0,2413068,41731,61550000,133848,348140,'PR키트 진행 등 참여율 1위'),
 ('바이오힐보 (국내)',195,0,5440064,24922,37900000,56333229,172219558,
  '퍼포먼스 광고 성과는 전량 국내에서 발생\n최대 조회수 764,503뷰'),
 ('바이오힐보 (해외·JP)',0,7,44255,29,1665965,0,0,
  '25.8~9월 빅빙세일 나노 시딩 테스트\n조회수·참여 모두 저조'),
 ('바이오힐보 (해외·US)',0,68,4739117,24134,30000000,0,0,
  '68건 중 36건 성과 확보 · 건당 131,642뷰(국내의 4.7배)\n최대 2,954,129뷰 = 전체 1위 / CPV는 청구 기준'),
 ('식물나라',140,0,2788821,20319,32490000,14985030,118255596,
  '빅&스몰웨이브의 시작(25.05)\n마이크로 IMC 성과 가장 잘 워킹하는 브랜드'),
 ('웨이크메이크',70,60,3143931,18097,38854315,16409851,113788050,
  '26.04 행사 및 올영픽 등 미존재로\nUGC 협업 수량 증대해도 좋을 것으로 판단'),
 ('루테카',124,0,1655285,9444,22500000,0,0,'퍼포먼스 광고 미집행\n(브랜드 fade out)'),
 ('컬러그램',45,32,1583237,8511,32447925,7670607,61405809,
  '전체 브랜드 중 ROAS 2위(800% 이상)\n→ 협업 수량 증가 필요'),
 ('케어플러스',32,0,188488,1839,3700000,0,0,
  '26.07 UGC 첫 협업 · 퍼포먼스 광고 미집행\n원고료 마이크로 5건×20만 + 나노 27건×10만'),
 ('필리밀리',27,0,1353401,1967,6300000,14086284,121998372,'전체 브랜드 중 ROAS 1위\n최저 CPA 확보'),
 ('올더베러',8,0,396952,3440,2000000,0,0,'퍼포먼스 광고 미집행'),
]
r0=8
for i,(nm,kr,gl,vw,eng,fee,ad,rev,memo) in enumerate(BD):
    r=r0+i
    vals={2:nm, 3:kr, 4:gl, 5:f'=SUM(C{r}:D{r})', 6:vw, 7:f'=F{r}/E{r}', 8:eng,
          9:round(fee/vw,2), 10:round(fee/eng), 11:ad, 12:rev,
          13:(f'=L{r}/K{r}' if ad else '-'), 14:memo}
    for c in range(2,15):
        cell=ws.cell(r,c)
        cell.value=vals[c]
        for k in ('font','fill','border','alignment'): setattr(cell,k,copy(tpl[c][k]))
        cell.number_format = DEC if c==9 else tpl[c]['number_format']
    ws.row_dimensions[r].height = 39.6 if '\n' in memo else 26.4

# 소계
r=r0+len(BD)
TOTFEE=492408205; TOTVW=56612765; TOTENG=390370   # 표준 기준 ⑤ (성과 집계분 2,086건)
vals={2:'소계', 3:f'=SUM(C{r0}:C{r-1})', 4:f'=SUM(D{r0}:D{r-1})', 5:f'=SUM(E{r0}:E{r-1})',
      6:f'=SUM(F{r0}:F{r-1})', 7:f'=F{r}/E{r}', 8:f'=SUM(H{r0}:H{r-1})',
      9:round(TOTFEE/TOTVW,2), 10:round(TOTFEE/TOTENG),
      11:f'=SUM(K{r0}:K{r-1})', 12:f'=SUM(L{r0}:L{r-1})', 13:f'=L{r}/K{r}', 14:'-'}
for c in range(2,15):
    cell=ws.cell(r,c); cell.value=vals[c]
    for k in ('font','fill','border','alignment'): setattr(cell,k,copy(tpl_s[c][k]))
    cell.number_format = DEC if c==9 else tpl_s[c]['number_format']
ws.row_dimensions[r].height=26.4
sub=r

# 머리말 · 각주
nf=copy(ws['B3'].font)
ws['B3']='※ 통계 : 25.04 ~ 26.09 기준 · 국내외 전체 2,139건(글로벌 167건 = JP 99 + US 68) · 26.09.07 갱신'
ws['B4']='※ RD : 표준 기준 ⑤ — 수량·비용 전체 2,139건 / 조회수·참여·CPV·CPE 성과 집계분 2,086건(IMD 2,050 + 바이오힐보 US 36)'
for k,t in enumerate([
 '※ 참여수 = 좋아요 + 댓글 / 평균 CPV = 크리에이터 원고료 ÷ 조회수 / 평균 CPE = 크리에이터 원고료 ÷ 참여수',
 '※ 소계 CPV·CPE = 크리에이터 원고료 492,408,205원 ÷ 조회수 56,612,765 / 참여 390,370 (표준 기준 ⑤) → CPV 8.70원 · CPE 1,261원. 바이오힐보 US 행은 원고료 미기재로 청구액 30,000,000원 기준이라 브랜드 행 단순 합과 소계가 일치하지 않는다.',
 '※ 세금계산서 기준으로 보면 900,344,264원 ÷ 조회수 56,612,765 = CPV 15.90원 / ÷ 참여 390,370 = CPE 2,306원. 원고료는 마크업·VAT 제외 금액으로 청구액의 54.7%다.',
 '※ 세금계산서 발행 확정 총액 900,344,264원 = 26.02~08 확정 840,514,264 + 25년 발행 59,830,000(바이오힐보 US 30,000,000 포함). + JP 미발행 8,904,000 → 진행 기준 909,248,264원.',
 '※ 좋아요 데이터 공란 505건(24.6%) 보정 시 참여 469,050 · 참여율 0.83% · CPE 1,050원(원고료)/1,920원(청구). 위 표는 미보정 공식 수치다.',
]):
    c=ws.cell(sub+1+k,2); c.value=t; c.font=copy(nf)

wb.calculation.fullCalcOnLoad=True
wb.save(OUT); print('saved', OUT, '| 소계행', sub)
