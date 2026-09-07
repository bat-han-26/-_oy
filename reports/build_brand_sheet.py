# -*- coding: utf-8 -*-
"""OYPB 브랜드별 요약표 — 데이터는 reports/data.py (SSOT)에서만 읽는다."""
import openpyxl, sys, os
from copy import copy
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import data as D
D.verify()
I = D.IDX
BASIS = {'F': '원고료', 'B': '청구'}
SRC='/root/.claude/uploads/ea1a909c-9825-524c-af08-aa69686f7bb5/8bba4283-__________.xlsx'
OUT='/home/user/-_oy/reports/OYPB_브랜드별_시딩_비용_광고성과_HAN_260907.xlsx'
MEMO2 = {  # 브랜드별 2줄 설명
 '아이디얼포맨':'25년 4Q UGC 협업 첫 시작을 바탕으로 \n250%대 ROAS는 300% 이상으로 올라 섬',
 '브링그린':'전체 퍼포먼스 광고비 44% 집행 · 전환 매출 49% 창출\nUGC&협력광고 중심 티트리 라인 스케일업 필요',
 '라운드어라운드':'PR키트 진행 등 참여율 1위(1.73%)\n광고는 281건 중 1건만 집행',
 '바이오힐보 (국내)':'퍼포먼스 광고 성과는 전량 국내에서 발생\n최대 조회수 764,503뷰',
 '바이오힐보 (해외·JP)':'26.06 NAD 크림 나노 시딩 테스트\n비용 건당 333,000(마크업 포함) · 조회수·참여 저조',
 '바이오힐보 (해외·US)':'틱톡 68건 · 비용 30,000,000(마크업 포함)\n2025 시딩 정산 59,830,000원 내 · 최대 2,954,129뷰 = 전체 1위',
 '식물나라':'빅&스몰웨이브의 시작(25.05)\n마이크로 IMC 성과 가장 잘 워킹하는 브랜드',
 '웨이크메이크':'JP 60건 건당 415,000 (21건은 26.09 진행 · 미발행)\n26.04 행사·올영픽 미존재로 협업 수량 증대 여지',
 '루테카':'퍼포먼스 광고 미집행\n(브랜드 fade out)',
 '컬러그램':'IMD 26.09.07로 103건 확정(v8 77건은 부분집합)\nJP 32건 건당 415,000 · ROAS 572%',
 '케어플러스':'26.07 UGC 첫 협업 · 퍼포먼스 광고 미집행\n원고료 마이크로 5건×20만 + 나노 27건×10만',
 '필리밀리':'전체 브랜드 중 ROAS 1위(866%)\n최저 CPA 확보',
 '올더베러':'퍼포먼스 광고 미집행',
}
wb=openpyxl.load_workbook(SRC); ws=wb['Sheet1']
DEC='_-* #,##0.0_-;\\-* #,##0.0_-;_-* "-"_-;_-@_-'
tpl ={c:{k:copy(getattr(ws.cell(8,c),k)) for k in ('font','fill','border','alignment')} for c in range(2,15)}
tpls={c:{k:copy(getattr(ws.cell(19,c),k)) for k in ('font','fill','border','alignment')} for c in range(2,15)}
nf ={c:ws.cell(8,c).number_format for c in range(2,15)}
nfs={c:ws.cell(19,c).number_format for c in range(2,15)}
TXT={12}  # 기준 열(문자)
def style(r,c,val,sub=False):
    cell=ws.cell(r,c); cell.value=val
    src=(tpls if sub else tpl); fmt=(nfs if sub else nf)
    base = src[c] if c<=14 else src[14]
    for k in ('font','fill','border','alignment'): setattr(cell,k,copy(base[k]))
    if c==11: cell.number_format='0.00%'
    elif c==14: cell.number_format=DEC
    elif c==12: cell.number_format='General'
    elif c<=14: cell.number_format=fmt[c]
    else: cell.number_format=fmt[6]
    if c==12: cell.alignment=copy(tpl[2]['alignment'])
r0=8
for i,b in enumerate(D.BRANDS):
    r=r0+i; kr,gl,ad=b[I['KR']],b[I['GL']],b[I['AD']]
    memo=MEMO2[b[I['NAME']]]
    vals={2:b[I['NAME']],3:kr,4:gl,5:f'=SUM(C{r}:D{r})',6:b[I['PF']],7:b[I['V']],8:b[I['L']],9:b[I['C']],
          10:f'=H{r}+I{r}',11:f'=J{r}/G{r}',12:b[I['COST']],13:BASIS[b[I['BASIS']]],14:f'=L{r}/G{r}',
          15:ad,16:b[I['REV']],17:(f'=P{r}/O{r}' if ad else '-'),18:memo}
    for c in range(2,19): style(r,c,vals[c])
    ws.row_dimensions[r].height = 39.6 if '\n' in memo else 26.4
r=r0+len(D.BRANDS)
S=lambda col: f'=SUM({col}{r0}:{col}{r-1})'
vals={2:'합계',3:S('C'),4:S('D'),5:S('E'),6:S('F'),7:S('G'),8:S('H'),9:S('I'),10:S('J'),
      11:f'=J{r}/G{r}',12:S('L'),13:'혼합',14:f'=L{r}/G{r}',15:S('O'),16:S('P'),17:f'=P{r}/O{r}',
      18:f'수량 {D.QTY:,} / 성과 집계 {D.PERF:,}'}
for c in range(2,19): style(r,c,vals[c],sub=True)
ws.row_dimensions[r].height=26.4
sub=r
# 헤더 재구성
G6=copy(ws.cell(6,3)._style); G7=copy(ws.cell(7,3)._style)
for mr in list(ws.merged_cells.ranges):
    if mr.min_row in (6,7) or mr.max_row in (6,7): ws.unmerge_cells(str(mr))
for c in range(2,19):
    ws.cell(6,c)._style=copy(G6); ws.cell(7,c)._style=copy(G7)
    ws.cell(6,c).value=None; ws.cell(7,c).value=None
for a,bb,val in [(3,6,'시딩 수량'),(7,11,'콘텐츠 성과 (실측)'),(12,14,'시딩 비용'),(15,17,'퍼포먼스 광고')]:
    ws.cell(6,a).value=val; ws.merge_cells(start_row=6,start_column=a,end_row=6,end_column=bb)
for c,val in [(2,'브랜드'),(18,'설명')]:
    ws.cell(6,c).value=val; ws.merge_cells(start_row=6,start_column=c,end_row=7,end_column=c)
for c,val in [(3,'KR'),(4,'글로벌'),(5,'전체'),(6,'성과 집계'),(7,'조회수'),(8,'좋아요'),(9,'댓글'),
              (10,'참여'),(11,'ER'),(12,'금액'),(13,'기준'),(14,'CPV'),(15,'광고비'),(16,'전환 매출'),(17,'ROAS')]:
    ws.cell(7,c).value=val
for col,w in (('K',9),('L',14),('M',8),('N',9),('O',14),('P',15),('Q',9),('R',40)): ws.column_dimensions[col].width=w
ws['B2']='OYPB 브랜드별 시딩 · 비용 · 성과  (팩트 집계)'
nfont=copy(ws['B3'].font)
ws['B3']=f'※ 통계 : 25.04 ~ 26.09 · 국내외 전체 {D.QTY:,}건(KR {D.KR_Q:,} + 글로벌 {D.GL_Q} = JP 99 + US 68) · 26.09.07 기준'
ws['B4']='※ RD : 모든 값은 원천 자료 기재값이다. 추정·안분·보정값은 사용하지 않았다.'
NOTES=[
 f'※ 수량 = 글로벌 시딩 라인 자료(JP 99 · US 68) + IMD 집계분 / 성과 집계 = IMD에 성과가 기재된 건수 {D.PERF:,}건 (미집계 {D.UNPERF}건 = US 32 + JP 26.09 21)',
 f'※ 조회수·좋아요·댓글 = IMD 실측값. 좋아요 데이터가 없는 {D.LIKE_BLANK}건(조회수 {D.LIKE_BLANK_V:,})은 추정하지 않았으므로 참여·ER은 실제보다 낮게(불리하게) 나온 값이다.',
 '※ 시딩 비용 기준 — [원고료] 마크업·VAT 제외. 국내는 IMD 행 단위 기재값, JP는 확정 단가 건당 415,000, 케어플러스는 시딩현황 3,700,000.',
 f'※ 시딩 비용 기준 — [청구] 마크업 포함. 바이오힐보 JP 7건 × 333,000 = 2,331,000 / 바이오힐보 US 틱톡 {D.INV_2025_US:,} (2025 시딩 정산 {D.INV_2025:,}원 내).',
 f'※ 합계 {D.COST:,}원 = 원고료 기준 {D.COST_F:,} + 청구 기준 {D.COST_B:,}. 기준이 섞여 있어 소계 CPV {D.COST/D.VIEW:.2f}원은 참고값이다.',
 f'※ 전체 청구 기준 — 세금계산서 발행 {D.INV_TOT:,}원 (26.02~08 확정 {D.INV_2026:,} + 2025 시딩 정산 {D.INV_2025:,}) ÷ 조회수 {D.VIEW:,} = CPV {D.INV_TOT/D.VIEW:.2f}원 / ÷ 참여 {D.ENGA:,} = CPE {D.INV_TOT/D.ENGA:,.0f}원 / ÷ {D.QTY:,}건 = 건당 {D.INV_TOT//D.QTY:,}원',
 f'※ JP 미발행 {D.JP_UNBILLED:,}원(웨이크메이크 21건 × 415,000) 별도 → 진행 기준 {D.INV_TOT+D.JP_UNBILLED:,}원',
 '※ 브랜드별 세금계산서 귀속은 발행 내역에 브랜드 미귀속액 70,431,189원이 있어 산출하지 않았다.',
]
for k,t in enumerate(NOTES):
    c=ws.cell(sub+1+k,2); c.value=t; c.font=copy(nfont)
wb.calculation.fullCalcOnLoad=True
wb.save(OUT); print('saved', OUT, '| 합계행', sub)
