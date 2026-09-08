# -*- coding: utf-8 -*-
"""브랜드별 UGC 수량·콘텐츠 성과·퍼포먼스 광고 성과 매트릭스
   원본 표(광고주 제출본) 레이아웃 그대로 재현 + 참여수/평균 CPV/평균 CPE 채움.
   모든 수치는 reports/data.py(SSOT) 단일 원천. 파생값은 전부 수식."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import data as D
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

D.verify()
I = D.IDX
B = {b[0]: b for b in D.BRANDS}

# ── 표 행 구성 (원본 표 순서). kind: 'n'=소계 포함, 's'=참고 소분류(소계 제외)
ROWS = [
    ('아이디얼포맨',        'n', ['아이디얼포맨']),
    ('브링그린',            'n', ['브링그린']),
    ('라운드어라운드',      'n', ['라운드어라운드']),
    ('바이오힐보',          'n', ['바이오힐보 (국내)', '바이오힐보 (해외·JP)', '바이오힐보 (해외·US)']),
    ('바이오힐보 (국내)',   's', ['바이오힐보 (국내)']),
    ('바이오힐보 (해외)',   's', ['바이오힐보 (해외·JP)', '바이오힐보 (해외·US)']),
    ('식물나라',            'n', ['식물나라']),
    ('웨이크메이크',        'n', ['웨이크메이크']),
    ('루테카',              'n', ['루테카']),
    ('컬러그램',            'n', ['컬러그램']),
    ('케어플러스',          'n', ['케어플러스']),
    ('필리밀리',            'n', ['필리밀리']),
    ('올더베러',            'n', ['올더베러']),
]
MEMO = {
 '아이디얼포맨':   '25년 4Q UGC 협업 첫 시작 · 최다 협업 브랜드(737건) · ROAS 304%',
 '브링그린':       '11개 브랜드 전체 광고비 43.8% · 전환 매출 48.8% 차지 → 협력광고 중심 스케일업',
 '라운드어라운드': '참여율(ER) 1.73% 전체 1위 · 퍼포먼스 광고는 1건만 집행(광고비 133,848원)',
 '바이오힐보':     '국내 195건 + 해외 75건(JP 7 · US 68) 합계 · 해외 성과 집계는 43건분',
 '바이오힐보 (국내)': '국내 195건만 · 퍼포먼스 광고비·매출은 전액 국내분 · 참고 소분류(소계 제외)',
 '바이오힐보 (해외)': 'JP 7건(26.06 NAD 크림, 건당 333,000 마크업 포함) + US 68건(25.08 틱톡, 3,000만원) · 성과 집계 43건 · 퍼포먼스 광고 미집행 · 참고 소분류(소계 제외)',
 '식물나라':       '빅&스몰웨이브의 시작(25.05) · 마이크로 IMC 성과 가장 잘 작동(ROAS 789%)',
 '웨이크메이크':   '26.04 행사·올영픽 등 미존재로 UGC 협업 수량 증대해도 좋을 것으로 판단 · JP 60건(건당 415,000) · 성과 집계 109건(26.09 JP 21건 미집계)',
 '루테카':         '퍼포먼스 광고 미집행 (브랜드 fade out)',
 '컬러그램':       'ROAS 572%(브랜드 5위) · 최신 IMD 반영으로 광고비 3.1배 재집계 · JP 32건(건당 415,000) → 협업 수량 증가 필요',
 '케어플러스':     '26.07 UGC 첫 협업 · 퍼포먼스 광고 미집행',
 '필리밀리':       '전체 브랜드 중 ROAS 1위(866%) · 최저 CPA 1,129원 확보 · 최소 수량(27건)으로 최고 효율',
 '올더베러':       '퍼포먼스 광고 미집행',
}

# ── 스타일
FONT = 'Arial'
def F(sz=9, b=False, color='000000'): return Font(name=FONT, size=sz, bold=b, color=color)
THIN = Side(style='thin', color='BFBFBF')
MED  = Side(style='medium', color='808080')
BOX  = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CTR  = Alignment(horizontal='center', vertical='center', wrap_text=True)
LFT  = Alignment(horizontal='left',   vertical='center', wrap_text=True)
RGT  = Alignment(horizontal='right',  vertical='center')
GRAY = PatternFill('solid', fgColor='D9D9D9')
YEL  = PatternFill('solid', fgColor='FFFF00')
BLUE = PatternFill('solid', fgColor='DCE6F1')
SUM_ = PatternFill('solid', fgColor='F2F2F2')
NUM, PCT, WON = '#,##0', '0%', '#,##0'

wb = Workbook()

# ══════════════════════════════════════════════════════════ Sheet 2: 산출근거
src = wb.create_sheet('산출근거')
SRC_HDR = ['브랜드', 'KR', '글로벌', '전체 수량', '성과 집계 건수', '조회수', '좋아요', '댓글',
           '시딩 비용', '비용 기준', '광고비', '전환 매출']
for j, h in enumerate(SRC_HDR, 1):
    c = src.cell(1, j, h); c.font = F(9, True); c.fill = GRAY; c.alignment = CTR; c.border = BOX
srow = {}
for i, b in enumerate(D.BRANDS, 2):
    src.cell(i, 1, b[0]).alignment = LFT
    for j, k in enumerate(['KR', 'GL', 'PF', 'V', 'L', 'C', 'COST'], 0):
        pass
    vals = [b[I['KR']], b[I['GL']], b[I['KR']] + b[I['GL']], b[I['PF']],
            b[I['V']], b[I['L']], b[I['C']], b[I['COST']],
            '원고료(마크업·VAT 제외)' if b[I['BASIS']] == 'F' else '청구(마크업 포함)',
            b[I['AD']], b[I['REV']]]
    for j, v in enumerate(vals, 2):
        c = src.cell(i, j, v)
        c.number_format = NUM if isinstance(v, int) else 'General'
        c.alignment = RGT if isinstance(v, int) else CTR
    for j in range(1, 13): src.cell(i, j).border = BOX; src.cell(i, j).font = F(9)
    srow[b[0]] = i
LAST = len(D.BRANDS) + 1
src.cell(LAST + 1, 1, '합계').font = F(9, True)
for j, col in enumerate(['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'L', 'M'], 0):
    pass
for col in ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'K', 'L']:
    c = src[f'{col}{LAST+1}']; c.value = f'=SUM({col}2:{col}{LAST})'
    c.font = F(9, True); c.number_format = NUM; c.alignment = RGT; c.fill = SUM_; c.border = BOX
src[f'A{LAST+1}'].fill = SUM_; src[f'A{LAST+1}'].border = BOX
src[f'J{LAST+1}'].fill = SUM_; src[f'J{LAST+1}'].border = BOX
for w, col in zip([22, 8, 9, 10, 13, 13, 11, 10, 14, 22, 14, 15], 'ABCDEFGHIJKL'):
    src.column_dimensions[col].width = w
n = LAST + 3
for t in ['[원천] reports/data.py (SSOT) — 세금계산서 월별 확정표·IMD 실측·2025 시딩 정산표 기재값. 추정·안분 없음.',
          '[시딩 비용] 국내 브랜드는 원고료(마크업·VAT 제외), 바이오힐보 해외 2건은 확정 청구액(마크업 포함) — 열 J에 기준 명기.',
          f'[미집계] 전체 {D.QTY:,}건 중 성과 집계 {D.PERF:,}건. 차이 {D.UNPERF}건 = 바이오힐보 US 32건 + 웨이크메이크 JP 26.09 21건.',
          f'[좋아요 공란] {D.LIKE_BLANK}건(조회수 {D.LIKE_BLANK_V:,})은 원천에 좋아요 미기재 → 추정하지 않음. 참여수·ER·CPE는 불리하게 편향됨.']:
    src.cell(n, 1, t).font = F(8, color='808080'); n += 1

# ══════════════════════════════════════════════════════════ Sheet 1: 브랜드별 성과
ws = wb.active; ws.title = '브랜드별_성과'
GROUPS = [('브랜드', 1, 1), ('UGC 수량', 2, 4), ('콘텐츠 성과', 5, 9), ('퍼포먼스 광고 성과', 10, 12), ('설명', 13, 13)]
HDR = ['브랜드', 'KR', '글로벌', '전체', '조회수', '평균 조회수', '참여수', '평균 CPV', '평균 CPE',
       '광고비', '전환 매출', 'ROAS', '설명']
for name, c1, c2 in GROUPS:
    ws.merge_cells(start_row=1, start_column=c1, end_row=1 if c1 != c2 else 2,
                   end_column=c2) if c1 != c2 else ws.merge_cells(start_row=1, start_column=c1, end_row=2, end_column=c2)
    c = ws.cell(1, c1, name); c.font = F(10, True); c.fill = GRAY; c.alignment = CTR
for j, h in enumerate(HDR, 1):
    if j in (1, 13): continue
    c = ws.cell(2, j, h); c.font = F(9, True); c.fill = YEL if j in (7, 8, 9) else GRAY; c.alignment = CTR
for r in (1, 2):
    for j in range(1, 14):
        ws.cell(r, j).border = Border(left=THIN, right=THIN, top=MED, bottom=MED if r == 2 else THIN)
ws.row_dimensions[1].height = 20; ws.row_dimensions[2].height = 22

R0 = 3
sum_rows = []
for k, (label, kind, members) in enumerate(ROWS):
    r = R0 + k
    q  = ' + '.join(f"산출근거!D{srow[m]}" for m in members)
    kr = ' + '.join(f"산출근거!B{srow[m]}" for m in members)
    gl = ' + '.join(f"산출근거!C{srow[m]}" for m in members)
    pf = ' + '.join(f"산출근거!E{srow[m]}" for m in members)
    vw = ' + '.join(f"산출근거!F{srow[m]}" for m in members)
    lk = ' + '.join(f"산출근거!G{srow[m]}" for m in members)
    cm = ' + '.join(f"산출근거!H{srow[m]}" for m in members)
    cs = ' + '.join(f"산출근거!I{srow[m]}" for m in members)
    ad = ' + '.join(f"산출근거!K{srow[m]}" for m in members)
    rv = ' + '.join(f"산출근거!L{srow[m]}" for m in members)
    ws.cell(r, 1, ('  ' if kind == 's' else '') + label)
    ws.cell(r, 2, f'={kr}')
    ws.cell(r, 3, f'=IF(({gl})=0,"-",{gl})')
    ws.cell(r, 4, f'={q}')
    ws.cell(r, 5, f'={vw}')
    ws.cell(r, 6, f'=IFERROR(ROUND(({vw})/({pf}),0),"-")')      # 평균 조회수 = 조회수 ÷ 성과 집계 건수
    ws.cell(r, 7, f'=({lk})+({cm})')                            # 참여수 = 좋아요 + 댓글
    ws.cell(r, 8, f'=IFERROR(ROUND(({cs})/({vw}),1),"-")')      # 평균 CPV = 시딩 비용 ÷ 조회수
    ws.cell(r, 9, f'=IFERROR(ROUND(({cs})/(({lk})+({cm})),0),"-")')  # 평균 CPE = 시딩 비용 ÷ 참여수
    ws.cell(r, 10, f'=IF(({ad})=0,"-",{ad})')
    ws.cell(r, 11, f'=IF(({rv})=0,"-",{rv})')
    ws.cell(r, 12, f'=IFERROR(IF(({ad})=0,"-",({rv})/({ad})),"-")')
    ws.cell(r, 13, MEMO[label])
    if kind == 'n': sum_rows.append(r)

RS = R0 + len(ROWS)          # 소계 행
NN = [r for r in sum_rows]
def add(col): return '+'.join(f'{col}{r}' for r in NN)
ws.cell(RS, 1, '소계')
ws.cell(RS, 2, f'={add("B")}')
# 글로벌 열은 0을 "-"로 표기하므로 SUM(개별셀) — SUM은 텍스트를 무시한다
ws.cell(RS, 3, '=SUM(' + ','.join(f'C{r}' for r in NN) + ')')
ws.cell(RS, 4, f'={add("D")}')
ws.cell(RS, 5, f'={add("E")}')
ws.cell(RS, 6, f'=ROUND(E{RS}/산출근거!E{LAST+1},0)')
ws.cell(RS, 7, f'={add("G")}')
ws.cell(RS, 8, f'=ROUND(산출근거!I{LAST+1}/E{RS},1)')
ws.cell(RS, 9, f'=ROUND(산출근거!I{LAST+1}/G{RS},0)')
ws.cell(RS, 10, f'=산출근거!K{LAST+1}')
ws.cell(RS, 11, f'=산출근거!L{LAST+1}')
ws.cell(RS, 12, f'=L{RS}/J{RS}')
ws.cell(RS, 13, '-')

# ── 서식
for r in range(R0, RS + 1):
    sub  = ROWS[r - R0][1] == 's' if r < RS else False
    tot  = (r == RS)
    fill = BLUE if sub else (SUM_ if tot else None)
    for j in range(1, 14):
        c = ws.cell(r, j)
        c.font = F(9, tot); c.border = BOX
        if fill: c.fill = fill
        if j == 1:   c.alignment = LFT
        elif j == 13: c.alignment = LFT; c.font = F(8, tot, '404040')
        else:
            c.alignment = RGT
            c.number_format = PCT if j == 12 else ('0.0' if j == 8 else NUM)
    ws.row_dimensions[r].height = 30 if not sub else 26
ws.row_dimensions[RS].height = 20

for w, col in zip([17, 8, 9, 9, 13, 12, 11, 10, 10, 14, 15, 8, 52], 'ABCDEFGHIJKLM'):
    ws.column_dimensions[col].width = w
ws.freeze_panes = 'B3'

n = RS + 2
NOTES = [
 f'※ 기준 — 수량: 전체 {D.QTY:,}건 / 조회수·참여수: 성과 집계 {D.PERF:,}건 (미집계 {D.UNPERF}건 = 바이오힐보 US 32 + 웨이크메이크 JP 26.09 21)',
 '① 참여수 = 좋아요 + 댓글 (원천 IMD 실측 합)',
 f'② 평균 CPV = 시딩 비용 ÷ 조회수 · ③ 평균 CPE = 시딩 비용 ÷ 참여수. 시딩 비용은 브랜드별 확인분 {D.COST:,}원'
 f' (원고료 기준 {D.COST_F:,} + 청구 기준 {D.COST_B:,}) — 세부는 [산출근거] 시트.',
 f'④ 세금계산서 발행 총액은 {D.INV_TOT:,}원이며 브랜드 귀속이 확정되지 않아 이 표의 CPV/CPE에 쓰지 않았다.'
 f' 발행액 기준 전체 CPV는 {D.INV_TOT/D.VIEW:.1f}원 · CPE는 {round(D.INV_TOT/D.ENGA):,}원.',
 '⑤ 평균 조회수 = 조회수 ÷ 성과 집계 건수 (전체 수량이 아님). 바이오힐보 해외·웨이크메이크만 두 값이 다르다.',
 '⑥ 바이오힐보 (국내)·(해외) 행은 참고 소분류로 소계에 포함하지 않는다 (바이오힐보 합계 행에 이미 포함).',
 f'⑦ 좋아요 공란 {D.LIKE_BLANK}건(조회수 {D.LIKE_BLANK_V:,})은 추정하지 않았다 → 참여수는 하한값, CPE는 상한값.',
 '⑧ ROAS = 전환 매출 ÷ 광고비. 광고 미집행 브랜드는 "-".',
]
for t in NOTES:
    c = ws.cell(n, 1, t); c.font = F(8, color='808080'); c.alignment = LFT
    ws.merge_cells(start_row=n, start_column=1, end_row=n, end_column=13); n += 1

wb.calculation.fullCalcOnLoad = True
OUT = 'reports/OYPB_브랜드별_성과매트릭스_HAN_260908.xlsx'
wb.save(OUT)
print('saved', OUT)
