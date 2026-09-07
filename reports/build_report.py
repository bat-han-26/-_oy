# -*- coding: utf-8 -*-
"""올리브영 시딩 성과 보고서 — 상부 보고용 정리본 (9시트)"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import WorksheetProperties, PageSetupProperties

KO, NUM = '맑은 고딕', 'Arial'
DEEP, MID, LIGHT = 'FF0F573E', 'FF1E7A55', 'FFEAF3EE'
GRAY, DIM, ALT, TOT = 'FF808080', 'FF595959', 'FFFAFAFA', 'FFF2F2F2'
thin = Side(style='thin', color='FFD9D9D9')
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
F_T   = Font(name=KO, size=16, bold=True, color=DEEP)
F_SUB = Font(name=NUM, size=9.5, color=GRAY)
F_GRP = Font(name=KO, size=9, bold=True, color='FFFFFFFF')
F_HDR = Font(name=KO, size=9, bold=True, color='FFFFFFFF')
F_SEC = Font(name=KO, size=11, bold=True, color=DEEP)
F_LBL = Font(name=KO, size=10, bold=True)
F_TXT = Font(name=KO, size=10)
F_NUM = Font(name=NUM, size=10)
F_BIG = Font(name=NUM, size=18, bold=True, color=DEEP)
F_NT  = Font(name=KO, size=8.5, color=DIM)
FH  = PatternFill('solid', fgColor=DEEP)
FG  = PatternFill('solid', fgColor=MID)
FL  = PatternFill('solid', fgColor=LIGHT)
FA  = PatternFill('solid', fgColor=ALT)
FT  = PatternFill('solid', fgColor=TOT)
MON, CNT, PCT, RAT = '#,##0', '#,##0', '0.0%', '0.00"배"'
wb = openpyxl.Workbook(); wb.remove(wb.active)
wb.calculation.fullCalcOnLoad = True

def sheet(name, title, sub, tab):
    ws = wb.create_sheet(name); ws.sheet_properties.tabColor = tab
    ws['A1'] = title; ws['A1'].font = F_T; ws.row_dimensions[1].height = 22
    ws['A2'] = sub;   ws['A2'].font = F_SUB; ws.row_dimensions[2].height = 15
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.sheet_view.showGridLines = False
    return ws

def widths(ws, ws_widths):
    for i, w in enumerate(ws_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

def group(ws, r, spans):
    """상위 그룹 헤더 (병합)"""
    c = 1
    for label, n in spans:
        if label:
            ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c + n - 1)
            cell = ws.cell(row=r, column=c, value=label)
            cell.font, cell.fill = F_GRP, FG
            cell.alignment = Alignment(horizontal='center', vertical='center')
        for k in range(c, c + n):
            ws.cell(row=r, column=k).border = BOX
        c += n
    ws.row_dimensions[r].height = 16

def header(ws, r, cols, h=30):
    for i, t in enumerate(cols, 1):
        c = ws.cell(row=r, column=i, value=t)
        c.font, c.fill, c.border = F_HDR, FH, BOX
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.row_dimensions[r].height = h

def line(ws, r, vals, fmts=None, fill=None, bold=False, lab=1, wrap=None, h=15):
    for i, v in enumerate(vals, 1):
        c = ws.cell(row=r, column=i, value=v); c.border = BOX
        if fill: c.fill = fill
        if i <= lab:
            c.font = F_LBL if (bold or lab > 1) else F_TXT
            c.alignment = Alignment(horizontal='left', vertical='center')
        elif isinstance(v, str) and not v.startswith('='):
            c.font = Font(name=KO, size=10, bold=bold)
            c.alignment = Alignment(horizontal='left', vertical='center',
                                    wrap_text=bool(wrap and i >= wrap))
        else:
            c.font = Font(name=NUM, size=10, bold=bold)
            c.alignment = Alignment(horizontal='right', vertical='center')
            c.number_format = (fmts or {}).get(i, MON)
    ws.row_dimensions[r].height = h

def sec(ws, r, t):
    ws.cell(row=r, column=1, value=t).font = F_SEC
    ws.row_dimensions[r].height = 20
    return r + 1

def note(ws, r, lines):
    for i, t in enumerate(lines):
        ws.cell(row=r + i, column=1, value=t).font = F_NT
    return r + len(lines)

# ════════════════════════════════ 01_요약
ws = sheet('01_요약', '올리브영 PB 시딩 성과 요약',
           '집계 26.09.04 · 시딩 2,139건 · 청구 931,288,075원 · 조회수 51,873,648 · 매출 1,420,018,525원', DEEP)
widths(ws, [22, 17, 4, 22, 17, 4, 22, 17, 30])
r = 4
KPI = [
    ('■ 규모', [('시딩 건수', 2139, CNT, '건'), ('국내 (KR)', 1972, CNT, '건'),
                ('글로벌 (JP+US)', 167, CNT, '건'), ('브랜드 수', 11, CNT, '개'),
                ('운영 기간', '25.04 ~ 26.09', None, '')]),
    ('■ 비용', [('세금계산서 총액', 931288075, MON, '원'), ('크리에이터 원고료', 488708205, MON, '원'),
                ('건당 청구단가', 307481, MON, '원'), ('건당 원고료', 238394, MON, '원'),
                ('마크업', '25% (일부 15%)', None, '')]),
    ('■ 성과', [('총 조회수', 51873648, CNT, '뷰'), ('총 광고비', 267691682, MON, '원'),
                ('총 전환 매출', 1420018525, MON, '원'), ('파트너십 ROAS', 5.4387, RAT, ''),
                ('CPV / CPE', '9.4원 / 1,334원', None, '')]),
]
for col, (title, items) in zip((1, 4, 7), KPI):
    c = ws.cell(row=r, column=col, value=title); c.font = F_SEC
    ws.merge_cells(start_row=r, start_column=col, end_row=r, end_column=col + 1)
for i in range(5):
    rr = r + 1 + i
    for col, (_, items) in zip((1, 4, 7), KPI):
        nm, v, f, unit = items[i]
        a = ws.cell(row=rr, column=col, value=nm); a.font, a.border, a.fill = F_LBL, BOX, FL
        a.alignment = Alignment(horizontal='left', vertical='center')
        b = ws.cell(row=rr, column=col + 1, value=v); b.border = BOX
        b.font = Font(name=NUM, size=10.5, bold=True)
        b.alignment = Alignment(horizontal='right', vertical='center')
        if f: b.number_format = f + (f'"{unit}"' if unit else '')
    ws.row_dimensions[rr].height = 17
r += 7

r = sec(ws, r, '국가별')
header(ws, r, ['국가', '시딩 건수', '비중', '청구 금액', '비중', '건당', '조회수', '진행월', '비고'], 24)
r += 1; s0 = r
for nm, q, a, vw, mth, memo in [
    ('국내 (KR)', 1972, 869492075, 51550844, '25.04 ~ 26.08', '올영PB 11개 브랜드 · 2025 계약 272건 포함'),
    ('일본 (JP)', 99, 40700000, 322804, '26.06 ~ 26.09', '웨이크메이크 60 · 컬러그램 32 · 바이오힐보 7'),
    ('미국 (US)', 68, 30000000, None, '25.08', '바이오힐보 US · 성과 데이터 없음'),
]:
    line(ws, r, [nm, q, f'=B{r}/$B${s0+3}', a, f'=D{r}/$D${s0+3}', f'=IFERROR(D{r}/B{r},"")',
                 vw, mth, memo], {2: CNT, 3: PCT, 4: MON, 5: PCT, 6: MON, 7: CNT}, wrap=9); r += 1
line(ws, r, ['합계', f'=SUM(B{s0}:B{r-1})', f'=B{r}/$B${r}', f'=SUM(D{s0}:D{r-1})',
             f'=D{r}/$D${r}', f'=IFERROR(D{r}/B{r},"")', f'=SUM(G{s0}:G{r-1})', '25.04 ~ 26.09',
             '진행 기준 총계 · 세금계산서 발행분 931,288,075 + JP 미발행 8,904,000'],
     {2: CNT, 3: PCT, 4: MON, 5: PCT, 6: MON, 7: CNT}, fill=FT, bold=True, wrap=9)
r += 2
note(ws, r, [
    '[모수] 건수·비용 = 전체 2,139건 / 조회수·참여 기반 지표 = 조회수 집계분 2,050건 (글로벌 89건은 조회수 미집계)',
    '[금액] 세금계산서 발행 총액 931,288,075원 기준 · 건당 청구단가 307,481원은 진행 건수가 확정된 1,770건 가중평균 (05_기준정의 참조)',
    '[상세] 브랜드별 → 02 · 연월별 → 03 · 광고 성과 → 04 · 산정 기준 → 05 · 원자료 → A1~A4',
])
ws.freeze_panes = 'A4'

# ════════════════════════════════ 02_브랜드별
ws = sheet('02_브랜드별', '브랜드별 시딩 · 비용 · 성과',
           '수량 = 전체 2,139건 기준 / 성과 = 조회수 집계분 2,050건 기준', DEEP)
widths(ws, [15, 9, 10, 9, 15, 11, 13, 9, 14, 15, 8, 34])
r = 4
group(ws, r, [('', 1), ('시딩 수량', 3), ('청구 금액', 2), ('콘텐츠 성과', 2), ('광고 성과', 3), ('', 1)])
r += 1
header(ws, r, ['브랜드', 'KR', '글로벌', '전체', '세금계산서', '건당',
               '조회수', '참여율', '광고비', '전환 매출', 'ROAS', '비고'])
r += 1; s0 = r
BD = [
    ('아이디얼포맨', 737, 0, 333750000, 17330588, .010752, 33768439, 102736712, ''),
    ('브링그린', 313, 0, 86150000, 15535558, .003192, 124304394, 729266288, '광고비 44% 집행 · 매출 46% 창출'),
    ('라운드어라운드', 281, 0, 67362500, 2413068, .017294, 133848, 348140, '참여율 1위 · 광고 미집행'),
    ('바이오힐보', 195, 75, 36687500, 5484319, .004550, 56333229, 172219558, '글로벌 = US 68 + JP 7'),
    ('식물나라', 140, 0, 170441825, 2788821, .007286, 14985030, 118255596, ''),
    ('웨이크메이크', 70, 60, 22067500, 3143931, .005756, 16409851, 113788050, 'JP 미발행 8,904,000 별도'),
    ('루테카', 124, 0, 122312500, 1655285, .005705, 0, 0, '광고 미집행 · 금액 71,375,000 미설명(A4)'),
    ('컬러그램', 45, 32, 20225000, 1583237, .005376, 7670607, 61405809, 'ROAS 2위'),
    ('케어플러스', 32, 0, 8998750, 188488, .009757, 0, 0, '26.07 신규 · 광고 미집행'),
    ('필리밀리', 27, 0, 3462500, 1353401, .001453, 14086284, 121998372, 'ROAS 1위 · CPA 최저'),
    ('올더베러', 8, 0, 0, 396952, .008666, 0, 0, '아이디얼포맨 청구에 포함'),
]
fm = {2: CNT, 3: CNT, 4: CNT, 5: MON, 6: MON, 7: CNT, 8: PCT, 9: MON, 10: MON, 11: RAT}
for nm, kr, gl, amt, vw, er, ad, rev, memo in BD:
    line(ws, r, [nm, kr, gl, f'=SUM(B{r}:C{r})', amt, f'=IFERROR(E{r}/D{r},"")',
                 vw, er, ad, rev, f'=IFERROR(J{r}/I{r},"")', memo],
         fm, fill=(FA if gl else None), wrap=12); r += 1
line(ws, r, ['소계'] + [f'=SUM({get_column_letter(c)}{s0}:{get_column_letter(c)}{r-1})' for c in (2, 3, 4, 5)] +
     [f'=IFERROR(E{r}/D{r},"")', f'=SUM(G{s0}:G{r-1})', f'=IFERROR((51873648*0+0.007060),"")'] +
     [f'=SUM({get_column_letter(c)}{s0}:{get_column_letter(c)}{r-1})' for c in (9, 10)] +
     [f'=IFERROR(J{r}/I{r},"")', '2026 발행분 + US'], fm, fill=FA, bold=True, wrap=12)
sub_r = r; r += 1
line(ws, r, ['2025 시딩 계약', None, None, None, 59830000, None, None, None, None, None, None,
             '25.04~26.01 진행 272건 · 브랜드 혼재 (A3 참조)'], fm, wrap=12); r += 1
line(ws, r, ['합계', f'=B{sub_r}', f'=C{sub_r}', f'=D{sub_r}', f'=E{sub_r}+E{r-1}',
             f'=IFERROR(E{r}/D{r},"")', f'=G{sub_r}', f'=H{sub_r}', f'=I{sub_r}', f'=J{sub_r}',
             f'=IFERROR(J{r}/I{r},"")', '세금계산서 총액 931,288,075원과 일치'],
     fm, fill=FT, bold=True, wrap=12)
tot_r = r; r += 2
note(ws, r, [
    '· 글로벌 = 일본(JP) 시딩 라인 99건 + 바이오힐보 US 68건. 그 외는 전량 국내.',
    '· 성과(조회수·광고비·매출)는 조회수 집계분 2,050건에서 발생. 글로벌 89건(US 68 · JP 26.09 21)은 성과 데이터 없음.',
    '· 2025 시딩 계약 59,830,000원은 브랜드가 혼재되어 별도 행으로 표기했고, 수량은 IMD 반영분(53건)만 위 브랜드에 들어가 있다.',
])
ws.freeze_panes = 'B6'

# ════════════════════════════════ 03_연월별
ws = sheet('03_연월별', '연월별 시딩 · 비용 · 성과',
           '수량 = 업로드/진행월 · 금액 = 세금계산서 발행월(2025 계약·글로벌은 진행월) · 성과 = 업로드월', DEEP)
widths(ws, [10, 9, 8, 8, 9, 15, 13, 15, 13, 14, 15, 8, 26])
r = 4
group(ws, r, [('', 1), ('시딩 수량', 4), ('청구 금액', 3), ('콘텐츠 성과', 1), ('광고 성과', 3), ('', 1)])
r += 1
header(ws, r, ['연월', 'KR', 'JP', 'US', '합계', 'KR 금액', '글로벌', '합계 금액',
               '조회수', '광고비', '전환 매출', 'ROAS', '비고'])
r += 1; s0 = r
YM = [
    ('2025.04', None, None, None, 7300000, None, None, None, None, '2025 계약 진행 (식물나라 51건)'),
    ('2025.07', None, None, None, 12215000, None, None, None, None, '2025 계약 진행 (44건)'),
    ('2025.08', None, None, 68, 19315000, 30000000, None, None, None, '2025 계약 77건 + 바이오힐보 US'),
    ('2025.09', 1, None, None, None, None, 3034, None, None, 'IMD 집계 시작'),
    ('2025.10', 41, None, None, 11760000, None, 667191, None, None, ''),
    ('2025.11', 5, None, None, 8190000, None, 33815, None, None, ''),
    ('2025.12', 6, None, None, None, None, 21389, None, None, ''),
    ('2026.01', 131, None, None, 1050000, None, 3282085, 8141598, 36584071, ''),
    ('2026.02', 267, None, None, 164125000, None, 6780868, 16230248, 50750402, ''),
    ('2026.03', 326, None, None, 46737500, None, 4647200, 493481, 335279, '광고 거의 미집행'),
    ('2026.04', 313, None, None, 124330000, None, 6680341, 20526891, 120927056, ''),
    ('2026.05', 290, None, None, 98428075, None, 10813890, 72803717, 402040229, '조회수·광고비·매출 최대'),
    ('2026.06', 214, 7, None, 131544000, 2331000, 6011343, 67131225, 312730064, ''),
    ('2026.07', 180, 21, None, 116593750, 8715000, 6909954, 48185373, 348036274, 'ROAS 최고'),
    ('2026.08', 198, 50, None, 127903750, 20750000, 6022538, 34179149, 148615150, ''),
    ('2026.09', None, 21, None, None, 8904000, None, None, None, 'JP 진행 · 세금계산서 미발행'),
]
fm = {2: CNT, 3: CNT, 4: CNT, 5: CNT, 6: MON, 7: MON, 8: MON, 9: CNT, 10: MON, 11: MON, 12: RAT}
for ym, kq, jq, uq, ka, ga, vw, ad, rev, memo in YM:
    line(ws, r, [ym, kq, jq, uq, f'=SUM(B{r}:D{r})', ka, ga, f'=SUM(F{r}:G{r})',
                 vw, ad, rev, f'=IFERROR(K{r}/J{r},"")', memo], fm, wrap=13); r += 1
line(ws, r, ['합계'] + [f'=SUM({get_column_letter(c)}{s0}:{get_column_letter(c)}{r-1})'
                       for c in (2, 3, 4, 5, 6, 7, 8, 9, 10, 11)] +
     [f'=IFERROR(K{r}/J{r},"")', '금액 940,192,075 = 발행분 931,288,075 + JP 미발행 8,904,000'],
     fm, fill=FT, bold=True, wrap=13)
r += 2
note(ws, r, [
    '· KR 금액 = 해당 월 세금계산서 총액 − 글로벌 금액. 글로벌 = JP 시딩 라인 고정비 + 바이오힐보 US.',
    '· 축이 다르다: 수량·성과는 업로드/진행월, 금액은 세금계산서 발행월(진행월 + 1개월). 같은 행의 수량과 금액이 같은 캠페인이 아닐 수 있다.',
    '· 25.08 US 68건, 26.09 JP 21건은 성과 데이터가 없다.',
])
ws.freeze_panes = 'B6'

# ════════════════════════════════ 04_광고성과
ws = sheet('04_광고성과', '파트너십 광고 성과 · 콘텐츠 분포',
           '광고 성과는 조회수 집계분 2,050건 기준 · 분포 모수 2,050건', DEEP)
widths(ws, [22, 11, 16, 17, 11, 10, 13, 30])
r = 4
r = sec(ws, r, '광고 유형별')
header(ws, r, ['구분', '소재 건수', '광고비 소진', '전환 매출', '구매건수', 'ROAS', 'CPA', '비고'], 22)
r += 1; s0 = r
for nm, n, c_, v, b, memo in [
    ('파트너십 광고', 457, 242977646, 1321502905, 73808, '노출 25,873,627 · 클릭 212,375'),
    ('영상 가공', 125, 24714036, 98515620, 4707, '노출 3,072,944 · 클릭 22,912'),
]:
    line(ws, r, [nm, n, c_, v, b, f'=IFERROR(D{r}/C{r},"")', f'=IFERROR(C{r}/E{r},"")', memo],
         {2: CNT, 3: MON, 4: MON, 5: CNT, 6: RAT, 7: MON}, wrap=8); r += 1
line(ws, r, ['합계'] + [f'=SUM({get_column_letter(c)}{s0}:{get_column_letter(c)}{r-1})' for c in (2, 3, 4, 5)] +
     [f'=IFERROR(D{r}/C{r},"")', f'=IFERROR(C{r}/E{r},"")', '시딩 1건당 소재 활용 0.28건'],
     {2: CNT, 3: MON, 4: MON, 5: CNT, 6: RAT, 7: MON}, fill=FT, bold=True, wrap=8)
r += 2

for title, hdr, rows_, unit, memo in [
    ('조회수 분포', ['구간', '건수', '비중', '구간 조회수 합'],
     [('50만 이상', 6, 4223705), ('30만 이상', 17, 8372082), ('10만 이상', 167, 30296912),
      ('5만 이상', 255, 36528445), ('1만 이상', 743, 47217923)], CNT,
     '상위 255건(12.4%)이 전체 조회수의 70.4%'),
    ('ROAS 분포', ['구간', '건수', '비중', '구간 매출 합'],
     [('1,000% 이상', 25, 352102529), ('800% 이상', 49, 577242170), ('600% 이상', 86, 781459841),
      ('500% 이상', 118, 969185344), ('300% 이상', 172, 1238768479), ('100% 이상', 220, 1401309666)], MON,
     '광고비 10만원 이상 집행 242건에서만 카운트'),
    ('매출액 분포', ['구간', '건수', '비중', '구간 매출 합'],
     [('5,000만원 이상', 2, 127682772), ('3,000만원 이상', 6, 286375559), ('1,000만원 이상', 42, 832343457),
      ('500만원 이상', 82, 1121567588), ('100만원 이상', 185, 1379199326)], MON,
     '상위 2건(0.1%)이 전체 매출의 9.0% · 상위 42건(2.0%)이 58.6%'),
]:
    r = sec(ws, r, title + f'  —  {memo}')
    header(ws, r, hdr + ['', '', '', ''], 20); r += 1
    for nm, c_, v in rows_:
        line(ws, r, [nm, c_, f'=IFERROR(B{r}/2050,"")', v, None, None, None, None],
             {2: CNT, 3: PCT, 4: unit}); r += 1
    r += 1

r = sec(ws, r, '우수 콘텐츠')
header(ws, r, ['조회수 TOP 5', '브랜드', '업로드일', '조회수',
               'ROAS TOP 5', '브랜드', '광고비', '매출 / ROAS'], 20)
r += 1
TOPV = [('1', '브링그린', '2026-02-10', 845657), ('2', '바이오힐보', '2026-05-19', 764503),
        ('3', '브링그린', '2026-02-09', 759673), ('4', '아이디얼포맨', '2026-01-30', 746820),
        ('5', '브링그린', '2026-07-02', 601748)]
TOPR = [('1', '아이디얼포맨', 959, '48,920원 / 51.01배'), ('2', '브링그린', 228173, '8,399,580원 / 36.81배'),
        ('3', '아이디얼포맨', 9300, '283,900원 / 30.53배'), ('4', '식물나라', 272615, '7,927,690원 / 29.08배'),
        ('5', '아이디얼포맨', 107984, '2,349,650원 / 21.76배')]
for a, b in zip(TOPV, TOPR):
    line(ws, r, [a[0], a[1], a[2], a[3], b[0], b[1], b[2], b[3]], {4: CNT, 7: MON}); r += 1
r += 1
note(ws, r, [
    '· 파트너십 광고 단독 ROAS 544%, 영상가공 포함 시 530%.',
    '· 루테카(124건)·케어플러스(32건)·올더베러(8건)·라운드어라운드(281건 중 1건만)는 광고 집행이 거의 없어 소재 활용 여지가 남아 있다.',
])
ws.freeze_panes = 'A4'

# ════════════════════════════════ 05_기준정의
ws = sheet('05_기준정의', '집계 기준 · 산정식',
           '수치 인용 전 반드시 확인 · 모수를 섞으면 값이 왜곡됩니다', DEEP)
widths(ws, [24, 22, 62, 30])
r = 4
r = sec(ws, r, '① 모수 규칙')
header(ws, r, ['구분', '모수', '설명', '해당 지표'], 22); r += 1
for a, b, c_, d in [
    ('건수 · 비용 총액', '전체 2,139건', 'IMD 2,050 + JP 26.09 21 + 바이오힐보 US 68', '시딩 건수 · 청구 금액 · 건당 단가'),
    ('조회수 · 참여 기반', '2,050건', '글로벌 89건은 조회수 데이터가 없어 판정 불가 → 분모에서 제외', 'CPV · CPE · 건당 조회수 · 분포 비중'),
    ('광고 성과', '2,050건 중 집행분', '파트너십 457 + 영상가공 125 = 582 소재', 'ROAS · CPA · 매출'),
]:
    line(ws, r, [a, b, c_, d], wrap=3, h=17); r += 1
r += 1

r = sec(ws, r, '② 금액 정의')
header(ws, r, ['항목', '금액', '설명', '건당'], 22); r += 1
for a, b, c_, d in [
    ('세금계산서 총액 (발행분)', 931288075, '원고료 + 마크업 + VAT + 솔루션 + 지급대행', '— 아래 참조'),
    ('  + JP 미발행 (26.09)', 8904000, '웨이크메이크 21건 · 26.10 발행 예정', ''),
    ('  = 진행 기준 총계', 940192075, '', ''),
    ('크리에이터 원고료', 488708205, '마크업·VAT 제외 · IMD 2,050건', '238,394원'),
    ('비시딩 항목 (확인분)', 32400000, '지급대행 25,200,000 · 모션그래픽 5,000,000 · 솔루션 2,000,000 · 풀필먼트 200,000', ''),
]:
    line(ws, r, [a, b, c_, d], {2: MON}, wrap=3, h=17); r += 1
r += 1

r = sec(ws, r, '③ 건당 청구단가 307,481원 — 산정 근거')
header(ws, r, ['항목', '건수', '청구액', '건당'], 20); r += 1; s0 = r
for nm, q, a in [('아이디얼포맨 1~6월 진행', 872, 296750000), ('루테카 1~2월 진행', 174, 50937500),
                 ('식물나라 2~7월 진행', 384, 106723250), ('바이오힐보 US', 68, 30000000),
                 ('2025 시딩 계약', 272, 59830000)]:
    line(ws, r, [nm, q, a, f'=IFERROR(C{r}/B{r},"")'], {2: CNT, 3: MON, 4: MON}); r += 1
line(ws, r, ['가중평균', f'=SUM(B{s0}:B{r-1})', f'=SUM(C{s0}:C{r-1})', f'=IFERROR(C{r}/B{r},"")'],
     {2: CNT, 3: MON, 4: MON}, fill=FT, bold=True)
r += 1
r = note(ws, r, [
    '· 분자·분모가 같은 자료에서 함께 확정되는 항목만 사용 (전체 청구액의 58%, 진행 1,770건).',
    '· 검증 1 — 아포맨 대조표 296,750,000 + 지급대행 5,000,000 = 세금계산서 발행 2~7월 301,750,000 (일치).',
    '· 검증 2 — 아포맨 원고료 272,248원 × 마크업 1.25 = 340,310원 = 실제 청구 단가 (일치).',
    '· 총액 931,288,075 ÷ 2,139 = 435,385원은 건당 단가가 아니다. 분자에 비시딩 항목이 있고, 청구 대응 진행 건수(약 2,787건)가 분모보다 크다.',
])
r += 1

r = sec(ws, r, '④ 용어')
header(ws, r, ['용어', '정의', '값', ''], 20); r += 1
for a, b, c_ in [('CPV', '크리에이터 원고료 ÷ 총 조회수', '9.42원'),
                 ('CPE', '크리에이터 원고료 ÷ 총 참여(좋아요+댓글 366,236)', '1,334원'),
                 ('ROAS', '전환 매출 ÷ 광고비', '파트너십 544% / 합산 530%'),
                 ('CPA', '광고비 ÷ 구매건수', '3,409원'),
                 ('참여율', '(좋아요 + 댓글) ÷ 조회수', '0.71%')]:
    line(ws, r, [a, b, c_, ''], wrap=2); r += 1
r += 1
note(ws, r, [
    '출처 — 세금계산서 총액표(26.09.01) · 2025 시딩 정산표 · 브랜드별 월별 대조표 · 글로벌 시딩 진행 라인 · IMD 성과 통계 v8(26.09.04) 원본데이터 2,050행',
    '개인정산 정보(주민등록번호·계좌·주소)는 원장 시트에만 있으며 본 보고서에는 포함하지 않음',
])
ws.freeze_panes = 'A4'

# ════════════════════════════════ A1_세금계산서
ws = sheet('A1_세금계산서', '[부속] 세금계산서 발행 내역', '프로젝트 · 국가 · 발행월별 · 2026년', 'FF9E9E9E')
MONS = ['2월', '3월', '4월', '5월', '6월', '7월', '8월']
widths(ws, [30, 14] + [13] * 7 + [15])
r = 4
header(ws, r, ['프로젝트', '국가'] + MONS + ['합계'], 20); r += 1
INV = [
    ('올영PB라이프 - 아이디얼포맨', '국내', {'2월': 71375000, '4월': 108000000, '5월': 74125000, '6월': 32125000, '7월': 16125000, '8월': 32000000}),
    ('올영PB라이프 - 식물나라', '국내', {'3월': 12500000, '4월': 13080000, '5월': 24303075, '6월': 19250000, '7월': 43408750, '8월': 57900000}),
    ('올영PB뷰티 - 브링그린', '국내', {'4월': 3250000, '6월': 15625000, '7월': 34875000, '8월': 32400000}),
    ('올영PB라이프 - 라운드어라운드', '국내', {'3월': 4675000, '6월': 62687500}),
    ('올영PB - 케어플러스', '국내', {'8월': 8998750}),
    ('올영PB - 필리밀리', '국내', {'7월': 3462500}),
    ('올영PB - 웨이크메이크', '일본(JP)', {'7월': 12937500, '8월': 9130000}),
    ('올영PB - 컬러그램', '일본(JP)', {'7월': 12000000, '8월': 8225000}),
    ('올영PB신성장 - 루테카', '국내+US 혼재', {'2월': 92750000, '3월': 29562500}),
    ('올영PB - 바이오힐보', '국내+JP 혼재', {'6월': 4187500, '7월': 2500000}),
]
s0 = r
for nm, ctry, vals in INV:
    line(ws, r, [nm, ctry] + [vals.get(m) for m in MONS] + [f'=SUM(C{r}:I{r})'],
         {k: MON for k in range(3, 11)}); r += 1
line(ws, r, ['2026년 합계', ''] + [f'=SUM({get_column_letter(c)}{s0}:{get_column_letter(c)}{r-1})' for c in range(3, 10)] +
     [f'=SUM(J{s0}:J{r-1})'], {k: MON for k in range(3, 11)}, fill=FT, bold=True)
inv_tot = r
r += 2
line(ws, r, ['2025.08  바이오힐보 US (68건)', '미국(US)', None, None, None, None, None, None, None, 30000000],
     {10: MON}); r += 1
line(ws, r, ['25.04~26.01  2025 시딩 계약 (272건)', '국내', None, None, None, None, None, None, None, 59830000],
     {10: MON}); r += 1
line(ws, r, ['세금계산서 총액 (발행분)', '', None, None, None, None, None, None, None,
             f'=J{inv_tot}+J{r-2}+J{r-1}'], {10: MON}, fill=FT, bold=True)
r += 2
note(ws, r, ['· 총액표에 국가 구분이 없어 글로벌 시딩 라인 자료로 귀속했다. 루테카·바이오힐보는 분리 근거가 없어 혼재로 둔다.'])
ws.freeze_panes = 'C5'

# ════════════════════════════════ A2_글로벌시딩
ws = sheet('A2_글로벌시딩', '[부속] 글로벌 시딩 진행 라인', '일본(JP) 99건 / 고정비 40,700,000원 · 미국(US) 68건 / 30,000,000원', 'FF9E9E9E')
widths(ws, [14, 42, 10, 9, 14, 9, 12, 12])
r = 4
header(ws, r, ['브랜드', '제품', '일정', '수량', '고정비', '건당', 'IMD 집계', ''], 20); r += 1
GJP = [
    ('웨이크메이크', '소블아 (7월 업로드 고정)', '7/20 ~', 3, 1245000, '내'),
    ('웨이크메이크', '소블아 (7월 업로드 자율)', '7/20 ~', 4, 1660000, '내'),
    ('웨이크메이크', '심리스파운데이션 (7월 업로드 자율)', '7/20 ~', 3, 1245000, '내'),
    ('웨이크메이크', '심리스 파운데이션 (8월 업로드)', '8/13 ~', 12, 4980000, '내'),
    ('웨이크메이크', '심리스 파운데이션', '8/13 ~', 7, 2905000, '내'),
    ('웨이크메이크', '소블아 오프체리', '8/13 ~', 10, 4150000, '내'),
    ('웨이크메이크', '소프트블러링아이팔레트', '9/14 ~', 21, 8904000, '외 (9월)'),
    ('컬러그램', '컬러커버틴트 + 탕후루 딥글레이즈 (7월 고정)', '7/20 ~', 3, 1245000, '내'),
    ('컬러그램', '컬러커버틴트 + 탕후루 딥글레이즈 (7월 자율)', '7/20 ~', 4, 1660000, '내'),
    ('컬러그램', '쉐딩스틱 (7월 업로드 자율)', '7/20 ~', 4, 1660000, '내'),
    ('컬러그램', '입체창조쉐딩스틱 (8월 업로드)', '8/13 ~', 4, 1660000, '내'),
    ('컬러그램', '컬러커버 틴트', '8/13 ~', 11, 4565000, '내'),
    ('컬러그램', '입체창조쉐딩스틱', '8/13 ~', 6, 2490000, '내'),
    ('바이오힐보', 'NAD 크림', '–', 7, 2331000, '내'),
]
s0 = r
for br, pr, dt, q, amt, win in GJP:
    line(ws, r, [f'{br}JP', pr, dt, q, amt, f'=IFERROR(E{r}/D{r},"")', win, ''],
         {4: CNT, 5: MON, 6: MON}, wrap=2); r += 1
line(ws, r, ['JP 합계', '', '', f'=SUM(D{s0}:D{r-1})', f'=SUM(E{s0}:E{r-1})',
             f'=IFERROR(E{r}/D{r},"")', '내 78 / 외 21', ''],
     {4: CNT, 5: MON, 6: MON}, fill=FT, bold=True); r += 1
line(ws, r, ['바이오힐보 US', '2025년 8월 올영PB 바이오힐보 US', '25.08', 68, 30000000,
             f'=IFERROR(E{r}/D{r},"")', '미수록', ''], {4: CNT, 5: MON, 6: MON}, wrap=2); r += 1
line(ws, r, ['글로벌 합계', '', '', f'=D{r-2}+D{r-1}', f'=E{r-2}+E{r-1}',
             f'=IFERROR(E{r}/D{r},"")', '', ''], {4: CNT, 5: MON, 6: MON}, fill=FT, bold=True)
r += 2
note(ws, r, [
    '· JP 건당 고정비는 415,000원이 기본 (소프트블러링아이팔레트 424,000 · 바이오힐보 NAD 333,000).',
    '· 고정비는 마크업·VAT 제외값이다. 세금계산서상 WM 22,067,500 + CG 20,225,000 = 42,292,500 중 고정비 29,465,000만 글로벌로 귀속했고, 차액 약 12.8백만(마크업·VAT 상당)은 KR에 남아 있다.',
    '· 9/14~ 업로드 21건은 26.09 진행분으로 세금계산서 미발행이다.',
])
ws.freeze_panes = 'A5'

# ════════════════════════════════ A3_2025계약·진행월
ws = sheet('A3_2025계약·진행월', '[부속] 2025 시딩 계약 · 브랜드별 진행월 대조',
           '2025 계약 272건 / 59,830,000원 · 진행월 대조표(아포맨·루테카·식물나라)', 'FF9E9E9E')
widths(ws, [18, 10, 10, 12, 10, 12, 12, 14, 13, 40])
r = 4
r = sec(ws, r, '2025 시딩 계약 (25.04 ~ 26.01 진행 · 전량 국내)')
header(ws, r, ['프로젝트', '진행월', '제리와콩나무', '원고료', 'BAT', '원고료',
               '솔루션', '총 원고료', '잔액', ''], 22); r += 1; s0 = r
for nm, mth, jq, jf, bq, bf, sol, tot, bal in [
    ('식물나라', '25.04', None, None, 51, 6790000, 510000, 7300000, 52700000),
    ('WM 산리오', '25.07', 23, 6785000, None, None, None, 6785000, 45915000),
    ('ID 올인원', '25.07', 12, 3540000, None, None, None, 3540000, 23060000),
    ('BOH 판테셀', '25.07', None, None, 9, 1800000, 90000, 1890000, 21170000),
    ('WM 코팅밤', '25.08', 37, 10915000, 40, 8000000, 400000, 19315000, 26600000),
    ('BOH 슈링크', '25.10', None, None, 39, 7800000, 390000, 8190000, 12980000),
    ('WM 실버크러시', '25.10', None, None, 17, 3400000, 170000, 3570000, 9410000),
    ('식물나라', '25.11', None, None, 39, 7800000, 390000, 8190000, 1220000),
    ('컬러그램', '26.01', None, None, 5, 1000000, 50000, 1050000, 170000),
]:
    line(ws, r, [nm, mth, jq, jf, bq, bf, sol, tot, bal, ''],
         {3: CNT, 4: MON, 5: CNT, 6: MON, 7: MON, 8: MON, 9: MON}); r += 1
line(ws, r, ['합계', ''] + [f'=SUM({get_column_letter(c)}{s0}:{get_column_letter(c)}{r-1})' for c in (3, 4, 5, 6, 7, 8)] +
     [170000, '272건 / 59,830,000원 · 잔여인원 0.8명 → 소진 완료'],
     {3: CNT, 4: MON, 5: CNT, 6: MON, 7: MON, 8: MON, 9: MON}, fill=FT, bold=True, wrap=10)
r += 2

r = sec(ws, r, '브랜드별 진행월 대조 (청구 대응 건수 근거)')
header(ws, r, ['브랜드', '진행월', '수량', '마크업 X', '마크업 O', '배수', '상태', '', '', '참고'], 22)
r += 1
PR = [
    ('아이디얼포맨', '26.1', 96, 25000000, 31250000, '확정', ''),
    ('아이디얼포맨', '26.2', 79, 32100000, 40125000, '확정', '프루아 69 + 선스틱 10'),
    ('아이디얼포맨', '26.3', 307, 82400000, 103000000, '확정', '지급대행 트루컴 5,000,000 별도'),
    ('아이디얼포맨', '26.4', 244, 59300000, 74125000, '확정', ''),
    ('아이디얼포맨', '26.5', 96, 25700000, 32125000, '확정', '올더베러 8건 포함'),
    ('아이디얼포맨', '26.6', 50, 12900000, 16125000, '확정', ''),
    ('아이디얼포맨 소계', '', 872, 237400000, 296750000, '', '세금계산서 2~7월 301,750,000과 일치'),
    ('루테카', '26.1', 88, 24200000, 30250000, '확정', ''),
    ('루테카', '26.2', 86, 16550000, 20687500, '확정', ''),
    ('루테카 소계', '', 174, 40750000, 50937500, '', '세금계산서 122,312,500과 71,375,000 차이 (A4 #1)'),
    ('식물나라', '26.2', 40, 10000000, 12500000, '확정', ''),
    ('식물나라', '26.3', 26, 6780000, 8475000, '확정', '지급대행 6,000,000 별도'),
    ('식물나라', '26.4', 54, 13998600, 17498250, '확정', '지급대행 3,200,000 별도'),
    ('식물나라', '26.5', 119, 17498250, 19250000, '확인 필요', '4월 마크업 이월 · 무가시딩 3,360,000'),
    ('식물나라', '26.6', 60, 15400000, 19250000, '확정', '7월 귀속'),
    ('식물나라', '26.7', 85, 24000000, 29750000, '미확정', '8월 귀속 · 마크업 25%→15%'),
    ('식물나라 소계', '', 384, 87676850, 106723250, '', '원본 파일 소계는 구버전(370건) — 재계산값'),
]
for nm, mth, q, x, o, st, memo in PR:
    is_sub = '소계' in nm
    line(ws, r, [nm, mth, q, x, o, f'=IFERROR(E{r}/D{r},"")', st, '', '', memo],
         {3: CNT, 4: MON, 5: MON, 6: RAT}, fill=(FA if is_sub else None), bold=is_sub, wrap=10); r += 1
r += 1
note(ws, r, ['· 마크업 O = 광고주 청구 기준(VAT 포함). 세금계산서 발행월 = 진행월 + 1개월, 지급대행은 별도 가산.'])
ws.freeze_panes = 'A5'

# ════════════════════════════════ A4_확인필요
ws = sheet('A4_확인필요', '[부속] 확인 필요 항목', '자료 간 불일치 · 미확정 사항 12건', 'FFC0504D')
widths(ws, [5, 24, 78, 16, 10])
r = 4
header(ws, r, ['#', '항목', '내용', '금액 영향', '우선순위'], 20); r += 1
GAPS = [
    (1, '루테카 발행액 미설명', '세금계산서 122,312,500 vs 진행 대조표 50,937,500 → 71,375,000 미설명. "올영PB신성장" 부문 타 항목 포함 여부 확인', 71375000, '높음'),
    (2, '청구 대응 진행 건수', '세금계산서에 수량 미기재 → 931,288,075원에 대응하는 진행 건수 확정 불가(추정 약 2,787건). 확정 시 건당 단가 확정 가능', None, '높음'),
    (3, 'IMD 국가 분류 오류', 'IMD는 일본 24건(WM 9·CG 8·BOH 7)만 분류하나 실제 99건. 54건(WM 30·CG 24)이 국내로 오분류', None, '높음'),
    (4, '바이오힐보 비용 귀속', 'IMD 202건·원고료 39,565,965인데 26년 발행액은 6,687,500뿐. 2025 계약 + US 30,000,000 + 브랜드사 직계약으로 분산 추정', None, '높음'),
    (5, '식물나라 소계 오류', '원본 대조표 소계(384건 / 68,177,200 / 95,145,250)가 구버전(370건) 값. 재계산 시 87,676,850 / 106,723,250', 11578000, '높음'),
    (6, '아포맨 누계 행 상충', '요약 행에 194,450,000 / 308,887,500 / 243,062,500 세 값 동시 기재. 월별 상세(237,400,000 / 296,750,000)가 정답', 12137500, '높음'),
    (7, 'JP 마크업률 미확정', 'JP 고정비 29,465,000만 글로벌로 귀속. 세금계산서 42,292,500과의 차액 약 12.8백만(마크업·VAT)은 KR에 잔류', 12827500, '중'),
    (8, '색조 3사 국내분 미발행', 'WM 국내 70건·CG 45건·FM 27건의 국내 시딩 비용이 26년 발행액에 없음. 하반기 계약으로 26.09 이후 발행 예정인지 확인', 34300000, '중'),
    (9, '비시딩 항목 전체 목록', '지급대행·모션그래픽·솔루션 32,400,000원은 확인분만. 전체 목록 확보 시 건당 단가 정밀화', None, '중'),
    (10, '2025년 상반기 데이터', 'IMD가 25.09.29부터라 25년 1~9월 업로드 데이터 없음. 2025 계약 272건 중 219건이 성과 집계 밖', None, '중'),
    (11, '루테카 US 수량', 'US TikTok 약 20건이 IMD·세금계산서 어느 쪽에서도 특정되지 않아 집계 제외', None, '낮음'),
    (12, '케어플러스 원고료', 'IMD 32건 원고료 0원(미입력). 세금계산서 발행액은 8,998,750', 8998750, '낮음'),
]
for n, item, desc, amt, pri in GAPS:
    line(ws, r, [n, item, desc, amt, pri], {1: CNT, 4: MON}, wrap=3, h=30); r += 1
r += 1
note(ws, r, ['· 금액 영향은 미설명·수정 필요 금액의 절대값이며 서로 중복될 수 있어 합산하지 않는다.'])
ws.freeze_panes = 'A5'

out = '/home/user/-_oy/reports/올리브영_시딩_성과보고서.xlsx'
wb.save(out); print('saved', out)
