# -*- coding: utf-8 -*-
"""올리브영 시딩 집계 (수량·비용) 워크북 — IMD 통계 워크북 서식 준용"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

KO   = '맑은 고딕'      # 한글 라벨 (원본 WenQuanYi Zen Hei → 윈도우/맥 호환 폰트로 대체)
NUM  = 'Arial'
GREEN= 'FF0F573E'
GRAY = 'FF808080'
NOTE = 'FF595959'
ALT  = 'FFFAFAFA'
TOT  = 'FFF2F2F2'
thin = Side(style='thin', color='FFD9D9D9')
BOX  = Border(left=thin, right=thin, top=thin, bottom=thin)

F_TITLE = Font(name=KO, size=15, bold=True, color=GREEN)
F_SUB   = Font(name=NUM, size=9, color=GRAY)
F_HDR   = Font(name=KO, size=9,  bold=True, color='FFFFFFFF')
F_LBL   = Font(name=KO, size=10, bold=True)
F_TXT   = Font(name=KO, size=10)
F_NUM   = Font(name=NUM, size=10)
F_NOTE  = Font(name=NUM, size=9, color=NOTE)
F_NOTEK = Font(name=KO,  size=9, color=NOTE)
FILL_H  = PatternFill('solid', fgColor=GREEN)
FILL_A  = PatternFill('solid', fgColor=ALT)
FILL_T  = PatternFill('solid', fgColor=TOT)

MONEY, MONEY_W, CNT, CNTW, RATIO, PCT = '#,##0', '#,##0\\원', '#,##0', '#,##0"건"', '0.00"배"', '0.0%'

wb = openpyxl.Workbook()
wb.remove(wb.active)
# LibreOffice 재계산 불가 환경 → 열 때 전체 재계산 강제
wb.calculation.fullCalcOnLoad = True


def head(ws, title, sub, headers, widths, hdr_row=4, heights=None):
    ws['A1'] = title; ws['A1'].font = F_TITLE; ws.row_dimensions[1].height = 18.75
    ws['A2'] = sub;   ws['A2'].font = F_SUB;   ws.row_dimensions[2].height = 14.25
    ws.row_dimensions[hdr_row].height = heights or 31.5
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=hdr_row, column=i, value=h)
        c.font, c.fill, c.border = F_HDR, FILL_H, BOX
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = ws.cell(row=hdr_row + 1, column=2)


def row(ws, r, vals, *, fill=None, bold_label=True, fmts=None, label_cols=1, wrap_from=None):
    for i, v in enumerate(vals, 1):
        c = ws.cell(row=r, column=i, value=v)
        c.border = BOX
        if i <= label_cols:
            c.font = F_LBL if bold_label else F_TXT
            c.alignment = Alignment(horizontal='left', vertical='center')
        elif isinstance(v, str) and not str(v).startswith('='):
            c.font = F_TXT
            c.alignment = Alignment(horizontal='left', vertical='center',
                                    wrap_text=bool(wrap_from and i >= wrap_from))
        else:
            c.font = F_NUM
            c.alignment = Alignment(horizontal='right', vertical='center')
            c.number_format = (fmts or {}).get(i, MONEY)
        if fill:
            c.fill = fill
    ws.row_dimensions[r].height = 14.25


def notes(ws, r, lines, col='A'):
    for i, t in enumerate(lines):
        c = ws[f'{col}{r + i}']; c.value = t
        c.font = F_NOTEK if any('가' <= ch <= '힣' for ch in t) else F_NOTE
    return r + len(lines)


# --- 1. 요약
ws = wb.create_sheet('요약')
head(ws, '올리브영 시딩 집계 — 전체 요약 (26.09.01 기준)',
     '전체 2,139건 / 940,192,075원  ·  국내(KR) 1,972건 869,492,075원 · 일본(JP) 99건 40,700,000원 · 미국(US) 68건 30,000,000원',
     ['국가', '수량', '수량 비중', '금액', '금액 비중', '건당', '진행월', '금액 기준', '내역'],
     [12, 10, 10, 16, 10, 12, 18, 22, 40], heights=22)
r = 5; s0 = r
fms = {2: CNT, 3: PCT, 4: MONEY_W, 5: PCT, 6: MONEY_W}
CTRY = [
    ('국내 (KR)', 1972, 869492075, '25.04 ~ 26.08', '세금계산서 발행액',
     '올영PB 11개 브랜드 · 2025 시딩 계약 272건 포함'),
    ('일본 (JP)', 99, 40700000, '26.06 ~ 26.09', '고정비 (마크업·VAT 제외)',
     '웨이크메이크 60 · 컬러그램 32 · 바이오힐보 7'),
    ('미국 (US)', 68, 30000000, '25.08', '세금계산서 발행액', '바이오힐보 US'),
]
for nm, q, a, mth, basis, memo in CTRY:
    row(ws, r, [nm, q, f'=B{r}/$B${s0 + len(CTRY)}', a, f'=D{r}/$D${s0 + len(CTRY)}',
                f'=IFERROR(D{r}/B{r},"")', mth, basis, memo], fmts=fms, wrap_from=9); r += 1
row(ws, r, ['전체 합계', f'=SUM(B{s0}:B{r - 1})', f'=B{r}/$B${r}',
            f'=SUM(D{s0}:D{r - 1})', f'=D{r}/$D${r}', f'=IFERROR(D{r}/B{r},"")',
            '25.04 ~ 26.09', '혼합', '글로벌 167건 = JP 99 + US 68'],
    fill=FILL_T, fmts=fms, wrap_from=9)
tot_c = r
r += 1
row(ws, r, ['(참고) 세금계산서 발행분', '', '', 931288075, '', '', '', '',
            '전체 940,192,075 − JP 미발행 8,904,000 (26.09 진행분 21건)'],
    fill=FILL_A, fmts={4: MONEY_W}, wrap_from=9)
r += 2

ws.cell(row=r, column=1, value='연월별 수량 · 금액').font = F_LBL
r += 1
for i2, h in enumerate(['연월', 'KR 수량', 'JP 수량', 'US 수량', '합계 수량',
                        'KR 금액', 'JP 금액', 'US 금액', '합계 금액'], 1):
    c = ws.cell(row=r, column=i2, value=h); c.font, c.fill, c.border = F_HDR, FILL_H, BOX
    c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
r += 1
ym0 = r
YM = [
    ('2025.04', None, None, None, 7300000,   None,     None),
    ('2025.07', None, None, None, 12215000,  None,     None),
    ('2025.08', None, None, 68,   19315000,  None,     30000000),
    ('2025.09', 1,    None, None, None,      None,     None),
    ('2025.10', 41,   None, None, 11760000,  None,     None),
    ('2025.11', 5,    None, None, 8190000,   None,     None),
    ('2025.12', 6,    None, None, None,      None,     None),
    ('2026.01', 131,  None, None, 1050000,   None,     None),
    ('2026.02', 267,  None, None, 164125000, None,     None),
    ('2026.03', 326,  None, None, 46737500,  None,     None),
    ('2026.04', 313,  None, None, 124330000, None,     None),
    ('2026.05', 290,  None, None, 98428075,  None,     None),
    ('2026.06', 214,  7,    None, 131544000, 2331000,  None),
    ('2026.07', 180,  21,   None, 116593750, 8715000,  None),
    ('2026.08', 198,  50,   None, 127903750, 20750000, None),
    ('2026.09', None, 21,   None, None,      8904000,  None),
]
fmm = {2: CNT, 3: CNT, 4: CNT, 5: CNT, 6: MONEY, 7: MONEY, 8: MONEY, 9: MONEY_W}
for ym, kq, jq, uq, ka, ja, ua in YM:
    row(ws, r, [ym, kq, jq, uq, f'=SUM(B{r}:D{r})', ka, ja, ua, f'=SUM(F{r}:H{r})'], fmts=fmm)
    r += 1
row(ws, r, ['합계'] + [f'=SUM({get_column_letter(c)}{ym0}:{get_column_letter(c)}{r - 1})'
                       for c in range(2, 10)], fill=FILL_T, fmts=fmm)
ym_tot = r
r += 2

ws.cell(row=r, column=1, value='브랜드별 수량').font = F_LBL
r += 1
for i2, h in enumerate(['브랜드', 'KR 수량', '글로벌 수량', '전체 수량', '글로벌 비중',
                        '글로벌 내역', '진행월', '', ''], 1):
    c = ws.cell(row=r, column=i2, value=h); c.font, c.fill, c.border = F_HDR, FILL_H, BOX
    c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
r += 1
bd0 = r
BD = [
    ('아이디얼포맨',     737, 0,  '', '25.09~26.07'),
    ('브링그린',         313, 0,  '', '26.02~26.08'),
    ('라운드어라운드',   281, 0,  '', '26.02~26.06'),
    ('바이오힐보',       195, 75, 'US 68 + JP 7', '25.08(US) · 25.10~26.08'),
    ('식물나라',         140, 0,  '', '25.11~26.05'),
    ('웨이크메이크',     70,  60, 'JP 60 (26.07 10 · 26.08 29 · 26.09 21)', '25.10 · 26.05~26.09'),
    ('루테카',           124, 0,  '', '26.01~26.03'),
    ('컬러그램',         45,  32, 'JP 32 (26.07 11 · 26.08 21)', '26.01 · 26.05~26.08'),
    ('케어플러스',       32,  0,  '', '26.07~26.08'),
    ('필리밀리',         27,  0,  '', '26.05~26.08'),
    ('올더베러',         8,   0,  '', '26.05~26.06'),
    ('딜라이트프로젝트', 0,   0,  '', '– (시딩 없음)'),
]
fmd = {2: CNT, 3: CNT, 4: CNT, 5: PCT}
for nm, kq, gq, gd, mth in BD:
    row(ws, r, [nm, kq, gq, f'=SUM(B{r}:C{r})', f'=IFERROR(C{r}/D{r},"")', gd, mth, None, None],
        fill=(FILL_A if gq else None), fmts=fmd, wrap_from=6); r += 1
row(ws, r, ['합계'] + [f'=SUM({get_column_letter(c)}{bd0}:{get_column_letter(c)}{r - 1})' for c in (2, 3, 4)] +
    [f'=IFERROR(C{r}/D{r},"")', 'JP 99 + US 68', '25.04~26.09', None, None],
    fill=FILL_T, fmts=fmd, wrap_from=6)
bd_tot = r
r += 2

ws.cell(row=r, column=1, value='검산').font = F_LBL
r += 1
for i2, h in enumerate(['항목', '값', '기대값', '판정', '설명'], 1):
    c = ws.cell(row=r, column=i2, value=h); c.font, c.fill, c.border = F_HDR, FILL_H, BOX
    c.alignment = Alignment(horizontal='center', vertical='center')
r += 1
for nm, f_, exp, desc in [
    ('전체 수량 (국가별)',   f'=B{tot_c}',  2139, 'KR 1,972 + JP 99 + US 68'),
    ('전체 수량 (연월별)',   f'=E{ym_tot}', 2139, '위 국가별 합계와 일치해야 함'),
    ('전체 수량 (브랜드별)', f'=D{bd_tot}', 2139, '위 국가별 합계와 일치해야 함'),
    ('전체 금액 (국가별)',   f'=D{tot_c}',  940192075, 'KR 869,492,075 + JP 40,700,000 + US 30,000,000'),
    ('전체 금액 (연월별)',   f'=I{ym_tot}', 940192075, '위 국가별 합계와 일치해야 함'),
    ('세금계산서 발행분',    f'=D{tot_c}-8904000', 931288075, '전체 − JP 미발행 8,904,000'),
]:
    row(ws, r, [nm, f_, exp, f'=IF(B{r}=C{r},"일치","불일치")', desc],
        fmts={2: (CNT if exp < 100000 else MONEY_W), 3: (CNT if exp < 100000 else MONEY_W)}, wrap_from=5)
    r += 1
r += 1
notes(ws, r, [
    '[국가 배분] 글로벌 = 글로벌 JP 시딩 진행 라인 99건 + 바이오힐보 US 68건. 그 외는 전부 KR.',
    '[금액 계산] KR 금액 = 해당 월 세금계산서 총액 − 글로벌 금액. 세금계산서 발행분 합계 931,288,075원은 그대로 유지된다.',
    '⚠️ JP 금액은 고정비(마크업·VAT 제외)다. 세금계산서 상 WM 22,067,500 + CG 20,225,000 = 42,292,500 중 고정비 29,465,000만 글로벌로 빠졌고, 차액 약 12.8백만(마크업·VAT 상당)은 KR에 남아 있다. JP 마크업률 확정 시 조정 필요.',
    '⚠️ 26.09 JP 진행분(웨이크메이크 21건 / 8,904,000)은 세금계산서 미발행이다. 발행 기준으로 볼 때는 931,288,075원을 쓴다.',
    '· 수량 축 = 업로드/진행월, 금액 축 = 세금계산서 발행월(2025 시딩분·JP는 진행월). 같은 행의 수량과 금액이 같은 캠페인을 가리키지 않을 수 있다.',
    '· IMD 원본데이터는 2,050건(국내 2,026 / 일본 24)이며, 여기에 JP 오분류 보정(54건)·JP 26.09분(21건)·US(68건)을 반영해 전체 2,139건으로 맞췄다.',
    '· 상세는 연월별_수량금액 / 브랜드별_수량정리 / 글로벌시딩_JP / 세금계산서_국가월별 시트 참조.',
])


# ─────────────────────────────────────────── 1-2. 연월별_수량금액
ws = wb.create_sheet('연월별_수량금액')
head(ws, '연월별 시딩 수량 · 금액 — KR / 글로벌',
     '글로벌 = 글로벌 JP 시딩 라인(99건) + 바이오힐보 US(68건) · 글로벌 금액은 해당 월 세금계산서 금액에서 차감해 KR과 분리 · KR = 세금계산서 월 총액 − 글로벌',
     ['연월', 'KR 수량', '글로벌 수량', '합계 수량', 'KR 금액', '글로벌 금액',
      '합계 금액 (세금계산서)', '글로벌 내역', '비고'],
     [10, 10, 11, 11, 16, 15, 17, 30, 40], heights=32)

# (연월, KR수량, 글로벌수량, 세금계산서 월총액, 글로벌금액, 글로벌내역, 비고)
YM = [
    ('2025.04', None, None, 7300000,   None,     '', '2025 시딩 진행월 · 식물나라 51건은 IMD 집계 밖'),
    ('2025.07', None, None, 12215000,  None,     '', '2025 시딩 진행월 · WM산리오+ID올인원+BOH판테셀 44건 IMD 밖'),
    ('2025.08', None, 68,   49315000,  30000000,  '바이오힐보 US 68건', '2025 시딩 WM코팅밤 19,315,000 + BOH US 30,000,000'),
    ('2025.09', 1,    0,    None,      None,      '', 'IMD 집계 시작 (09-29)'),
    ('2025.10', 41,   0,    11760000,  None,      '', '2025 시딩 BOH슈링크+WM실버크러시 진행월'),
    ('2025.11', 5,    0,    8190000,   None,      '', '2025 시딩 식물나라 진행월'),
    ('2025.12', 6,    0,    None,      None,      '', ''),
    ('2026.01', 131,  0,    1050000,   None,      '', '2025 시딩 컬러그램 진행월'),
    ('2026.02', 267,  0,    164125000, None,      '', '아포맨 71,375,000 + 루테카 92,750,000'),
    ('2026.03', 326,  0,    46737500,  None,      '', '식물나라 12,500,000 + 라운드어라운드 4,675,000 + 루테카 29,562,500'),
    ('2026.04', 313,  0,    124330000, None,      '', '아포맨 108,000,000 + 식물나라 13,080,000 + 브링그린 3,250,000'),
    ('2026.05', 290,  0,    98428075,  None,      '', '아포맨 74,125,000 + 식물나라 24,303,075'),
    ('2026.06', 214,  7,    133875000, 2331000,   '바이오힐보JP NAD 크림 7건', ''),
    ('2026.07', 180,  21,   125308750, 8715000,   'WM 10건 4,150,000 + CG 11건 4,565,000 (7/20~)', ''),
    ('2026.08', 198,  50,   148653750, 20750000,  'WM 29건 12,035,000 + CG 21건 8,715,000 (8/13~)', ''),
]
r = 5; s0 = r
fmy = {2: CNT, 3: CNT, 4: CNT, 5: MONEY_W, 6: MONEY_W, 7: MONEY_W}
for ym, kq, gq, tot, ga, gd, memo in YM:
    row(ws, r, [ym, kq, gq,
                (f'=SUM(B{r}:C{r})' if (kq is not None or gq is not None) else None),
                (f'=G{r}-F{r}' if tot else None), ga, tot, gd, memo],
        fmts=fmy, wrap_from=8)
    ws.row_dimensions[r].height = 26
    r += 1
row(ws, r, ['합계 (세금계산서 발행분)'] +
    [f'=SUM({get_column_letter(c)}{s0}:{get_column_letter(c)}{r - 1})' for c in (2, 3, 4, 5, 6, 7)] +
    ['', 'KR 869,492,075 + 글로벌 61,796,000 = 931,288,075 (세금계산서 총액 일치)'],
    fill=FILL_T, fmts=fmy, wrap_from=9)
sum_row = r
ws.row_dimensions[r].height = 26
r += 1
row(ws, r, ['2026.09', 0, 21, f'=SUM(B{r}:C{r})', None, 8904000, 8904000,
            'WM 소프트블러링아이팔레트 21건 (9/14~)', '세금계산서 미발행 — 위 합계에 미포함'],
    fill=FILL_A, fmts=fmy, wrap_from=8)
sep_row = r
ws.row_dimensions[r].height = 26
r += 1
row(ws, r, ['총계 (진행 기준)'] +
    [f'={get_column_letter(c)}{sum_row}+{get_column_letter(c)}{sep_row}' for c in (2, 3, 4, 5, 6, 7)] +
    ['', 'KR 869,492,075 + 글로벌 70,700,000 = 940,192,075'],
    fill=FILL_T, fmts=fmy, wrap_from=9)
tot_ym = r
ws.row_dimensions[r].height = 26
r += 2

ws.cell(row=r, column=1, value='검산').font = F_LBL
r += 1
for i2, h in enumerate(['항목', '값', '기대값', '판정', '설명'], 1):
    c = ws.cell(row=r, column=i2, value=h); c.font, c.fill, c.border = F_HDR, FILL_H, BOX
    c.alignment = Alignment(horizontal='center', vertical='center')
r += 1
for nm, formula, exp, desc in [
    ('세금계산서 발행분 총액', f'=G{sum_row}', 931288075, '세금계산서 총액표 전체 총계와 일치해야 함'),
    ('전체 수량 (진행 기준)',  f'=D{tot_ym}', 2139, 'KR 1,972 + 글로벌 167 (JP 99 + US 68) — 요약·브랜드별 시트와 일치'),
    ('글로벌 금액 (진행 기준)', f'=F{tot_ym}', 70700000, '바이오힐보 US 30,000,000 + 글로벌 JP 고정비 40,700,000'),
    ('글로벌 수량 (진행 기준)', f'=C{tot_ym}', 167, 'BOH US 68 + JP 99'),
]:
    row(ws, r, [nm, formula, exp, f'=IF(B{r}=C{r},"일치","불일치")', desc],
        fmts={2: MONEY_W if exp > 10000 else CNT, 3: MONEY_W if exp > 10000 else CNT}, wrap_from=5)
    ws.row_dimensions[r].height = 22
    r += 1
r += 1

ws.cell(row=r, column=1, value='2025 시딩 계약 수량 (진행월 · IMD 축과 중복 가능 → 위 합계 수량 미포함)').font = F_LBL
r += 1
for i2, h in enumerate(['진행월', '수량', '금액', '', '', '', '', '내역'], 1):
    c = ws.cell(row=r, column=i2, value=h); c.font, c.fill, c.border = F_HDR, FILL_H, BOX
    c.alignment = Alignment(horizontal='center', vertical='center')
r += 1
c25 = r
for ym, q, a, memo in [
    ('2025.04', 51, 7300000,  '식물나라'),
    ('2025.07', 44, 12215000, 'WM 산리오 23 + ID 올인원 12 + BOH 판테셀 9'),
    ('2025.08', 77, 19315000, 'WM 코팅밤 (제리와콩나무 37 + BAT 40)'),
    ('2025.10', 56, 11760000, 'BOH 슈링크 39 + WM 실버크러시 17'),
    ('2025.11', 39, 8190000,  '식물나라'),
    ('2026.01', 5,  1050000,  '컬러그램'),
]:
    row(ws, r, [ym, q, a, None, None, None, None, memo], fmts={2: CNT, 3: MONEY_W}, wrap_from=8); r += 1
row(ws, r, ['합계', f'=SUM(B{c25}:B{r - 1})', f'=SUM(C{c25}:C{r - 1})',
            None, None, None, None, '272건 / 59,830,000원 · 25.04~25.08 진행분 172건은 IMD 집계 밖'],
    fill=FILL_T, fmts={2: CNT, 3: MONEY_W}, wrap_from=8)
r += 2
notes(ws, r, [
    '[계산 방식] 글로벌 금액 = 글로벌 JP 시딩 라인 고정비(업로드 일자 기준 월 배분) + 바이오힐보 US. KR 금액 = 해당 월 세금계산서 총액 − 글로벌 금액.',
    '· 글로벌 JP 라인은 전량 26년 JP 시딩 건이므로, 세금계산서 상 웨이크메이크·컬러그램·바이오힐보 발행액에서 이 금액만큼이 글로벌로 빠진다.',
    '· ⚠️ JP 라인 금액은 고정비(마크업·VAT 제외)다. 따라서 JP분의 마크업·VAT는 아직 KR 금액에 남아 있다 (26.07~08 잔여 WM·CG 발행액 약 21.5백만). JP 마크업률이 확정되면 그만큼 글로벌로 더 옮겨야 한다.',
    '· 26.09 진행분(WM 21건 / 8,904,000)은 세금계산서 미발행이라 발행분 합계에서 제외하고 진행 기준 총계에만 포함했다.',
    '· 축이 다르다: 수량은 업로드/진행월, 금액은 세금계산서 발행월(2025 시딩분은 진행월).',
    '· 2026.07~08 KR 수량은 IMD 국내(184 / 248)에서 글로벌 오분류분(7월 4건 · 8월 50건)을 차감한 값이다.',
    '· 출처: 세금계산서 총액표, 2025 시딩 정산표, 글로벌 JP 시딩 진행 라인 자료, IMD 원본데이터 2,050행.',
])

# ─────────────────────────────────────────── 2. 세금계산서_국가월별
ws = wb.create_sheet('세금계산서_국가월별')
MON = ['1월', '2월', '3월', '4월', '5월', '6월', '7월', '8월']
head(ws, '세금계산서 기준 — 프로젝트 · 국가 · 월별 (2026년 발행월)',
     '금액의 정답 기준 · 국내/일본/혼재를 같은 월이라도 행 분리 · 단위: 원',
     ['프로젝트', '국가'] + MON + ['합계'], [34, 16] + [13] * 8 + [15], heights=20)

inv = [
    ('국내', [
        ('올영PB라이프 - 아이디얼포맨 (퍼펙트올인원)', {'2월': 71375000, '4월': 108000000, '5월': 74125000, '6월': 32125000, '7월': 16125000, '8월': 32000000}),
        ('올영PB라이프 - 식물나라',                   {'3월': 12500000, '4월': 13080000, '5월': 24303075, '6월': 19250000, '7월': 43408750, '8월': 57900000}),
        ('올영PB뷰티 - 브링그린',                     {'4월': 3250000, '6월': 15625000, '7월': 34875000, '8월': 32400000}),
        ('올영PB라이프 - 라운드어라운드',              {'3월': 4675000, '6월': 62687500}),
        ('올영PB - 케어플러스',                       {'8월': 8998750}),
        ('올영PB - 필리밀리',                         {'7월': 3462500}),
    ]),
    ('일본(JP) 캠페인', [
        ('올영PB - 웨이크메이크', {'7월': 12937500, '8월': 9130000}),
        ('올영PB - 컬러그램',    {'7월': 12000000, '8월': 8225000}),
    ]),
    ('국내+US 혼재', [('올영PB신성장 - 루테카', {'2월': 92750000, '3월': 29562500})]),
    ('국내+JP 혼재', [('올영PB - 바이오힐보',   {'6월': 4187500, '7월': 2500000})]),
]
r = 5
sub_rows, blocks = [], []
for ctry, items in inv:
    start = r
    for name, vals in items:
        row(ws, r, [name, ctry] + [vals.get(m) for m in MON] +
            [f'=SUM(C{r}:J{r})'], fmts={k: MONEY for k in range(3, 11)} | {11: MONEY_W}); r += 1
    row(ws, r, [f'{ctry} 소계', ''] +
        [f'={get_column_letter(c)}{start}+' .join([]) or f'=SUM({get_column_letter(c)}{start}:{get_column_letter(c)}{r - 1})' for c in range(3, 11)] +
        [f'=SUM(K{start}:K{r - 1})'], fill=FILL_A, fmts={k: MONEY for k in range(3, 11)} | {11: MONEY_W})
    sub_rows.append(r); blocks.append((start, r - 1)); r += 1
row(ws, r, ['2026년 월 합계', ''] +
    [f'=' + '+'.join(f'{get_column_letter(c)}{s}' for s in sub_rows) for c in range(3, 11)] +
    ['=' + '+'.join(f'K{s}' for s in sub_rows)], fill=FILL_T,
    fmts={k: MONEY for k in range(3, 11)} | {11: MONEY_W})
r += 2
r = notes(ws, r, [
    '국가 귀속 근거 — 세금계산서 총액표 자체에는 국가 구분이 없어 아래 근거로 분리했다.',
    '· 웨이크메이크·컬러그램 → 일본(JP) 확정: 글로벌 JP 시딩 진행 라인 자료(WM 60건 · CG 32건)로 확인됨. IMD 일본 건 총원고료(WM 23,004,315 / CG 20,297,925)도 각 발행액(22,067,500 / 20,225,000)과 근접 일치.',
    '   반대로 두 브랜드 국내 원고료(WM 15,850,000 / CG 12,150,000)는 발행액에 없음 → 색조 3사 국내분은 하반기 별도 계약으로 26.9월 이후 발행 추정.',
    '· 연월별_수량금액 시트에서는 JP 라인 고정비(31,796,000)만 글로벌로 차감했다. 잔여분(마크업·VAT 상당)은 KR에 남아 있다.',
    '· 루테카 → 국내+US 혼재: 계약이 KR 80건 + US TikTok 20건 구성. 국가별 분리 근거 없음.',
    '· 바이오힐보 → 국내+JP 혼재: 26.6월 일본 7건(1,665,965)이 6월 발행액 4,187,500에 포함.',
    '· 아포맨·식물나라·브링그린·라운드어라운드·필리밀리·케어플러스 → 국내 전량 (견적서·IMD 모두 국내 IG 단일).',
    '· 1월 발행분 없음 (루테카·아포맨 모두 2월부터 발행).',
])
r += 1
ws.cell(row=r, column=1, value='별건 — 미국(US)').font = F_LBL
r += 1
for i, h in enumerate(['프로젝트', '국가', '진행월', '수량', '금액'], 1):
    c = ws.cell(row=r, column=i, value=h); c.font, c.fill, c.border = F_HDR, FILL_H, BOX
    c.alignment = Alignment(horizontal='center', vertical='center')
r += 1
row(ws, r, ['2025년 8월 올영PB 바이오힐보 US', '미국', '25.8', 68, 30000000],
    fmts={4: CNTW, 5: MONEY_W})

# ─────────────────────────────────────────── 3. 글로벌시딩_JP
ws = wb.create_sheet('글로벌시딩_JP')
head(ws, '글로벌(JP) 시딩 — 진행 라인별 수량 · 고정비',
     '전량 인스타그램 · 마이크로 · 고정비 기준(마크업·VAT 제외) · 이 수량을 전체·브랜드별 수량에서 차감해 국내 순수량을 산출',
     ['브랜드', '제품', '채널', '등급', '일정', '비용 구분', '고정비', '수량', '건당', 'IMD 집계기간'],
     [14, 42, 11, 9, 9, 9, 14, 9, 12, 13])
GJP = [
    ('웨이크메이크', '소블아 (7월 업로드 고정)',                          '7/20 ~', 1245000, 3,  '내'),
    ('웨이크메이크', '소블아 (7월 업로드 자율)',                          '7/20 ~', 1660000, 4,  '내'),
    ('웨이크메이크', '심리스파운데이션 (7월 업로드 자율)',                 '7/20 ~', 1245000, 3,  '내'),
    ('웨이크메이크', '심리스 파운데이션 (8월 업로드)',                     '8/13 ~', 4980000, 12, '내'),
    ('웨이크메이크', '심리스 파운데이션',                                  '8/13 ~', 2905000, 7,  '내'),
    ('웨이크메이크', '소블아 오프체리',                                    '8/13 ~', 4150000, 10, '내'),
    ('웨이크메이크', '소프트블러링아이팔레트',                             '9/14 ~', 8904000, 21, '외'),
    ('컬러그램',   '컬러커버틴트 + 탕후루 딥글레이즈 (7월 업로드 고정)', '7/20 ~', 1245000, 3,  '내'),
    ('컬러그램',   '컬러커버틴트 + 탕후루 딥글레이즈 (7월 업로드 자율)', '7/20 ~', 1660000, 4,  '내'),
    ('컬러그램',   '쉐딩스틱 (7월 업로드 자율)',                          '7/20 ~', 1660000, 4,  '내'),
    ('컬러그램',   '입체창조쉐딩스틱 (8월 업로드)',                        '8/13 ~', 1660000, 4,  '내'),
    ('컬러그램',   '컬러커버 틴트',                                        '8/13 ~', 4565000, 11, '내'),
    ('컬러그램',   '입체창조쉐딩스틱',                                     '8/13 ~', 2490000, 6,  '내'),
    ('바이오힐보', 'NAD 크림',                                             '',       2331000, 7,  '내'),
]
r = 5
fmg = {7: MONEY_W, 8: CNT, 9: MONEY_W}
sub = []
for brand in ['웨이크메이크', '컬러그램', '바이오힐보']:
    start = r
    for br, pr, dt, amt, q, win in GJP:
        if br != brand:
            continue
        row(ws, r, [f'올리브영PB_{br}JP', pr, '인스타그램', '마이크로', dt, '고정비', amt, q,
                    f'=IFERROR(G{r}/H{r},"")', '내' if win == '내' else '외 (9월)'],
            fmts=fmg, wrap_from=2); r += 1
    row(ws, r, [f'{brand}JP 소계', '', '', '', '', '',
                f'=SUM(G{start}:G{r - 1})', f'=SUM(H{start}:H{r - 1})',
                f'=IFERROR(G{r}/H{r},"")', ''], fill=FILL_A, fmts=fmg)
    sub.append(r); r += 1
row(ws, r, ['글로벌(JP) 총계', '', '', '', '', '',
            '=' + '+'.join(f'G{x}' for x in sub), '=' + '+'.join(f'H{x}' for x in sub),
            f'=IFERROR(G{r}/H{r},"")', ''], fill=FILL_T, fmts=fmg)
tot_g = r
r += 2
ws.cell(row=r, column=1, value='IMD 집계기간(≤26.08.30) 구분').font = F_LBL
r += 1
for i, h in enumerate(['구분', '웨이크메이크', '컬러그램', '바이오힐보', '합계 수량', '합계 고정비'], 1):
    c = ws.cell(row=r, column=i, value=h); c.font, c.fill, c.border = F_HDR, FILL_H, BOX
    c.alignment = Alignment(horizontal='center', vertical='center')
r += 1
ws.row_dimensions[r].height = 14.25
win_first = r
row(ws, r, ['기간 내 (7~8월 업로드)', 39, 32, 7, f'=SUM(B{r}:D{r})', 31796000],
    fmts={2: CNT, 3: CNT, 4: CNT, 5: CNT, 6: MONEY_W}); r += 1
row(ws, r, ['기간 외 (9/14~ 업로드)', 21, 0, 0, f'=SUM(B{r}:D{r})', 8904000],
    fmts={2: CNT, 3: CNT, 4: CNT, 5: CNT, 6: MONEY_W}); r += 1
row(ws, r, ['브랜드 합계', f'=SUM(B{win_first}:B{r - 1})', f'=SUM(C{win_first}:C{r - 1})',
            f'=SUM(D{win_first}:D{r - 1})', f'=SUM(E{win_first}:E{r - 1})',
            f'=SUM(F{win_first}:F{r - 1})'], fill=FILL_T,
    fmts={2: CNT, 3: CNT, 4: CNT, 5: CNT, 6: MONEY_W})
r += 2
notes(ws, r, [
    '· 건당 고정비는 415,000원이 기본이며, 소프트블러링아이팔레트만 424,000원 · 바이오힐보 NAD 크림만 333,000원이다.',
    '· 고정비는 마크업·VAT 제외 값이다. 세금계산서 상 웨이크메이크 22,067,500 · 컬러그램 20,225,000과 직접 비교하면 안 된다.',
    '· 9/14~ 업로드분(웨이크메이크 소프트블러링아이팔레트 21건 / 8,904,000원)은 IMD 집계 종료(26.08.30) 이후라 2,050건 안에 없다.',
    '· 바이오힐보 NAD 크림 7건은 일정 미기재이며 IMD 일본 26.06 7건과 수량이 일치한다.',
    '· 출처: 글로벌 시딩 진행 라인 자료 (담당자 제공).',
])

# ─────────────────────────────────────────── 4. 2025시딩_계약
ws = wb.create_sheet('2025시딩_계약')
head(ws, '2025년 시딩 계약 — 25.4월 ~ 26.1월 진행 (전량 국내)',
     '총 원고료 = 원고료 + 솔루션 이용료 · 계약 잔액 170,000 (잔여인원 0.8명) → 사실상 소진 완료',
     ['프로젝트', '국가', '진행월', '제리와콩나무 건수', '제리와콩나무 원고료',
      'BAT 건수', 'BAT 원고료', '솔루션 이용료', '총 원고료', '잔액'],
     [18, 10, 10, 14, 15, 10, 14, 13, 15, 15])
d25 = [
    ('식물나라',      '25.4',  None, None,     51, 6790000,  510000, 7300000,  52700000),
    ('WM 산리오',     '25.7',  23,   6785000,  None, None,   None,   6785000,  45915000),
    ('ID 올인원',     '25.7',  12,   3540000,  None, None,   None,   3540000,  23060000),
    ('BOH 판테셀',    '25.7',  None, None,     9,  1800000,  90000,  1890000,  21170000),
    ('WM 코팅밤',     '25.8',  37,   10915000, 40, 8000000,  400000, 19315000, 26600000),
    ('BOH 슈링크',    '25.10', None, None,     39, 7800000,  390000, 8190000,  12980000),
    ('WM 실버크러시', '25.10', None, None,     17, 3400000,  170000, 3570000,  9410000),
    ('식물나라',      '25.11', None, None,     39, 7800000,  390000, 8190000,  1220000),
    ('컬러그램',      '26.1',  None, None,     5,  1000000,  50000,  1050000,  170000),
]
r = 5; s = r
fm = {4: CNT, 5: MONEY, 6: CNT, 7: MONEY, 8: MONEY, 9: MONEY_W, 10: MONEY_W}
for nm, mth, jq, jf, bq, bf, sol, tot, bal in d25:
    row(ws, r, [nm, '국내', mth, jq, jf, bq, bf, sol, tot, bal], fmts=fm); r += 1
row(ws, r, ['합계', '', ''] + [f'=SUM({get_column_letter(c)}{s}:{get_column_letter(c)}{r - 1})' for c in range(4, 10)] + [170000],
    fill=FILL_T, fmts=fm)
r += 1
row(ws, r, ['검산 (원고료+솔루션)', '', '', '', '', '', f'=E{r - 1}+G{r - 1}+H{r - 1}', '', f'=I{r - 1}', ''],
    fill=FILL_A, fmts={7: MONEY_W, 9: MONEY_W})
r += 2
notes(ws, r, [
    '· 272건 / 59,830,000원 (제리와콩나무 72건 + BAT 200건) · 솔루션(스프레이 AI) 이용료 총 2,000,000원 포함.',
    '· G열+H열 합계 = I열 합계 = 59,830,000원 으로 검산 일치.',
    '· 별건: 바이오힐보 US 68건 / 30,000,000원 (25.8월) — 세금계산서_국가월별 시트 하단 참조.',
    '· 출처: 2025년 시딩 정산표 (담당자 제공).',
])

# ─────────────────────────────────────────── 5. 진행월_대조표
ws = wb.create_sheet('진행월_대조표')
head(ws, '진행월 기준 대조표 — 아이디얼포맨 · 루테카 · 식물나라 (전량 국내)',
     '수량·마크업 전/후·확정 여부 · 마크업 25%(기본) → 식물나라 26.7월부터 15% 변동',
     ['브랜드', '국가', '진행월', '수량', '마크업 X', '마크업 O', '배수', '상태', '참고'],
     [14, 14, 10, 10, 15, 15, 8, 12, 58])
r = 5


def block(brand, ctry, rows_, sub_note):
    global r
    s = r
    for mth, q, x, o, st, memo in rows_:
        row(ws, r, [brand, ctry, mth, q, x, o,
                    (f'=IFERROR(F{r}/E{r},"")' if x else None), st, memo],
            fmts={4: CNT, 5: MONEY_W, 6: MONEY_W, 7: RATIO}, wrap_from=9); r += 1
    row(ws, r, [f'{brand} 소계', '', ''] +
        [f'=SUM({get_column_letter(c)}{s}:{get_column_letter(c)}{r - 1})' for c in range(4, 7)] +
        [f'=IFERROR(F{r}/E{r},"")', '', sub_note], fill=FILL_A,
        fmts={4: CNT, 5: MONEY_W, 6: MONEY_W, 7: RATIO}, wrap_from=9)
    r += 2


block('아이디얼포맨', '국내', [
    ('26.1', 96,  25000000, 31250000,  '확정', ''),
    ('26.2', 79,  32100000, 40125000,  '확정', '2월 프루아 69건 + 선스틱(런칭) 10건'),
    ('26.3', 307, 82400000, 103000000, '확정', '2월 프루아 잔여 112건 추가 / 지급대행(트루컴 1건 5,000,000) 제외'),
    ('26.4', 244, 59300000, 74125000,  '확정', ''),
    ('26.5', 96,  25700000, 32125000,  '확정', '올더베러 8건 포함'),
    ('26.6', 50,  12900000, 16125000,  '확정', ''),
], '업데이트 완료 — 원본 파일의 "1~6월 누계" 요약 행(194,450,000 / 308,887,500 / 243,062,500)은 상충. 이 월별 상세가 세금계산서와 검산되는 값이다.')

block('루테카', '국내+US 혼재', [
    ('26.1', 88, 24200000, 30250000, '확정', ''),
    ('26.2', 86, 16550000, 20687500, '확정', ''),
], '변동 X — 세금계산서 발행액 122,312,500과 71,375,000 차이. "올영PB신성장" 부문 타 항목 동반 청구 추정 → 분해 필요.')

block('식물나라', '국내', [
    ('26.2', 40,  10000000, 12500000, '확정',      ''),
    ('26.3', 26,  6780000,  8475000,  '확정',      '지급대행 6,000,000 별도'),
    ('26.4', 54,  13998600, 17498250, '확정',      '지급대행 3,200,000 별도'),
    ('26.5', 119, 17498250, 19250000, '확인 필요', '4월 마크업 1,864,500 이월 / 무가시딩 업체지급 3,360,000 (3,200,000 + 수수료 5%)'),
    ('26.6', 60,  15400000, 19250000, '확정',      '7월 귀속'),
    ('26.7', 85,  24000000, 29750000, '미확정',    '8월 귀속 (안투안 55건 + 준석 30건) · 마크업 25%→15% · 광고주 청구 전'),
    ('26.8', None, None,    None,     '미확정',    '전략1 총 67건 예상 (릴스 47 + 피드 20, 나노) · 예산 12,161,250 · 마크업 25%→15%'),
], '재계산값 — 원본 파일 소계(384건 / 68,177,200 / 95,145,250)는 갱신 누락(구버전 370건 합계와 일치). 지급대행 9,200,000 별도.')

notes(ws, r, [
    '· 마크업 X = 크리에이터 원고료 / 마크업 O = 광고주 청구 기준(마크업 포함) · 배수 = 마크업O ÷ 마크업X.',
    '· 식물나라 26.8월은 금액 미확정이라 수량·금액 공란 (예산 12,161,250원 별도 표기).',
    '· 출처: 올영PB 브랜드별 월별 대조표 (담당자 제공, Table-1-2 기준).',
])

# --- 7. 브랜드별_수량정리
ws = wb.create_sheet('브랜드별_수량정리')
head(ws, '브랜드별 수량 — KR / 글로벌 (전체 2,139건 기준)',
     '전체 2,139건 = KR 1,972 + 글로벌 167 (JP 99 + US 68) · IMD 원본 2,050건에 JP 오분류 보정·JP 26.09분·US를 반영한 값',
     ['브랜드', 'KR 수량', '글로벌 수량', '전체 수량', '글로벌 비중',
      'IMD 원본 건수', '글로벌 내역', '진행월', '비고'],
     [16, 11, 12, 11, 10, 12, 30, 22, 34])
bd = [
    ('아이디얼포맨',     737, 0,  737, '', '25.09~26.07', ''),
    ('브링그린',         313, 0,  313, '', '26.02~26.08', ''),
    ('라운드어라운드',   281, 0,  281, '', '26.02~26.06', ''),
    ('바이오힐보',       195, 75, 202, 'US 68 + JP 7', '25.08(US) · 25.10~26.08', 'US 68건은 IMD 미포함'),
    ('식물나라',         140, 0,  140, '', '25.11~26.05', ''),
    ('루테카',           124, 0,  124, '', '26.01~26.03', 'US TT 약 20건은 수량 특정 불가로 제외'),
    ('웨이크메이크',     70,  60, 109, 'JP 60 (26.07 10 · 26.08 29 · 26.09 21)', '25.10 · 26.05~26.09',
     'IMD는 JP 9건만 분류 · 26.09분 21건 IMD 미포함'),
    ('컬러그램',         45,  32, 77,  'JP 32 (26.07 11 · 26.08 21)', '26.01 · 26.05~26.08',
     'IMD는 JP 8건만 분류'),
    ('케어플러스',       32,  0,  32,  '', '26.07~26.08', ''),
    ('필리밀리',         27,  0,  27,  '', '26.05~26.08', ''),
    ('올더베러',         8,   0,  8,   '', '26.05~26.06', ''),
    ('딜라이트프로젝트', 0,   0,  0,   '', '–', '시딩 데이터 없음'),
]
r = 5; s0 = r
fmb = {2: CNT, 3: CNT, 4: CNT, 5: PCT, 6: CNT}
for nm, kq, gq, imd, gd, mth, memo in bd:
    row(ws, r, [nm, kq, gq, f'=SUM(B{r}:C{r})', f'=IFERROR(C{r}/D{r},"")', imd, gd, mth, memo],
        fill=(FILL_A if gq else None), fmts=fmb, wrap_from=7); r += 1
row(ws, r, ['합계'] + [f'=SUM({get_column_letter(c)}{s0}:{get_column_letter(c)}{r - 1})' for c in (2, 3, 4)] +
    [f'=IFERROR(C{r}/D{r},"")', f'=SUM(F{s0}:F{r - 1})', 'JP 99 + US 68', '25.04~26.09',
     'KR 1,972 + 글로벌 167 = 2,139'], fill=FILL_T, fmts=fmb, wrap_from=7)
tot_bd = r
r += 2
ws.cell(row=r, column=1, value='검산').font = F_LBL
r += 1
for i2, h in enumerate(['항목', '값', '기대값', '판정', '설명'], 1):
    c = ws.cell(row=r, column=i2, value=h); c.font, c.fill, c.border = F_HDR, FILL_H, BOX
    c.alignment = Alignment(horizontal='center', vertical='center')
r += 1
for nm, f_, exp, desc in [
    ('전체 수량', f'=D{tot_bd}', 2139, 'KR 1,972 + 글로벌 167'),
    ('IMD 원본 건수', f'=F{tot_bd}', 2050, 'IMD 원본데이터 행수'),
    ('IMD → 전체 차이', f'=D{tot_bd}-F{tot_bd}', 89, 'JP 26.09분 21 + US 68'),
]:
    row(ws, r, [nm, f_, exp, f'=IF(B{r}=C{r},"일치","불일치")', desc], fmts={2: CNT, 3: CNT}, wrap_from=5); r += 1
r += 1
notes(ws, r, [
    '· IMD 원본데이터는 일본을 24건(WM 9 · CG 8 · BOH 7)만 잡고 있으나 글로벌 자료 기준 JP는 99건이다. IMD 집계기간 내 78건 중 54건(WM 30 · CG 24)이 국내로 오분류돼 있어 KR 수량에서 차감했다.',
    '· IMD 2,050건 → 전체 2,139건 차이 89건 = JP 26.09 진행분 21건(집계기간 밖) + 바이오힐보 US 68건(IMD 미수록).',
    '· 루테카 US TikTok 약 20건은 IMD·세금계산서 어느 쪽에서도 수량이 특정되지 않아 집계에서 제외했다 (불일치_확인필요 참조).',
])


# ─────────────────────────────────────────── 7. 귀속검증
ws = wb.create_sheet('귀속검증')
head(ws, '귀속 규칙 검증 — 세금계산서 발행월 = 진행월 + 1개월 (지급대행 별도 가산)',
     '아이디얼포맨(국내)으로 완전 검증됨 · 단위: 원',
     ['검증 항목', '진행 기준 금액', '발행 기준 금액', '차이', '판정', '설명'],
     [30, 17, 17, 14, 10, 50])
r = 5
ver = [
    ('발행 2월 ↔ 진행 1월 + 2월', 31250000 + 40125000, 71375000, '진행 1월 31,250,000 + 2월 40,125,000이 2월에 일괄 발행'),
    ('발행 4월 ↔ 진행 3월 + 지급대행', 103000000 + 5000000, 108000000, '진행 3월 103,000,000 + 지급대행(트루컴 1건) 5,000,000'),
    ('발행 5월 ↔ 진행 4월', 74125000, 74125000, '1개월 시차 그대로'),
    ('발행 6월 ↔ 진행 5월', 32125000, 32125000, '1개월 시차 그대로'),
    ('발행 7월 ↔ 진행 6월', 16125000, 16125000, '1개월 시차 그대로'),
]
s = r
for nm, a, b, memo in ver:
    row(ws, r, [nm, a, b, f'=C{r}-B{r}', f'=IF(D{r}=0,"일치","불일치")', memo],
        fmts={2: MONEY_W, 3: MONEY_W, 4: MONEY_W}, wrap_from=6); r += 1
row(ws, r, ['아포맨 총액 검증', 296750000 + 5000000 + 32000000, 333750000,
            f'=C{r}-B{r}', f'=IF(D{r}=0,"일치","불일치")',
            '진행 1~6월 296,750,000 + 지급대행 5,000,000 + 8월 발행 32,000,000 (7월 진행분)'],
    fill=FILL_T, fmts={2: MONEY_W, 3: MONEY_W, 4: MONEY_W}, wrap_from=6)
r += 2
notes(ws, r, [
    '· 대조표의 "식물나라 6월 → 7월귀속", "7월 → 8월 귀속" 주석도 같은 규칙이다.',
    '· 이 규칙 덕분에 진행월 기준 수량과 발행월 기준 금액을 서로 환산할 수 있다.',
    '· 단, 식물나라는 발행 7월 43,408,750(선케어 86건)·8월 57,900,000(산리오 콜라보)에 대응하는 진행월 행이 대조표에 없어 완전 환산이 불가하다.',
])

# ─────────────────────────────────────────── 8. 불일치_확인필요
ws = wb.create_sheet('불일치_확인필요')
head(ws, '소스 간 불일치 — 확인 필요 항목',
     '세금계산서(발행월) · 대조표(진행월) · IMD(업로드일) 3자 교차 검증 결과',
     ['#', '항목', '내용', '금액 영향', '우선순위'],
     [6, 22, 88, 16, 11])
r = 5
gaps = [
    (1, '루테카 발행액', '세금계산서 122,312,500 vs 진행 대조표 50,937,500 vs 거래명세서 49,225,000 → 최대 71,375,000 미설명. "올영PB신성장" 부문 타 항목 포함 여부 확인', 71375000, '높음'),
    (2, '식물나라 소계 갱신 누락', '대조표 소계(68,177,200 / 95,145,250)가 구버전(370건) 값. 월별 상세 재계산 시 87,676,850 / 106,723,250', 11578000, '높음'),
    (3, '아포맨 누계 행 상충', '요약 행에 194,450,000 / 308,887,500 / 243,062,500 세 값이 동시 기재. 월별 상세 237,400,000 / 296,750,000로 통일 필요', 12137500, '높음'),
    (4, '바이오힐보 비용 귀속', 'IMD 202건 · 원고료 39,565,965인데 26년 발행액은 6,687,500뿐. 2025 시딩 계약(판테셀·슈링크) + US 30,000,000 + 브랜드사 직계약으로 분산 추정', None, '높음'),
    (5, '색조 3사 국내분 미발행', 'WM 국내 100건 · CG 국내 69건 · FM 27건의 국내 시딩 비용이 26년 발행액에 없음. 하반기 계약으로 26.9월 이후 발행 예정인지 확인', 34300000, '중'),
    (6, '진행 vs 업로드 건수 갭', '아포맨 872건(진행) vs 734건(IMD, 올더베러 8 포함) / 루테카 174 vs 124 / 식물나라 384 vs 140 → 미업로드 + IMD 반영 지연 분해 필요', None, '중'),
    (7, '미국 건 IMD 누락', 'IMD 원본데이터에 US 행이 0건. 바이오힐보 US 68건, 루테카 US TT 약 20건이 집계 밖', 30000000, '중'),
    (8, '2025년 1~9월 데이터', 'IMD가 25-09-29부터라 25년 상반기 업로드 데이터 없음. 2025 시딩 계약 272건과 대조하면 실제 270여건', None, '중'),
    (9, '케어플러스 원고료 미입력', 'IMD 32건 원고료 0원. 세금계산서 발행액은 8,998,750', 8998750, '낮음'),
    (11, 'IMD 국가 분류 오류', 'IMD 원본은 일본을 24건(WM 9 · CG 8 · BOH 7)으로 잡지만 글로벌 자료 기준 78건(WM 39 · CG 32 · BOH 7). 54건(WM 30 · CG 24)이 국내로 오분류 → 원본데이터 국가 컬럼 보정 필요', None, '높음'),
    (10, '담당자 제공 수량 차이', '스프레이 AI 시트 2,170여건(2025 270 + 2026 1,900) vs 구두 "총 2,100여건" → 70건 차이. 제안서 인용 전 확정', None, '낮음'),
]
for n, item, desc, amt, pri in gaps:
    row(ws, r, [n, item, desc, amt, pri], fmts={1: CNT, 4: MONEY_W}, wrap_from=3)
    ws.row_dimensions[r].height = 30
    r += 1
r += 1
notes(ws, r, ['· 금액 영향은 미설명·수정 필요 금액의 절대값이며, 서로 중복될 수 있어 합산하지 않는다.'])

# ─────────────────────────────────────────── 9. 단가_참고
ws = wb.create_sheet('단가_참고')
head(ws, '단가 · 계약 조건 참고',
     '견적서 · 예산안 · 대조표에서 확인된 값',
     ['구분', '값', '출처 / 비고'], [30, 40, 60], heights=20)
r = 5
for k, v, src in [
    ('마크업', '25% (기본) → 식물나라 26.7월부터 15%', '대조표 참고 컬럼 · 견적서'),
    ('지급대행 수수료', '5~10%', '식물나라 무가시딩 업체지급 5% · 3Q 예산안'),
    ('마크업 배수 (실측)', '1.25 (아포맨 일부 1.59 · 식물나라 7월 2.49)', '대조표 배수 컬럼'),
    ('국내 나노', '225,000원/건', '3Q 예산안'),
    ('국내 마이크로', '350,000원/건 (기본 250,000 + 인센티브 100,000)', '3Q 예산안 · 2Q 견적서'),
    ('국내 매크로', '600,000~800,000원/건', '2Q 견적서 · 3Q 예산안'),
    ('일본 고정', '424,000원/건', '색조 8월 견적서'),
    ('일본 자율', '530,000원/건', '색조 8월 견적서'),
    ('색조 국내 건당', '418,500원 (고료 400,000 + 인센티브 기댓값 15,000 + 배송 3,500)', '색조 하반기 운영 시트'),
    ('색조 일본 건당', '424,000원 (고료 400,000 + 인센티브 기댓값 15,000 + 배송 9,000)', '색조 하반기 운영 시트'),
    ('트위터(X) 나노', '500,000원/건', '3Q 예산안'),
    ('UGC 인센티브', '100,000원/건 (10만뷰 달성 시, 달성률 15% 가정)', '색조 하반기 운영 시트'),
    ('스프레이 AI 솔루션', '26.11~27.10 무상 지원 (연 50,400,000원 상당)', '스프레이_AI지원 시트 — 비용 집계 제외'),
    ('2025 시딩 솔루션 이용료', '총 2,000,000원 (총 원고료에 포함)', '2025 시딩 정산표'),
]:
    row(ws, r, [k, v, src], wrap_from=3); r += 1
r += 2
ws.cell(row=r, column=1, value='출처 문서').font = F_LBL
r += 1
for i, h in enumerate(['문서', '용도', ''], 1):
    c = ws.cell(row=r, column=i, value=h); c.font, c.fill, c.border = F_HDR, FILL_H, BOX
    c.alignment = Alignment(horizontal='center', vertical='center')
r += 1
for doc, use in [
    ('올영PB UGC 세금계산서 기준 총액표 (26.09.01)', '금액 기준 — 전체 총계 931,288,075원'),
    ('2025년 시딩 정산표 (제리와콩나무 / BAT 분할)', '2025 시딩 272건 / 59,830,000원 · 잔액 170,000원'),
    ('올영PB 브랜드별 월별 대조표 (아포맨·루테카·식물나라)', '진행월 기준 수량 · 마크업 전/후 · 확정 여부'),
    ('올리브영 PB UGC 시딩 성과 통계 (IMD 26.09.01) — 원본데이터 2,050행', '국가(국내·일본)별 건수 · 성과'),
    ('동 파일 — 브랜드별_종합 / 연월별_시딩건수 / 스프레이_AI지원', '브랜드별 건수·ROAS · 담당자 제공 운용 실적'),
]:
    row(ws, r, [doc, use, ''], bold_label=False, wrap_from=1); r += 1
r += 1
notes(ws, r, ['· 개인정산 정보(주민등록번호·계좌·주소 등)는 원장 시트에만 있고 이 워크북에는 포함하지 않았다.'])

out = '/home/user/-_oy/reports/올리브영_시딩_집계_260901.xlsx'
wb.save(out)
print('saved', out)
