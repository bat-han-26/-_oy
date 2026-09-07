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

# ════════════════════════════════════════════════════════════════════
#  데이터는 reports/data.py (SSOT) 에서만 읽는다. 이 파일에 숫자 하드코딩 금지.
# ════════════════════════════════════════════════════════════════════
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import data as D
D.verify()
I = D.IDX
BASIS = {'F': '원고료', 'B': '청구'}
f2 = '#,##0.00'

# ════════════════════════════════════════════════════════════════════ 01_요약
ws = sheet('01_요약', '올리브영 PB 시딩 성과 요약',
           f'26.09.07 기준 · 팩트 집계 · 시딩 {D.QTY:,}건 · 청구 {D.INV_TOT:,}원 · 조회수 {D.VIEW:,} · 매출 {D.REV:,}원', DEEP)
widths(ws, [26, 17, 4, 26, 17, 4, 24, 17, 34])
r = 4
KPI = [
    ('■ 규모', [('시딩 건수', D.QTY, CNT, '건'), ('국내 (KR)', D.KR_Q, CNT, '건'),
                ('글로벌 (JP+US)', D.GL_Q, CNT, '건'), ('브랜드 수', 11, CNT, '개'),
                ('운영 기간', '25.04 ~ 26.09', None, '')]),
    ('■ 비용', [('세금계산서 발행', D.INV_TOT, MON, '원'), ('건당 청구액', D.INV_TOT // D.QTY, MON, '원'),
                ('시딩 비용 확인분', D.COST, MON, '원'), ('청구액 대비 비중', D.COST / D.INV_TOT, PCT, ''),
                ('퍼포먼스 광고비', D.AD, MON, '원')]),
    ('■ 성과', [('총 조회수', D.VIEW, CNT, '뷰'), ('총 참여 (좋아요+댓글)', D.ENGA, CNT, ''),
                ('총 전환 매출', D.REV, MON, '원'), ('파트너십 ROAS', D.PA['rev'] / D.PA['ad'], RAT, ''),
                ('CPV / CPE', f"{D.INV_TOT/D.VIEW:.1f}원 / {D.INV_TOT/D.ENGA:,.0f}원", None, '')]),
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

r = sec(ws, r, '전체 단가  —  세금계산서 발행액 기준')
header(ws, r, ['항목', '금액 / 값', '수량', '조회수', 'CPV', '참여', 'CPE', 'ER', '건당 청구액'], 24)
r += 1; t0 = r
line(ws, r, ['세금계산서 발행 확정', D.INV_TOT, D.QTY, D.VIEW, f'=B{r}/D{r}', D.ENGA,
             f'=B{r}/F{r}', f'=F{r}/D{r}', f'=B{r}/C{r}'],
     {2: MON, 3: CNT, 4: CNT, 5: f2 + '"원"', 6: CNT, 7: MON, 8: PCT, 9: MON},
     fill=FT, bold=True, wrap=9); r += 1
line(ws, r, ['  (구성) 시딩 비용 확인분', D.COST, None, None, f'=B{r}/$D${t0}', None, f'=B{r}/$F${t0}',
             f'=B{r}/$B${t0}', f'=B{r}/$C${t0}'],
     {2: MON, 5: f2 + '"원"', 7: MON, 8: PCT, 9: MON}, wrap=9); r += 1
line(ws, r, ['      · 원고료 기준 (국내 + WM·CG JP)', D.COST_F, None, None, None, None, None,
             f'=B{r}/$B${t0}', '마크업·VAT 제외'], {2: MON, 8: PCT}, wrap=9); r += 1
line(ws, r, ['      · 청구 기준 (BOH JP + BOH US)', D.COST_B, None, None, None, None, None,
             f'=B{r}/$B${t0}', 'BOH JP 건당 333,000(마크업 포함) + BOH US 틱톡 30,000,000'],
     {2: MON, 8: PCT}, wrap=9); r += 1
line(ws, r, ['  (구성) 그 외', f'=B{t0}-B{t0+1}', None, None, None, None, None,
             f'=B{r}/$B${t0}', '마크업·VAT·솔루션·지급대행·비시딩 항목·성과 미집계 건 대응분'],
     {2: MON, 8: PCT}, wrap=9)
r += 2

r = sec(ws, r, '국가별  —  세금계산서 발행액 기준')
header(ws, r, ['국가', '시딩 건수', '비중', '발행액', '비중', '건당', '조회수', '진행월', '근거'], 24)
r += 1; s0 = r
for nm, q, a, vw, mth, memo in [
    ('국내 (KR)', D.KR_Q, D.INV_KR, 52_536_293, '25.04 ~ 26.09', '총 발행액 − JP·US 명시 라인'),
    ('일본 (JP)', 99, D.INV_JP, 407_439, '26.06 ~ 26.09', '세금계산서 WM 22,067,500 + CG 20,225,000'),
    ('미국 (US)', 68, D.INV_2025_US, 4_739_117, '25.08', '2025 시딩 정산 59,830,000원 내 틱톡 시딩분'),
]:
    line(ws, r, [nm, q, f'=B{r}/$B${s0+3}', a, f'=D{r}/$D${s0+3}', f'=IFERROR(D{r}/B{r},"")',
                 vw, mth, memo], {2: CNT, 3: PCT, 4: MON, 5: PCT, 6: MON, 7: CNT}, wrap=9); r += 1
line(ws, r, ['합계', f'=SUM(B{s0}:B{r-1})', f'=B{r}/$B${r}', f'=SUM(D{s0}:D{r-1})',
             f'=D{r}/$D${r}', f'=IFERROR(D{r}/B{r},"")', f'=SUM(G{s0}:G{r-1})', '25.04 ~ 26.09',
             '조회수는 IMD 국가 태그 기준'],
     {2: CNT, 3: PCT, 4: MON, 5: PCT, 6: MON, 7: CNT}, fill=FT, bold=True, wrap=9)
r += 2
note(ws, r, [
    '[팩트 원칙] 모든 수치는 원천 자료 기재값이다. 추정·안분·보정값은 사용하지 않았다 (05_기준정의 ②).',
    f'[모수] 수량·비용 = 전체 {D.QTY:,}건 / 조회수·참여 = IMD 성과 집계분 {D.PERF:,}건 · 미집계 {D.UNPERF}건(US 32 + JP 26.09 21)',
    f'[참여수] 실측 좋아요 {D.LIKE:,} + 댓글 {D.CMNT:,} = {D.ENGA:,}. 좋아요 데이터가 없는 {D.LIKE_BLANK}건(조회수 {D.LIKE_BLANK_V:,})은 추정하지 않았다 → 실제 참여·ER은 이 값 이상 (A4-3)',
    f'[금액] 세금계산서 발행 {D.INV_TOT:,}원 = 26.02~08 확정 {D.INV_2026:,} + 2025 시딩 정산 {D.INV_2025:,}. + JP 미발행 {D.JP_UNBILLED:,} → 진행 기준 {D.INV_TOT+D.JP_UNBILLED:,}',
    f'[시딩 비용] 확인분 {D.COST:,}원 = 원고료 기준 {D.COST_F:,} + 청구 기준 {D.COST_B:,}. 기준이 다른 값이 섞여 있어 02_브랜드별 기준 열에 건별 표기했다.',
    '[상세] 브랜드별 → 02 · 연월별 → 03 · 광고 성과 → 04 · 산정 기준 → 05 · 원자료 → A1~A4',
])
ws.freeze_panes = 'A4'

# ════════════════════════════════════════════════════════════════════ 02_브랜드별
ws = sheet('02_브랜드별', '브랜드별 시딩 · 비용 · 성과',
           '팩트 집계 · 수량은 라인 자료, 성과는 IMD 실측, 비용은 원고료/청구 기준을 열에 표기', DEEP)
widths(ws, [17, 7, 8, 8, 8, 13, 11, 9, 10, 8, 14, 8, 8, 14, 15, 8, 30])
r = 4
group(ws, r, [('', 1), ('시딩 수량', 4), ('콘텐츠 성과 (실측)', 5), ('시딩 비용', 3), ('퍼포먼스 광고', 3), ('', 1)])
r += 1
header(ws, r, ['브랜드', 'KR', '글로벌', '전체', '성과\n집계', '조회수', '좋아요', '댓글',
               '참여', 'ER', '금액', '기준', 'CPV', '광고비', '전환 매출', 'ROAS', '비고'])
r += 1; s0 = r
fm = {2: CNT, 3: CNT, 4: CNT, 5: CNT, 6: CNT, 7: CNT, 8: CNT, 9: CNT, 10: PCT,
      11: MON, 13: f2, 14: MON, 15: MON, 16: RAT}
for b in D.BRANDS:
    kr, gl, ad = b[I['KR']], b[I['GL']], b[I['AD']]
    line(ws, r, [b[I['NAME']], kr, gl, kr + gl, b[I['PF']], b[I['V']], b[I['L']], b[I['C']],
                 f'=G{r}+H{r}', f'=I{r}/F{r}', b[I['COST']], BASIS[b[I['BASIS']]], f'=K{r}/F{r}',
                 ad, b[I['REV']], (f'=O{r}/N{r}' if ad else '–'), b[I['MEMO']]],
         fm, fill=(FA if gl else None), wrap=17); r += 1
line(ws, r, ['합계'] + [f'=SUM({get_column_letter(c)}{s0}:{get_column_letter(c)}{r-1})' for c in (2, 3, 4, 5, 6, 7, 8, 9)] +
     [f'=I{r}/F{r}', f'=SUM(K{s0}:K{r-1})', '혼합', f'=K{r}/F{r}'] +
     [f'=SUM({get_column_letter(c)}{s0}:{get_column_letter(c)}{r-1})' for c in (14, 15)] +
     [f'=O{r}/N{r}', f'수량 {D.QTY:,} / 성과 집계 {D.PERF:,}'], fm, fill=FT, bold=True, wrap=17)
r += 2
note(ws, r, [
    '· 수량 = 글로벌 시딩 라인 자료(JP 99 · US 68) + IMD 집계분. 성과 집계는 IMD에 성과가 기재된 건수다.',
    '· 시딩 비용 기준 — [원고료] 마크업·VAT 제외. 국내는 IMD 행 단위 기재값, JP는 확정 단가 건당 415,000, 케어플러스는 시딩현황 3,700,000.',
    '· 시딩 비용 기준 — [청구] 마크업 포함. 바이오힐보 JP 건당 333,000(7건 2,331,000) · 바이오힐보 US 틱톡 30,000,000(2025 시딩 정산 59,830,000원 내).',
    f'· 합계 {D.COST:,}원 = 원고료 기준 {D.COST_F:,} + 청구 기준 {D.COST_B:,}. 기준이 섞여 있어 소계 CPV {D.COST/D.VIEW:.2f}원은 참고값이다.',
    '· IMD의 JP 원고료 기재값은 라인 총액을 행마다 반복 기재한 값이라 사용하지 않았다 (A2 각주 · A4-10).',
    '· 브랜드별 세금계산서 귀속은 발행 내역에 브랜드 미귀속액 70,431,189원이 있어 산출하지 않았다 (A1). 전체 청구 기준 단가는 01_요약.',
    f'· 참여는 실측값이다. 좋아요 데이터가 없는 {D.LIKE_BLANK}건은 추정하지 않았으므로 실제 ER은 표기값 이상이다 (A4-3).',
])
ws.freeze_panes = 'B6'

# ════════════════════════════════════════════════════════════════════ 03_연월별
ws = sheet('03_연월별', '연월별 시딩 · 성과 · 발행액',
           '수량·성과 = 업로드/진행월 (IMD 실측) / 발행액 = 세금계산서 발행월 확정값', DEEP)
widths(ws, [13, 11, 14, 15, 16, 16, 9, 36])
r = 4
header(ws, r, ['연월', '시딩 건수', '조회수', '광고비', '전환 매출', '세금계산서 발행', 'ROAS', '비고'], 24)
r += 1; s0 = r
fmy = {2: CNT, 3: CNT, 4: MON, 5: MON, 6: MON, 7: RAT}
for ym, q, vw, ad, rev, memo in D.YM:
    line(ws, r, [ym, q, vw, ad, rev, D.INV_MONTHLY.get(ym), (f'=E{r}/D{r}' if ad else '–'), memo],
         fmy, wrap=8); r += 1
line(ws, r, ['2025 시딩 정산', None, None, None, None, D.INV_2025, '–',
             f'2025 시딩 전체 운영비 · 그중 바이오힐보 US 틱톡 {D.INV_2025_US:,}'], fmy, wrap=8); r += 1
line(ws, r, ['합계'] + [f'=SUM({get_column_letter(c)}{s0}:{get_column_letter(c)}{r-1})' for c in (2, 3, 4, 5, 6)] +
     [f'=E{r}/D{r}', f'발행 {D.INV_TOT:,} + JP 미발행 {D.JP_UNBILLED:,} = 진행 기준 {D.INV_TOT+D.JP_UNBILLED:,}'],
     fmy, fill=FT, bold=True, wrap=8)
r += 2
note(ws, r, [
    '· 축이 다르다 — 수량·성과는 업로드/진행월, 발행액은 세금계산서 발행월(진행월 + 1개월 · 지급대행 별도 가산). 같은 행의 수량과 금액이 같은 캠페인이 아니다.',
    '· 26.02~08 발행액은 26.09.07 확정 총액표 기재값이다. 국가별·브랜드별로 나누지 않았다 — 총액표에 구분이 없다.',
    '· 25.08 US 68건 중 32건, 26.09 JP 21건은 성과 데이터가 없다 (조회수·광고비·매출 공란).',
    f'· 26.09 JP 21건은 세금계산서 미발행이다 (원고료 {D.JP_UNBILLED:,} = 21건 × 415,000).',
])
ws.freeze_panes = 'B5'

# ════════════════════════════════════════════════════════════════════ 04_광고성과
ws = sheet('04_광고성과', '퍼포먼스 광고 성과 · 콘텐츠 분포',
           f'IMD 실측 · 조회수·매출 분포 모수 {D.PERF:,}건 · ROAS 분포 모수 {D.DIST_ROAS_BASE}건(광고비 10만원 이상 집행분)', DEEP)
widths(ws, [22, 11, 16, 17, 11, 10, 13, 32])
r = 4
r = sec(ws, r, '광고 유형별')
header(ws, r, ['구분', '소재 건수', '광고비 소진', '전환 매출', '구매건수', 'ROAS', 'CPA', '비고'], 22)
r += 1; s0 = r
for nm, d, memo in [('파트너십 광고', D.PA, '국내 1,871건에서 발생 (집행 8개 브랜드)'), ('영상 가공', D.VE, '')]:
    line(ws, r, [nm, d['n'], d['ad'], d['rev'], d['buy'], f'=IFERROR(D{r}/C{r},"")', f'=IFERROR(C{r}/E{r},"")', memo],
         {2: CNT, 3: MON, 4: MON, 5: CNT, 6: RAT, 7: MON}, wrap=8); r += 1
line(ws, r, ['합계'] + [f'=SUM({get_column_letter(c)}{s0}:{get_column_letter(c)}{r-1})' for c in (2, 3, 4, 5)] +
     [f'=IFERROR(D{r}/C{r},"")', f'=IFERROR(C{r}/E{r},"")',
      f"시딩 1건당 소재 활용 {(D.PA['n']+D.VE['n'])/D.QTY:.2f}건"],
     {2: CNT, 3: MON, 4: MON, 5: CNT, 6: RAT, 7: MON}, fill=FT, bold=True, wrap=8)
r += 2
for title, hdr, rows_, unit, base, memo in [
    ('조회수 분포', ['구간', '건수', '비중', '구간 조회수 합'], D.DIST_VIEW, CNT, D.PERF,
     '상위 271건(12.8%)이 전체 조회수의 72.9%'),
    ('ROAS 분포', ['구간', '건수', '비중', '구간 매출 합'], D.DIST_ROAS, MON, D.DIST_ROAS_BASE,
     f'광고비 10만원 이상 집행 {D.DIST_ROAS_BASE}건 기준 (비중 = ÷{D.DIST_ROAS_BASE})'),
    ('매출액 분포', ['구간', '건수', '비중', '구간 매출 합'], D.DIST_REV, MON, D.PERF,
     '상위 2건(0.1%)이 전체 매출의 8.6% · 상위 44건(2.1%)이 57.8%'),
]:
    r = sec(ws, r, title + f'  —  {memo}')
    header(ws, r, hdr + ['', '', '', ''], 20); r += 1
    for nm, c_, v in rows_:
        line(ws, r, [nm, c_, f'=IFERROR(B{r}/{base},"")', v, None, None, None, None],
             {2: CNT, 3: PCT, 4: unit}); r += 1
    r += 1
r = sec(ws, r, '우수 콘텐츠')
header(ws, r, ['조회수 TOP 5', '브랜드', '업로드일', '조회수',
               'ROAS TOP 5', '브랜드', '광고비', '매출 / ROAS'], 20)
r += 1
for a, b in zip(D.TOP_VIEW, D.TOP_ROAS):
    line(ws, r, [a[0], a[1], a[2], a[3], b[0], b[1], b[2], b[3]], {4: CNT, 7: MON}); r += 1
r += 1
note(ws, r, [
    f"· 파트너십 광고 ROAS {D.PA['rev']/D.PA['ad']*100:.1f}%, 영상가공 포함 시 {D.REV/D.AD*100:.1f}%. 총 구매 {D.BUY:,}건 · CPA {D.AD/D.BUY:,.0f}원.",
    '· 파트너십 성과는 국내 1,871건에서 발생. 글로벌 77건과 미집행 3개 브랜드(루테카·케어플러스·올더베러) 164건은 광고비·매출이 0이라 제외해도 ROAS는 동일하다.',
    f'· ROAS TOP 5는 광고비 10만원 이상 집행 {D.DIST_ROAS_BASE}건 내 순위다. 소액 집행 건은 배수가 과대해져 모수에서 제외했다.',
    '· 루테카(124건)·케어플러스(32건)·올더베러(8건)·라운드어라운드(281건 중 1건만)는 광고 집행이 거의 없어 소재 활용 여지가 남아 있다.',
])
ws.freeze_panes = 'A4'

# ════════════════════════════════════════════════════════════════════ 05_기준정의
ws = sheet('05_기준정의', '집계 기준 · 팩트 / 비팩트 구분',
           '이 보고서에 들어간 값과, 근거가 없어 넣지 않은 값을 구분해 적었다', DEEP)
widths(ws, [26, 24, 68, 28])
r = 4
r = sec(ws, r, '① 팩트 — 원천 자료 기재값만 사용')
header(ws, r, ['항목', '값', '원천 자료', '비고'], 22); r += 1
for a, b, c_, d in [
    ('세금계산서 발행 총액', f'{D.INV_TOT:,}원', f'월별 확정 총액표(26.09.07) {D.INV_2026:,} + 2025 시딩 정산 {D.INV_2025:,}',
     f'2025 시딩은 전량 {D.INV_2025:,}원 내에서 운영 · 그중 BOH US 틱톡 {D.INV_2025_US:,}'),
    ('JP 발행액', f'{D.INV_JP:,}원', '세금계산서 JP 명시 라인 — 웨이크메이크 22,067,500 + 컬러그램 20,225,000', '바이오힐보 JP는 별도 라인 없음'),
    ('시딩 수량', f'{D.QTY:,}건', f'글로벌 시딩 라인(JP 99 · US 68) + IMD 집계 2,076건', f'KR {D.KR_Q:,} + 글로벌 {D.GL_Q}'),
    ('조회수 · 좋아요 · 댓글', f'{D.VIEW:,} / {D.LIKE:,} / {D.CMNT:,}',
     'IMD 성과 통계 v8(26.09.04) + 컬러그램(26.09.07) + 바이오힐보 US(26.09.07)', f'성과 집계분 {D.PERF:,}건 실측'),
    ('시딩 비용 (원고료 기준)', f'{D.COST_F:,}원', 'IMD 행 단위 기재값 + JP 확정 단가 415,000 + 케어플러스 시딩현황 3,700,000', '마크업·VAT 제외'),
    ('시딩 비용 (청구 기준)', f'{D.COST_B:,}원', '바이오힐보 JP 7건 × 333,000 = 2,331,000 / 바이오힐보 US 틱톡 30,000,000', '마크업 포함'),
    ('광고 성과', f'{D.AD:,} → {D.REV:,}원', 'IMD 파트너십광고 · 영상가공 컬럼 실측', f"소재 {D.PA['n']+D.VE['n']}건 · 구매 {D.BUY:,}건"),
]:
    line(ws, r, [a, b, c_, d], wrap=3, h=28); r += 1
r += 1
r = sec(ws, r, '② 넣지 않은 값 — 근거가 없어 추정을 배제한 항목')
header(ws, r, ['항목', '왜 넣지 않았나', '영향', '대안'], 22); r += 1
for a, b, c_, d in [
    (f'좋아요 결측 {D.LIKE_BLANK}건', f'IMD에 좋아요 데이터가 없다(조회수 {D.LIKE_BLANK_V:,}). 대부분 댓글은 있어 실제 0이 아닌 수집 누락으로 보인다',
     f'참여 {D.ENGA:,} · ER {D.ENGA/D.VIEW*100:.2f}% · CPE {D.INV_TOT/D.ENGA:,.0f}원이 실제보다 낮게(불리하게) 나온다', 'IMD에서 좋아요를 채우면 해소 (A4-3)'),
    ('브랜드별 세금계산서 귀속', '발행 내역 대비 브랜드 미귀속액 70,431,189원이 있고, 총액표에 브랜드·국가 구분이 없다',
     '브랜드별 청구 기준 CPV·CPE를 낼 수 없다', '브랜드별은 시딩 비용(원고료/청구) 기준 CPV로 표기'),
    ('바이오힐보 US 원고료', 'IMD에 원고료 기재가 없다. 청구 30,000,000원만 확인된다',
     'US 68건은 청구 기준으로만 표기된다', '기준 열에 "청구"로 명시'),
    ('청구 대응 진행 건수', '세금계산서에 수량 컬럼이 없다. 2025 계약 219건 등 성과 집계 밖 건이 청구에 포함돼 있다',
     f'건당 청구액 {D.INV_TOT//D.QTY:,}원(= {D.INV_TOT:,} ÷ {D.QTY:,})은 실제 시딩 1건 단가보다 높다', '대조표로 확인되는 1,770건 실단가 290,531원 병기 (③)'),
    ('국가별 조회수', 'IMD 국가 태그가 부정확하다(JP 일부가 국내로 분류)', '국가별 조회수는 태그 기준값이다', '수량·금액은 라인 자료 기준으로 별도 관리 (A4-5)'),
]:
    line(ws, r, [a, b, c_, d], wrap=2, h=34); r += 1
r += 1
r = sec(ws, r, '③ 건당 단가 두 가지')
header(ws, r, ['구분', '건수', '금액', '건당'], 20); r += 1
line(ws, r, ['전체 청구 기준', D.QTY, D.INV_TOT, f'=C{r}/B{r}'], {2: CNT, 3: MON, 4: MON}, fill=FT, bold=True); r += 1
s0 = r
for nm, q, a in [('아이디얼포맨 1~6월 진행', 872, 296_750_000), ('루테카 1~2월 진행', 174, 50_937_500),
                 ('식물나라 2~7월 진행', 384, 106_723_250), ('2025 시딩 정산 (국내 272 + US 68)', 340, D.INV_2025)]:
    line(ws, r, [nm, q, a, f'=IFERROR(C{r}/B{r},"")'], {2: CNT, 3: MON, 4: MON}); r += 1
line(ws, r, ['정합 표본 가중평균', f'=SUM(B{s0}:B{r-1})', f'=SUM(C{s0}:C{r-1})', f'=IFERROR(C{r}/B{r},"")'],
     {2: CNT, 3: MON, 4: MON}, fill=FT, bold=True)
r += 1
r = note(ws, r, [
    '· 정합 표본 = 진행 대조표와 세금계산서에서 건수·금액이 함께 확정되는 항목만 (전체 청구액의 57%).',
    '· 검증 1 — 아포맨 대조표 296,750,000 + 지급대행 5,000,000 = 세금계산서 발행 2~7월 301,750,000 (일치).',
    '· 검증 2 — 아포맨 원고료 272,248원 × 마크업 1.25 = 340,310원 = 위 표 실제 청구 단가 (일치).',
    '· 두 값의 차이는 전체 청구액에 비시딩 항목(지급대행 25,200,000 · 모션그래픽 5,000,000 · 솔루션 2,000,000 · 풀필먼트 200,000)과 성과 집계 밖 건의 청구분이 포함돼 있기 때문이다.',
])
r += 1
r = sec(ws, r, '④ 산정식')
header(ws, r, ['용어', '정의', '값', ''], 20); r += 1
for a, b, c_ in [('CPV', f'세금계산서 발행액 {D.INV_TOT:,} ÷ 조회수 {D.VIEW:,}', f'{D.INV_TOT/D.VIEW:.2f}원'),
                 ('CPE', f'세금계산서 발행액 {D.INV_TOT:,} ÷ 참여 {D.ENGA:,}', f'{D.INV_TOT/D.ENGA:,.0f}원'),
                 ('CPV (시딩 비용)', f'시딩 비용 {D.COST:,} ÷ 조회수 {D.VIEW:,}', f'{D.COST/D.VIEW:.2f}원'),
                 ('건당 청구액', f'세금계산서 발행액 {D.INV_TOT:,} ÷ 시딩 {D.QTY:,}건', f'{D.INV_TOT//D.QTY:,}원'),
                 ('ER (참여율)', f'(좋아요 {D.LIKE:,} + 댓글 {D.CMNT:,}) ÷ 조회수 {D.VIEW:,}', f'{D.ENGA/D.VIEW*100:.2f}%'),
                 ('ROAS', '전환 매출 ÷ 광고비', f"파트너십 {D.PA['rev']/D.PA['ad']*100:.1f}% / 합산 {D.REV/D.AD*100:.1f}%"),
                 ('CPA', '광고비 ÷ 구매건수', f'{D.AD/D.BUY:,.0f}원')]:
    line(ws, r, [a, b, c_, ''], wrap=2); r += 1
r += 1
note(ws, r, [
    '출처 — 세금계산서 월별 확정 총액표(26.09.07) · 2025 시딩 정산표 · 브랜드별 진행월 대조표 · 글로벌 시딩 진행 라인 · IMD 성과 통계 v8(26.09.04) · IMD 컬러그램(26.09.07) 103행 · IMD 바이오힐보 US(26.09.07) 36행 · 식물나라 시딩현황(케어플러스 원고료)',
    '데이터 단일 원천 — reports/data.py (모든 산출물이 이 파일에서만 값을 읽는다 · 저장 전 21항목 자체 검증)',
    '개인정산 정보(주민등록번호·계좌·주소·연락처)는 원장 시트에만 있으며 본 보고서에는 포함하지 않음',
])
ws.freeze_panes = 'A4'
# ════════════════════════════════ A1_세금계산서
ws = sheet('A1_세금계산서', '[부속] 세금계산서 발행 내역', '발행 확정 총액(26.09.07) 및 프로젝트별 내역 대조', 'FF9E9E9E')
MONS = ['2월', '3월', '4월', '5월', '6월', '7월', '8월']
widths(ws, [30, 14] + [13] * 7 + [15])
r = 4

r = sec(ws, r, '① 2026년 발행 확정 총액  —  기준 자료')
header(ws, r, ['구분', '국가'] + MONS + ['합계'], 20); r += 1
line(ws, r, ['발행 확정 (26.09.07)', '구분 없음', 92750000, 88812500, 129388689, 120981325,
             133875000, 94312500, 180394250, f'=SUM(C{r}:I{r})'],
     {k: MON for k in range(3, 11)}, fill=FT, bold=True)
fix_r = r; r += 2

r = sec(ws, r, '② 프로젝트별 내역  —  참고 (합계는 ①이 우선)')
header(ws, r, ['프로젝트', '국가'] + MONS + ['합계'], 20); r += 1
INV = [
    ('올영PB라이프 - 아이디얼포맨', '국내', {'4월': 108000000, '5월': 74125000, '6월': 32125000, '7월': 16125000, '8월': 32000000}),
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
line(ws, r, ['내역 소계', ''] + [f'=SUM({get_column_letter(c)}{s0}:{get_column_letter(c)}{r-1})' for c in range(3, 10)] +
     [f'=SUM(J{s0}:J{r-1})'], {k: MON for k in range(3, 11)}, fill=FA, bold=True)
sub_r = r; r += 1
line(ws, r, ['차이 (① − ②)', '귀속 확인필요'] +
     [f'={get_column_letter(c)}{fix_r}-{get_column_letter(c)}{sub_r}' for c in range(3, 10)] +
     [f'=J{fix_r}-J{sub_r}'], {k: MON for k in range(3, 11)}, fill=FT, bold=True)
r += 2

r = sec(ws, r, '③ 세금계산서 총 발행액 · 국가 배분')
header(ws, r, ['구분', '국가', '', '', '', '', '', '', '', '금액'], 20); r += 1
line(ws, r, ['2026.02~08 발행 확정', '구분 없음', None, None, None, None, None, None, None,
             f'=J{fix_r}'], {10: MON}); r += 1
line(ws, r, ['2025 시딩 정산 (국내 272건 + US 68건)', '국내+US', None, None, None, None, None, None, None,
             59830000], {10: MON}); r += 1
line(ws, r, ['세금계산서 총액 (발행 확정)', '', None, None, None, None, None, None, None,
             f'=J{r-2}+J{r-1}'], {10: MON}, fill=FT, bold=True)
tot_r = r; r += 1
line(ws, r, ['  글로벌 · 일본(JP) 발행분', '일본(JP)', None, None, None, None, None, None, None, 31796000],
     {10: MON}); r += 1
line(ws, r, ['  글로벌 · 바이오힐보 US 틱톡 (2025 시딩 정산 내)', '미국(US)', None, None, None, None, None, None,
             None, 30000000], {10: MON}); r += 1
line(ws, r, ['  국내 (KR) = 총액 − 글로벌', '국내', None, None, None, None, None, None, None,
             f'=J{tot_r}-J{r-2}-J{r-1}'], {10: MON}, fill=FA, bold=True); r += 1
line(ws, r, ['  + JP 미발행 (26.09)', '일본(JP)', None, None, None, None, None, None, None, 8715000],
     {10: MON}); r += 1
line(ws, r, ['진행 기준 총계', '', None, None, None, None, None, None, None,
             f'=J{tot_r}+J{r-1}'], {10: MON}, fill=FT, bold=True)
r += 2
note(ws, r, [
    '· 총 발행액 900,344,264원 = 26.02~08 확정 840,514,264 + 2025 시딩 정산 59,830,000.',
    '· 2025년 시딩은 전량 59,830,000원 안에서 운영되었고, 그중 바이오힐보 US 틱톡 시딩이 30,000,000원이다. 국내 272건 대응액은 29,830,000원.',
    '· 기존 집계 931,288,075원은 US 30,000,000원을 2025 시딩 정산과 별도로 한 번 더 더한 값이었다 → 중복 30,000,000. 여기에 26년 확정 −943,811을 반영해 순 −30,943,811원.',
    '· 확정 반영으로 빠진 것 — 26.02 아이디얼포맨 계상분 71,375,000원 (확정 2월 92,750,000원은 루테카 단독과 일치).',
    '· 확정 반영으로 더해진 것 — 26.03 +42,075,000 / 26.04 +5,058,689 / 26.05 +22,553,250 / 26.08 +31,740,500, 26.07 −30,996,250 (7→8월 발행 이동 추정) → 순 +70,431,189. 프로젝트별 귀속은 확인 필요.',
    '· 총액표에 국가 구분이 없어 글로벌 시딩 라인 자료로 귀속했다. 루테카·바이오힐보는 분리 근거가 없어 혼재로 둔다.',
])
ws.freeze_panes = 'C5'

# ════════════════════════════════ A2_글로벌시딩
ws = sheet('A2_글로벌시딩', '[부속] 글로벌 시딩 진행 라인', '일본(JP) 99건 / 원고료 40,511,000원 (WM·CG 건당 415,000 · BOH 333,000) · 미국(US) 68건 / 30,000,000원', 'FF9E9E9E')
widths(ws, [14, 42, 10, 9, 14, 9, 12, 12])
r = 4
header(ws, r, ['브랜드', '제품', '일정', '수량', '원고료', '건당', 'IMD 집계', ''], 20); r += 1
GJP = [
    ('웨이크메이크', '소블아 (7월 업로드 고정)', '7/20 ~', 3, 1245000, '내'),
    ('웨이크메이크', '소블아 (7월 업로드 자율)', '7/20 ~', 4, 1660000, '내'),
    ('웨이크메이크', '심리스파운데이션 (7월 업로드 자율)', '7/20 ~', 3, 1245000, '내'),
    ('웨이크메이크', '심리스 파운데이션 (8월 업로드)', '8/13 ~', 12, 4980000, '내'),
    ('웨이크메이크', '심리스 파운데이션', '8/13 ~', 7, 2905000, '내'),
    ('웨이크메이크', '소블아 오프체리', '8/13 ~', 10, 4150000, '내'),
    ('웨이크메이크', '소프트블러링아이팔레트', '9/14 ~', 21, 8715000, '외 (9월)'),
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
    '· JP 원고료 단가 확정 — 웨이크메이크·컬러그램 건당 415,000원 / 바이오힐보 NAD 건당 333,000원 (마크업·VAT 제외).',
    '· 마크업 검증 — WM 발행 대응 39건 16,185,000 × 1.25 × 1.1 = 22,254,375 vs 세금계산서 22,067,500 (차이 −186,875 · −0.8%로 정합).',
    '· CG 32건 13,280,000 × 1.25 × 1.1 = 18,260,000 vs 세금계산서 20,225,000 → +1,965,000(+10.8%) 초과. 국내분 혼재 또는 마크업률 상이 여부 확인 필요 (A4-9).',
    '· IMD의 JP 원고료 기재값(WM 23,004,315 · CG 20,297,925 · BOH 1,665,965 = 44,968,205)은 라인 단위 총액을 행마다 반복 기재한 값이다. 2,706,390 / 2,255,325 두 금액이 WM·CG에 동일하게 등장한다 → 집계에 사용하지 않고 위 확정 단가를 적용했다.',
    '· 9/14~ 업로드 21건은 26.09 진행분으로 세금계산서 미발행이다 (21 × 415,000 = 8,715,000).',
])
ws.freeze_panes = 'A5'

# ════════════════════════════════ A3_2025계약·진행월
ws = sheet('A3_2025계약·진행월', '[부속] 2025 시딩 계약 · 브랜드별 진행월 대조',
           '2025 시딩 정산 59,830,000원 (국내 272건 29,830,000 + 바이오힐보 US 틱톡 68건 30,000,000) · 진행월 대조표', 'FF9E9E9E')
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
     [170000, '국내 272건 / 29,830,000원 · 잔여인원 0.8명 → 소진 완료'],
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
ws = sheet('A4_확인필요', '[부속] 확인 필요 항목', '자료 간 불일치 · 미확정 사항 · 해결·확정 4건 포함 · 26.09.07 갱신', 'FFC0504D')
widths(ws, [5, 24, 78, 16, 10])
r = 4
header(ws, r, ['#', '항목', '내용', '금액 영향', '우선순위'], 20); r += 1
GAPS = [
    (1, '발행액 브랜드 미귀속 70,431,189원', '26.02~08 발행 확정액이 프로젝트별 내역보다 70,431,189원 많다(03 +42,075,000 / 04 +5,058,689 / 05 +22,553,250 / 07 −30,996,250 / 08 +31,740,500). 이 때문에 브랜드별 청구 기준 단가를 산출하지 못했다', 70431189, '높음'),
    (2, '청구 대응 진행 건수', '세금계산서에 수량 컬럼이 없어 900,344,264원에 대응하는 진행 건수를 확정할 수 없다. 2025 계약 219건 등 성과 집계 밖 건이 청구에 포함돼 있다. 확정되면 건당 청구액이 정확해진다', None, '높음'),
    (3, '좋아요 데이터 공란 519건', 'IMD 2,112건 중 519건의 좋아요 칸이 공란(조회수 8,976,727). 대부분 댓글이 있어 실제 0이 아닌 수집 누락. 추정하지 않았으므로 참여 391,960 · ER 0.68% · CPE 2,297원은 실제보다 불리한 값이다', None, '높음'),
    (4, '타 브랜드 IMD 최신화', '컬러그램이 v8 77건 → 최신 103건으로 26건 누락돼 있었다(광고비 3.1배·매출 2.2배 차이). 필리밀리(27건 ER 0.18%) 등 다른 브랜드도 동일 누락 가능성 → 브랜드별 최신 IMD 확인 필요', None, '높음'),
    (5, 'IMD 국가 분류 오류', 'IMD는 일본 24건(WM 9·CG 8·BOH 7)만 분류하나 실제 99건. 54건(WM 30·CG 24)이 국내로 오분류', None, '높음'),
    (6, '바이오힐보 국내 비용 귀속', '국내 195건·원고료 37,900,000인데 26년 발행액은 6,687,500뿐. 2025 계약 + 브랜드사 직계약으로 분산 추정', None, '높음'),
    (7, '식물나라 소계 오류', '원본 대조표 소계(384건 / 68,177,200 / 95,145,250)가 구버전(370건) 값. 재계산 시 87,676,850 / 106,723,250', 11578000, '높음'),
    (8, '아포맨 누계 행 상충', '요약 행에 194,450,000 / 308,887,500 / 243,062,500 세 값 동시 기재. 월별 상세(237,400,000 / 296,750,000)가 정답', 12137500, '높음'),
    (9, '컬러그램 JP 청구액 초과', 'CG JP 32건 원고료 13,280,000 × 마크업 1.25 × VAT 1.1 = 18,260,000인데 세금계산서는 20,225,000 → +1,965,000(+10.8%). WM은 −0.8%로 정합하므로 CG 발행액에 국내분이 섞였거나 마크업률이 다름', 1965000, '중'),
    (10, 'IMD JP 원고료 기재 오류', 'IMD의 JP 원고료가 라인 총액을 행마다 반복 기재한 값(2,706,390 / 2,255,325이 WM·CG에 동일 등장, 합 44,968,205 · 건당 255만원). 확정 단가 40,511,000으로 교체했으나 IMD 원본 수정 필요', 4457205, '중'),
    (11, '색조 3사 국내분 미발행', 'WM 국내 70건·CG 45건·FM 27건의 국내 시딩 비용이 26년 발행액에 없음. 하반기 계약으로 26.09 이후 발행 예정인지 확인', 34300000, '중'),
    (12, '비시딩 항목 전체 목록', '지급대행·모션그래픽·솔루션 32,400,000원은 확인분만. 전체 목록 확보 시 건당 단가 정밀화', None, '중'),
    (13, '2025년 상반기 데이터', 'IMD가 25.09.29부터라 25년 1~9월 업로드 데이터 없음. 2025 계약 272건 중 219건이 성과 집계 밖', None, '중'),
    (14, '루테카 US 수량', 'US TikTok 약 20건이 IMD·세금계산서 어느 쪽에서도 특정되지 않아 집계 제외', None, '낮음'),
    (15, '바이오힐보 US 미집계 32건', 'US 68건 중 36건만 성과 확보(4,739,117뷰). 나머지 32건 조회수·참여 데이터 미확보 → 실제 성과는 현재 수치 이상', None, '중'),
    (16, '바이오힐보 US 원고료', 'US 36건 원고료 미기재. 청구 30,000,000원만 확인되어 CPV는 청구 기준으로 산출(국내·JP와 기준 상이)', 30000000, '중'),
    (17, '케어플러스 청구 차이', '협의 비용 합 7,825,000(+VAT 8,607,500) vs 세금계산서 8,998,750 → 391,250 차이', 391250, '낮음'),
    ('해결', '루테카 발행액 71,375,000', '26.02 확정액이 92,750,000(루테카 단독)으로 확정되어 아이디얼포맨 2월 계상분 71,375,000은 제외 확정', None, '완료'),
    ('해결', '케어플러스 원고료', '시딩현황에서 확인 — 마이크로 5건×200,000 + 나노 27건×100,000 = 3,700,000원. CPV 19.63원 / CPE 2,012원', None, '완료'),
    ('확정', '2025 시딩 정산 구조', '2025 시딩은 전량 59,830,000원 내에서 운영. 그중 바이오힐보 US 틱톡 68건이 30,000,000원, 국내 272건이 29,830,000원(건당 109,669원). 국내 272건에 무상건·소액건이 섞여 있어 이 단가가 정상 — 오류 아님', None, '완료'),
    ('해결', '바이오힐보 US 성과', 'IMD 26.09.07로 36건 확보 — 4,739,117뷰 / 좋아요 23,068 / 댓글 1,066 / 최대 2,954,129뷰(전체 1위)', None, '완료'),
]
for n, item, desc, amt, pri in GAPS:
    line(ws, r, [n, item, desc, amt, pri], ({4: MON} if isinstance(n, str) else {1: CNT, 4: MON}),
         wrap=3, h=30, fill=(FA if isinstance(n, str) else None)); r += 1
r += 1
note(ws, r, ['· 금액 영향은 미설명·수정 필요 금액의 절대값이며 서로 중복될 수 있어 합산하지 않는다.'])
ws.freeze_panes = 'A5'

out = '/home/user/-_oy/reports/올리브영_시딩_성과보고서_HAN_260907.xlsx'
wb.save(out); print('saved', out)
