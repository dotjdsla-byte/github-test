council = None
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

KF = "맑은 고딕"
def F(sz=10, b=False, color="000000", it=False):
    return Font(name=KF, size=sz, bold=b, color=color, italic=it)

THIN = Side(style="thin", color="A6A6A6")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

NAVY   = PatternFill("solid", fgColor="1F3864")   # title
HEADF  = PatternFill("solid", fgColor="2E5C8A")   # column headers
GRPF   = PatternFill("solid", fgColor="D9E2F3")   # group label col
SUBF   = PatternFill("solid", fgColor="EDF2F9")   # sub-item col
YEL    = PatternFill("solid", fgColor="FFF2CC")   # user input
ALT    = PatternFill("solid", fgColor="F7F9FC")

WRAP = Alignment(horizontal="left", vertical="center", wrap_text=True)
CTR  = Alignment(horizontal="center", vertical="center", wrap_text=True)

wb = openpyxl.Workbook()

# =====================================================================
# SHEET 1 - 제품 요약 (고객 요청 양식 기준)
# =====================================================================
ws = wb.active
ws.title = "제품 요약"

PRODUCTS = ["Unitor™ FuelPower™\nConditioner",
            "Unitor™ FuelPower™\nCatalyst",
            "Unitor™ DieselPower™\nLubricity"]

ws.merge_cells("A1:F1")
ws["A1"] = "Unitor™ 연료유 첨가제 제품 자료 (제출용)"
ws["A1"].font = F(16, True, "FFFFFF"); ws["A1"].fill = NAVY; ws["A1"].alignment = CTR
ws.row_dimensions[1].height = 34

ws.merge_cells("A2:F2")
ws["A2"] = ("※ 고객 요청 양식 기준 재구성 · 제품 3종 (FuelPower Conditioner / FuelPower Catalyst / DieselPower Lubricity)\n"
            "※ 제품별 적용 유종이 상이하여 열(column)을 「유종」이 아닌 「제품」 기준으로 배열하고, 유종 적용 범위는 별도 행 및 '유종별 적용' 시트에 표기")
ws["A2"].font = F(9, color="404040"); ws["A2"].alignment = WRAP
ws.row_dimensions[2].height = 30

HDR = 4
hdrs = ["구분", "항목", "제품 ①", "제품 ②", "제품 ③", "비고"]
for i, h in enumerate(hdrs, start=1):
    c = ws.cell(HDR, i, h)
    c.font = F(10, True, "FFFFFF"); c.fill = HEADF; c.alignment = CTR; c.border = BOX
ws.row_dimensions[HDR].height = 22

# rows: (group, item, p1, p2, p3, remark, kind)
#   kind: '' normal | 'input' yellow | 'num' numeric blue | 'formula'
R = []
R.append(("기본 정보", "제품명", PRODUCTS[0], PRODUCTS[1], PRODUCTS[2], "", "name"))
R.append((None, "제조사 / 브랜드", "UNITOR\n(Wilhelmsen Ships Service)", "UNITOR\n(Wilhelmsen Ships Service)", "UNITOR\n(Wilhelmsen Ships Service)", "", ""))
R.append((None, "국내 총판 · 대리점", "WSS Korea", "WSS Korea", "WSS Korea", "", ""))
R.append((None, "제품 번호", "확인 필요", "확인 필요", "650-779094", "제품 매뉴얼 기재 번호", "input"))

R.append(("적용 범위", "적용 유종",
          "HSFO / VLSFO /\nBio-VLSFO (B30) / Bio-HFO\n※ 잔사유 전용 (MGO 미적용)",
          "HSFO / VLSFO /\nBio-VLSFO (B30) / Bio-HFO\n※ 잔사유 전용 (MGO 미적용)",
          "LSMGO / MGO / MDO / ULSD\nBio-LSMGO (조건부)\n※ 증류유 전용",
          "상세: '유종별 적용' 시트 참조", ""))
R.append((None, "사용 목적",
          "아스팔텐 슬러지 분산 및 침전 억제\n연료 안정성 · 블렌드 호환성 확보\n(Stability & Dispersant)",
          "연소 효율 개선 (Combustion Improver)\n+ 경미한 슬러지 분산\n(Mild sludge dispersancy)",
          "윤활성 개선 (Lubricity Improver)\n연료펌프 플런저 · 인젝터\n마모 및 고착 방지", "", ""))
R.append((None, "기능별 주요 성분",
          "아스팔텐 분산제 계열\n(상세 조성 SDS 별첨)",
          "철계 연소촉매\n(Iron-based catalyst)",
          "윤활성 향상제 계열\n(상세 조성 SDS 별첨)",
          "Conditioner · Lubricity 상세 조성 확인 필요", ""))

R.append(("성상", "밀도 (kg/m³ @15℃)", "0.84", "0.96", "확인 필요", "", ""))
R.append((None, "점도 (cSt)", "-", "-", "-", "", ""))
R.append((None, "인화점 (℃)", "61℃ 초과", "61℃ 초과", "확인 필요", "", ""))

R.append(("투입 방법", "투입 지점",
          "Storage TK\n(벙커링 시 투입 권장)",
          "Storage TK\n(벙커링 시 투입 권장)",
          "Storage → Settling / Service\n(Transfer Pump 이용)", "", ""))
R.append((None, "취급 · 주의사항", "-", "-",
          "보관온도 20~40℃\n천연고무 · NBR · PVC · 폴리우레탄 씰 부적합\n유동점강하제 · 세탄가향상제 · 소포제 병용 가능",
          "DP Lubricity 제품 매뉴얼 기준", ""))

R.append(("투입 비율\n(상세)", "투입기준 ①\n평상시 (General)",
          "1L : 15 MT\n(1 : 15,000)", "1L : 15 MT\n(1 : 15,000)", "1L : 10 MT\n(0.1 L/ton, 100 ppm)", "", ""))
R.append((None, "투입기준 ②\n최대 (Max.)",
          "1L : 5 MT\n(1 : 5,000)\n※ 고 TSP · 불안정 연료 · 블렌드 불호환 시",
          "1L : 4 MT\n(1 : 4,000)",
          "1L : 5 MT\n(0.2 L/ton, 200 ppm)", "", ""))
R.append((None, "투입기준 ③\n최소 (Min.)",
          "1L : 25 MT\n(1 : 25,000)", "1L : 15 MT\n(1 : 15,000)", "1L : 10 MT\n(0.1 L/ton)", "", ""))
R.append((None, "실측 시험 근거 투입량",
          "Hot Spin 시험 기준\n(제품 자료상 1 : 10,000 적용 사례 있음)",
          "100 ppm = 1L : 10 MT\n★ SFOC 개선 실측 데이터는 전부\n100 ppm 조건에서 취득",
          "125 ppm = 0.125 L/ton\n→ HFRR 460 µm 미만 달성",
          "★ 연료절감을 목적으로 할 경우\nCatalyst는 1L : 10 MT 적용 권장", ""))

R.append(("투입량 · 비용", "평상시 기준 처리량\n(1 L 당 처리 톤수, MT)", 15, 15, 10,
          "아래 계산식의 입력값 (파란색)", "num"))
R.append((None, "1,000 M/T 기준 투입량 (L)", None, None, None, "1,000 ÷ 처리톤수 (자동 계산)", "qty"))
R.append((None, "제품단가 (USD / LITER)", None, None, None, "★ 노란색 셀에 단가 입력", "price"))
R.append((None, "1,000 M/T 기준 예상 비용 (USD)", None, None, None, "투입량 × 단가 (자동 계산)", "cost"))

R.append(("공급 조건", "포장 규격", "확인 필요", "확인 필요", "확인 필요", "표준 CAN / DRUM 규격 확인 후 기입", "input"))
R.append((None, "보급가능지역 - 국내", "가능", "가능", "가능", "", ""))
R.append((None, "보급가능지역 - 국외", "가능\n(WSS 글로벌 네트워크)", "가능\n(WSS 글로벌 네트워크)", "가능\n(WSS 글로벌 네트워크)",
          "구체적 공급 가능 항구 목록 별첨 필요", ""))

R.append(("기대 효과", "연료절감\n[보증 여부]",
          "단독 기준 연료절감률 미제시\n(보증 불가)\n※ Catalyst 병용 시 SFOC 최대 1.31% 개선\n  (CCS 입회 실선 엔진 시험)",
          "VLSFO : SFOC 1.454% 개선\nBio-VLSFO (B30) : SFOC 최대 1.31% 개선\n[보증 불가 — 실측 시험 자료 기준]",
          "해당 없음\n(윤활성 개선 제품)",
          "실측 시험 자료는 'Reference 상세' 시트 참조", ""))
R.append((None, "기타 효과",
          "Hot Spin 침전 평균 35.5% 감소\n(8개 항구 RMG380 실측, 최고 46.67%)\n정유기 · 필터 막힘 저감\n블렌드 간 호환성 개선",
          "ECN (점화성 지수) 6.25% 개선\nIgnition Delay 4.92% 단축\n(IP 541 / FIA-100 FCA, VPS)",
          "HFRR 마모흔 535 µm (ULSD 단독, 부적합)\n→ 100 ppm 396 µm / 150 ppm 355 µm\nISO 8217 (520) · EN 590 (460) 모두 충족",
          "", ""))

R.append(("검증 · 실적", "제3자 검증 /\n선급 · 시험기관",
          "Lloyd's Register FOBAS\nAdditive Performance\nCertification Scheme",
          "Veritas Petroleum Services (VPS)\nChina Classification Society (CCS)\nWärtsilä Letter of No Objection",
          "Lloyd's Register\nProduct Verification Scheme\nISO 12156-1 (HFRR)", "", ""))
R.append((None, "주요 Reference",
          "8개 항구 RMG380 Hot Spin 실측\nLR FOBAS 인증",
          "Tokyo Univ. of Marine Science &\nTechnology (VLSFO)\nShanghai Maritime Univ. /\nMAN ME-C 6S35 (B30)",
          "사내 Lab Test (ULSD, ISO 12156-1)\nLR Product Verification Scheme",
          "상세: 'Reference 상세' 시트 참조", ""))
R.append((None, "주요고객 · 판매실적\n(최근 3년)", "작성 필요", "작성 필요", "작성 필요",
          "★ 거래처별 판매실적 내부 확인 후 기입", "input"))
R.append(("담당자", "담당자 연락처", "이름 :\n이메일 :\n전화번호 :", "이름 :\n이메일 :\n전화번호 :", "이름 :\n이메일 :\n전화번호 :",
          "★ 기입 필요", "input"))

r = HDR + 1
group_spans = []   # (start_row, end_row, label)
cur_label, cur_start = None, None
qty_row = price_row = cost_row = num_row = None

for grp, item, p1, p2, p3, rem, kind in R:
    if grp is not None:
        if cur_label is not None:
            group_spans.append((cur_start, r - 1, cur_label))
        cur_label, cur_start = grp, r

    ws.cell(r, 2, item).font = F(10, True)
    ws.cell(r, 2).fill = SUBF
    ws.cell(r, 2).alignment = WRAP

    if kind == "num":
        num_row = r
        for i, v in enumerate([p1, p2, p3]):
            c = ws.cell(r, 3 + i, v)
            c.font = F(10, True, "0000FF"); c.alignment = CTR
            c.number_format = '0"  MT / 1L"'
    elif kind == "qty":
        qty_row = r
        for i in range(3):
            col = get_column_letter(3 + i)
            c = ws.cell(r, 3 + i, f"=ROUND(1000/{col}{num_row},0)")
            c.font = F(10, True); c.alignment = CTR
            c.number_format = '#,##0"  L"'
    elif kind == "price":
        price_row = r
        for i in range(3):
            c = ws.cell(r, 3 + i, None)
            c.fill = YEL; c.font = F(10, True, "0000FF"); c.alignment = CTR
            c.number_format = '"USD "#,##0.00'
    elif kind == "cost":
        cost_row = r
        for i in range(3):
            col = get_column_letter(3 + i)
            c = ws.cell(r, 3 + i, f'=IF({col}{price_row}="","",{col}{qty_row}*{col}{price_row})')
            c.font = F(10, True); c.alignment = CTR
            c.number_format = '"USD "#,##0;("USD "#,##0);-'
    else:
        for i, v in enumerate([p1, p2, p3]):
            c = ws.cell(r, 3 + i, v)
            c.alignment = WRAP
            if kind == "name":
                c.font = F(11, True, "1F3864"); c.alignment = CTR
            elif v == "확인 필요" or v == "작성 필요":
                c.font = F(10, True, "C00000"); c.fill = YEL; c.alignment = CTR
            elif kind == "input":
                c.font = F(10); c.fill = YEL
            else:
                c.font = F(10)

    rc = ws.cell(r, 6, rem)
    rc.font = F(9, color="595959"); rc.alignment = WRAP
    if rem.startswith("★"):
        rc.font = F(9, True, "C00000")

    for i in range(1, 7):
        ws.cell(r, i).border = BOX
    r += 1

group_spans.append((cur_start, r - 1, cur_label))
last_row = r - 1

for s, e, label in group_spans:
    if e > s:
        ws.merge_cells(start_row=s, start_column=1, end_row=e, end_column=1)
    c = ws.cell(s, 1, label)
    c.font = F(10, True, "1F3864"); c.fill = GRPF
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for rr in range(s, e + 1):
        ws.cell(rr, 1).border = BOX

note = last_row + 2
ws.merge_cells(start_row=note, start_column=1, end_row=note, end_column=6)
n = ws.cell(note, 1,
    "▣ 작성 안내\n"
    "  · 노란색 셀 = 기입 필요 항목 (제품단가 / 포장 규격 / 판매실적 / 담당자 연락처 / 제품번호 · 성상 일부)\n"
    "  · 파란색 숫자 = 계산 입력값. '평상시 기준 처리량'을 바꾸면 투입량과 예상 비용이 자동으로 재계산됩니다.\n"
    "  · 예상 비용 = 1,000 M/T 기준 투입량 × 제품단가. 단가 입력 전에는 공란으로 표시됩니다.\n"
    "  · ★ Catalyst의 SFOC 개선 실측치(1.454% / 1.31%)는 100 ppm(1L : 10 MT) 조건에서 취득한 값입니다. "
    "평상시 권장 투입비율(1L : 15 MT)과 다르므로, 연료절감을 제안 근거로 사용할 경우 투입량을 1L : 10 MT(1,000 M/T 당 100 L)로 산정해야 합니다.\n"
    "  · 본 자료의 모든 수치는 제3자 시험기관 · 선급 실측 자료에 근거하며, 연료절감률에 대한 보증은 제공되지 않습니다.")
n.font = F(9, color="404040"); n.alignment = WRAP
n.fill = PatternFill("solid", fgColor="FFF9E6"); n.border = BOX
ws.row_dimensions[note].height = 108

ws.column_dimensions["A"].width = 12
ws.column_dimensions["B"].width = 22
for col in ("C", "D", "E"):
    ws.column_dimensions[col].width = 40
ws.column_dimensions["F"].width = 30
for rr in range(HDR + 1, last_row + 1):
    ws.row_dimensions[rr].height = 46
ws.freeze_panes = "C5"

# =====================================================================
# SHEET 2 - 유종별 적용
# =====================================================================
ws2 = wb.create_sheet("유종별 적용")
ws2.merge_cells("A1:H1")
ws2["A1"] = "제품별 · 유종별 적용 범위"
ws2["A1"].font = F(14, True, "FFFFFF"); ws2["A1"].fill = NAVY; ws2["A1"].alignment = CTR
ws2.row_dimensions[1].height = 30
ws2.merge_cells("A2:H2")
ws2["A2"] = "● 권장 (Recommended)    △ 조건부 적용 (Case by case)    –  해당 없음 (N/A)"
ws2["A2"].font = F(9, color="404040"); ws2["A2"].alignment = CTR

fuels = ["HSFO", "VLSFO", "Bio-VLSFO\n/ Bio-HFO\n(B30)", "LSMGO / MGO\n/ MDO", "Bio-LSMGO\n/ Bio-MDO", "B100", "기타 용도\n(Tank Mixing 등)"]
for i, h in enumerate(["제품명"] + fuels, start=1):
    c = ws2.cell(4, i, h)
    c.font = F(10, True, "FFFFFF"); c.fill = HEADF; c.alignment = CTR; c.border = BOX
ws2.row_dimensions[4].height = 42

MTX = [
    ("Unitor™ FuelPower™ Conditioner", ["●", "●", "●", "–", "–", "–", "●"]),
    ("Unitor™ FuelPower™ Catalyst",    ["●", "●", "●", "–", "–", "–", "–"]),
    ("Unitor™ DieselPower™ Lubricity", ["–", "–", "–", "●", "△", "△", "–"]),
]
rr = 5
for name, marks in MTX:
    c = ws2.cell(rr, 1, name); c.font = F(10, True); c.alignment = WRAP; c.border = BOX
    for i, m in enumerate(marks, start=2):
        mc = ws2.cell(rr, i, m)
        mc.alignment = CTR; mc.border = BOX
        mc.font = F(13, True, "1F3864" if m == "●" else ("BF8F00" if m == "△" else "A6A6A6"))
    ws2.row_dimensions[rr].height = 30
    rr += 1

rr += 1
ws2.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=8)
c = ws2.cell(rr, 1,
    "▣ 적용 근거 및 단서\n"
    "  · FuelPower Conditioner / Catalyst : 잔사유(Residual) 전용 제품으로 MGO 등 증류유에는 적용하지 않습니다.\n"
    "  · 기타 용도(Tank Mixing) : 서로 다른 벙커 간 혼합 시 발생하는 불호환(Incompatibility) 완화 목적의 Conditioner 적용을 의미합니다.\n"
    "  · DieselPower Lubricity — Bio-LSMGO : FAME 성분 자체가 윤활성을 보완하므로 윤활성 부적합 가능성이 낮습니다. 성적서상 HFRR 값 확인 후 적용을 판단합니다.\n"
    "  · DieselPower Lubricity — B100 : HVO(수소첨가 식물성 오일)에 한해 윤활성 저하가 나타납니다. 일반 FAME B100은 해당하지 않습니다.\n"
    "  · 본 표는 3개 제품에 한정된 것으로, 미생물 · 저온유동성 · 산화안정성 · Ash · Soot 등 그 밖의 이슈에 대해서는 별도 제품군이 적용됩니다.")
c.font = F(9, color="404040"); c.alignment = WRAP
c.fill = PatternFill("solid", fgColor="FFF9E6"); c.border = BOX
ws2.row_dimensions[rr].height = 96

ws2.column_dimensions["A"].width = 38
for i in range(2, 9):
    ws2.column_dimensions[get_column_letter(i)].width = 16

# =====================================================================
# SHEET 3 - Reference 상세
# =====================================================================
ws3 = wb.create_sheet("Reference 상세")
ws3.merge_cells("A1:H1")
ws3["A1"] = "제품별 실측 시험 · 제3자 검증 Reference"
ws3["A1"].font = F(14, True, "FFFFFF"); ws3["A1"].fill = NAVY; ws3["A1"].alignment = CTR
ws3.row_dimensions[1].height = 30
ws3.merge_cells("A2:H2")
ws3["A2"] = "※ 시험/인증 기관, 적용 유종, 시험 엔진·선박, 주요 결과 요약"
ws3["A2"].font = F(9, color="404040"); ws3["A2"].alignment = WRAP

h3 = ["No.", "제품명", "적용 유종", "시험 / 인증 기관", "시험 엔진 · 선박", "주요 결과 요약", "시험 구분", "비고"]
for i, h in enumerate(h3, start=1):
    c = ws3.cell(4, i, h)
    c.font = F(10, True, "FFFFFF"); c.fill = HEADF; c.alignment = CTR; c.border = BOX
ws3.row_dimensions[4].height = 24

REFS = [
    (1, "Unitor™ FuelPower™ Conditioner", "HSFO / VLSFO\n(RMG380)",
     "Lloyd's Register FOBAS\nAdditive Performance\nCertification Scheme", "-",
     "Hot Spin 침전량 평균 35.5% 감소\n(8개 항구 실측, 최고 46.67%)\n분리온도 40 / 98℃ 조건",
     "Actual Trial\n(실측 시험)", "항구별 상세 결과 하단 표 참조"),
    (2, "Unitor™ FuelPower™ Conditioner\n+ Unitor™ FuelPower™ Catalyst",
     "Bio-VLSFO (B30)\nB30 = Bio 30%,\nFAME 순도 98.8%",
     "China Classification\nSociety (CCS)\nWärtsilä Letter of No Objection",
     "MAN ME-C 6S35\n(Shanghai Maritime Univ.)",
     "SFOC 최대 1.31% 개선\nCCS 전 구간 입회, 2초 간격 데이터 기록",
     "Actual Trial\n(실측 시험)", "CCS 검증 완료\nHot Spin 성능은 LR 검증"),
    (3, "Unitor™ FuelPower™ Catalyst", "VLSFO",
     "Veritas Petroleum\nServices (VPS)\nWärtsilä Letter of No Objection", "-",
     "SFOC 1.454% 개선\nECN (점화성 지수) 6.25% 개선\nIgnition Delay 4.92% 단축\n(IP 541 / FIA-100 FCA 연소분석)",
     "Actual Trial\n(실측 시험)", "SFOC 시험은 Tokyo Univ. of Marine\nScience & Technology 에서 고객사 실시"),
    (4, "Unitor™ FuelPower™ Catalyst",
     "Bio-VLSFO (B30)",
     "China Classification\nSociety (CCS)\nWärtsilä Letter of No Objection",
     "MAN ME-C 6S35\n(Shanghai Maritime Univ.)",
     "SFOC 최대 1.31% 개선\n(Engine Trial)", "Actual Trial\n(실측 시험)", "CCS 검증 완료"),
    (5, "Unitor™ DieselPower™ Lubricity", "LSMGO / ULSD",
     "Lloyd's Register\nProduct Verification Scheme\nISO 12156-1 (HFRR)", "-",
     "HFRR 마모흔 (WS1.4, µm)\n· ULSD 단독 : 535 (부적합)\n· + 100 ppm : 396 (합격)\n· + 150 ppm : 355 (합격)",
     "Lab Test\n(사내 시험)", "ISO 8217 한계 520 µm,\nEN 590 한계 460 µm 모두 충족"),
]
rr = 5
for row in REFS:
    for i, v in enumerate(row, start=1):
        c = ws3.cell(rr, i, v)
        c.border = BOX
        c.alignment = CTR if i in (1, 3, 7) else WRAP
        c.font = F(9, True) if i == 2 else F(9)
    if rr % 2 == 1:
        for i in range(1, 9):
            if not ws3.cell(rr, i).fill.fgColor.rgb or ws3.cell(rr, i).fill.patternType is None:
                ws3.cell(rr, i).fill = ALT
    ws3.row_dimensions[rr].height = 66
    rr += 1

rr += 1
sub = rr
ws3.merge_cells(start_row=sub, start_column=1, end_row=sub, end_column=8)
c = ws3.cell(sub, 1, "[상세] Unitor™ FuelPower™ Conditioner — Hot Spin Test 항구별 실측 결과 (RMG380)")
c.font = F(11, True, "FFFFFF"); c.fill = HEADF; c.alignment = Alignment(horizontal="left", vertical="center")
ws3.row_dimensions[sub].height = 24
rr += 1

hs_hdr = ["항구", "점도\n(cSt @50℃)", "MCR (%)", "분리온도\n(℃)", "Before\n(침전 %)", "After\n(침전 %)", "개선율 (%)", "비고"]
for i, h in enumerate(hs_hdr, start=1):
    c = ws3.cell(rr, i, h)
    c.font = F(9, True); c.fill = GRPF; c.alignment = CTR; c.border = BOX
ws3.row_dimensions[rr].height = 32
rr += 1

HS = [
    ("Zeebrugge", 335.2, 11.6, 98, 0.19, 0.13),
    ("Tuapse", 18.9, 2.96, 40, 0.24, 0.20),
    ("Amsterdam", 206.6, 10.91, 98, 0.15, 0.08),
    ("Flushing", 239.7, 9.08, 98, 0.13, 0.09),
    ("Duqm", 230.7, 7.15, 98, 0.33, 0.20),
    ("Amsterdam", 314.4, 10.02, 98, 0.16, 0.10),
    ("Vlissingen", 242.8, 10.5, 98, 0.13, 0.07),
    ("Callao", 335.1, 8.84, 98, 0.17, 0.11),
]
hs_start = rr
for port, visc, mcr, temp, bef, aft in HS:
    ws3.cell(rr, 1, port).font = F(9)
    ws3.cell(rr, 2, visc).number_format = '#,##0.0'
    ws3.cell(rr, 3, mcr).number_format = '0.00'
    ws3.cell(rr, 4, temp).number_format = '0"℃"'
    ws3.cell(rr, 5, bef).number_format = '0.00'
    ws3.cell(rr, 6, aft).number_format = '0.00'
    ws3.cell(rr, 7, f"=(E{rr}-F{rr})/E{rr}").number_format = '0.00%'
    ws3.cell(rr, 8, "저점도 · 분리온도 40℃ 조건" if port == "Tuapse" else "")
    ws3.cell(rr, 8).font = F(8, color="595959")
    for i in range(1, 9):
        ws3.cell(rr, i).border = BOX
        if i != 1 and i != 8:
            ws3.cell(rr, i).font = F(9)
            ws3.cell(rr, i).alignment = CTR
    ws3.cell(rr, 1).alignment = Alignment(horizontal="left", vertical="center")
    ws3.cell(rr, 8).alignment = WRAP
    ws3.row_dimensions[rr].height = 18
    rr += 1
hs_end = rr - 1

ws3.cell(rr, 1, "평균").font = F(9, True)
ws3.cell(rr, 1).alignment = CTR
for i in (2, 3, 5, 6):
    col = get_column_letter(i)
    c = ws3.cell(rr, i, f"=AVERAGE({col}{hs_start}:{col}{hs_end})")
    c.number_format = '#,##0.00'; c.font = F(9, True); c.alignment = CTR
ws3.cell(rr, 4, "-").alignment = CTR; ws3.cell(rr, 4).font = F(9, True)
c = ws3.cell(rr, 7, f"=AVERAGE(G{hs_start}:G{hs_end})")
c.number_format = '0.00%'; c.font = F(10, True, "1F3864"); c.alignment = CTR
ws3.cell(rr, 8, "최고 개선율 46.67% (Amsterdam)").font = F(8, True, "1F3864")
ws3.cell(rr, 8).alignment = WRAP
for i in range(1, 9):
    ws3.cell(rr, i).border = BOX
    ws3.cell(rr, i).fill = GRPF
avg_row = rr

rr += 2
ws3.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=8)
c = ws3.cell(rr, 1,
    "▣ 출처 및 단서\n"
    "  · Hot Spin 실측 : Wilhelmsen Ships Service 사내 실측 데이터 (RMG380, 분리온도 40 / 98℃) · Lloyd's Register FOBAS 인증\n"
    "  · Hot Spin Test 는 저장 중 발생할 슬러지 생성 가능성을 사전 측정하는 시험으로, 원심관 바닥 축적량 약 0.2% 이상을 위험 신호로 봅니다.\n"
    "  · HFRR 마모흔은 시험구(⌀6 mm 볼)에 생긴 자국의 지름이며, 펌프 내부 치수가 아닙니다. 한계선은 실제 분사장비 고장 이력에서 정해졌습니다.\n"
    "  · SFOC 개선 수치는 해당 시험 조건(연료 · 엔진 · 부하 · 투입량)에서의 실측 결과이며, 타 선박에서의 동일한 결과를 보증하지 않습니다.")
c.font = F(9, color="404040"); c.alignment = WRAP
c.fill = PatternFill("solid", fgColor="FFF9E6"); c.border = BOX
ws3.row_dimensions[rr].height = 82

widths3 = [22, 30, 20, 30, 26, 40, 14, 30]
for i, w in enumerate(widths3, start=1):
    ws3.column_dimensions[get_column_letter(i)].width = w

# =====================================================================
# SHEET 4 - 투입량·비용 산출
# =====================================================================
ws4 = wb.create_sheet("투입량·비용 산출")
ws4.merge_cells("A1:G1")
ws4["A1"] = "벙커량 기준 투입량 · 비용 산출"
ws4["A1"].font = F(14, True, "FFFFFF"); ws4["A1"].fill = NAVY; ws4["A1"].alignment = CTR
ws4.row_dimensions[1].height = 30
ws4.merge_cells("A2:G2")
ws4["A2"] = "※ 노란색 셀(벙커량 · 제품단가)만 입력하면 투입량과 비용이 자동 계산됩니다."
ws4["A2"].font = F(9, color="404040"); ws4["A2"].alignment = WRAP

ws4["A4"] = "벙커량 (M/T)"
ws4["A4"].font = F(11, True); ws4["A4"].alignment = Alignment(horizontal="right", vertical="center")
ws4["B4"] = 1000
ws4["B4"].fill = YEL; ws4["B4"].font = F(12, True, "0000FF"); ws4["B4"].alignment = CTR
ws4["B4"].number_format = '#,##0"  MT"'; ws4["B4"].border = BOX
ws4.merge_cells("C4:G4")
ws4["C4"] = "← 실제 벙커 수량을 입력하십시오 (기본값 1,000 M/T)"
ws4["C4"].font = F(9, color="595959"); ws4["C4"].alignment = Alignment(horizontal="left", vertical="center")

h4 = ["제품명", "적용 유종", "투입 시나리오", "투입 비율\n(1 L 당 처리 MT)", "필요 투입량 (L)", "제품단가\n(USD / L)", "예상 비용 (USD)"]
for i, h in enumerate(h4, start=1):
    c = ws4.cell(6, i, h)
    c.font = F(10, True, "FFFFFF"); c.fill = HEADF; c.alignment = CTR; c.border = BOX
ws4.row_dimensions[6].height = 36

CALC = [
    ("Unitor™ FuelPower™ Conditioner", "HSFO / VLSFO /\nBio-VLSFO", "평상시 (General)", 15),
    ("", "", "최대 (Max.)", 5),
    ("", "", "최소 (Min.)", 25),
    ("Unitor™ FuelPower™ Catalyst", "HSFO / VLSFO /\nBio-VLSFO", "평상시 (General)", 15),
    ("", "", "★ 연료절감 실측 조건 (100 ppm)", 10),
    ("", "", "최대 (Max.)", 4),
    ("Unitor™ DieselPower™ Lubricity", "LSMGO / MGO /\nMDO", "평상시 (General, 100 ppm)", 10),
    ("", "", "시험 근거 (125 ppm)", 8),
    ("", "", "최대 (Max., 200 ppm)", 5),
]
rr = 7
blocks = []
bstart = rr
for name, fuel, scen, ratio in CALC:
    if name and rr != bstart:
        blocks.append((bstart, rr - 1))
        bstart = rr
    ws4.cell(rr, 1, name).font = F(10, True); ws4.cell(rr, 1).alignment = WRAP
    ws4.cell(rr, 2, fuel).font = F(9); ws4.cell(rr, 2).alignment = WRAP
    c = ws4.cell(rr, 3, scen); c.font = F(9, True, "C00000") if scen.startswith("★") else F(9)
    c.alignment = WRAP
    c = ws4.cell(rr, 4, ratio); c.font = F(10, True, "0000FF"); c.alignment = CTR
    c.number_format = '0"  MT / 1L"'
    c = ws4.cell(rr, 5, f"=ROUND($B$4/D{rr},0)"); c.font = F(10, True); c.alignment = CTR
    c.number_format = '#,##0"  L"'
    c = ws4.cell(rr, 6, None); c.fill = YEL; c.font = F(10, True, "0000FF"); c.alignment = CTR
    c.number_format = '"USD "#,##0.00'
    c = ws4.cell(rr, 7, f'=IF(F{rr}="","",E{rr}*F{rr})'); c.font = F(10, True); c.alignment = CTR
    c.number_format = '"USD "#,##0;("USD "#,##0);-'
    for i in range(1, 8):
        ws4.cell(rr, i).border = BOX
    ws4.row_dimensions[rr].height = 30
    rr += 1
blocks.append((bstart, rr - 1))

for s, e in blocks:
    if e > s:
        ws4.merge_cells(start_row=s, start_column=1, end_row=e, end_column=1)
        ws4.merge_cells(start_row=s, start_column=2, end_row=e, end_column=2)
    ws4.cell(s, 1).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws4.cell(s, 2).alignment = CTR

rr += 1
ws4.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=7)
c = ws4.cell(rr, 1,
    "▣ 산출 기준\n"
    "  · 필요 투입량 (L) = 벙커량 (M/T) ÷ 투입 비율 (1 L 당 처리 MT), 리터 단위 반올림\n"
    "  · 예상 비용 (USD) = 필요 투입량 × 제품단가. 단가 입력 전에는 공란으로 표시됩니다.\n"
    "  · ★ FuelPower Catalyst 의 SFOC 개선 실측치(VLSFO 1.454% / Bio-VLSFO 1.31%)는 100 ppm(1L : 10 MT) 조건에서 취득한 값입니다. "
    "연료절감을 근거로 제안할 경우 이 행을 기준으로 산출하십시오.\n"
    "  · Conditioner 와 Catalyst 를 병용할 경우 두 제품의 투입량과 비용을 각각 합산해야 합니다. "
    "(Bio-VLSFO B30 실선 시험은 Catalyst 1L + Conditioner 1L per 10 MT 조건에서 수행되었습니다.)")
c.font = F(9, color="404040"); c.alignment = WRAP
c.fill = PatternFill("solid", fgColor="FFF9E6"); c.border = BOX
ws4.row_dimensions[rr].height = 90

widths4 = [34, 20, 30, 18, 18, 16, 20]
for i, w in enumerate(widths4, start=1):
    ws4.column_dimensions[get_column_letter(i)].width = w

# =====================================================================
out = "/home/user/github-test/Unitor_연료유첨가제_제품자료_3종.xlsx"
wb.save(out)
print("saved", out)
