from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.worksheet.page import PageMargins

wb = Workbook()
ws = wb.active
ws.title = "INVOICE"
ws.sheet_view.showGridLines = False

F = "Arial"
def font(size=10, bold=False, color="000000"):
    return Font(name=F, size=size, bold=bold, color=color)
thin = Side(style="thin", color="000000")
med = Side(style="medium", color="000000")
grey = PatternFill("solid", fgColor="D9D9D9")
light = PatternFill("solid", fgColor="F2F2F2")

# Columns: A label, B description, C currency, D amount
widths = {"A": 17, "B": 46, "C": 13, "D": 18}
for k, v in widths.items():
    ws.column_dimensions[k].width = v

def put(cell, value, size=10, bold=False, h="left", v="center", indent=None, color="000000", fmt=None):
    c = ws[cell]
    c.value = value
    c.font = font(size, bold, color)
    if indent is None: indent = 1 if h == "right" else 0
    c.alignment = Alignment(horizontal=h, vertical=v, indent=indent)
    if fmt:
        c.number_format = fmt
    return c

# ---------- Header (company) ----------
ws.merge_cells("A1:D1")
put("A1", "YOUR COMPANY NAME", 18, True)
ws.row_dimensions[1].height = 30
put("A2", "Address line 1", 9, color="404040")
put("A3", "Address line 2", 9, color="404040")
put("C2", "TEL", 9, h="right", color="404040");    put("D2", "00-00-000-0000", 9)
put("C3", "FAX", 9, h="right", color="404040");    put("D3", "00-00-000-0000", 9)
put("C4", "E-MAIL", 9, h="right", color="404040"); put("D4", "email@company.com", 9)
for col in "ABCD":
    ws[f"{col}5"].border = Border(bottom=thin)

# ---------- Title ----------
ws.row_dimensions[6].height = 8
ws.merge_cells("A7:D7")
put("A7", "INVOICE", 18, True, h="center")
ws["A7"].fill = grey
ws.row_dimensions[7].height = 32
ws.row_dimensions[8].height = 8

# ---------- Info block ----------
info = [
    ("BILL NO.", "INV-2026-0001"),
    ("OWNER", "Customer company name"),
    ("CC", "-"),
    ("YOUR REF. NO.", "Customer PO / reference no."),
    ("RE", "Vessel / project name"),
]
r = 9
for label, val in info:
    put(f"A{r}", f"{label} :", 10, True)
    ws.merge_cells(f"B{r}:D{r}")
    put(f"B{r}", val, 10)
    r += 1
# Date & currency on the last line, right side
put(f"A{r}", "DATE :", 10, True)
put(f"B{r}", "2026-11-27", 10, fmt="yyyy-mm-dd")
from datetime import date
ws[f"B{r}"].value = date(2026, 11, 27)
put(f"C{r}", "CURRENCY :", 10, True, h="right")
put(f"D{r}", "KRW", 10, True, h="center")
CUR = f"$D${r}"
ws.row_dimensions[r + 1].height = 10

# ---------- Description table ----------
top = r + 2
ws.merge_cells(f"A{top}:B{top}")
put(f"A{top}", "DESCRIPTION", 10, True, h="center")
ws.merge_cells(f"C{top}:D{top}")
put(f"C{top}", "AMOUNT", 10, True, h="center")
for col in "ABCD":
    ws[f"{col}{top}"].fill = light
ws.row_dimensions[top].height = 20

AMT = '#,##0;-#,##0;"-"'
rows = []  # (desc, indent, bold, amount or None)
rows.append(("", 0, False, None))
rows.append(("PLEASE SETTLE THE AMOUNT BELOW TO OUR ACCOUNT.", 0, False, None))
rows.append(("", 0, False, None))
rows.append(("1. STORE PARTS", 0, True, 45520594))
rows.append(("", 0, False, None))
rows.append(("2. ADDITIONAL CHARGE", 0, True, None))
rows.append(("1) PACKING CHARGE", 1, False, None))
rows.append(("- TON BAG * 6PCS", 2, False, 30000))
rows.append(("- WOODEN PALLET * 8PCS", 2, False, 360000))
rows.append(("2) TRUCKING CHARGE", 1, False, None))
rows.append(("- BUSAN - YEOSU ROUND TRIP, 5TON TRUCK", 2, False, 520000))
rows.append(("", 2, False, None))   # spare
rows.append(("", 2, False, None))   # spare
rows.append(("", 2, False, None))   # spare
rows.append(("", 0, False, None))

r = top + 1
first_item = r
for desc, ind, bold, amt in rows:
    ws.merge_cells(f"A{r}:B{r}")
    put(f"A{r}", desc if desc else None, 10, bold, indent=ind * 2 + 1)
    c = put(f"D{r}", amt, 10, h="right", fmt=AMT)
    put(f"C{r}", f'=IF(D{r}="","",{CUR})', 10, h="center")
    ws.row_dimensions[r].height = 17
    r += 1
last_item = r - 1

# Discount
ws.merge_cells(f"A{r}:B{r}")
put(f"A{r}", "3. DISCOUNT", 10, True, indent=1)
put(f"C{r}", f"={CUR}", 10, h="center")
put(f"D{r}", 0, 10, h="right", fmt='"-"#,##0;"-"#,##0;"-"')
disc = r
ws.row_dimensions[r].height = 17
r += 1
ws.merge_cells(f"A{r}:B{r}"); ws.row_dimensions[r].height = 10; r += 1

# Sub total
ws.merge_cells(f"A{r}:B{r}")
put(f"A{r}", "4. SUB TOTAL", 10, True, indent=1)
put(f"C{r}", f"={CUR}", 10, h="center")
put(f"D{r}", f"=SUM(D{first_item}:D{last_item})-D{disc}", 10, True, h="right", fmt=AMT)
sub = r
ws.row_dimensions[r].height = 17
r += 1
ws.merge_cells(f"A{r}:B{r}"); ws.row_dimensions[r].height = 10
body_end = r
r += 1

# Total
total = r
ws.merge_cells(f"A{r}:B{r}")
put(f"A{r}", "TOTAL AMOUNT", 11, True, h="center")
put(f"C{r}", f"={CUR}", 11, True, h="center")
put(f"D{r}", f"=D{sub}", 11, True, h="right", fmt=AMT)
for col in "ABCD":
    ws[f"{col}{r}"].fill = light
ws.row_dimensions[r].height = 22

# Borders for table area (top .. total)
for row in range(top, total + 1):
    for col in "ABCD":
        L = med if col == "A" else (thin if col in "CD" else None)
        R = med if col == "D" else None
        T = med if row == top else (thin if row in (top + 1, total) else None)
        B = med if row == total else (thin if row == top else None)
        ws[f"{col}{row}"].border = Border(left=L, right=R, top=T, bottom=B)

# ---------- Account details ----------
r = total + 2
put(f"A{r}", "ACCOUNT DETAILS", 10, True)
for col in "ABCD":
    ws[f"{col}{r}"].border = Border(bottom=thin)
r += 1
acct = [
    ("BANK NAME", "Bank name (Branch)"),
    ("ACCOUNT NO.", "000-000-000000"),
    ("BENEFICIARY", "Account holder name"),
    ("SWIFT CODE", "XXXXXXXX"),
]
for label, val in acct:
    put(f"A{r}", label, 10, color="404040", indent=1)
    ws.merge_cells(f"B{r}:D{r}")
    put(f"B{r}", val, 10)
    ws.row_dimensions[r].height = 18
    r += 1
last_row = r - 1

# ---------- Print setup ----------
ws.print_area = f"A1:D{last_row}"
ws.page_setup.paperSize = ws.PAPERSIZE_A4
ws.page_setup.orientation = "portrait"
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 1
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.page_margins = PageMargins(left=0.6, right=0.6, top=0.7, bottom=0.7)
ws.print_options.horizontalCentered = True

# ---------- Guide sheet ----------
g = wb.create_sheet("작성 안내")
g.sheet_view.showGridLines = False
g.column_dimensions["A"].width = 4
g.column_dimensions["B"].width = 90
KF = "Malgun Gothic"
lines = [
    ("Invoice 작성 안내", True),
    ("", False),
    ("1. 상단 회사명 / 주소 / TEL / FAX / E-MAIL 을 자사 정보로 바꿔 주세요 (한 번만 수정하면 계속 사용).", False),
    ("2. BILL NO., OWNER, CC, YOUR REF. NO., RE, DATE 를 건별로 입력합니다.", False),
    ("3. CURRENCY 칸(KRW/USD 등)을 바꾸면 표의 통화 표시가 모두 같이 바뀝니다.", False),
    ("4. DESCRIPTION 에 항목을, 오른쪽 AMOUNT 칸에 금액을 입력하세요. 금액이 있는 줄에만 통화가 표시됩니다.", False),
    ("   (현재 입력된 STORE PARTS / PACKING / TRUCKING 항목은 예시입니다. 빈 줄 3개는 추가 항목용입니다.)", False),
    ("5. DISCOUNT 는 양수로 입력하세요 (예: 50000 → '-50,000' 으로 표시되고 합계에서 차감). 없으면 0.", False),
    ("6. SUB TOTAL 과 TOTAL AMOUNT 는 자동 계산됩니다 (수식 칸이므로 직접 입력하지 마세요).", False),
    ("7. 하단 ACCOUNT DETAILS 에 입금 계좌 정보를 입력하세요.", False),
    ("8. 인쇄/PDF 저장 시 A4 한 장에 맞춰 출력되도록 설정되어 있습니다.", False),
    ("", False),
    ("※ 항목 줄을 더 추가하려면 빈 줄 사이에 '행 삽입' 하면 SUB TOTAL 합계 범위가 자동으로 늘어납니다.", False),
]
for i, (t, b) in enumerate(lines, start=2):
    c = g[f"B{i}"]
    c.value = t
    c.font = Font(name=KF, size=14 if b else 10, bold=b)

out = "/home/user/github-test/invoice/invoice_template.xlsx"
wb.save(out)
print(out, "top", top, "items", first_item, last_item, "disc", disc, "sub", sub, "total", total)
