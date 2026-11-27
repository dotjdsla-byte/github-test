"""Build 4 logbook-supply invoices (one per vessel, pre-rename vessel names) for 고려문화사."""
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.worksheet.page import PageMargins

OUT = "/home/user/github-test/invoice/koryo"
INVOICE_DATE = date(2026, 11, 27)  # 발행일: 필요 시 수정

SELLER = {
    "name": "KORYO MUNHWASA (고려문화사)",
    "addr1": "2F, 260-10, Euljiro 3-ga, Jung-gu,",
    "addr2": "Seoul, Republic of Korea",
    "ceo": "KIM CHAE GI (김채기)",
    "email": "212happy@naver.com",
}
BANK = [
    ("BANK NAME", "STANDARD CHARTERED BANK KOREA (SC제일은행)  (BRANCH: EULJIRO)"),
    ("ACCOUNT NO.", "108-20-133606"),
    ("BENEFICIARY", "KORYO MUNHWASA KIM CHAE GI (고려문화사 김채기)"),
    ("SWIFT CODE", "SCBLKRSE"),
]
ITEMS = [
    ("1. SUPPLY OF LOGBOOKS", 316800),
    ("2. FREIGHT CHARGE FOR LOGBOOK DELIVERY", 8800),
]
VESSELS = [  # (선명 - 변경 전, 선주)
    ("NORTH LIGHT", "ASTERON PTE. LTD."),
    ("NORTH MOON", "NORVEX PTE. LTD."),
    ("NORTH OCEAN", "TRIVANTA PTE. LTD."),
    ("NORTH VALLEY", "VELLORI PTE. LTD."),
]

F = "Arial"
thin = Side(style="thin", color="000000")
med = Side(style="medium", color="000000")
grey = PatternFill("solid", fgColor="D9D9D9")
light = PatternFill("solid", fgColor="F2F2F2")
AMT = '#,##0;-#,##0;"-"'


def build(bill_no, vessel, owner):
    wb = Workbook()
    ws = wb.active
    ws.title = "INVOICE"
    ws.sheet_view.showGridLines = False
    for k, v in {"A": 17, "B": 46, "C": 13, "D": 18}.items():
        ws.column_dimensions[k].width = v

    def put(cell, value, size=10, bold=False, h="left", indent=None, color="000000", fmt=None):
        c = ws[cell]
        c.value = value
        c.font = Font(name=F, size=size, bold=bold, color=color)
        if indent is None:
            indent = 1 if h == "right" else 0
        c.alignment = Alignment(horizontal=h, vertical="center", indent=indent)
        if fmt:
            c.number_format = fmt
        return c

    # Header
    ws.merge_cells("A1:D1")
    put("A1", SELLER["name"], 18, True)
    ws.row_dimensions[1].height = 30
    put("A2", SELLER["addr1"], 9, color="404040")
    put("A3", SELLER["addr2"], 9, color="404040")
    put("C2", "CEO", 9, h="right", color="404040"); put("D2", SELLER["ceo"], 9)
    put("C3", "E-MAIL", 9, h="right", color="404040"); put("D3", SELLER["email"], 9)
    for col in "ABCD":
        ws[f"{col}5"].border = Border(bottom=thin)
    ws.row_dimensions[6].height = 8
    ws.merge_cells("A7:D7")
    put("A7", "INVOICE", 18, True, h="center")
    ws["A7"].fill = grey
    ws.row_dimensions[7].height = 32
    ws.row_dimensions[8].height = 8

    # Info
    r = 9
    for label, val in [("BILL NO.", bill_no), ("OWNER", owner), ("CC", "-"),
                       ("YOUR REF. NO.", "-"), ("RE", f"M/V {vessel}")]:
        put(f"A{r}", f"{label} :", 10, True)
        ws.merge_cells(f"B{r}:D{r}")
        put(f"B{r}", val, 10)
        r += 1
    put(f"A{r}", "DATE :", 10, True)
    put(f"B{r}", INVOICE_DATE, 10, fmt="yyyy-mm-dd")
    put(f"C{r}", "CURRENCY :", 10, True, h="right")
    put(f"D{r}", "KRW", 10, True, h="center")
    CUR = f"$D${r}"
    ws.row_dimensions[r + 1].height = 10

    # Table
    top = r + 2
    ws.merge_cells(f"A{top}:B{top}")
    put(f"A{top}", "DESCRIPTION", 10, True, h="center")
    ws.merge_cells(f"C{top}:D{top}")
    put(f"C{top}", "AMOUNT", 10, True, h="center")
    for col in "ABCD":
        ws[f"{col}{top}"].fill = light
    ws.row_dimensions[top].height = 20

    rows = [("", None), ("PLEASE SETTLE THE AMOUNT BELOW TO OUR ACCOUNT.", None), ("", None)]
    for desc, amt in ITEMS:
        rows += [(desc, amt), ("", None)]
    rows += [("", None)] * 4  # spare lines

    r = top + 1
    first = r
    for desc, amt in rows:
        ws.merge_cells(f"A{r}:B{r}")
        put(f"A{r}", desc or None, 10, bool(amt), indent=1)
        put(f"D{r}", amt, 10, h="right", fmt=AMT)
        put(f"C{r}", f'=IF(D{r}="","",{CUR})', 10, h="center")
        ws.row_dimensions[r].height = 17
        r += 1
    last = r - 1

    ws.merge_cells(f"A{r}:B{r}")
    put(f"A{r}", f"{len(ITEMS) + 1}. DISCOUNT", 10, True, indent=1)
    put(f"C{r}", f"={CUR}", 10, h="center")
    put(f"D{r}", 0, 10, h="right", fmt='"-"#,##0;"-"#,##0;"-"')
    disc = r
    r += 1
    ws.merge_cells(f"A{r}:B{r}"); ws.row_dimensions[r].height = 10; r += 1
    ws.merge_cells(f"A{r}:B{r}")
    put(f"A{r}", f"{len(ITEMS) + 2}. SUB TOTAL", 10, True, indent=1)
    put(f"C{r}", f"={CUR}", 10, h="center")
    put(f"D{r}", f"=SUM(D{first}:D{last})-D{disc}", 10, True, h="right", fmt=AMT)
    sub = r
    r += 1
    ws.merge_cells(f"A{r}:B{r}"); ws.row_dimensions[r].height = 10; r += 1
    total = r
    ws.merge_cells(f"A{r}:B{r}")
    put(f"A{r}", "TOTAL AMOUNT", 11, True, h="center")
    put(f"C{r}", f"={CUR}", 11, True, h="center")
    put(f"D{r}", f"=D{sub}", 11, True, h="right", fmt=AMT)
    for col in "ABCD":
        ws[f"{col}{r}"].fill = light
    ws.row_dimensions[r].height = 22

    for row in range(top, total + 1):
        for col in "ABCD":
            L = med if col == "A" else (thin if col in "CD" else None)
            R = med if col == "D" else None
            T = med if row == top else (thin if row in (top + 1, total) else None)
            B = med if row == total else (thin if row == top else None)
            ws[f"{col}{row}"].border = Border(left=L, right=R, top=T, bottom=B)

    # Account details
    r = total + 2
    put(f"A{r}", "ACCOUNT DETAILS", 10, True)
    for col in "ABCD":
        ws[f"{col}{r}"].border = Border(bottom=thin)
    r += 1
    for label, val in BANK:
        put(f"A{r}", label, 10, color="404040", indent=1)
        ws.merge_cells(f"B{r}:D{r}")
        put(f"B{r}", val, 10)
        ws.row_dimensions[r].height = 18
        r += 1

    ws.print_area = f"A1:D{r - 1}"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.6, right=0.6, top=0.7, bottom=0.7)
    ws.print_options.horizontalCentered = True

    path = f"{OUT}/INVOICE_{vessel.replace(' ', '_')}.xlsx"
    wb.save(path)
    return path


if __name__ == "__main__":
    for i, (vessel, owner) in enumerate(VESSELS, start=1):
        print(build(f"KRM-{INVOICE_DATE:%y}-{i:03d}", vessel, owner))
