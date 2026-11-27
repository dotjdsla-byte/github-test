"""Delivery Note / Receipt x4 (one per vessel, pre-rename names) - IS530.1 camera 1 SET each."""
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.worksheet.page import PageMargins

OUT = "/home/user/github-test/invoice/koryo/delivery_note"
DATE = "2026-09-28"          # 납품일: 직접 수정
PLACE = ""                   # 납품 장소: 직접 입력
REF_NO = "GOS-26-0927"
ORDER_DATE = "2026-09-28"
VESSELS = ["NORTH LIGHT", "NORTH MOON", "NORTH OCEAN", "NORTH VALLEY"]
ITEMS = [  # (no, part no, description, remark, qty, unit)
    (1, "NONE", "Intrinsically Safe SmartPhone Camera", "Model: IS530.1", 1, "SET"),
]
SELLER = {
    "name": "KORYO MUNHWASA (고려문화사)",
    "addr": "2F, 260-10, Euljiro 3-ga, Jung-gu, Seoul, Republic of Korea",
    "contact": "CEO : KIM CHAE GI (김채기)     E-MAIL : 212happy@naver.com",
}

F = "Arial"
BLUE = "1F5FBF"
thin = Side(style="thin", color="000000")
hair = Side(style="hair", color="999999")
med = Side(style="medium", color="000000")
dark = PatternFill("solid", fgColor="404040")
COLS = "ABCDEF"


def build(vessel):
    wb = Workbook()
    ws = wb.active
    ws.title = "DELIVERY NOTE"
    ws.sheet_view.showGridLines = False
    for k, v in zip(COLS, [6, 14, 52, 7, 7, 16]):
        ws.column_dimensions[k].width = v

    def put(cell, value, size=10, bold=False, h="left", color="000000", fill=None, wrap=False, indent=0):
        c = ws[cell]
        c.value = value
        c.font = Font(name=F, size=size, bold=bold, color=color)
        c.alignment = Alignment(horizontal=h, vertical="center", wrap_text=wrap, indent=indent)
        if fill:
            c.fill = fill
        return c

    def merge(rng):
        ws.merge_cells(rng)

    # Header
    merge("A1:F1"); put("A1", SELLER["name"], 18, True, "center"); ws.row_dimensions[1].height = 30
    merge("A2:F2"); put("A2", SELLER["addr"], 9, h="center", color="404040")
    merge("A3:F3"); put("A3", SELLER["contact"], 9, h="center", color="404040")
    merge("A4:F4"); put("A4", "#### Non-use of Materials containing asbestos ####", 9, h="center", color=BLUE)
    ws.row_dimensions[5].height = 8
    merge("B6:E6"); put("B6", "DELIVERY NOTE / RECEIPT", 16, True, "center", "FFFFFF", dark)
    for col in "BCDE":
        ws[f"{col}6"].fill = dark
    ws.row_dimensions[6].height = 30
    ws.row_dimensions[7].height = 8

    put("D8", "DATE :", 10, True); merge("E8:F8"); put("E8", DATE, 10)
    put("A9", "VESSEL NAME :", 10, True); merge("A9:B9")
    put("C9", vessel, 11, True, indent=1)
    put("D9", "PLACE :", 10, True); merge("E9:F9"); put("E9", PLACE, 10)
    ws.row_dimensions[10].height = 6

    # Table header
    top = 11
    for col, t in zip(COLS, ["NO.", "PART NO", "DESCRIPTION", "Q'TY", "", "REMARK"]):
        put(f"{col}{top}", t, 10, True, "center", "FFFFFF", dark)
    merge(f"D{top}:E{top}")
    ws.row_dimensions[top].height = 20

    r = top + 1
    lines = [("", None)]
    lines += [(f"REF NO. : {REF_NO}", "blue"), (f"ORDER DATE. : {ORDER_DATE}", "blue")]
    lines += [(f"MACHINERY: 선명: {vessel}", "head"), ("MAKER: -", "head"),
              ("TYPE: Intrinsically Safe SmartPhone Camera", "head")]
    for text, kind in lines:
        if kind == "blue":
            merge(f"A{r}:C{r}")
            put(f"A{r}", text, 10, True, color=BLUE, indent=1)
        elif kind == "head":
            put(f"C{r}", text, 10, True, indent=1)
        ws.row_dimensions[r].height = 17
        r += 1

    for no, part, desc, remark, qty, unit in ITEMS:
        for col in "ABDEF":
            merge(f"{col}{r}:{col}{r + 1}")
        put(f"A{r}", no, 10, h="center")
        put(f"B{r}", part, 10, indent=1)
        put(f"C{r}", desc, 10, indent=1)
        put(f"C{r + 1}", f"Remark: {remark}", 10, color=BLUE, indent=1)
        put(f"D{r}", qty, 10, h="center")
        put(f"E{r}", unit, 10, h="center")
        ws.row_dimensions[r].height = 17
        ws.row_dimensions[r + 1].height = 17
        r += 2

    blank_end = r + 18
    for rr in range(r, blank_end):
        ws.row_dimensions[rr].height = 17
    last = blank_end - 1

    # Borders: vertical column lines + faint row lines
    vlines = {"A": (med, thin), "B": (thin, thin), "C": (thin, thin), "D": (thin, hair),
              "E": (hair, thin), "F": (thin, med)}
    for rr in range(top, last + 1):
        for col in COLS:
            L, R = vlines[col]
            T = med if rr == top else None
            B = med if rr == last else hair
            ws[f"{col}{rr}"].border = Border(left=L, right=R, top=T, bottom=B)

    # Signature box (receipt)
    s = last + 1
    for rr in range(s, s + 5):
        ws.row_dimensions[rr].height = 18
        for col in COLS:
            ws[f"{col}{rr}"].border = Border(
                left=med if col == "A" else None, right=med if col == "F" else None,
                bottom=med if rr == s + 4 else None)
    put(f"A{s}", "RECEIVED THE ABOVE ITEMS IN GOOD ORDER AND CONDITION.", 9, color="404040", indent=1)
    merge(f"A{s}:C{s}")
    merge(f"D{s + 3}:F{s + 3}")
    for col in "DEF":
        ws[f"{col}{s + 3}"].border = Border(bottom=thin, right=med if col == "F" else None)
    merge(f"D{s + 4}:F{s + 4}")
    put(f"D{s + 4}", "MASTER / CHIEF ENGINEER", 10, True, "center")
    ws[f"F{s + 4}"].border = Border(right=med, bottom=med)
    ws[f"D{s + 4}"].border = Border(bottom=med)
    ws[f"E{s + 4}"].border = Border(bottom=med)

    merge(f"A{s + 6}:F{s + 6}")
    put(f"A{s + 6}", "(1 - 1)", 9, h="center", color="404040")

    ws.print_area = f"A1:F{s + 6}"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.5, right=0.5, top=0.6, bottom=0.6)
    ws.print_options.horizontalCentered = True

    path = f"{OUT}/DELIVERY_NOTE_{vessel.replace(' ', '_')}.xlsx"
    wb.save(path)
    return path


if __name__ == "__main__":
    for v in VESSELS:
        print(build(v))
