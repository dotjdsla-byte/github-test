# -*- coding: utf-8 -*-
"""CARGO PART DAILY LOG - A3 landscape (left = page 32, right = page 33)."""
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

OUT = "/home/user/github-test/forms/CARGO_PART_DAILY_LOG_A3.xlsx"

FONT = "Arial"
THIN = Side(style="thin", color="000000")
MED = Side(style="medium", color="000000")

# ---------------- column map ----------------
A1, A2, A3, A4 = 1, 2, 3, 4          # page32 left block label cols
AV = [5, 6, 7, 8]                     # page32 left block value cols (tank 1..4)
B1, B2, B3 = 9, 10, 11                # page32 right block label cols
BV = [12, 13, 14, 15]                 # page32 right block value cols (tank 1..4)
GUT = 16                              # gutter / fold
C1, C2, C3, CV = 17, 18, 19, 20       # page33 left block
D1, D2, D3, D4 = 21, 22, 23, 24       # page33 right block label cols
DV = list(range(25, 35))              # page33 right block: 10 fine value cols

# Widths are sized so the sheet prints at 100% (no shrink) on A3 landscape:
# LibreOffice/Excel renders ~7.27 pt per width unit here, and the A3 printable
# width at 0.2" margins is ~1154 pt, i.e. a budget of ~158 units.
WIDTHS = {
    A1: 2.3, A2: 3.4, A3: 7.0, A4: 10.8,
    AV[0]: 5.1, AV[1]: 5.1, AV[2]: 5.1, AV[3]: 5.1,
    B1: 4.6, B2: 7.0, B3: 8.8,
    BV[0]: 5.0, BV[1]: 5.0, BV[2]: 5.0, BV[3]: 5.0,
    GUT: 1.4,
    C1: 1.9, C2: 4.2, C3: 13.8, CV: 8.6,
    D1: 1.9, D2: 4.0, D3: 5.2, D4: 13.0,
}
for c in DV:
    WIDTHS[c] = 1.9

LEFT_FIRST, LEFT_LAST = A1, BV[3]
RIGHT_FIRST, RIGHT_LAST = C1, DV[-1]

R_TABLE = 4          # first table row
NROWS = 42           # rows of the tallest block (page 32 right block)
R_LAST = R_TABLE + NROWS - 1   # 45

wb = Workbook()
ws = wb.active
ws.title = "CARGO PART DAILY LOG"

for col, w in WIDTHS.items():
    ws.column_dimensions[get_column_letter(col)].width = w

# ---------------- helpers ----------------
CEN = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEF = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
VERT = Alignment(horizontal="center", vertical="center", textRotation=255, wrap_text=False)

def put(r, c, v, align=CEN, size=8, bold=False, rows=1, cols=1):
    """write value at (r,c), merging `rows` x `cols` starting there."""
    if rows > 1 or cols > 1:
        ws.merge_cells(start_row=r, start_column=c, end_row=r + rows - 1, end_column=c + cols - 1)
    cell = ws.cell(row=r, column=c, value=v)
    cell.font = Font(name=FONT, size=size, bold=bold)
    cell.alignment = align
    return cell

def merge_cols(r, c1, c2, v, align=CEN, size=8, bold=False, rows=1):
    return put(r, c1, v, align=align, size=size, bold=bold, rows=rows, cols=c2 - c1 + 1)

# =========================================================
# HEADER
# =========================================================
ws.row_dimensions[1].height = 26
ws.row_dimensions[2].height = 16
ws.row_dimensions[3].height = 3

merge_cols(1, A1, A2, "PAGE", size=8, bold=False)
put(1, A3, 32, size=10, bold=True)
merge_cols(1, A4, BV[1], "CARGO PART DAILY LOG", size=18, bold=True)
merge_cols(1, BV[2], BV[3], "DATE :", align=Alignment(horizontal="left", vertical="bottom"), size=10, bold=True)

merge_cols(2, A1, A3, "SHIP'S NAME :", align=Alignment(horizontal="left", vertical="bottom"), size=10, bold=True)
merge_cols(2, A4, AV[3], None)
merge_cols(2, B1, B2, "VOY. NO. :", align=Alignment(horizontal="left", vertical="bottom"), size=10, bold=True)
merge_cols(2, B3, BV[3], None)

merge_cols(1, DV[4], DV[6], "PAGE", size=8)
merge_cols(1, DV[7], DV[9], 33, size=10, bold=True)
merge_cols(2, C1, C2, "FROM :", align=Alignment(horizontal="left", vertical="bottom"), size=10, bold=True)
merge_cols(2, C3, CV, None)
merge_cols(2, D1, D2, "TO :", align=Alignment(horizontal="left", vertical="bottom"), size=10, bold=True)
merge_cols(2, D3, DV[9], None)

# underline for the fill-in blanks in the header
def underline(r, c1, c2):
    for c in range(c1, c2 + 1):
        cell = ws.cell(row=r, column=c)
        cell.border = Border(bottom=THIN)

underline(1, BV[2], BV[3])
underline(2, A4, AV[3])
underline(2, B3, BV[3])
underline(2, C3, CV)
underline(2, D3, DV[9])

for r in range(R_TABLE, R_LAST + 1):
    ws.row_dimensions[r].height = 18.2

# =========================================================
# BLOCK A : page 32 - left block
# =========================================================
r = R_TABLE
# atmospheric conditions
for txt in ("ATMOS. Press (mb.A)", "ATMOS. Temp. (℃)", "SEA WATER Temp. (℃)"):
    merge_cols(r, A1, A4, txt, align=LEF)
    merge_cols(r, AV[0], AV[3], None)
    r += 1
# tank number header
merge_cols(r, A1, A4, None)
for i, n in enumerate((1, 2, 3, 4)):
    put(r, AV[i], n, bold=True)
r += 1

put(r, A1, "TRUNK\nDK", rows=2, cols=2, size=7)
put(r, A3, "Temp.(℃)", rows=2)
put(r, A4, "AFT", align=LEF)
put(r + 1, A4, "MID", align=LEF)
r += 2

INSUL_ITEMS = ["CEILING", "Port Upp. Cham", "Port Mid Side W", "Port Low. Cham",
               "Btm Aft Mid", "Btm Mid", "Tran Wall Fwd", "Tran Wall Aft"]
put(r, A1, "INSUL.\nSPACE", rows=10, cols=2, size=7)
put(r, A3, "Temp.(℃)", rows=8)
for i, t in enumerate(INSUL_ITEMS):
    put(r + i, A4, t, align=LEF)
put(r + 8, A3, "Press\n(mbar)", size=7)
put(r + 8, A4, "IS N2", align=LEF)
put(r + 9, A3, "Diff. Press\n(mbar)", size=7)
put(r + 9, A4, "IS/IBS N2", align=LEF)
r += 10

put(r, A1, "IBS", rows=1, cols=2, size=7)
put(r, A3, "Press\n(mbar)", size=7)
put(r, A4, "IBS N2", align=LEF)
r += 1

put(r, A1, "SIDE\nWALL", rows=1, cols=2, size=7)
put(r, A3, "Temp.(℃)")
put(r, A4, "PORT", align=LEF)
r += 1

put(r, A1, "DUCT\nKEEL", rows=2, cols=2, size=7)
put(r, A3, "Temp.(℃)", rows=2)
put(r, A4, "MID", align=LEF)
put(r + 1, A4, "AFT", align=LEF)
r += 2

# ---- SMR ----
smr_start = r
put(r, A1, "SMR", rows=20, align=VERT)

put(r, A2, "BOG", rows=3, align=VERT)
put(r, A3, "Flow\n(kg/h)", size=7)
put(r + 1, A3, "Press (bar)", size=7)
put(r + 2, A3, "Temp.(℃)", size=7)
put(r, A4, "PFHE Inlet", rows=3)
r += 3

put(r, A2, "PFHE", rows=4, align=VERT)
put(r, A3, "Temp.(℃)", rows=4)
for i, t in enumerate(("Top", "Upper Mid", "Lower Mid", "Bottom")):
    put(r + i, A4, t, align=LEF)
r += 4

put(r, A2, "LIQ.", rows=2, align=VERT, size=7)
put(r, A3, "Temp.(℃)", size=7)
put(r + 1, A3, "Flow\n(kg/h)", size=7)
put(r, A4, "PFHE Outlet", rows=2)
r += 2

put(r, A2, "MR COMP.", rows=11, align=VERT, size=7)
put(r, A3, "Press (bar)", size=7)
put(r + 1, A3, "Temp.(℃)", size=7)
put(r, A4, "SUCTION", rows=2)
put(r + 2, A3, "Press (bar)", size=7)
put(r + 3, A3, "Temp.(℃)", size=7)
put(r + 2, A4, "DISCHARGE", rows=2)
put(r + 4, A3, "Temp.(℃)", size=7)
put(r + 4, A4, "AFT. CLR.", align=LEF)
put(r + 5, A3, "Temp.(℃)", rows=2, size=7)
put(r + 5, A4, "OIL CLR. Inlet", align=LEF)
put(r + 6, A4, "OIL CLR. Outlet", align=LEF)
merge_cols(r + 7, A3, A4, "Power (kW)", align=LEF)
merge_cols(r + 8, A3, A4, "Ampere (A)", align=LEF)
merge_cols(r + 9, A3, A4, "Capacity (%)", align=LEF)
put(r + 10, A3, "Press (bar)", size=7)
put(r + 10, A4, "C.F.W", align=LEF)
r += 11
assert r - smr_start == 20, r - smr_start

merge_cols(r, A1, A4, "FORE VENT VALVE SET Press (mb.G)", align=LEF, size=7.5)
merge_cols(r, AV[0], AV[3], None)
r += 1
A_LAST = r - 1

# every remaining value cell of block A
for rr in range(R_TABLE + 4, A_LAST):
    for c in AV:
        ws.cell(row=rr, column=c)

# =========================================================
# BLOCK B : page 32 - right block
# =========================================================
r = R_TABLE
merge_cols(r, B1, B3, None)
for i, n in enumerate((1, 2, 3, 4)):
    put(r, BV[i], n, bold=True)
r += 1

put(r, B1, "CARGO\nTANK", rows=8)
merge_cols(r, B2, B3, "Press (mb.G)", align=LEF)
merge_cols(r + 1, B2, B3, "LEVEL (mtr)", align=LEF)
merge_cols(r + 2, B2, B3, "LIQ. VOL. (M³)", align=LEF)
put(r + 3, B2, "Temp.(℃)", rows=5)
for i, t in enumerate(("TOP", "95%", "60%", "30%", "BTM")):
    put(r + 3 + i, B3, t)
r += 8

put(r, B1, "VAP\nHDR", rows=3)
merge_cols(r, B2, B3, "Press (Gauge/ABS.)", align=LEF)
merge_cols(r + 1, B2, B3, "X-OVER Press (bar.G)", align=LEF)
merge_cols(r + 2, B2, B3, "X-OVER Temp. (℃)", align=LEF)
r += 3

put(r, B1, "LIQ.\nHDR.", rows=4)
put(r, B2, "Press\n(bar.G)", rows=2)
put(r, B3, "FWD")
put(r + 1, B3, "AFT")
put(r + 2, B2, "Temp.(℃)", rows=2)
put(r + 2, B3, "FWD")
put(r + 3, B3, "AFT")
r += 4

put(r, B1, "STRIP\nHDR.", rows=2)
put(r, B2, "Temp.(℃)", rows=2)
put(r, B3, "FWD")
put(r + 1, B3, "AFT")
r += 2

merge_cols(r, B1, B3, "VAP. FLOW TO ATM.(ton/h)", align=LEF)
r += 1

put(r, B1, "INSUL.\nN2\nSYS.", rows=5)
put(r, B2, "press", rows=3)
put(r, B3, "BLEED (bar.G)")
put(r + 1, B3, "TO IBS (mb.G)")
put(r + 2, B3, "TO IS (mb.G)")
merge_cols(r + 3, B2, B3, "VAP.HDR./PRI.HDR.DIF.P.", align=LEF, size=7)
merge_cols(r + 4, B2, B3, "FLOW RATE (NM³/H)", align=LEF)
r += 5

merge_cols(r, B1, B3, "WATER DET. SYS. COND.", align=LEF)
r += 1

put(r, B1, "N2\nGEN.", rows=10)
merge_cols(r, B2, B3, "RUN NO.", align=CEN)
put(r + 1, B2, "Press\n(bar.G)", rows=3)
put(r + 1, B3, "BUFFER TANK")
put(r + 2, B3, "FEED AIR")
put(r + 3, B3, "FILTER Outlet")
put(r + 4, B2, "Temp.(℃)", rows=3)
put(r + 4, B3, "COMP. Outlet")
put(r + 5, B3, "FEED AIR")
put(r + 6, B3, "HEATER Outlet")
merge_cols(r + 7, B2, B3, "O2 CONTENT (%)", align=LEF)
merge_cols(r + 8, B2, B3, "DEW POINT (℃)", align=LEF)
merge_cols(r + 9, B2, B3, "NITRIGEN FLOW (nm3/h)", align=LEF)
r += 10

put(r, B1, "GAS\nDETEC.", rows=7)
put(r, B2, None)
put(r, B3, "TANK")
for i, t in enumerate(("IBS\nGasdome", "IBS\nLiq.dome", "IS\nGas dome", "VENT\nMAST")):
    put(r, BV[i], t, size=7)
put(r + 1, B2, "VOLUME (%)", rows=4)
for i in range(4):
    put(r + 1 + i, B3, i + 1)
merge_cols(r + 5, B2, B3, "GAS VENT DRAIN TANK", align=LEF)
merge_cols(r + 5, BV[0], BV[3], None)
merge_cols(r + 6, B2, B3, "HP VENT MAST", align=LEF)
merge_cols(r + 6, BV[0], BV[3], None)
r += 7
B_LAST = r - 1
assert B_LAST == R_LAST, (B_LAST, R_LAST)

# =========================================================
# BLOCK C : page 33 - left block (L/D COMP. gas press / temp)
# =========================================================
STAGES = ["GAS INLET ST1", "GAS DISCH ST1", "GAS INLET ST2", "GAS DISCH ST2",
          "GAS INLET ST3", "GAS DISCH ST3", "GAS INLET ST4", "GAS DISCH ST4",
          "GAS AFTER COOLER"]
r = R_TABLE
put(r, C1, "L/D COMP.", rows=24, align=VERT, size=7)
merge_cols(r, C2, C3, "RUN NO.")
r += 1
put(r, C2, "Press\n(bar)", rows=9)
for i, t in enumerate(STAGES):
    put(r + i, C3, t, align=LEF)
r += 9
put(r, C2, "Temp.\n(℃)", rows=9)
for i, t in enumerate(STAGES):
    put(r + i, C3, t, align=LEF)
r += 9
put(r, C2, "Flow\n(kg/h)", rows=2)
put(r, C3, "ST1 INLET", align=LEF)
put(r + 1, C3, "ST4 OUTLET", align=LEF)
r += 2
put(r, C2, "VALVE", rows=3)
for i, t in enumerate(("VDV POSITION (%)", "ST4 TO ST2 ASV (%)", "ST4 TO MIST ASV (%)")):
    put(r + i, C3, t, align=LEF)
r += 3
C_LAST = r - 1

# =========================================================
# BLOCK D : page 33 - right block
# =========================================================
r = R_TABLE
put(r, D1, "GLYCOL WATER SYSTEM", rows=12, align=VERT, size=7)
put(r, D2, "LEVEL", rows=2)
merge_cols(r, D3, D4, "RESER. TANK (cm)", align=LEF)
merge_cols(r + 1, D3, D4, "EXP. TANK (cm)", align=LEF)
put(r + 2, D2, "STM.\nHTR.", rows=7)
merge_cols(r + 2, D3, D4, "RUN NO.")
put(r + 3, D3, "Press\n(bar.G)", rows=3)
for i, t in enumerate(("INLET", "OUTLET", "HEATING STEAM")):
    put(r + 3 + i, D4, t, align=LEF)
put(r + 6, D3, "Temp.\n(℃)", rows=3)
for i, t in enumerate(("INLET", "OUTLET", "CONT.V. OUTLET")):
    put(r + 6 + i, D4, t, align=LEF)
r += 9
# coffer dam : 5 columns across the value area (2 fine cols each)
merge_cols(r, D2, D4, "COFFER DAM")
for i in range(5):
    put(r, DV[i * 2], i + 1, cols=2, bold=True)
put(r + 1, D2, "Temp.\n(℃)", rows=2, cols=2)
put(r + 1, D4, "AVE.")
put(r + 2, D4, "MIN.")
for k in (1, 2):
    for i in range(5):
        put(r + k, DV[i * 2], None, cols=2)
r += 3

# ---- vapourizer ----
put(r, D1, "VAPOURIZER", rows=7, align=VERT, size=7)
merge_cols(r, D2, D4, "RUN NO.")
put(r, DV[0], "LNG", cols=5, bold=True)
put(r, DV[5], "FORCING", cols=5, bold=True)
put(r + 1, D2, "Press\n(bar.G)", rows=3, cols=2)
for i, t in enumerate(("INLET GAS", "OUTLET GAS", "INLET STEAM")):
    put(r + 1 + i, D4, t, align=LEF)
put(r + 4, D2, "Temp.\n(℃)", rows=3, cols=2)
for i, t in enumerate(("INLET GAS", "OUTLET GAS", "CONDENSATE")):
    put(r + 4 + i, D4, t, align=LEF)
for k in range(1, 7):
    put(r + k, DV[0], None, cols=5)
    put(r + k, DV[5], None, cols=5)
r += 7
merge_cols(r, D1, D4, "CONDENSATE LEVEL (cm)", align=LEF)
put(r, DV[0], None, cols=5)
put(r, DV[5], None, cols=5)
r += 1

# ---- L/D comp. motor & LO ----
put(r, D1, "L/D COMP.", rows=5, align=VERT, size=7)
put(r, D2, "MOTOR", rows=3, cols=2)
for i, t in enumerate(("MOTOR AMP. (A)", "WIND Temp.(U/V/W,℃)", "BRG (DE/NDE,℃)")):
    put(r + i, D4, t, align=LEF)
put(r + 3, D2, "LO", rows=2, cols=2)
put(r + 3, D4, "Press (bar)", align=LEF)
put(r + 4, D4, "Temp. (℃)", align=LEF)
for i in range(5):
    put(r + i, DV[0], None, cols=10)
r += 5
D_LAST = r - 1

# value cells for block C and the glycol part of block D
for rr in range(R_TABLE, C_LAST + 1):
    ws.cell(row=rr, column=CV)
for rr in range(R_TABLE, R_TABLE + 9):
    put(rr, DV[0], None, cols=10)

# =========================================================
# REMARKS area (page 33, below the tables)
# =========================================================
REM_FIRST = D_LAST + 1
for rr in range(REM_FIRST, R_LAST + 1):
    merge_cols(rr, RIGHT_FIRST, RIGHT_LAST, None)

# =========================================================
# BORDERS
# =========================================================
def box(r1, c1, r2, c2):
    """thin grid inside + medium outline"""
    for rr in range(r1, r2 + 1):
        for cc in range(c1, c2 + 1):
            cell = ws.cell(row=rr, column=cc)
            cell.border = Border(
                left=MED if cc == c1 else THIN,
                right=MED if cc == c2 else THIN,
                top=MED if rr == r1 else THIN,
                bottom=MED if rr == r2 else THIN,
            )

box(R_TABLE, A1, A_LAST, AV[3])
box(R_TABLE, B1, B_LAST, BV[3])
box(R_TABLE, C1, C_LAST, CV)
box(R_TABLE, D1, D_LAST, DV[-1])
box(REM_FIRST, RIGHT_FIRST, R_LAST, RIGHT_LAST)

# strip inner borders of merged ranges so merges render as one cell
for mr in list(ws.merged_cells.ranges):
    r1, r2 = mr.min_row, mr.max_row
    c1, c2 = mr.min_col, mr.max_col
    outer = ws.cell(row=r1, column=c1).border
    ob = ws.cell(row=r2, column=c2).border
    top, left, bottom, right = outer.top, outer.left, ob.bottom, ob.right
    for rr in range(r1, r2 + 1):
        for cc in range(c1, c2 + 1):
            ws.cell(row=rr, column=cc).border = Border(
                top=top if rr == r1 else None,
                bottom=bottom if rr == r2 else None,
                left=left if cc == c1 else None,
                right=right if cc == c2 else None,
            )

# =========================================================
# PAGE SETUP : A3 landscape, one page
# =========================================================
ws.page_setup.paperSize = 8              # A3
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 1
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.page_margins.left = 0.25
ws.page_margins.right = 0.25
ws.page_margins.top = 0.2
ws.page_margins.bottom = 0.2
ws.page_margins.header = 0.1
ws.page_margins.footer = 0.1
ws.print_area = f"A1:{get_column_letter(RIGHT_LAST)}{R_LAST}"
ws.sheet_view.showGridLines = False
ws.freeze_panes = "A4"

wb.save(OUT)
print("saved:", OUT)
print("rows:", R_LAST, "cols:", RIGHT_LAST)
print("total width:", round(sum(WIDTHS.values()), 1))
