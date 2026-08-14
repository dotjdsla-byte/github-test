"""Build the EAL product / maker summary workbook from six SK Shipping lube charts.

Source PDFs (all vessels operated by SK Shipping):
  PUTERI SABAH   IMO 9975521  Shell            Lubrication Survey v2, 15.01.2025
  AL SAKHAMAH    IMO 9986051  TotalEnergies    Lubrication Chart R3, 16-Apr-2025
  BU FINTAS      IMO 9976824  TotalEnergies    Lubrication Chart R2, 07-Aug-2024
  MARVEL DOVE    IMO 9964182  Chevron          Lubrication Chart, rev. 17-Jun-2024
  PRISM AGILITY  IMO 9810549  TotalEnergies    Lubrication Chart R3, 11-Jan-2024
  SK AUDACE      IMO 9693161  TotalEnergies    Lubrication Chart R3, 11-Jan-2024

Every row below is transcribed from those charts; nothing is inferred.
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

KF = "맑은 고딕"
def F(sz=10, b=False, color="000000"):
    return Font(name=KF, size=sz, bold=b, color=color)

THIN = Side(style="thin", color="A6A6A6")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

NAVY  = PatternFill("solid", fgColor="1F3864")
HEADF = PatternFill("solid", fgColor="2E5C8A")
GRPF  = PatternFill("solid", fgColor="D9E2F3")
YEL   = PatternFill("solid", fgColor="FFF2CC")
ALT   = PatternFill("solid", fgColor="F7F9FC")
WARN  = PatternFill("solid", fgColor="FCE4E4")

WRAP = Alignment(horizontal="left", vertical="center", wrap_text=True)
CTR  = Alignment(horizontal="center", vertical="center", wrap_text=True)

# 선박별 EAL 적용 상세 --------------------------------------------------
# (선박, IMO, 공급사, 개정, 장비, 적용부위, 제품, 오일메이커, 장비메이커, 모델, 표기, 비고)
DETAIL = [
    # ---- PUTERI SABAH (Shell) : "(VGP Compliant)" 표기
    ("PUTERI SABAH", "9975521", "Shell", "v2 / 2025-01-15",
     "Accommodation Ladder", "Wire Ropes", "Shell Naturelle S2 Grease A600P 1.5",
     "Shell", "Samgong", "", "(VGP Compliant)", ""),
    ("PUTERI SABAH", "9975521", "Shell", "v2 / 2025-01-15",
     "Pilot Ladder", "Wire Ropes", "Shell Naturelle S2 Grease A600P 1.5",
     "Shell", "Samgong", "", "(VGP Compliant)", ""),
    ("PUTERI SABAH", "9975521", "Shell", "v2 / 2025-01-15",
     "Emergency Towing System", "Towing Wire", "Shell Naturelle S2 Grease A600P 1.5",
     "Shell", "Tanktec", "", "(VGP Compliant)", ""),
    ("PUTERI SABAH", "9975521", "Shell", "v2 / 2025-01-15",
     "Provision Crane", "Open Gears & Wire Ropes", "Shell Naturelle S2 Grease A600P 1.5",
     "Shell", "Sangsangin Industry Co Ltd", "", "(VGP Compliant)", ""),
    ("PUTERI SABAH", "9975521", "Shell", "v2 / 2025-01-15",
     "Compressor Room Crane", "Open Gears & Wire Ropes", "Shell Naturelle S2 Grease A600P 1.5",
     "Shell", "Sangsangin Industry Co Ltd", "", "(VGP Compliant)", ""),
    ("PUTERI SABAH", "9975521", "Shell", "v2 / 2025-01-15",
     "Manifold Handling Crane", "Open Gears & Wire Ropes", "Shell Naturelle S2 Grease A600P 1.5",
     "Shell", "Sangsangin Industry Co Ltd", "", "(VGP Compliant)", ""),
    ("PUTERI SABAH", "9975521", "Shell", "v2 / 2025-01-15",
     "Lifeboat Davit & Winch", "Wire Ropes", "Shell Naturelle S2 Grease A600P 1.5",
     "Shell", "Oriental Precision (OPCO)", "HGD-062-26", "(VGP Compliant)", ""),
    ("PUTERI SABAH", "9975521", "Shell", "v2 / 2025-01-15",
     "General Vessel Lubrication", "Open Gears & Wire Ropes", "Shell Naturelle S2 Grease A600P 1.5",
     "Shell", "(전선 공통)", "", "(VGP Compliant)", ""),
    ("PUTERI SABAH", "9975521", "Shell", "v2 / 2025-01-15",
     "Propeller Bonnet / Cap", "Grease Filling", "Shell Naturelle S5 Grease V120P 2",
     "Shell", "HHI (EMD)", "", "(VGP Compliant)", ""),
    ("PUTERI SABAH", "9975521", "Shell", "v2 / 2025-01-15",
     "General Vessel Lubrication", "Grease Points", "Shell Naturelle S5 Grease V120P 2",
     "Shell", "(전선 공통)", "", "(VGP Compliant)", ""),
    ("PUTERI SABAH", "9975521", "Shell", "v2 / 2025-01-15",
     "Rudder Carrier", "Grease Filling", "Margrease EP 0",
     "확인 필요", "3K Industry", "MSK-602-T5M", "(VGP Compliant)",
     "차트 주석 (B) = Shell Marine 미공급 품목. 제조사 확인 필요"),

    # ---- AL SAKHAMAH (TotalEnergies) : "(EAL)" 표기
    ("AL SAKHAMAH", "9986051", "TotalEnergies\nLubmarine", "R3 / 2025-04-16",
     "Emergency Towing System", "Towing Pennant", "BIOADHESIVE PLUS",
     "TotalEnergies", "(차트상 미기재)", "", "(EAL)", ""),
    ("AL SAKHAMAH", "9986051", "TotalEnergies\nLubmarine", "R3 / 2025-04-16",
     "Rudder Carrier", "Rudder Trunk", "BIOADHESIVE PLUS",
     "TotalEnergies", "FLUTEK", "", "(요약표 EAL 분류)",
     "본문 표에는 (EAL) 표기 없음. 2p 요약표에서 BIOADHESIVE PLUS 항목으로 분류"),
    ("AL SAKHAMAH", "9986051", "TotalEnergies\nLubmarine", "R3 / 2025-04-16",
     "Propeller Cap", "Cap", "BIOMULTIS EP 2",
     "TotalEnergies", "SILLA METAL CO.", "", "(EAL)", ""),

    # ---- BU FINTAS (TotalEnergies) : "(EAL)" 표기
    ("BU FINTAS", "9976824", "TotalEnergies\nLubmarine", "R2 / 2024-08-07",
     "Propeller Cap", "Cap", "BIOMULTIS EP 2",
     "TotalEnergies", "SILLA METAL CO.", "", "(EAL)", ""),
    ("BU FINTAS", "9976824", "TotalEnergies\nLubmarine", "R2 / 2024-08-07",
     "Steering Gear", "Grease Pump", "BIOMULTIS EP 2",
     "TotalEnergies", "FLUTEK", "FE21", "(EAL)", ""),
    ("BU FINTAS", "9976824", "TotalEnergies\nLubmarine", "R2 / 2024-08-07",
     "Rudder Carrier", "Grease Filling", "BIOMULTIS EP 2",
     "TotalEnergies", "MACGREGOR", "", "(EAL)", ""),
    ("BU FINTAS", "9976824", "TotalEnergies\nLubmarine", "R2 / 2024-08-07",
     "General Greasing", "Grease Points", "BIOMULTIS EP 2",
     "TotalEnergies", "(전선 공통)", "", "(EAL)", ""),
    ("BU FINTAS", "9976824", "TotalEnergies\nLubmarine", "R2 / 2024-08-07",
     "Emergency Towing System", "Wire Ropes", "BIOADHESIVE PLUS",
     "TotalEnergies", "TANKTECH", "", "(EAL)", ""),
    ("BU FINTAS", "9976824", "TotalEnergies\nLubmarine", "R2 / 2024-08-07",
     "General Greasing", "Wire Ropes", "BIOADHESIVE PLUS",
     "TotalEnergies", "(전선 공통)", "", "(EAL)", ""),
    ("BU FINTAS", "9976824", "TotalEnergies\nLubmarine", "R2 / 2024-08-07",
     "General Greasing", "Open Gears", "BIO OG+",
     "TotalEnergies", "(전선 공통)", "", "(EAL)", ""),
    ("BU FINTAS", "9976824", "TotalEnergies\nLubmarine", "R2 / 2024-08-07",
     "(차트 COMMENTS Note 3)", "초기 충전유", "MOBIL ARCTIC EAL 32",
     "ExxonMobil", "(차트상 미기재)", "", "제품명에 EAL 포함",
     "★ 냉동기 압축기용 POE 오일. 제품명의 EAL 은 상표이며 VGP 환경친화 윤활유 아님"),

    # ---- MARVEL DOVE (Chevron) : "(VGP Compliant)" 표기
    ("MARVEL DOVE", "9964182", "Chevron", "2024-06-17",
     "Propeller Bonnet", "Grease Points", "CLARITY SYN EA GREASE",
     "Chevron", "HHI-EMD", "", "(VGP Compliant)", ""),
    ("MARVEL DOVE", "9964182", "Chevron", "2024-06-17",
     "Accommodation Ladder", "Wire Ropes", "CLARITY SYN EA GREASE",
     "Chevron", "SAMGONG", "SH81731", "(VGP compliant)", ""),
    ("MARVEL DOVE", "9964182", "Chevron", "2024-06-17",
     "Pilot Ladder Reel", "Wire Ropes", "CLARITY SYN EA GREASE",
     "Chevron", "SAMGONG", "SH81731", "(VGP compliant)", ""),
    ("MARVEL DOVE", "9964182", "Chevron", "2024-06-17",
     "Bosun Store Davit", "Open Gears & Wire Ropes", "CLARITY SYN EA GREASE",
     "Chevron", "A-TECH", "", "(VGP compliant)", ""),
    ("MARVEL DOVE", "9964182", "Chevron", "2024-06-17",
     "Emergency Cargo Pump Davit", "Wire Ropes", "CLARITY SYN EA GREASE",
     "Chevron", "SHIN MYUNG", "", "(VGP compliant)", ""),
    ("MARVEL DOVE", "9964182", "Chevron", "2024-06-17",
     "Remedy Handling Davit", "Wire Ropes", "CLARITY SYN EA GREASE",
     "Chevron", "SHIN MYUNG", "", "(VGP compliant)", ""),
    ("MARVEL DOVE", "9964182", "Chevron", "2024-06-17",
     "Emergency Towing System", "Towing Wire", "CLARITY SYN EA GREASE",
     "Chevron", "KTMI", "", "(VGP Compliant)", ""),
    ("MARVEL DOVE", "9964182", "Chevron", "2024-06-17",
     "Life & Rescue Boat Davit", "Wire Ropes", "CLARITY SYN EA GREASE",
     "Chevron", "ORIENTAL", "HGD-062-26 / BWE-10", "(VGP Compliant)", ""),
    ("MARVEL DOVE", "9964182", "Chevron", "2024-06-17",
     "Open Gears and Wire Ropes\n(Miscellaneous)", "Brush Applied", "CLARITY SYN EA GREASE",
     "Chevron", "(전선 공통)", "", "(VGP Compliant)", ""),
    ("MARVEL DOVE", "9964182", "Chevron", "2024-06-17",
     "Rudder Carrier", "Grease Points", "CLARITY SYN EA GREASE 0",
     "Chevron", "3K INDUSTRY", "", "(VGP Compliant)", ""),
    ("MARVEL DOVE", "9964182", "Chevron", "2024-06-17",
     "Unit Cooler - ECR", "Crankcase", "MOBIL Arctic EAL 32",
     "ExxonMobil", "HI-AIR KOREA", "", "제품명에 EAL 포함",
     "★ 냉동기 압축기용 POE 오일 · Initial Fill & Owner Supplied. VGP 환경친화 윤활유 아님"),
    ("MARVEL DOVE", "9964182", "Chevron", "2024-06-17",
     "Unit Cooler - Switch Board Room", "Crankcase", "MOBIL Arctic EAL 32",
     "ExxonMobil", "HI-AIR KOREA", "", "제품명에 EAL 포함",
     "★ 상동 · Initial Fill & Owner Supplied"),
    ("MARVEL DOVE", "9964182", "Chevron", "2024-06-17",
     "Unit Cooler - Work Shop", "Crankcase", "MOBIL Arctic EAL 32",
     "ExxonMobil", "HI-AIR KOREA", "", "제품명에 EAL 포함",
     "★ 상동 · Initial Fill & Owner Supplied"),

    # ---- PRISM AGILITY (TotalEnergies) : "(EAL)" 표기
    ("PRISM AGILITY", "9810549", "TotalEnergies\nLubmarine", "R3 / 2024-01-11",
     "Bow Thruster", "Enclosed Gears", "BIONEPTAN HT 100",
     "TotalEnergies", "KAWASAKI HEAVY INDUSTRIES (KHI)", "KT-219B5", "(EAL)", ""),
    ("PRISM AGILITY", "9810549", "TotalEnergies\nLubmarine", "R3 / 2024-01-11",
     "Bow Thruster", "Grease Points", "BIOMULTIS EP 2",
     "TotalEnergies", "KAWASAKI HEAVY INDUSTRIES (KHI)", "KT-219B5", "(EAL)", ""),
    ("PRISM AGILITY", "9810549", "TotalEnergies\nLubmarine", "R3 / 2024-01-11",
     "Hi-FIN Inside", "Grease Points", "BIOMULTIS EP 2",
     "TotalEnergies", "HYUNDAI HEAVY INDUSTRIES (HHI)", "FPP", "(요약표 EAL 분류)",
     "본문 표에는 (EAL) 표기 없음. 2p 요약표에서 BIOMULTIS EP 2 항목으로 분류"),
    ("PRISM AGILITY", "9810549", "TotalEnergies\nLubmarine", "R3 / 2024-01-11",
     "Bow Thruster", "초기 충전유 (First oil)", "MOBIL SHC AWARE GEAR 100",
     "ExxonMobil", "KAWASAKI HEAVY INDUSTRIES (KHI)", "KT-219B5", "COMMENTS Note 1",
     "★ 차트 표에는 (EAL) 미표기. SHC AWARE 는 ExxonMobil 의 EAL 기어유 제품군"),
    ("PRISM AGILITY", "9810549", "TotalEnergies\nLubmarine", "R3 / 2024-01-11",
     "Rudder Carrier", "Grease Points", "HOUTON TECTYL G OS 5550 ECO",
     "Houghton (Tectyl)", "3K INDUSTRY", "MSB-360-T6M", "COMMENTS Note 3",
     "★ 차트 표에는 'Maker supply' 로만 기재. 차트 원문 표기 'HOUTON' (Houghton 오기로 추정)"),

    # ---- SK AUDACE (TotalEnergies) : "(EAL)" 표기
    ("SK AUDACE", "9693161", "TotalEnergies\nLubmarine", "R3 / 2024-01-11",
     "Stern Tube", "Bearings & Seals", "BIONEPTAN 150",
     "TotalEnergies", "WARTSILA", "", "(EAL)",
     "★ 본 6척 중 유일하게 Stern Tube 에 EAL 적용 (타 5척은 광유)"),
    ("SK AUDACE", "9693161", "TotalEnergies\nLubmarine", "R3 / 2024-01-11",
     "Bow Thruster", "Enclosed Gears", "BIONEPTAN HT 100",
     "TotalEnergies", "KHI", "", "(EAL)", ""),
    ("SK AUDACE", "9693161", "TotalEnergies\nLubmarine", "R3 / 2024-01-11",
     "Windlass & Mooring Winch", "Grease Points", "BIOMULTIS EP 2",
     "TotalEnergies", "FLUTEK", "", "(EAL)", ""),
    ("SK AUDACE", "9693161", "TotalEnergies\nLubmarine", "R3 / 2024-01-11",
     "Cargo Machinery Room Crane", "Grease Points", "BIOMULTIS EP 2",
     "TotalEnergies", "ORIENTAL", "", "(EAL)", ""),
    ("SK AUDACE", "9693161", "TotalEnergies\nLubmarine", "R3 / 2024-01-11",
     "Hose Handling Crane", "Grease Points", "BIOMULTIS EP 2",
     "TotalEnergies", "ORIENTAL", "", "(EAL)", ""),
    ("SK AUDACE", "9693161", "TotalEnergies\nLubmarine", "R3 / 2024-01-11",
     "Provision Crane", "Grease Points", "BIOMULTIS EP 2",
     "TotalEnergies", "ORIENTAL", "", "(EAL)", ""),
    ("SK AUDACE", "9693161", "TotalEnergies\nLubmarine", "R3 / 2024-01-11",
     "Emergency Towing System", "Grease Points", "BIOMULTIS EP 2",
     "TotalEnergies", "TANKTECH", "", "(EAL)", ""),
    ("SK AUDACE", "9693161", "TotalEnergies\nLubmarine", "R3 / 2024-01-11",
     "Rescue Boat Davit Winch", "Grease Points", "BIOMULTIS EP 2",
     "TotalEnergies", "ORIENTAL", "", "(EAL)", ""),
    ("SK AUDACE", "9693161", "TotalEnergies\nLubmarine", "R3 / 2024-01-11",
     "Lifeboat Davit Winch", "Grease Points", "BIOMULTIS EP 2",
     "TotalEnergies", "ORIENTAL", "", "(EAL)", ""),
    ("SK AUDACE", "9693161", "TotalEnergies\nLubmarine", "R3 / 2024-01-11",
     "General Lubrication", "Grease Points", "BIOMULTIS EP 2",
     "TotalEnergies", "(전선 공통)", "", "(EAL)", ""),
    ("SK AUDACE", "9693161", "TotalEnergies\nLubmarine", "R3 / 2024-01-11",
     "Cargo Machinery Room Crane", "Wire Ropes", "BIOADHESIVE PLUS",
     "TotalEnergies", "ORIENTAL", "", "(EAL)", ""),
    ("SK AUDACE", "9693161", "TotalEnergies\nLubmarine", "R3 / 2024-01-11",
     "Hose Handling Crane", "Wire Ropes", "BIOADHESIVE PLUS",
     "TotalEnergies", "ORIENTAL", "", "(EAL)", ""),
    ("SK AUDACE", "9693161", "TotalEnergies\nLubmarine", "R3 / 2024-01-11",
     "Provision Crane", "Wire Ropes", "BIOADHESIVE PLUS",
     "TotalEnergies", "ORIENTAL", "", "(EAL)", ""),
    ("SK AUDACE", "9693161", "TotalEnergies\nLubmarine", "R3 / 2024-01-11",
     "Rescue Boat Davit Winch", "Wire Ropes", "BIOADHESIVE PLUS",
     "TotalEnergies", "ORIENTAL", "", "(EAL)", ""),
    ("SK AUDACE", "9693161", "TotalEnergies\nLubmarine", "R3 / 2024-01-11",
     "Lifeboat Davit Winch", "Wire Ropes", "BIOADHESIVE PLUS",
     "TotalEnergies", "ORIENTAL", "", "(EAL)", ""),
    ("SK AUDACE", "9693161", "TotalEnergies\nLubmarine", "R3 / 2024-01-11",
     "General Lubrication", "Wire Ropes", "BIOADHESIVE PLUS",
     "TotalEnergies", "(전선 공통)", "", "(EAL)", ""),
    ("SK AUDACE", "9693161", "TotalEnergies\nLubmarine", "R3 / 2024-01-11",
     "Rudder Carrier", "Grease Pump", "MOBIL SHC AWARE GREASE EP 2",
     "ExxonMobil", "FLUTEK", "", "COMMENTS Note 3",
     "★ 차트 표에는 'Maker supply' 로만 기재. SHC AWARE 는 ExxonMobil 의 EAL 제품군"),
    ("SK AUDACE", "9693161", "TotalEnergies\nLubmarine", "R3 / 2024-01-11",
     "Packaged Type Unit Cooler — ECR", "Crankcase (Synthetic)", "MOBIL ARCTIC EAL 32",
     "ExxonMobil", "HI-AIR KOREA", "", "COMMENTS Note 1",
     "★ 냉동기 압축기용 POE 오일. 제품명의 EAL 은 상표이며 VGP 환경친화 윤활유 아님"),
    ("SK AUDACE", "9693161", "TotalEnergies\nLubmarine", "R3 / 2024-01-11",
     "Packaged Type Unit Cooler — MSBD Room", "Crankcase (Synthetic)", "MOBIL ARCTIC EAL 32",
     "ExxonMobil", "HI-AIR KOREA", "", "COMMENTS Note 1", "★ 상동"),
    ("SK AUDACE", "9693161", "TotalEnergies\nLubmarine", "R3 / 2024-01-11",
     "Packaged Type Unit Cooler — Workshop", "Crankcase (Synthetic)", "MOBIL ARCTIC EAL 32",
     "ExxonMobil", "HI-AIR KOREA", "", "COMMENTS Note 1", "★ 상동"),
]

# 제품별 요약 ----------------------------------------------------------
# (제품명, 오일메이커, 제품유형, 적용선박, 주요 적용부위, 표기, 비고)
PRODUCTS = [
    ("BIOMULTIS EP 2", "TotalEnergies\nLubmarine", "생분해성 다목적 EP 그리스",
     "AL SAKHAMAH\nBU FINTAS\nPRISM AGILITY\nSK AUDACE",
     "Propeller Cap · Rudder Carrier ·\nSteering Gear Grease Pump ·\nBow Thruster · Hi-FIN Inside ·\nWinch · Crane · Davit ·\nEmergency Towing · General Greasing",
     "(EAL)", "4개 선박 · 총 15개 부위. 본 6척 중 최다 적용 EAL 제품"),
    ("BIOADHESIVE PLUS", "TotalEnergies\nLubmarine", "생분해성 점착성 와이어로프 그리스",
     "AL SAKHAMAH\nBU FINTAS\nSK AUDACE",
     "Emergency Towing System\n(Towing Pennant / Wire Ropes) ·\nRudder Trunk · Crane · Davit Winch ·\nGeneral Greasing",
     "(EAL)", "3개 선박 · 총 10개 부위"),
    ("BIO OG+", "TotalEnergies\nLubmarine", "생분해성 개방기어용 그리스",
     "BU FINTAS", "General Greasing — Open Gears", "(EAL)", "1개 선박 · 1개 부위"),
    ("BIONEPTAN HT 100", "TotalEnergies\nLubmarine", "생분해성 밀폐기어유 (ISO VG 100)",
     "PRISM AGILITY\nSK AUDACE", "Bow Thruster — Enclosed Gears", "(EAL)",
     "2개 선박 · 총 2개 부위"),
    ("BIONEPTAN 150", "TotalEnergies\nLubmarine", "생분해성 선미관유 (ISO VG 150)",
     "SK AUDACE", "Stern Tube — Bearings & Seals", "(EAL)",
     "★ 본 6척 중 유일한 Stern Tube EAL 적용. 나머지 5척은 광유\n"
     "(Melina S 30 / Atlanta Marine D 3005 / Veritas 800 Marine 30)"),
    ("Shell Naturelle S2 Grease\nA600P 1.5", "Shell", "생분해성 와이어로프 그리스",
     "PUTERI SABAH",
     "Accommodation / Pilot Ladder ·\nTowing Wire · Crane Open Gears &\nWire Ropes · Lifeboat Davit ·\nGeneral Lubrication",
     "(VGP Compliant)", "구 제품명 : Naturelle S2 Wire Rope Lubricant A (차트 부록 기재)"),
    ("Shell Naturelle S5 Grease\nV120P 2", "Shell", "생분해성 다목적 그리스 (NLGI 2)",
     "PUTERI SABAH", "Propeller Bonnet / Cap Grease Filling ·\nGeneral Grease Points",
     "(VGP Compliant)", "1개 선박 · 2개 부위"),
    ("CLARITY SYN EA GREASE", "Chevron", "합성 생분해성 그리스",
     "MARVEL DOVE",
     "Propeller Bonnet · Wire Ropes\n(Accommodation / Pilot Ladder,\nDavit 4종) · Towing Wire ·\nOpen Gears (Brush Applied)",
     "(VGP Compliant)", "1개 선박 · 9개 부위"),
    ("CLARITY SYN EA GREASE 0", "Chevron", "합성 생분해성 그리스 (0번 주도)",
     "MARVEL DOVE", "Rudder Carrier — Grease Points", "(VGP Compliant)",
     "Rudder Carrier 전용으로 별도 지정"),
    ("MOBIL SHC AWARE GEAR 100", "ExxonMobil", "합성 생분해성 기어유 (ISO VG 100)",
     "PRISM AGILITY", "Bow Thruster 초기 충전유", "COMMENTS Note 1",
     "★ 차트 본문 표에는 (EAL) 미표기 — COMMENTS 주석에만 기재"),
    ("MOBIL SHC AWARE GREASE EP 2", "ExxonMobil", "합성 생분해성 EP 그리스",
     "SK AUDACE", "Rudder Carrier — Grease Pump", "COMMENTS Note 3",
     "★ 차트 본문 표에는 'Maker supply' 로만 기재 — COMMENTS 주석에만 제품명 명시"),
    ("HOUTON TECTYL G OS 5550 ECO", "Houghton (Tectyl)", "환경친화 러더캐리어 그리스",
     "PRISM AGILITY", "Rudder Carrier — Grease Points", "COMMENTS Note 3",
     "★ 차트 본문 표에는 'Maker supply' 로만 기재 — COMMENTS 주석에만 제품명 명시"),
    ("Margrease EP 0", "확인 필요", "러더캐리어 그리스",
     "PUTERI SABAH", "Rudder Carrier — Grease Filling", "(VGP Compliant)",
     "★ 차트 주석 (B) = Shell Marine 미공급 품목. 제조사 확인 필요"),
    ("MOBIL Arctic EAL 32", "ExxonMobil", "냉동기 압축기유 (POE)",
     "MARVEL DOVE\nBU FINTAS\nSK AUDACE", "Unit Cooler Crankcase\n(MARVEL DOVE 3개소 · SK AUDACE 3개소) ·\nBU FINTAS 초기 충전유",
     "제품명에 EAL 포함",
     "★ 제품명의 'EAL' 은 상표이며 VGP 환경친화 윤활유(EAL) 아님. "
     "밀폐 냉동회로용으로 해수 접촉부가 아님 — EAL 목록 집계 시 제외 검토 필요"),
]

wb = openpyxl.Workbook()

# =====================================================================
# SHEET 1 - EAL 제품·메이커
# =====================================================================
ws = wb.active
ws.title = "EAL 제품·메이커"

ws.merge_cells("A1:G1")
ws["A1"] = "선박 6척 Lubrication Chart — EAL(환경친화 윤활유) 제품 · 메이커 종합"
ws["A1"].font = F(15, True, "FFFFFF"); ws["A1"].fill = NAVY; ws["A1"].alignment = CTR
ws.row_dimensions[1].height = 32

ws.merge_cells("A2:G2")
ws["A2"] = ("※ 대상 : PUTERI SABAH · AL SAKHAMAH · BU FINTAS · MARVEL DOVE · PRISM AGILITY · SK AUDACE (SK SHIPPING)\n"
            "※ TotalEnergies 차트는 「(EAL)」, Shell · Chevron 차트는 「(VGP Compliant)」로 표기 — 표기만 다를 뿐 동일한 환경친화 윤활유 지정입니다.")
ws["A2"].font = F(9, color="404040"); ws["A2"].alignment = WRAP
ws.row_dimensions[2].height = 30

hdr = ["No.", "EAL 제품명", "오일 메이커\n(브랜드)", "제품 유형", "적용 선박", "주요 적용 부위", "비고"]
for i, h in enumerate(hdr, start=1):
    c = ws.cell(4, i, h)
    c.font = F(10, True, "FFFFFF"); c.fill = HEADF; c.alignment = CTR; c.border = BOX
ws.row_dimensions[4].height = 30

r = 5
for n, (name, maker, kind, vessels, parts, mark, note) in enumerate(PRODUCTS, start=1):
    vals = [n, name, maker, kind, vessels, parts, note]
    for i, v in enumerate(vals, start=1):
        c = ws.cell(r, i, v)
        c.border = BOX
        c.alignment = CTR if i in (1, 3) else WRAP
        c.font = F(10, True) if i == 2 else F(9)
    if note.startswith("★"):
        ws.cell(r, 7).font = F(9, True, "C00000")
        ws.cell(r, 7).fill = WARN
    if maker == "확인 필요":
        ws.cell(r, 3).fill = YEL; ws.cell(r, 3).font = F(10, True, "C00000")
    elif r % 2 == 1:
        for i in range(1, 7):
            ws.cell(r, i).fill = ALT
    ws.row_dimensions[r].height = 62
    r += 1
last = r - 1

r += 1
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
c = ws.cell(r, 1,
    "▣ 확인 사항\n"
    "  · 「EAL」 문자열 검색만으로는 PUTERI SABAH(Shell) · MARVEL DOVE(Chevron) 두 척이 누락됩니다. "
    "두 차트는 EAL 을 「(VGP Compliant)」로 표기하며, 해당 제품은 Shell Naturelle 계열과 Chevron Clarity Syn EA 계열입니다.\n"
    "  · MOBIL Arctic EAL 32 는 제품명에 EAL 이 들어가지만 냉동기 압축기용 POE 오일입니다. "
    "해수 접촉부(oil-to-sea interface)가 아니므로 VGP 상 EAL 과 성격이 다릅니다 — 집계 목적에 따라 제외를 검토하십시오.\n"
    "  · MOBIL SHC AWARE GEAR 100 과 HOUTON TECTYL G OS 5550 ECO 는 PRISM AGILITY 차트 본문 표가 아닌 COMMENTS 주석에만 있어 "
    "표 검색으로는 잡히지 않습니다.\n"
    "  · Margrease EP 0 (PUTERI SABAH Rudder Carrier) 는 VGP Compliant 로 지정되었으나 Shell 미공급 품목이라 제조사가 차트에 없습니다.\n"
    "  · 노란색 셀 = 원문에서 확인되지 않아 별도 확인이 필요한 항목입니다.")
c.font = F(9, color="404040"); c.alignment = WRAP
c.fill = PatternFill("solid", fgColor="FFF9E6"); c.border = BOX
ws.row_dimensions[r].height = 118

for col, w in zip("ABCDEFG", [5, 26, 18, 26, 18, 34, 40]):
    ws.column_dimensions[col].width = w
ws.freeze_panes = "A5"

# =====================================================================
# SHEET 2 - 선박별 상세
# =====================================================================
ws2 = wb.create_sheet("선박별 상세")
ws2.merge_cells("A1:K1")
ws2["A1"] = "선박별 EAL 적용 상세 (Lubrication Chart 원문 기준)"
ws2["A1"].font = F(14, True, "FFFFFF"); ws2["A1"].fill = NAVY; ws2["A1"].alignment = CTR
ws2.row_dimensions[1].height = 30
ws2.merge_cells("A2:K2")
ws2["A2"] = "※ 각 행은 차트 원문 표기를 그대로 옮긴 것입니다. 장비 메이커는 차트의 EQUIPMENT MAKER 열 기준입니다."
ws2["A2"].font = F(9, color="404040"); ws2["A2"].alignment = WRAP

h2 = ["선박명", "IMO", "차트 공급사", "개정", "장비 (Equipment)", "적용 부위",
      "EAL 제품명", "오일 메이커", "장비 메이커", "모델", "차트 표기 / 비고"]
for i, h in enumerate(h2, start=1):
    c = ws2.cell(4, i, h)
    c.font = F(10, True, "FFFFFF"); c.fill = HEADF; c.alignment = CTR; c.border = BOX
ws2.row_dimensions[4].height = 30

r = 5
vessel_spans = []
cur, start = None, 5
for row in DETAIL:
    ves, imo, sup, rev, eq, part, prod, omk, emk, model, mark, note = row
    if cur is None:
        cur = ves
    elif ves != cur:
        vessel_spans.append((start, r - 1))
        cur, start = ves, r
    remark = mark if not note else f"{mark}\n{note}"
    for i, v in enumerate([ves, imo, sup, rev, eq, part, prod, omk, emk, model, remark], start=1):
        c = ws2.cell(r, i, v)
        c.border = BOX
        c.alignment = CTR if i in (1, 2, 3, 4, 10) else WRAP
        c.font = F(9, True) if i == 7 else F(9)
    if note.startswith("★"):
        ws2.cell(r, 11).font = F(9, True, "C00000")
        for i in range(1, 12):
            ws2.cell(r, i).fill = WARN
    if omk == "확인 필요":
        ws2.cell(r, 8).fill = YEL; ws2.cell(r, 8).font = F(9, True, "C00000")
    ws2.row_dimensions[r].height = 40
    r += 1
vessel_spans.append((start, r - 1))

for s, e in vessel_spans:
    if e > s:
        for col in (1, 2, 3, 4):
            ws2.merge_cells(start_row=s, start_column=col, end_row=e, end_column=col)
    for col in (1, 2, 3, 4):
        ws2.cell(s, col).alignment = CTR
        ws2.cell(s, col).font = F(10, True, "1F3864") if col == 1 else F(9)
        ws2.cell(s, col).fill = GRPF

detail_last = r - 1
r += 1
ws2.merge_cells(start_row=r, start_column=1, end_row=r, end_column=11)
c = ws2.cell(r, 1,
    "▣ 출처\n"
    "  · PUTERI SABAH — Shell Lubrication Survey, Version 2, 2025-01-15 (Vessel Code 806179)\n"
    "  · AL SAKHAMAH — TotalEnergies Lubmarine Lubrication Chart R3, 2025-04-16 (NAVITEC No. 32788)\n"
    "  · BU FINTAS — TotalEnergies Lubmarine Lubrication Chart R2, 2024-08-07 (NAVITEC No. 32574)\n"
    "  · MARVEL DOVE — Chevron Marine Lubricants Lubrication Chart, 최종 개정 2024-06-17 (Provisional)\n"
    "  · PRISM AGILITY — TotalEnergies Lubmarine Lubrication Chart R3, 2024-01-11 (NAVITEC No. 24221)\n  · SK AUDACE — TotalEnergies Lubmarine Lubrication Chart R3, 2024-01-11 (NAVITEC No. 22000)\n"
    "  · 분홍색 행 = 표기 방식이 달라 별도 판단이 필요한 항목 (제품명에 EAL 포함 / COMMENTS 주석 기재 / 요약표 분류)")
c.font = F(9, color="404040"); c.alignment = WRAP
c.fill = PatternFill("solid", fgColor="FFF9E6"); c.border = BOX
ws2.row_dimensions[r].height = 100

for col, w in zip("ABCDEFGHIJK", [15, 11, 15, 14, 26, 22, 27, 15, 28, 20, 40]):
    ws2.column_dimensions[col].width = w
ws2.freeze_panes = "E5"

# =====================================================================
# SHEET 0 - 한눈에 보기 (부위별 요약 + Klüber / SKF 대조)
# =====================================================================
OK   = "1F7A3D"   # 승인 확인
NG   = "C00000"   # 확인 필요
GREY = "808080"   # 해당 없음

# (부위, 현재제품, 오일메이커, 적용선박, 장비메이커, Klüber대응, Klüber승인, SKF등재, 상태)
#   상태: 'ok' | 'chk' | 'na' | 'warn'
OVERVIEW = [
    ("Stern Tube\n(선미관)", "BIONEPTAN 150", "TotalEnergies", "SK AUDACE", "WÄRTSILÄ",
     "Klüberbio RM 2-150", "승인 — Wärtsilä Japan · Sweden(Cedervall) · UK\n(BIO FKM seal rings)",
     "등재 — Bioneptan 150 / 150 cSt /\nSterntube Oil / FKM Pod · FKM Bio", "ok"),
    ("Stern Tube\n(선미관)", "ATLANTA MARINE D 3005\nMelina S 30 · Veritas 800 Marine 30",
     "TotalEnergies / Shell /\nChevron", "나머지 5척", "KEMEL · WÄRTSILÄ",
     "Klüberbio RM 2-100 / RM 2-150",
     "승인 — KEMEL(선박별 개별승인 필요) ·\nWärtsilä · SKF Marine(Simplex)",
     "해당 없음 (광유는 EAL 아님)",
     "warn"),

    ("Bow Thruster\n— 기어유", "BIONEPTAN HT 100", "TotalEnergies", "PRISM AGILITY\nSK AUDACE", "KHI",
     "Klüberbio EG 2-100", "승인 — Kawasaki Heavy Industries\n(2018-02 및 2018-08 자료 모두 확인)",
     "등재 — Bioneptan HT 100 / 100 cSt /\nMultipurpose Oil", "ok"),
    ("Bow Thruster\n— 기어유", "MOBIL SHC AWARE GEAR 100", "ExxonMobil", "PRISM AGILITY\n(초기 충전유)", "KHI",
     "Klüberbio EG 2-100", "승인 — Kawasaki Heavy Industries",
     "등재 — Mobil SHC Aware Gear 100 /\n100 cSt / Gear Oil", "ok"),
    ("Bow Thruster\n— 그리스", "BIOMULTIS EP 2", "TotalEnergies", "PRISM AGILITY\nSK AUDACE", "KHI",
     "Klüberbio AG 39-602 계열", "자료 내 KHI 그리스 승인 없음",
     "대상 아님 (SKF 리스트는 오일 전용)", "chk"),

    ("Propeller Cap /\nBonnet", "BIOMULTIS EP 2", "TotalEnergies", "AL SAKHAMAH\nBU FINTAS", "SILLA METAL CO.",
     "Klüberbio BM 32-142 /\nAG 39-602 / LG 39-701 N",
     "자료 내 SILLA METAL 없음\n(승인: MAN D&T · Mecklenburger Metallguss ·\nSistemar · Stone Marine)",
     "대상 아님 (그리스)", "chk"),
    ("Propeller Cap /\nBonnet", "Shell Naturelle S5 Grease\nV120P 2", "Shell", "PUTERI SABAH", "HHI (EMD)",
     "Klüberbio BM 32-142 /\nAG 39-602 / LG 39-701 N", "자료 내 HHI (EMD) 없음",
     "대상 아님 (그리스)", "chk"),
    ("Propeller Cap /\nBonnet", "CLARITY SYN EA GREASE", "Chevron", "MARVEL DOVE", "HHI-EMD",
     "Klüberbio BM 32-142 /\nAG 39-602 / LG 39-701 N", "자료 내 HHI-EMD 없음",
     "대상 아님 (그리스)", "chk"),

    ("Rudder Carrier", "CLARITY SYN EA GREASE 0", "Chevron", "MARVEL DOVE", "3K INDUSTRY",
     "Klüberbio LG 39-701 N", "승인 — 3K / Korea\n(rudder shaft and pump)",
     "대상 아님 (그리스)", "ok"),
    ("Rudder Carrier", "HOUTON TECTYL G OS 5550 ECO", "Houghton (Tectyl)", "PRISM AGILITY", "3K INDUSTRY",
     "Klüberbio LG 39-701 N", "승인 — 3K / Korea\n(rudder shaft and pump)",
     "대상 아님 (그리스)", "ok"),
    ("Rudder Carrier", "Margrease EP 0", "확인 필요", "PUTERI SABAH", "3K Industry",
     "Klüberbio LG 39-701 N", "승인 — 3K / Korea\n(rudder shaft and pump)",
     "대상 아님 (그리스)", "ok"),
    ("Rudder Carrier", "MOBIL SHC AWARE GREASE EP 2", "ExxonMobil", "SK AUDACE", "FLUTEK",
     "Klüberbio LG 39-701 N /\nAG 39-602", "자료 내 FLUTEK 없음", "대상 아님 (그리스)", "chk"),
    ("Rudder Carrier", "BIOMULTIS EP 2", "TotalEnergies", "BU FINTAS", "MACGREGOR",
     "Klüberbio LG 39-701 N /\nAG 39-602", "자료 내 MACGREGOR 없음", "대상 아님 (그리스)", "chk"),
    ("Rudder Carrier\n(Rudder Trunk)", "BIOADHESIVE PLUS", "TotalEnergies", "AL SAKHAMAH", "FLUTEK",
     "Klüberbio LG 39-701 N /\nAG 39-602", "자료 내 FLUTEK 없음", "대상 아님 (그리스)", "chk"),

    ("Wire Ropes /\nTowing Wire", "BIOADHESIVE PLUS", "TotalEnergies", "AL SAKHAMAH\nBU FINTAS\nSK AUDACE",
     "ORIENTAL · TANKTECH 등", "Klüber 자료 범위 밖", "— (Klüber 자료는 선미관 · 스러스터 ·\n씰/베어링 그리스 한정)",
     "대상 아님 (그리스)", "na"),
    ("Wire Ropes /\nTowing Wire", "Shell Naturelle S2 Grease\nA600P 1.5", "Shell", "PUTERI SABAH",
     "Samgong · Tanktec ·\nSangsangin · OPCO", "Klüber 자료 범위 밖", "—", "대상 아님 (그리스)", "na"),
    ("Wire Ropes /\nTowing Wire", "CLARITY SYN EA GREASE", "Chevron", "MARVEL DOVE",
     "SAMGONG · A-TECH ·\nSHIN MYUNG · KTMI · ORIENTAL", "Klüber 자료 범위 밖", "—", "대상 아님 (그리스)", "na"),

    ("Open Gears", "BIO OG+", "TotalEnergies", "BU FINTAS", "(전선 공통)",
     "Klüber 자료 범위 밖", "—", "대상 아님 (그리스)", "na"),
    ("Grease Points\n(일반)", "BIOMULTIS EP 2", "TotalEnergies", "BU FINTAS · SK AUDACE\nPRISM AGILITY",
     "Winch · Crane · Davit 등", "Klüberbio AG 39-602 계열", "—", "대상 아님 (그리스)", "na"),

    ("참고 — EAL 아님", "MOBIL Arctic EAL 32", "ExxonMobil", "MARVEL DOVE\nBU FINTAS\nSK AUDACE",
     "HI-AIR KOREA", "해당 없음", "해당 없음",
     "해당 없음", "warn"),
]

ws0 = wb.create_sheet("한눈에 보기", 0)
ws0.merge_cells("A1:I1")
ws0["A1"] = "EAL 적용 한눈에 보기 — 부위별 요약 및 Klüber · SKF 승인 대조"
ws0["A1"].font = F(15, True, "FFFFFF"); ws0["A1"].fill = NAVY; ws0["A1"].alignment = CTR
ws0.row_dimensions[1].height = 32
ws0.merge_cells("A2:I2")
ws0["A2"] = ("※ 선박 6척 · 적용 60건을 「적용 부위」 기준으로 압축한 표입니다. 건별 원문은 '선박별 상세' 시트를 보십시오.\n"
             "※ Klüber 승인 = 해당 장비 메이커가 Klüber EAL 제품을 승인했는지 · SKF 등재 = 현재 사용 제품이 SKF Marine EAL 리스트에 있는지")
ws0["A2"].font = F(9, color="404040"); ws0["A2"].alignment = WRAP
ws0.row_dimensions[2].height = 32

h0 = ["적용 부위", "현재 사용 제품", "오일 메이커", "적용 선박", "장비 메이커",
      "Klüber 대응 제품", "Klüber 승인 (장비 메이커 기준)", "SKF Marine EAL 리스트 등재", "판정"]
for i, h in enumerate(h0, start=1):
    c = ws0.cell(4, i, h)
    c.font = F(10, True, "FFFFFF"); c.fill = HEADF; c.alignment = CTR; c.border = BOX
ws0.row_dimensions[4].height = 34

VERDICT = {"ok": ("승인 확인", OK), "chk": ("확인 필요", NG),
           "na": ("해당 없음", GREY), "warn": ("검토 필요", "BF8F00")}

r = 5
part_spans = []
cur, start = None, 5
for part, prod, omk, ves, emk, klu, appr, skf, state in OVERVIEW:
    if cur is None:
        cur = part
    elif part != cur:
        part_spans.append((start, r - 1)); cur, start = part, r
    label, color = VERDICT[state]
    for i, v in enumerate([part, prod, omk, ves, emk, klu, appr, skf, label], start=1):
        c = ws0.cell(r, i, v)
        c.border = BOX
        c.alignment = CTR if i in (1, 3, 9) else WRAP
        c.font = F(10, True) if i == 2 else F(9)
    ws0.cell(r, 9).font = F(9, True, color)
    if state == "ok":
        ws0.cell(r, 7).font = F(9, True, OK)
    elif state == "chk":
        ws0.cell(r, 7).font = F(9, True, NG); ws0.cell(r, 7).fill = YEL
    elif state == "warn":
        for i in range(1, 10):
            ws0.cell(r, i).fill = WARN
    if omk == "확인 필요":
        ws0.cell(r, 3).fill = YEL; ws0.cell(r, 3).font = F(9, True, NG)
    ws0.row_dimensions[r].height = 46
    r += 1
part_spans.append((start, r - 1))

for s, e in part_spans:
    if e > s:
        ws0.merge_cells(start_row=s, start_column=1, end_row=e, end_column=1)
    ws0.cell(s, 1).alignment = CTR
    ws0.cell(s, 1).font = F(10, True, "1F3864")
    ws0.cell(s, 1).fill = GRPF

r += 1
ws0.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)
c = ws0.cell(r, 1,
    "▣ 읽는 법\n"
    "  · 승인 확인 = 현재 쓰는 제품과 같은 용도로 해당 장비 메이커가 Klüber EAL 을 승인한 이력이 있음 → 대체 검토 가능\n"
    "  · 확인 필요 = 장비 메이커가 Klüber 승인 목록에 없음 → Klüber 로 바꾸려면 해당 메이커 승인을 별도로 받아야 함\n"
    "  · 해당 없음 = 와이어로프 · 개방기어용 그리스로, Klüber 자료(선미관 · 스러스터 · 씰/베어링) 범위 밖\n"
    "  · SKF Marine EAL 리스트는 Simplex 선미관 씰용 「오일」 전용 목록입니다. 그리스가 미등재인 것은 결격이 아니라 대상이 아닌 것입니다.\n\n"
    "▣ 눈에 띄는 두 가지\n"
    "  · Rudder Carrier — 3K INDUSTRY 장비를 쓰는 3척(PUTERI SABAH · MARVEL DOVE · PRISM AGILITY)은 "
    "Klüber 가 3K 로부터 rudder shaft & pump 승인을 이미 받아둔 상태입니다. 3척이 지금 서로 다른 제품(Margrease · Clarity Syn EA 0 · Tectyl ECO)을 "
    "쓰고 있어 Klüberbio LG 39-701 N 하나로 통일할 여지가 있습니다.\n"
    "  · Stern Tube — 6척 중 SK AUDACE 만 EAL(BIONEPTAN 150)입니다. 나머지 5척은 광유이며, 씰 메이커가 KEMEL · WÄRTSILÄ 라 "
    "Klüberbio RM 2-100 / RM 2-150 승인 범위 안에 들어옵니다. 다만 KEMEL 은 선박별 개별 승인이 필요합니다.")
c.font = F(9, color="404040"); c.alignment = WRAP
c.fill = PatternFill("solid", fgColor="FFF9E6"); c.border = BOX
ws0.row_dimensions[r].height = 168

for col, w in zip("ABCDEFGHI", [16, 27, 15, 17, 24, 22, 34, 30, 11]):
    ws0.column_dimensions[col].width = w
ws0.freeze_panes = "B5"

# =====================================================================
# SHEET 3 - Klüber · SKF 참조 (원문 발췌)
# =====================================================================
# (장비/씰 메이커, 용도, Klüber 제품, 승인 내용, 출처)
KLUBER = [
    ("3K / Korea", "Rudder shaft & pump", "Klüberbio LG 39-701 N", "Approved", "승인현황 p.6", True),
    ("Hanil Lubtec / Korea", "Rudder shaft & pump", "Klüberbio LG 39-701 N", "Approved", "승인현황 p.6", False),
    ("Kawasaki Heavy Industries / Japan", "Thruster 기어유", "Klüberbio EG 2-100", "Approved",
     "승인현황 p.3 · 2018-08 업데이트 p.6", True),
    ("Kawasaki Heavy Industries / Japan", "CP propeller hub", "Klüberbio EG 2-68", "Approved", "승인현황 p.5", False),
    ("Hyundai Heavy Industries / Korea", "Thruster 기어유", "Klüberbio EG 2-100", "Approved",
     "승인현황 p.3 · 2018-08 업데이트 p.6", True),
    ("KTE Nakashima / Korea", "Thruster 기어유", "Klüberbio EG 2-100", "Approved", "승인현황 p.3", False),
    ("HwaSeung R&A / Korea", "Thruster 프로펠러축 씰", "Klüberbio EG 2-68 / 100 / 150", "Approved", "승인현황 p.2", False),
    ("WÄRTSILÄ Japan (former JMT)", "Stern tube 씰", "Klüberbio RM 2-100 / RM 2-150",
     "BIO FKM seal rings 승인", "승인현황 p.1", True),
    ("WÄRTSILÄ Sweden AB (Cedervall)", "Stern tube 씰", "Klüberbio RM 2-100 / RM 2-150",
     "RM 2-150 승인 · RM 2-100 은 신형 face seal 용", "승인현황 p.1", True),
    ("WÄRTSILÄ UK (Deep Sea Seals)", "Stern tube 씰", "Klüberbio RM 2-100 / RM 2-150",
     "여러 씰 타입 승인 (최신 목록 별도 확인)", "승인현황 p.1", True),
    ("WÄRTSILÄ Japan", "Rudder shaft 씰", "Klüberbio AG 39-602", "Approved", "승인현황 p.6", True),
    ("WÄRTSILÄ Propulsion / Holland", "Thruster 기어유", "Klüberbio EG 2-150",
     "승인 (사전 협의 필요)", "승인현황 p.4 · 업데이트 p.9", True),
    ("KEMEL / Japan", "Stern tube 씰", "Klüberbio RM 2-100 / RM 2-150",
     "승인 — 단, 선박별 개별 승인 필요", "승인현황 p.1", True),
    ("KEMEL / Japan", "Thruster 씰", "Klüberbio EG 2-100 / EG 2-150", "신형 BIO seal ring 승인", "승인현황 p.2", True),
    ("SKF Marine (Simplex, 구 Blohm+Voss)", "Stern tube 씰", "Klüberbio RM 2-100 / RM 2-150",
     "FKM (Viton Pod · Viton Bio) 씰링 재질 승인", "승인현황 p.1", True),
    ("SKF Marine (Simplex)", "Thruster 씰", "Klüberbio EG 2-68 / 100 / 150",
     "FKM (Viton Pod · Viton BIO) 승인", "승인현황 p.2 · 업데이트 p.12~14", True),
    ("MAN Diesel & Turbo SE / Denmark", "Propeller cap", "Klüberbio BM 32-142 / AG 39-602 /\nLG 39-701 N",
     "Approved", "승인현황 p.6", True),
    ("Mecklenburger Metallguss / Germany", "Propeller cap", "Klüberbio BM 32-142 / AG 39-602 /\nLG 39-701 N",
     "Approved", "승인현황 p.6", True),
    ("Sistemar Propeller / Spain", "Propeller cap", "Klüberbio AG 39-602", "Approved", "승인현황 p.6", False),
    ("Stone Marine / UK", "Propeller cap", "Klüberbio AG 39-602", "Approved", "승인현황 p.6", False),
    ("Becker Marine / Germany", "Rudder shaft", "Klüberbio AG 39-602", "Approved", "승인현황 p.6", False),
]

# (브랜드, 제품명, 점도, 분류, 씰링재질, 6척 사용 여부)
SKF_LIST = [
    ("Total Lubmarine", "Bioneptan 150", "150", "Sterntube Oil", "FKM Pod, FKM Bio", "SK AUDACE 사용 중"),
    ("Total Lubmarine", "Bioneptan HT 100", "100", "Multipurpose Oil", "FKM Pod, FKM Bio",
     "PRISM AGILITY · SK AUDACE 사용 중"),
    ("Total Lubmarine", "Bioneptan 100", "100", "Sterntube Oil", "FKM Pod, FKM Bio", "—"),
    ("Total Lubmarine", "Carter Bio 150", "150", "Sterntube Oil", "FKM Pod, FKM Bio", "—"),
    ("Total Lubmarine", "Biohydran TMP 100", "100", "Hydraulic Oil", "FKM Pod, FKM Bio", "—"),
    ("ExxonMobil", "Mobil SHC Aware Gear 100", "100", "Gear Oil", "FKM Pod, FKM Bio", "PRISM AGILITY 사용 중"),
    ("ExxonMobil", "Mobil SHC Aware Gear 68", "68", "Gear Oil", "FKM Pod, FKM Bio", "—"),
    ("ExxonMobil", "Mobil SHC Aware Gear 150", "150", "Gear Oil", "FKM Pod, FKM Bio", "—"),
    ("ExxonMobil", "Mobil SHC Aware ST 100", "100", "Sterntube Oil", "FKM Pod, FKM Bio",
     "— (SKF 협의 필요 품목)"),
    ("Shell", "Naturelle S4 Stern Tube Fluid 100", "100", "Sterntube Oil", "FKM Pod, FKM Bio", "—"),
    ("Shell", "Naturelle S4 Gear Fluid 68 / 100 / 150", "68 / 100 / 150", "Gear Oil", "FKM Pod, FKM Bio", "—"),
    ("Chevron", "Clarity Synthetic EA Gear Oil 100 / 150", "100 / 150", "Gear Oil", "FKM Pod, FKM Bio", "—"),
    ("Chevron", "Clarity Synthetic EA Hydraulic Oil 68 / 100", "68 / 100", "Hydraulic Oil", "FKM Pod, FKM Bio", "—"),
    ("Klüber", "Klüberbio RM 2-100", "100", "Sterntube Oil", "FKM Pod, FKM Bio", "대응 후보"),
    ("Klüber", "Klüberbio RM 2-150", "150", "Sterntube Oil", "FKM Pod, FKM Bio", "대응 후보"),
    ("Klüber", "Klüberbio RM 8-100", "100", "Sterntube Oil", "FKM Pod, FKM Bio", "—"),
    ("Klüber", "Klüberbio EG 2-68 / 100 / 150", "68 / 100 / 150", "Gear Oil", "FKM Pod, FKM Bio", "대응 후보"),
    ("Klüber", "Klüberbio LR 9-46 / 68", "46 / 68", "Hydraulic Oil", "FKM Pod, FKM Bio", "—"),
]

ws3 = wb.create_sheet("Klüber·SKF 참조")
ws3.merge_cells("A1:F1")
ws3["A1"] = "Klüber EAL 승인 현황 · SKF Marine EAL 리스트 (원문 발췌)"
ws3["A1"].font = F(14, True, "FFFFFF"); ws3["A1"].fill = NAVY; ws3["A1"].alignment = CTR
ws3.row_dimensions[1].height = 30

ws3.merge_cells("A3:F3")
c = ws3.cell(3, 1, "[1] Klüber EAL 승인 현황 — 6척 장비 · 씰 메이커 관련 발췌")
c.font = F(11, True, "FFFFFF"); c.fill = HEADF
c.alignment = Alignment(horizontal="left", vertical="center")
ws3.row_dimensions[3].height = 24

for i, h in enumerate(["장비 · 씰 메이커", "용도", "Klüber 제품", "승인 내용", "출처", "6척 해당"], start=1):
    c = ws3.cell(4, i, h)
    c.font = F(9, True); c.fill = GRPF; c.alignment = CTR; c.border = BOX
ws3.row_dimensions[4].height = 22

r = 5
for mk, use, prod, appr, src, rel in KLUBER:
    for i, v in enumerate([mk, use, prod, appr, src, "●" if rel else ""], start=1):
        c = ws3.cell(r, i, v)
        c.border = BOX; c.font = F(9)
        c.alignment = CTR if i in (5, 6) else WRAP
    if rel:
        ws3.cell(r, 6).font = F(12, True, OK)
        for i in range(1, 7):
            ws3.cell(r, i).fill = PatternFill("solid", fgColor="EAF4EC")
    ws3.row_dimensions[r].height = 30
    r += 1

r += 2
ws3.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
c = ws3.cell(r, 1, "[2] SKF Marine EAL 리스트 — 6척 관련 브랜드 및 Klüber 대응 후보 발췌")
c.font = F(11, True, "FFFFFF"); c.fill = HEADF
c.alignment = Alignment(horizontal="left", vertical="center")
ws3.row_dimensions[r].height = 24
r += 1

for i, h in enumerate(["브랜드", "제품명", "점도 (cSt @40℃)", "분류", "씰링 재질", "6척 사용 여부"], start=1):
    c = ws3.cell(r, i, h)
    c.font = F(9, True); c.fill = GRPF; c.alignment = CTR; c.border = BOX
ws3.row_dimensions[r].height = 22
r += 1

for brand, prod, visc, kind, seal, used in SKF_LIST:
    for i, v in enumerate([brand, prod, visc, kind, seal, used], start=1):
        c = ws3.cell(r, i, v)
        c.border = BOX; c.font = F(9)
        c.alignment = CTR if i in (1, 3, 4, 5) else WRAP
    if "사용 중" in used:
        for i in range(1, 7):
            ws3.cell(r, i).fill = PatternFill("solid", fgColor="EAF4EC")
        ws3.cell(r, 6).font = F(9, True, OK)
    elif used == "대응 후보":
        ws3.cell(r, 6).font = F(9, True, "1F3864")
    ws3.row_dimensions[r].height = 22
    r += 1

r += 1
ws3.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
c = ws3.cell(r, 1,
    "▣ 출처 및 단서\n"
    "  · Klüber EAL Approval Status, 2018-02-13 (GATE-MOG / Dirk Fabry, 6p)\n"
    "  · Update Klüberbio EG 2 oils APPROVALS, 2018-08-28 v4.1 — Wilhelmsen Ships Service Singapore 교육자료\n"
    "  · SKF Marine 「Find the right EAL」 리스트 (출력일 2026-08-14, SKF Marine GmbH 5p)\n"
    "  · Klüber 자료는 2018년 기준입니다. 실제 제안 전 최신 승인 목록을 Klüber 및 해당 장비 메이커에 재확인하십시오.\n"
    "  · SKF 리스트는 Simplex 선미관 씰 시스템용 「오일」 전용이며, 중간축 베어링 · 스러스트 베어링 · 기어 · 클러치에는 적용되지 않습니다.\n"
    "  · SKF 원문 주의 : 「모든 bio oil 이 EAL 인 것은 아니다」 — EAL 여부는 오일 제조사의 책임이며 SKF 는 제조사 정보에 의존합니다.\n"
    "  · EAL 사용 시 Simplex 립링 씰은 FKM Pod 또는 FKM Bio 재질이어야 합니다. 기존 씰 재질이 다르면 씰 교체가 필요합니다.")
c.font = F(9, color="404040"); c.alignment = WRAP
c.fill = PatternFill("solid", fgColor="FFF9E6"); c.border = BOX
ws3.row_dimensions[r].height = 130

for col, w in zip("ABCDEF", [34, 24, 30, 40, 24, 26]):
    ws3.column_dimensions[col].width = w

out = "/home/user/github-test/EAL_제품_메이커_정리_6척.xlsx"
wb.save(out)
print("saved", out, "| sheets:", wb.sheetnames)
print("overview:", len(OVERVIEW), "| products:", len(PRODUCTS),
      "| detail:", len(DETAIL), "| kluber:", len(KLUBER), "| skf:", len(SKF_LIST))
