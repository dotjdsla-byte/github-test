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

out = "/home/user/github-test/EAL_제품_메이커_정리_6척.xlsx"
wb.save(out)
print("saved", out, "| products:", len(PRODUCTS), "| detail rows:", len(DETAIL))
