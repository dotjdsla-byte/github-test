const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";               // 13.333 x 7.5
pres.author = "Fuel Quality Training";
pres.title = "선박 연료 품질 동향 - Section 01";

/* ---------- design tokens (from reference deck) ---------- */
const BG      = "F4F6FB";
const NAVY    = "1B3A5C";
const NAVY_D  = "17334F";
const ORANGE  = "E17B34";
const BLUE    = "2E5C8A";
const WHITE   = "FFFFFF";
const BODY    = "3D4A5C";
const MUTED   = "8A94A6";
const LINE    = "DFE5F0";
const RED     = "C0392B";
const GREEN   = "2E7D5B";

const SERIF = "바탕";          // Korean myeongjo — on every Korean Windows
const SANS  = "맑은 고딕";      // Malgun Gothic

const M    = 0.62;             // side margin
const CW   = 13.333 - M * 2;   // content width = 12.093
const FOOT = "Classified as Internal - This content is proprietary information intended for employees and partners only.";

const sh = () => ({ type: "outer", color: "9BAAC4", blur: 10, offset: 2, angle: 90, opacity: 0.22 });

/* ---------- helpers ---------- */
function base(s, num, eyebrow, title, opts = {}) {
  s.background = { color: opts.bg || BG };
  const dark = !!opts.dark;
  s.addText(eyebrow, {
    x: M, y: 0.44, w: 8, h: 0.28, margin: 0,
    fontFace: SANS, fontSize: 11.5, bold: true, charSpacing: 2.4,
    color: ORANGE, align: "left", valign: "middle",
  });
  s.addText(title, {
    x: M, y: 0.76, w: CW, h: 0.78, margin: 0,
    fontFace: SERIF, fontSize: opts.titleSize || 36, bold: false,
    color: dark ? WHITE : NAVY, align: "left", valign: "middle",
  });
  foot(s, num, dark);
}

function foot(s, num, dark) {
  s.addText(FOOT, {
    x: M, y: 7.02, w: 9.6, h: 0.3, margin: 0,
    fontFace: SANS, fontSize: 8.5, color: dark ? "6C7C93" : MUTED, valign: "middle",
  });
  s.addText(num, {
    x: 12.2, y: 7.02, w: 0.55, h: 0.3, margin: 0,
    fontFace: SANS, fontSize: 9.5, color: dark ? "6C7C93" : MUTED, align: "right", valign: "middle",
  });
}

function card(s, x, y, w, h, fill) {
  s.addShape(pres.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.14,
    fill: { color: fill || WHITE }, line: { color: fill ? fill : "EDF1F8", width: 0.75 },
    shadow: sh(),
  });
}

function banner(s, text, y) {
  const yy = y === undefined ? 5.94 : y;
  s.addShape(pres.ShapeType.roundRect, {
    x: M, y: yy, w: CW, h: 0.86, rectRadius: 0.1,
    fill: { color: NAVY }, line: { color: NAVY, width: 0 },
  });
  s.addText(
    [
      { text: "“ ", options: { fontFace: SERIF, fontSize: 17, color: ORANGE, bold: true } },
      { text, options: { fontFace: SANS, fontSize: 14, color: WHITE, bold: true } },
      { text: " ”", options: { fontFace: SERIF, fontSize: 17, color: ORANGE, bold: true } },
    ],
    { x: M + 0.25, y: yy, w: CW - 0.5, h: 0.86, margin: 0, align: "center", valign: "middle" }
  );
}

// small caps English label used on every card in the reference deck
function label(s, txt, x, y, w, color) {
  s.addText(txt, {
    x, y, w, h: 0.24, margin: 0,
    fontFace: SANS, fontSize: 9.5, bold: true, charSpacing: 1.8,
    color: color || ORANGE, valign: "middle",
  });
}

function src(s, txt, y) {
  s.addText(txt, {
    x: M, y: y === undefined ? 6.72 : y, w: CW, h: 0.24, margin: 0,
    fontFace: SANS, fontSize: 8.5, italic: true, color: MUTED, valign: "middle",
  });
}

/* =======================================================================
   S1 — 표지
   ======================================================================= */
{
  const s = pres.addSlide();
  s.background = { color: NAVY };

  s.addShape(pres.ShapeType.roundRect, {
    x: -1.6, y: -2.2, w: 7.2, h: 7.2, rectRadius: 0.5,
    fill: { color: NAVY_D }, line: { width: 0 }, rotate: 20,
  });
  s.addShape(pres.ShapeType.roundRect, {
    x: 9.4, y: 4.0, w: 6.2, h: 6.2, rectRadius: 0.5,
    fill: { color: NAVY_D }, line: { width: 0 }, rotate: 20,
  });

  s.addText("SECTION 01", {
    x: M, y: 1.62, w: 6, h: 0.3, margin: 0,
    fontFace: SANS, fontSize: 12.5, bold: true, charSpacing: 3.2, color: ORANGE, valign: "middle",
  });
  s.addText("선박 연료,\n무엇이 달라졌는가", {
    x: M, y: 2.06, w: 9.6, h: 1.86, margin: 0, lineSpacing: 50,
    fontFace: SERIF, fontSize: 42, color: WHITE, valign: "top",
  });
  s.addText("VLSFO · MGO 품질 동향과 선상 대응 — 2024~2025 실측 데이터 기반", {
    x: M, y: 4.02, w: 10, h: 0.34, margin: 0,
    fontFace: SANS, fontSize: 14.5, color: "AFC0D8", valign: "middle",
  });

  const chips = [
    ["83%", "VLSFO + MGO 공급 비중"],
    ["483 ppm", "Cat fines 실측 최대값"],
    ["71%", "정유기 · 필터 관련 이슈"],
  ];
  const cw = 3.5, gap = 0.42;
  chips.forEach(([big, small], i) => {
    const x = M + i * (cw + gap);
    s.addShape(pres.ShapeType.roundRect, {
      x, y: 4.72, w: cw, h: 1.16, rectRadius: 0.12,
      fill: { color: NAVY_D }, line: { color: "2C4E70", width: 0.75 },
    });
    s.addText(big, {
      x: x + 0.26, y: 4.88, w: cw - 0.5, h: 0.5, margin: 0,
      fontFace: SANS, fontSize: 25, bold: true, color: ORANGE, valign: "middle",
    });
    s.addText(small, {
      x: x + 0.26, y: 5.36, w: cw - 0.5, h: 0.32, margin: 0,
      fontFace: SANS, fontSize: 10.5, color: "AFC0D8", valign: "middle",
    });
  });

  src(s, "출처: Viswa Group, Global Fuel Characteristics 2024 / 2025 · CIMAC Guideline, Design and operation of fuel cleaning systems for diesel engines (2024-09 v2)", 6.28);
  foot(s, "01", true);
  s.addNotes("Section 01의 목적은 '연료가 나빠졌다'는 인상이 아니라, 왜 서류상 합격한 연료가 실제로 문제를 일으키는지를 데이터로 보여주는 것입니다. 세 숫자(83% / 483ppm / 71%)를 먼저 던지고 시작합니다.");
}

/* =======================================================================
   S2 — 시장 구조
   ======================================================================= */
{
  const s = pres.addSlide();
  base(s, "02", "01   시장 구조", "지금 배에 들어오는 연료");

  card(s, M, 1.74, 6.42, 4.02);
  s.addChart(
    pres.ChartType.bar,
    [
      { name: "2024", labels: ["VLSFO", "LSMGO"], values: [55, 29] },
      { name: "2025", labels: ["VLSFO", "LSMGO"], values: [50, 33] },
    ],
    {
      x: M + 0.18, y: 1.96, w: 6.06, h: 3.6,
      barDir: "col", barGrouping: "clustered", barGapWidthPct: 60,
      chartColors: [BLUE, ORANGE],
      showLegend: true, legendPos: "t", legendFontSize: 10.5, legendFontFace: SANS, legendColor: BODY,
      showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '0"%"',
      dataLabelFontSize: 11.5, dataLabelFontFace: SANS, dataLabelColor: BODY, dataLabelFontBold: true,
      catAxisLabelFontSize: 12, catAxisLabelFontFace: SANS, catAxisLabelColor: NAVY,
      catAxisLabelFontBold: true, catGridLine: { style: "none" }, catAxisLineShow: false,
      valAxisHidden: true, valGridLine: { style: "none" }, valAxisMaxVal: 65,
      valAxisLineShow: false, showTitle: false,
    }
  );

  const notes = [
    ["VLSFO", "55% → 50%", "여전히 최다 공급 연료. 다만 비중은 줄고 있음", BLUE],
    ["LSMGO", "29% → 33%", "가장 빠르게 늘어난 연료. 증가분을 그대로 흡수", ORANGE],
  ];
  notes.forEach(([nm, delta, desc, col], i) => {
    const y = 1.74 + i * 1.32;
    card(s, M + 6.72, y, 5.37, 1.16);
    s.addText(nm, {
      x: M + 6.96, y: y + 0.14, w: 2.2, h: 0.32, margin: 0,
      fontFace: SANS, fontSize: 13, bold: true, color: col, valign: "middle",
    });
    s.addText(delta, {
      x: M + 9.0, y: y + 0.14, w: 2.9, h: 0.32, margin: 0,
      fontFace: SANS, fontSize: 14, bold: true, color: NAVY, align: "right", valign: "middle",
    });
    s.addText(desc, {
      x: M + 6.96, y: y + 0.54, w: 4.9, h: 0.44, margin: 0,
      fontFace: SANS, fontSize: 11, color: BODY, valign: "top",
    });
  });

  card(s, M + 6.72, 4.38, 5.37, 1.38, NAVY);
  s.addText("83%", {
    x: M + 6.96, y: 4.56, w: 1.9, h: 0.58, margin: 0,
    fontFace: SANS, fontSize: 32, bold: true, color: ORANGE, valign: "middle",
  });
  s.addText("두 연료가 전체 공급의 대부분을 차지합니다.\nVLSFO와 MGO의 특성을 아는 것이 곧 연료 관리입니다.", {
    x: M + 8.86, y: 4.56, w: 3.0, h: 1.0, margin: 0, lineSpacing: 16,
    fontFace: SANS, fontSize: 10.5, color: "D6E0EE", valign: "middle",
  });

  src(s, "출처: Viswa Group, Global Fuel Characteristics 2024 · 2025 (공급 비중 기준)", 6.0);
  s.addNotes("2025년 들어 VLSFO는 줄고 MGO가 늘었습니다. 두 연료를 합치면 83%입니다. 뒤에서 다루겠지만 MGO가 늘었다고 안심할 일은 아닙니다 — 슬라이드 8에서 확인합니다.");
}

/* =======================================================================
   S3 — 착시
   ======================================================================= */
{
  const s = pres.addSlide();
  base(s, "03", "02   착시", "합격증은 안전을 보증하지 않는다");

  card(s, M, 1.74, 3.62, 3.62);
  label(s, "ON PAPER", M + 0.28, 1.96, 2.6);
  s.addText("4.8%", {
    x: M + 0.28, y: 2.24, w: 3.06, h: 0.86, margin: 0,
    fontFace: SANS, fontSize: 50, bold: true, color: NAVY, valign: "middle",
  });
  s.addText("2025년 VLSFO 규격 부적합률", {
    x: M + 0.28, y: 3.12, w: 3.06, h: 0.28, margin: 0,
    fontFace: SANS, fontSize: 11.5, bold: true, color: BODY, valign: "middle",
  });
  s.addText(
    [
      { text: "2024년 4.3% → 2025년 4.8%", options: { fontFace: SANS, fontSize: 11, color: BODY, breakLine: true } },
      { text: "숫자만 보면 20척 중 1척.", options: { fontFace: SANS, fontSize: 11, color: BODY, breakLine: true } },
      { text: "「대체로 문제없다」는 인상을 줍니다.", options: { fontFace: SANS, fontSize: 11, color: BODY } },
    ],
    { x: M + 0.28, y: 3.5, w: 3.06, h: 1.0, margin: 0, lineSpacing: 17, valign: "top" }
  );
  s.addShape(pres.ShapeType.roundRect, {
    x: M + 0.28, y: 4.66, w: 3.06, h: 0.56, rectRadius: 0.08,
    fill: { color: "FDF2E9" }, line: { width: 0 },
  });
  s.addText("그런데 실제 손상은 계속 발생합니다", {
    x: M + 0.38, y: 4.66, w: 2.86, h: 0.56, margin: 0,
    fontFace: SANS, fontSize: 10.5, bold: true, color: ORANGE, align: "center", valign: "middle",
  });

  card(s, M + 3.86, 1.74, 8.23, 3.62);
  label(s, "WHAT GC-MS ACTUALLY FOUND", M + 4.14, 1.96, 6.0, BLUE);
  s.addText("ISO 8217 항목은 전부 정상, 정밀 분석에서만 잡힌 오염", {
    x: M + 4.14, y: 2.22, w: 7.7, h: 0.3, margin: 0,
    fontFace: SANS, fontSize: 13, bold: true, color: NAVY, valign: "middle",
  });

  const ports = [
    ["ARA", "2024.12~2025.1", "산가 전량 1.5 초과 (항구 평균 0.6) · Ca 50~130 ppm (평균 13 ppm)", "7척 — 연료펌프 · 정유기"],
    ["Istanbul", "2025.4", "스티렌 · DCPD 10,000 ppm 초과 · 유리지방산 1,000 ppm 초과", "5척 — 주기 · 발전기 연료펌프"],
    ["New Orleans", "2025.4", "총 산성분 4,300 ppm · 인(P) 75~90 ppm", "플런저 침식부식 · 강재 부식"],
    ["New Orleans", "2025.10", "칼륨(K) 100~250 ppm · 알코올 4 wt% 초과", "4척 — 터보차저 파울링"],
  ];
  ports.forEach(([p, when, found, effect], i) => {
    const y = 2.58 + i * 0.68;
    if (i > 0) {
      s.addShape(pres.ShapeType.line, {
        x: M + 4.14, y: y - 0.04, w: 7.7, h: 0, line: { color: LINE, width: 0.75 },
      });
    }
    s.addText(p, {
      x: M + 4.14, y: y + 0.02, w: 1.6, h: 0.26, margin: 0,
      fontFace: SANS, fontSize: 11.5, bold: true, color: NAVY, valign: "middle",
    });
    s.addText(when, {
      x: M + 4.14, y: y + 0.28, w: 1.6, h: 0.24, margin: 0,
      fontFace: SANS, fontSize: 9, color: MUTED, valign: "middle",
    });
    s.addText(found, {
      x: M + 5.82, y: y + 0.02, w: 4.0, h: 0.5, margin: 0, lineSpacing: 14,
      fontFace: SANS, fontSize: 10, color: BODY, valign: "middle",
    });
    s.addText(effect, {
      x: M + 9.92, y: y + 0.02, w: 1.92, h: 0.5, margin: 0, lineSpacing: 14,
      fontFace: SANS, fontSize: 10, bold: true, color: RED, valign: "middle",
    });
  });

  banner(s, "ISO 8217은 「인도 시점」의 품질을 정의할 뿐, 「엔진 입구」에서 요구되는 품질을 정의하지 않습니다.", 5.54);
  src(s, "출처: Viswa Group 2025 (항구별 사례) · CIMAC Guideline 2024 §2", 6.48);
  s.addNotes("핵심 슬라이드입니다. 부적합률 4.8%는 'ISO 8217 항목' 기준입니다. 네 항구 사례 모두 ISO 항목은 정상이었고, GC-MS 정밀 분석에서만 오염이 잡혔습니다. 즉 시험성적서가 깨끗하다는 것이 연료가 깨끗하다는 뜻은 아닙니다.");
}

/* =======================================================================
   S4 — 중질화
   ======================================================================= */
{
  const s = pres.addSlide();
  base(s, "04", "03   원인 ①", "연료는 더 무거워지고 있다");

  const rows = [
    ["VISCOSITY", "점도 @50°C", "155", "172", "cSt", "분리 온도를 더 높여야 함"],
    ["DENSITY", "밀도 @15°C", "947", "951", "kg/m³", "물 · 고형분 분리가 더 어려움"],
    ["MCR", "잔류탄소분", "6.25", "6.61", "%", "연소 후 퇴적물 증가"],
    ["AL + SI", "촉매 미분 평균", "18", "21", "ppm", "마모성 입자 부하 상승"],
  ];
  const cwid = (CW - 0.36 * 3) / 4;
  rows.forEach(([en, ko, a, b, unit, note], i) => {
    const x = M + i * (cwid + 0.36);
    card(s, x, 1.74, cwid, 3.28);
    label(s, en, x + 0.24, 1.96, cwid - 0.48);
    s.addText(ko, {
      x: x + 0.24, y: 2.2, w: cwid - 0.48, h: 0.3, margin: 0,
      fontFace: SANS, fontSize: 13, bold: true, color: NAVY, valign: "middle",
    });

    s.addText("2024", {
      x: x + 0.24, y: 2.66, w: cwid - 0.48, h: 0.22, margin: 0,
      fontFace: SANS, fontSize: 9, color: MUTED, valign: "middle",
    });
    s.addText(a, {
      x: x + 0.24, y: 2.86, w: cwid - 0.48, h: 0.4, margin: 0,
      fontFace: SANS, fontSize: 20, color: MUTED, valign: "middle",
    });

    s.addShape(pres.ShapeType.line, {
      x: x + 0.24, y: 3.36, w: cwid - 0.48, h: 0, line: { color: LINE, width: 0.75 },
    });

    s.addText("2025", {
      x: x + 0.24, y: 3.46, w: cwid - 0.48, h: 0.22, margin: 0,
      fontFace: SANS, fontSize: 9, bold: true, color: ORANGE, valign: "middle",
    });
    s.addText(
      [
        { text: b, options: { fontFace: SANS, fontSize: 32, bold: true, color: NAVY } },
        { text: "  " + unit, options: { fontFace: SANS, fontSize: 11, color: MUTED } },
      ],
      { x: x + 0.24, y: 3.66, w: cwid - 0.48, h: 0.56, margin: 0, valign: "middle" }
    );

    s.addShape(pres.ShapeType.roundRect, {
      x: x + 0.24, y: 4.32, w: cwid - 0.48, h: 0.5, rectRadius: 0.07,
      fill: { color: "EEF2F9" }, line: { width: 0 },
    });
    s.addText(note, {
      x: x + 0.34, y: 4.32, w: cwid - 0.68, h: 0.5, margin: 0,
      fontFace: SANS, fontSize: 10, color: BLUE, align: "center", valign: "middle",
    });
  });

  banner(s, "점도가 오르면 분리 온도를 더 올려야 합니다. 그런데 VLSFO는 과열되면 오히려 불안정해집니다.", 5.24);
  src(s, "출처: Viswa Group 2024 · 2025, VLSFO 전 세계 평균 (Table 1) · CIMAC Guideline 2024 §7.2", 6.24);
  s.addNotes("네 항목 모두 한 해 사이에 나빠졌습니다. 중요한 건 상충 관계입니다 — 점도가 높아져 분리 온도를 올려야 하는데, CIMAC은 VLSFO가 열에 민감해서 장시간 고온 노출 시 불안정해진다고 경고합니다. 분리기 운전이 그만큼 좁은 창에서 이뤄져야 합니다.");
}

/* =======================================================================
   S5 — Cat fines
   ======================================================================= */
{
  const s = pres.addSlide();
  base(s, "05", "04   원인 ②", "60 ppm 합격, 15 ppm 필요");

  const steps = [
    ["AS DELIVERED", "인도 시점 규격", "60", "ppm", "ISO 8217:2024 상한\n(저점도 등급은 40~50 ppm)", BLUE, "EEF2F9"],
    ["AS SUPPLIED", "실제 공급 최대값", "483", "ppm", "유럽 · ARA 실측 최대\n규격 상한의 8배", RED, "FDECEA"],
    ["AT ENGINE INLET", "엔진 입구 요구치", "7~15", "ppm", "다수 OEM 권장 상한\nCIMAC Guideline §6.3", GREEN, "E9F5F0"],
  ];
  const sw = 3.58, sgap = 0.62;
  steps.forEach(([en, ko, num, unit, desc, col, tint], i) => {
    const x = M + i * (sw + sgap);
    card(s, x, 1.74, sw, 3.3);
    label(s, en, x + 0.26, 1.96, sw - 0.52, col);
    s.addText(ko, {
      x: x + 0.26, y: 2.2, w: sw - 0.52, h: 0.3, margin: 0,
      fontFace: SANS, fontSize: 13, bold: true, color: NAVY, valign: "middle",
    });
    s.addText(
      [
        { text: num, options: { fontFace: SANS, fontSize: 48, bold: true, color: col } },
        { text: " " + unit, options: { fontFace: SANS, fontSize: 14, color: MUTED } },
      ],
      { x: x + 0.26, y: 2.6, w: sw - 0.52, h: 0.94, margin: 0, valign: "middle" }
    );
    s.addShape(pres.ShapeType.roundRect, {
      x: x + 0.26, y: 3.66, w: sw - 0.52, h: 0.78, rectRadius: 0.08,
      fill: { color: tint }, line: { width: 0 },
    });
    s.addText(desc, {
      x: x + 0.36, y: 3.66, w: sw - 0.72, h: 0.78, margin: 0, lineSpacing: 15,
      fontFace: SANS, fontSize: 10, color: BODY, align: "center", valign: "middle",
    });
    if (i < 2) {
      s.addText("▶", {
        x: x + sw + 0.06, y: 3.06, w: 0.5, h: 0.4, margin: 0,
        fontFace: SANS, fontSize: 15, color: "B9C4D6", align: "center", valign: "middle",
      });
    }
  });

  s.addText(
    [
      { text: "촉매 미분(Al+Si)은 모스 경도 8의 세라믹 입자입니다.  ", options: { fontFace: SANS, fontSize: 11, color: BODY } },
      { text: "2행정", options: { fontFace: SANS, fontSize: 11, bold: true, color: NAVY } },
      { text: " 라이너 · 피스톤링 · 링그루브 마모와 스커핑,  ", options: { fontFace: SANS, fontSize: 11, color: BODY } },
      { text: "4행정", options: { fontFace: SANS, fontSize: 11, bold: true, color: NAVY } },
      { text: " 분사노즐 마모 → 불완전 연소 → 터보차저 퇴적으로 이어집니다.", options: { fontFace: SANS, fontSize: 11, color: BODY } },
    ],
    { x: M, y: 5.14, w: CW, h: 0.34, margin: 0, align: "center", valign: "middle" }
  );

  banner(s, "기준점은 벙커 매니폴드가 아니라 엔진 입구에 있습니다. 규격 합격은 출발점일 뿐입니다.", 5.56);
  src(s, "출처: CIMAC Guideline 2024 §6.3 · §14 (엔진 입구 목표: 점도 2~20 cSt, Al+Si 15 ppm 미만, 수분 0.3% V/V 미만) · Viswa Group 2025", 6.56);
  s.addNotes("이 슬라이드가 Section 01 전체의 논리 축입니다. 세 숫자가 각각 다른 지점을 가리킵니다 — 규격은 인도 시점, 실측은 공급 현실, 그리고 OEM이 실제로 요구하는 건 엔진 입구 7~15ppm입니다. 그 간극을 메우는 것이 본선의 연료 처리 시스템입니다.");
}

/* =======================================================================
   S6 — Shelf life
   ======================================================================= */
{
  const s = pres.addSlide();
  base(s, "06", "05   원인 ③", "벙커링 당일 합격, 두 달 뒤 불합격");

  card(s, M, 1.74, 7.66, 3.72);

  const head = [
    { text: "샘플", options: { fontFace: SANS, fontSize: 10.5, bold: true, color: WHITE, align: "left" } },
    { text: "벙커링 당일", options: { fontFace: SANS, fontSize: 10.5, bold: true, color: WHITE, align: "center" } },
    { text: "문제 발생 시", options: { fontFace: SANS, fontSize: 10.5, bold: true, color: WHITE, align: "center" } },
    { text: "경과 기간", options: { fontFace: SANS, fontSize: 10.5, bold: true, color: WHITE, align: "center" } },
    { text: "판정", options: { fontFace: SANS, fontSize: 10.5, bold: true, color: WHITE, align: "center" } },
  ];
  const data = [
    ["VLSFO 01", "0.08%", "0.29%", "5 개월"],
    ["VLSFO 02", "0.05%", "0.40%", "2 개월"],
    ["VLSFO 03", "0.03%", "0.10%", "1 개월"],
    ["VLSFO 05", "0.04%", "0.30%", "2 개월"],
  ];
  const body = data.map(([a, b, c, d]) => [
    { text: a, options: { fontFace: SANS, fontSize: 11, bold: true, color: NAVY, align: "left" } },
    { text: b, options: { fontFace: SANS, fontSize: 12, bold: true, color: GREEN, align: "center" } },
    { text: c, options: { fontFace: SANS, fontSize: 12, bold: true, color: RED, align: "center" } },
    { text: d, options: { fontFace: SANS, fontSize: 11, color: BODY, align: "center" } },
    { text: "합격 → 불합격", options: { fontFace: SANS, fontSize: 10, color: BODY, align: "center" } },
  ]);

  s.addTable([head, ...body], {
    x: M + 0.26, y: 2.28, w: 7.14, colW: [1.5, 1.42, 1.42, 1.4, 1.4],
    rowH: [0.36, 0.5, 0.5, 0.5, 0.5],
    border: { type: "solid", color: LINE, pt: 0.75 },
    fill: { color: WHITE },
    valign: "middle", margin: 4,
  });
  // header fill
  s.addShape(pres.ShapeType.rect, {
    x: M + 0.26, y: 2.28, w: 7.14, h: 0.36, fill: { color: NAVY }, line: { width: 0 },
  });
  s.addText(
    [
      { text: "샘플", options: { fontFace: SANS, fontSize: 10.5, bold: true, color: WHITE } },
    ],
    { x: M + 0.3, y: 2.28, w: 1.42, h: 0.36, margin: 0, valign: "middle" }
  );
  ["벙커링 당일", "문제 발생 시", "경과 기간", "판정"].forEach((t, i) => {
    const xs = [1.5, 2.92, 4.34, 5.74];
    const ws = [1.42, 1.42, 1.4, 1.4];
    s.addText(t, {
      x: M + 0.26 + xs[i], y: 2.28, w: ws[i], h: 0.36, margin: 0,
      fontFace: SANS, fontSize: 10.5, bold: true, color: WHITE, align: "center", valign: "middle",
    });
  });

  label(s, "ISO 8217 TOTAL SEDIMENT LIMIT  0.10%", M + 0.26, 1.96, 5.4, BLUE);
  s.addText("전 샘플이 벙커링 당일 규격을 통과했습니다.", {
    x: M + 0.26, y: 4.86, w: 7.14, h: 0.3, margin: 0,
    fontFace: SANS, fontSize: 11.5, bold: true, color: NAVY, valign: "middle",
  });

  card(s, M + 7.9, 1.74, 4.19, 3.72, NAVY);
  s.addText("왜 중요한가", {
    x: M + 8.16, y: 1.98, w: 3.67, h: 0.32, margin: 0,
    fontFace: SANS, fontSize: 13, bold: true, color: ORANGE, valign: "middle",
  });
  s.addText(
    [
      { text: "시험성적서는 「그날」의 상태입니다.", options: { fontFace: SANS, fontSize: 11.5, bold: true, color: WHITE, breakLine: true } },
      { text: "VLSFO는 저장 중에도 계속 변합니다.", options: { fontFace: SANS, fontSize: 11.5, color: "D6E0EE", breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "슬러징의 4대 원인", options: { fontFace: SANS, fontSize: 10.5, bold: true, color: ORANGE, breakLine: true } },
      { text: "① 초기 침전물 과다", options: { fontFace: SANS, fontSize: 10.5, color: "D6E0EE", breakLine: true } },
      { text: "② 시간 경과에 따른 열화", options: { fontFace: SANS, fontSize: 10.5, color: "D6E0EE", breakLine: true } },
      { text: "③ 블렌드 간 불호환", options: { fontFace: SANS, fontSize: 10.5, color: "D6E0EE", breakLine: true } },
      { text: "④ 오염물 혼입", options: { fontFace: SANS, fontSize: 10.5, color: "D6E0EE", breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "VLSFO는 HSFO보다 열화에 취약하고 저장 수명이 짧습니다.", options: { fontFace: SANS, fontSize: 10.5, italic: true, color: "AFC0D8" } },
    ],
    { x: M + 8.16, y: 2.4, w: 3.67, h: 2.86, margin: 0, lineSpacing: 17, valign: "top" }
  );

  banner(s, "TSP · TSA 시험 결과는 VLSFO의 장기 안정성을 보증하지 않습니다.", 5.62);
  src(s, "출처: Viswa Group 2025, Table 3 (실제 문제 사례 샘플) · CIMAC Congress 2023, Paper No. 131", 6.62);
  s.addNotes("Section 01에서 가장 강한 카드입니다. 네 샘플 모두 벙커링 당일에는 ISO 한계 0.10% 아래였습니다. 그런데 1~5개월 뒤 최대 0.40%까지 올라갔습니다. 즉 '받을 때 합격'이 '쓸 때 안전'을 뜻하지 않습니다. 이것이 저장·관리가 시험성적서만큼 중요한 이유입니다.");
}

/* =======================================================================
   S7 — Purifier / Filter
   ======================================================================= */
{
  const s = pres.addSlide();
  base(s, "07", "06   현상", "문제는 결국 정유기에서 터진다");

  card(s, M, 1.74, 5.5, 3.62);
  s.addChart(
    pres.ChartType.doughnut,
    [{
      name: "VLSFO 운항 이슈",
      labels: ["정유기 막힘", "정유기 + 필터", "필터 막힘", "펌프 고착", "기타"],
      values: [40, 21, 9, 1, 29],
    }],
    {
      x: M + 0.3, y: 1.94, w: 4.9, h: 3.24,
      holeSize: 62,
      chartColors: [NAVY, BLUE, "6E93BF", ORANGE, "DDE4EF"],
      showLegend: true, legendPos: "b", legendFontSize: 10, legendFontFace: SANS, legendColor: BODY,
      showValue: true, dataLabelPosition: "ctr", dataLabelFormatCode: '0"%"',
      dataLabelFontSize: 10.5, dataLabelFontFace: SANS, dataLabelColor: WHITE, dataLabelFontBold: true,
      showTitle: false,
    }
  );
  s.addText(
    [
      { text: "71", options: { fontFace: SANS, fontSize: 40, bold: true, color: ORANGE } },
      { text: "%", options: { fontFace: SANS, fontSize: 20, bold: true, color: ORANGE } },
    ],
    { x: M + 2.0, y: 2.94, w: 1.5, h: 0.5, margin: 0, align: "center", valign: "middle" }
  );
  s.addText("정유기·필터", {
    x: M + 2.0, y: 3.4, w: 1.5, h: 0.24, margin: 0,
    fontFace: SANS, fontSize: 8.5, color: BODY, align: "center", valign: "middle",
  });

  card(s, M + 5.78, 1.74, 6.31, 3.62);
  label(s, "WHAT IT MEANS ON BOARD", M + 6.06, 1.96, 5.0, BLUE);
  s.addText("연료 문제는 엔진이 아니라 정화 장비에서 먼저 나타납니다", {
    x: M + 6.06, y: 2.22, w: 5.75, h: 0.32, margin: 0,
    fontFace: SANS, fontSize: 13, bold: true, color: NAVY, valign: "middle",
  });

  const items = [
    ["정유기 막힘", "40%", "슬러지 부하가 분리 능력을 초과. 심하면 보울 손상까지"],
    ["정유기 + 필터 동시", "21%", "이미 불안정해진 연료가 계통 전체를 막는 단계"],
    ["필터 막힘", "9%", "왁스 석출 또는 응집된 아스팔텐"],
    ["연료펌프 고착", "1%", "가장 늦게, 가장 비싸게 드러나는 증상"],
  ];
  items.forEach(([t, pct, d], i) => {
    const y = 2.62 + i * 0.66;
    if (i > 0) {
      s.addShape(pres.ShapeType.line, {
        x: M + 6.06, y: y - 0.05, w: 5.75, h: 0, line: { color: LINE, width: 0.75 },
      });
    }
    s.addShape(pres.ShapeType.roundRect, {
      x: M + 6.06, y: y + 0.08, w: 0.62, h: 0.3, rectRadius: 0.06,
      fill: { color: i === 3 ? "FDF2E9" : "EEF2F9" }, line: { width: 0 },
    });
    s.addText(pct, {
      x: M + 6.06, y: y + 0.08, w: 0.62, h: 0.3, margin: 0,
      fontFace: SANS, fontSize: 10.5, bold: true, color: i === 3 ? ORANGE : BLUE,
      align: "center", valign: "middle",
    });
    s.addText(t, {
      x: M + 6.8, y: y + 0.04, w: 2.1, h: 0.38, margin: 0,
      fontFace: SANS, fontSize: 11.5, bold: true, color: NAVY, valign: "middle",
    });
    s.addText(d, {
      x: M + 8.94, y: y + 0.02, w: 2.87, h: 0.44, margin: 0, lineSpacing: 14,
      fontFace: SANS, fontSize: 10, color: BODY, valign: "middle",
    });
  });

  banner(s, "2025년 보고된 VLSFO 문제의 70%가 정유기 슬러징과 필터 막힘이었습니다.", 5.54);
  src(s, "출처: Viswa Group 2025, Figure 12 (VLSFO 운항 이슈 분포)", 6.48);
  s.addNotes("정유기와 필터는 연료 상태를 가장 먼저 알려주는 지표입니다. 슬러지 배출량이 늘거나 필터 차압이 빨리 오르면 그 자체가 경고 신호입니다. 엔진에서 증상이 보일 때는 이미 늦습니다.");
}

/* =======================================================================
   S8 — MGO
   ======================================================================= */
{
  const s = pres.addSlide();
  base(s, "08", "07   사각지대", "MGO는 안전지대가 아니다");

  const cards3 = [
    ["FILTER CHOKING", "필터 막힘", "67", "%", "MGO 운항 이슈 중 최다.\n오염물 혼입 또는 왁스 석출이 원인", NAVY],
    ["FLASH POINT", "인화점 미달", "28", "%", "MGO 부적합 항목 2위.\n품질이 아니라 화재 안전 문제", RED],
    ["BIO-MGO", "바이오 혼합 MGO", "19", "%", "일반 MGO 부적합률 3.0%의 6배.\nBDN 표기와 실제 FAME 함량 불일치", ORANGE],
  ];
  const w3 = (CW - 0.42 * 2) / 3;
  cards3.forEach(([en, ko, num, unit, desc, col], i) => {
    const x = M + i * (w3 + 0.42);
    card(s, x, 1.74, w3, 2.80);
    label(s, en, x + 0.26, 1.96, w3 - 0.52, col === NAVY ? BLUE : col);
    s.addText(ko, {
      x: x + 0.26, y: 2.2, w: w3 - 0.52, h: 0.3, margin: 0,
      fontFace: SANS, fontSize: 13.5, bold: true, color: NAVY, valign: "middle",
    });
    s.addText(
      [
        { text: num, options: { fontFace: SANS, fontSize: 46, bold: true, color: col } },
        { text: unit, options: { fontFace: SANS, fontSize: 18, bold: true, color: col } },
      ],
      { x: x + 0.26, y: 2.6, w: w3 - 0.52, h: 0.86, margin: 0, valign: "middle" }
    );
    s.addText(desc, {
      x: x + 0.26, y: 3.54, w: w3 - 0.52, h: 0.72, margin: 0, lineSpacing: 16,
      fontFace: SANS, fontSize: 10.5, color: BODY, valign: "top",
    });
  });

  card(s, M, 4.70, CW, 0.70, "FDF2E9");
  s.addText(
    [
      { text: "MGO 부적합 항목 구성  ", options: { fontFace: SANS, fontSize: 10.5, bold: true, color: ORANGE } },
      { text: "유동점 31%  ·  인화점 28%  ·  황분 12%  ·  외관 9%  ·  수분 7%  ·  점도 6%  ·  총침전물 3%  ·  회분 2%  ·  윤활성 2%",
        options: { fontFace: SANS, fontSize: 10.5, color: BODY } },
    ],
    { x: M + 0.3, y: 4.70, w: CW - 0.6, h: 0.70, margin: 0, valign: "middle" }
  );

  banner(s, "MGO 공급이 늘었다고 안심할 수 없습니다. 인화점 미달은 품질 문제가 아니라 안전 문제입니다.", 5.56);
  src(s, "출처: Viswa Group 2025, Figure 9 · Figure 16 (MGO 부적합 항목 및 운항 이슈 분포)", 6.50);
  s.addNotes("MGO는 '깨끗한 연료'라는 인식이 있지만 운항 이슈의 67%가 필터 막힘입니다. 특히 인화점 미달 28%는 MARPOL·SOLAS상 60°C 최저 기준과 직결되는 안전 항목입니다. 바이오 혼합 MGO는 부적합률이 일반 MGO의 6배입니다.");
}

/* =======================================================================
   S9 — 분리기의 경계
   ======================================================================= */
{
  const s = pres.addSlide();
  base(s, "09", "08   경계", "원심분리기가 하는 일, 하지 못하는 일");

  card(s, M, 1.74, 5.86, 3.5);
  label(s, "WITHIN THE SCOPE OF FUEL CLEANING", M + 0.28, 1.96, 5.3, GREEN);
  s.addText("기계적 정화로 처리되는 영역", {
    x: M + 0.28, y: 2.22, w: 5.3, h: 0.32, margin: 0,
    fontFace: SANS, fontSize: 14, bold: true, color: NAVY, valign: "middle",
  });
  const canDo = [
    ["촉매 미분 (Al+Si)", "CIMAC: 효율적인 정화 시스템이 유일하게 효과적인 방법"],
    ["수분 (담수 · 해수)", "엔진 입구 0.3% V/V 미만 유지"],
    ["침전물 · 응집 슬러지", "충분히 큰 입자로 응집된 경우에 한함"],
  ];
  canDo.forEach(([t, d], i) => {
    const y = 2.66 + i * 0.78;
    s.addShape(pres.ShapeType.roundRect, {
      x: M + 0.28, y: y + 0.04, w: 0.26, h: 0.26, rectRadius: 0.13,
      fill: { color: GREEN }, line: { width: 0 },
    });
    s.addText(t, {
      x: M + 0.68, y: y, w: 5.0, h: 0.32, margin: 0,
      fontFace: SANS, fontSize: 12, bold: true, color: NAVY, valign: "middle",
    });
    s.addText(d, {
      x: M + 0.68, y: y + 0.32, w: 5.0, h: 0.4, margin: 0, lineSpacing: 14,
      fontFace: SANS, fontSize: 10, color: BODY, valign: "top",
    });
  });
  s.addText("단, CFR 규격 · 분리 온도 · 유량이 모두 맞을 때에 한합니다.", {
    x: M + 0.28, y: 4.82, w: 5.3, h: 0.28, margin: 0,
    fontFace: SANS, fontSize: 10, italic: true, color: MUTED, valign: "middle",
  });

  card(s, M + 6.23, 1.74, 5.86, 3.5, NAVY);
  label(s, "OUTSIDE THE SCOPE OF FUEL CLEANING", M + 6.51, 1.96, 5.3, ORANGE);
  s.addText("기계적 정화로 처리되지 않는 영역", {
    x: M + 6.51, y: 2.22, w: 5.3, h: 0.32, margin: 0,
    fontFace: SANS, fontSize: 14, bold: true, color: WHITE, valign: "middle",
  });
  const cannot = [
    ["화학 오염물", "CNSL · 유리지방산 · 스티렌 · 알코올 — 용해되어 있어 분리 불가"],
    ["시간 경과 안정성 열화", "저장 중 진행. 받을 때의 시험 결과로는 예측되지 않음"],
    ["블렌드 간 불호환", "혼합 자체가 원인. CIMAC은 서로 다른 연료의 혼합을 권하지 않음"],
  ];
  cannot.forEach(([t, d], i) => {
    const y = 2.66 + i * 0.78;
    s.addShape(pres.ShapeType.roundRect, {
      x: M + 6.51, y: y + 0.04, w: 0.26, h: 0.26, rectRadius: 0.13,
      fill: { color: ORANGE }, line: { width: 0 },
    });
    s.addText(t, {
      x: M + 6.91, y: y, w: 5.0, h: 0.32, margin: 0,
      fontFace: SANS, fontSize: 12, bold: true, color: WHITE, valign: "middle",
    });
    s.addText(d, {
      x: M + 6.91, y: y + 0.32, w: 5.0, h: 0.4, margin: 0, lineSpacing: 14,
      fontFace: SANS, fontSize: 10, color: "AFC0D8", valign: "top",
    });
  });
  s.addText("→ Section 02 에서 다룹니다", {
    x: M + 6.51, y: 4.82, w: 5.3, h: 0.28, margin: 0,
    fontFace: SANS, fontSize: 10.5, bold: true, italic: true, color: ORANGE, valign: "middle",
  });

  banner(s, "일상 분석으로 검출되지 않는 유해 물질은 어떤 연료 정화 시스템의 처리 범위 밖에 있습니다.", 5.46);
  src(s, "출처: CIMAC Guideline 2024 §6.3 · §6.4 · §7.2 — 인용문은 §6.4 원문", 6.46);
  s.addNotes("Section 01의 결론입니다. 분리기를 잘 돌리는 것은 필수지만, 분리기가 손댈 수 없는 영역이 분명히 존재한다는 것을 CIMAC이 직접 명시하고 있습니다. Section 02는 바로 이 영역을 다룹니다 — 분리기의 대체가 아니라, 분리기가 닿지 못하는 구간에 대한 보완입니다.");
}

pres.writeFile({ fileName: "/tmp/claude-0/-home-user-github-test/9f72a13f-ac8c-5519-89c5-583ea05e90c6/scratchpad/fuel_section01.pptx" })
  .then(f => console.log("WROTE " + f));
