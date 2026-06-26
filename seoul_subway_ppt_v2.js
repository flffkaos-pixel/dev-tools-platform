const pptxgen = require("pptxgenjs");

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.title = '서울 지하철 유동인구 기반 주거·업무 지역 분류 및 팝업스토어 입지 전략 분석';
pres.author = '이어드림스쿨6기';

const C = {
  navy:  "1B3A6B", blue:  "2980B9", red:   "C0392B",
  gray:  "7F8C8D", light: "ECF0F1", white: "FFFFFF",
  accent:"F39C12", dark:  "1C2833", mid:   "34495E",
};

const mkShadow = () => ({ type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.12 });

// ── Slide 1 : Cover ─────────────────────────────────────────
{
  let sl = pres.addSlide();
  sl.background = { color: C.navy };
  sl.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:10, h:0.08, fill:{color:C.accent} });

  sl.addShape(pres.shapes.OVAL, { x:0.7, y:1.3, w:1.2, h:1.2, fill:{color:C.blue}, shadow:mkShadow() });
  sl.addShape(pres.shapes.OVAL, { x:1.3, y:1.5, w:1.0, h:1.0, fill:{color:C.red},  shadow:mkShadow() });
  sl.addShape(pres.shapes.OVAL, { x:0.9, y:2.0, w:0.8, h:0.8, fill:{color:C.gray}, shadow:mkShadow() });

  sl.addText("서울 지하철 유동인구 기반", { x:2.6, y:1.1, w:7, h:0.7,  fontSize:28, bold:true, color:C.white, align:"left" });
  sl.addText("주거·업무 지역 분류 및",      { x:2.6, y:1.75,w:7, h:0.65, fontSize:26, bold:true, color:C.accent, align:"left" });
  sl.addText("팝업스토어 입지 전략 분석",    { x:2.6, y:2.35,w:7, h:0.65, fontSize:26, bold:true, color:C.white, align:"left" });
  sl.addShape(pres.shapes.RECTANGLE, { x:2.6, y:3.1, w:2.5, h:0.05, fill:{color:C.accent} });
  sl.addText("출근 시간대(07~09시) 승하차 데이터를 활용한\n주거 타운 / 오피스 타운 / 복합 지역 분류 모델", {
    x:2.6, y:3.3, w:6.8, h:0.8, fontSize:13, color:"BDC3C7", align:"left"
  });
  sl.addShape(pres.shapes.RECTANGLE, { x:0, y:5.1, w:10, h:0.525, fill:{color:C.dark} });
  sl.addText("이어드림스쿨6기  |  데이터 분석 프로젝트", { x:0.5, y:5.15, w:9, h:0.42, fontSize:11, color:"7F8C8D", align:"left", valign:"middle" });
}

// ── Slide 2 : Agenda ─────────────────────────────────────────
{
  let sl = pres.addSlide();
  sl.background = { color: C.light };
  sl.addText("목차", { x:0.5, y:0.3, w:9, h:0.6, fontSize:32, bold:true, color:C.navy });
  sl.addShape(pres.shapes.RECTANGLE, { x:0.5, y:0.85, w:1.2, h:0.06, fill:{color:C.accent} });
  const items = [
    { n:"01", t:"분석 배경",       s:"Situation · Complication · Question · Answer" },
    { n:"02", t:"분석 방법론",     s:"지역특성지표 X 정의 및 통계적 가설 설정" },
    { n:"03", t:"가설 검정 및 통계 검증", s:"Kruskal-Wallis H Test / 효과크기 분석" },
    { n:"04", t:"시각화 분석",     s:"박스플롯 · 막대그래프 기반 데이터 해석" },
    { n:"05", t:"클러스터별 전략 제안", s:"주거 타운 · 오피스 타운 · 복합 지역 맞춤 전략" },
    { n:"06", t:"결론 및 제언",    s:"공간 목적성 × 시간 방향성 매칭 프레임워크" },
  ];
  items.forEach((it, i) => {
    const y = 1.1 + i * 0.72;
    sl.addShape(pres.shapes.OVAL, { x:0.5, y, w:0.55, h:0.55, fill:{color:C.navy} });
    sl.addText(it.n, { x:0.5, y, w:0.55, h:0.55, fontSize:13, bold:true, color:C.white, align:"center", valign:"middle" });
    sl.addText(it.t, { x:1.2, y, w:8, h:0.32, fontSize:16, bold:true, color:C.dark });
    sl.addText(it.s, { x:1.2, y:y+0.32, w:8, h:0.28, fontSize:11, color:C.gray });
  });
}

// ── Slide 3 : SCQA Background ───────────────────────────────
{
  let sl = pres.addSlide();
  sl.background = { color: C.white };
  sl.addText("01  분석 배경", { x:0.5, y:0.3, w:9, h:0.55, fontSize:26, bold:true, color:C.navy });
  sl.addShape(pres.shapes.RECTANGLE, { x:0.5, y:0.82, w:1.5, h:0.05, fill:{color:C.accent} });

  const scqa = [
    { label:"Situation",     color:C.navy,  text:"서울 전역 지하철 유동인구는 도시 공간 구조에 따라 독특한 주기성과 방향성을 지니고 있으며, 기업들이 팝업스토어를 마케팅 수단으로 적극 활용하고 있습니다." },
    { label:"Complication",  color:C.red,   text:"단순 총유동인구 중심 입지 선정은 높은 임대료 대비 타겟 고객 불일치를 초래하며, 주거지·업무지의 상반된 시간대별 특성을 반영하지 못해 마케팅 효율이 저하됩니다." },
    { label:"Question",      color:C.blue,  text:"출근 시간대(07~09시) 승하차 데이터를 통해 역사를 '주거 타운 / 오피스 타운 / 복합 지역'으로 분류하고, 팝업스토어 최적 입지·운영 전략은 무엇인가?" },
    { label:"Answer",        color:C.accent,text:"지역특성지표 X 수식을 통해 3클러스터로 분류하고, 비모수 통계 검정으로 유의성 검증. 오피스 타운 '평일 5일 압축 운영', 주거 타운 '주말·야간 집중 운영' 전략 도출." },
  ];
  scqa.forEach((q, i) => {
    const y = 1.0 + i * 1.1;
    sl.addShape(pres.shapes.RECTANGLE, { x:0.5, y, w:0.09, h:0.95, fill:{color:q.color} });
    sl.addShape(pres.shapes.RECTANGLE, { x:0.68, y, w:8.82, h:0.95, fill:{color:C.light}, shadow:mkShadow() });
    sl.addText(q.label, { x:0.85, y:y+0.08, w:2, h:0.32, fontSize:13, bold:true, color:q.color });
    sl.addText(q.text,  { x:0.85, y:y+0.4,  w:8.4, h:0.5,  fontSize:12, color:C.dark });
  });
}

// ── Slide 4 : Methodology Formula ───────────────────────────
{
  let sl = pres.addSlide();
  sl.background = { color: C.dark };
  sl.addText("02  분석 방법론", { x:0.5, y:0.3, w:9, h:0.55, fontSize:26, bold:true, color:C.white });
  sl.addShape(pres.shapes.RECTANGLE, { x:0.5, y:0.82, w:1.5, h:0.05, fill:{color:C.accent} });

  sl.addText("지역특성지표 X", { x:0.5, y:1.1, w:9, h:0.5, fontSize:20, bold:true, color:C.accent });
  sl.addShape(pres.shapes.RECTANGLE, { x:0.5, y:1.65, w:9, h:1.1, fill:{color:C.navy}, shadow:mkShadow() });
  sl.addText("X  =  ( P 승차  −  P 하차 )  /  ( P 승차  +  P 하차 )", {
    x:0.5, y:1.65, w:9, h:1.1, fontSize:26, fontFace:"Consolas", bold:true, color:C.white, align:"center", valign:"middle"
  });

  const zones = [
    { x:0.5,  color:C.blue, title:"X → +1", sub:"주거 타운 Bed Town",     desc:"출근 시간대\n승차 압도적" },
    { x:3.6,  color:C.red,  title:"X → −1", sub:"오피스 타운 Office Town",desc:"출근 시간대\n하차 압도적" },
    { x:6.7,  color:C.gray, title:"X → 0",  sub:"복합 지역 Mixed Zone",   desc:"승차·하차\n균형 유지" },
  ];
  zones.forEach(z => {
    sl.addShape(pres.shapes.RECTANGLE, { x:z.x, y:2.95, w:2.8, h:2.25, fill:{color:C.mid}, shadow:mkShadow() });
    sl.addShape(pres.shapes.RECTANGLE, { x:z.x, y:2.95, w:2.8, h:0.06, fill:{color:z.color} });
    sl.addText(z.title, { x:z.x, y:3.05, w:2.8, h:0.5, fontSize:18, bold:true, color:z.color, align:"center" });
    sl.addText(z.sub,  { x:z.x, y:3.52, w:2.8, h:0.65,fontSize:13, bold:true, color:C.white, align:"center" });
    sl.addText(z.desc, { x:z.x, y:4.2,  w:2.8, h:0.8, fontSize:11, color:"BDC3C7", align:"center" });
  });
  sl.addShape(pres.shapes.RECTANGLE, { x:0.5, y:5.28, w:4.3, h:0.05, fill:{color:C.blue} });
  sl.addText("H₀ : 지역유형별 지표 X 차이 없음    H₁ : 통계적으로 유의미한 차이 존재", {
    x:0.5, y:5.1, w:9, h:0.38, fontSize:11, color:"95A5A6", italic:true
  });
}

// ── Slide 5 : Statistical Validation ─────────────────────────
{
  let sl = pres.addSlide();
  sl.background = { color: C.white };
  sl.addText("03  통계적 가설 검정", { x:0.5, y:0.3, w:9, h:0.55, fontSize:26, bold:true, color:C.navy });
  sl.addShape(pres.shapes.RECTANGLE, { x:0.5, y:0.82, w:1.5, h:0.05, fill:{color:C.accent} });

  // Two big stat boxes
  const stats2 = [
    { label:"Kruskal-Wallis H 통계량", val:"263.84", color:C.navy },
    { label:"p-value (유의확률)",       val:"1.2×10⁻⁵⁷", color:C.red },
  ];
  stats2.forEach((s, i) => {
    const bx = 0.5 + i * 4.7;
    sl.addShape(pres.shapes.RECTANGLE, { x:bx, y:1.05, w:4.4, h:1.55, fill:{color:C.light}, shadow:mkShadow() });
    sl.addShape(pres.shapes.RECTANGLE, { x:bx, y:1.05, w:0.09, h:1.55, fill:{color:s.color} });
    sl.addText(s.val, { x:bx+0.2, y:1.15, w:4.0, h:0.85, fontSize:34, bold:true, color:s.color, align:"center" });
    sl.addText(s.label, { x:bx+0.2, y:2.0, w:4.0, h:0.42, fontSize:13, color:C.dark, align:"center" });
  });

  // Verdict
  sl.addShape(pres.shapes.RECTANGLE, { x:0.5, y:2.78, w:9, h:0.85, fill:{color:C.navy}, shadow:mkShadow() });
  sl.addText("✅  귀무가설(H₀) 강력 기각 — 지역유형별 출근 시간대 승하차 구조 차이가 우연이 아님 통계적으로 입증", {
    x:0.7, y:2.78, w:8.6, h:0.85, fontSize:14, bold:true, color:C.white, valign:"middle"
  });

  // Effect size
  sl.addShape(pres.shapes.RECTANGLE, { x:0.5, y:3.82, w:4.4, h:1.0, fill:{color:C.light}, shadow:mkShadow() });
  sl.addShape(pres.shapes.RECTANGLE, { x:0.5, y:3.82, w:0.09, h:1.0, fill:{color:C.accent} });
  sl.addText("효과크기 η²H  =  0.32 (Large)", { x:0.7, y:3.9,  w:4.0, h:0.5, fontSize:16, bold:true, color:C.accent });
  sl.addText("공간 분류 모델이 전체 분산의 약 32%를 설명하며,\n강건한 실질적 유의성을 갖춤", { x:0.7, y:4.4, w:4.0, h:0.5, fontSize:11, color:C.dark });

  // Table
  sl.addText("클러스터별 표본 정보", { x:5.2, y:3.82, w:4.3, h:0.35, fontSize:13, bold:true, color:C.navy });
  sl.addTable([
    [{ text:"지역 유형", options:{bold:true,fill:{color:C.navy},color:C.white} },
     { text:"N",   options:{bold:true,fill:{color:C.navy},color:C.white} },
     { text:"평균 지표 X̄", options:{bold:true,fill:{color:C.navy},color:C.white} }],
    [{ text:"주거 타운",    options:{fill:{color:"D6EAF8"}} }, "142", {text:"+0.63",options:{bold:true,color:C.blue}}],
    [{ text:"오피스 타운",  options:{fill:{color:"FDEDEC"}} }, "85",  {text:"−0.71",options:{bold:true,color:C.red}}],
    [{ text:"복합 지역",    options:{fill:{color:"F2F3F4"}} }, "118", {text:"+0.02",options:{bold:true,color:C.gray}}],
  ], {
    x:5.2, y:4.22, w:4.3, h:1.0,
    fontSize:11, color:C.dark, align:"center", valign:"middle",
    border:{pt:0.5, color:"D5D8DC"},
  });
}

// ── Slide 6 : Chart 1 (Boxplot) ─────────────────────────────
{
  let sl = pres.addSlide();
  sl.background = { color: C.white };
  sl.addText("04  데이터 시각화 — 박스플롯", { x:0.5, y:0.3, w:9, h:0.55, fontSize:26, bold:true, color:C.navy });
  sl.addShape(pres.shapes.RECTANGLE, { x:0.5, y:0.82, w:1.5, h:0.05, fill:{color:C.accent} });

  // Key insight bar
  sl.addShape(pres.shapes.RECTANGLE, { x:0.5, y:0.98, w:9, h:0.62, fill:{color:"FEF9E7"}, shadow:mkShadow() });
  sl.addShape(pres.shapes.RECTANGLE, { x:0.5, y:0.98, w:0.09, h:0.62, fill:{color:C.accent} });
  sl.addText("📊 핵심 관찰: 주거 타운(X̄=+0.63), 오피스 타운(X̄=−0.71), 복합 지역(X̄≈0)으로 명확히 분리가능. 지표 X가 0을 중심으로 대칭적으로 분포하지 않으며 극단적 값을 가짐.", {
    x:0.7, y:0.98, w:8.7, h:0.62, fontSize:11.5, color:C.dark, valign:"middle"
  });

  // Insert boxplot chart image
  sl.addImage({
    path:"C:/Users/중진공39/Desktop/chart_boxplot.png",
    x:0.5, y:1.75, w:9, h:3.7,
  });
}

// ── Slide 7 : Chart 2 (Barplot) ─────────────────────────────
{
  let sl = pres.addSlide();
  sl.background = { color: C.white };
  sl.addText("04  데이터 시각화 — 막대그래프", { x:0.5, y:0.3, w:9, h:0.55, fontSize:26, bold:true, color:C.navy });
  sl.addShape(pres.shapes.RECTANGLE, { x:0.5, y:0.82, w:1.5, h:0.05, fill:{color:C.accent} });

  sl.addShape(pres.shapes.RECTANGLE, { x:0.5, y:0.98, w:9, h:0.62, fill:{color:"FEF9E7"}, shadow:mkShadow() });
  sl.addShape(pres.shapes.RECTANGLE, { x:0.5, y:0.98, w:0.09, h:0.62, fill:{color:C.accent} });
  sl.addText("📊 핵심 관찰: 주거 타운은 출근 승차(185K)가 하차(42K)의 4.4배. 오피스 타운은 반대로 하차(245K)가 승차(39K)의 6.3배. 복합 지역은 승차·하차가 거의 균형(약 92K).", {
    x:0.7, y:0.98, w:8.7, h:0.62, fontSize:11.5, color:C.dark, valign:"middle"
  });

  sl.addImage({
    path:"C:/Users/중진공39/Desktop/chart_bar.png",
    x:0.5, y:1.75, w:9, h:3.7,
  });
}

// ── Slide 8 : Office Town Strategy ─────────────────────────
{
  let sl = pres.addSlide();
  sl.background = { color: C.white };
  sl.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:0.2, h:5.625, fill:{color:C.red} });
  sl.addText("05  오피스 타운 (Office Town) 입지 전략", { x:0.5, y:0.3, w:9, h:0.55, fontSize:24, bold:true, color:C.navy });
  sl.addShape(pres.shapes.RECTANGLE, { x:0.5, y:0.82, w:1.5, h:0.05, fill:{color:C.red} });

  sl.addShape(pres.shapes.RECTANGLE, { x:0.5, y:1.05, w:1.6, h:0.5, fill:{color:C.red}, shadow:mkShadow() });
  sl.addText("RED ZONE", { x:0.5, y:1.05, w:1.6, h:0.5, fontSize:13, bold:true, color:C.white, align:"center", valign:"middle" });
  sl.addText("평일 5일 압축 운영 스케줄", { x:2.3, y:1.05, w:6, h:0.5, fontSize:18, bold:true, color:C.dark, valign:"middle" });

  sl.addShape(pres.shapes.RECTANGLE, { x:0.5, y:1.7, w:9, h:1.0, fill:{color:"FDEDEC"}, shadow:mkShadow() });
  sl.addText("⏰  골든 아워", { x:0.65, y:1.75, w:3, h:0.35, fontSize:14, bold:true, color:C.red });
  sl.addText("평일 점심시간 11:30~13:30  及  퇴근 시간 18:00~20:00  집중 구조  |  주말 유동인구 소멸", { x:0.65, y:2.1, w:8.7, h:0.3, fontSize:12, color:C.dark });
  sl.addText("✓ 월~금 5일 압축 운영 권장   |   주말 운영 지양 (비용 대비 효과 극히 낮음)", { x:0.65, y:2.4, w:8.7, h:0.28, fontSize:11, bold:true, color:C.red });

  const strats = [
    { title:"구매력 높은 직장인 타겟", pts:["테이크아웃 F&B 브랜드","직장인 에센셜 헬스케어","고단가 소형 가젯"] },
    { title:"즉각적 단기 몰입형 체험", pts:["스트레스 해소형 복합 공간","스마트워치·이어폰 체험존","인테리어 샘플링 부스"] },
    { title:"입지 선정 기준",         pts:["역 출근 하차인구 20만명+","업무빌딩 밀집 반경 300m","식당·카페 동선 인접"] },
  ];
  strats.forEach((s, i) => {
    const bx = 0.5 + i * 3.1;
    sl.addShape(pres.shapes.RECTANGLE, { x:bx, y:2.9, w:2.95, h:2.3, fill:{color:C.light}, shadow:mkShadow() });
    sl.addShape(pres.shapes.RECTANGLE, { x:bx, y:2.9, w:0.09, h:2.3, fill:{color:C.red} });
    sl.addText(s.title, { x:bx+0.2, y:2.95, w:2.6, h:0.45, fontSize:12, bold:true, color:C.red });
    s.pts.forEach((p, j) => sl.addText("• "+p, { x:bx+0.2, y:3.4+j*0.42, w:2.6, h:0.36, fontSize:11, color:C.dark }));
  });
}

// ── Slide 9 : Bed Town Strategy ───────────────────────────
{
  let sl = pres.addSlide();
  sl.background = { color: C.white };
  sl.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:0.2, h:5.625, fill:{color:C.blue} });
  sl.addText("05  주거 타운 (Bed Town) 입지 전략", { x:0.5, y:0.3, w:9, h:0.55, fontSize:24, bold:true, color:C.navy });
  sl.addShape(pres.shapes.RECTANGLE, { x:0.5, y:0.82, w:1.5, h:0.05, fill:{color:C.blue} });

  sl.addShape(pres.shapes.RECTANGLE, { x:0.5, y:1.05, w:1.6, h:0.5, fill:{color:C.blue}, shadow:mkShadow() });
  sl.addText("BLUE ZONE", { x:0.5, y:1.05, w:1.6, h:0.5, fontSize:13, bold:true, color:C.white, align:"center", valign:"middle" });
  sl.addText("주말 및 평일 야간 집중 운영 스케줄", { x:2.3, y:1.05, w:6, h:0.5, fontSize:18, bold:true, color:C.dark, valign:"middle" });

  sl.addShape(pres.shapes.RECTANGLE, { x:0.5, y:1.7, w:9, h:1.0, fill:{color:"D6EAF8"}, shadow:mkShadow() });
  sl.addText("⏰  골든 아워", { x:0.65, y:1.75, w:3, h:0.35, fontSize:14, bold:true, color:C.blue });
  sl.addText("평일 저녁 19:30 이후 귀가 동선  及  주말 전 시간대 유동인구 집중  |  평일 낮 시간대 다운사이징 권장", { x:0.65, y:2.1, w:8.7, h:0.3, fontSize:12, color:C.dark });
  sl.addText("✓ 평일 낮 운영 인건비 최소화   |   주말 全天 운영으로 체류 시간 극대화", { x:0.65, y:2.4, w:8.7, h:0.28, fontSize:11, bold:true, color:C.blue });

  const strats = [
    { title:"생활 밀착형 '슬리퍼 상권'", pts:["홈리빙 가구·소품","가성비 구독형 밀키트","반려동물 케어용품"] },
    { title:"로컬 커뮤니티 연계 마케팅",  pts:["주차장 확장이용 부스","아파트 단위 공동구매","동네 맘카페 연계 쿠폰"] },
    { title:"입지 선정 기준",            pts:["역 출근 승차인구 15만명+","주거단지 밀집 반경 500m","시장·백화점 인접 역세권"] },
  ];
  strats.forEach((s, i) => {
    const bx = 0.5 + i * 3.1;
    sl.addShape(pres.shapes.RECTANGLE, { x:bx, y:2.9, w:2.95, h:2.3, fill:{color:C.light}, shadow:mkShadow() });
    sl.addShape(pres.shapes.RECTANGLE, { x:bx, y:2.9, w:0.09, h:2.3, fill:{color:C.blue} });
    sl.addText(s.title, { x:bx+0.2, y:2.95, w:2.6, h:0.45, fontSize:12, bold:true, color:C.blue });
    s.pts.forEach((p, j) => sl.addText("• "+p, { x:bx+0.2, y:3.4+j*0.42, w:2.6, h:0.36, fontSize:11, color:C.dark }));
  });
}

// ── Slide 10 : Mixed Zone Strategy ─────────────────────────
{
  let sl = pres.addSlide();
  sl.background = { color: C.white };
  sl.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:0.2, h:5.625, fill:{color:C.gray} });
  sl.addText("05  복합 및 혼합 상권 (Mixed Zone) 입지 전략", { x:0.5, y:0.3, w:9, h:0.55, fontSize:24, bold:true, color:C.navy });
  sl.addShape(pres.shapes.RECTANGLE, { x:0.5, y:0.82, w:1.5, h:0.05, fill:{color:C.gray} });

  sl.addShape(pres.shapes.RECTANGLE, { x:0.5, y:1.05, w:1.6, h:0.5, fill:{color:C.gray}, shadow:mkShadow() });
  sl.addText("GRAY ZONE", { x:0.5, y:1.05, w:1.6, h:0.5, fontSize:13, bold:true, color:C.white, align:"center", valign:"middle" });
  sl.addText("주중·주말 상시 운영 골든 존", { x:2.3, y:1.05, w:6, h:0.5, fontSize:18, bold:true, color:C.dark, valign:"middle" });

  sl.addShape(pres.shapes.RECTANGLE, { x:0.5, y:1.7, w:9, h:1.0, fill:{color:"F2F3F4"}, shadow:mkShadow() });
  sl.addText("⏰  골든 아워", { x:0.65, y:1.75, w:3, h:0.35, fontSize:14, bold:true, color:C.gray });
  sl.addText("주중/주말 관계없이 상시 높은 배후 유동인구 유지  |  체류 시간이 가장 긴 목적형 방문객 비율 높음", { x:0.65, y:2.1, w:8.7, h:0.3, fontSize:12, color:C.dark });
  sl.addText("✓ 여가·트렌드 소비 목적来店 → 인스타그램 바이럴 + 브랜드 콜라보레이션 최적 zone", { x:0.65, y:2.4, w:8.7, h:0.28, fontSize:11, bold:true, color:C.gray });

  const strats = [
    { title:"고감도 비주얼 마케팅",  pts:["인스타그램 바이럴 최적화","포토존·입간판 설치","SNS Influencer 콜라보"] },
    { title:"대형 브랜드 콜라보",    pts:["패션/뷰티 플래그십","카페 브랜드 한정 팝업","엔터테인먼트 체험존"] },
    { title:"입지 선정 기준",        pts:["역 총 승하차 10만명+","20~30대 유동비중 高","상업시설·카페 밀집지"] },
  ];
  strats.forEach((s, i) => {
    const bx = 0.5 + i * 3.1;
    sl.addShape(pres.shapes.RECTANGLE, { x:bx, y:2.9, w:2.95, h:2.3, fill:{color:C.light}, shadow:mkShadow() });
    sl.addShape(pres.shapes.RECTANGLE, { x:bx, y:2.9, w:0.09, h:2.3, fill:{color:C.gray} });
    sl.addText(s.title, { x:bx+0.2, y:2.95, w:2.6, h:0.45, fontSize:12, bold:true, color:C.gray });
    s.pts.forEach((p, j) => sl.addText("• "+p, { x:bx+0.2, y:3.4+j*0.42, w:2.6, h:0.36, fontSize:11, color:C.dark }));
  });
}

// ── Slide 11 : Conclusion ────────────────────────────────────
{
  let sl = pres.addSlide();
  sl.background = { color: C.navy };
  sl.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:10, h:0.08, fill:{color:C.accent} });
  sl.addText("06  결론 및 제언", { x:0.5, y:0.35, w:9, h:0.6, fontSize:28, bold:true, color:C.white });
  sl.addShape(pres.shapes.RECTANGLE, { x:0.5, y:0.92, w:1.5, h:0.05, fill:{color:C.accent} });

  sl.addShape(pres.shapes.RECTANGLE, { x:0.5, y:1.15, w:9, h:1.05, fill:{color:C.dark}, shadow:mkShadow() });
  sl.addText("핵심 발견", { x:0.7, y:1.22, w:2, h:0.35, fontSize:13, bold:true, color:C.accent });
  sl.addText("오프라인 마케팅의 성패는 단순 유동인구 '양(Quantity)'이 아닌,\n'공간의 목적성(Quality)'과 '시간의 방향성(Flow)' 정교한 매칭에 있다.", {
    x:0.7, y:1.52, w:8.6, h:0.65, fontSize:14, color:C.white, italic:true
  });

  const take = [
    { zone:"주거 타운", color:C.blue, msg:"평일 저녁 19:30~ · 주말 운영\n생활 밀착형·가정 소비성향 강함" },
    { zone:"오피스 타운", color:C.red, msg:"평일 점심·퇴근길 집중\n구매력 높지만 피로도 높은 직장인" },
    { zone:"복합 지역",  color:C.gray,msg:"주중·주말 상시\n트렌드·여가 목적의 체류형 방문객" },
  ];
  take.forEach((t, i) => {
    const bx = 0.5 + i * 3.1;
    sl.addShape(pres.shapes.RECTANGLE, { x:bx, y:2.4, w:2.95, h:1.5, fill:{color:C.mid}, shadow:mkShadow() });
    sl.addShape(pres.shapes.RECTANGLE, { x:bx, y:2.4, w:2.95, h:0.07, fill:{color:t.color} });
    sl.addText(t.zone, { x:bx, y:2.5, w:2.95, h:0.4, fontSize:14, bold:true, color:t.color, align:"center" });
    sl.addText(t.msg, { x:bx+0.1, y:2.9, w:2.75, h:0.9, fontSize:11, color:"BDC3C7", align:"center" });
  });

  sl.addShape(pres.shapes.RECTANGLE, { x:0.5, y:4.1, w:9, h:0.95, fill:{color:"1A252F"}, shadow:mkShadow() });
  sl.addText("도출된 데이터 모델 및 클러스터별 운영 프레임워크는\n한정된 마케팅 예산 안에서 ROI 극대화 및 타겟 불일치 리스크 최소화를 위한\n신뢰도 높은 의사결정 시스템 역할을 수행합니다.", {
    x:0.7, y:4.17, w:8.6, h:0.8, fontSize:12, color:"95A5A6", align:"center"
  });
  sl.addShape(pres.shapes.RECTANGLE, { x:0, y:5.2, w:10, h:0.425, fill:{color:C.dark} });
  sl.addText("이어드림스쿨6기  |  데이터 분석 프로젝트", { x:0.5, y:5.2, w:9, h:0.42, fontSize:11, color:"7F8C8D", align:"left", valign:"middle" });
}

// ── Save ──────────────────────────────────────────────────────
pres.writeFile({ fileName: "C:/Users/중진공39/Desktop/서울지하철_팝업스토어입지전략_v2.pptx" })
  .then(() => console.log("✅ Saved: 서울지하철_팝업스토어입지전략_v2.pptx"))
  .catch(e => console.error("Error:", e));