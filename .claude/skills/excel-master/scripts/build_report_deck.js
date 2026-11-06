/* 월간 영업실적 임원 보고 덱(7장) 빌더 — build_monthly_report.py 가 내보낸 *_deck.json 을 읽습니다.
 *   node build_report_deck.js 보고_deck.json "MBE_11월_영업실적_보고_가안.pptx"
 * pptxgenjs 필요. 테마 적용(apply_theme.js)은 pptx 스킬이 있을 때만 수행합니다.
 */
const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");

const [, , jsonPath, outPath] = process.argv;
if (!jsonPath || !outPath) { console.error("사용법: node build_report_deck.js deck.json out.pptx"); process.exit(1); }
const D = JSON.parse(fs.readFileSync(jsonPath, "utf8"));

const THEME = {
  name: "MBE 월간 보고", headFontFace: "맑은 고딕", bodyFontFace: "맑은 고딕",
  colors: { dk1: "1B1F2A", lt1: "FFFFFF", dk2: "14213D", lt2: "F3F5F9", accent1: "1F3864", accent2: "8EA9DB", accent3: "FFC000",
            accent4: "C00000", accent5: "0070C0", accent6: "7F7F7F", hlink: "0070C0", folHlink: "7F7F7F" },
};
const H = THEME.colors;                       // 차트·표 옵션은 hex 필요
const fmt = (n) => Math.round(n).toLocaleString("en-US");
const f1 = (n) => n.toLocaleString("en-US", { minimumFractionDigits: 1, maximumFractionDigits: 1 });
const pct = (v, d = 1) => (v * 100).toFixed(d) + "%";
const spct = (v, d = 1) => (v > 0 ? "+" : "") + (v * 100).toFixed(d) + "%";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";                   // 13.33 x 7.5 in
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = `MBE ${D.month}월 영업실적 보고 (가안)`;
pres.author = D.author;
const C = pres.SchemeColor;

// ---------- 레이아웃 ----------
pres.defineSlideMaster({
  title: "TITLE", background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.2, w: 8.2, h: 1.7, fontSize: 44, bold: true, color: C.background1, align: "left", valign: "top", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "subtitle", type: "body", x: 0.8, y: 4.2, w: 8.2, h: 1.2, fontSize: 20, color: C.accent2, align: "left", valign: "top", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT", background: { color: C.background1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.35, w: 12.1, h: 1.0, fontSize: 30, bold: true, color: C.text2, align: "left", valign: "middle", margin: 0 }, text: "" } },
    { text: { text: `MBE ${D.month}월 영업실적 보고 (가안)`, options: { x: 0.6, y: 7.05, w: 7, h: 0.3, fontSize: 10, color: C.accent6, margin: 0 } } },
  ],
  slideNumber: { x: 12.1, y: 7.05, w: 0.6, h: 0.3, fontSize: 10, color: C.accent6, align: "right" },
});
pres.defineSlideMaster({
  title: "CLOSE", background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.35, w: 12.1, h: 1.0, fontSize: 30, bold: true, color: C.background1, align: "left", valign: "middle", margin: 0 }, text: "" } },
  ],
  slideNumber: { x: 12.1, y: 7.05, w: 0.6, h: 0.3, fontSize: 10, color: C.accent2, align: "right" },
});

const chartText = { catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", legendFontFace: "+mn-lt" };
const quiet = { valGridLine: { color: "E7E6E6", size: 0.75 }, catGridLine: { style: "none" }, catAxisLabelColor: "404040", valAxisLabelColor: "7F7F7F" };

function card(slide, name, x, y, w, h, fill) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: fill || C.background2 }, line: { color: fill || C.background2, width: 0 }, objectName: name });
}
function caption(slide, text, x = 0.6, y = 6.7, w = 12.1) {
  slide.addText(text, { x, y, w, h: 0.3, fontSize: 10, color: C.accent6, margin: 0, isTextBox: true, objectName: "출처 주석" });
}

// ---------- 1. 표지 ----------
pres.addSection({ title: "개요" });
let s = pres.addSlide({ masterName: "TITLE", sectionTitle: "개요" });
s.addText(`MBE ${D.month}월\n영업실적 보고`, { placeholder: "title" });
s.addText(`기준일 ${D.asof}\n가안 · ${D.author}`, { placeholder: "subtitle" });
card(s, "표지 핵심 카드", 9.3, 1.9, 3.4, 3.5, C.accent1);
s.addText(pct(D.progress_rate), { x: 9.3, y: 2.25, w: 3.4, h: 1.3, fontSize: 60, bold: true, color: C.accent3, align: "center", margin: 0, isTextBox: true, objectName: "표지 숫자" });
s.addText("일 목표 진도 대비", { x: 9.3, y: 3.6, w: 3.4, h: 0.45, fontSize: 18, color: C.background1, align: "center", margin: 0, isTextBox: true });
s.addText(`${D.month}월 1~${D.asof.match(/(\d+)일/)[1]}일 매출 ${fmt(D.month_sales)}백만원`, { x: 9.5, y: 4.25, w: 3.0, h: 0.8, fontSize: 14, color: C.accent2, align: "center", margin: 0, isTextBox: true });
s.addNotes("표지. 헤드라인 숫자는 일일 보고서(공식) 기준 진도 대비 달성률입니다.");

// ---------- 2. 핵심 요약 ----------
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "개요" });
s.addText(`${D.month}월 매출은 진도 대비 ${pct(D.progress_rate)}, 마감 예상은 목표의 ${pct(D.fc_rate)}`, { placeholder: "title" });
const stats = [
  { n: pct(D.progress_rate), l: "진도 대비 달성", sub: `실적 ${fmt(D.month_sales)} · 진도 목표 ${fmt(D.progress_target)}`, c: C.text2 },
  { n: spct(D.month_growth), l: "전년 같은 기간 대비", sub: `전년 같은 기간 ${fmt(D.month_ly)}백만원`, c: C.accent4 },
  { n: fmt(D.fc_base), l: "마감 예상 (백만원)", sub: `범위 ${fmt(D.fc_low)}~${fmt(D.fc_high)} · 목표 ${fmt(D.month_target)}`, c: C.text2 },
];
stats.forEach((st, i) => {
  const x = 0.6 + i * 4.2;
  card(s, `핵심 카드 ${i + 1}`, x, 1.7, 3.9, 2.6);
  s.addText(st.n, { x: x + 0.3, y: 1.9, w: 3.3, h: 1.1, fontSize: 48, bold: true, color: st.c, margin: 0, isTextBox: true, objectName: `핵심 숫자 ${i + 1}` });
  s.addText(st.l, { x: x + 0.3, y: 3.0, w: 3.3, h: 0.45, fontSize: 18, bold: true, color: C.text1, margin: 0, isTextBox: true });
  s.addText(st.sub, { x: x + 0.3, y: 3.5, w: 3.3, h: 0.6, fontSize: 12, color: C.accent6, margin: 0, isTextBox: true });
});
s.addText([
  { text: `목표 달성에는 잔여 ${D.days_left}일 일평균 ${fmt(D.need_daily)}백만원이 필요합니다. 계획 일평균 ${fmt(D.plan_daily)}의 ${(D.need_daily / D.plan_daily).toFixed(2)}배이고 전일 실적은 ${fmt(D.y_day_sales)}입니다`, options: { bullet: true, breakLine: true, paraSpaceAfter: 8 } },
  { text: `1~${D.month}월 누계 ${fmt(D.ytd_sales)}백만원으로 누계 목표 ${fmt(D.ytd_target)}의 ${pct(D.ytd_rate)}, 전년 대비 ${spct(D.ytd_growth)}입니다`, options: { bullet: true, breakLine: true, paraSpaceAfter: 8 } },
  { text: D.neg_months >= 2 ? `전년비는 ${D.neg_months}개월 연속 마이너스입니다` : `전년 같은 기간 대비 ${spct(D.month_growth)}입니다`, options: { bullet: true } },
], { x: 0.6, y: 4.65, w: 12.1, h: 1.9, fontSize: 16, color: C.text1, margin: 0, valign: "top", isTextBox: true, objectName: "요약 문장" });
caption(s, "기준: 일일 보고서(ERP). 진도 대비 = 실적 ÷ (월 목표 × 일 목표 진도).");
s.addNotes("결론 먼저. 모든 숫자는 일일 보고서 공식 기준입니다. 마감 예상은 전년 월별 비중과 주중·주말 가중으로 추정한 값입니다.");

// ---------- 3. 유통별 진도 + 월별 추이 ----------
pres.addSection({ title: "실적 분석" });
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "실적 분석" });
const low = [...D.channels].sort((a, b) => a.prog - b.prog)[0];
s.addText(`${low.name}이 진도 대비 ${pct(low.prog)}로 가장 낮고` + (D.neg_months >= 2 ? `, 전년 대비 하락은 ${D.neg_months}개월째입니다` : `, 전체는 전년 같은 기간 대비 ${spct(D.month_growth)}입니다`), { placeholder: "title" });
const chNames = D.channels.map((c) => c.name);
const chColors = D.channels.map((c) => (c.name === low.name ? H.accent4 : H.accent1)).concat([H.accent6]);
s.addText("유통별 진도 달성", { x: 0.6, y: 1.5, w: 5.8, h: 0.4, fontSize: 16, bold: true, color: C.text2, margin: 0, isTextBox: true });
s.addChart(pres.charts.BAR, [{ name: "진도 달성", labels: chNames.concat(["합계"]), values: D.channels.map((c) => c.prog).concat([D.progress_rate]) }], {
  x: 0.6, y: 1.95, w: 5.8, h: 4.5, barDir: "bar", chartColors: chColors, barGapWidthPct: 45, showValue: true, dataLabelPosition: "outEnd",
  dataLabelFormatCode: "0.0%", dataLabelFontSize: 14, dataLabelColor: "1B1F2A", dataLabelFontBold: true,
  catAxisLabelFontSize: 14, valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: 1, catAxisOrientation: "maxMin", showLegend: false,
  valGridLine: { style: "none" }, catGridLine: { style: "none" }, catAxisLabelColor: "1B1F2A", ...chartText,
});
s.addText("MB 월별 매출, 올해와 전년", { x: 6.9, y: 1.5, w: 5.8, h: 0.4, fontSize: 16, bold: true, color: C.text2, margin: 0, isTextBox: true });
const months = D.monthly.map((m) => m.month);
s.addChart([
  { type: pres.charts.BAR, data: [{ name: "26년", labels: months, values: D.monthly.map((m) => m.sales) }], options: { chartColors: [H.accent1], barGapWidthPct: 50 } },
  { type: pres.charts.LINE, data: [{ name: "25년(전년)", labels: months, values: D.monthly.map((m) => m.ly) }], options: { chartColors: ["A6A6A6"], lineSize: 2, lineDataSymbol: "none", lineDash: "dash" } },
], {
  x: 6.9, y: 1.95, w: 5.8, h: 4.5, showLegend: true, legendPos: "b", legendFontSize: 12, catAxisLabelFontSize: 12, valAxisLabelFontSize: 12,
  valAxisLabelFormatCode: "#,##0", ...quiet, ...chartText,
});
caption(s, `${D.month}월은 진행 중이라 전년 같은 기간과 비교합니다. 단위: 백만원. 월별은 일일 보고서(브랜드 합계) 기준.`);
s.addNotes("왼쪽은 일 목표 진도 대비 달성률(채널 전체 실적 ÷ 채널 진도 목표). 오른쪽 월별 매출의 11월은 1~5일만 집계된 진행 중인 값입니다.");

// ---------- 4. 증감 분해 ----------
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "실적 분석" });
const wf = [...D.waterfall].sort((a, b) => a.delta - b.delta);
const total = wf.reduce((a, b) => a + b.delta, 0);
const worst = wf[0];
s.addText(`전년 대비 누계는 ${fmt(total)}백만원 줄었고 ${worst.name}이 ${fmt(worst.delta)}로 가장 큽니다`, { placeholder: "title" });
s.addText("유통별 누계 증감액 (백만원)", { x: 0.6, y: 1.5, w: 7.4, h: 0.4, fontSize: 16, bold: true, color: C.text2, margin: 0, isTextBox: true });
s.addChart(pres.charts.BAR, [{ name: "증감액", labels: wf.map((w) => (w.name === "기타" ? "기타(미분류)" : w.name)), values: wf.map((w) => w.delta) }], {
  x: 0.6, y: 1.95, w: 7.4, h: 4.5, barDir: "bar", chartColors: [H.accent5], invertedColors: [H.accent4], barGapWidthPct: 45, showValue: true,
  dataLabelPosition: "outEnd", dataLabelFormatCode: "#,##0;-#,##0", dataLabelFontSize: 14, dataLabelColor: "1B1F2A", dataLabelFontBold: true,
  catAxisLabelFontSize: 14, valAxisHidden: true, valAxisMinVal: Math.floor(wf[0].delta / 1000 - 0.8) * 1000, valAxisMaxVal: 1200, catAxisOrientation: "maxMin", showLegend: false,
  valGridLine: { style: "none" }, catGridLine: { style: "none" }, catAxisLabelColor: "1B1F2A", catAxisLabelPos: "low", ...chartText,
});
const share = (worst.delta / total) * 100;
const plus = wf.filter((w) => w.delta > 0);
[
  { n: fmt(total), l: "누계 증감 합계", sub: "매장 합계 기준 (전년 동기 대비)", c: C.accent4 },
  { n: `${worst.name} ${share.toFixed(0)}%`, l: "감소에서 차지하는 비중", sub: `${fmt(worst.delta)}백만원`, c: C.text2 },
].forEach((st, i) => {
  const y = 1.7 + i * 2.35;
  card(s, `증감 카드 ${i + 1}`, 8.5, y, 4.2, 2.1);
  s.addText(st.n, { x: 8.8, y: y + 0.2, w: 3.6, h: 0.9, fontSize: 36, bold: true, color: st.c, margin: 0, isTextBox: true });
  s.addText(st.l, { x: 8.8, y: y + 1.1, w: 3.6, h: 0.4, fontSize: 16, bold: true, color: C.text1, margin: 0, isTextBox: true });
  s.addText(st.sub, { x: 8.8, y: y + 1.5, w: 3.6, h: 0.4, fontSize: 12, color: C.accent6, margin: 0, isTextBox: true });
});
caption(s, `매장 행 합계 기준입니다. 공식 누계 전년 동기보다 ${fmt(Math.abs(D.diff_ytd_py))}백만원 작아 공식 누계 전년비(${spct(D.ytd_growth)})와 다릅니다.` + (plus.length ? ` 기타(미분류)는 유통에 배정되지 않은 매출입니다.` : ""));
s.addNotes("전년 대비 증감을 유통별로 나눈 값입니다. 마스터 DB 매장 합계 기준이며, 공식 보고서 누계와는 전년 동기에서 차이가 있어 주석에 밝혔습니다.");

// ---------- 5. 점검 ----------
pres.addSection({ title: "점검과 실행" });
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "점검과 실행" });
const need = D.judges["확인 필요"] || 0, warn = D.judges["주의"] || 0;
s.addText(`확인 필요 ${need}건 중 우선 점검 매장 ${D.priority.length}곳과 신상 할인 ${D.new_dc_diff > 0 ? "확대" : "점검"}가 핵심입니다`, { placeholder: "title" });
s.addText("우선 점검 매장 (진도 70% 미만, 부족액 큰 순)", { x: 0.6, y: 1.5, w: 7.6, h: 0.4, fontSize: 16, bold: true, color: C.text2, margin: 0, isTextBox: true });
const hdr = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: H.accent1 }, align: "center", fontSize: 14 } });
const rowsP = [[hdr("순위"), hdr("매장"), hdr("유통"), hdr("진도 달성"), hdr("본사 담당")]].concat(
  D.priority.map((p) => [
    { text: String(p.rank), options: { align: "center" } }, { text: p.name, options: { bold: true } }, { text: p.channel, options: { align: "center" } },
    { text: pct(p.prog), options: { align: "right", color: H.accent4, bold: true } }, { text: p.owner, options: { align: "center" } }]));
s.addTable(rowsP, { x: 0.6, y: 1.95, w: 7.6, colW: [0.8, 3.0, 1.2, 1.4, 1.2], rowH: 0.62, fontSize: 14, color: "1B1F2A", valign: "middle", border: { type: "solid", pt: 0.75, color: "D9D9D9" } });
const risks = [
  { n: `${need}건`, l: "확인 필요 항목", sub: `주의 ${warn}건 · 일일 보고서 자동 판정` },
  { n: `${pct(D.new_dc)}`, l: "신상 할인율", sub: `전년보다 ${D.new_dc_diff > 0 ? "+" : ""}${D.new_dc_diff.toFixed(1)}%p ${D.new_dc_diff > 0 ? "확대" : "변동"}` },
  { n: spct(D.worst_channel.growth), l: `${D.worst_channel.name} 전년 같은 기간 대비`, sub: D.worst_channel.growth < 0 ? "하락 폭이 유통 중 가장 큼" : "증감이 유통 중 가장 낮음" },
];
risks.forEach((r, i) => {
  const y = 1.55 + i * 1.7;
  card(s, `리스크 카드 ${i + 1}`, 8.7, y, 4.0, 1.5);
  s.addText(r.n, { x: 8.95, y: y + 0.12, w: 3.5, h: 0.65, fontSize: 30, bold: true, color: i === 0 ? C.accent4 : C.text2, margin: 0, isTextBox: true });
  s.addText(r.l, { x: 8.95, y: y + 0.77, w: 3.5, h: 0.35, fontSize: 14, bold: true, color: C.text1, margin: 0, isTextBox: true });
  s.addText(r.sub, { x: 8.95, y: y + 1.1, w: 3.5, h: 0.3, fontSize: 11, color: C.accent6, margin: 0, isTextBox: true });
});
caption(s, "본사 담당은 매장 정보의 본사 담당자 기준입니다. 원인은 확인 전이므로 가설로만 관리합니다.");
s.addNotes("우선 점검 매장은 일일 보고서 점검 기준(진도 70% 미만, 월 목표 50백만 이상, 부족액 큰 순 5곳)과 같습니다.");

// ---------- 6. 다음 주 액션 ----------
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "점검과 실행" });
const A = D.actions;
const dept = D.low_channel;
s.addText(`다음 주 액션 ${A.length}건을 제안합니다 (담당·기한은 확정 전)`, { placeholder: "title" });
const short = A.map((a) => a.basis.split('. ')[0].replace(/\(진도 70% 미만\)/, '').trim());
const ownerShort = (o) => (o.includes(",") ? o.split(",")[0].trim() + " 외" : o);
const rowsA = [[hdr("과제"), hdr("근거"), hdr("담당"), hdr("기한")]].concat(A.map((a, i) => [
  { text: a.task, options: { bold: true } }, { text: short[i] }, { text: ownerShort(a.owner), options: { align: "center" } }, { text: a.due.replace("~ 매일", "~"), options: { align: "center" } }]));
s.addTable(rowsA, { x: 0.6, y: 1.55, w: 12.1, colW: [3.2, 5.6, 1.6, 1.7], rowH: 0.68, fontSize: 14, color: "1B1F2A", valign: "middle", border: { type: "solid", pt: 0.75, color: "D9D9D9" } });
caption(s, "담당 \"외\"는 우선 점검 매장의 본사 담당 여러 명을 뜻합니다. 상세 근거는 엑셀 '이슈·액션' 시트에 있습니다.");
s.addNotes("담당과 기한은 제안입니다. 확정 후 엑셀 이슈·액션 시트와 함께 갱신하세요.");

// ---------- 7. 제안과 결정 요청 ----------
s = pres.addSlide({ masterName: "CLOSE", sectionTitle: "점검과 실행" });
s.addText("프로모션 2안과 결정이 필요한 3가지", { placeholder: "title" });
D.promos.forEach((p, i) => {
  const x = 0.6 + i * 3.95;
  card(s, `프로모션 ${p.id}`, x, 1.6, 3.75, 4.6, H.accent1);
  s.addText(`${p.id}안`, { x: x + 0.3, y: 1.8, w: 1.2, h: 0.5, fontSize: 22, bold: true, color: C.accent3, margin: 0, isTextBox: true });
  s.addText(p.title, { x: x + 0.3, y: 2.3, w: 3.15, h: 0.85, fontSize: 18, bold: true, color: C.background1, margin: 0, valign: "top", isTextBox: true });
  s.addText(p.basis, { x: x + 0.3, y: 3.2, w: 3.15, h: 1.75, fontSize: 14, color: "FFFFFF", margin: 0, valign: "top", isTextBox: true });
  s.addText([
    { text: `대상  ${p.target}`, options: { breakLine: true } },
    { text: `기간  ${p.period}`, options: { breakLine: true } },
    { text: `측정  ${p.metric}` },
  ], { x: x + 0.3, y: 5.0, w: 3.15, h: 1.05, fontSize: 12, color: C.accent2, margin: 0, valign: "top", isTextBox: true, objectName: `프로모션 ${p.id} 조건` });
});
s.addText("결정 요청", { x: 8.6, y: 1.6, w: 4.1, h: 0.5, fontSize: 22, bold: true, color: C.accent3, margin: 0, isTextBox: true });
s.addText([
  { text: `다음 주 액션 ${A.length}건의 담당과 기한 확정`, options: { bullet: { type: "number" }, breakLine: true, paraSpaceAfter: 14 } },
  { text: "프로모션 A안·B안 진행 여부", options: { bullet: { type: "number" }, breakLine: true, paraSpaceAfter: 14 } },
  { text: `하락 원인(${[...new Set([D.worst_channel.name, D.low_channel.name])].join("·")}) 확인 담당 지정`, options: { bullet: { type: "number" } } },
], { x: 8.6, y: 2.25, w: 4.1, h: 3.4, fontSize: 18, color: "FFFFFF", margin: 0, valign: "top", isTextBox: true, objectName: "결정 요청 목록" });
s.addText("가안입니다. 원인은 가설이며 효과는 실행 후 측정합니다.", { x: 8.6, y: 5.8, w: 4.1, h: 0.5, fontSize: 12, color: C.accent2, margin: 0, isTextBox: true });
s.addNotes("프로모션은 가설 기반 제안입니다. 승인되면 대상 매장과 기간을 확정하고 행사 전후 일 매출·객단가로 효과를 측정합니다.");

(async () => {
  await pres.writeFile({ fileName: outPath });
  const hits = require("child_process").spawnSync("sh", ["-c", "ls /root/.claude/skills/synced/*/pptx/scripts/apply_theme.js 2>/dev/null | head -1"], { encoding: "utf8" }).stdout.trim();
  if (hits) {
    const { applyTheme } = require(hits);
    await applyTheme(outPath, THEME);
    console.log("테마 적용 완료");
  } else {
    console.log("apply_theme.js 없음: 테마 색 적용 생략(스킬 설치 후 다시 실행)");
  }
  console.log("완료:", outPath);
})();
