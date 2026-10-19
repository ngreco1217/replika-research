// Builds the 48" x 36" Vanderbilt research poster (editable .pptx).
// Usage: NODE_PATH=<dir with pptxgenjs> node build_poster.js <out.pptx>
const path = require("path");
const pptxgen = require("pptxgenjs");

const OUT = process.argv[2] || path.join(__dirname, "Greco_NHIS_Poster.pptx");
const LOGO = path.join(__dirname, "vanderbilt_v.png");

// ---------- palette & type ----------
const BLACK = "1C1C1C";
const GOLD = "CFAE70";
const WHITE = "FFFFFF";
const TINT = "F6F0E3"; // light gold tint for callout boxes
const GRAY = "8A8A8A";
const DARKGRAY = "4A4A4A";
const FONT = "Arial";

const pres = new pptxgen();
pres.defineLayout({ name: "POSTER_48x36", width: 48, height: 36 });
pres.layout = "POSTER_48x36";
pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
pres.title = "Moderate Physical Activity and Self-Reported Depression Outcomes (NHIS 2010-2023)";
pres.author = "Nicolas Greco";

const s = pres.addSlide();
s.background = { color: WHITE };

// ---------- helpers ----------
const M = 0.75; // outer margin
const PAD = 0.4; // inner panel padding
const HEAD_H = 0.9; // panel heading bar height

function text(t, o) {
  s.addText(t, Object.assign({ isTextBox: true, fontFace: FONT, color: BLACK, margin: 0, valign: "top" }, o));
}

function panel(x, y, w, h, title) {
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: WHITE }, line: { color: GOLD, width: 2.5 } });
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h: HEAD_H, fill: { color: BLACK }, line: { color: BLACK, width: 0.5 } });
  text(title.toUpperCase(), {
    x, y, w, h: HEAD_H, fontSize: 38, bold: true, color: GOLD, align: "center", valign: "middle", charSpacing: 1,
  });
  return { cx: x + PAD, cw: w - 2 * PAD, top: y + HEAD_H + 0.35 };
}

function numCircle(x, y, d, n, fs) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: GOLD }, line: { color: BLACK, width: 1.5 } });
  text(String(n), { x, y, w: d, h: d, fontSize: fs || 30, bold: true, align: "center", valign: "middle" });
}

// bullet list: items are arrays of runs or strings
function bullets(items, o) {
  const runs = [];
  items.forEach((it, i) => {
    const parts = typeof it === "string" ? [{ text: it }] : it;
    parts.forEach((p, j) => {
      const opt = Object.assign({}, p.options || {});
      if (j === 0) opt.bullet = { indent: 28 };
      if (j === parts.length - 1 && i < items.length - 1) opt.breakLine = true;
      runs.push({ text: p.text, options: opt });
    });
  });
  text(runs, Object.assign({ fontSize: 24, paraSpaceAfter: 10 }, o));
}
const B = (t) => ({ text: t, options: { bold: true } });
const R = (t) => ({ text: t });

// ---------- header band ----------
const HDR_H = 6.95;
s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 48, h: HDR_H, fill: { color: BLACK }, line: { color: BLACK, width: 0 } });
// logo (aspect 900:632)
s.addImage({ path: LOGO, x: M + 0.35, y: 0.75, w: 4.4, h: 4.4 * 632 / 900, altText: "Vanderbilt University V logo" });
text("VANDERBILT", { x: M, y: 4.05, w: 5.1, h: 0.6, fontSize: 38, color: WHITE, align: "center", fontFace: "Times New Roman" });
text("UNIVERSITY", { x: M, y: 4.65, w: 5.1, h: 0.6, fontSize: 38, color: WHITE, align: "center", fontFace: "Times New Roman" });

const TX = 6.6, TW = 47.25 - TX;
text(
  "The Association Between Moderate Physical Activity and Self-Reported Depression Outcomes: Evidence from the National Health Interview Survey, 2010–2023",
  { x: TX, y: 0.75, w: TW, h: 3.65, fontSize: 85, bold: true, color: GOLD, align: "center", valign: "middle", lineSpacingMultiple: 0.9 }
);
text([{ text: "Nicolas Greco", options: { underline: { style: "sng" } } }], {
  x: TX, y: 4.45, w: TW, h: 0.9, fontSize: 56, bold: true, color: WHITE, align: "center", valign: "middle",
});
text("Departments of Economics & Psychology, Vanderbilt University, Nashville, TN", {
  x: TX + 4, y: 5.4, w: TW - 8, h: 0.5, fontSize: 30, color: WHITE, align: "center", valign: "middle",
});
text("Independent study conducted at Pace University (Pforzheimer Honors College), New York, NY.", {
  x: TX + 4, y: 5.9, w: TW - 8, h: 0.5, fontSize: 26, italic: true, color: "D9D9D9", align: "center", valign: "middle",
});
// in-press badge (bottom right of header)
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 40.95, y: 5.25, w: 6.3, h: 1.25, rectRadius: 0.2, fill: { color: GOLD }, line: { color: GOLD, width: 0 } });
text([
  { text: "Paper in press –", options: { bold: true, fontSize: 20, breakLine: true } },
  { text: "University of Michigan Undergraduate", options: { fontSize: 18, breakLine: true } },
  { text: "Journal of Public Health", options: { fontSize: 18 } },
], { x: 41.15, y: 5.3, w: 5.9, h: 1.15, align: "center", valign: "middle", objectName: "In-press badge" });

// ---------- columns ----------
const Y0 = 7.5, Y1 = 36 - M;
const GAP = 0.6;
const LX = M, LW = 13.5;
const MX = LX + LW + GAP, MW = 18.3;
const RX = MX + MW + GAP, RW = 13.5;

// ===== LEFT: Background =====
let p = panel(LX, Y0, LW, 12.05, "Background");
let y = p.top;
const stats = [
  ["8.3%", [B("of U.S. adults"), R(" (≈21 million) had a major depressive episode in 2021 (NIMH)")]],
  ["280M", [B("people worldwide"), R(" live with depression (WHO)")]],
  ["17.2%", [B("of U.S. adults aged 18–25"), R(" had past-year depression in 2020, up from 10.3% in 2015 (Goodwin et al., 2022)")]],
];
stats.forEach(([num, label]) => {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: p.cx, y, w: p.cw, h: 1.2, rectRadius: 0.12, fill: { color: BLACK }, line: { color: BLACK, width: 0 } });
  text(num, { x: p.cx + 0.2, y, w: 3.6, h: 1.2, fontSize: 60, bold: true, color: GOLD, align: "center", valign: "middle" });
  text(label.map((r) => ({ text: r.text, options: Object.assign({ color: WHITE }, r.options) })), {
    x: p.cx + 4.0, y: y + 0.05, w: p.cw - 4.2, h: 1.1, fontSize: 24, valign: "middle",
  });
  y += 1.32;
});
y += 0.15;
// barriers vs physical activity
const cardW = (p.cw - 0.3) / 2;
[["Treatment barriers", ["High costs", "Long waitlists", "Provider shortages"], "F2F2F2"],
 ["Physical activity", ["Low-cost", "No prescription", "Flexible, anywhere"], TINT]].forEach(([title, items, fill], i) => {
  const cx = p.cx + i * (cardW + 0.3);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: cx, y, w: cardW, h: 2.25, rectRadius: 0.12, fill: { color: fill }, line: { color: i ? GOLD : "BFBFBF", width: 1.5 } });
  text(title, { x: cx + 0.25, y: y + 0.18, w: cardW - 0.5, h: 0.5, fontSize: 24, bold: true });
  bullets(items, { x: cx + 0.25, y: y + 0.7, w: cardW - 0.5, h: 1.45, paraSpaceAfter: 2 });
});
y += 2.5;
text("Prior evidence", { x: p.cx, y, w: p.cw, h: 0.5, fontSize: 26, bold: true });
y += 0.55;
bullets([
  [R("Physical activity: "), B("17% lower odds"), R(" of developing depression (Kandola et al., 2019)")],
  [R("Half the recommended activity: "), B("18% lower risk"), R("; full amount: "), B("25%"), R(" (Pearce et al., 2022)")],
], { x: p.cx, y, w: p.cw, h: 1.6 });
y += 1.75;
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: p.cx, y, w: p.cw, h: 1.35, rectRadius: 0.12, fill: { color: TINT }, line: { color: GOLD, width: 2 } });
text([B("Research gap: "), R("few nationally representative studies compare how often people exercise across several dimensions of depression.")], {
  x: p.cx + 0.25, y: y + 0.1, w: p.cw - 0.5, h: 1.15, fontSize: 24, valign: "middle",
});
const bgEnd = y + 1.35 + 0.35;
console.log("Background content ends", bgEnd.toFixed(2), "panel ends", (Y0 + 12.05).toFixed(2));

// ===== LEFT: Aims =====
const AY = Y0 + 12.05 + 0.5;
p = panel(LX, AY, LW, 6.8, "Aims");
y = p.top;
text([B("Question: "), R("Is how often U.S. adults do moderate physical activity linked to their self-reported depression and general health?")],
  { x: p.cx, y, w: p.cw, h: 1.15, fontSize: 24 });
y += 1.2;
text([B("Hypothesis: "), R("More frequent moderate activity is associated with better depression-related outcomes.")],
  { x: p.cx, y, w: p.cw, h: 0.85, fontSize: 24 });
y += 0.95;
// conceptual model
const mY = y;
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: p.cx, y: mY, w: 4.3, h: 2.34, rectRadius: 0.15, fill: { color: BLACK }, line: { color: BLACK, width: 0 } });
text([
  { text: "Moderate physical activity", options: { bold: true, color: GOLD, fontSize: 24, breakLine: true } },
  { text: "None · Low · Medium · High", options: { color: WHITE, fontSize: 20 } },
], { x: p.cx + 0.2, y: mY, w: 3.9, h: 2.34, align: "center", valign: "middle" });
const outX = p.cx + 6.15, outW = p.cw - 6.15;
["General health", "Depression frequency", "Depression level"].forEach((o, i) => {
  const oy = mY + i * 0.82;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: outX, y: oy, w: outW, h: 0.7, rectRadius: 0.1, fill: { color: WHITE }, line: { color: GOLD, width: 2.5 } });
  text(o, { x: outX, y: oy, w: outW, h: 0.7, fontSize: 22, bold: true, align: "center", valign: "middle" });
});
s.addShape(pres.shapes.LINE, { x: p.cx + 4.4, y: mY + 1.17, w: 1.65, h: 0, line: { color: BLACK, width: 5, endArrowType: "triangle" } });
text("linked to?", { x: p.cx + 4.3, y: mY + 0.62, w: 1.85, h: 0.45, fontSize: 18, italic: true, align: "center" });
// controls
s.addShape(pres.shapes.LINE, { x: p.cx + 5.2, y: mY + 1.3, w: 0, h: 1.25, line: { color: GRAY, width: 2.5, dashType: "dash" } });
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: p.cx + 2.6, y: mY + 2.55, w: 5.2, h: 0.6, rectRadius: 0.1, fill: { color: "F2F2F2" }, line: { color: GRAY, width: 1.5, dashType: "dash" } });
text("Controls: demographic & socioeconomic", { x: p.cx + 2.6, y: mY + 2.55, w: 5.2, h: 0.6, fontSize: 18, align: "center", valign: "middle", color: DARKGRAY });
console.log("Aims model ends", (mY + 3.15).toFixed(2), "panel ends", (AY + 6.8).toFixed(2));

// ===== LEFT: Methods =====
const MYp = AY + 6.8 + 0.5;
p = panel(LX, MYp, LW, Y1 - MYp, "Methods");
y = p.top;
const tx = p.cx + 1.15, tw = p.cw - 1.15;
const chip = (cx, cy, cw, label, fill, bold) => {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: cx, y: cy, w: cw, h: 0.62, rectRadius: 0.1, fill: { color: fill }, line: { color: fill === WHITE ? GOLD : fill, width: 1.5 } });
  text(label, { x: cx, y: cy, w: cw, h: 0.62, fontSize: 18, bold: !!bold, align: "center", valign: "middle", color: fill === BLACK ? GOLD : BLACK });
};
const steps = [
  { h: 1.6, title: "Data", body: "IPUMS NHIS 2010–2023 · N = 332,521 U.S. adults · self-reported, nationally representative" },
  { h: 1.45, title: "Exposure: moderate activity, times/week", chips: [["None 0", WHITE], ["Low 1–3", WHITE], ["Medium 4–6", WHITE], ["High 7+", WHITE]] },
  { h: 1.45, title: "Outcomes (3 separate models)", chips: [["General health", TINT], ["Depression frequency", TINT], ["Depression level", TINT]] },
  { h: 1.55, title: "OLS regression vs. no activity", chips: [["(1) Activity only", WHITE], ["(2) + Demographic", WHITE], ["(3) + Socioeconomic", BLACK]] },
];
const stepTop = y;
steps.forEach((st, i) => {
  if (i < steps.length - 1) {
    s.addShape(pres.shapes.LINE, { x: p.cx + 0.4, y: y + 0.8, w: 0, h: st.h - 0.8, line: { color: GOLD, width: 3 } });
  }
  numCircle(p.cx, y, 0.8, i + 1);
  text(st.title, { x: tx, y: y + 0.08, w: tw, h: 0.5, fontSize: 24, bold: true });
  if (st.body) text(st.body, { x: tx, y: y + 0.6, w: tw, h: 0.85, fontSize: 24 });
  if (st.chips) {
    const n = st.chips.length, g = 0.15;
    const widths = n === 4 ? [2.35, 2.35, 2.65, 2.25] : n === 3 && i === 2 ? [3.15, 4.05, 3.25] : [3.3, 3.55, 3.6];
    let cx = tx;
    st.chips.forEach(([lab, fill], k) => { chip(cx, y + 0.68, widths[k], lab, fill, fill === BLACK); cx += widths[k] + g; });
  }
  y += st.h;
});
text("★ Specification (3), with all controls, is preferred.", { x: tx, y: y + 0.0, w: tw, h: 0.45, fontSize: 20, italic: true, color: DARKGRAY });
console.log("Methods ends", (y + 0.45).toFixed(2), "panel ends", Y1.toFixed(2));

// ===== MIDDLE: Results =====
p = panel(MX, Y0, MW, Y1 - Y0, "Results");
y = p.top;
// headline callout
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: p.cx, y, w: p.cw, h: 3.0, rectRadius: 0.2, fill: { color: BLACK }, line: { color: GOLD, width: 3 } });
text([
  { text: "−2.92", options: { fontSize: 100, bold: true, color: GOLD, breakLine: true } },
  { text: "percentage points", options: { fontSize: 24, color: WHITE } },
], { x: p.cx + 0.3, y, w: 5.6, h: 3.0, align: "center", valign: "middle", lineSpacingMultiple: 0.9 });
text([
  { text: "Adults active 4–6 times a week were 2.92 percentage points less likely to report frequent depressive symptoms than inactive adults.", options: { fontSize: 28, bold: true, color: WHITE, breakLine: true } },
  { text: "95% CI: −4.4 to −1.4 · Specification 3", options: { fontSize: 22, color: GOLD } },
], { x: p.cx + 6.2, y: y + 0.2, w: p.cw - 6.6, h: 2.6, valign: "middle", paraSpaceAfter: 8 });
y += 3.4;

// ---- Figure 1: dot-and-whisker (drawn as editable shapes) ----
text("Figure 1. Change in each outcome vs. no activity, by activity level (Spec 3)", { x: p.cx, y, w: p.cw, h: 0.5, fontSize: 24, bold: true });
y += 0.6;
// legend
const legY = y + 0.08;
s.addShape(pres.shapes.OVAL, { x: p.cx + 0.1, y: legY + 0.06, w: 0.3, h: 0.3, fill: { color: BLACK }, line: { color: BLACK, width: 1.5 } });
text("Significant (95% CI excludes 0)", { x: p.cx + 0.55, y: legY, w: 5.4, h: 0.42, fontSize: 18, valign: "middle" });
s.addShape(pres.shapes.OVAL, { x: p.cx + 6.2, y: legY + 0.06, w: 0.3, h: 0.3, fill: { color: WHITE }, line: { color: GRAY, width: 2.5 } });
text("Not significant", { x: p.cx + 6.65, y: legY, w: 2.6, h: 0.42, fontSize: 18, valign: "middle" });
s.addShape(pres.shapes.OVAL, { x: p.cx + 9.5, y: legY + 0.06, w: 0.3, h: 0.3, fill: { color: GOLD }, line: { color: BLACK, width: 1.5 } });
text("Headline finding", { x: p.cx + 9.95, y: legY, w: 3.2, h: 0.42, fontSize: 18, valign: "middle" });
y += 0.75;

// Spec 3 estimates, in percentage points (Tables 2-4 x 100). Negative = less likely than inactive adults.
const FIG = [
  { title: "A. Poor general health", range: [-40, 0], ticks: [-40, -30, -20, -10, 0],
    rows: [[-22.7, -23.8, -21.6, true, false, "−22.7"], [-34.2, -35.7, -32.7, true, false, "−34.2"], [-35.9, -37.9, -33.9, true, false, "−35.9"]] },
  { title: "B. Frequent depression", range: [-5, 1], ticks: [-4, -2, 0],
    rows: [[-0.454, -1.5, 0.6, false, false, "−0.454"], [-2.92, -4.4, -1.4, true, true, "−2.92"], [-2.50, -4.5, -0.5, true, false, "−2.50"]] },
  { title: "C. Depression level", range: [-4, 3], ticks: [-4, -2, 0, 2],
    rows: [[-0.324, -1.9, 1.3, false, false, "−0.324"], [-1.45, -3.7, 0.8, false, false, "−1.45"], [-0.397, -3.3, 2.5, false, false, "−0.397"]] },
];
const fmt = (v) => (v < 0 ? "−" : "") + String(Math.abs(v));
const labW = 3.3;
const plotX0 = p.cx + labW + 0.2, plotW = p.cw - labW - 0.2;
const pg = 0.5, pw = (plotW - 2 * pg) / 3;
const titleY = y, plotTop = y + 0.6, rowH = 1.75, plotH = rowH * 3;
// y-axis title (rotated) and row labels
text("Activity level (times/week)", { x: p.cx - 2.0 + 0.25, y: plotTop + plotH / 2 - 0.25, w: 4.4, h: 0.5, fontSize: 20, bold: true, align: "center", valign: "middle", rotate: 270 });
["Low (1–3)", "Medium (4–6)", "High (7+)"].forEach((lab, r) => {
  text(lab, { x: p.cx + 0.55, y: plotTop + r * rowH + rowH / 2 - 0.25, w: labW - 0.45, h: 0.5, fontSize: 22, bold: true, align: "right", valign: "middle" });
});
FIG.forEach((f, k) => {
  const px = plotX0 + k * (pw + pg);
  const inner = 0.3; // keep whiskers off the frame
  const X = (v) => px + inner + ((v - f.range[0]) / (f.range[1] - f.range[0])) * (pw - 2 * inner);
  text(f.title, { x: px, y: titleY, w: pw, h: 0.5, fontSize: 22, bold: true, align: "center" });
  s.addShape(pres.shapes.RECTANGLE, { x: px, y: plotTop, w: pw, h: plotH, fill: { color: "FAFAFA" }, line: { color: "BFBFBF", width: 1 } });
  f.ticks.forEach((t) => {
    s.addShape(pres.shapes.LINE, { x: X(t), y: plotTop, w: 0, h: plotH, line: t === 0 ? { color: BLACK, width: 2, dashType: "dash" } : { color: "E0E0E0", width: 1 } });
    s.addShape(pres.shapes.LINE, { x: X(t), y: plotTop + plotH, w: 0, h: 0.12, line: { color: DARKGRAY, width: 1.5 } });
    text(fmt(t), { x: X(t) - 0.6, y: plotTop + plotH + 0.15, w: 1.2, h: 0.4, fontSize: 18, align: "center" });
  });
  f.rows.forEach(([est, lo, hi, sig, star, lab], r) => {
    const cy = plotTop + r * rowH + rowH / 2 + 0.15;
    const col = sig ? BLACK : GRAY;
    s.addShape(pres.shapes.LINE, { x: X(lo), y: cy, w: X(hi) - X(lo), h: 0, line: { color: col, width: 4 } });
    [lo, hi].forEach((v) => s.addShape(pres.shapes.LINE, { x: X(v), y: cy - 0.13, w: 0, h: 0.26, line: { color: col, width: 4 } }));
    const d = star ? 0.42 : 0.34;
    s.addShape(pres.shapes.OVAL, { x: X(est) - d / 2, y: cy - d / 2, w: d, h: d,
      fill: { color: star ? GOLD : sig ? BLACK : WHITE }, line: { color: sig ? BLACK : GRAY, width: sig ? 1.5 : 3 } });
    text(lab, { x: X(est) - 0.6, y: cy - 0.72, w: 1.2, h: 0.38, fontSize: 18, bold: sig, align: "center", valign: "middle", color: sig ? BLACK : DARKGRAY, fill: { color: "FAFAFA" } });
  });
});
y = plotTop + plotH + 0.6;
text("Difference in likelihood vs. inactive adults (percentage points)", { x: plotX0, y, w: plotW, h: 0.45, fontSize: 20, bold: true, align: "center" });
y += 0.5;
text("Dots = estimates; bars = 95% confidence intervals. Values below 0 mean less likely than adults with no moderate activity. Each panel has its own scale.",
  { x: p.cx, y, w: p.cw, h: 0.75, fontSize: 18, italic: true, color: DARKGRAY });
y += 1.05;

// ---- Figure 2: sample size (native chart) ----
text("Figure 2. Sample size by moderate activity level (N = 332,521; % of sample)", { x: p.cx, y, w: p.cw, h: 0.5, fontSize: 24, bold: true });
y += 0.55;
const fig2H = 5.8;
s.addChart(pres.charts.BAR, [{
  name: "Respondents",
  labels: ["None (0) – 37.5%", "Low (1–3) – 44.0%", "Medium (4–6) – 12.2%", "High (7+) – 6.3%"],
  values: [124773, 146241, 40660, 20847],
}], {
  x: p.cx, y, w: p.cw, h: fig2H, barDir: "bar", catAxisOrientation: "maxMin",
  chartColors: [BLACK], barGapWidthPct: 45,
  showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "#,##0", dataLabelFontSize: 20, dataLabelFontBold: true, dataLabelColor: BLACK, dataLabelFontFace: FONT,
  catAxisLabelFontSize: 20, catAxisLabelColor: BLACK, catAxisLabelFontFace: FONT,
  valAxisLabelFontSize: 18, valAxisLabelColor: DARKGRAY, valAxisLabelFontFace: FONT, valAxisLabelFormatCode: "#,##0",
  valAxisMaxVal: 175000, valAxisMinVal: 0, valAxisMajorUnit: 25000,
  showCatAxisTitle: true, catAxisTitle: "Activity level (times/week)", catAxisTitleFontSize: 20, catAxisTitleColor: BLACK, catAxisTitleFontFace: FONT,
  showValAxisTitle: true, valAxisTitle: "Number of respondents", valAxisTitleFontSize: 20, valAxisTitleColor: BLACK, valAxisTitleFontFace: FONT,
  valGridLine: { color: "E0E0E0", size: 1 }, catGridLine: { style: "none" },
  showLegend: false, showTitle: false,
});
y += fig2H + 0.35;

// ---- Table 1: compact Spec 3 results ----
text("Table 1. Specification 3 estimates, percentage points [95% CI]", { x: p.cx, y, w: p.cw, h: 0.5, fontSize: 24, bold: true });
y += 0.6;
const hdr = (t) => ({ text: t, options: { bold: true, color: GOLD, fill: { color: BLACK }, align: "center", valign: "middle", fontSize: 20 } });
const cell = (est, ci, sig) => ({
  text: [
    { text: est + (sig ? "*" : ""), options: { bold: sig, fontSize: 22, breakLine: true } },
    { text: ci, options: { fontSize: 18, color: DARKGRAY } },
  ],
  options: { align: "center", valign: "middle" },
});
const rowLab = (t) => ({ text: t, options: { bold: true, fontSize: 20, valign: "middle" } });
const tbl = [
  [hdr("Activity level"), hdr("General health (poor)"), hdr("Depression frequency"), hdr("Depression level")],
  [rowLab("Low (1–3)"), cell("−22.7", "[−23.8, −21.6]", true), cell("−0.454", "[−1.5, 0.6]", false), cell("−0.324", "[−1.9, 1.3]", false)],
  [rowLab("Medium (4–6)"), cell("−34.2", "[−35.7, −32.7]", true), cell("−2.92", "[−4.4, −1.4]", true), cell("−1.45", "[−3.7, 0.8]", false)],
  [rowLab("High (7+)"), cell("−35.9", "[−37.9, −33.9]", true), cell("−2.50", "[−4.5, −0.5]", true), cell("−0.397", "[−3.3, 2.5]", false)],
  [rowLab("Observations"), { text: "161,344", options: { align: "center", fontSize: 20 } }, { text: "160,974", options: { align: "center", fontSize: 20 } }, { text: "66,089", options: { align: "center", fontSize: 20 } }],
  [rowLab("R²"), { text: "0.204", options: { align: "center", fontSize: 20 } }, { text: "0.252", options: { align: "center", fontSize: 20 } }, { text: "0.038", options: { align: "center", fontSize: 20 } }],
];
const colW = [3.5, (p.cw - 3.5) / 3, (p.cw - 3.5) / 3, (p.cw - 3.5) / 3];
s.addTable(tbl, {
  x: p.cx, y, w: p.cw, colW, rowH: [0.65, 1.0, 1.0, 1.0, 0.6, 0.6],
  fontFace: FONT, color: BLACK, border: { type: "solid", pt: 1, color: "BFBFBF" }, margin: 0.06, valign: "middle",
});
y += 0.65 + 1.0 * 3 + 0.6 * 2 + 0.15;
text("* Statistically significant (95% CI excludes 0). Reference group: no moderate activity.", { x: p.cx, y, w: p.cw, h: 0.45, fontSize: 18, italic: true, color: DARKGRAY });
console.log("Results ends", (y + 0.45).toFixed(2), "panel ends", Y1.toFixed(2));

// ===== RIGHT: Discussion =====
const DH = 9.75;
p = panel(RX, Y0, RW, DH, "Discussion");
y = p.top;
bullets([
  [B("Modest but meaningful: "), R("4–6 sessions a week was linked to fewer frequent depressive symptoms")],
  [B("Better self-rated health "), R("at every activity level (−22.7 to −35.9 pp)")],
  [B("Fits prior work "), R("linking activity to lower depression risk (Kandola et al., 2019; Pearce et al., 2022)")],
  [B("Plausible pathways: "), R("less stress, better sleep, greater well-being")],
  [B("No link with depression level: "), R("all CIs include 0 (R² = 0.038)")],
  [B("Socioeconomic status matters: "), R("welfare receipt consistently predicted worse outcomes (Alzahrani et al., 2024)")],
], { x: p.cx, y, w: p.cw, h: 5.0, paraSpaceAfter: 12 });
y += 5.15;
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: p.cx, y, w: p.cw, h: 2.95, rectRadius: 0.12, fill: { color: "F2F2F2" }, line: { color: GRAY, width: 1.5 } });
text("Limitations", { x: p.cx + 0.25, y: y + 0.15, w: p.cw - 0.5, h: 0.5, fontSize: 26, bold: true });
bullets([
  "Self-reported data: social desirability & extreme response bias",
  "Unobserved confounders (e.g., sleep, genetics)",
  "Cross-sectional data show association, not cause",
  "Income measured only in categories",
], { x: p.cx + 0.25, y: y + 0.7, w: p.cw - 0.5, h: 2.15, fontSize: 24, paraSpaceAfter: 2 });
console.log("Discussion ends", (y + 2.95).toFixed(2), "panel ends", (Y0 + DH).toFixed(2));

// ===== RIGHT: Conclusions =====
const CY = Y0 + DH + 0.5, CH = 8.45;
p = panel(RX, CY, RW, CH, "Conclusions");
y = p.top;
const msgs = [
  "Being moderately active 4–6 times a week was linked to a lower chance of feeling depressed often.",
  "Any regular moderate activity – even 1–3 times a week – was linked to better self-rated health.",
  "Exercise is not a cure: it was not linked to how severe depression feels, so it should complement, not replace, treatment.",
];
const msgH = [1.3, 1.3, 1.7];
msgs.forEach((m, i) => {
  numCircle(p.cx, y + 0.05, 0.8, i + 1);
  text(m, { x: p.cx + 1.1, y, w: p.cw - 1.1, h: msgH[i], fontSize: 24, bold: i === 0 });
  y += msgH[i] + 0.1;
});
y += 0.05;
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: p.cx, y, w: p.cw, h: 2.35, rectRadius: 0.12, fill: { color: TINT }, line: { color: GOLD, width: 2 } });
text("Future directions", { x: p.cx + 0.25, y: y + 0.15, w: p.cw - 0.5, h: 0.5, fontSize: 26, bold: true });
bullets([
  "Test activity as a treatment for people clinically diagnosed with depression",
  "Track real-time, individual-level improvement",
], { x: p.cx + 0.25, y: y + 0.7, w: p.cw - 0.5, h: 1.55, fontSize: 24, paraSpaceAfter: 2 });
console.log("Conclusions ends", (y + 2.35).toFixed(2), "panel ends", (CY + CH).toFixed(2));

// ===== RIGHT: Acknowledgements =====
const AKY = CY + CH + 0.5;
p = panel(RX, AKY, RW, Y1 - AKY, "Acknowledgements");
y = p.top - 0.1;
text("I thank Professor Osborne for his helpful feedback on the manuscript, and my mom and dad for their constant support and encouragement.",
  { x: p.cx, y, w: p.cw, h: 1.3, fontSize: 24 });
y += 1.4;
text([B("Funding: "), R("This independent study received no external funding.")], { x: p.cx, y, w: p.cw, h: 0.4, fontSize: 20 });
y += 0.45;
text([B("Data: "), R("IPUMS NHIS v8.1 (Blewett et al., 2025).")], { x: p.cx, y, w: p.cw, h: 0.4, fontSize: 20 });
y += 0.55;
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: p.cx, y, w: p.cw, h: 2.05, rectRadius: 0.12, fill: { color: TINT }, line: { color: GOLD, width: 2 } });
text([
  { text: "Author Contribution Statement", options: { bold: true, fontSize: 24, breakLine: true } },
  { text: "NG (sole author) conceived and designed the study, prepared the data, performed all analyses, interpreted the results, wrote the manuscript, and designed this poster.", options: { fontSize: 22 } },
], { x: p.cx + 0.2, y: y + 0.1, w: p.cw - 0.4, h: 1.85, valign: "middle" });
y += 2.25;
text("Selected references", { x: p.cx, y, w: p.cw, h: 0.35, fontSize: 18, bold: true });
y += 0.38;
const refs = [
  "Blewett, L. A., et al. (2025). IPUMS Health Surveys: NHIS (Version 8.1). doi.org/10.18128/D071.V8.1",
  "Goodwin, R. D., et al. (2022). Am. J. Prev. Med., 63(5), 726–733.",
  "Kandola, A., et al. (2019). Neurosci. Biobehav. Rev., 107, 525–539.",
  "Pearce, M., et al. (2022). JAMA Psychiatry, 79(6), 550–559.",
  "Alzahrani, H., et al. (2024). J. Adv. Trends Med. Res., 862–869.",
  "NIMH (n.d.). Major Depression. · WHO (2025). Depressive disorder (depression).",
];
text(refs.map((r, i) => ({ text: r, options: { breakLine: i < refs.length - 1 } })), {
  x: p.cx, y, w: p.cw, h: 2.2, fontSize: 16, color: DARKGRAY, paraSpaceAfter: 2,
});
console.log("Ack ends", (y + 2.2).toFixed(2), "panel ends", Y1.toFixed(2));

pres.writeFile({ fileName: OUT }).then(() => console.log("wrote", OUT));
