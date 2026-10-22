// Builds the 48" x 36" Vanderbilt research poster, v2 (editable .pptx).
// Usage: NODE_PATH=<dir with pptxgenjs, react, react-dom, react-icons, sharp> node build_poster_v2.js <out.pptx>
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");

const OUT = process.argv[2] || path.join(__dirname, "Greco_NHIS_Poster_48x36_v2.pptx");
const LOGO = path.join(__dirname, "vanderbilt_v.png");

// ---------- palette ----------
const BLACK = "1C1C1C";
const GOLD = "CFAE70";
const WHITE = "FFFFFF";
const TINT = "F7F1E4";
const SOFT = "F4F4F4";
const GRAY = "8A8A8A";
const DARKGRAY = "4A4A4A";
const LINEGRAY = "D9D9D9";
// activity ramp: none -> low -> medium -> high (used everywhere activity level appears)
const RAMP = { none: "BDBDBD", low: "E8D6A8", med: "CFAE70", high: "8C6A2F" };
const RAMP_TXT = { none: BLACK, low: BLACK, med: BLACK, high: WHITE };
const FONT = "Arial";

// ---------- assets ----------
async function icon(name, color, px = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(fa[name], { color: "#" + color, size: px }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// Header background: charcoal, soft gold glow behind the title, and a heartbeat line along the bottom.
async function headerBackground() {
  const W = 4800, H = 720, base = 672;
  const beat = (x0, a) => // one heartbeat complex starting at x0, amplitude a (px)
    `L ${x0} ${base} q 25 -${a * 0.18} 50 0 L ${x0 + 80} ${base} l 18 ${a * 0.15} l 26 -${a} l 30 ${a * 1.35} l 20 -${a * 0.5} ` +
    `L ${x0 + 220} ${base} q 45 -${a * 0.3} 90 0`;
  let d = `M 0 ${base} `;
  [60, 300, 4160, 4430].forEach((x, i) => { d += beat(x, i % 2 ? 70 : 110) + " "; });
  d += `L ${W} ${base}`;
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${W * 2}" height="${H * 2}" viewBox="0 0 ${W} ${H}">
    <defs>
      <radialGradient id="g" cx="0.5" cy="0.38" r="0.5" gradientTransform="translate(0.5 0.38) scale(1 0.36) translate(-0.5 -0.38)">
        <stop offset="0" stop-color="#CFAE70" stop-opacity="0.22"/>
        <stop offset="1" stop-color="#CFAE70" stop-opacity="0"/>
      </radialGradient>
      <linearGradient id="l" x1="0" x2="1" y1="0" y2="0">
        <stop offset="0" stop-color="#CFAE70" stop-opacity="0.95"/>
        <stop offset="0.3" stop-color="#CFAE70" stop-opacity="0.35"/>
        <stop offset="0.7" stop-color="#CFAE70" stop-opacity="0.35"/>
        <stop offset="1" stop-color="#CFAE70" stop-opacity="0.95"/>
      </linearGradient>
    </defs>
    <rect width="${W}" height="${H}" fill="#1C1C1C"/>
    <rect width="${W}" height="${H}" fill="url(#g)"/>
    <path d="${d}" fill="none" stroke="url(#l)" stroke-width="6" stroke-linejoin="round" stroke-linecap="round"/>
  </svg>`;
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

(async () => {
  const I = {};
  const goldIcons = ["FaBrain", "FaBullseye", "FaFlask", "FaChartBar", "FaComments", "FaLightbulb", "FaHandsHelping", "FaBookOpen"];
  const blackIcons = ["FaDollarSign", "FaHourglassHalf", "FaUserMd", "FaRunning", "FaBan", "FaMapMarkerAlt", "FaDatabase", "FaWalking",
    "FaClipboardList", "FaLayerGroup", "FaHeartbeat", "FaCalendarAlt", "FaTachometerAlt", "FaBalanceScale", "FaMinusCircle", "FaCommentDots",
    "FaQuestion", "FaLink", "FaCoins", "FaBookOpen", "FaCalendarCheck", "FaHandHoldingHeart", "FaSearch", "FaMicroscope", "FaPuzzlePiece"];
  for (const n of goldIcons) I[n + "_g"] = await icon(n, GOLD);
  for (const n of blackIcons) I[n] = await icon(n, BLACK);
  I.FaRocket_g = await icon("FaRocket", GOLD);
  const HEADER_BG = await headerBackground();

  const pres = new pptxgen();
  pres.defineLayout({ name: "POSTER_48x36", width: 48, height: 36 });
  pres.layout = "POSTER_48x36";
  pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
  pres.title = "Moderate Physical Activity and Self-Reported Depression Outcomes (NHIS 2010-2023)";
  pres.author = "Nicolas Greco";
  const s = pres.addSlide();
  s.background = { color: WHITE };

  // ---------- helpers ----------
  const M = 0.75;
  const text = (t, o) => s.addText(t, Object.assign({ isTextBox: true, fontFace: FONT, color: BLACK, margin: 0, valign: "top" }, o));
  const B = (t, extra) => ({ text: t, options: Object.assign({ bold: true }, extra || {}) });
  const R = (t, extra) => ({ text: t, options: Object.assign({}, extra || {}) });
  const SUP = (t, extra) => ({ text: t, options: Object.assign({ superscript: true }, extra || {}) }); // numbered citation
  const shadow = () => ({ type: "outer", blur: 10, offset: 3, angle: 90, color: "000000", opacity: 0.16 });
  const rrect = (x, y, w, h, o) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, Object.assign({ x, y, w, h, rectRadius: 0.15 }, o));
  const circ = (x, y, d, o) => s.addShape(pres.shapes.OVAL, Object.assign({ x, y, w: d, h: d }, o));
  const img = (data, x, y, w, h, alt) => s.addImage({ data, x, y, w, h: h || w, altText: alt || "" });
  // icon centered in a filled circle
  const iconBadge = (data, x, y, d, fill, line) => {
    circ(x, y, d, { fill: { color: fill }, line: { color: line || fill, width: 1.5 } });
    img(data, x + d * 0.22, y + d * 0.22, d * 0.56);
  };
  const PAD = 0.45;
  // card with floating black heading pill + gold icon; returns content box
  function card(x, y, w, h, title, iconData) {
    rrect(x, y, w, h, { rectRadius: 0.3, fill: { color: WHITE }, line: { color: GOLD, width: 2 }, shadow: shadow() });
    rrect(x + 0.25, y + 0.25, w - 0.5, 0.95, { rectRadius: 0.475, fill: { color: BLACK }, line: { color: BLACK, width: 0 } });
    circ(x + 0.33, y + 0.3, 0.84, { fill: { color: "2E2E2E" }, line: { color: GOLD, width: 2 } });
    img(iconData, x + 0.51, y + 0.48, 0.48);
    text(title.toUpperCase(), { x: x + 1.3, y: y + 0.25, w: w - 2.6, h: 0.95, fontSize: 40, bold: true, color: GOLD, align: "center", valign: "middle", charSpacing: 2 });
    return { cx: x + PAD, cw: w - 2 * PAD, top: y + 1.2 + 0.3, bottom: y + h - 0.3 };
  }
  const log = (name, end, bottom) => console.log(`${name.padEnd(14)} content ends ${end.toFixed(2)}  limit ${bottom.toFixed(2)}${end > bottom + 0.001 ? "  <-- OVERFLOW" : ""}`);

  // =====================================================================
  // HEADER
  // =====================================================================
  const HDR_H = 7.25;
  s.addImage({ data: HEADER_BG, x: 0, y: 0, w: 48, h: HDR_H, altText: "Header background with heartbeat line" });
  img("image/png;base64," + require("fs").readFileSync(LOGO).toString("base64"), M + 0.3, 0.8, 4.3, 4.3 * 632 / 900, "Vanderbilt University V logo");
  text("VANDERBILT", { x: M, y: 3.95, w: 4.9, h: 0.6, fontSize: 38, color: WHITE, align: "center", fontFace: "Times New Roman", charSpacing: 1 });
  text("UNIVERSITY", { x: M, y: 4.55, w: 4.9, h: 0.6, fontSize: 38, color: WHITE, align: "center", fontFace: "Times New Roman", charSpacing: 1 });

  const TX = 6.2, TW = 34.6;
  text("The Association Between Moderate Physical Activity and Self-Reported Depression Outcomes", {
    x: TX, y: 0.75, w: TW, h: 2.75, fontSize: 85, bold: true, color: GOLD, align: "center", valign: "middle", lineSpacingMultiple: 0.92,
  });
  text("Evidence from the National Health Interview Survey, 2010–2023", {
    x: TX, y: 3.55, w: TW, h: 0.8, fontSize: 48, color: WHITE, align: "center", valign: "middle",
  });
  text([{ text: "Nicolas Greco", options: { underline: { style: "sng" } } }], {
    x: TX, y: 4.45, w: TW, h: 0.95, fontSize: 56, bold: true, color: WHITE, align: "center", valign: "middle",
  });
  text("Departments of Economics & Psychology, Vanderbilt University, Nashville, TN", {
    x: TX, y: 5.45, w: TW, h: 0.45, fontSize: 28, color: WHITE, align: "center", valign: "middle",
  });
  text("Independent study conducted at Pace University (Pforzheimer Honors College), New York, NY.", {
    x: TX, y: 5.92, w: TW, h: 0.45, fontSize: 24, italic: true, color: "D0D0D0", align: "center", valign: "middle",
  });
  // in-press badge
  rrect(41.25, 1.0, 6.0, 3.3, { rectRadius: 0.3, fill: { color: "262626" }, line: { color: GOLD, width: 3 } });
  circ(43.65, 1.25, 1.2, { fill: { color: BLACK }, line: { color: GOLD, width: 2 } });
  img(I.FaBookOpen_g, 43.92, 1.52, 0.66);
  text([
    { text: "Paper in press –", options: { bold: true, fontSize: 26, color: GOLD, breakLine: true } },
    { text: "University of Michigan Undergraduate Journal of Public Health", options: { fontSize: 20, color: WHITE } },
  ], { x: 41.5, y: 2.55, w: 5.5, h: 1.6, align: "center", valign: "middle", objectName: "In-press badge" });

  // =====================================================================
  // COLUMNS
  // =====================================================================
  const Y0 = 7.75, Y1 = 36 - M, GAP = 0.55;
  const LX = M, LW = 13.4, MX = LX + LW + GAP, MW = 18.6, RX = MX + MW + GAP, RW = 13.4;

  // ----------------------------- BACKGROUND -----------------------------
  const BG_H = 11.0;
  let c = card(LX, Y0, LW, BG_H, "Background", I.FaBrain_g);
  let y = c.top;
  const tileW = (c.cw - 0.4) / 3;
  [
    ["21M", [R("U.S. adults had a major depressive episode in 2021 (8.3%)"), SUP("1")]],
    ["280M", [R("people live with depression worldwide"), SUP("2")]],
    ["17.2%", [R("of U.S. adults aged 18–25 had depression in 2020, up from 10.3% in 2015"), SUP("3")]],
  ].forEach(([num, lab], i) => {
    const tx = c.cx + i * (tileW + 0.2);
    rrect(tx, y, tileW, 2.45, { rectRadius: 0.2, fill: { color: BLACK }, line: { color: BLACK, width: 0 }, shadow: shadow() });
    text(num, { x: tx, y: y + 0.08, w: tileW, h: 1.0, fontSize: 64, bold: true, color: GOLD, align: "center", valign: "middle" });
    text(lab.map((q) => ({ text: q.text, options: Object.assign({ color: WHITE }, q.options) })), { x: tx + 0.15, y: y + 1.08, w: tileW - 0.3, h: 1.3, fontSize: 18, color: WHITE, align: "center", valign: "top" });
  });
  y += 2.6;
  // barriers vs physical activity
  const halfW = (c.cw - 0.3) / 2;
  [
    ["Treatment barriers", SOFT, "BFBFBF", [["FaDollarSign", "High costs"], ["FaHourglassHalf", "Long waitlists"], ["FaUserMd", "Provider shortages"]]],
    ["Physical activity", TINT, GOLD, [["FaCoins", "Low-cost"], ["FaBan", "No prescription"], ["FaMapMarkerAlt", "Flexible, anywhere"]]],
  ].forEach(([title, fill, line, items], k) => {
    const bx = c.cx + k * (halfW + 0.3);
    rrect(bx, y, halfW, 2.4, { rectRadius: 0.2, fill: { color: fill }, line: { color: line, width: 2 } });
    text(title, { x: bx + 0.25, y: y + 0.12, w: halfW - 0.5, h: 0.5, fontSize: 24, bold: true });
    items.forEach(([ic, lab], j) => {
      const iy = y + 0.66 + j * 0.57;
      iconBadge(I[ic], bx + 0.25, iy, 0.52, WHITE, k ? GOLD : "BFBFBF");
      text(lab, { x: bx + 0.95, y: iy, w: halfW - 1.1, h: 0.52, fontSize: 24, valign: "middle" });
    });
  });
  // arrow between the two cards
  circ(c.cx + halfW - 0.2, y + 0.85, 0.7, { fill: { color: BLACK }, line: { color: GOLD, width: 2 } });
  text("→", { x: c.cx + halfW - 0.2, y: y + 0.85, w: 0.7, h: 0.7, fontSize: 28, bold: true, color: GOLD, align: "center", valign: "middle" });
  y += 2.6;
  // prior evidence mini bar chart
  text("Prior evidence: activity is linked to lower depression risk", { x: c.cx, y, w: c.cw, h: 0.45, fontSize: 24, bold: true });
  y += 0.55;
  const pe = [
    ["Any activity", "4", 17, "17% lower odds"],
    ["Half the recommended dose", "5", 18, "18% lower risk"],
    ["Full recommended dose", "5", 25, "25% lower risk"],
  ];
  const peLabW = 5.0, peBarX = c.cx + peLabW + 0.15, peBarW = c.cw - peLabW - 0.15 - 2.6;
  pe.forEach(([lab, src, v, vl], i) => {
    const py = y + i * 0.7;
    text([B(lab), SUP(src, { bold: true })], { x: c.cx, y: py + 0.04, w: peLabW, h: 0.6, fontSize: 20, align: "right", valign: "middle" });
    rrect(peBarX, py + 0.12, peBarW * v / 25, 0.45, { rectRadius: 0.22, fill: { color: i === 2 ? RAMP.high : i === 1 ? RAMP.med : RAMP.low }, line: { color: BLACK, width: 0.75 } });
    text(vl, { x: peBarX + peBarW * v / 25 + 0.12, y: py + 0.05, w: 2.6, h: 0.6, fontSize: 20, bold: true, valign: "middle" });
  });
  y += 2.15;
  rrect(c.cx, y, c.cw, 1.3, { rectRadius: 0.2, fill: { color: BLACK }, line: { color: GOLD, width: 2.5 } });
  iconBadge(I.FaSearch, c.cx + 0.25, y + 0.27, 0.75, GOLD);
  text([B("The gap: ", { color: GOLD }), R("few nationally representative studies compare how often people exercise across several dimensions of depression.", { color: WHITE })],
    { x: c.cx + 1.2, y: y + 0.08, w: c.cw - 1.4, h: 1.14, fontSize: 24, valign: "middle" });
  log("Background", y + 1.3, c.bottom);

  // ----------------------------- AIMS -----------------------------
  const AY = Y0 + BG_H + 0.5, AH = 6.7;
  c = card(LX, AY, LW, AH, "Aims", I.FaBullseye_g);
  y = c.top;
  const qh = (letter, label, body, h) => {
    circ(c.cx, y + 0.02, 0.62, { fill: { color: GOLD }, line: { color: BLACK, width: 1.5 } });
    text(letter, { x: c.cx, y: y + 0.02, w: 0.62, h: 0.62, fontSize: 28, bold: true, align: "center", valign: "middle" });
    text([B(label + " "), R(body)], { x: c.cx + 0.85, y, w: c.cw - 0.85, h, fontSize: 24 });
    y += h + 0.1;
  };
  qh("Q", "Question:", "Is how often U.S. adults do moderate physical activity linked to their self-reported depression and general health?", 0.85);
  qh("H", "Hypothesis:", "Since activity can lower stress and aid sleep, more frequent activity is linked to better depression outcomes.", 0.85);
  // conceptual model
  const mY = y, exW = 4.3, exH = 2.3;
  rrect(c.cx, mY, exW, exH, { rectRadius: 0.2, fill: { color: BLACK }, line: { color: BLACK, width: 0 }, shadow: shadow() });
  text("Moderate physical activity", { x: c.cx + 0.15, y: mY + 0.15, w: exW - 0.3, h: 0.85, fontSize: 22, bold: true, color: GOLD, align: "center", valign: "middle" });
  // rising ramp bars = frequency
  ["none", "low", "med", "high"].forEach((k, i) => {
    const bw = 0.62, bh = 0.22 + i * 0.22, bx = c.cx + 0.55 + i * 0.83;
    s.addShape(pres.shapes.RECTANGLE, { x: bx, y: mY + 1.9 - bh, w: bw, h: bh, fill: { color: RAMP[k] }, line: { color: RAMP[k], width: 0 } });
  });
  text("frequency", { x: c.cx, y: mY + 1.9, w: exW, h: 0.35, fontSize: 16, italic: true, color: "D0D0D0", align: "center", valign: "middle" });
  const outX = c.cx + exW + 1.65, outW = c.cw - exW - 1.65;
  [["FaHeartbeat", "General health"], ["FaCalendarAlt", "Depression frequency"], ["FaTachometerAlt", "Depression level"]].forEach(([ic, lab], i) => {
    const oy = mY + i * 0.8;
    rrect(outX, oy, outW, 0.7, { rectRadius: 0.35, fill: { color: WHITE }, line: { color: GOLD, width: 2.5 } });
    img(I[ic], outX + 0.2, oy + 0.15, 0.4);
    text(lab, { x: outX + 0.7, y: oy, w: outW - 0.8, h: 0.7, fontSize: 20, bold: true, valign: "middle" });
    // fan arrows
    const sx = c.cx + exW + 0.1, sy = mY + exH / 2, ex = outX - 0.08, ey = oy + 0.35;
    s.addShape(pres.shapes.LINE, { x: sx, y: Math.min(sy, ey), w: ex - sx, h: Math.abs(ey - sy) || 0.001, flipV: ey < sy,
      line: { color: BLACK, width: 3, endArrowType: "triangle" } });
  });
  rrect(outX, mY + 2.45, outW, 0.55, { rectRadius: 0.3, fill: { color: SOFT }, line: { color: GRAY, width: 1.5, dashType: "dash" } });
  text("+ demographic & socioeconomic controls", { x: outX, y: mY + 2.45, w: outW, h: 0.55, fontSize: 16, italic: true, color: DARKGRAY, align: "center", valign: "middle" });
  log("Aims", mY + 3.0, c.bottom);

  // ----------------------------- METHODS -----------------------------
  const MYp = AY + AH + 0.5;
  c = card(LX, MYp, LW, Y1 - MYp, "Methods", I.FaFlask_g);
  y = c.top;
  const tx = c.cx + 1.15, tw = c.cw - 1.15;
  const steps = [
    { h: 1.42, ic: "FaDatabase", title: [B("Data: IPUMS NHIS"), SUP("6", { bold: true }), B(", 2010–2023")] },
    { h: 1.38, ic: "FaWalking", title: "Exposure: moderate activity, times/week" },
    { h: 1.38, ic: "FaClipboardList", title: "Outcomes: 3 self-reported measures" },
    { h: 2.75, ic: "FaLayerGroup", title: "Model: OLS regression vs. no activity" },
  ];
  steps.forEach((st, i) => {
    if (i < steps.length - 1) s.addShape(pres.shapes.LINE, { x: c.cx + 0.45, y: y + 0.9, w: 0, h: st.h - 0.9, line: { color: GOLD, width: 4 } });
    iconBadge(I[st.ic], c.cx, y, 0.9, GOLD, BLACK);
    text(st.title, { x: tx, y: y + 0.05, w: tw, h: 0.5, fontSize: 24, bold: true });
    y += st.h;
  });
  // step 1 detail
  let sy = c.top + 0.62;
  text([B("N = 332,521"), R(" U.S. adults · self-reported · nationally representative")], { x: tx, y: sy, w: tw, h: 0.85, fontSize: 24 });
  // step 2 chips (ramp)
  sy = c.top + 1.42 + 0.62;
  const chipW = (tw - 0.45) / 4;
  [["none", "None", "0"], ["low", "Low", "1–3"], ["med", "Medium", "4–6"], ["high", "High", "7+"]].forEach(([k, a, b], i) => {
    rrect(tx + i * (chipW + 0.15), sy, chipW, 0.62, { rectRadius: 0.31, fill: { color: RAMP[k] }, line: { color: k === "low" ? GOLD : RAMP[k], width: 1 } });
    text([B(a + " "), R(b)], { x: tx + i * (chipW + 0.15), y: sy, w: chipW, h: 0.62, fontSize: 18, color: RAMP_TXT[k], align: "center", valign: "middle" });
  });
  // step 3 chips
  sy = c.top + 1.42 + 1.38 + 0.62;
  [["FaHeartbeat", "General health", 3.75], ["FaCalendarAlt", "Dep. frequency", 3.9], ["FaTachometerAlt", "Dep. level", 3.4]].forEach(([ic, lab, w], i, arr) => {
    const ox = tx + arr.slice(0, i).reduce((a, q) => a + q[2] + 0.15, 0);
    rrect(ox, sy, w, 0.62, { rectRadius: 0.31, fill: { color: WHITE }, line: { color: GOLD, width: 2 } });
    img(I[ic], ox + 0.2, sy + 0.13, 0.36);
    text(lab, { x: ox + 0.7, y: sy, w: w - 0.75, h: 0.62, fontSize: 18, bold: true, valign: "middle" });
  });
  // step 4 nested specifications (boxes inside boxes = nested models)
  sy = c.top + 1.42 + 1.38 + 1.38 + 0.6;
  const nx = tx, nw = tw;
  rrect(nx, sy, nw, 1.72, { rectRadius: 0.2, fill: { color: BLACK }, line: { color: GOLD, width: 2.5 } });
  text([B("(3) + Socioeconomic", { color: GOLD }), R(": marital status, income, poverty ratio, meds", { color: WHITE }), R("  ★ preferred", { color: GOLD, italic: true })], { x: nx + 0.2, y: sy + 0.05, w: nw - 0.4, h: 0.42, fontSize: 18 });
  rrect(nx + 0.25, sy + 0.52, nw - 0.5, 1.0, { rectRadius: 0.16, fill: { color: "5A4A2A" }, line: { color: GOLD, width: 1.5 } });
  text([B("(2) + Demographic", { color: WHITE }), R(": age, sex, race", { color: WHITE })], { x: nx + 0.45, y: sy + 0.56, w: nw - 0.9, h: 0.4, fontSize: 18 });
  rrect(nx + 0.5, sy + 0.97, nw - 1.0, 0.45, { rectRadius: 0.12, fill: { color: GOLD }, line: { color: GOLD, width: 1 } });
  text([B("(1) Activity only")], { x: nx + 0.7, y: sy + 0.97, w: nw - 1.4, h: 0.45, fontSize: 18, valign: "middle" });
  text("How to read: each estimate is the percentage-point difference vs. inactive adults.", { x: tx, y: sy + 1.8, w: tw, h: 0.4, fontSize: 18, italic: true, color: DARKGRAY });
  log("Methods", sy + 2.2, c.bottom);

  // =====================================================================
  // RESULTS
  // =====================================================================
  c = card(MX, Y0, MW, Y1 - Y0, "Results", I.FaChartBar_g);
  y = c.top;
  // headline callout
  rrect(c.cx, y, c.cw, 3.1, { rectRadius: 0.3, fill: { color: BLACK }, line: { color: GOLD, width: 3 }, shadow: shadow() });
  circ(c.cx + 0.35, y + 0.3, 1.0, { fill: { color: "2E2E2E" }, line: { color: GOLD, width: 2 } });
  img(I.FaRocket_g, c.cx + 0.6, y + 0.55, 0.5);
  text([
    { text: "−2.92", options: { fontSize: 115, bold: true, color: GOLD, breakLine: true } },
    { text: "percentage points", options: { fontSize: 24, color: WHITE } },
  ], { x: c.cx + 1.0, y, w: 5.9, h: 3.1, align: "center", valign: "middle", lineSpacingMultiple: 0.88 });
  text([
    { text: "KEY FINDING", options: { fontSize: 20, bold: true, color: GOLD, charSpacing: 3, breakLine: true } },
    { text: "Adults active 4–6 times a week were 2.92 percentage points less likely to report frequent depressive symptoms than inactive adults.", options: { fontSize: 28, bold: true, color: WHITE, breakLine: true } },
    { text: "95% CI: −4.4 to −1.4 · Specification 3", options: { fontSize: 22, color: GOLD } },
  ], { x: c.cx + 6.9, y: y + 0.2, w: c.cw - 7.3, h: 2.7, valign: "middle", paraSpaceAfter: 8 });
  y += 3.75;

  // ---- Figure 1: dot-and-whisker ----
  text([B("Figure 1. "), R("Change in each outcome vs. no activity, by activity level (Spec 3, 95% CI)")], { x: c.cx, y, w: c.cw, h: 0.5, fontSize: 24 });
  y += 0.62;
  // legend
  const lg = (x, fill, line, label, lw) => {
    circ(x, y + 0.07, 0.32, { fill: { color: fill }, line: { color: line, width: 2 } });
    text(label, { x: x + 0.42, y, w: lw, h: 0.46, fontSize: 18, valign: "middle" });
  };
  lg(c.cx + 3.6, RAMP.med, BLACK, "Significant (filled, colored by activity level)", 6.2);
  lg(c.cx + 10.6, WHITE, GRAY, "Not significant (hollow)", 3.6);
  y += 0.62;
  const FIG = [
    { title: "A. Poor general health", ic: "FaHeartbeat", range: [-40, 0], ticks: [-40, -30, -20, -10, 0],
      rows: [[-22.7, -23.8, -21.6, true, "−22.7"], [-34.2, -35.7, -32.7, true, "−34.2"], [-35.9, -37.9, -33.9, true, "−35.9"]] },
    { title: "B. Frequent depression", ic: "FaCalendarAlt", range: [-5, 1], ticks: [-4, -2, 0],
      rows: [[-0.454, -1.5, 0.6, false, "−0.454"], [-2.92, -4.4, -1.4, true, "−2.92", true], [-2.50, -4.5, -0.5, true, "−2.50"]] },
    { title: "C. Depression level", ic: "FaTachometerAlt", range: [-4, 3], ticks: [-4, -2, 0, 2],
      rows: [[-0.324, -1.9, 1.3, false, "−0.324"], [-1.45, -3.7, 0.8, false, "−1.45"], [-0.397, -3.3, 2.5, false, "−0.397"]] },
  ];
  const fmt = (v) => (v < 0 ? "−" : "") + Math.abs(v);
  const rowKeys = ["low", "med", "high"], rowLabs = [["Low", "1–3×/wk"], ["Medium", "4–6×/wk"], ["High", "7+×/wk"]];
  const labW = 2.75, plotX0 = c.cx + 0.55 + labW + 0.2, plotW = c.cx + c.cw - plotX0, pg = 0.4, pw = (plotW - 2 * pg) / 3;
  const titleY = y, plotTop = y + 0.62, rowH = 1.72, plotH = rowH * 3;
  text("Activity level", { x: c.cx - 2.2 + 0.25, y: plotTop + plotH / 2 - 0.25, w: 4.4, h: 0.5, fontSize: 20, bold: true, align: "center", valign: "middle", rotate: 270 });
  rowKeys.forEach((k, r) => {
    const ly = plotTop + r * rowH + rowH / 2 - 0.42;
    rrect(c.cx + 0.6, ly, labW, 0.84, { rectRadius: 0.42, fill: { color: RAMP[k] }, line: { color: k === "low" ? GOLD : RAMP[k], width: 1 } });
    text([B(rowLabs[r][0], { breakLine: true, fontSize: 22 }), R(rowLabs[r][1], { fontSize: 16 })], { x: c.cx + 0.6, y: ly, w: labW, h: 0.84, color: RAMP_TXT[k], align: "center", valign: "middle" });
  });
  FIG.forEach((f, k) => {
    const px = plotX0 + k * (pw + pg), inner = 0.3;
    const X = (v) => px + inner + ((v - f.range[0]) / (f.range[1] - f.range[0])) * (pw - 2 * inner);
    img(I[f.ic], px + 0.05, titleY + 0.06, 0.38);
    text(f.title, { x: px + 0.5, y: titleY, w: pw - 0.5, h: 0.5, fontSize: 21, bold: true, valign: "middle" });
    rrect(px, plotTop, pw, plotH, { rectRadius: 0.12, fill: { color: "FAFAFA" }, line: { color: LINEGRAY, width: 1 } });
    // shaded "better" side (left of zero)
    s.addShape(pres.shapes.RECTANGLE, { x: px + 0.06, y: plotTop + 0.06, w: X(0) - px - 0.06, h: plotH - 0.12, fill: { color: TINT, transparency: 30 }, line: { color: TINT, width: 0 } });
    f.ticks.forEach((t) => {
      s.addShape(pres.shapes.LINE, { x: X(t), y: plotTop, w: 0, h: plotH, line: t === 0 ? { color: BLACK, width: 2.25, dashType: "dash" } : { color: "E3E3E3", width: 1 } });
      s.addShape(pres.shapes.LINE, { x: X(t), y: plotTop + plotH, w: 0, h: 0.12, line: { color: DARKGRAY, width: 1.5 } });
      text(fmt(t), { x: X(t) - 0.6, y: plotTop + plotH + 0.15, w: 1.2, h: 0.4, fontSize: 18, align: "center" });
    });
    f.rows.forEach(([est, lo, hi, sig, lab, star], r) => {
      const cy = plotTop + r * rowH + rowH / 2 + 0.18;
      const col = sig ? BLACK : GRAY;
      s.addShape(pres.shapes.LINE, { x: X(lo), y: cy, w: X(hi) - X(lo), h: 0, line: { color: col, width: 4.5 } });
      [lo, hi].forEach((v) => s.addShape(pres.shapes.LINE, { x: X(v), y: cy - 0.14, w: 0, h: 0.28, line: { color: col, width: 4.5 } }));
      if (star) circ(X(est) - 0.42, cy - 0.42, 0.84, { fill: { color: GOLD, transparency: 55 }, line: { color: GOLD, width: 2.5 } });
      const d = star ? 0.46 : 0.38;
      circ(X(est) - d / 2, cy - d / 2, d, { fill: { color: sig ? RAMP[rowKeys[r]] : WHITE }, line: { color: sig ? BLACK : GRAY, width: sig ? 2 : 3 } });
      text(lab, { x: X(est) - 0.6, y: cy - (star ? 0.86 : 0.74), w: 1.2, h: 0.36, fontSize: 18, bold: sig, align: "center", valign: "middle", color: sig ? BLACK : DARKGRAY, fill: { color: "FAFAFA" } });
    });
  });
  y = plotTop + plotH + 0.6;
  text("Difference vs. inactive adults (percentage points)", { x: plotX0, y, w: plotW, h: 0.42, fontSize: 20, bold: true, align: "center" });
  y += 0.45;
  text([R("◄ ", { color: "8C6A2F" }), R("shaded side = less likely than inactive adults", { italic: true }), R(" · each panel has its own scale", { italic: true })],
    { x: plotX0, y, w: plotW, h: 0.38, fontSize: 18, color: DARKGRAY, align: "center" });
  y += 0.45;
  const takeaway = (body) => {
    rrect(c.cx, y, c.cw, 0.85, { rectRadius: 0.15, fill: { color: TINT }, line: { color: GOLD, width: 1.5 } });
    text([B("Takeaway: ", { color: "8C6A2F" }), R(body)], { x: c.cx + 0.25, y, w: c.cw - 0.5, h: 0.85, fontSize: 20, valign: "middle" });
  };
  takeaway("better general health at every activity level; fewer frequent depressive symptoms only at 4–6 and 7+ sessions a week.");
  y += 1.25;

  // ---- Figure 2: 1,000-dot unit chart of the sample ----
  text([B("Figure 2. "), R("Who was studied: sample by activity level (N = 332,521)")], { x: c.cx, y, w: c.cw, h: 0.5, fontSize: 24 });
  y += 0.62;
  const groups = [["none", 375, "None (0)", "124,773", "37.5%"], ["low", 440, "Low (1–3)", "146,241", "44.0%"],
                  ["med", 122, "Medium (4–6)", "40,660", "12.2%"], ["high", 63, "High (7+)", "20,847", "6.3%"]];
  const COLS = 100, ROWS = 10, gx0 = c.cx + 0.75, gw = c.cw - 0.75, pitch = gw / COLS, dot = pitch * 0.8;
  let idx = 0;
  groups.forEach(([k, n]) => {
    for (let j = 0; j < n; j++, idx++) {
      const col = Math.floor(idx / ROWS), row = idx % ROWS;
      circ(gx0 + col * pitch + (pitch - dot) / 2, y + row * pitch + (pitch - dot) / 2, dot,
        { fill: { color: RAMP[k] }, line: { color: k === "low" ? "D9C48F" : RAMP[k], width: 0.5 } });
    }
  });
  const gh = ROWS * pitch;
  text("1 dot = 0.1%", { x: c.cx - 1.0 + 0.2, y: y + gh / 2 - 0.2, w: 2.0, h: 0.4, fontSize: 16, bold: true, align: "center", valign: "middle", rotate: 270 });
  // x axis: cumulative share of sample
  const axY = y + gh + 0.08;
  s.addShape(pres.shapes.LINE, { x: gx0, y: axY, w: gw, h: 0, line: { color: DARKGRAY, width: 1.5 } });
  [0, 25, 50, 75, 100].forEach((t) => {
    const tx2 = gx0 + gw * t / 100;
    s.addShape(pres.shapes.LINE, { x: tx2, y: axY, w: 0, h: 0.12, line: { color: DARKGRAY, width: 1.5 } });
    text(t + "%", { x: tx2 - 0.6, y: axY + 0.15, w: 1.2, h: 0.38, fontSize: 18, align: "center" });
  });
  text("Cumulative share of sample (%)", { x: gx0, y: axY + 0.52, w: gw, h: 0.42, fontSize: 20, bold: true, align: "center" });
  y = axY + 1.05;
  const lgW = c.cw / 4;
  groups.forEach(([k, , lab, n, pct], i) => {
    const lx = c.cx + i * lgW;
    circ(lx, y + 0.12, 0.5, { fill: { color: RAMP[k] }, line: { color: BLACK, width: 1 } });
    text([B(lab, { fontSize: 18, breakLine: true }), R(n + " · ", { fontSize: 18 }), B(pct, { fontSize: 18 })], { x: lx + 0.6, y, w: lgW - 0.65, h: 0.8, valign: "middle" });
  });
  y += 0.9;
  takeaway("most adults were inactive or low-activity; only 12.2% reached the 4–6 times-a-week level linked to fewer depressive symptoms.");
  y += 1.25;

  // ---- Table 1 ----
  text([B("Table 1. "), R("Specification 3 estimates, percentage points [95% CI]")], { x: c.cx, y, w: c.cw, h: 0.5, fontSize: 24 });
  y += 0.6;
  const hdr = (t) => ({ text: t, options: { bold: true, color: GOLD, fill: { color: BLACK }, align: "center", valign: "middle", fontSize: 20 } });
  const cell = (est, ci, sig) => ({
    text: [{ text: est + (sig ? "*" : ""), options: { bold: sig, fontSize: 22, breakLine: true } }, { text: ci, options: { fontSize: 18, color: DARKGRAY } }],
    options: { align: "center", valign: "middle", fill: { color: sig ? TINT : WHITE } },
  });
  const rl = (k, t) => ({ text: t, options: { bold: true, fontSize: 20, valign: "middle", align: "center", color: RAMP_TXT[k], fill: { color: RAMP[k] } } });
  const plain = (t) => ({ text: t, options: { align: "center", fontSize: 20 } });
  const tbl = [
    [hdr("Activity level"), hdr("General health (poor)"), hdr("Depression frequency"), hdr("Depression level")],
    [rl("low", "Low (1–3)"), cell("−22.7", "[−23.8, −21.6]", true), cell("−0.454", "[−1.5, 0.6]", false), cell("−0.324", "[−1.9, 1.3]", false)],
    [rl("med", "Medium (4–6)"), cell("−34.2", "[−35.7, −32.7]", true), cell("−2.92", "[−4.4, −1.4]", true), cell("−1.45", "[−3.7, 0.8]", false)],
    [rl("high", "High (7+)"), cell("−35.9", "[−37.9, −33.9]", true), cell("−2.50", "[−4.5, −0.5]", true), cell("−0.397", "[−3.3, 2.5]", false)],
    [{ text: "Observations", options: { bold: true, fontSize: 20, align: "center" } }, plain("161,344"), plain("160,974"), plain("66,089")],
    [{ text: "R²", options: { bold: true, fontSize: 20, align: "center" } }, plain("0.204"), plain("0.252"), plain("0.038")],
  ];
  const cw3 = (c.cw - 3.6) / 3;
  s.addTable(tbl, { x: c.cx, y, w: c.cw, colW: [3.6, cw3, cw3, cw3], rowH: [0.7, 1.15, 1.15, 1.15, 0.6, 0.6],
    fontFace: FONT, color: BLACK, border: { type: "solid", pt: 1, color: LINEGRAY }, margin: 0.05, valign: "middle" });
  y += 0.7 + 1.15 * 3 + 0.6 * 2 + 0.12;
  text("* Significant (95% CI excludes 0); highlighted. Reference group: no moderate activity.", { x: c.cx, y, w: c.cw, h: 0.42, fontSize: 18, italic: true, color: DARKGRAY });
  log("Results", y + 0.42, c.bottom);

  // =====================================================================
  // RIGHT COLUMN
  // =====================================================================
  // ----------------------------- DISCUSSION -----------------------------
  const DH = 9.35;
  c = card(RX, Y0, RW, DH, "Discussion", I.FaComments_g);
  y = c.top;
  const insights = [
    ["FaCalendarCheck", [B("4–6 sessions a week "), R("had the largest link with fewer frequent depressive symptoms")], 0.95],
    ["FaHeartbeat", [B("Better self-rated health "), R("at every activity level (−22.7 to −35.9 pp)")], 0.6],
    ["FaBookOpen", [B("Fits prior work "), R("linking activity to lower depression risk"), SUP("4,5")], 0.6],
    ["FaBalanceScale", [B("Socioeconomic status matters: "), R("welfare receipt consistently predicted worse outcomes"), SUP("7")], 0.95],
    ["FaMinusCircle", [B("No link with depression level: "), R("all CIs include 0 (R² = 0.038)")], 0.6],
  ];
  insights.forEach(([ic, runs, h]) => {
    iconBadge(I[ic], c.cx, y + h / 2 - 0.36, 0.72, TINT, GOLD);
    text(runs, { x: c.cx + 0.95, y, w: c.cw - 0.95, h, fontSize: 24, valign: "middle" });
    y += h + 0.18;
  });
  y += 0.05;
  // limitations as 2x2 tiles
  text("Limitations", { x: c.cx, y, w: c.cw, h: 0.5, fontSize: 26, bold: true });
  y += 0.58;
  const ltW = (c.cw - 0.25) / 2;
  [["FaCommentDots", "Self-report bias"], ["FaQuestion", "Unmeasured confounders"], ["FaLink", "Association, not cause"], ["FaPuzzlePiece", "Income only in categories"]]
    .forEach(([ic, lab], i) => {
      const lx = c.cx + (i % 2) * (ltW + 0.25), ly = y + Math.floor(i / 2) * 1.12;
      rrect(lx, ly, ltW, 0.97, { rectRadius: 0.18, fill: { color: SOFT }, line: { color: "BFBFBF", width: 1.25 } });
      img(I[ic], lx + 0.2, ly + 0.27, 0.43);
      text(lab, { x: lx + 0.8, y: ly, w: ltW - 0.9, h: 0.97, fontSize: 24, bold: true, valign: "middle" });
    });
  y += 2.21;
  log("Discussion", y, c.bottom);

  // ----------------------------- CONCLUSIONS -----------------------------
  const CY = Y0 + DH + 0.5, CH = 8.5;
  c = card(RX, CY, RW, CH, "Conclusions", I.FaLightbulb_g);
  y = c.top;
  const msgs = [
    ["Being moderately active 4–6 times a week was linked to a lower chance of feeling depressed often.", 0.95],
    ["Any regular moderate activity – even 1–3 times a week – was linked to better self-rated health.", 0.95],
    ["Activity was not linked to how strongly depression is felt, so it is a low-cost complement to care, not a cure.", 0.95],
  ];
  msgs.forEach(([m, h], i) => {
    text(String(i + 1), { x: c.cx, y: y - 0.12, w: 0.9, h: 1.0, fontSize: 66, bold: true, color: GOLD, align: "center", valign: "top" });
    text(m, { x: c.cx + 1.05, y, w: c.cw - 1.05, h, fontSize: 24, bold: i === 0 });
    y += h + 0.12;
  });
  rrect(c.cx, y, c.cw, 1.2, { rectRadius: 0.2, fill: { color: TINT }, line: { color: GOLD, width: 2 } });
  text([B("Big picture: "), R("as depression rises and care stays hard to reach, regular moderate activity is a low-cost tool worth promoting.")],
    { x: c.cx + 0.25, y: y + 0.1, w: c.cw - 0.5, h: 1.0, fontSize: 24, valign: "middle" });
  y += 1.4;
  rrect(c.cx, y, c.cw, 2.05, { rectRadius: 0.2, fill: { color: BLACK }, line: { color: GOLD, width: 2.5 }, shadow: shadow() });
  circ(c.cx + 0.25, y + 0.25, 0.8, { fill: { color: "2E2E2E" }, line: { color: GOLD, width: 2 } });
  img(I.FaRocket_g, c.cx + 0.45, y + 0.45, 0.4);
  text("Future directions", { x: c.cx + 1.25, y: y + 0.25, w: c.cw - 1.5, h: 0.8, fontSize: 26, bold: true, color: GOLD, valign: "middle" });
  text("Test activity as a treatment for people clinically diagnosed with depression, tracking real-time, individual-level change.",
    { x: c.cx + 0.3, y: y + 1.0, w: c.cw - 0.6, h: 0.95, fontSize: 22, color: WHITE });
  log("Conclusions", y + 2.05, c.bottom);

  // ----------------------------- ACKNOWLEDGEMENTS -----------------------------
  const AKY = CY + CH + 0.5;
  c = card(RX, AKY, RW, Y1 - AKY, "Acknowledgements", I.FaHandsHelping_g);
  y = c.top - 0.05;
  text("I thank Professor Osborne for his helpful feedback on the manuscript, and my mom and dad for their constant support and encouragement.",
    { x: c.cx, y, w: c.cw, h: 1.3, fontSize: 24 });
  y += 1.38;
  text([B("Funding: "), R("This independent study received no external funding.", { breakLine: true }), B("Data: "), R("IPUMS NHIS v8.1 (Blewett et al., 2025).")],
    { x: c.cx, y, w: c.cw, h: 0.8, fontSize: 20 });
  y += 0.92;
  rrect(c.cx, y, c.cw, 1.95, { rectRadius: 0.2, fill: { color: TINT }, line: { color: GOLD, width: 2 } });
  text([
    { text: "Author Contribution Statement", options: { bold: true, fontSize: 22, breakLine: true } },
    { text: "NG (sole author) conceived and designed the study, prepared the data, performed all analyses, interpreted the results, wrote the manuscript, and designed this poster.", options: { fontSize: 20 } },
  ], { x: c.cx + 0.2, y: y + 0.08, w: c.cw - 0.4, h: 1.8, valign: "middle" });
  y += 2.1;
  const refs = [
    "1. NIMH (n.d.). Major Depression.",
    "2. WHO (2025). Depressive disorder (depression).",
    "3. Goodwin, R. D., et al. (2022). Am. J. Prev. Med., 63(5), 726–733.",
    "4. Kandola, A., et al. (2019). Neurosci. Biobehav. Rev., 107, 525–539.",
    "5. Pearce, M., et al. (2022). JAMA Psychiatry, 79(6), 550–559.",
    "6. Blewett, L. A., et al. (2025). IPUMS Health Surveys: NHIS, v8.1. doi.org/10.18128/D071.V8.1",
    "7. Alzahrani, H., et al. (2024). J. Adv. Trends Med. Res., 862–869.",
  ];
  text([B("Selected references", { fontSize: 18, color: BLACK, breakLine: true })].concat(refs.map((r, i) => R(r, { breakLine: i < refs.length - 1 }))),
    { x: c.cx, y, w: c.cw, h: 2.4, fontSize: 16, color: DARKGRAY, paraSpaceAfter: 1 });
  log("Acknowledgements", y + 2.4, c.bottom);

  await pres.writeFile({ fileName: OUT });
  console.log("wrote", OUT);
})();
