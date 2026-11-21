// Six-slide companion deck: an executive summary, one slide per task (titles as in the brief), commitment. Text from slides.json.
// Usage: node build_slides.js <slides.json> <out.pptx>
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const [src, out] = process.argv.slice(2);
const D = JSON.parse(fs.readFileSync(src, "utf8"));

const NAVY = "13203A", BLUE = "2C5BE0", BG = "F4F6F9", CARD = "FFFFFF", CARDLINE = "E1E6EE", TINT = "E8EFFF",
  MUTED = "5B667A", TEXT = "1F2937", NAVY2 = "22325A";
const FONT = "Arial";

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10 x 5.625 in
pres.author = "Fedric Intan Damai";
pres.title = "Indonesia DEX Affiliate Plan";
pres.theme = { headFontFace: FONT, bodyFontFace: FONT };

pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: BG },
  objects: [
    { text: { text: "Fedric Intan Damai · OKX DEX BD Affiliate (Indonesia)", options: { x: 0.5, y: 5.22, w: 6, h: 0.25, fontSize: 9, color: MUTED, fontFace: FONT, margin: 0 } } },
    { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.28, w: 9, h: 0.5, fontSize: 22, bold: true, color: NAVY, fontFace: FONT, valign: "top", align: "left", margin: 0 }, text: "" } },
  ],
  slideNumber: { x: 9.1, y: 5.22, w: 0.4, h: 0.25, fontSize: 9, color: MUTED, fontFace: FONT, align: "right" },
});
pres.defineSlideMaster({ title: "DARK", background: { color: NAVY }, objects: [] });

const kicker = (s, t, y = 0.32) => s.addText(t, { x: 0.5, y, w: 9, h: 0.25, fontSize: 11, bold: true, color: "8FB0FF", fontFace: FONT, margin: 0, charSpacing: 1, isTextBox: true });
const answer = (s, t) => s.addText([{ text: "Answer: ", options: { bold: true, color: BLUE } }, { text: t, options: { color: TEXT } }],
  { x: 0.5, y: 0.8, w: 9, h: 0.5, fontSize: 12, fontFace: FONT, margin: 0, valign: "top", isTextBox: true });
const card = (s, x, y, w, h, fill = CARD) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: CARDLINE, width: 0.75 }, rectRadius: 0.08 });
function bullets(s, items, x, y, w, h, size = 11) {
  s.addText(items.map((t, i) => ({ text: t, options: { bullet: { indent: 11 }, breakLine: i < items.length - 1, paraSpaceAfter: 3 } })),
    { x, y, w, h, fontSize: size, color: TEXT, fontFace: FONT, valign: "top", margin: 0, isTextBox: true });
}
const cardTitle = (s, t, x, y, w) => s.addText(t, { x, y, w, h: 0.3, fontSize: 13, bold: true, color: NAVY, fontFace: FONT, margin: 0, isTextBox: true });
const foot = (s, t) => s.addText(t, { x: 0.5, y: 4.86, w: 9, h: 0.32, fontSize: 9, color: MUTED, fontFace: FONT, margin: 0, valign: "top", isTextBox: true });
function tbl(s, head, rows, opt) {
  const hl = opt.hl || [];
  const data = [head.map((h, i) => ({ text: h, options: { bold: true, color: NAVY, align: opt.align[i], fill: { color: "E2E6EC" } } }))];
  rows.forEach((r, ri) => data.push(r.map((c, i) => ({ text: String(c), options: {
    align: opt.align[i], bold: (opt.boldLast && ri === rows.length - 1) || (i === 0 && opt.boldFirst) || false,
    fill: { color: hl.includes(ri) ? TINT : CARD } } }))));
  s.addTable(data, { x: opt.x, y: opt.y, w: opt.w, colW: opt.colW, fontSize: opt.size || 10, fontFace: FONT, color: TEXT,
    border: { type: "solid", pt: 0.5, color: "D5DAE1" }, margin: [2, 5, 2, 5], rowH: opt.rowH || 0.26, valign: "middle" });
}

// ---------------------------------------------------------------- 1. executive summary
{
  const s = pres.addSlide({ masterName: "DARK" });
  kicker(s, "OKX P4 DEX BD AFFILIATE · INDONESIA · EXECUTIVE SUMMARY", 0.5);
  s.addText("Indonesia DEX Affiliate Plan", { x: 0.5, y: 0.85, w: 8.6, h: 0.7, fontSize: 34, bold: true, color: "FFFFFF", fontFace: FONT, margin: 0, valign: "top", isTextBox: true });
  s.addText("Approach three trading circles first, pilot B on a growth deal, and turn the first 30 days into live, activated partners.",
    { x: 0.5, y: 1.6, w: 8.6, h: 0.65, fontSize: 15, color: "DCE3F0", fontFace: FONT, margin: 0, valign: "top", isTextBox: true });
  s.addText("Fedric Intan Damai · November 2026", { x: 0.5, y: 2.35, w: 8, h: 0.3, fontSize: 12, color: "C9D3E6", fontFace: FONT, margin: 0, isTextBox: true });
  D.stats.forEach(([big, small], i) => {
    const x = 0.5 + i * 2.3;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 2.95, w: 2.1, h: 1.75, fill: { color: NAVY2 }, line: { color: NAVY2 }, rectRadius: 0.08 });
    s.addText(big, { x: x + 0.18, y: 3.08, w: 1.8, h: 0.55, fontSize: 26, bold: true, color: "FFFFFF", fontFace: FONT, margin: 0, isTextBox: true });
    s.addText(small, { x: x + 0.18, y: 3.68, w: 1.78, h: 0.95, fontSize: 10.5, color: "C9D3E6", fontFace: FONT, margin: 0, valign: "top", isTextBox: true });
  });
  s.addText("FTD = first-time trader · FnF = friends and family, the power traders a partner brings · DEX = decentralised exchange · LP = liquidity provider",
    { x: 0.5, y: 4.92, w: 9, h: 0.3, fontSize: 9, color: "9FB0CC", fontFace: FONT, margin: 0, isTextBox: true });
}

// ---------------------------------------------------------------- 2. Task 1
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  s.addText("Task 1: Find partners beyond airdrop communities", { placeholder: "title" });
  answer(s, D.t1_answer);
  tbl(s, ["#", "Partner", "Evidence of fit", "Priority", "FTDs", "FnF traders", "Swaps a month"], D.t1_rows, {
    x: 0.5, y: 1.38, w: 9, colW: [0.3, 2.0, 3.3, 0.8, 0.6, 1.0, 1.0], align: ["left", "left", "left", "left", "right", "right", "right"],
    hl: [0, 1, 2], boldLast: true, size: 10, rowH: 0.31 });
  foot(s, D.t1_foot);
}

// ---------------------------------------------------------------- 3. Task 2
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  s.addText("Task 2: Make a commercial decision", { placeholder: "title" });
  answer(s, D.t2_answer);
  tbl(s, ["Partner", "Their ask", "Month 1 at their ask", "Our deal: at 50%", "At forecast", "At 2x", "At 3x"], D.t2_rows,
    { x: 0.5, y: 1.38, w: 9, colW: [2.1, 1.3, 1.3, 1.2, 1.05, 1.0, 1.05], align: ["left", "left", "right", "right", "right", "right", "right"], hl: [0], boldFirst: true, size: 10.5, rowH: 0.3 });
  card(s, 0.5, 2.78, 4.4, 2.0); card(s, 5.1, 2.78, 4.4, 2.0);
  cardTitle(s, "Counter-offer: a growth deal", 0.7, 2.88, 4.0);
  bullets(s, D.t2_deal, 0.7, 3.22, 4.0, 1.5, 11);
  cardTitle(s, "Before signing, and if volume halves", 5.3, 2.88, 4.0);
  bullets(s, D.t2_right, 5.3, 3.22, 4.0, 1.5, 11);
  foot(s, D.t2_foot);
}

// ---------------------------------------------------------------- 4. Task 3
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  s.addText("Task 3: Show how you would win and activate a partner", { placeholder: "title" });
  answer(s, D.t3_answer);
  D.t3_cols.forEach(([t, items], i) => {
    const x = 0.5 + i * 3.05;
    card(s, x, 1.42, 2.85, 3.35, i === 0 ? TINT : CARD);
    cardTitle(s, t, x + 0.18, 1.52, 2.5);
    bullets(s, items, x + 0.18, 1.88, 2.52, 2.8, 11);
  });
  foot(s, D.t3_foot);
}

// ---------------------------------------------------------------- 5. Task 4
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  s.addText("Task 4: Demonstrate practical execution and persistence", { placeholder: "title" });
  answer(s, D.t4_answer);
  tbl(s, ["Week", "Focus", "Messages", "Meetings", "Live", "Activated", "FTDs", "Trades a day"], D.t4_rows, {
    x: 0.5, y: 1.38, w: 9, colW: [1.15, 3.5, 0.85, 0.8, 0.6, 0.85, 0.5, 0.75], align: ["left", "left", "right", "right", "right", "right", "right", "right"],
    hl: [3], size: 10, rowH: 0.36 });
  D.t4_cards.forEach(([t, b], i) => {
    const x = 0.5 + i * 3.05;
    card(s, x, 3.42, 2.85, 1.35);
    s.addText(t, { x: x + 0.15, y: 3.5, w: 2.6, h: 0.28, fontSize: 12, bold: true, color: NAVY, fontFace: FONT, margin: 0, isTextBox: true });
    s.addText(b, { x: x + 0.15, y: 3.8, w: 2.6, h: 0.92, fontSize: 10, color: TEXT, fontFace: FONT, margin: 0, valign: "top", isTextBox: true });
  });
  foot(s, D.t4_foot);
}

// ---------------------------------------------------------------- 6. commitment
{
  const s = pres.addSlide({ masterName: "DARK" });
  kicker(s, "COMMITMENT AND HOW I WORKED", 0.55);
  s.addText("Ready to start in 2 weeks, full-time, on the ground across Indonesia and Southeast Asia", { x: 0.5, y: 0.9, w: 8.8, h: 1.2, fontSize: 28, bold: true, color: "FFFFFF", fontFace: FONT, margin: 0, valign: "top", isTextBox: true });
  const blocks = [
    ["Commitment", ["Earliest start: in 2 weeks", "Full-time: yes, from 2 weeks", "Other commitments: marketer role at Surgence Labs",
      "In person: across Indonesia and Southeast Asia; partner meetings and events welcome"]],
    ["How I worked", ["I set the strategy, the selection logic and every judgment call", "AI gathered public data, automated repetitive work and fact-checked",
      "I re-checked every part and ran independent fact-checks with several AI agents", "Public information only; every source is linked in the write-up and spreadsheet"]],
  ];
  blocks.forEach(([t, items], i) => {
    const x = 0.5 + i * 4.6;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 2.4, w: 4.4, h: 2.5, fill: { color: NAVY2 }, line: { color: NAVY2 }, rectRadius: 0.08 });
    s.addText(t, { x: x + 0.2, y: 2.52, w: 4, h: 0.32, fontSize: 14, bold: true, color: "FFFFFF", fontFace: FONT, margin: 0, isTextBox: true });
    s.addText(items.map((x2, j) => ({ text: x2, options: { bullet: { indent: 12 }, breakLine: j < items.length - 1, paraSpaceAfter: 5 } })),
      { x: x + 0.2, y: 2.92, w: 4.0, h: 1.9, fontSize: 12, color: "DCE3F0", fontFace: FONT, margin: 0, valign: "top", isTextBox: true });
  });
}

pres.writeFile({ fileName: out }).then(() => console.log("wrote", out));
