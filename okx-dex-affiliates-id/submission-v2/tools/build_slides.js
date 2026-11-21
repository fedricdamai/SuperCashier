// Six-slide companion deck (main points only), numbers from slides.json.
// Usage: node build_slides.js <slides.json> <out.pptx>
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const [src, out] = process.argv.slice(2);
const D = JSON.parse(fs.readFileSync(src, "utf8"));

const NAVY = "13203A", BLUE = "2C5BE0", BG = "F4F6F9", CARD = "FFFFFF", CARDLINE = "E1E6EE", TINT = "E8EFFF",
  MUTED = "5B667A", TEXT = "1F2937", GREEN = "1E7A45", NAVY2 = "22325A";
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
    { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.3, w: 9, h: 0.9, fontSize: 24, bold: true, color: NAVY, fontFace: FONT, valign: "top", align: "left", margin: 0 }, text: "" } },
  ],
  slideNumber: { x: 9.1, y: 5.22, w: 0.4, h: 0.25, fontSize: 9, color: MUTED, fontFace: FONT, align: "right" },
});
pres.defineSlideMaster({ title: "DARK", background: { color: NAVY }, objects: [] });

const kicker = (s, t, y = 0.32) => s.addText(t, { x: 0.5, y, w: 9, h: 0.25, fontSize: 11, bold: true, color: BLUE, fontFace: FONT, margin: 0, charSpacing: 1, isTextBox: true });
const card = (s, x, y, w, h, fill = CARD) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: CARDLINE, width: 0.75 }, rectRadius: 0.08 });
function bullets(s, items, x, y, w, h, size = 12) {
  s.addText(items.map((t, i) => ({ text: t, options: { bullet: { indent: 12 }, breakLine: i < items.length - 1, paraSpaceAfter: 4 } })),
    { x, y, w, h, fontSize: size, color: TEXT, fontFace: FONT, valign: "top", margin: 0, isTextBox: true });
}
const cardTitle = (s, t, x, y, w) => s.addText(t, { x, y, w, h: 0.32, fontSize: 14, bold: true, color: NAVY, fontFace: FONT, margin: 0, isTextBox: true });
const foot = (s, t) => s.addText(t, { x: 0.5, y: 4.88, w: 9, h: 0.3, fontSize: 9, color: MUTED, fontFace: FONT, margin: 0, valign: "top", isTextBox: true });
function tbl(s, head, rows, opt) {
  const hl = opt.hl || [];
  const data = [head.map((h, i) => ({ text: h, options: { bold: true, color: NAVY, align: opt.align[i], fill: { color: "E2E6EC" } } }))];
  rows.forEach((r, ri) => data.push(r.map((c, i) => ({ text: String(c), options: {
    align: opt.align[i], bold: (opt.boldLast && ri === rows.length - 1) || (i === 0 && opt.boldFirst) || false,
    fill: { color: hl.includes(ri) ? TINT : CARD } } }))));
  s.addTable(data, { x: opt.x, y: opt.y, w: opt.w, colW: opt.colW, fontSize: opt.size || 11, fontFace: FONT, color: TEXT,
    border: { type: "solid", pt: 0.5, color: "D5DAE1" }, margin: [3, 5, 3, 5], rowH: opt.rowH || 0.26, valign: "middle" });
}

// ---------------------------------------------------------------- 1. title and summary
{
  const s = pres.addSlide({ masterName: "DARK" });
  kicker(s, "OKX P4 DEX BD AFFILIATE · INDONESIA · TAKE-HOME", 0.55);
  s.addText("Three trading circles first, one funded pilot, 30 days to traction", { x: 0.5, y: 0.9, w: 8.6, h: 1.3, fontSize: 32, bold: true, color: "FFFFFF", fontFace: FONT, margin: 0, valign: "top", isTextBox: true });
  s.addText("Fedric Intan Damai · November 2026", { x: 0.5, y: 2.25, w: 8, h: 0.3, fontSize: 13, color: "C9D3E6", fontFace: FONT, margin: 0, isTextBox: true });
  const stats = [
    [String(D.top3_ftd), "first-time traders in month one from Kevin Susanto, Bevan and Ponyin"],
    ["B first", `the one pilot with no downside: positive from ${D.S.B50} at half volume`],
    [`${D.activated} of 8`, "partners activated by day 30 (10+ FTDs, 50+ trades a day)"],
    [String(D.tpd30), `trades a day by day 30, from about ${D.ftd30} first-time traders`],
  ];
  stats.forEach(([big, small], i) => {
    const x = 0.5 + i * 2.3;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 3.05, w: 2.1, h: 1.85, fill: { color: NAVY2 }, line: { color: NAVY2 }, rectRadius: 0.08 });
    s.addText(big, { x: x + 0.18, y: 3.2, w: 1.8, h: 0.6, fontSize: 28, bold: true, color: "FFFFFF", fontFace: FONT, margin: 0, isTextBox: true });
    s.addText(small, { x: x + 0.18, y: 3.85, w: 1.78, h: 0.95, fontSize: 11, color: "C9D3E6", fontFace: FONT, margin: 0, valign: "top", isTextBox: true });
  });
}

// ---------------------------------------------------------------- 2. Task 1
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  s.addText("Kevin, Bevan and Ponyin first: they trade in public and each brings a circle", { placeholder: "title" });
  const rows = D.partners.map((p, i) => [i + 1, p.name, p.proof, p.fnf, p.ftd, p.swaps]);
  rows.push(["", "Total", "", D.total.fnf, D.total.ftd, D.total.swaps]);
  tbl(s, ["#", "Partner", "Verified proof", "FnF", "FTDs", "Swaps / month"], rows, {
    x: 0.5, y: 1.35, w: 9, colW: [0.35, 2.25, 3.95, 0.7, 0.7, 1.05], align: ["left", "left", "left", "right", "right", "right"],
    hl: [0, 1, 2], boldLast: true, size: 11, rowH: 0.3 });
  foot(s, "Ranked on proof of trading, then the FnF traders they bring, then rival flow to move, then FTDs and trades a day. FTDs = reach × 2% sign up × 15% trade; FnF traders at $547 a day each. Meet in person: Kevin and Bevan (Jakarta). Ponyin opens ApexFNF and the CVC team (0xRusdi, Tomket clan, Setya).");
}

// ---------------------------------------------------------------- 3. Task 2
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  s.addText("Pilot B first and fund its growth; C starts at 2x, A follows on its own", { placeholder: "title" });
  const S = D.S;
  tbl(s, ["Partner", "Their ask", "Month 1 at their ask", "Our deal: at 50%", "At forecast", "At 2x", "At 3x"], [
    ["B · repeat traders", "35%, no fee", S.B100_ask, S.B50, S.B100, S.B200, S.B300],
    ["C · educator", "$3,000 + 10%", S.C100_ask, S.C50, S.C100, S.C200, S.C300],
    ["A · airdrop", "$2,000 + 20%", S.A100_ask, S.A50, S.A100, S.A200, S.A300],
  ], { x: 0.5, y: 1.35, w: 9, colW: [1.9, 1.35, 1.35, 1.25, 1.1, 1.0, 1.05], align: ["left", "left", "right", "right", "right", "right", "right"], hl: [0], boldFirst: true, size: 11, rowH: 0.3 });
  card(s, 0.5, 2.88, 4.4, 1.9); card(s, 5.1, 2.88, 4.4, 1.9);
  cardTitle(s, "The growth deal", 0.7, 2.97, 4.0);
  bullets(s, ["B: no retainer; share 30% rising to 45% by bracket",
    "C: $500 retainer + $500 at forecast; share 15% rising to 30%; co-hosted event",
    "Cash bonus within 7 days at 1x, 2x and 3x forecast",
    `B's $250 competition pays back above ${D.payback}; stop it at week 3 if off track`], 0.7, 3.32, 4.0, 1.4, 11.5);
  cardTitle(s, "Before signing, and if volume halves", 5.3, 2.97, 4.0);
  bullets(s, ["Evidence: 90 days of dashboards, proof of users, wallet addresses we check onchain",
    `At 50%, B is still ${S.B50}: check market, incentive or the group itself`,
    "No fixed cost extras: weekly payouts, rate ratchet, 5% network override"], 5.3, 3.32, 4.0, 1.4, 11.5);
  foot(s, "Month 1 = volume × 0.04% − retainer − revenue share − bonus − campaign spend (brief's net revenue rate). Every scenario in the spreadsheet's Task 2 tab stays positive.");
}

// ---------------------------------------------------------------- 4. Task 3
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  s.addText("Kevin Susanto: open with his own trade, offer fee share plus milestone cash", { placeholder: "title" });
  const cols = [
    ["The offer", ["No retainer", "Fee share 30%, rising to 40% as volume grows", "Cash within 7 days at $250K, $500K and $1M",
      "Co-hosted live session + 7-day competition", `About ${D.kev.share} share + ${D.kev.bonus} bonus at his forecast`]],
    ["The messages", ["Hook: +$53,566 in a day, #3 on fomo", "Who I am and the offer in line two", "What he and his followers get",
      "Next step: reply, coffee in Jakarta, link live in 48h", "Follow-up holds his first-wave slot", "Higher pay elsewhere: we're stable, rates grow, bisa dibahas"]],
    ["First 14 days", ["Days 1-2: link, terms, compliance", "Days 2-4: 3 FnF traders through the full flow", "Day 5: live session + competition",
      "Days 8-10: fix the weakest funnel step; UI/UX feedback", `Day 14: ${D.kev.plan[5].replace(" by day 14", "")}`]],
  ];
  cols.forEach(([t, items], i) => {
    const x = 0.5 + i * 3.05;
    card(s, x, 1.35, 2.85, 3.4, i === 0 ? TINT : CARD);
    cardTitle(s, t, x + 0.18, 1.47, 2.5);
    bullets(s, items, x + 0.18, 1.85, 2.52, 2.8, 11.5);
  });
  foot(s, `Kevin's first month at base case: ${D.kev.ftd} FTDs, ~${D.kev.tpd} trades a day, ${D.kev.swaps} of swaps. The fee share is a published OKX feature; brackets, bonus and campaign need OKX approval.`);
}

// ---------------------------------------------------------------- 5. Task 4
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  s.addText(`30 days: 8 partners live, ${D.activated} activated, about ${D.ftd30} new traders`, { placeholder: "title" });
  tbl(s, ["Week", "Outreach", "Meetings", "Live", "Activated", "FTDs", "Trades a day"], D.weeks, {
    x: 0.5, y: 1.35, w: 5.6, colW: [1.15, 0.8, 0.8, 0.6, 0.85, 0.55, 0.85], align: ["left", "right", "right", "right", "right", "right", "right"],
    hl: [3], size: 10, rowH: 0.34 });
  s.addText(`By day 30: ${D.swaps30} of swaps from our 8 plus pilot B; ${D.rev30} of revenue at 0.04%.`, { x: 0.5, y: 3.25, w: 5.6, h: 0.5, fontSize: 11, color: TEXT, fontFace: FONT, margin: 0, isTextBox: true });
  const cards = [
    ["Activated partner", "10+ FTDs and 50+ trades a day within 30 days; 30%+ still trading after 14 days; under 5% flagged wallets"],
    ["20 messages, no meetings", "Go through connections of connections: direct intros from Kevin, Bevan and Ponyin's circles, then lead with the offer"],
    ["Day 30 decision", "Continue if activated; renegotiate if FnF trade but FTDs fall short; stop under 5 FTDs or over 10% flagged"],
  ];
  cards.forEach(([t, b], i) => {
    const y = 1.35 + i * 1.15;
    card(s, 6.3, y, 3.2, 1.02);
    s.addText(t, { x: 6.45, y: y + 0.08, w: 2.9, h: 0.28, fontSize: 12.5, bold: true, color: NAVY, fontFace: FONT, margin: 0, isTextBox: true });
    s.addText(b, { x: 6.45, y: y + 0.36, w: 2.95, h: 0.62, fontSize: 10, color: TEXT, fontFace: FONT, margin: 0, valign: "top", isTextBox: true });
  });
  foot(s, "Waves: Kevin, Bevan and Ponyin on day 10; GuarEmperor, Brandon Salim and Crypto Ndo on day 17; DegenRex and Meteora LP Army on day 24. Pilot B live from day 10; C starts once B passes 2x.");
}

// ---------------------------------------------------------------- 6. commitment
{
  const s = pres.addSlide({ masterName: "DARK" });
  kicker(s, "COMMITMENT AND METHOD", 0.55);
  s.addText("Ready to start in 2 weeks, full-time, on the ground across Indonesia and Southeast Asia", { x: 0.5, y: 0.9, w: 8.8, h: 1.2, fontSize: 28, bold: true, color: "FFFFFF", fontFace: FONT, margin: 0, valign: "top", isTextBox: true });
  const blocks = [
    ["Commitment", ["Earliest start: in 2 weeks", "Full-time: yes, from 2 weeks", "Other commitments: marketer role at Surgence Labs", "In person: Indonesia and SEA; partner meetings and events welcome"]],
    ["How I worked", ["I set the strategy, the selection logic and every judgment call", "AI gathered public data, automated repetitive work and fact-checked", "I re-checked every part and ran independent fact-checks with several AI agents", "Public information only"]],
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
