// Build the take-home write-up (.docx) from the JSON blocks that make_writeup.py exports.
// Usage: node build_writeup.js <doc_blocks.json> <out.docx> [bodyPt]
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle,
  AlignmentType, LevelFormat, HeadingLevel, Footer, PageNumber, ExternalHyperlink, PageBreak,
} = require("docx");

const [src, out, bodyArg] = process.argv.slice(2);
const blocks = JSON.parse(fs.readFileSync(src, "utf8"));
const BODY = Math.round(parseFloat(bodyArg || "11") * 2); // half-points
const SMALL = 18; // 9 pt
const TBL = 16; // 8 pt
const FONT = "Arial";
const NAVY = "13203A", GREY = "5B667A", LINK = "0563C1";
const PAGE_W = 12240, MARGIN = 1080, CONTENT_W = PAGE_W - 2 * MARGIN;

const decode = (s) => s.replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">")
  .replace(/&quot;/g, '"').replace(/&#x27;/g, "'").replace(/&#39;/g, "'");

// "<b>x</b> y <i>z</i>" -> TextRuns
function runs(html, size, extra = {}) {
  const res = [];
  let bold = false, italics = false;
  for (const part of html.split(/(<\/?[bi]>)/)) {
    if (part === "<b>") bold = true;
    else if (part === "</b>") bold = false;
    else if (part === "<i>") italics = true;
    else if (part === "</i>") italics = false;
    else if (part) res.push(new TextRun({ text: decode(part), bold: bold || extra.bold || undefined, italics: italics || extra.italics || undefined,
      size, font: FONT, color: extra.color }));
  }
  return res;
}

const border = { style: BorderStyle.SINGLE, size: 4, color: "C8CCD2" };
const borders = { top: border, bottom: border, left: border, right: border };

function table(b) {
  const total = b.widths.reduce((x, y) => x + y, 0);
  const cols = b.widths.map((w) => Math.floor((w / total) * CONTENT_W));
  cols[cols.length - 1] += CONTENT_W - cols.reduce((x, y) => x + y, 0);
  const mk = (txt0, i, head, bold) => {
    let txt = String(txt0), fill = head ? "EEF1F5" : undefined;
    if (txt.startsWith("§")) { fill = "2C5BE0"; txt = txt.slice(1); }
    else if (txt.startsWith("¤")) { fill = "8FB0FF"; txt = txt.slice(1); }
    return new TableCell({
    borders, width: { size: cols[i], type: WidthType.DXA },
    shading: fill ? { fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
    margins: { top: 30, bottom: 30, left: 60, right: 60 },
    children: String(txt).split("\n").map((line) => new Paragraph({
      alignment: !head && b.num_from !== null && i >= b.num_from ? AlignmentType.RIGHT : AlignmentType.LEFT,
      spacing: { before: 0, after: 0 },
      children: runs(line, b.size || TBL, { bold: head || bold }),
    })),
  }); };
  const rows = [new TableRow({ tableHeader: true, cantSplit: true, children: b.head.map((h, i) => mk(h, i, true, false)) })];
  b.rows.forEach((r, ri) => rows.push(new TableRow({
    cantSplit: true, children: r.map((c, i) => mk(c, i, false, b.bold_last && ri === b.rows.length - 1)),
  })));
  return new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: cols, rows });
}

const children = [];
let pendingBreak = false;
for (const b of blocks) {
  const size = b.small ? SMALL : BODY;
  if (b.type === "h1") children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: b.text })] }));
  else if (b.type === "h2") { children.push(new Paragraph({ heading: HeadingLevel.HEADING_2, pageBreakBefore: pendingBreak || undefined, children: [new TextRun({ text: b.text })] })); pendingBreak = false; }
  else if (b.type === "pagebreak") pendingBreak = true;
  else if (b.type === "p") {
    const kids = [];
    if (b.label) kids.push(new TextRun({ text: decode(b.label) + " ", bold: true, size, font: FONT, color: b.small ? GREY : undefined }));
    kids.push(...runs(b.html, size, { color: b.small ? GREY : undefined }));
    children.push(new Paragraph({ spacing: { before: 0, after: 60 }, children: kids }));
  } else if (b.type === "ul" || b.type === "ol") {
    b.items.forEach((it, i) => children.push(new Paragraph({
      numbering: { reference: b.type === "ul" ? "bullets" : "numbers", level: 0 },
      spacing: { before: 0, after: i === b.items.length - 1 ? 70 : 0 },
      children: runs(it, size, { color: b.small ? GREY : undefined }),
    })));
  } else if (b.type === "quote") {
    b.lines.forEach((line, i) => children.push(new Paragraph({
      indent: { left: 360 }, spacing: { before: 0, after: i === b.lines.length - 1 ? 90 : 30 },
      border: { left: { style: BorderStyle.SINGLE, size: 12, color: "C8CCD2", space: 8 } },
      children: runs(line, BODY - 1, { italics: true }),
    })));
  } else if (b.type === "table") {
    children.push(table(b));
    children.push(new Paragraph({ spacing: { before: 0, after: 50 }, children: [] }));
  } else if (b.type === "links") {
    children.push(new Paragraph({ spacing: { before: 80, after: 20 }, keepNext: true,
      children: [new TextRun({ text: b.group, bold: true, size: SMALL + 2, font: FONT, color: NAVY })] }));
    b.items.forEach((it) => children.push(new Paragraph({
      numbering: { reference: "bullets", level: 0 }, spacing: { before: 0, after: 0 },
      children: [new TextRun({ text: it.label + ": ", size: SMALL, font: FONT }),
        new ExternalHyperlink({ link: it.url, children: [new TextRun({ text: it.url.replace(/^https?:\/\//, ""), size: SMALL, font: FONT, color: LINK, underline: {} })] })],
    })));
  }
}

const doc = new Document({
  creator: "Fedric Intan Damai",
  title: "Indonesia DEX Affiliate Plan",
  styles: {
    default: { document: { run: { font: FONT, size: BODY, color: "1A1A1A" } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 32, bold: true, color: NAVY }, paragraph: { spacing: { before: 0, after: 60 }, keepNext: true, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 25, bold: true, color: NAVY }, paragraph: { spacing: { before: 120, after: 50 }, keepNext: true, outlineLevel: 1 } },
    ],
  },
  numbering: {
    config: [
      { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 360, hanging: 240 } } } }] },
      { reference: "numbers", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 360, hanging: 260 } } } }] },
    ],
  },
  sections: [{
    properties: { page: { size: { width: PAGE_W, height: 15840 }, margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [
      new TextRun({ text: "Fedric Intan Damai · OKX DEX BD Affiliate (Indonesia) · page ", size: 16, color: GREY, font: FONT }),
      new TextRun({ children: [PageNumber.CURRENT], size: 16, color: GREY, font: FONT }),
    ] })] }) },
    children,
  }],
});
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(out, buf); console.log("wrote", out); });
