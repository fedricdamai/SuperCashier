"""Regenerate deck_slides_v2.md from deck_v2.html (on-slide content, footer, speaker notes)."""
import re
import sys
from pathlib import Path

from bs4 import BeautifulSoup, NavigableString

src, out = Path(sys.argv[1]), Path(sys.argv[2])
soup = BeautifulSoup(src.read_text(), "html.parser")


def text(el):
    t = el.get_text(" ", strip=True) if not isinstance(el, NavigableString) else str(el).strip()
    t = re.sub(r"\s+", " ", t)
    t = re.sub(r"\s+([,.;:)])", r"\1", t).replace("( ", "(")
    return t


def sup_text(el):
    """Text with <sup>n</sup> rendered as superscript digits and <i>n</i> badges as circled numbers."""
    el = BeautifulSoup(str(el), "html.parser")
    for sp in el.find_all("sup"):
        sp.replace_with("".join("⁰¹²³⁴⁵⁶⁷⁸⁹"[int(c)] if c.isdigit() else c for c in sp.get_text()))
    for i in el.find_all("i"):
        if i.get_text().strip().isdigit():
            i.replace_with("①②③④⑤"[int(i.get_text().strip()) - 1] + " ")
    return text(el)


def table_md(tbl):
    rows = []
    for tr in tbl.find_all("tr"):
        cells = [sup_text(c).replace("|", "/") for c in tr.find_all(["th", "td"])]
        rows.append("| " + " | ".join(cells) + " |")
    head = rows[0]
    sep = "|" + "|".join("---" for _ in tbl.find("tr").find_all(["th", "td"])) + "|"
    return "\n".join([head, sep] + rows[1:])


lines = ["# OKX DEX BD Affiliate (Indonesia) take-home: deck text v2 (6 slides)", "",
         "Text export of `deck_v2.html` (the designed deck): on-slide content, footer, then speaker notes. "
         "Every number ties to `OKX_DEX_BD_Affiliate_Indonesia_Prospect_Spreadsheet.xlsx`.", ""]
for n, wrap in enumerate(soup.select("section.slide-wrap"), 1):
    frame = wrap.select_one(".frame")
    lines += [f"## Slide {n}: {text(frame.h2)}", f"*{text(frame.select_one('.kicker')).upper()}*", ""]
    body = frame.select_one(".s-body")
    for el in body.find_all(recursive=False):
        cls = el.get("class", [])
        if "tscroll" in cls:
            lines += [table_md(el.find("table")), ""]
        elif "steps" in cls:
            for row in el.select(".steprow"):
                lines.append("Steps: " + " → ".join(f"{sup_text(s.b)} ({sup_text(s).replace(sup_text(s.b), '', 1).strip()})" for s in row.select(".step")))
            ban = el.select_one(".banner")
            lines += ["", f"**{text(ban)}**", ""]
        elif "why" in cls:
            chain = el.select_one(".chain")
            parts = [sup_text(c) for c in chain.select(".c")]
            lines += [f"**{text(chain.select_one('.lbl'))}** (filter, in order): " + " → ".join(parts), ""]
            picks = el.select(".pick")
            hdrs = [text(p.select_one(".hd b")) + " · " + text(p.select_one(".score")).lower() for p in picks]
            lines += ["| | " + " | ".join(hdrs) + " |", "|---|" + "---|" * len(picks)]
            dts = [sup_text(d) for d in picks[0].select("dt")]
            for k, dt in enumerate(dts):
                lines.append(f"| {dt} | " + " | ".join(text(p.select("dd")[k]) for p in picks) + " |")
            lines.append("| ④ Segment and targets | " + " | ".join(sup_text(p.select_one(".kpi")).replace("④ ", "") for p in picks) + " |")
            tail = el.select_one(".tail")
            if tail:
                lines += ["", sup_text(tail).replace(tail.b.get_text(), f"**{tail.b.get_text()}**", 1)]
            lines.append("")
        elif "formula" in cls or "para" in cls:
            lines += [sup_text(el), ""]
        elif "cards" in cls or "s6" in cls:
            for card in el.select(".card"):
                lines.append("**[card]**")
                lines.append(f"**{sup_text(card.h3)}**")
                stat = card.select_one(".stat")
                if stat:
                    lines.append(f"{text(stat.select_one('.big'))}: {text(stat.p)}")
                for li in card.select("li"):
                    lines.append(f"- {sup_text(li)}")
                for p in card.find_all("p", recursive=False):
                    lines.append(sup_text(p))
                lines.append("")
        elif "card" in cls:
            lines.append("**[card]**")
            for p in el.find_all("p"):
                lines.append(sup_text(p))
            lines.append("")
    lines += [f"**Footer:** {sup_text(frame.select_one('.foot'))}", "", "**Speaker notes:**", ""]
    for el in wrap.select(".notes .body > *"):
        if el.name == "blockquote":
            lines += [f"> {text(el)}", ""]
        else:
            lines += [text(el), ""]
    lines += ["---", ""]
md = "\n".join(lines)
assert "—" not in md
out.write_text(md)
print("written", len(md))
