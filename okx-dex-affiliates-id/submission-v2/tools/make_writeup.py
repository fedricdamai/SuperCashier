"""Build the write-up blocks (JSON for build_writeup.js) from the recalculated submission workbook.

Usage: python make_writeup.py <submission.xlsx> <out.json>
Every number and every one-sentence answer comes from the workbook, so the write-up and the sheet always match.
"""
import html
import json
import sys
from pathlib import Path

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).parent))
from layout_v6 import MI, T1 as L1, T2 as L2, T3 as L3, T4 as L4  # noqa: E402

wb = load_workbook(sys.argv[1], data_only=True)
out = Path(sys.argv[2])
t1, t2, t3, t4, mi = (wb[n] for n in ("Task 1", "Task 2", "Task 3", "Task 4", "Miscellaneous"))
E = html.escape


def k(x):
    return f"${x / 1000:,.0f}K"


def m(x):
    return f"${x / 1e6:.2f}M"


def usd(x):
    return f"-${-x:,.0f}" if x < 0 else f"${x:,.0f}"


def h1(t): return dict(type="h1", text=t)
def h2(t): return dict(type="h2", text=t)
def p(txt, label=None, small=False): return dict(type="p", label=label, html=txt, small=small)
def ul(items, small=False): return dict(type="ul", items=items, small=small)
def ol(items): return dict(type="ol", items=items)
def quote(lines_): return dict(type="quote", lines=lines_)
def table(head, rows, widths, num_from=None, bold_last=False):
    return dict(type="table", head=head, rows=[[str(c) for c in r] for r in rows], widths=widths, num_from=num_from, bold_last=bold_last)


def ans(ws, row, b_col, c_col):
    """(question, one-sentence answer, detail bullets) from an Answers row."""
    q, a = ws.cell(row, 1).value, ws.cell(row, b_col).value
    d = ws.cell(row, c_col).value or ""
    return q, a, [x.lstrip("• ").strip() for x in d.split("\n") if x.strip()]


def bold_lead(s):
    """'A: text' -> '<b>A:</b> text' (labels before the first colon)."""
    if ": " in s and len(s.split(": ", 1)[0]) <= 40:
        h, rest = s.split(": ", 1)
        return f"<b>{E(h)}:</b> {E(rest)}"
    return E(s)


N0, TOT, A0 = L1["N0"], L1["TOT"], L1["A0"]
T1A = {r: ans(t1, L1["ANS0"] + r, 2, 5) for r in range(5)}
T2A = {r: ans(t2, L2["ANS0"] + r, 2, 8) for r in range(8)}
T4A = {r: ans(t4, L4["ANS0"] + r, 2, 6) for r in range(6)}
W = {k_: L4[k_] for k_ in ("LIVE", "ACT", "FTD", "TPD", "SW8", "SWP", "REV", "REVN", "RUN", "W0", "FOCUS")}
ftd = [t1.cell(N0 + i, 6).value for i in range(8)]
swaps = [t1.cell(N0 + i, 16).value for i in range(8)]
top3_ftd, top3_sw = sum(ftd[:3]), sum(swaps[:3])
T = dict(ftd=t1.cell(TOT, 6).value, fnf=t1.cell(TOT, 8).value, swaps=t1.cell(TOT, 16).value, rev=t1.cell(TOT, 18).value)
chk = dict(noair=t1.cell(L1["CHK"], 5).value, meet=t1.cell(L1["CHK"], 7).value, types=t1.cell(L1["CHK"], 9).value)
SC = lambda pn, j, c: t2.cell(L2["SC0"] + 4 * "ABC".index(pn) + j, c).value  # j: 0 50%, 1 100%, 2 200%, 3 300%
wk = {c: [t4[f"{c}{W[x]}"].value for x in ("LIVE", "ACT", "FTD", "TPD")] for c in "BCDE"}
day30_sw = t4[f"E{W['SW8']}"].value + t4[f"E{W['SWP']}"].value
I0 = L3["ITEMS0"]
T3 = [(t3.cell(I0 + i, 1).value, t3.cell(I0 + i, 2).value, t3.cell(I0 + i, 3).value) for i in range(8)]
PLAN = [(t3.cell(L3["PLAN0"] + i, 1).value, t3.cell(L3["PLAN0"] + i, 3).value, t3.cell(L3["PLAN0"] + i, 2).value) for i in range(6)]
k_share, k_bonus = t3.cell(L3["SHARE"], 2).value, t3.cell(L3["BONUS"], 2).value
assert t1.cell(N0, 1).value.startswith("Kevin") and t1.cell(N0 + 2, 1).value.startswith("Ponyin")
names = ["Kevin Susanto", "Bevan / Narrative Labs", "Ponyin + CVC network", "GuarEmperor + FnF circle", "Brandon Salim + a79",
         "Crypto Ndo", "DegenRex", "Meteora LP Army"]

H = [h1("Indonesia DEX Affiliate Plan"),
     p("OKX P4 DEX BD Affiliate take-home · Fedric Intan Damai · November 2026. Companion files: a 6-slide deck and the prospect "
       "spreadsheet, where every number below is a live formula."),
     p("FTD = first-time trader (a referred user's first trade). FnF = friends and family: the power traders a partner brings "
       "personally. DEX = decentralised exchange. KOL = key opinion leader. LP = liquidity provider. SOL = Solana's token. "
       "USDT = Tether's US-dollar stablecoin.", "Terms:", small=True)]

H.append(h2("Executive summary"))
H.append(ul([
    f"<b>Partners:</b> Kevin Susanto, Bevan and Ponyin first: public proof they trade, plus their own power traders. "
    f"{top3_ftd:.0f} of {T['ftd']:.0f} FTDs and {m(top3_sw)} of {m(T['swaps'])} in first-month swaps.",
    f"<b>Deal:</b> pilot B first, on a growth deal: rising share, a milestone cash bonus, OKX-funded campaigns. Only B makes money at "
    f"its own ask ({usd(SC('B', 1, 16))}), even at half volume.",
    "<b>Pitch:</b> one benefit up front for Kevin: a share of every swap his followers make, next to his existing OKX link.",
    f"<b>30 days:</b> {wk['E'][0]} partners live plus pilot B, {wk['E'][1]} activated, about {wk['E'][2]:,.0f} FTDs and "
    f"{wk['E'][3]:,.0f} trades a day, and {m(day30_sw)} of swaps by day 30.",
]))

# ---------------------------------------------------------------- Task 1
H.append(h2("Task 1: Find partners beyond airdrop communities"))
H.append(p(E(T1A[0][1].replace("in section A", "below")), "Answer:"))
rows = []
for i in range(8):
    r = A0 + i
    url = str(t1.cell(r, 3).value).split(" ; ")[0].replace("https://", "")
    n = N0 + i
    rows.append([f"<b>{i + 1}. {E(names[i])}</b>\n{E(url)}", *[E(str(t1.cell(r, c).value)) for c in (4, 5, 6, 7, 8)],
                 f"<b>{t1.cell(r, 9).value}:</b> {E(str(t1.cell(r, 10).value))}\n<i>{t1.cell(n, 6).value:.0f} FTDs · {k(t1.cell(n, 16).value)}</i>"])
H.append(table(["Partner and public link", "Audience and market", "Evidence of fit", "Acquisition route", "Proposed partnership",
                "Main risk", "Priority and forecast"], rows, [15, 14, 16, 13, 15, 14, 13]))
H.append(p(f"{chk['noair']} of 8 not airdrop-led, {chk['types']:.0f} partner types, {chk['meet']} meetable in person. GuarEmperor (#4) trades "
           f"actively and calls token buys, but his feed mixes those calls with whitelist and mint posts, so he starts in the second wave.", small=True))
H.append(p(E(T1A[1][1]), "Why these 3 first:"))
H.append(ul([bold_lead(x) for x in T1A[1][2]]))
H.append(p(E(T1A[3][1]), "How we pick the FnF power traders:"))
H.append(ul([bold_lead(x) for x in T1A[3][2][:2]]))
H.append(p(f"people reached × 2% sign up × 15% trade = FTDs, each swapping ~$2,000 in month one, plus FnF power traders from "
           f"day 4. All 8: {T['ftd']:.0f} FTDs, {T['fnf']} FnF power traders, {m(T['swaps'])} of swaps, {usd(T['rev'])} at 0.04%. "
           f"Rates from three Indonesian referral books.", "Forecast (italics in the table):", small=True))

# ---------------------------------------------------------------- Task 2
H.append(h2("Task 2: Make a commercial decision"))
deal = {"A": "20% share only; later", "B": "30% rising to 45%; 10% milestone bonus; $250 competition",
        "C": "$500 + $500 retainer; 15% rising to 30%; bonus; $300 event"}
askt = {"A": "$2,000 + 20%", "B": "35%, no fee", "C": "$3,000 + 10%"}
rows = [[pn, askt[pn], usd(SC(pn, 1, 16)), deal[pn], usd(SC(pn, 0, 14)), usd(SC(pn, 1, 14)), usd(SC(pn, 2, 14)), usd(SC(pn, 3, 14))]
        for pn in "ABC"]
H.append(table(["Partner", "Their ask", "Month 1 at their ask", "Our growth deal", "Ours at 50%", "Ours at forecast", "At 2x", "At 3x"],
               rows, [8, 12, 11, 33, 9, 10, 8, 9], num_from=4))
for i, lab in enumerate(["1. First-month contribution:", "2. Preferred pilot:", "3. Counter-offer:", "To make it more attractive at no fixed cost:",
                         "4. Evidence before signing:", "5. If B reaches only 50% of forecast:", "Is the campaign spend worth it?",
                         "Other costs to investigate:"]):
    q, a, d = T2A[i]
    H.append(p(E(a), lab))
    if i in (0, 2, 4):
        H.append(ul([bold_lead(x) for x in (d[:3] if i == 4 else d)]))
    elif i == 1:
        H.append(p("; ".join(bold_lead(x) for x in d) + ".", small=True))
    elif i == 5:
        H.append(p("; ".join(bold_lead(x) for x in d) + ".", "Then find out why.", small=True))
H.append(p("month 1 = revenue at 0.04% − retainer − share − bonus − campaign spend; brackets are marginal.", "Method:", small=True))

# ---------------------------------------------------------------- Task 3
H.append(h2("Task 3: Show how you would win and activate a partner"))
H.append(p(f"Kevin Susanto, our #1. No retainer; fee share 30% to $250K of monthly swaps, 35% to $500K, 40% above; a 10% cash "
           f"bonus at each milestone within 7 days; a co-hosted live session and competition. Month one: ~{usd(k_share)} share "
           f"plus {usd(k_bonus)} bonus.", "Partner and offer:"))
msg = T3[0][2]
H.append(p(E(T3[0][1]), f"First message ({len(msg.split())} words; Bahasa version in the spreadsheet):"))
H.append(quote([E(x) for x in msg.split("\n\n")]))
H.append(p(E(T3[2][1]), "Follow-up, day 4:"))
H.append(quote([E(T3[2][2])]))
H.append(p(E(T3[3][1]), "If he says another platform pays more:"))
H.append(quote([E(x) for x in T3[3][2].split("\n")]))
H.append(p("", "First-meeting questions:"))
H.append(ol([E(T3[4 + i][2]) for i in range(3)]))
H.append(table(["Days", "Activation plan (day 1 = agreement signed)", "Target"], [[a, E(b), E(c)] for a, b, c in PLAN], [10, 64, 26]))
H.append(p("the fee share and follower discount are published OKX DEX referral features; the cash bonus, the co-hosted session "
           "and the launch-wave terms are offered subject to OKX approval.", "No unverified promises:", small=True))

# ---------------------------------------------------------------- Task 4
H.append(h2("Task 4: Demonstrate practical execution and persistence"))
H.append(p(E(T4A[0][1].replace(" (section A)", "")), "Answer:"))
focus = [t4.cell(W["FOCUS"], 2 + j).value for j in range(4)]
w0 = W["W0"]
rows = []
for j, (c, lab) in enumerate(zip("BCDE", ("1 · days 1-7", "2 · days 8-14", "3 · days 15-21", "4 · days 22-30"))):
    rows.append([lab, E(focus[j]), t4[f"{c}{w0 + 1}"].value, t4[f"{c}{w0 + 3}"].value, f"{wk[c][0]} + B" if j else "0",
                 wk[c][1], f"{wk[c][2]:,.0f}", f"{wk[c][3]:,.0f}"])
H.append(table(["Week", "Focus", "Messages", "Meetings", "Live", "Activated", "FTDs", "Trades a day"], rows,
               [12, 40, 9, 9, 7, 9, 6, 8], num_from=2))
sw8, swp, rev30, revn30 = (t4.cell(W[x], 5).value for x in ("SW8", "SWP", "REV", "REVN"))
H.append(p(f"Swaps by day 30: {m(sw8)} from our 8 and {m(swp)} from pilot B: {usd(rev30)} of revenue at 0.04% ({usd(revn30)} after "
           f"partner shares).", small=True))
for i, lab in ((1, "In-person prospecting day:"), (2, "If 20 messages bring 2 replies and no meetings:"), (3, "Day 30, per partner:"),
               (4, "Activated partner:"), (5, "How the numbers land inside 30 days:")):
    q, a, d = T4A[i]
    H.append(p(E(a), lab))
    if i in (1, 2):
        H.append(p("; ".join(bold_lead(x) for x in d) + ".", small=True))

H.append(h2("Commitment"))
H.append(p("<b>Earliest start:</b> in 2 weeks. <b>Full-time:</b> yes, from 2 weeks. <b>Other commitments:</b> marketer role at "
           "Surgence Labs. <b>In person:</b> available across Indonesia and Southeast Asia, and willing to attend partner meetings "
           "and events in person."))


# ---------------------------------------------------------------- appendix: sources and resources
H.append(dict(type="pagebreak"))
H.append(h2("How I worked and sources"))
H.append(p("I set the strategy, the selection logic and every judgment call. AI gathered public data, automated repetitive work and "
           "fact-checked. I re-checked every part myself and ran independent fact-checks with several AI agents. Public information only."))
H.append(ul([
    "<b>Screening:</b> 100 public Indonesian prospects scored on one filter: proof that they trade, the power traders they can bring, "
    "rival flow we can move, and fees from memecoin traders. 8 made the list; the rest are in the research workbook with the reason each was cut.",
    "<b>Measuring:</b> follower counts and median views on recent posts through the X API; Bevan's wallet on Solscan; Discord invite "
    "counts (members and members online); Telegram channel previews; event organiser posts (20-21 November 2026).",
    "<b>Modelling:</b> one formula for every partner and every hypothetical. Rates come from three Indonesian referral books "
    "(@0xSarimin, @0xRusdi, @AiSlimx) and Bevan's published wallet; every assumption is a yellow input, so the numbers move when one changes.",
    "<b>Checking:</b> each figure traced to its public source below; AI fact-checks run independently by several agents, then re-checked by hand.",
]))
groups = {}
r = MI["SRC0"]
while mi.cell(r, 1).value:
    groups.setdefault(mi.cell(r, 1).value, []).append((mi.cell(r, 2).value, mi.cell(r, 3).value))
    r += 1
for g, items in groups.items():
    H.append(dict(type="links", group=g, items=[dict(label=a, url=u) for a, u in items]))

out.write_text(json.dumps(H, indent=1, ensure_ascii=False))
print("blocks", len(H), "sources", sum(len(v) for v in groups.values()))
