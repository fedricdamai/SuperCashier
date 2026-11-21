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
T2A = {r: ans(t2, L2["ANS0"] + r, 2, 8) for r in range(7)}
T4A = {r: ans(t4, L4["ANS0"] + r, 2, 6) for r in range(6)}
W = {k_: L4[k_] for k_ in ("LIVE", "ACT", "FTD", "TPD", "SW8", "SWP", "REV", "REVN", "RUN", "W0", "FOCUS")}
ftd = [t1.cell(N0 + i, 6).value for i in range(8)]
swaps = [t1.cell(N0 + i, 16).value for i in range(8)]
top3_ftd, top3_sw = sum(ftd[:3]), sum(swaps[:3])
T = dict(ftd=t1.cell(TOT, 6).value, fnf=t1.cell(TOT, 8).value, swaps=t1.cell(TOT, 16).value, rev=t1.cell(TOT, 18).value)
CASES = ("50% of forecast", "Forecast", "Bonus threshold", "A volume you choose")
SC = lambda pn, j, c: t2.cell(L2["SC0"] + 4 * "ABC".index(pn) + j, c).value  # j: 0 50%, 1 forecast, 2 threshold, 3 chosen volume
IN = lambda key, pn: t2.cell(L2[key], 2 + "ABC".index(pn)).value
wk = {c: [t4[f"{c}{W[x]}"].value for x in ("LIVE", "ACT", "FTD", "TPD")] for c in "BCDE"}
day30_sw = t4[f"E{W['SW8']}"].value + t4[f"E{W['SWP']}"].value
I0 = L3["ITEMS0"]
T3 = [(t3.cell(I0 + i, 1).value, t3.cell(I0 + i, 2).value, t3.cell(I0 + i, 3).value) for i in range(8)]
PLAN = [(t3.cell(L3["PLAN0"] + i, 1).value, t3.cell(L3["PLAN0"] + i, 3).value, t3.cell(L3["PLAN0"] + i, 2).value) for i in range(6)]
O = lambda j: t3.cell(L3["OFFER0"] + j, 2).value
k_earn1, k_earnt, k_keept = (t3.cell(L3[x], 2).value for x in ("EARN1", "EARNT", "KEEPT"))
assert t1.cell(N0, 1).value.startswith("Kevin") and t1.cell(N0 + 2, 1).value.startswith("0xRusdi")
names = ["Kevin Susanto", "Bevan / Narrative Labs", "0xRusdi + CVC network", "Brandon Salim + a79", "GuarEmperor + FnF circle",
         "Crypto Ndo", "DegenRex", "Meteora LP Army"]
mm = lambda x: f"${x / 1e6:,.0f}M" if x % 1e6 == 0 else m(x)

H = [h1("Indonesia DEX Affiliate Plan – Fedric Intan Damai")]
H.append(h2("Executive summary"))
H.append(ul([
    f"<b>Partners:</b> Kevin Susanto, Bevan and 0xRusdi first: public proof they trade, plus their own power traders. "
    f"{top3_ftd:.0f} of {T['ftd']:.0f} FTDs and {m(top3_sw)} of {m(T['swaps'])} in first-month swaps.",
    f"<b>Deal:</b> pilot C first once its volume is confirmed, on a growth deal: its share as asked, its {usd(IN('BONUS', 'C'))} fee paid "
    f"only once its swaps pass {mm(IN('THR', 'C'))}, and 50% above that. OKX makes {usd(SC('C', 1, 10))} at forecast instead of losing "
    f"{usd(-SC('C', 1, 6))}.",
    f"<b>Pitch:</b> a share of every swap Kevin's followers make, and his own VIP-style tiers once his link earns OKX "
    f"{usd(O(2))} of fees in a month.",
    f"<b>30 days:</b> {wk['E'][0]} partners live plus pilot C, {wk['E'][1]} activated, about {wk['E'][2]:,.0f} FTDs and "
    f"{wk['E'][3]:,.0f} trades a day, and {m(day30_sw)} of swaps by day 30.",
]))

# ---------------------------------------------------------------- Task 1 (layout of the user's edited write-up)
H.append(h2("Task 1: Find partners beyond airdrop communities"))
rows = []
for i in range(8):
    r = A0 + i
    us = [u.strip().replace("https://", "") for u in str(t1.cell(r, 3).value).split(" ; ")]
    links = "\n".join(E(u) for u in (us[:2] if "CVC" in str(us) else us[:1]))
    n = N0 + i
    rows.append([f"<b>{i + 1}. {E(names[i])}</b>\n{links}\n\n{E(t1.cell(r, 11).value)}", *[E(str(t1.cell(r, c).value)) for c in (4, 5, 6, 7, 8)],
                 f"<b>{t1.cell(r, 9).value}:</b> {E(str(t1.cell(r, 10).value))}\n<i>{t1.cell(n, 6).value:.0f} FTDs · {k(t1.cell(n, 16).value)}</i>"])
H.append(table(["Partner and public link", "Audience and market", "Evidence of fit", "Acquisition route", "Proposed partnership",
                "Main risk", "Priority and forecast"], rows, [15, 14, 16, 13, 15, 14, 13]))
H.append(p("people reached × 2% sign up × 15% trade = FTDs, each swapping ~$2,000 in month one, plus FnF power traders from day 4.",
           "Forecast (italics in the table):", small=True))
H.append(p(f"8 affiliates, {T['ftd']:.0f} FTDs, {T['fnf']} FnF power traders, {m(T['swaps'])} of swaps monthly, {usd(T['rev'])} at 0.04%.",
           "Total:", small=True))
H.append(p(E(T1A[1][1]), "Why these 3 first:"))
H.append(ul([bold_lead(x) for x in T1A[1][2]]))
H.append(p("From here, there are possibilities of automatic pickups by other colleagues and similar traders, with additional impact, as "
           "they're also the thought leaders on the DEX side of trading."))

# ---------------------------------------------------------------- Task 2
H.append(h2("Task 2: Make a commercial decision"))
askt = {pn: (f"{usd(IN('FEE', pn))} + " if IN("FEE", pn) else "") + f"{IN('ASK', pn):.0%}" + ("" if IN("FEE", pn) else ", no fee") for pn in "ABC"}
deal = {pn: (f"{IN('SHARE', pn):.0%} share; {usd(IN('BONUS', pn))} paid once swaps pass {mm(IN('THR', pn))}; {IN('UP', pn):.0%} above"
             if IN("BONUS", pn) else f"{IN('SHARE', pn):.0%} share, no fee (its own terms)") for pn in "ABC"}
beats = {pn: (mm(IN("THR", pn)) if IN("BONUS", pn) else "Same terms") for pn in "ABC"}
order = sorted("ABC", key=lambda pn: IN("ORDER", pn))
rows = [[pn, askt[pn], usd(SC(pn, 1, 6)), deal[pn], usd(SC(pn, 0, 10)), usd(SC(pn, 1, 10)), beats[pn]] for pn in order]
H.append(table(["Partner", "Their ask", "Month 1 at their ask", "Our growth deal", "Ours at 50%", "Ours at forecast",
                "Pays them more than their ask above"], rows, [8, 12, 11, 33, 10, 11, 15], num_from=4))
labels = ["1. First-month contribution:", "2. Preferred pilot:", "3. Counter-offer, the growth deal:",
          "Why it pays them more than a retainer:", "4. Evidence before signing:", "5. If C reaches only 50% of forecast:",
          "Other costs, and what we do about them:"]
for i, lab in enumerate(labels):
    q, a, d = T2A[i]
    H.append(p(E(a), lab))
    if i in (0, 2, 6):
        H.append(ul([bold_lead(x) for x in d]))
    elif i in (1, 3, 4):
        H.append(p("; ".join(bold_lead(x) for x in (d[:3] if i == 4 else d)) + ".", small=True))
    elif i == 5:
        H.append(p("; ".join(bold_lead(x) for x in d) + ".", "Then find out why.", small=True))
H.append(p("month 1 = revenue at 0.04% − share − bonus (paid only in a month above the threshold).", "Method:", small=True))

# ---------------------------------------------------------------- Task 3
H.append(h2("Task 3: Show how you would win and activate a partner"))
H.append(p(f"Kevin Susanto, our #1. No retainer. He starts on OKX DEX's standard referral share ({O(1):.0%}, rising with OKX's published "
           f"tiers). Once his link earns OKX {usd(O(2))} of fees in a month (about {mm(O(3))} of memecoin swaps at 0.5%), his own "
           f"VIP-style tiers start: {O(4):.0%}, then {O(6):.0%} at {mm(O(5))} and {O(8):.0%} at {mm(O(7))}, plus a co-hosted live "
           f"session and competition. At the target he earns {usd(k_earnt)} a month and OKX keeps {usd(k_keept)}.", "Partner and offer:"))
msg = T3[0][2]
H.append(p(E(T3[0][1]), f"First message ({len(msg.split())} words; Bahasa version in the spreadsheet):"))
H.append(quote([E(x) for x in msg.split("\n\n")]))
H.append(p(E(T3[2][1]), "Follow-up, day 4:"))
H.append(quote([E(T3[2][2])]))
H.append(p(E(T3[3][1]), "If he says another platform pays more:"))
H.append(quote([E(T3[3][2])]))
H.append(p("", "First-meeting questions (his needs, how we help, the lightest way to run it):"))
H.append(ol([E(T3[4 + i][2]) for i in range(3)]))
H.append(table(["Days", "Activation plan (day 1 = agreement signed)", "Target"], [[a, E(b), E(c)] for a, b, c in PLAN], [10, 64, 26]))
H.append(p("the referral share and follower discount are published OKX DEX referral features; his personal tiers, the co-hosted "
           "session and the launch terms are offered subject to OKX approval.", "No unverified promises:", small=True))

# ---------------------------------------------------------------- Task 4
H.append(h2("Task 4: Demonstrate practical execution and persistence"))
H.append(p(E(T4A[0][1].replace(" (section A)", "")), "Answer:"))
focus = [t4.cell(W["FOCUS"], 2 + j).value for j in range(4)]
w0 = W["W0"]
rows = []
for j, (c, lab) in enumerate(zip("BCDE", ("1 · days 1-7", "2 · days 8-14", "3 · days 15-21", "4 · days 22-30"))):
    rows.append([lab, E(focus[j]), t4[f"{c}{w0 + 1}"].value, t4[f"{c}{w0 + 3}"].value, f"{wk[c][0]} + C" if j else "0",
                 wk[c][1], f"{wk[c][2]:,.0f}", f"{wk[c][3]:,.0f}"])
H.append(table(["Week", "Focus", "Messages", "Meetings", "Live", "Activated", "FTDs", "Trades a day"], rows,
               [12, 40, 9, 9, 7, 9, 6, 8], num_from=2))
sw8, swp, rev30, revn30 = (t4.cell(W[x], 5).value for x in ("SW8", "SWP", "REV", "REVN"))
H.append(p(f"Swaps by day 30: {m(sw8)} from our 8 and {m(swp)} from pilot C: {usd(rev30)} of revenue at 0.04% ({usd(revn30)} after "
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
