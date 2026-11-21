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
from layout_v6 import MI, T1 as L1, T2 as L2, T3 as L3, T4 as L4, GT as GTL  # noqa: E402

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
T4A = {r: ans(t4, L4["ANS0"] + r, 2, 6) for r in range(6)}
W = {k_: L4[k_] for k_ in ("LIVE", "ACT", "FTD", "TPD", "SW8", "REV", "REVN", "RUN", "W0", "FOCUS", "PIL0")}
ftd = [t1.cell(N0 + i, 6).value for i in range(8)]
swaps = [t1.cell(N0 + i, 16).value for i in range(8)]
top3_ftd, top3_sw = sum(ftd[:3]), sum(swaps[:3])
T = dict(ftd=t1.cell(TOT, 6).value, fnf=t1.cell(TOT, 8).value, swaps=t1.cell(TOT, 16).value, rev=t1.cell(TOT, 18).value)
SCv = lambda pn, j, c: t2.cell(L2["SC0"] + 4 * "ABC".index(pn) + j, c).value  # j: 0 50%, 1 forecast, 2 $10M, 3 $20M
IN = lambda key, pn: t2.cell(L2[key], 2 + "ABC".index(pn)).value
minsh = t2.cell(L2["MINSH"], 2).value
wk = {c: [t4[f"{c}{W[x]}"].value for x in ("LIVE", "ACT", "FTD", "TPD")] for c in "BCDE"}
sw8_30 = t4.cell(W["SW8"], 5).value
pil30 = t4.cell(W["PIL0"] + 1, 5).value
I0 = L3["ITEMS0"]
T3 = [(t3.cell(I0 + i, 1).value, t3.cell(I0 + i, 2).value, t3.cell(I0 + i, 3).value) for i in range(8)]
PLAN = [(t3.cell(L3["PLAN0"] + i, 1).value, t3.cell(L3["PLAN0"] + i, 3).value, t3.cell(L3["PLAN0"] + i, 2).value) for i in range(6)]
O = lambda j: t3.cell(L3["OFFER0"] + j, 2).value
k_earn1, k_earnt, k_keept = (t3.cell(L3[x], 2).value for x in ("EARN1", "EARNT", "KEEPT"))
assert t1.cell(N0, 1).value.startswith("Kevin") and t1.cell(N0 + 2, 1).value.startswith("0xRusdi") and t1.cell(N0 + 5, 1).value.startswith("Michael")
names = ["Kevin Susanto", "Bevan / Narrative Labs", "0xRusdi + CVC network", "Brandon Salim + a79", "GuarEmperor + FnF circle",
         "Michael Wyann / Laba Lab", "Crypto Ndo", "DegenRex"]
mm = lambda x: f"${x / 1e6:,.0f}M" if x % 1e6 == 0 else m(x)

H = [h1("Indonesia DEX Affiliate Plan – Fedric Intan Damai")]
H.append(h2("Executive summary"))
H.append(ul([
    f"<b>Partners:</b> Kevin Susanto, Bevan and 0xRusdi first: public proof they trade, plus their own power traders. "
    f"{top3_ftd:.0f} of {T['ftd']:.0f} FTDs and {m(top3_sw)} of {m(T['swaps'])} in first-month swaps.",
    f"<b>Deal:</b> pilot C first once its volume is confirmed: a smaller share plus milestone payouts, so OKX keeps at least "
    f"{minsh:.0%} of revenue in every case, and {usd(SCv('C', 1, 10))} at forecast instead of losing {usd(-SCv('C', 1, 6))}.",
    f"<b>Pitch:</b> a share of every swap Kevin's followers make, and his own VIP-style tiers once his link earns OKX "
    f"{usd(O(2))} of fees in a month.",
    f"<b>30 days:</b> {wk['E'][0]} partners live, {wk['E'][1]} activated, about {wk['E'][2]:,.0f} FTDs, {wk['E'][3]:,.0f} trades a day "
    f"and {m(sw8_30)} of swaps from our 8; pilot C adds {mm(pil30)} on its own track.",
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

# ---------------------------------------------------------------- Task 2: tables, one column per partner
H.append(h2("Task 2: Make a commercial decision"))
heads = ["", "A · airdrop community", "B · repeat on-chain traders", "C · trading educator"]
cell_ = lambda pn, j: f"OKX {usd(SCv(pn, j, 10))} · partner {usd(SCv(pn, j, 9))} ({SCv(pn, j, 11):.0%} to OKX)"
rows = [["<b>Their ask, at forecast</b>"] + [f"OKX {usd(SCv(pn, 1, 6))} · partner {usd(SCv(pn, 1, 5))}" for pn in "ABC"]]
for j, lab in enumerate(("Our deal at 50% of forecast", "Our deal at forecast", "Our deal at $10M a month", "Our deal at $20M a month")):
    rows.append([f"<b>{lab}</b>"] + [cell_(pn, j) for pn in "ABC"])
H.append(table(["Month 1"] + heads[1:], rows, [22, 26, 26, 26]))
qa = []
for i in (0, 1, 2, 4, 5):
    rw = L2["ANS0"] + i
    q, s1 = t2.cell(rw, 1).value, t2.cell(rw, 2).value
    qa.append([f"<b>{E(q)}</b>\n<i>{E(s1)}</i>"] + [E(str(t2.cell(rw, c).value)) for c in (6, 10, 14)])
H.append(table(["Question and answer"] + heads[1:], qa, [28, 24, 24, 24]))
rw = L2["ANS0"] + 6
H.append(p(E(t2.cell(rw, 2).value), "Other costs, and what we do about them:"))
oc = str(t2.cell(rw, 6).value).split(". ")
H.append(ul([bold_lead(x.rstrip(".")) for x in oc]))
H.append(p("month 1 = revenue at 0.04% − share − milestone payout (paid only in a month that reaches it).", "Method:", small=True))

# ---------------------------------------------------------------- Task 3: one table
H.append(h2("Task 3: Show how you would win and activate a partner"))
offer = (f"No retainer. Starts on OKX DEX's standard referral share ({O(1):.0%}). Once his link earns OKX {usd(O(2))} of fees in a month "
         f"(about {mm(O(3))} of memecoin swaps at 0.5%), his own VIP-style tiers start: {O(4):.0%}, then {O(6):.0%} at {mm(O(5))} and "
         f"{O(8):.0%} at {mm(O(7))}; plus a co-hosted live session and competition. At the target he earns {usd(k_earnt)} a month and OKX keeps {usd(k_keept)}.")
msg = T3[0][2]
qs = "\n".join(f"{i + 1}. {E(T3[4 + i][2])}" for i in range(3))
rows = [["<b>Partner and offer</b>", "Kevin Susanto, our #1: standard share now, his own tiers once his link proves itself", E(offer)],
        [f"<b>First message</b>\n{len(msg.split())} words", E(T3[0][1]), "<i>" + "</i>\n<i>".join(E(x) for x in msg.split("\n\n")) + "</i>"],
        ["<b>Follow-up, day 4</b>", E(T3[2][1]), f"<i>{E(T3[2][2])}</i>"],
        ["<b>If another platform pays more</b>", E(T3[3][1]), f"<i>{E(T3[3][2])}</i>"],
        ["<b>First-meeting questions</b>", "His needs, how we help, and the lightest way to run it", qs]]
H.append(table(["Item", "In one sentence", "Text"], rows, [14, 19, 67]))
H.append(table(["Days", "Activation plan (day 1 = agreement signed)", "Target"], [[a, E(b), E(c)] for a, b, c in PLAN], [10, 64, 26]))
H.append(p("the referral share and follower discount are published OKX DEX referral features; his personal tiers, the co-hosted "
           "session and the launch terms are offered subject to OKX approval. Bahasa version of the message: spreadsheet, Task 3.",
           "No unverified promises:", small=True))

# ---------------------------------------------------------------- Task 4: Gantt and KPIs
H.append(h2("Task 4: Demonstrate practical execution and persistence"))
gt = wb["Task 4 Gantt"]
G0 = GTL["G0"]
blocks = [(d, d + 2) for d in range(1, 31, 3)]
rows = []
r = G0
while gt.cell(r, 1).value:
    name, track, kpi, s0, s1 = (gt.cell(r, c).value for c in range(1, 6))
    mark = "¤" if track == "Pilot C" else "§"
    rows.append([("<b>Pilot C:</b> " + E(name.split(": ", 1)[1]) if track == "Pilot C" else E(name)), E(str(kpi))] +
                [mark if (s0 <= b1 and s1 >= b0) else "" for b0, b1 in blocks])
    r += 1
H.append(dict(type="table", head=["Activity (dark: our 8 partners; light: pilot C, its own track)", "KPI"] + [f"{b0}-{b1}" for b0, b1 in blocks],
              rows=rows, widths=[33, 23] + [4.4] * 10, num_from=None, bold_last=False, size=14))
K0 = GTL["KPI0"]
krow = []
for i in range(8):
    rr = K0 + i
    lab = gt.cell(rr, 1).value
    vals = [gt.cell(rr, c).value for c in range(2, 6)]
    money = "swaps" in lab.lower() or "revenue" in lab.lower()
    krow.append([f"<b>{E(lab)}</b>"] + [(m(v) if v >= 1e6 else (k(v) if v >= 1e4 else usd(v))) if money else f"{v:,.0f}" for v in vals])
H.append(table(["KPI", "Day 7", "Day 14", "Day 21", "Day 30"], krow, [40, 15, 15, 15, 15], num_from=1))
for i, lab in ((1, "In-person prospecting day:"), (2, "If 20 messages bring 2 replies and no meetings:"), (3, "Day 30, per partner:"),
               (4, "Activated partner:")):
    q, a, d = T4A[i]
    H.append(p(E(a), lab))
    if i == 1:
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
