"""Export every slide text and number from the recalculated submission workbook to JSON.

Usage: python slides_data.py <submission.xlsx> <out.json>
"""
import json
import sys
from pathlib import Path

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).parent))
from layout_v6 import T1 as L1, T2 as L2, T3 as L3, T4 as L4, GT  # noqa: E402

wb = load_workbook(sys.argv[1], data_only=True)
t1, t2, t3, t4 = (wb[n] for n in ("Task 1", "Task 2", "Task 3", "Task 4"))


def k(x):
    return f"${x / 1000:,.0f}K"


def m(x):
    return f"${x / 1e6:.2f}M"


def usd(x):
    return f"-${-x:,.0f}" if x < 0 else f"${x:,.0f}"


N0, TOT, A0 = L1["N0"], L1["TOT"], L1["A0"]
names = ["Kevin Susanto", "Bevan / Narrative Labs", "0xRusdi + CVC network", "Brandon Salim + a79", "GuarEmperor + FnF",
         "Michael Wyann / Laba Lab", "Crypto Ndo", "DegenRex"]
proof = ["#3 on fomo's 24-hour profit board (+$53,566)", "Wallet +374 SOL; 1 SOL to 10 live", "~$3.4M referred on BasedBot; $8,463 earned",
         "fomo +$32,185 in one day", "Active trader; calls token buys ($PRL 3.5x)", "Paid membership; 121K on YouTube",
         "Existing trades on OKX Wallet; 59.3K on X", "Won Clash of Trenchers live"]
assert t1.cell(N0 + 5, 1).value.startswith("Michael")
rows = [[i + 1, names[i], proof[i], t1.cell(A0 + i, 9).value, round(t1.cell(N0 + i, 6).value), str(t1.cell(N0 + i, 8).value),
         k(t1.cell(N0 + i, 16).value)] for i in range(8)]
rows.append(["", "Total", "", "", round(t1.cell(TOT, 6).value), str(t1.cell(TOT, 8).value), m(t1.cell(TOT, 16).value)])
top3_ftd = round(sum(t1.cell(N0 + i, 6).value for i in range(3)))

SCv = lambda pn, j, c: t2.cell(L2["SC0"] + 4 * "ABC".index(pn) + j, c).value
minsh = t2.cell(L2["MINSH"], 2).value
cellv = lambda pn, j: f"{usd(SCv(pn, j, 10))} ({SCv(pn, j, 11):.0%})"
t2rows = [["Their ask, at forecast"] + [usd(SCv(pn, 1, 6)) for pn in "ABC"]]
for j, lab in enumerate(("Our deal at 50% of forecast", "Our deal at forecast", "Our deal at $10M a month", "Our deal at $20M a month")):
    t2rows.append([lab] + [cellv(pn, j) for pn in "ABC"])
W = L4
wk = {c: [t4[f"{c}{W[x]}"].value for x in ("LIVE", "ACT", "FTD", "TPD")] for c in "BCDE"}
sw8 = t4.cell(W["SW8"], 5).value
pil = t4.cell(W["PIL0"] + 1, 5).value
O = lambda j: t3.cell(L3["OFFER0"] + j, 2).value
k_earnt, k_keept = t3.cell(L3["EARNT"], 2).value, t3.cell(L3["KEEPT"], 2).value
plan = [t3.cell(L3["PLAN0"] + i, 2).value for i in range(6)]
gt = wb["Task 4 Gantt"]
gantt = []
r = GT["G0"]
blocks = [(d, d + 6) for d in (1, 8, 15, 22)]
blocks[-1] = (22, 30)
while gt.cell(r, 1).value:
    name, track, kpi, s0, s1 = (gt.cell(r, c).value for c in range(1, 6))
    gantt.append([name, str(kpi), track] + [1 if (s0 <= b1 and s1 >= b0) else 0 for b0, b1 in blocks])
    r += 1
kpi = []
for i in range(GT["KPI0"], GT["KPI0"] + 9):
    lab = gt.cell(i, 1).value
    if lab in ("Partners live", "FTDs (running total)", "Trades a day", "Swaps (running total)", "Pilot C: swaps (running total)"):
        vals = [gt.cell(i, c).value for c in range(2, 6)]
        money = "swaps" in lab.lower()
        kpi.append([lab.replace(" (running total)", "")] + [(f"${v / 1e6:.1f}M" if v >= 1e6 else k(v)) if money else f"{v:,.0f}" for v in vals])

D = dict(
    stats=[[str(top3_ftd), f"FTDs in month one from Kevin Susanto, Bevan and 0xRusdi (of {t1.cell(TOT, 6).value:.0f} from all 8)"],
           ["C first", f"once its volume is confirmed; OKX keeps at least {minsh:.0%} of revenue in every case"],
           [f"{wk['E'][1]} of 8", "partners activated by day 30 (10+ FTDs and 50+ trades a day)"],
           [k(sw8), f"of swaps by day 30 from our 8 ({wk['E'][3]:,.0f} trades a day); pilot C adds ${pil / 1e6:.0f}M"]],
    t1_answer=t1.cell(L1["ANS0"] + 1, 2).value,
    t1_rows=rows,
    t1_foot=("FTDs = people reached × 2% sign up × 15% trade. FnF power traders: $1,000 a day each from day 4. "
             "6 of 8 not airdrop-led; 5 partner types; meet in person: Kevin, Bevan and 0xRusdi (Jakarta)."),
    t2_answer=t2.cell(L2["ANS0"] + 1, 2).value,
    t2_rows=t2rows,
    t2_deal=["A: 15% share; payouts $300 at $4M up to $2,000 at $15M",
             "B: 20% to $2M, 25%, 30%, then 35% above $10M; payouts $100 to $800",
             "C: 10% share; payouts $500 at $6M up to $3,000 at $20M",
             "No fixed fee: payouts only in a month that reaches the milestone"],
    t2_right=["Evidence: 90 days of dashboards and the wallets of top members, checked on-chain",
              f"At 50%, C still makes OKX {usd(SCv('C', 0, 10))}: no milestone, nothing extra paid",
              "Other costs: campaigns past the first milestone (capped at 10% of revenue above it) and fee discounts for named FnF traders or past a volume level"],
    t2_foot="Cells: month 1 for OKX (OKX's share of revenue). Month 1 = volume × 0.04% − share − milestone payout. Every case: spreadsheet, Task 2.",
    t3_answer="Kevin Susanto, our #1: a share of every swap his followers make, and his own VIP-style tiers once his link earns OKX $10,000 a month.",
    t3_cols=[["The offer", ["No retainer", f"Standard OKX DEX share ({O(1):.0%}) to start",
                            f"Tiers from ${O(3) / 1e6:.0f}M of monthly swaps: {O(4):.0%}, {O(6):.0%}, {O(8):.0%}", "Co-hosted live session + 7-day competition",
                            f"At target: {usd(k_earnt)} to him, {usd(k_keept)} to OKX a month"]],
             ["The messages", ["Warm opener: I follow your trades", "Offer in two sentences", "Runs alongside his OKX link: nothing changes",
                               "Next step: coffee or a call, Calendly or Telegram", "Follow-up: a spot kept in the first group",
                               "Pays more elsewhere? Share it; we come back in 48h"]],
             ["First 14 days", ["Days 1-2: link, terms, shared dashboard", f"Days 2-4: {plan[1].split(' (')[0]}", "Day 5: live session + competition",
                                "Days 8-10: fix the weakest funnel step; product feedback", f"Day 14: {plan[5].replace(' by day 14', '')}"]]],
    t3_foot="The referral share and follower discount are published OKX DEX referral features; his personal tiers, the session and launch terms are subject to OKX approval.",
    t4_answer=f"Our 8: {wk['E'][0]} live, {wk['E'][1]} activated, ~{wk['E'][2]:,.0f} FTDs, {k(sw8)} of swaps by day 30; pilot C is its own track.",
    gantt=gantt,
    kpi=kpi,
    t4_foot="Activated: 10+ FTDs and 50+ trades a day within 30 days, 30%+ still trading after 14 days, under 5% flagged wallets. Day by day: spreadsheet, Task 4 Gantt.",
)
json.dump(D, open(sys.argv[2], "w"), indent=1, ensure_ascii=False)
print(json.dumps(D)[:300])
