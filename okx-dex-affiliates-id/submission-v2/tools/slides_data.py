"""Export every slide text and number from the recalculated submission workbook to JSON.

Usage: python slides_data.py <submission.xlsx> <out.json>
"""
import json
import sys
from pathlib import Path

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).parent))
from layout_v6 import T1 as L1, T2 as L2, T3 as L3, T4 as L4  # noqa: E402

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
         "Crypto Ndo", "DegenRex", "Meteora LP Army"]
proof = ["#3 on fomo's 24-hour profit board (+$53,566)", "Wallet +374 SOL; 1 SOL to 10 live", "~$3.4M referred on BasedBot; $8,463 earned",
         "fomo +$32,185 in one day", "Active trader; calls token buys ($PRL 3.5x)", "Existing trades on OKX Wallet; 59.3K on X",
         "Won Clash of Trenchers live", "462 and 671 at two community nights"]
assert t1.cell(N0 + 2, 1).value.startswith("0xRusdi") and t1.cell(N0 + 3, 1).value.startswith("Brandon")
rows = [[i + 1, names[i], proof[i], t1.cell(A0 + i, 9).value, round(t1.cell(N0 + i, 6).value), str(t1.cell(N0 + i, 8).value),
         k(t1.cell(N0 + i, 16).value)] for i in range(8)]
rows.append(["", "Total", "", "", round(t1.cell(TOT, 6).value), str(t1.cell(TOT, 8).value), m(t1.cell(TOT, 16).value)])
top3_ftd = round(sum(t1.cell(N0 + i, 6).value for i in range(3)))

SCv = lambda pn, j, c: t2.cell(L2["SC0"] + 4 * "ABC".index(pn) + j, c).value
SC = lambda pn, j, c: usd(SCv(pn, j, c))
IN = lambda key, pn: t2.cell(L2[key], 2 + "ABC".index(pn)).value
mm = lambda x: f"${x / 1e6:,.0f}M" if x % 1e6 == 0 else m(x)
t2rows = [["C · trading educator", "$3,000 + 10%", SC("C", 1, 6), SC("C", 0, 10), SC("C", 1, 10), mm(IN("THR", "C"))],
          ["B · repeat on-chain traders", "35%, no fee", SC("B", 1, 6), SC("B", 0, 10), SC("B", 1, 10), "Same terms"],
          ["A · airdrop community", "$2,000 + 20%", SC("A", 1, 6), SC("A", 0, 10), SC("A", 1, 10), mm(IN("THR", "A"))]]
W = L4
wk = {c: [t4[f"{c}{W[x]}"].value for x in ("LIVE", "ACT", "FTD", "TPD")] for c in "BCDE"}
w0 = W["W0"]
t4rows = []
for j, (c, lab) in enumerate(zip("BCDE", ("1 · days 1-7", "2 · days 8-14", "3 · days 15-21", "4 · days 22-30"))):
    t4rows.append([lab, t4.cell(W["FOCUS"], 2 + j).value, t4[f"{c}{w0 + 1}"].value, t4[f"{c}{w0 + 3}"].value,
                   f"{wk[c][0]} + C" if j else "0", wk[c][1], f"{wk[c][2]:,.0f}", f"{wk[c][3]:,.0f}"])
sw30 = t4.cell(W["SW8"], 5).value + t4.cell(W["SWP"], 5).value
rev30 = t4.cell(W["REV"], 5).value
O = lambda j: t3.cell(L3["OFFER0"] + j, 2).value
k_earnt, k_keept = t3.cell(L3["EARNT"], 2).value, t3.cell(L3["KEEPT"], 2).value
plan = [t3.cell(L3["PLAN0"] + i, 2).value for i in range(6)]

D = dict(
    stats=[[str(top3_ftd), f"FTDs in month one from Kevin Susanto, Bevan and 0xRusdi (of {t1.cell(TOT, 6).value:.0f} from all 8)"],
           ["C first", f"once its volume is confirmed: OKX makes {SC('C', 1, 10)} at forecast instead of losing {usd(-SCv('C', 1, 6))}"],
           [f"{wk['E'][1]} of 8", "partners activated by day 30 (10+ FTDs and 50+ trades a day)"],
           [m(sw30), f"of swaps by day 30 from our 8 and pilot C ({wk['E'][3]:,.0f} trades a day)"]],
    t1_answer=t1.cell(L1["ANS0"] + 1, 2).value,
    t1_rows=rows,
    t1_foot=("FTDs = people reached × 2% sign up × 15% trade. FnF power traders: $1,000 a day each from day 4. "
             "6 of 8 not airdrop-led; 5 partner types; meet in person: Kevin, Bevan and 0xRusdi (Jakarta)."),
    t2_answer=t2.cell(L2["ANS0"] + 1, 2).value,
    t2_rows=t2rows,
    t2_deal=[f"C: {IN('SHARE', 'C'):.0%} share; its $3,000 paid the month swaps pass {mm(IN('THR', 'C'))}; 50% of revenue above that",
             f"A: {IN('SHARE', 'A'):.0%} share; its $2,000 paid the month swaps pass {mm(IN('THR', 'A'))}; 50% above",
             "B: its own terms, 35% and no fee: simple and done",
             f"C at $12M earns {SC('C', 3, 9)} on our deal against {SC('C', 3, 5)} on its ask"],
    t2_right=["Evidence: 90 days of dashboards and the wallet addresses of top members, checked on-chain",
              f"At 50%, C still makes OKX {SC('C', 0, 10)}: the bonus just isn't triggered",
              "Other costs: none in month 1; campaigns only past the threshold, capped at 10% of the revenue above it"],
    t2_foot="Month 1 = volume × 0.04% − share − bonus (paid only in a month above the threshold). Every scenario: spreadsheet, Task 2.",
    t3_answer="Kevin Susanto, our #1: a share of every swap his followers make, and his own VIP-style tiers once his link earns OKX $10,000 a month.",
    t3_cols=[["The offer", ["No retainer", f"Standard OKX DEX share ({O(1):.0%}) to start",
                            f"Tiers from {mm(O(3))} of monthly swaps: {O(4):.0%}, {O(6):.0%}, {O(8):.0%}", "Co-hosted live session + 7-day competition",
                            f"At target: {usd(k_earnt)} to him, {usd(k_keept)} to OKX a month"]],
             ["The messages", ["Warm opener: I follow your trades", "Offer in two sentences", "Runs alongside his OKX link: nothing changes",
                               "Next step: coffee or a call, Calendly or Telegram", "Follow-up: a spot kept in the first group",
                               "Pays more elsewhere? Share it; we come back in 48h"]],
             ["First 14 days", ["Days 1-2: link, terms, compliance check", f"Days 2-4: {plan[1].split(' (')[0]}", "Day 5: live session + competition",
                                "Days 8-10: fix the weakest funnel step; product feedback", f"Day 14: {plan[5].replace(' by day 14', '')}"]]],
    t3_foot="The referral share and follower discount are published OKX DEX referral features; his personal tiers, the session and launch terms are subject to OKX approval.",
    t4_answer=t4.cell(L4["ANS0"], 2).value.replace(" (section A)", ""),
    t4_rows=t4rows,
    t4_cards=[["Activated partner", "10+ FTDs and 50+ trades a day within 30 days; 30%+ still trading after 14 days; under 5% flagged wallets"],
              ["20 messages, 0 meetings", "Connections of connections: direct introductions from Kevin, Bevan and 0xRusdi's circles, with the offer up front"],
              ["Day 30 decision", "Continue if activated; renegotiate if power traders trade but FTDs fall short; stop under 5 FTDs or over 10% flagged"]],
    t4_foot=(f"Rolling: partners go live as they sign (3 by day 8, 6 by day 14, all 8 by day 20). Pilot C live from day 10 once its "
             f"evidence checks out. By day 30: {m(sw30)} of swaps, {usd(rev30)} of revenue at 0.04%."),
)
json.dump(D, open(sys.argv[2], "w"), indent=1, ensure_ascii=False)
print(json.dumps(D)[:300])
