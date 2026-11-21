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
names = ["Kevin Susanto", "Bevan / Narrative Labs", "Ponyin + CVC network", "GuarEmperor + FnF", "Brandon Salim + a79",
         "Crypto Ndo", "DegenRex", "Meteora LP Army"]
proof = ["#3 on fomo's 24-hour profit board (+$53,566)", "Wallet +374 SOL; 1 SOL to 10 live", "~$170K a month on Trojan; BasedBot referral book",
         "1.9% engagement; members post gains", "fomo +$32,185 in one day", "Trades only on OKX Wallet; 59.3K on X",
         "Won Clash of Trenchers live", "462 and 671 at two community nights"]
rows = [[i + 1, names[i], proof[i], t1.cell(A0 + i, 9).value, round(t1.cell(N0 + i, 6).value),
         f"{t1.cell(N0 + i, 8).value} × {usd(t1.cell(N0 + i, 10).value)}", k(t1.cell(N0 + i, 16).value)] for i in range(8)]
rows.append(["", "Total", "", "", round(t1.cell(TOT, 6).value), str(t1.cell(TOT, 8).value), m(t1.cell(TOT, 16).value)])
top3_ftd = round(sum(t1.cell(N0 + i, 6).value for i in range(3)))

SC = lambda pn, j, c: usd(t2.cell(L2["SC0"] + 4 * "ABC".index(pn) + j, c).value)
t2rows = [["B · repeat on-chain traders", "35%, no fee", SC("B", 1, 16), SC("B", 0, 14), SC("B", 1, 14), SC("B", 2, 14), SC("B", 3, 14)],
          ["C · trading educator", "$3,000 + 10%", SC("C", 1, 16), SC("C", 0, 14), SC("C", 1, 14), SC("C", 2, 14), SC("C", 3, 14)],
          ["A · airdrop community", "$2,000 + 20%", SC("A", 1, 16), SC("A", 0, 14), SC("A", 1, 14), SC("A", 2, 14), SC("A", 3, 14)]]
payback = t2.cell(L2["PAY"], 2).value
W = L4
wk = {c: [t4[f"{c}{W[x]}"].value for x in ("LIVE", "ACT", "FTD", "TPD")] for c in "BCDE"}
w0 = W["W0"]
t4rows = []
for j, (c, lab) in enumerate(zip("BCDE", ("1 · days 1-7", "2 · days 8-14", "3 · days 15-21", "4 · days 22-30"))):
    t4rows.append([lab, t4.cell(W["FOCUS"], 2 + j).value, t4[f"{c}{w0 + 1}"].value, t4[f"{c}{w0 + 3}"].value,
                   f"{wk[c][0]} + B" if j else "0", wk[c][1], f"{wk[c][2]:,.0f}", f"{wk[c][3]:,.0f}"])
sw30 = t4.cell(W["SW8"], 5).value + t4.cell(W["SWP"], 5).value
rev30 = t4.cell(W["REV"], 5).value
k_share, k_bonus = t3.cell(L3["SHARE"], 2).value, t3.cell(L3["BONUS"], 2).value
plan = [t3.cell(L3["PLAN0"] + i, 2).value for i in range(6)]

D = dict(
    stats=[[str(top3_ftd), f"FTDs in month one from Kevin Susanto, Bevan and Ponyin (of {t1.cell(TOT, 6).value:.0f} from all 8)"],
           ["B first", "the one pilot that makes money at its own ask, even at half volume"],
           [f"{wk['E'][1]} of 8", "partners activated by day 30 (10+ FTDs and 50+ trades a day)"],
           [m(sw30), f"of swaps by day 30 from our 8 and pilot B ({wk['E'][3]:,.0f} trades a day)"]],
    t1_answer=t1.cell(L1["ANS0"] + 1, 2).value,
    t1_rows=rows,
    t1_foot=("FTDs = people reached × 2% sign up × 15% trade. FnF power traders: $1,000 a day each, 2x in Kevin's whale circle, "
             "1.5x in Ponyin's network. 6 of 8 not airdrop-led; 8 partner types; meet in person: Kevin and Bevan (Jakarta)."),
    t2_answer=t2.cell(L2["ANS0"] + 1, 2).value,
    t2_rows=t2rows,
    t2_deal=["B: no retainer; share 30% rising to 45% by bracket; $250 competition",
             "C: $500 retainer + $500 at forecast; share 15% rising to 30%; $300 event",
             "A: 20% share only, after B and C",
             "Cash bonus within 7 days at 1x, 2x and 3x forecast"],
    t2_right=["Evidence: 90 days of dashboards and the wallet addresses of top members, checked on-chain",
              f"At 50%, B still makes {SC('B', 0, 14)}: check the market, the incentive, then the group",
              f"B's campaign pays for itself above {payback:.1f}x forecast; stop it at week 3 if off track"],
    t2_foot="Month 1 = volume × 0.04% − retainer − revenue share − bonus − campaign spend. Every scenario: spreadsheet, Task 2.",
    t3_answer="Kevin Susanto, our #1: lead with one benefit, a share of every swap his followers make, next to the OKX link he already has.",
    t3_cols=[["The offer", ["No retainer", "Fee share 30%, rising to 40% as volume grows", "Cash bonus within 7 days at $250K, $500K and $1M",
                            "Co-hosted live session + 7-day competition", f"Month one: ~{usd(k_share)} share + {usd(k_bonus)} bonus"]],
             ["The messages", ["Hook: his fomo top-traders day, lightly", "One benefit up front: he earns on every swap", "Runs alongside his OKX link: nothing changes",
                               "Next step: one-page brief, Calendly or Telegram", "Follow-up: real first-wave deadline", "Pays more elsewhere? Send your data; we beat it in 48h"]],
             ["First 14 days", ["Days 1-2: link, terms, compliance check", f"Days 2-4: {plan[1].split(' (')[0]}", "Day 5: live session + competition",
                                "Days 8-10: fix the weakest funnel step; product feedback", f"Day 14: {plan[5].replace(' by day 14', '')}"]]],
    t3_foot="The fee share and follower discount are published OKX DEX referral features; the bonus, session and launch terms are subject to OKX approval.",
    t4_answer=t4.cell(L4["ANS0"], 2).value.replace(" (section A)", ""),
    t4_rows=t4rows,
    t4_cards=[["Activated partner", "10+ FTDs and 50+ trades a day within 30 days; 30%+ still trading after 14 days; under 5% flagged wallets"],
              ["20 messages, 0 meetings", "Connections of connections: direct introductions from Kevin, Bevan and Ponyin's circles, with the offer up front"],
              ["Day 30 decision", "Continue if activated; renegotiate if power traders trade but FTDs fall short; stop under 5 FTDs or over 10% flagged"]],
    t4_foot=(f"Rolling: partners go live as they sign (3 by day 8, 6 by day 14, all 8 by day 20). Pilot B live from day 8; C starts once B "
             f"passes 2x. By day 30: {m(sw30)} of swaps, {usd(rev30)} of revenue at 0.04%."),
)
json.dump(D, open(sys.argv[2], "w"), indent=1, ensure_ascii=False)
print(json.dumps(D)[:300])
