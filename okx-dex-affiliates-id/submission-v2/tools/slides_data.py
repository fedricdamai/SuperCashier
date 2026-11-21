"""Export the numbers the slides need from the recalculated submission workbook to JSON.

Usage: python slides_data.py <submission.xlsx> <out.json>
"""
import json
import sys

from openpyxl import load_workbook

wb = load_workbook(sys.argv[1], data_only=True)
t1, t2, t3, t4 = (wb[n] for n in ("Task 1", "Task 2", "Task 3", "Task 4"))


def k(x):
    return f"${x / 1000:,.0f}K"


def usd(x):
    return f"-${-x:,.0f}" if x < 0 else f"${x:,.0f}"


names = ["Kevin Susanto", "Bevan / Narrative Labs", "Ponyin + CVC network", "GuarEmperor + FnF", "Brandon Salim + a79",
         "Crypto Ndo", "DegenRex", "Meteora LP Army"]
proof = ["#3 on fomo 24h PnL, +$53.6K", "Wallet +374 SOL; 1 to 10 SOL live", "Trojan ~$170K/month; BasedBot referral book",
         "1.9% engagement; member PnL posts", "fomo +$32K in one day", "Trades on OKX Wallet; 59.3K X",
         "Won Clash of Trenchers live", "462 + 671 at two meetups"]
partners = []
for i in range(8):
    r = 37 + i
    partners.append(dict(name=names[i], proof=proof[i], prio=t1.cell(12 + i, 5).value, fnf=t1.cell(r, 8).value,
                         ftd=round(t1.cell(r, 6).value), swaps=k(t1.cell(r, 13).value)))
S = {}
for pn in "ABC":
    for j, lab in enumerate(("50", "100", "200", "300")):
        row = 24 + 4 * "ABC".index(pn) + j
        S[f"{pn}{lab}"] = usd(t2.cell(row, 14).value)
        S[f"{pn}{lab}_ask"] = usd(t2.cell(row, 16).value)
kev = dict(ftd=round(t1["F37"].value), tpd=round(t1["I37"].value), swaps=k(t1["M37"].value), share=usd(t3["B37"].value),
           bonus=usd(t3["B38"].value), plan=[t3.cell(r, 3).value for r in range(15, 21)])
wk = {c: [t4[f"{c}{r}"].value for r in (10, 12, 13, 14, 15)] for c in "BCDE"}
weeks = [["1 · days 1-7", 25, 3, "0", 0, 0, 0],
         ["2 · days 8-14", 20, 6, f"{wk['C'][0]} + B", wk["C"][1], round(wk["C"][2]), round(wk["C"][3])],
         ["3 · days 15-21", 15, 5, f"{wk['D'][0]} + B", wk["D"][1], round(wk["D"][2]), round(wk["D"][3])],
         ["4 · days 22-30", 10, 4, f"{wk['E'][0]} + B", wk["E"][1], round(wk["E"][2]), round(wk["E"][3])]]
data = dict(partners=partners, total=dict(fnf=t1["H45"].value, ftd=round(t1["F45"].value), swaps=f"${t1['M45'].value / 1e6:.2f}M"),
            top3_ftd=round(sum(t1.cell(37 + i, 6).value for i in range(3))), S=S, payback=f"{t2['B37'].value:.1f}x",
            kev=kev, weeks=weeks, activated=wk["E"][1], ftd30=round(wk["E"][2]), tpd30=round(wk["E"][3]),
            swaps30=k(wk["E"][4]), rev30=usd(t4["E17"].value))
json.dump(data, open(sys.argv[2], "w"), indent=1)
print(json.dumps(data)[:600])
