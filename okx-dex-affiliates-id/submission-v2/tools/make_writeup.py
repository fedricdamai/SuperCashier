"""Build the Google Doc write-up (HTML for Drive import) from the recalculated submission workbook.

Usage: python make_doc.py <submission.xlsx> <out.html>
Google Docs turns "$...$" pairs into equations when the second "$" follows a non-space character,
so the script refuses to write any block where that would happen.
"""
import html
import re
import sys
from pathlib import Path

from openpyxl import load_workbook

xlsx, out = Path(sys.argv[1]), Path(sys.argv[2])
wb = load_workbook(xlsx, data_only=True)
t1, t2, t3, t4 = (wb[n] for n in ("Task 1", "Task 2", "Task 3", "Task 4"))
E = html.escape


def k(x):
    return f"${x / 1000:,.0f}K"


def m(x):
    return f"${x / 1e6:.2f}M"


def usd(x):
    return f"-${-x:,.0f}" if x < 0 else f"${x:,.0f}"


P = []
for i in range(8):
    r = 37 + i
    P.append(dict(pool=t1.cell(r, 3).value, ftd=t1.cell(r, 6).value, fnf=t1.cell(r, 8).value, ftd_sw=t1.cell(r, 10).value,
                  fnf_sw=t1.cell(r, 11).value, swaps=t1.cell(r, 13).value, rev=t1.cell(r, 15).value, prio=t1.cell(12 + i, 5).value))
T = dict(ftd=t1["F45"].value, fnf=t1["H45"].value, swaps=t1["M45"].value)
kev = P[0]
top3_ftd = sum(P[i]["ftd"] for i in range(3))
wk = {c: [t4[f"{c}{r}"].value for r in (10, 12, 13, 14, 15)] for c in "BCDE"}
pil30, rev30, rev30_after = t4["E16"].value, t4["E17"].value, t4["E18"].value
# Task 2 scenario rows: A 24-27, B 28-31, C 32-35; N (14) ours, O (15) partner on ours, P (16) ask, Q (17) partner on ask
S = {(pn, j): t2.cell(24 + 4 * "ABC".index(pn) + j, 14).value for pn in "ABC" for j in range(4)}
SO = {(pn, j): t2.cell(24 + 4 * "ABC".index(pn) + j, 15).value for pn in "ABC" for j in range(4)}
SA = {(pn, j): t2.cell(24 + 4 * "ABC".index(pn) + j, 16).value for pn in "ABC" for j in range(4)}
SQ = {(pn, j): t2.cell(24 + 4 * "ABC".index(pn) + j, 17).value for pn in "ABC" for j in range(4)}
payback, c_cap = t2["B37"].value, t2["B38"].value
k_share, k_bonus = t3["B37"].value, t3["B38"].value
TASK3 = {t3.cell(r, 1).value: t3.cell(r, 2).value for r in range(4, 12)}
msg = t3["B4"].value
PLAN = [(t3.cell(r, 1).value, t3.cell(r, 2).value, t3.cell(r, 3).value) for r in range(15, 21)]
assert t1["A37"].value.startswith("Kevin") and t1["A39"].value.startswith("Ponyin") and t2["A28"].value == "B"

names = ["Kevin Susanto", "Bevan / Narrative Labs", "Ponyin + CVC network", "GuarEmperor + FnF circle", "Brandon Salim + a79",
         "Crypto Ndo", "DegenRex", "Meteora LP Army"]
types = ["Large KOL, active trader", "Trader-educator, paid community", "Trader KOL, FnF clans, CVC team",
         "Alpha caller, tools Discord", "Celebrity trader, own trader group", "Trench creator", "Champion trader, TG calls",
         "LP community (events)"]
proof = ["#3 on fomo 24h PnL, +$53,566 (19-20 Nov)", "Wallet: ~1,080 SOL swapped, +374 SOL; 1 to 10 SOL live",
         "Trojan PnL about $170K a month; 0xRusdi's BasedBot referral book", "1.9% engagement; members post PnL from his tools",
         "fomo +$32,185 in one day; code on every trade", "Trades on OKX Wallet; 59.3K X",
         "Won Clash of Trenchers live, 1 to 10 SOL", "462 and 671 at two meetups"]
route = ["OKX regional team; in person, Jakarta", "DM, then in person at Narrative Labs", "DM Ponyin; he opens the CVC team",
         "DM; intro via All Good", "Intro via Kevin", "X DM", "DM; Ave Indonesia", "X DM to the hosts"]


def table(head, rows, widths, num_from=None, bold_last=False):
    cols = "<colgroup>" + "".join(f'<col width="{w}%">' for w in widths) + "</colgroup>"
    h = "".join(f'<th width="{w}%" style="width:{w}%;border:1px solid #c8ccd2;padding:2px 4px;background:#eef1f5;text-align:left">{E(str(x))}</th>'
                for x, w in zip(head, widths))
    body = ""
    for i, row in enumerate(rows):
        b = bold_last and i == len(rows) - 1
        tds = ""
        for j, x in enumerate(row):
            al = "right" if num_from is not None and j >= num_from else "left"
            v = f"<b>{E(str(x))}</b>" if b else E(str(x))
            tds += f'<td style="border:1px solid #c8ccd2;padding:2px 4px;text-align:{al}">{v}</td>'
        body += f"<tr>{tds}</tr>"
    return f'<table width="100%" style="border-collapse:collapse;width:100%;font-size:8.5pt">{cols}<tr>{h}</tr>{body}</table>'


def p(txt, label=None):
    lab = f"<b>{E(label)}</b> " if label else ""
    return f'<p style="margin:0 0 5pt 0">{lab}{txt}</p>'


def ul(items):
    return '<ul style="margin:0 0 5pt 0">' + "".join(f'<li style="margin:0">{x}</li>' for x in items) + "</ul>"


def h2(t):
    return f'<h2 style="font-size:12.5pt;margin:8pt 0 3pt 0">{E(t)}</h2>'




# ---- block output for the .docx builder (overrides the HTML helpers above)
def table(head, rows, widths, num_from=None, bold_last=False):
    return dict(type="table", head=[str(x) for x in head], rows=[[str(x) for x in r] for r in rows], widths=widths,
                num_from=num_from, bold_last=bold_last)


def p(txt, label=None):
    return dict(type="p", label=label, html=txt)


def ul(items):
    return dict(type="ul", items=items)


def h2(t):
    return dict(type="h2", text=t)


H = [dict(type="h1", text="Indonesia DEX Affiliate Plan"),
     p("OKX P4 DEX BD Affiliate take-home · Fedric Intan Damai · November 2026. Companion files: a 6-slide deck and the prospect "
       "spreadsheet (every number below is a formula there).")]

H.append(h2("Summary"))
H.append(p(f"Approach Kevin Susanto, Bevan (Narrative Labs) and Ponyin first. Each has public proof that he trades and brings his own "
           f"circle of active traders (FnF), and Ponyin also opens the CVC Ventures circle. Together they bring {top3_ftd:.0f} first-time "
           f"traders (FTDs) in their first month. Pilot B first: it has the clearest evidence and no fixed fee, so I would put campaign "
           f"money behind it to push it to 2-3x its forecast, then start C, and let A follow. By day 30: 8 partners live, "
           f"{wk['E'][1]} activated, about {wk['E'][2]:.0f} FTDs and {wk['E'][3]:.0f} trades a day."))

H.append(h2("Task 1: eight partners whose audiences already trade"))
H.append(p(f"Every partner is sized one way: reach (median views plus members online) × 2% sign-up × 15% trade = FTDs, each "
           f"swapping about $2,000 in month one, plus the FnF power traders it onboards in days 1-3 at $547 a day each. Kevin: "
           f"{kev['pool']:,.0f} reach, {kev['ftd']:.0f} FTDs, {k(kev['ftd_sw'])}; plus {kev['fnf']} FnF traders, {k(kev['fnf_sw'])}; "
           f"{k(kev['swaps'])} in month one, {usd(kev['rev'])} at 0.04%."))
rows = [[f"{i + 1}. {names[i]}", types[i], proof[i], route[i], P[i]["prio"], P[i]["fnf"], f"{P[i]['ftd']:.0f}", k(P[i]["swaps"])] for i in range(8)]
rows.append(["Total", "", "", "", "", T["fnf"], f"{T['ftd']:.0f}", m(T["swaps"])])
H.append(table(["Partner", "Type", "Verified proof", "Route", "Priority", "FnF", "FTDs", "Swaps a month"], rows,
               [17, 16, 26, 14, 8, 5, 6, 8], num_from=5, bold_last=True))
H.append(p("ranked on proof of trading, then the FnF traders they bring, then rival flow we can move, then FTDs and trades a day.",
           "Why these three first:"))
H.append(ul([
    f"<b>Kevin Susanto:</b> the most FTDs ({kev['ftd']:.0f}) from the largest measured reach (41,248 median views); #3 on fomo's 24-hour "
    f"PnL board; OKX's regional team already works with him, so he is the fastest yes. Risk: his existing OKX link, so DEX tracking stays separate.",
    "<b>Bevan / Narrative Labs:</b> his published wallet shows ~1,080 SOL swapped May to July (+374 SOL), and he turned 1 SOL into 10 live "
    "at Web3 Week Asia. He founded a paid Masterclass with a 9,238-member Discord (984 online); we can meet him at Narrative Labs.",
    f"<b>Ponyin and his circle:</b> about $170K of Trojan volume in a month; ApexFNF is his FnF; and he opens the CVC Ventures team "
    f"(0xRusdi's referral book and Tomket clan, Setya's airdrop traders): {P[2]['fnf']} FnF traders and {P[2]['ftd']:.0f} FTDs from one introduction.",
]))
H.append(p("GuarEmperor (#4) has the best engagement of the large accounts (1.9%) but his circle's trading is not verified yet, so he "
           "starts in wave 2. Verified: trading proof, reach and community counts (X API, Solscan, Discord, organiser posts, 20-21 Nov "
           "2026); assumed: the 2% and 15% rates and volume per trader. Brief checks: 6 of 8 not airdrop-led, 6 partner types, 2 "
           "meetable in person (Kevin and Bevan). All fields for each partner: spreadsheet, Task 1."))

H.append(h2("Task 2: pilot B first, on a growth deal"))
deal = {"A": "20% share only, later", "B": "30% rising to 45%, bonus, $250 campaign",
        "C": "$500 + $500 retainer, 15% rising to 30%, bonus, $300 event"}
askt = {"A": "$2,000 + 20%", "B": "35%, no fee", "C": "$3,000 + 10%"}
rows = [[pn, askt[pn], usd(SA[(pn, 1)]), deal[pn], usd(S[(pn, 0)]), usd(S[(pn, 1)]), usd(S[(pn, 2)]), usd(S[(pn, 3)]),
         f"{usd(SQ[(pn, 1)])} / {usd(SO[(pn, 1)])}"] for pn in "ABC"]
H.append(table(["Partner", "Their ask", "Month 1, their ask", "Our deal", "At 50%", "At forecast", "At 2x", "At 3x",
                "Partner earns at forecast (ask / ours)"], rows, [9, 11, 10, 23, 8, 9, 8, 8, 14], num_from=4))
H.append(p(f"Month 1 = revenue at 0.04% − retainer − revenue share − bonus − campaign spend. B at forecast: $800 − $240 share − $80 "
           f"bonus − $250 campaign = {usd(S[('B', 1)])}; at its own ask, $800 − $280 = {usd(SA[('B', 1)])}. A at its ask: $1,600 − $2,000 "
           f"− $320, a $720 loss. C at its ask: $2,400 − $3,000 − $240, an $840 loss."))
H.append(p("B: the clearest evidence (repeat onchain traders), no fixed fee, positive even at half volume. I would spend campaign "
           "money on B to push it to 2-3x its forecast, then start C to expand (its volume has to be proven first), and let A follow "
           "on share only: an airdrop audience is not our target and arrives with campaigns on its own.", "Preferred pilot:"))
H.append(p("no retainer for B, and a small one for C ($500 at signing, $500 more only when it reaches forecast) instead of $3,000. "
           "Revenue share in brackets that rise as volume grows (B: 30% up to forecast, then 35%, 40% and 45%; C: 15% rising to 30%, "
           "above its 10% ask). A cash bonus at 100%, 200% and 300% of forecast, paid within 7 days and worth 10% of the revenue at that "
           "milestone. OKX-funded campaigns: a trading competition for B, a co-hosted event for C, plus 10% of revenue above forecast "
           "reinvested in more of them. Every scenario stays positive.", "Counter-offer, a growth deal:"))
H.append(p(f"B's $250 campaign pays for itself once B passes {payback:.1f}x its forecast; if B is not on track by week 3, it stops. "
           f"C's upfront retainer stays at $500 because above {usd(c_cap)} C turns negative at half volume.", "Is the campaign spend worth it?"))
H.append(p("weekly payouts and the bonus within 7 days; a rate ratchet (a bracket reached is next month's starting rate, if retention "
           "holds); a network override (5% of revenue from partners they introduce, for 3 months); member perks from the campaign pool "
           "(competitions, fee discounts, early access to OKX Boost and X Launch).", "To make it more attractive at no fixed cost:"))
H.append(p("90 days of past numbers (referral or bot dashboards: volume, active traders, retention by week); proof of users "
           "(community size, activity and location); trading history and the wallet addresses of their most active members, which we "
           "check onchain for repeat trading and wash trades; past partnership results and a reference.", "Evidence before signing:"))
H.append(p(f"B stays positive at {usd(S[('B', 0)])} on our deal and {usd(SA[('B', 0)])} on its own ask, because there is no fixed fee, "
           f"and its bracket share falls to 30% automatically. Then I would find out why: the market (overall memecoin DEX volume), the "
           f"incentive (activity dropping after the competition) or the group itself (active wallets falling). The first two are normal "
           f"swings: adjust the campaign. If the group has stopped carrying its weight, renegotiate or stop.", "If volume reaches only 50%:"))
H.append(p("campaign budgets per community, bonus payouts, follower fee discounts (0-20%), gas subsidies, content fees, fraud screening "
           "and tracking tools, compliance review (POJK 6/2026, Law 4/2026), translation and BD time.", "Other costs to investigate:"))

H.append(h2("Task 3: winning and activating Kevin Susanto"))
H.append(p(f"no retainer; fee share of 30% on the first $250K of monthly referred swaps, 35% to $500K and 40% above; a cash bonus of 10% "
           f"of the swap fees at $250K, $500K and $1M, paid within 7 days; and a co-hosted live session with a 7-day trading competition "
           f"for his Telegram: about {usd(k_share)} in share and {usd(k_bonus)} in bonus at his forecast. Brackets, bonus and campaign "
           f"are proposals until OKX approves them.", "Offer:"))
H.append(p(f"<i>{E(msg)}</i>", f"First message ({len(msg.split())} words; hook, who I am and the proposal, what's in it for him and his followers, next step):"))
H.append(p(f"<i>{E(TASK3['Follow-up (day 4, no reply)'])}</i>", "Follow-up, day 4 (the deadline is real: our first wave):"))
H.append(p(f"<i>{E(TASK3[chr(82) + 'eply to ' + chr(39) + 'another platform pays more' + chr(39)])}</i>", "If he says another platform pays more:"))
H.append(p("First-meeting questions:"))
H.append(dict(type="ol", items=[E(TASK3[f"Meeting question {i}"]) for i in (1, 2, 3)]))
short = ["Separate DEX link and discount, tracked apart from his okx.ac link; terms and payout wallet in writing; compliance check",
         "Activate his circle: 3 traders onboarded 1:1 through the full flow (wallet, funding, first trade, Track alerts)",
         "Launch the campaign: co-hosted live session on his Telegram and a 7-day trading competition",
         "Push the flow: 2 X threads, a pinned guide and an FAQ for first-time DEX users",
         "Fix the funnel step by step (click, wallet, deposit, first and second trade); UI/UX feedback to product",
         "Competition results, second session, review; agree days 15-30 and the next bracket"]
H.append(table(["Days", "Activation plan", "Target"], [[a, short[i], c] for i, (a, b, c) in enumerate(PLAN)], [9, 63, 28]))

H.append(h2("Task 4: the first 30 days"))
rows = [["1 · days 1-7", "Score 25 prospects; message all", 25, 3, 0, 0, 0, 0],
        ["2 · days 8-14", "Sign top 3 and B; Jakarta day", 20, 6, f"{wk['C'][0]} + B", wk["C"][1], f"{wk['C'][2]:.0f}", f"{wk['C'][3]:.0f}"],
        ["3 · days 15-21", "Wave 2 live; B campaign", 15, 5, f"{wk['D'][0]} + B", wk["D"][1], f"{wk['D'][2]:.0f}", f"{wk['D'][3]:.0f}"],
        ["4 · days 22-30", "DegenRex, Meteora night; review; decide", 10, 4, f"{wk['E'][0]} + B", wk["E"][1], f"{wk['E'][2]:.0f}", f"{wk['E'][3]:.0f}"]]
H.append(table(["Week", "Focus", "Outreach", "Meetings", "Live", "Activated", "FTDs", "Trades a day"], rows,
               [13, 35, 9, 9, 9, 9, 7, 9], num_from=2))
H.append(p(f"Swaps by day 30: {k(wk['E'][4])} from our 8 and {m(pil30)} from pilot B, so {usd(rev30)} of revenue at 0.04%, or "
           f"{usd(rev30_after)} after partner shares. C starts once B passes 2x (month 2)."))
H.append(p("at least 10 FTDs and 50 trades a day within 30 days live, at least 30% of FTDs still trading after 14 days, and under 5% "
           "flagged wallets (wash trades, self-referral, sybil clusters).", "Activated partner:"))
H.append(p("Wednesday of week 2 in Jakarta: Kevin Susanto at 10:00 (through OKX's regional team), Bevan at Narrative Labs at 13:00 "
           "to sign and onboard his first traders, then an LP Army night or W3AI Wednesday at JURA. Fallback: a same-day call.", "In-person day:"))
H.append(p("go through connections of connections. Ask the partners already talking to us and their circles (Kevin, Bevan, Ponyin) "
           "for direct intros to the decision-maker, skip the public inbox, and lead with an easy offer: fee share plus a cash bonus "
           "within 7 days of each milestone.", "If 20 messages bring 2 replies and no meetings:"))
H.append(p("continue if activated with under 5% flagged wallets; renegotiate if the FnF traders trade but FTDs or trades a day fall "
           "short; stop if under 5 FTDs or over 10% flagged.", "Day 30 decision per partner:"))

H.append(h2("Commitment"))
H.append(ul(["<b>Earliest start:</b> in 2 weeks.", "<b>Full-time:</b> yes, from 2 weeks.",
             "<b>Other commitments:</b> marketer role at Surgence Labs.",
             "<b>In person:</b> available across Indonesia and Southeast Asia, and willing to attend partner meetings and events in person."]))
H.append(h2("How I worked and sources"))
H.append(p("I set the strategy, the selection logic and every judgment call. AI gathered public data, automated repetitive work "
           "(screening 100 prospects, building the model) and fact-checked. I re-checked every part myself and ran independent "
           "fact-checks with several AI agents. Public information only."))
H.append(p("Sources: OKX DEX docs, DefiLlama, Solscan, X, Discord, Luma, OJK; every link is in the spreadsheet (Miscellaneous, section G)."))
import json
out.write_text(json.dumps(H, indent=1, ensure_ascii=False))
print("blocks", len(H))
