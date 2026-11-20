"""Build the submission workbook (Task 1-4 + Miscellaneous) and the research workbook (same + research tabs).

Usage: python build_workbooks.py <legacy_model.xlsx> <prospects.csv> <REVIEW.md> <out_dir>
Partner text for the six kept partners comes from the legacy model; numbers are live formulas.
"""
import csv
import re
import sys
from copy import copy
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

LEGACY, PROSPECTS, REVIEW, OUTDIR = map(Path, sys.argv[1:5])
L = load_workbook(LEGACY, data_only=True)

# ---------------------------------------------------------------- styles (match the original workbook)
F_BODY = Font(name="Calibri", size=11)
F_BOLD = Font(name="Calibri", size=11, bold=True)
F_TITLE = Font(name="Calibri", size=14, bold=True)
F_SECTION = Font(name="Calibri", size=12, bold=True)
F_LINK = Font(name="Calibri", size=11, color="0563C1", underline="single")
FILL_HEAD = PatternFill("solid", fgColor="F3F3F3")
FILL_INPUT = PatternFill("solid", fgColor="FFF3C4")
THIN = Side(style="thin", color="D0D0D0")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")
TOP = Alignment(vertical="top")
MONEY, MONEY2, PCT, PCT2, NUM1 = '$#,##0', '$#,##0.00', '0%', '0.00%', '0.0'


def put(ws, ref, value, font=F_BODY, fmt=None, fill=None, wrap=True, border=False):
    c = ws[ref]
    c.value = value
    c.font = font
    c.alignment = WRAP if wrap else TOP
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    if border:
        c.border = BOX
    return c


def header(ws, row, labels, col=1):
    for i, h in enumerate(labels):
        c = ws.cell(row, col + i, h)
        c.font, c.fill, c.alignment, c.border = F_BOLD, FILL_HEAD, WRAP, BOX


def cell(ws, row, col, value, fmt=None, inp=False, bold=False, wrap=True):
    c = ws.cell(row, col, value)
    c.font = F_BOLD if bold else F_BODY
    c.alignment = WRAP if wrap else TOP
    c.border = BOX
    if fmt:
        c.number_format = fmt
    if inp:
        c.fill = FILL_INPUT
    return c


def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


# ---------------------------------------------------------------- partners (rank order)
SL, PP, FP, DP, OF = (L["Shortlist (Task 1)"], L["Partner plan (3 scenarios)"], L["FTD plan"],
                      L["30-day plan"], L["Offers & support"])
KEEP = [  # name key, shortlist row, partner-plan row, FTD-plan row, ranking row, go-live, event, FnF text, in person, airdrop-led
    ("0xRusdi", 8, 26, 6, 8, 10, "No", "His BasedBot referral book: his 5 most active referred traders", "No", "No"),
    ("Crypto Ndo", 9, 21, 7, 9, 10, "No", "Traders in his circle (FnF): 5 onboarded first", "No", "No"),
    ("Ponyin", 10, 28, 8, 10, 10, "No", "His trading circle (FnF): 5 onboarded first", "No", "No"),
    ("TwipsX", 11, 27, 9, 11, 17, "No", "5 heavy traders from his promo network", "No", "No"),
    ("Tomket Lovers", 12, 24, 10, 12, 24, "No", "Its fomo clan of 13-15 active traders: 5 onboarded first", "No", "Yes"),
    ("Meteora LP Army", 13, 23, 11, 13, 17, "Yes", "5 LP whales from the community", "Yes", "No"),
]
partners = []
for key, s, p, f, rk, live, ev, fnf, inperson, airdrop in KEEP:
    partners.append(dict(
        name=SL.cell(s, 1).value, link=SL.cell(s, 2).value, type=SL.cell(s, 3).value, audience=SL.cell(s, 5).value,
        evidence=SL.cell(s, 6).value, route=SL.cell(s, 7).value, proposal=SL.cell(s, 8).value, risk=SL.cell(s, 9).value,
        reason=SL.cell(s, 10).value, assumptions=SL.cell(s, 11).value, vampire=SL.cell(s, 12).value, fnf=fnf,
        in_person=inperson, airdrop=airdrop, segment=PP.cell(p, 2).value, aud_n=PP.cell(p, 20).value,
        rate=PP.cell(p, 21).value, days=PP.cell(p, 22).value, share=PP.cell(p, 14).value, ptraders=5,
        basis=PP.cell(p, 23).value, deliverables=FP.cell(f, 3).value,
        scores=tuple(DP.cell(rk, c).value for c in range(2, 9)), go_live=live, event=ev, short=key))
partners.insert(4, dict(
    short="GuarEmperor", name="GuarEmperor (@GuarEmperor) and his FnF circle", link="https://x.com/GuarEmperor",
    type="Alpha-call creator with an FnF trading circle (X)",
    audience="Indonesian DeFi and memecoin traders following his alpha calls and smart-money wallet lists",
    evidence=("~63-90K X followers across two public snapshots; listed by ICN among 10 X crypto influencers growing the "
              "Indonesian cryptosphere; pinned 'alpha results' record (Dec 2025); public smart-money wallet lists "
              "('Stable Meme Wallets', 'Arc Smart Money Wallets', 'Top 60 Wallets') formatted for BasedBot and GMGN"),
    assumptions="Size and activity of his FnF circle; share of followers who trade rather than farm airdrops; current follower count",
    route="X DM; sister account @emperorjournal_",
    proposal=("OKX DEX referral fee share with a follower discount; his FnF circle (5 traders) onboarded 1:1 in days 1-3; "
              "his wallet lists and alpha calls executed through OKX Wallet Meme Mode and Track"),
    risk="Airdrop and testnet guides in his history (farming audience); alpha calls need front-running and disclosure checks",
    reason="Brings an FnF trading circle, and followers who already copy his wallet lists on BasedBot and GMGN",
    vampire="BasedBot and GMGN (his wallet lists are built for their import)",
    fnf="His FnF circle: 5 traders onboarded first (size to confirm on the first call)", in_person="No", airdrop="No",
    segment="Memecoin trenchers", aud_n=4410, rate=0.01, days=15, share=0.3, ptraders=5,
    basis="~7% of ~63K X followers see a typical post (lower of two public snapshots)",
    deliverables="Days 1-3: his FnF circle (5 traders) onboarded 1:1; then 6 alpha threads executed on OKX Wallet and 1 X Spaces with the referral link",
    scores=(3, 4, 3, 5, 3, 5, 3), go_live=17, event="No"))
partners.append(dict(
    short="Superteam Indonesia", name="Superteam Indonesia (Solana community)", link="https://luma.com/7gbaanxq",
    type="Solana ecosystem community (events)", audience="Solana users, traders and builders in Jakarta and Bali",
    evidence=("Solana Superteam Indonesia Ecosystem Showcase in Jakarta: 502 went (Luma); runs Superteam Hub Jakarta with "
              "Indonesia Crypto Hub; its Country Lead was named a Coinvestasi 2025 most impactful figure"),
    assumptions="Share of members who trade memecoins rather than build; capacity of a co-hosted trader night",
    route="Superteam Hub Jakarta in person; biweekly community call; Country Lead via LinkedIn",
    proposal=("Co-hosted Solana trader night at Superteam Hub with wallets funded on site; Meme Mode and Track walkthrough; "
              "2 active traders from the community onboarded 1:1; tracked host link; fee share"),
    risk="Builder-first community; Solana Foundation neutrality may limit exchange-branded events",
    reason="Meetable every week in Jakarta and already onchain, but builder-led: proof of trading still to check",
    vampire="Solana wallets and bots its members already use (Phantom, Jupiter, GMGN; to confirm)",
    fnf="2 active traders from the community (builder-led, so fewer than the trench partners)", in_person="Yes", airdrop="No",
    segment="Memecoin trenchers", aud_n=150, rate=0.2, days=28, share=0.25, ptraders=2,
    basis="One co-hosted trader night of ~150 (its Jakarta showcase drew 502); wallets funded on site, so these FTDs trade from day 1",
    deliverables="Day 1: co-hosted Solana trader night at Superteam Hub Jakarta with wallets funded on site; 2 active traders onboarded in days 1-3",
    scores=(2, 4, 2, 3, 5, 4, 4), go_live=24, event="Yes"))
assert len(partners) == 8


def build(research=False):
    wb = Workbook()
    t1 = wb.active
    t1.title = "Task 1"
    t2, t3, t4, mi = (wb.create_sheet(n) for n in ("Task 2", "Task 3", "Task 4", "Miscellaneous"))
    M, T1, T2 = "Miscellaneous!", "'Task 1'!", "'Task 2'!"

    # ============================================================ Miscellaneous (inputs first: other tabs read them)
    put(mi, "A1", "Miscellaneous: assumptions, products and fees, benchmarks, growth path, brief checklist, sources and longlist", F_TITLE, wrap=False)
    put(mi, "A3", "A. Assumptions (yellow cells are inputs; every tab recalculates)", F_SECTION, wrap=False)
    header(mi, 4, ["Assumption", "Pessimistic", "BAU", "Optimistic", "Basis"])
    for i, r in enumerate(range(6, 17)):
        row = 5 + i
        label = PP.cell(r, 1).value
        cell(mi, row, 1, label)
        fmt = {5: "0.0", 6: PCT, 8: MONEY, 10: PCT, 12: MONEY, 15: PCT2}.get(row)
        for j, col in enumerate((2, 3, 4)):
            cell(mi, row, 2 + j, PP.cell(r, col).value, fmt=fmt, inp=True)
        basis = PP.cell(r, 5).value or ""
        basis = basis.replace("columns T-U", "Task 1, section C").replace(
            "between @AiSlimx (~$8K a month, OKX Wallet Boost) and", "above @AiSlimx (~$550 a day on OKX Wallet Boost) and below")
        if label.startswith("Power traders per partner"):
            label = "FnF power traders per partner (scales each partner's own count in Task 1)"
            mi.cell(row, 1).value = label
        cell(mi, row, 5, basis)
    # rows: 5 FTD mult, 6 share trading, 7 retail trades, 8 retail size, 9 power traders, 10 share active,
    #       11 power trades, 12 power size, 13 start day, 14 days in month, 15 net revenue rate
    put(mi, "A17", "B. Trade profiles by segment (BAU; pessimistic and optimistic scale every segment by the ratios in row 24)", F_SECTION, wrap=False)
    header(mi, 18, ["Segment", "Retail trades per active day", "Retail average trade", "Retail $ a day per FTD",
                    "Power trades per active day", "Power average trade", "Power $ a day per trader", "Basis"])
    segs = [("Memecoin trenchers", "=C7", "=C8", "=C11", "=C12", "Rows 7-12"),
            ("DeFi / LP users", 3, 50, 10, 150, "Assumption: fewer, larger swaps when LPs rebalance")]
    for i, (name, rt, rs, pt, ps, basis) in enumerate(segs):
        r = 19 + i
        cell(mi, r, 1, name)
        cell(mi, r, 2, rt, inp=not str(rt).startswith("="))
        cell(mi, r, 3, rs, fmt=MONEY, inp=not str(rs).startswith("="))
        cell(mi, r, 4, f"=B{r}*C{r}*$C$6", fmt=MONEY)
        cell(mi, r, 5, pt, inp=not str(pt).startswith("="))
        cell(mi, r, 6, ps, fmt=MONEY, inp=not str(ps).startswith("="))
        cell(mi, r, 7, f"=E{r}*F{r}*$C$10", fmt=MONEY)
        cell(mi, r, 8, basis)
    header(mi, 23, ["Scenario scale vs BAU", "Retail, pessimistic", "Retail, optimistic", "Power, pessimistic", "Power, optimistic"])
    cell(mi, 24, 1, "Scale")
    for col, f in zip("BCDE", ["=(B6*B7*B8)/(C6*C7*C8)", "=(D6*D7*D8)/(C6*C7*C8)",
                              "=(B10*B11*B12)/(C10*C11*C12)", "=(D10*D11*D12)/(C10*C11*C12)"]):
        c = cell(mi, 24, "ABCDE".index(col) + 1, f, fmt="0.00")
    SEG = f"{M}$A$19:$A$20"

    # C. products and fees
    put(mi, "A26", "C. OKX products and fees (published fee vs the brief's 0.04%)", F_SECTION, wrap=False)
    PR = L["Products & revenue"]
    header(mi, 27, [PR.cell(2, c).value for c in range(1, 10)])
    for i, r in enumerate(range(3, 13)):
        row = 28 + i
        for c in range(1, 10):
            v = PR.cell(r, c).value
            if c == 6:
                v = f"=D{row}*E{row}"
            elif c == 7:
                v = f"=E{row}*$C$15"
            elif c == 8 and r == 9:
                v = "Multi-chain traders moving funds between chains"
            elif c == 8 and r == 11:
                v = "Not a target: fees go to a rival venue"
            fmt = MONEY2 if c in (6, 7) else (PCT2 if c == 4 or (c == 3 and isinstance(v, (int, float))) else (MONEY if c == 5 else None))
            cell(mi, row, c, v, fmt=fmt, inp=c in (4, 5) and v not in (None, ""))
    MEME_FEE = f"{M}$D$28"
    header(mi, 39, ["OKX DEX actuals (DefiLlama, mid-Nov 2026)", "Value", "Note"])
    for i, r in enumerate(range(15, 18)):
        row = 40 + i
        cell(mi, row, 1, PR.cell(r, 1).value)
        cell(mi, row, 2, "=B41/B40" if r == 17 else PR.cell(r, 2).value, fmt=PCT2 if r == 17 else MONEY)
        cell(mi, row, 3, PR.cell(r, 3).value)

    # D. rival benchmark
    put(mi, "A44", "D. What rival trading tools pay referrers (secondary sources, Nov 2026)", F_SECTION, wrap=False)
    header(mi, 45, [OF.cell(34, c).value for c in range(1, 6)])
    for i, r in enumerate(range(35, 40)):
        row = 46 + i
        cell(mi, row, 1, OF.cell(r, 1).value)
        cell(mi, row, 2, OF.cell(r, 2).value, fmt=PCT2, inp=True)
        cell(mi, row, 3, OF.cell(r, 3).value, fmt=PCT, inp=True)
        cell(mi, row, 4, f"=B{row}*C{row}*1000000", fmt=MONEY)
        cell(mi, row, 5, OF.cell(r, 5).value)
    put(mi, "A51", OF.cell(40, 1).value, wrap=False)

    # E. growth path
    put(mi, "A53", "E. How the number grows (0.04%, before partner shares; each partner's first 30 days live, BAU)", F_SECTION, wrap=False)
    header(mi, 54, ["Step", "Swaps", "Revenue (0.04%)", "Why it's realistic"])
    steps = [
        ("1. Posts first; FnF power traders join mid-month (day 15)",
         f"={T1}K45+{T1}L45*($C$14-15+1)/($C$14-$C$13+1)", "The usual sequence: content first, intros later"),
        ("2. FnF power traders first (onboarded by day 4): our plan", f"={T1}M45",
         "FnF power traders carry most swaps, so every earlier day counts"),
        ("3. 10 FnF power traders (not 5) at the 3 partners with books or clans",
         "=B56+3*5*$G$19*($C$14-$C$13+1)", "0xRusdi's BasedBot book, Ponyin's trading circle, Tomket Lovers' 13-15-member clan"),
        ("4. Plus the Partner B pilot ($2M a month, brief)", f"=B57+{T2}C5", "The one pilot the brief allows; positive on its own terms"),
        ("5. $10K a month", "=10000/$C$15", "A and C on revenue share only if their 90-day data holds, plus more partners at this profile: a month 4-6 goal"),
    ]
    for i, (a, b, d) in enumerate(steps):
        row = 55 + i
        cell(mi, row, 1, a)
        cell(mi, row, 2, b, fmt=MONEY)
        cell(mi, row, 3, f"=B{row}*$C$15", fmt=MONEY)
        cell(mi, row, 4, d)

    # F. brief checklist
    put(mi, "A61", "F. Brief checklist: where each request is answered", F_SECTION, wrap=False)
    checklist = [
        ("Format: up to 6 slides plus a prospect spreadsheet", "Deck (6 slides) + this workbook"),
        ("Public information only; no contact, spend or private contacts", "All tabs; slide 6"),
        ("Task 1: 8 potential partners", "Task 1, section A; slide 2"),
        ("Task 1: at least 6 whose main proposition is not airdrops", "Task 1, brief checks (row 21)"),
        ("Task 1: at least 3 partner types", "Task 1, column D (4 types)"),
        ("Task 1: at least 2 you could meet in person", "Task 1, column O; Task 4, in-person day"),
        ("Task 1: link, audience and market, evidence, route, proposed partnership, main risk, priority with reason", "Task 1, section A"),
        ("Task 1: top 3 and why you'd approach them first", "Task 1, rows 4-8; slide 2"),
        ("Task 1: verified facts vs assumptions; follower count alone is not evidence", "Task 1, columns H and I; deck footers; slide 6"),
        ("Task 2: first-month contribution for each partner, with the calculation", "Task 2, rows 3-7; slide 3"),
        ("Task 2: preferred pilot and reasoning", "Task 2, row 9"),
        ("Task 2: counter-offer or pilot structure", "Task 2, rows 10-11"),
        ("Task 2: evidence before signing", "Task 2, row 12"),
        ("Task 2: chosen partner's economics at 50% of forecast", "Task 2, columns K-L and rows 16-20"),
        ("Task 2: exclude other costs, then say which you'd investigate", "Task 2, row 13"),
        ("Task 2: pilot only one partner at first", "Task 2, row 14"),
        ("Task 3: use your highest-priority real prospect", "Task 3 (0xRusdi); slide 4"),
        ("Task 3: first outreach message of 120 words or fewer", "Task 3, row 4 (word count in the label)"),
        ("Task 3: follow-up after no response", "Task 3, row 6"),
        ("Task 3: reply to 'another platform offers a higher commission'", "Task 3, row 7"),
        ("Task 3: activation plan for the first 14 days after agreement", "Task 3, rows 13-20"),
        ("Task 3: 3 most important first-meeting questions", "Task 3, rows 8-10"),
        ("Task 3: no unverified product, incentive or commercial promises", "Task 3, row 11"),
        ("Task 4: weekly targets for outreach, meetings, pilots and activated partners", "Task 4, section A; slide 5"),
        ("Task 4: one in-person prospecting day (where, whom, access, fallback)", "Task 4, section E; slide 5"),
        ("Task 4: response to 20 messages, 2 replies, 0 meetings", "Task 4, section D (outreach reply rate); slide 5"),
        ("Task 4: results that mean continue, renegotiate or stop", "Task 4, section E (day 30); slide 5"),
        ("Task 4: activated-partner definition", "Task 4, section B note and section D; slide 5"),
        ("Task 4: commitment statement", "Slide 6"),
        ("Cite the public sources you use", "Miscellaneous, section G; deck footers"),
        ("Explain how you used AI and what you verified yourself", "Slide 6"),
        ("Be ready to adapt when assumptions change", "Miscellaneous, section A (yellow inputs)"),
        ("Extra: Task 2's five answers for our real top 3", "Task 2, rows 22-33"),
        ("Extra: 100 prospects screened, with the reason each was cut", "Miscellaneous, section H"),
    ]
    header(mi, 62, ["#", "Brief request", "Where it's answered"] + (["Status"] if research else []))
    for i, (req, where) in enumerate(checklist):
        row = 63 + i
        cell(mi, row, 1, i + 1)
        cell(mi, row, 2, req)
        cell(mi, row, 3, where)
        if research:
            cell(mi, row, 4, "To fill" if where == "Slide 6" and ("commitment" in req or "AI" in req) else "Done")
    last_ck = 63 + len(checklist) - 1

    # G. sources
    g0 = last_ck + 2
    put(mi, f"A{g0}", "G. Sources", F_SECTION, wrap=False)
    sources = [
        ("OKX DEX fee schedule", "https://web3.okx.com/dex-fees"),
        ("OKX DEX interface-fee notice (effective 12 May 2026)", "https://web3.okx.com/help/update-on-okx-dex-interface-fee"),
        ("OKX DEX referral programme", "https://web3.okx.com/help/whats-dex-referral-program"),
        ("OKX Wallet Meme Mode", "https://www.okx.com/learn/meme-mode"),
        ("OKX DEX fees and volume (DefiLlama)", "https://defillama.com/protocol/okx-dex"),
        ("Partner public pages (8)", "Task 1, column C"),
        ("Meteora LP Army Jakarta (462 went)", "https://luma.com/76fvdt2j"),
        ("MET Bali (671 went)", "https://luma.com/pkr27xr4"),
        ("Superteam Indonesia Ecosystem Showcase (502 went)", "https://luma.com/7gbaanxq"),
        ("Superteam Hub Jakarta (with Indonesia Crypto Hub)", "https://luma.com/2fpz2q5p"),
        ("W3AI Wednesday at JURA (weekly)", "https://luma.com/zkq937gq"),
        ("GuarEmperor wallet lists (BasedBot, GMGN)", "https://github.com/ariefzzz5421/wallet-tracker-transfer"),
        ("ICN: 10 X crypto influencers in Indonesia", "https://indonesiacrypto.network/blog/10-x-twitter-crypto-influencers-growing-the-indonesian-cryptosphere"),
        ("ICN community sizes (Mar 2026)", "https://indonesiacrypto.network/blog/crypto-community-indonesia"),
        ("POJK 6/2026 finfluencer certification", "https://economy.okezone.com/read/2026/07/08/320/3228975/ojk-wajibkan-sertifikasi-influencer-kripto-ini-tujuannya"),
        ("Law 4/2026 (OJK powers over offshore platforms)", "https://www.hsfkramer.com/ru/notes/indonesia/2026-posts/indonesia-reshapes-digital-asset-and-cryptocurrency-framework-with-2026-financial-services-omnibus-law"),
        ("Akademi Crypto inquiry halted (Oct 2026)", "https://justisio.com/berita/2026-10-26-polda-metro-hentikan-penyelidikan-laporan-dugaan-penipuan-kripto-timot"),
        ("OKX Alpha Traders Summit, Bali", "https://bitpinas.com/feature/okx-alpha-traders-summit/"),
        ("Rival referral rates", "pump.fun: PrimeXBT, Crypto Briefing; Axiom: affiliate guides (unconfirmed); fomo: Datawallet"),
        ("Indonesian X evidence", "@0xRusdi BasedBot dashboard; @0xSarimin Trojan referrals; @AiSlimx OKX Wallet Boost; @Ponyin Trojan results; @TomketLovers fomo clan; @cryptondo OKX Wallet use"),
    ]
    header(mi, g0 + 1, ["Topic", "Source"])
    for i, (topic, url) in enumerate(sources):
        row = g0 + 2 + i
        cell(mi, row, 1, topic)
        c = cell(mi, row, 2, url)
        if url.startswith("http"):
            c.hyperlink = url
            c.font = F_LINK
    h0 = g0 + 2 + len(sources) + 1

    # H. longlist
    put(mi, f"A{h0}", "H. Longlist: 100 public Indonesian prospects screened", F_SECTION, wrap=False)
    put(mi, f"A{h0 + 1}", "Each prospect checked against the same filter: proof of trading, FnF power traders it can bring, rival flow to move, trencher fees.", wrap=False)
    LL = L["Longlist (100 screened)"]
    rows = []
    for r in range(15, 115):
        rec = [LL.cell(r, c).value for c in range(1, 11)]
        name = rec[1]
        if name == "vemi.pramudya":
            rec = [None, partners[4]["name"], "Trading creator", partners[4]["link"], "Memecoin trenchers", "Y", "In the 8",
                   "#5: FnF circle plus followers who copy his wallet lists on BasedBot and GMGN", "Wave 2 (day 17)",
                   "https://github.com/ariefzzz5421/wallet-tracker-transfer"]
        elif name == "Hyperliquid Indonesia (community)":
            rec[6], rec[7], rec[8] = "Excluded by rule", "Rival perp venue: OKX can't send traffic there", ""
        elif name == "Superteam Indonesia":
            rec[6], rec[7], rec[8] = "In the 8", "#8: meetable every week in Jakarta; Solana users already onchain", "Wave 3 (day 24): co-hosted trader night"
        rows.append(rec)
    order = {"In the 8": 0, "Bench (month 2)": 1, "Access channel": 2, "Lead to verify": 3, "Excluded by rule": 4, "Legal hold": 5, "Screened out": 6}
    rows.sort(key=lambda x: (order[x[6]], str(x[1]).lower()))
    header(mi, h0 + 3, ["Screen result", "Count"])
    for i, res in enumerate(order):
        cell(mi, h0 + 4 + i, 1, res)
        cell(mi, h0 + 4 + i, 2, f'=COUNTIF($G${h0 + 14}:$G${h0 + 113},A{h0 + 4 + i})')
    cell(mi, h0 + 11, 1, "Total", bold=True)
    cell(mi, h0 + 11, 2, f"=SUM(B{h0 + 4}:B{h0 + 10})", bold=True)
    header(mi, h0 + 13, ["#", "Prospect", "Type", "Public link", "Segment", "Onchain users? (Y/N)", "Screen result", "Why", "Use in month 1-2", "Source"])
    for i, rec in enumerate(rows):
        row = h0 + 14 + i
        rec[0] = i + 1
        for c, v in enumerate(rec, 1):
            cell(mi, row, c, v)
    widths(mi, {"A": 44, "B": 30, "C": 30, "D": 30, "E": 34, "F": 16, "G": 18, "H": 50, "I": 40, "J": 44})
    mi.freeze_panes = "A3"

    # ============================================================ Task 1
    put(t1, "A1", "Task 1: 8 partners beyond airdrop communities", F_TITLE, wrap=False)
    put(t1, "A2", "Each partner's first 30 days live, base case (BAU); revenue at the brief's 0.04%. Yellow cells are inputs.", wrap=False)
    put(t1, "A4", "Why these 3 first", F_BOLD, wrap=False)
    put(t1, "B4", "Filter, in order: 1 Proof they already trade → 2 FnF power traders they bring (~90% of swaps) → 3 Rival flow we move to OKX → 4 Trencher fees: 0.5%", F_BOLD, wrap=False)
    lines = [
        ("1. 0xRusdi", "Proof: $8,463 of BasedBot referral earnings at the 25% tier, the only public proof of a referred trading book · FnF: his referral book · Flow: BasedBot users → OKX", 37),
        ("2. Crypto Ndo", "Proof: says he trades only on OKX Wallet; largest trencher audience on X (59.3K) · FnF: traders in his circle · Flow: his fomo referral links → OKX", 38),
        ("3. Ponyin", "Proof: ~$170K of Trojan volume posted for March 2026 · FnF: his trading circle · Flow: Trojan users → OKX", 39),
    ]
    for i, (a, b, r) in enumerate(lines):
        put(t1, f"A{5 + i}", a, F_BOLD, wrap=False)
        put(t1, f"B{5 + i}", f'="{b} · Trencher: "&TEXT(G{r},"0")&" FTDs, "&TEXT(I{r},"0")&" trades a day, $"&TEXT(M{r}/1000,"0")&"K swaps"', wrap=False)
    put(t1, "A8", "Reach alone isn't enough", F_BOLD, wrap=False)
    put(t1, "B8", "TwipsX (#4) brings more FTDs than any top-3 partner but has no proof its readers trade; Indonesia's biggest crypto names (100 screened, Miscellaneous section H) have none either.", wrap=False)

    put(t1, "A10", "A. The 8 partners (all fields the brief asks for)", F_SECTION, wrap=False)
    header(t1, 11, ["#", "Partner", "Public link", "Partner type", "Priority", "Priority reason", "Audience and market",
                    "Evidence of fit (verified)", "Assumptions to verify", "Acquisition route", "Proposed partnership",
                    "Main risk", "FnF / power traders they bring", "Rival tool their traders use now", "Meetable in person", "Airdrop-led"])
    for i, p in enumerate(partners):
        r = 12 + i
        vals = [i + 1, p["name"], p["link"], p["type"], f'=IF(J{26 + i}<=3,"High",IF(J{26 + i}<=7,"Medium","Low"))', p["reason"],
                p["audience"], p["evidence"], p["assumptions"], p["route"], p["proposal"], p["risk"], p["fnf"], p["vampire"],
                p["in_person"], p["airdrop"]]
        for c, v in enumerate(vals, 1):
            cc = cell(t1, r, c, v)
            if c == 3:
                cc.hyperlink = str(v).split(" ; ")[0]
                cc.font = F_LINK
    cell(t1, 21, 1, "Brief checks", bold=True)
    cell(t1, 21, 2, "Partners")
    cell(t1, 21, 3, "=COUNTA(B12:B19)")
    cell(t1, 21, 4, "Not airdrop-led (6+ needed)")
    cell(t1, 21, 5, '=COUNTIF(P12:P19,"No")')
    cell(t1, 21, 6, "Meetable in person (2+ needed)")
    cell(t1, 21, 7, '=COUNTIF(O12:O19,"Yes")')
    cell(t1, 21, 8, "Partner types (3+ needed): 4 = X trading creators, a trading clan, an LP community, a Solana ecosystem community")

    put(t1, "A23", "B. Ranking (1-5 per criterion; yellow = inputs)", F_SECTION, wrap=False)
    header(t1, 24, ["Partner", "Proof of real trading", "FTD potential", "FnF / power-trader access", "Rival flow to move",
                    "Reachability", "Segment value", "Risk (5 = low)", "Weighted score", "Rank"])
    cell(t1, 25, 1, "Weight", bold=True)
    for c, w in enumerate([0.2, 0.2, 0.15, 0.15, 0.1, 0.1, 0.1], 2):
        cell(t1, 25, c, w, fmt=PCT, inp=True)
    cell(t1, 25, 9, "=SUM(B25:H25)", fmt=PCT)
    for i, p in enumerate(partners):
        r = 26 + i
        cell(t1, r, 1, f"=B{12 + i}")
        for c, s in enumerate(p["scores"], 2):
            cell(t1, r, c, s, inp=True)
        cell(t1, r, 9, f"=SUMPRODUCT(B{r}:H{r},$B$25:$H$25)", fmt="0.00")
        cell(t1, r, 10, f"=RANK(I{r},$I$26:$I$33)")

    put(t1, "A35", "C. First 30 days live, by partner (BAU unless stated; revenue at the brief's 0.04%)", F_SECTION, wrap=False)
    header(t1, 36, ["Partner", "Segment", "Audience (post readers or attendees)", "FTD rate", "Days an FTD trades", "Partner share",
                    "FTDs", "FnF power traders", "Trades a day (day 30)", "Swaps a day (day 30)", "Retail swaps",
                    "FnF power-trader swaps", "Swaps (BAU)", "Swaps (pessimistic)", "Swaps (optimistic)", "Revenue (0.04%)",
                    "Paid to partner", "OKX keeps", "Audience basis"])
    for i, p in enumerate(partners):
        r = 37 + i
        seg = f"MATCH($B{r},{SEG},0)"
        rd, pd = f"INDEX({M}$D$19:$D$20,{seg})", f"INDEX({M}$G$19:$G$20,{seg})"
        rt, pt = f"INDEX({M}$B$19:$B$20,{seg})", f"INDEX({M}$E$19:$E$20,{seg})"
        cell(t1, r, 1, f"=B{12 + i}")
        cell(t1, r, 2, p["segment"], inp=True)
        cell(t1, r, 3, p["aud_n"], fmt="#,##0", inp=True)
        cell(t1, r, 4, p["rate"], fmt=PCT, inp=True)
        cell(t1, r, 5, p["days"], inp=True)
        cell(t1, r, 6, p["share"], fmt=PCT, inp=True)
        cell(t1, r, 7, f"=C{r}*D{r}*{M}$C$5", fmt=NUM1)
        cell(t1, r, 8, p["ptraders"], inp=True)
        cell(t1, r, 9, f"=G{r}*{M}$C$6*{rt}+H{r}*{M}$C$10*{pt}", fmt=NUM1)
        cell(t1, r, 10, f"=G{r}*{rd}+H{r}*{pd}", fmt=MONEY)
        cell(t1, r, 11, f"=G{r}*{rd}*E{r}", fmt=MONEY)
        cell(t1, r, 12, f"=H{r}*{pd}*({M}$C$14-{M}$C$13+1)", fmt=MONEY)
        cell(t1, r, 13, f"=K{r}+L{r}", fmt=MONEY)
        cell(t1, r, 14, f"=C{r}*D{r}*{M}$B$5*{rd}*{M}$B$24*E{r}+H{r}*{M}$B$9/{M}$C$9*{pd}*{M}$D$24*({M}$B$14-{M}$B$13+1)", fmt=MONEY)
        cell(t1, r, 15, f"=C{r}*D{r}*{M}$D$5*{rd}*{M}$C$24*E{r}+H{r}*{M}$D$9/{M}$C$9*{pd}*{M}$E$24*({M}$D$14-{M}$D$13+1)", fmt=MONEY)
        cell(t1, r, 16, f"=M{r}*{M}$C$15", fmt=MONEY2)
        cell(t1, r, 17, f"=P{r}*F{r}", fmt=MONEY2)
        cell(t1, r, 18, f"=P{r}-Q{r}", fmt=MONEY2)
        cell(t1, r, 19, p["basis"])
    cell(t1, 45, 1, "Total, 8 partners", bold=True)
    for c in range(7, 19):
        if c == 6:
            continue
        col = "ABCDEFGHIJKLMNOPQRS"[c - 1]
        cell(t1, 45, c, f"=SUM({col}37:{col}44)", fmt=t1.cell(37, c).number_format, bold=True)
    checks = [
        ("FnF power traders' share of swaps", "=L45/M45", PCT),
        ("Swaps if no FnF power traders arrive (retail FTDs only)", "=K45", MONEY),
        ("Revenue in that case (0.04%)", f"=K45*{M}$C$15", MONEY),
        ("Retail swaps per FTD in month 1", "=K45/G45", MONEY),
        ("Monthly swaps after day 30 (run-rate)", "=J45*30", MONEY),
        ("Published-fee upside before partner shares (trenchers 0.5%; LP at the blend in C53)",
         f'=SUMIF(B37:B44,"Memecoin trenchers",M37:M44)*{MEME_FEE}+SUMIF(B37:B44,"DeFi / LP users",M37:M44)*C53', MONEY),
    ]
    put(t1, "A47", "Checks", F_BOLD, wrap=False)
    for i, (a, f, fmt) in enumerate(checks):
        cell(t1, 48 + i, 1, a)
        cell(t1, 48 + i, 2, f, fmt=fmt)
    cell(t1, 53, 3, 0.0015, fmt=PCT2, inp=True)
    put(t1, "A55", "Where prospects come from: Telegram directories and t.me previews, X (KOL posts, PnL and referral screenshots, "
                   "public wallet lists), Luma event pages, rival ambassador and campaign lists (Bitget Wallet, BasedBot, Trojan, "
                   "fomo clans, pump.fun), GMGN and Kolscan KOL wallets. Hyperliquid Indonesia and Bael calls were screened out "
                   "as rival-venue and rival-wallet ties.", wrap=False)
    widths(t1, {"A": 30, "B": 30, "C": 30, "D": 26, "E": 12, "F": 34, "G": 30, "H": 46, "I": 34, "J": 26, "K": 44, "L": 34,
                "M": 34, "N": 30, "O": 14, "P": 12, "Q": 14, "R": 14, "S": 44})
    t1.freeze_panes = "C12"

    # ============================================================ Task 2
    LT2 = L["Task 2 numbers"]
    put(t2, "A1", "Task 2: one pilot from partners A, B and C (net revenue rate 0.04%, from the brief)", F_TITLE, wrap=False)
    header(t2, 3, [LT2.cell(3, c).value for c in range(1, 13)])
    for i, r in enumerate((4, 5, 6)):
        row = 4 + i
        cell(t2, row, 1, LT2.cell(r, 1).value)
        cell(t2, row, 2, LT2.cell(r, 2).value)
        cell(t2, row, 3, LT2.cell(r, 3).value, fmt=MONEY, inp=True)
        cell(t2, row, 4, f"=C{row}*{M}$C$15", fmt=MONEY)
        cell(t2, row, 5, LT2.cell(r, 5).value, fmt=MONEY, inp=True)
        cell(t2, row, 6, LT2.cell(r, 6).value, fmt=PCT, inp=True)
        cell(t2, row, 7, f"=D{row}-E{row}-D{row}*F{row}", fmt=MONEY)
        cell(t2, row, 8, LT2.cell(r, 8).value, fmt=PCT, inp=True)
        cell(t2, row, 9, f"=D{row}-D{row}*H{row}", fmt=MONEY)
        cell(t2, row, 10, f'=IF(E{row}=0,"No fixed fee",E{row}/({M}$C$15*(1-F{row})))', fmt=MONEY)
        cell(t2, row, 11, f"=D{row}/2-E{row}-D{row}/2*F{row}", fmt=MONEY)
        cell(t2, row, 12, f"=D{row}/2-D{row}/2*H{row}", fmt=MONEY)
    cell(t2, 7, 1, "Total", bold=True)
    for col in "CDEGIKL":
        cell(t2, 7, "ABCDEFGHIJKL".index(col) + 1, f"=SUM({col}4:{col}6)", fmt=MONEY, bold=True)
    for r in range(9, 15):
        cell(t2, r, 1, LT2.cell(r, 1).value, bold=True)
        cell(t2, r, 2, LT2.cell(r, 2).value)
    put(t2, "A16", "Partner B's own payout, and the step-up case", F_SECTION, wrap=False)
    header(t2, 17, ["Case", "Full volume", "At 50% volume"])
    for i, (a, b, c) in enumerate([("B's payout on its ask (35%)", "=D5*F5", "=D5/2*F5"),
                                   ("B's payout on our counter (25%)", "=D5*H5", "=D5/2*H5"),
                                   ("OKX month 1 if B reaches the 35% step-up inside month 1", "=D5-D5*F5", "=D5/2-D5/2*F5")]):
        cell(t2, 18 + i, 1, a)
        cell(t2, 18 + i, 2, b, fmt=MONEY)
        cell(t2, 18 + i, 3, c, fmt=MONEY)
    put(t2, "A22", "Our real top 3 on the same test (each partner's first 30 days live, BAU, 0.04%)", F_SECTION, wrap=False)
    header(t2, 23, ["Partner", "Swaps, first 30 days (BAU)", "Revenue (0.04%)", "Fixed fee (our offer)", "Revenue share (our offer)",
                    "1. Month-1 contribution", "5. At 50% of forecast", "Partner earns (0.04%)", "Partner earns at 50%",
                    "Volume the illustrative fixed fee (B33) needs to break even", "OKX keeps at the published 0.5% fee", "4. Evidence before signing"])
    evid = ["BasedBot dashboard export: referred traders, volume per trader, days active, chains; share of his referred traders in Indonesia",
            "fomo referral stats (active referrals, volume); share of followers trading daily; disclosure on his own launchpad (@flapdotsh)",
            "Trojan dashboard or wallet history for him and his circle; how many in the circle trade daily"]
    for k in range(3):
        r, src = 24 + k, 37 + k
        cell(t2, r, 1, f"={T1}A{src}")
        cell(t2, r, 2, f"={T1}M{src}", fmt=MONEY)
        cell(t2, r, 3, f"=B{r}*{M}$C$15", fmt=MONEY2)
        cell(t2, r, 4, 0, fmt=MONEY, inp=True)
        cell(t2, r, 5, f"={T1}F{src}", fmt=PCT)
        cell(t2, r, 6, f"=C{r}-D{r}-C{r}*E{r}", fmt=MONEY2)
        cell(t2, r, 7, f"=C{r}/2-D{r}-C{r}/2*E{r}", fmt=MONEY2)
        cell(t2, r, 8, f"=C{r}*E{r}", fmt=MONEY2)
        cell(t2, r, 9, f"=C{r}/2*E{r}", fmt=MONEY2)
        cell(t2, r, 10, f"=$B$33/({M}$C$15*(1-E{r}))", fmt=MONEY)
        cell(t2, r, 11, f"=B{r}*{MEME_FEE}*(1-E{r})", fmt=MONEY2)
        cell(t2, r, 12, evid[k])
    cell(t2, 27, 1, "Top 3 total", bold=True)
    for col in "BCDFGHIK":
        cell(t2, 27, "ABCDEFGHIJKL".index(col) + 1, f"=SUM({col}24:{col}26)", fmt=t2[f"{col}24"].number_format, bold=True)
    answers = [
        ("2. Preferred pilot", "All three go live together in wave 1 (day 10); the one-pilot limit applies to the brief's A, B and C. If only one could go first: 0xRusdi, the only partner with public proof of a referred trading book."),
        ("3. Pilot structure", "Revenue share only: 30% base, 40% once at least $1M of referred swaps and 50 wallets trading on 3+ days; follower discount up to 10% from the partner's share; 30-day clawback for wash trading, self-referral and sybil wallets; no fixed fee."),
        ("Why no fixed fee", '="At 0.04%, even a $"&TEXT(B33,"#,##0")&" fixed fee needs about $"&TEXT(J24/1000000,"0.0")&"M of swaps to break even, while each of these partners is forecast at about $"&TEXT(B27/3/1000,"0")&"K. A fixed fee only comes back with 90 days of data, and at published fees (column K)."'),
        ("5. If volume reaches only 50%", '="Contribution halves from $"&TEXT(F27,"0")&" to $"&TEXT(G27,"0")&" for the three but stays positive, because there is no fixed fee. Each partner\'s payout halves too ($"&TEXT(H24,"0")&" to $"&TEXT(I24,"0")&" for 0xRusdi at 0.04%), so the offer leans on support and FnF onboarding, not the rate."'),
    ]
    for i, (a, b) in enumerate(answers):
        cell(t2, 29 + i, 1, a, bold=True)
        cell(t2, 29 + i, 2, b)
    cell(t2, 33, 1, "Illustrative fixed fee to test", bold=True)
    cell(t2, 33, 2, 500, fmt=MONEY, inp=True)
    widths(t2, {"A": 30, "B": 30, "C": 14, "D": 14, "E": 14, "F": 14, "G": 16, "H": 16, "I": 16, "J": 18, "K": 16, "L": 44})

    # ============================================================ Task 3
    LT3 = L["Task 3 outreach"]
    put(t3, "A1", "Task 3: win and activate 0xRusdi, our #1 prospect", F_TITLE, wrap=False)
    header(t3, 3, ["Item", "Text"] + (["Word count"] if research else []))
    for r in range(4, 12):
        cell(t3, r, 1, LT3.cell(r, 1).value, bold=True)
        cell(t3, r, 2, LT3.cell(r, 2).value)
        if research and r <= 7:
            cell(t3, r, 3, f'=LEN(TRIM(B{r}))-LEN(SUBSTITUTE(TRIM(B{r})," ",""))+1')
    t3["A4"].value = '="First message (X DM), "&(LEN(TRIM(B4))-LEN(SUBSTITUTE(TRIM(B4)," ",""))+1)&" words"'
    t3["B11"].value = ("The messages promise no rates, rewards or features OKX hasn't confirmed. The fee share (up to 50% by tier) and "
                       "the follower discount are published OKX DEX referral features; any specific rate needs OKX approval.")
    put(t3, "A13", "14-day activation plan (day 1 = agreement signed)", F_SECTION, wrap=False)
    header(t3, 14, ["Days", "What happens", "Target"])
    for i, r in enumerate(range(14, 20)):
        row = 15 + i
        cell(t3, row, 1, LT3.cell(r, 1).value, bold=True)
        txt = (LT3.cell(r, 2).value or "").replace("his 5 most active referred traders", "his 5 most active referred traders (his FnF power traders)")
        cell(t3, row, 2, txt)
        cell(t3, row, 3, LT3.cell(r, 3).value)
    put(t3, "A22", "Offer to 0xRusdi (OKX DEX referral programme; any rate needs OKX approval)", F_SECTION, wrap=False)
    header(t3, 23, ["Term", "Value"])
    offer = [("Base revenue share", f"={T1}F37", PCT, True), ("Step-up share", 0.4, PCT, True),
             ("Step-up trigger", OF.cell(6, 6).value, None, False), ("Follower discount", "Up to 10%, from the partner's share", None, False),
             ("Swap fee his traders pay", "0.5% on memecoins (published OKX DEX fee)", None, False),
             ("What he needs most", OF.cell(19, 6).value, None, False), ("Support from us", OF.cell(19, 7).value, None, False),
             ("He earns in his first 30 days (published fee)", f"={T1}M37*{MEME_FEE}*{T1}F37", MONEY2, False),
             ("Rules for all partners", OF.cell(10, 1).value.replace("Rules for all: ", ""), None, False)]
    for i, (a, b, fmt, inp) in enumerate(offer):
        cell(t3, 24 + i, 1, a, bold=True)
        cell(t3, 24 + i, 2, b, fmt=fmt, inp=inp and not str(b).startswith("="))
    put(t3, "A34", "Support every partner gets", F_SECTION, wrap=False)
    for i, r in enumerate(range(26, 32)):
        cell(t3, 35 + i, 1, f"{i + 1}.")
        cell(t3, 35 + i, 2, OF.cell(r, 1).value)
    widths(t3, {"A": 36, "B": 110, "C": 40 if not research else 40})

    # ============================================================ Task 4
    put(t4, "A1", "Task 4: first 30 days (calendar days 1-30 of the job)", F_TITLE, wrap=False)
    put(t4, "A2", "Main metrics: FTDs, then trades a day; swaps and revenue follow. Partners go live in three waves: days 10, 17 and 24.", wrap=False)
    put(t4, "A4", "A. Weekly targets", F_SECTION, wrap=False)
    header(t4, 5, ["Target", "Week 1 (days 1-7)", "Week 2 (days 8-14)", "Week 3 (days 15-21)", "Week 4 (days 22-30)", "Day 30", "How it's set"])
    R0, R1 = 25, 32  # ramp rows (partners); 33 total; 34 Partner B
    weekly = [
        ("Prospects scored", 25, 20, 15, 10, "=SUM(B6:E6)", "Score everyone before messaging"),
        ("Outreach messages (warm intro first)", 25, 20, 15, 10, "=SUM(B7:E7)", "100% of scored"),
        ("Replies", 8, 6, 5, 3, "=SUM(B8:E8)", "~30% of messages"),
        ("Meetings (call or in person)", 3, 6, 5, 4, "=SUM(B9:E9)", "~80% of replies; the Jakarta day is in week 2"),
        ("Pilots live (running total, our 8)", 0, f'=COUNTIF($C${R0}:$C${R1},"<=14")', f'=COUNTIF($C${R0}:$C${R1},"<=21")',
         f'=COUNTIF($C${R0}:$C${R1},"<=30")', "=E10",
         "Waves: top 3 on day 10; TwipsX, GuarEmperor and the Meteora night on day 17; Tomket Lovers and the Superteam night on day 24. If one of the top 3 says no in week 1, TwipsX moves into wave 1"),
        ("Partner B pilot (the brief's hypothetical)", "Signing", "Live from day 10", "Live", "Live", "Live", "The one A/B/C pilot; counted separately because its volume is the brief's estimate"),
        ("Activated partners (running total, our 8)", 0, 0, 0, f'=COUNTIF(R{R0}:R{R1},"Yes")', "=E12", "FTDs >= 10, trades a day >= 50 and swaps >= $100K within 30 days live"),
        ("FTDs (running total)", 0, f"=SUM(H{R0}:H{R1})", f"=SUM(I{R0}:I{R1})", f"=SUM(J{R0}:J{R1})", "=E13", "Task 1 FTDs x share of each partner's month live; event FTDs on the night"),
        ("Trades a day (last day of the week)", 0, f"=SUM(K{R0}:K{R1})", f"=SUM(L{R0}:L{R1})", f"=SUM(M{R0}:M{R1})", "=E14", "FTDs x share trading x trades + FnF power traders x share active x trades"),
        ("Swaps from our 8 (running total)", 0, f"=SUM(N{R0}:N{R1})", f"=SUM(O{R0}:O{R1})", f"=SUM(P{R0}:P{R1})", "=E15", "Each partner's live days in the calendar month"),
        ("Swaps from Partner B (running total)", 0, "=N34", "=O34", "=P34", "=E16", "$2M a month from day 10"),
        ("Revenue at 0.04% (running total)", f"=(B15+B16)*{M}$C$15", f"=(C15+C16)*{M}$C$15", f"=(D15+D16)*{M}$C$15", f"=(E15+E16)*{M}$C$15", "=E17", "Swaps x 0.04%"),
        ("Revenue after partner shares (running total)", 0,
         f"=(SUMPRODUCT(N{R0}:N{R1},1-$G${R0}:$G${R1})+N34*(1-{T2}$H$5))*{M}$C$15",
         f"=(SUMPRODUCT(O{R0}:O{R1},1-$G${R0}:$G${R1})+O34*(1-{T2}$H$5))*{M}$C$15",
         f"=(SUMPRODUCT(P{R0}:P{R1},1-$G${R0}:$G${R1})+P34*(1-{T2}$H$5))*{M}$C$15", "=E18",
         "Revenue x (1 - each partner's share; B at our 25%)"),
        ("Monthly run-rate on day 30 (our 8 + B)", "", "", "", f"=(SUM(Q{R0}:Q{R1})+Q34)*30", "=E19",
         f'="Swaps a day on day 30 x 30 -> $"&TEXT(E19*{M}$C$15,"#,##0")&" a month at 0.04%"'),
    ]
    for i, rowvals in enumerate(weekly):
        r = 6 + i
        for c, v in enumerate(rowvals, 1):
            fmt = None
            if r in (13, 14):
                fmt = NUM1
            elif r in (15, 16, 19):
                fmt = MONEY
            elif r in (17, 18):
                fmt = MONEY2
            cell(t4, r, c, v, fmt=fmt if 2 <= c <= 6 else None, inp=(r <= 9 and 2 <= c <= 5), bold=(c == 1))

    put(t4, "A22", "B. Calendar ramp behind the weekly targets (links to Task 1)", F_SECTION, wrap=False)
    header(t4, 24, ["Partner", "Event night?", "Go-live day", "Live days by day 14", "Live days by day 21", "Live days by day 30",
                    "Partner share", "FTDs by day 14", "FTDs by day 21", "FTDs by day 30", "Trades a day, day 14",
                    "Trades a day, day 21", "Trades a day, day 30", "Swaps by day 14", "Swaps by day 21", "Swaps by day 30",
                    "Swaps a day on day 30", "Activated by day 30?"])
    for i, p in enumerate(partners):
        r, t = R0 + i, 37 + i
        seg = f"MATCH({T1}$B{t},{SEG},0)"
        rd, pd = f"INDEX({M}$D$19:$D$20,{seg})", f"INDEX({M}$G$19:$G$20,{seg})"
        rt, ptr = f"INDEX({M}$B$19:$B$20,{seg})", f"INDEX({M}$E$19:$E$20,{seg})"
        ftd, ptn, start, month = f"{T1}$G{t}", f"{T1}$H{t}", f"{M}$C$13", f"{M}$C$14"
        cell(t4, r, 1, f"={T1}A{t}")
        cell(t4, r, 2, p["event"], inp=True)
        cell(t4, r, 3, p["go_live"], inp=True)
        for k, d in enumerate((14, 21, 30)):
            cell(t4, r, 4 + k, f"=MAX(0,{d}-$C{r})")
        cell(t4, r, 7, f"={T1}F{t}", fmt=PCT)
        for k, live in enumerate("DEF"):
            L_ = f"{live}{r}"
            cell(t4, r, 8 + k, f'=IF({L_}=0,0,IF($B{r}="Yes",{ftd},{ftd}*{L_}/{month}))', fmt=NUM1)
            fcol = "HIJ"[k]
            cell(t4, r, 11 + k, f"=IF({L_}=0,0,{fcol}{r}*{M}$C$6*{rt}+IF({L_}>={start},{ptn}*{M}$C$10*{ptr},0))", fmt=NUM1)
            cell(t4, r, 14 + k, f'=IF({L_}=0,0,IF($B{r}="Yes",{ftd}*{rd}*{L_},{ftd}*{rd}*{L_}^2/(2*{month}))+{ptn}*{pd}*MAX(0,{L_}-{start}+1))', fmt=MONEY)
        cell(t4, r, 17, f"=IF(F{r}=0,0,J{r}*{rd}+IF(F{r}>={start},{ptn}*{pd},0))", fmt=MONEY)
        cell(t4, r, 18, f'=IF(AND(J{r}>=10,M{r}>=50,P{r}>=100000),"Yes","No")')
    cell(t4, 33, 1, "Total, our 8", bold=True)
    for c in list(range(8, 18)):
        col = "ABCDEFGHIJKLMNOPQR"[c - 1]
        cell(t4, 33, c, f"=SUM({col}{R0}:{col}{R1})", fmt=t4.cell(R0, c).number_format, bold=True)
    cell(t4, 34, 1, "Partner B pilot (brief)")
    cell(t4, 34, 2, "No")
    cell(t4, 34, 3, 10, inp=True)
    for k, d in enumerate((14, 21, 30)):
        cell(t4, 34, 4 + k, f"=MAX(0,{d}-C34)")
        cell(t4, 34, 14 + k, f"={T2}$C$5/{M}$C$14*{'DEF'[k]}34", fmt=MONEY)
    cell(t4, 34, 17, f"={T2}$C$5/{M}$C$14", fmt=MONEY)
    put(t4, "A35", "Activated partner: at least 10 FTDs, at least 50 trades a day, at least $100K of swaps, at least 30% of FTDs still "
                   "trading in days 15-30 and under 5% flagged wallets. The ramp checks the first three; retention and flags come from the dashboard.", wrap=False)

    put(t4, "A37", "C. What each partner delivers in its first 30 days live", F_SECTION, wrap=False)
    header(t4, 38, ["Partner", "Deliverables", "FnF power traders", "FTD target (day 30)", "Trades a day (day 30)", "Swaps (first 30 days)"])
    for i, p in enumerate(partners):
        r, t = 39 + i, 37 + i
        cell(t4, r, 1, f"={T1}A{t}")
        cell(t4, r, 2, p["deliverables"])
        cell(t4, r, 3, f"={T1}H{t}")
        cell(t4, r, 4, f"={T1}G{t}", fmt=NUM1)
        cell(t4, r, 5, f"={T1}I{t}", fmt=NUM1)
        cell(t4, r, 6, f"={T1}M{t}", fmt=MONEY)

    put(t4, "A48", "D. What we monitor, and what we change when it slips", F_SECTION, wrap=False)
    header(t4, 49, [DP.cell(35, c).value for c in range(1, 8)])
    for i, r in enumerate(range(36, 45)):
        row = 50 + i
        for c in range(1, 8):
            v = DP.cell(r, c).value
            if isinstance(v, str):
                v = v.replace("Power traders active", "FnF power traders active").replace("Swaps per power trader", "Swaps per FnF power trader")
            cell(t4, row, c, v, bold=(c == 1))

    put(t4, "A60", "E. Operating cadence, in-person day and day-30 decisions", F_SECTION, wrap=False)
    header(t4, 61, ["When", "What", "Output"])
    cadence = [(DP.cell(r, 1).value, DP.cell(r, 2).value, DP.cell(r, 3).value) for r in range(47, 53)]
    cadence[3] = ("Week 2, Wednesday",
                  "In-person day in Jakarta, set on a Wednesday so a weekly event is always on. 10:00 Superteam Hub Jakarta (run "
                  "with Indonesia Crypto Hub): agree the Solana trader night with the hub team, arranged through the Luma host page "
                  "5 days ahead. 13:00-15:00 two coffee meetings with Jakarta-based trench creators from the pipeline. 19:00 the next "
                  "LP Army night if one is listed on Luma; otherwise W3AI Wednesday at JURA Kemanggisan (weekly). Fallback: same-day "
                  "video call, the next Jakarta prospect on the list, and the evening event regardless", "Meetings held even if some cancel")
    for i, (a, b, c) in enumerate(cadence):
        row = 62 + i
        cell(t4, row, 1, a, bold=True)
        cell(t4, row, 2, (b or "").replace("power-trader track", "FnF power-trader track"))
        cell(t4, row, 3, c)
    widths(t4, {"A": 34, "B": 60, "C": 16, "D": 18, "E": 18, "F": 18, "G": 40, "H": 12, "I": 12, "J": 12, "K": 12, "L": 12,
                "M": 12, "N": 14, "O": 14, "P": 14, "Q": 14, "R": 14})
    t4.freeze_panes = "B6"

    # ============================================================ research-only tabs
    if research:
        rv = wb.create_sheet("Checks and history")
        put(rv, "A1", "Checks and history (research copy only)", F_TITLE, wrap=False)
        put(rv, "A3", "A. Verify before sending", F_SECTION, wrap=False)
        header(rv, 4, ["#", "Check", "Where it shows"])
        todo = [
            ("Re-open the 8 partner links and re-check follower counts and quoted posts on-platform", "Task 1, columns C and H"),
            ("GuarEmperor: confirm his FnF circle (who, how many trade daily) and his current follower count (63K vs 90K snapshots)", "Task 1 row 16; ranking row 30 (FnF score 3 until confirmed)"),
            ("Superteam Indonesia: confirm a trader night is possible and its capacity (model assumes ~150)", "Task 1 row 19 and row 44"),
            ("OKX DEX referral tiers: confirm the live tier table (regional help pages differ)", "Task 3 offer; deck footers"),
            ("Meme Mode limit orders: confirm before mentioning on a call", "Miscellaneous, products table"),
            ("BasedBot fee and 0xRusdi's tier history behind the ~$3.4M estimate", "Task 1 row 12, column I"),
            ("Legal: promotion of OKX DEX in Indonesia under Law 4/2026; partner POJK 6/2026 certification status", "Task 4, compliance gate"),
            ("Fill slide 6: commitment and the AI-use line", "Deck slide 6"),
        ]
        for i, (a, b) in enumerate(todo):
            cell(rv, 5 + i, 1, i + 1)
            cell(rv, 5 + i, 2, a)
            cell(rv, 5 + i, 3, b)
        put(rv, "A14", "B. Partner changes after review", F_SECTION, wrap=False)
        header(rv, 15, ["Change", "Why"])
        for i, (a, b) in enumerate([
                ("Hyperliquid Indonesia replaced by GuarEmperor and his FnF circle", "Hyperliquid is a rival venue: OKX can't send traffic there"),
                ("Bael calls replaced by Superteam Indonesia", "Bael is a Bitget Wallet Indonesia ambassador; Superteam keeps 2 partners meetable in person"),
                ("Power traders shown as FnF power traders, with a per-partner count", "Makes the FnF circle each partner brings visible; Superteam set to 2 as a builder-led community"),
                ("Four cryptokolz.com-only prospects removed from the longlist", "Promotional directory; replaced by four licensed exchanges screened out as competitors")]):
            cell(rv, 16 + i, 1, a)
            cell(rv, 16 + i, 2, b)
        put(rv, "A21", "C. Review findings (from REVIEW.md)", F_SECTION, wrap=False)
        header(rv, 22, ["#", "Where", "What was wrong", "Fix"])
        r = 23
        for line in REVIEW.read_text().splitlines():
            m = re.match(r"\| (E\d+) \| (.*?) \| (.*?) \| (.*?) \|", line)
            if m:
                for c, v in enumerate(m.groups(), 1):
                    cell(rv, r, c, v.replace("**", ""))
                r += 1
        widths(rv, {"A": 40, "B": 70, "C": 70, "D": 70})
    return wb


for research, name in ((False, "OKX_DEX_BD_Affiliate_Indonesia_Prospect_Spreadsheet.xlsx"), (True, "OKX_research_and_checks.xlsx")):
    wb = build(research)
    wb.save(OUTDIR / name)
    print("saved", name)
