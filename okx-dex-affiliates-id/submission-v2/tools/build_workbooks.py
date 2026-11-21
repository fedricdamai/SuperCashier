"""Build the submission workbook (Task 1-4 + Miscellaneous) and the research workbook (same + research tabs).

Usage: python build_workbooks.py <legacy_model.xlsx> <research_sheet.xlsx> <REVIEW.md> <out_dir>

The volume model is the research sheet's method (FTDs = reach pool P x sign-up rate s x activation a; swaps = FTDs x
monthly swaps per FTD v, at low/base/high rates) plus an FnF line: the power traders each partner introduces in days 1-3.
Every number is a live formula on the yellow inputs.
"""
import re
import sys
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

LEGACY, RESEARCH, REVIEW, OUTDIR = map(Path, sys.argv[1:5])
L = load_workbook(LEGACY, data_only=True)
R = load_workbook(RESEARCH, data_only=True)

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
MONEY, MONEY2, PCT, PCT1, PCT2, NUM0, NUM1 = '$#,##0', '$#,##0.00', '0%', '0.0%', '0.00%', '#,##0', '0.0'


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


def link(c, url):
    c.hyperlink = url
    c.font = F_LINK


def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


def col(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


# ---------------------------------------------------------------- partners (rank order = score order)
SL = L["Shortlist (Task 1)"]
TREN, LP = "Memecoin trenchers", "DeFi / LP users"
partners = [
    dict(short="Kevin Susanto", name="Kevin Susanto (@KevinSusanto)",
         link="https://x.com/KevinSusanto ; https://t.me/KevinSusantoAlpha",
         type="Large KOL and active onchain trader; hub of Jakarta's degen circle",
         audience="Regional SEA and Indonesian crypto audience, English posts; Jakarta",
         evidence=("fomo Top Traders 24h PnL, 19-20 Nov 2026: #3, +$53,566 (screenshot); a tracker logged ~$402K across 99 of "
                   "his fomo trades; X 168,874, median 41,248 views a post; Telegram KevinSusantoAlpha 2,860; Degen Night speaker "
                   "(31 Oct) with Bevan and Brandon Salim; attends OKX events (annual event 7 Oct, OKX x McLaren)"),
         assumptions=("Engagement is low (0.84%), so the base FTD count may be high; whether his OKX link (okx.ac/join/VVIP7777) "
                      "is a CEX affiliate deal OKX already pays, and which DEX volume would be new; his Lighter advisory terms"),
         route="Through OKX's regional marketing team, which already works with him; then DM; in person in Jakarta",
         proposal=("OKX DEX referral fee share on a separate DEX link, no fixed fee; 3 traders from his circle onboarded 1:1 in "
                   "days 1-3; a co-hosted onchain session for his Telegram; intros to Brandon Salim and the Degen Night circle"),
         risk="Already holds an OKX referral link, so DEX and CEX tracking must not overlap; advises Lighter, a rival perp DEX; runs volume giveaways",
         reason="Most FTDs of the 8, a verified top fomo trader, and the fastest yes: OKX's regional team already works with him",
         vampire="fomo (fee code KevinSusanto); Binance, Bitget and Tokocrypto links",
         fnf="His Jakarta trading circle and KevinSusantoAlpha Telegram: 3 traders onboarded first",
         in_person="Yes", airdrop="No (volume giveaways)", segment=TREN, pool=41820, ptraders=3, share=0.3, members=2860, per_member="trade",
         basis="Measured (research): median 41,248 views a post + 20% of 2,860 Telegram subscribers",
         deliverables="Days 1-3: 3 traders from his circle onboarded 1:1; day 5: co-hosted onchain session for KevinSusantoAlpha; 2 market threads with the link",
         need="Clear, non-exclusive terms next to his existing exchange deals, with DEX tracking kept apart from his OKX link",
         support="Tracking agreed with OKX's regional team; one co-hosted onchain session; Bahasa and English Meme Mode guide; wallet-level referral dashboard",
         check="fomo referral stats (active referrals, volume); with OKX's regional team, what his okx.ac link already earns and which DEX volume would be new; Lighter advisory terms",
         line="Proof: #3 on fomo's 24h PnL board (+$53.6K, 19-20 Nov); ~$402K across 99 fomo trades (tracker) · FnF: his Jakarta circle and Telegram · Flow: fomo (code KevinSusanto) → OKX · Meet: Jakarta, through OKX's regional team",
         scores=(4, 5, 4, 4, 5, 5, 3), go_live=10, event="No"),
    dict(short="Bevan / Narrative Labs", name="Bevan Pramudito / Narrative Labs (@bevanpramudt)",
         link="https://x.com/bevanpramudt ; https://narrativelabs.id",
         type="Trader-educator with a paid trading community and own tools",
         audience="Intermediate and advanced Indonesian memecoin and onchain traders (its site: 'advanced, not for beginners'); Jakarta",
         evidence=("Published wallet 8DC3hp...qGxg (Solscan): 91 swaps from 14 May to 30 Jul 2026, 352.6 SOL bought and 726.8 SOL "
                   "sold (~1,080 SOL), +374 SOL net; 'The Kingslayer' at Clash of Trenchers (W3W, 12 Nov): one of 3 traders to turn "
                   "1 SOL into 10 SOL in 20 minutes; Discord 9,238 members, 984 online (20 Nov 2026); paid Masterclass (29 of 40 "
                   "promo seats sold); member meetup in Jakarta on 29 Oct; X 26,983, median 7,545 views a post"),
         assumptions=("No swaps on the published wallet since 30 Jul: which wallet he trades from now; self-reported +Rp13.7B PnL "
                      "(Feb-Aug 2026); share of Discord members who trade weekly; GMGN and Mobee referral terms"),
         route="X DM, then a Narrative Labs session in Jakarta (in person)",
         proposal=("OKX DEX referral fee share with a follower discount; 5 Masterclass traders onboarded 1:1 in days 1-3; a "
                   "co-branded Discord workshop where he trades his own workflow on Meme Mode and Track; scope an OKX DEX API "
                   "trade button for his Narrative Alpha tool"),
         risk="Already pays out GMGN (BEVANP) and Mobee referral codes; published wallet idle since 30 Jul",
         reason="Verified onchain trading (wallet and a live 1-to-10 SOL run), a paid community he founded, GMGN flow to move, meetable in Jakarta",
         vampire="GMGN (code BEVANP); Mobee (code BEVAN)",
         fnf="His paid Masterclass traders: 5 onboarded first (29 seats sold in the promo batch)",
         in_person="Yes", airdrop="No", segment=TREN, pool=9240, ptraders=5, share=0.3, members=9238, per_member="trade",
         basis="Measured (research, 20 Nov 2026): median 7,545 + 711 (Narrative Labs) views a post + 984 Discord members online",
         deliverables=("Days 1-3: 5 Masterclass traders onboarded 1:1 in Jakarta; day 5: co-branded Discord workshop on Meme Mode "
                       "and Track, replay pinned; 2 X posts and a Narrative Labs announcement with the link"),
         need="A trading tool his paid members can't get elsewhere, plus income on the execution he already teaches",
         support=("OKX DEX API docs and a technical contact to scope a Narrative Alpha trade button; co-hosted Discord workshop; "
                  "Bahasa Meme Mode and Track guide; wallet-level referral dashboard"),
         check=("Which wallet he trades from now (the published one is idle since 30 Jul) and its history; Discord active-member "
                "stats; Masterclass enrolment and how many members trade weekly; GMGN (BEVANP) referral stats"),
         line="Proof: published wallet, ~1,080 SOL swapped May-Jul, +374 SOL; 1 SOL to 10 SOL live at W3W · FnF: his paid Masterclass traders · Flow: GMGN (code BEVANP) → OKX · Meet: Jakarta",
         scores=(5, 3, 5, 4, 5, 5, 3), go_live=10, event="No"),
    dict(short="GuarEmperor + FnF", name="GuarEmperor / Emperor Journal (@GuarEmperor) and his FnF circle",
         link="https://x.com/GuarEmperor ; https://x.com/emperorjournal_",
         type="Alpha caller with an FnF circle and a tools Discord",
         audience="Indonesian and global CT degens: memecoin traders on new chains, NFT minters, whitelist hunters",
         evidence=("X 94,421, median 15,885 views a post, 1.9% engagement (~47 replies a post); members post results from his "
                   "tools (GMGN +$802, +$387); public smart-money wallet lists formatted for BasedBot and GMGN; 'Enjoy being poor' "
                   "Telegram ~16.5K (third-party snapshot)"),
         assumptions="Size and activity of his FnF circle; Emperor Journal Discord size (invite-gated); share of followers who trade rather than farm",
         route="X DM; warm intro through All Good (@mudacumasekali), an Emperor Journal member; sister account @emperorjournal_",
         proposal=("OKX DEX referral fee share with a follower discount; his FnF circle (5 traders) onboarded 1:1 in days 1-3; OKX "
                   "DEX API for his Discord tools; alpha calls executed through Meme Mode and Track"),
         risk="WL, mint and airdrop posts in his feed; alpha calls need front-running and disclosure checks",
         reason="Best engagement among the big accounts and an FnF circle that trades from his tools",
         vampire="GMGN and BasedBot (his wallet lists are built for their import)",
         fnf="His FnF circle and Emperor Journal members: 5 onboarded first", in_person="No", airdrop="Partly (WL and mint posts)",
         segment=TREN, pool=19185, ptraders=5, share=0.3, members=16500, per_member="nft",
         basis="Measured (research): median 15,885 views a post + 20% of ~16.5K Telegram (third-party snapshot); Discord size not public",
         deliverables="Days 1-3: his FnF circle (5 traders) onboarded 1:1; then 6 alpha threads executed on OKX Wallet and 1 X Spaces with the referral link",
         need="Better execution for the calls his members already copy",
         support="OKX DEX API docs for his Discord tools; Meme Mode and Track guide; wallet-level referral dashboard",
         check="Emperor Journal Discord size and activity; wallets or PnL for his FnF circle; share of members trading memecoins rather than farming",
         line="Proof: members post results from his tools (GMGN +$802, +$387); 1.9% engagement, best of the big accounts · FnF: his circle and Emperor Journal · Flow: GMGN and BasedBot wallet lists → OKX",
         scores=(3, 4, 5, 4, 4, 4, 3), go_live=10, event="No"),
    dict(short="CVC Ventures network",
         name="CVC Ventures network: 0xRusdi (+ Tomket Lovers), Ponyin (+ ApexFNF), Setya",
         link="https://x.com/CVC_Ventures ; https://x.com/0xRusdi ; https://x.com/Ponyin ; https://x.com/ApexFNF_io ; https://x.com/setyamickala ; https://x.com/TomketLovers",
         type="Trader network: three trencher KOLs and two FnF clans",
         audience="Indonesian Solana and multi-chain memecoin traders on BasedBot, Trojan and fomo",
         evidence=("0xRusdi: $8,463 of BasedBot referral earnings at the 25% tier (dashboard, 26 Apr 2026); Ponyin: Trojan PnL card, "
                   "March 2026, ~$170K of volume (1.13K SOL bought, 1.24K sold); Tomket Lovers: fomo clan of 13-15 active traders; "
                   "Setya: X 52,661, CMIF 2025; CVC_Ventures team: Setya, Ponyin, 0xRusdi (X, Nov 2026)"),
         assumptions=("Who signs for the network; ApexFNF and Tomket clan sizes; how far the four audiences overlap (reach adds them "
                      "up, so it is an upper bound); Setya's role beyond Airdrop Finder"),
         route="X DM to 0xRusdi first (strongest proof), then a group call with the CVC_Ventures team",
         proposal=("One network deal: a code per member on one shared dashboard; 15 FnF traders onboarded 1:1 in days 1-3 (5 from "
                   "0xRusdi's referral book, 5 from ApexFNF, 5 from the Tomket clan); a network trading competition; fee share split "
                   "inside the network"),
         risk=("Four parties to align; Tomket Lovers is on Bitget Wallet Indonesia's Superfans list; Ponyin's Trojan ties and Dawn "
               "role; Setya co-runs Airdrop Finder (airdrop audience)"),
         reason="The strongest proof and the most FnF traders on the list, but four parties make it the slowest to sign",
         vampire="BasedBot (0xRusdi), Trojan (Ponyin), fomo (Tomket clan); Bitget Wallet (Tomket)",
         fnf="15: 0xRusdi's 5 most active referred traders, 5 from ApexFNF (Ponyin's FnF), 5 from the Tomket Lovers clan",
         in_person="No", airdrop="Partly (Setya's Airdrop Finder; Tomket campaigns)", segment=TREN, pool=25120, ptraders=15, share=0.3,
         members=None, per_member="trade",
         basis=("Mixed: Ponyin measured (4,628 median views); 0xRusdi and Setya at 20% of followers (25.8K, 52.7K); Tomket Lovers "
                "4,800 (its promo views). Audiences overlap, so this is an upper bound"),
         deliverables="Days 1-3: 15 FnF traders onboarded 1:1 (5 each from 0xRusdi's book, ApexFNF and the Tomket clan); then each member posts with his code and the network runs a 7-day trading competition",
         need="One deal that pays every member and clan fairly, with terms at least as good as BasedBot's 25% tier",
         support="One shared dashboard with per-member codes; multi-chain Meme Mode guide; network leaderboard and competition (needs OKX approval)",
         check="", line="", scores=(5, 4, 5, 5, 2, 5, 1), go_live=17, event="No"),
    dict(short="Brandon Salim + a79", name="Brandon Salim (@brandonsalim) and a79, his trader group",
         link="https://x.com/brandonsalim ; https://x.com/a79",
         type="Celebrity trader with his own application-only trader group",
         audience="Mass Indonesian audience (actor fanbase) plus stock and crypto retail traders; Jakarta",
         evidence=("X 186,757, median 14,342 views over his 25 latest original posts; fomo 24h PnL +$32,185 (pinned, 4 Sep 2026), "
                   "fomo leaderboard #440; posts his fomo referral code on almost every trade since September; founder of a79 "
                   "(a79.world, application-only; X 3,543); Degen Night speaker (31 Oct) with Kevin Susanto and Bevan"),
         assumptions="a79 member count (not public); his fomo referral results; how many fans trade onchain; he says he trades 'peanuts money'",
         route="Warm intro through Kevin Susanto; he already follows my X account; or his management via brandonsalim.com",
         proposal=("OKX DEX referral fee share with a follower discount, in the same 'trade with me, X% off fees' format he runs on "
                   "fomo; 5 a79 traders onboarded 1:1 in days 1-3; an a79 group code with clan rewards; a co-hosted 'actor vs "
                   "trenchers' trading competition"),
         risk="Celebrity reach may not trade onchain; his fomo tie; needs POJK 6/2026 finfluencer certification before paid posts",
         reason="Second-largest measured reach, a habit of posting a trading referral code, and his own trader group",
         vampire="fomo (code on almost every trade post)",
         fnf="a79, his application-only trader group: 5 traders onboarded first",
         in_person="Yes", airdrop="No", segment=TREN, pool=14342, ptraders=5, share=0.3, members=None, per_member="trade",
         basis="Measured (research): median 14,342 views over his 25 latest original posts; a79 size not public",
         deliverables="Days 1-3: 5 a79 traders onboarded 1:1; then trade posts with the OKX link in his fomo format and an a79 trading competition",
         need="The same trade-with-me format he runs on fomo, multichain and non-exclusive",
         support="Multichain Meme Mode guide (Solana, BSC, Base); a79 group code and leaderboard; wallet-level referral dashboard; compliance check of posts",
         check="", line="", scores=(3, 4, 3, 5, 4, 5, 2), go_live=17, event="No"),
    dict(short="Crypto Ndo", name=SL.cell(9, 1).value, link=SL.cell(9, 2).value, type=SL.cell(9, 3).value,
         audience=SL.cell(9, 5).value, evidence=SL.cell(9, 6).value, assumptions=SL.cell(9, 11).value,
         route=SL.cell(9, 7).value, proposal=SL.cell(9, 8).value, risk=SL.cell(9, 9).value,
         reason=SL.cell(9, 10).value, vampire=SL.cell(9, 12).value,
         fnf="Traders in his circle: 5 onboarded first", in_person="No", airdrop="No",
         segment=TREN, pool=11860, ptraders=5, share=0.3, members=None, per_member="trade",
         basis="Estimated: 20% of 59.3K X followers (median views not yet measured)",
         deliverables="Days 1-3: 5 traders from his circle; then 8 Meme Mode trade threads and 2 X Spaces",
         need="A reason to move his followers' referral flow from fomo to OKX (he already trades on OKX Wallet)",
         support="Co-branded Meme Mode preset guide in Bahasa; weekly wallet-level dashboard; Bahasa support chat; compliance check of posts",
         check="", line="", scores=(3, 3, 4, 4, 3, 5, 4), go_live=17, event="No"),
    dict(short="DegenRex", name="DegenRex (@degen2ex)", link="https://x.com/degen2ex ; https://t.me/degenrex",
         type="Champion trader with a Telegram call channel",
         audience="Indonesian memecoin trenchers on BSC, Base and Solana; Jakarta event crowd",
         evidence=("Won Clash of Trenchers by Ave.ai at Web3 Week Asia, Jakarta (11-12 Nov 2026): group winner on day 1, then the "
                   "final, 1 SOL to 10 SOL in 20 minutes, live on stage (Ave.ai, Ave Indonesia and W3W posts); Telegram degenrex "
                   "1,845; median 1,328 views a post"),
         assumptions="His prize share and trading volume in the final (not disclosed); how many Telegram members trade his calls",
         route="X DM; Jakarta events; through Ave Indonesia's organisers",
         proposal=("Trader ambassador: OKX DEX fee share; 3 traders from his call channel onboarded 1:1; live 'OKX DEX champion' "
                   "sessions in his Telegram; defend his title in an OKX DEX trading competition"),
         risk="Small reach; Ave.ai and four.meme already back him",
         reason="The strongest live, verified trading proof on the list, and multichain (BSC, Base, Solana) like OKX DEX",
         vampire="Ave.ai and four.meme (BSC)",
         fnf="3 traders from his Telegram call channel, onboarded first", in_person="Yes", airdrop="No",
         segment=TREN, pool=1697, ptraders=3, share=0.3, members=1845, per_member="trade",
         basis="Measured (research): median 1,328 views a post + 20% of 1,845 Telegram subscribers (the research's 1.5x boost for the win is left out)",
         deliverables="Days 1-3: 3 traders from his call channel onboarded 1:1; then 2 live trading sessions in his Telegram and a title-defence trading competition",
         need="A recurring stage as a champion, and routing for his BSC and Base plays",
         support="Competition prize pool (needs OKX approval); multichain Meme Mode guide; wallet-level referral dashboard",
         check="", line="", scores=(5, 2, 3, 3, 4, 5, 4), go_live=24, event="No"),
    dict(short="Meteora LP Army", name="Meteora Indonesia 'LP Army' (hosts: Evil Panda, Friday7th)",
         link=SL.cell(13, 2).value, type=SL.cell(13, 3).value, audience=SL.cell(13, 5).value,
         evidence=SL.cell(13, 6).value + "; co-host Friday7th posts a DLMM bot journal (+1.946 SOL over 155 trades, 1-17 Nov)",
         assumptions=SL.cell(13, 11).value, route=SL.cell(13, 7).value, proposal=SL.cell(13, 8).value,
         risk=SL.cell(13, 9).value, reason=SL.cell(13, 10).value, vampire=SL.cell(13, 12).value,
         fnf="5 LP whales from the community", in_person="Yes", airdrop="No",
         segment=LP, pool=2982, ptraders=5, share=0.25, members=462, per_member="trade",
         basis="Estimated: 20% of MeteoraIDN's 13,411 X followers + 300 at one co-hosted night (past nights: 462 and 671)",
         deliverables="Day 1: co-hosted community night with wallets funded on site; 5 LP whales onboarded in days 1-3",
         need="Event costs and good speakers",
         support="Co-hosted community nights (venue and food budget, needs OKX approval); OKX speaker on swaps and Bridge; on-site wallet set-up desk",
         check="", line="", scores=(3, 1, 4, 3, 5, 3, 4), go_live=24, event="Yes"),
]
assert len(partners) == 8
partners.insert(2, partners.pop(3))  # Ponyin + CVC network to #3, GuarEmperor to #4
_net, _ge = partners[2], partners[3]
_net.update(
    short="Ponyin + CVC network",
    name="Ponyin (@Ponyin) + CVC Ventures network: ApexFNF, 0xRusdi (+ Tomket Lovers), Setya",
    route="X DM to Ponyin; he opens the CVC team (0xRusdi, Setya) and ApexFNF, his FnF",
    reason="Ponyin's verified trading opens a whole circle: ApexFNF, 0xRusdi's referral book with the Tomket clan, and Setya's airdrop traders",
    check="Ponyin's Trojan dashboard or wallet history; ApexFNF member wallets and activity; 0xRusdi's BasedBot referral export; Tomket clan size; who signs for the network",
    line="Proof: Ponyin's Trojan PnL card, ~$170K in March 2026; 0xRusdi's $8,463 BasedBot referral pay · FnF: ApexFNF and the Tomket clan · Network: one DM to Ponyin opens the CVC team (0xRusdi, Setya) · Flow: Trojan, BasedBot, fomo → OKX",
    scores=(5, 4, 5, 5, 3, 5, 1), go_live=10)
_ge.update(go_live=17)
for _p in partners:
    if _p["short"] not in ("Kevin Susanto", "Bevan / Narrative Labs"):
        _p["in_person"] = "No"
partners[7]["route"] = "X DM to the hosts (Evil Panda, Friday7th)"
WEIGHTS = [0.15, 0.25, 0.2, 0.1, 0.15, 0.05, 0.1]


def ramp(p, L, s=0.02, a=0.15, v=2000, vf=16400, start=4, days=30, tr=25, tf=120):
    """Task 4 ramp for a posted partner (base case), L days after go-live: FTDs, trades a day, swaps."""
    ftd_total = p["pool"] * s * a
    ftd = ftd_total * min(L, days) / days
    fnf_tpd = p["ptraders"] * vf / days / tf if L >= start else 0
    tpd = ftd * v / days / tr + fnf_tpd
    swaps = ftd_total * v / days * L * L / (2 * days) + p["ptraders"] * vf / days * max(0, L - start + 1)
    return ftd, tpd, swaps, fnf_tpd

N0, N1, TOT = 37, 44, 45  # section C rows


def build(research=False):
    wb = Workbook()
    t1 = wb.active
    t1.title = "Task 1"
    t2, t3, t4, mi = (wb.create_sheet(n) for n in ("Task 2", "Task 3", "Task 4", "Miscellaneous"))
    M, T1, T2 = "Miscellaneous!", "'Task 1'!", "'Task 2'!"

    # ============================================================ Miscellaneous (inputs first: other tabs read them)
    put(mi, "A1", "Miscellaneous: assumptions, products and fees, benchmarks, growth path, brief checklist, sources and longlist", F_TITLE, wrap=False)
    put(mi, "A3", "A. Assumptions (yellow cells are inputs; every tab recalculates). Swaps per partner = reach P x s x a x v, plus FnF power traders", F_SECTION, wrap=False)
    header(mi, 4, ["Assumption", "Low", "Base", "High", "Basis"])
    A = [  # row 5 onwards
        ("Sign-up rate s: share of the reach pool P that signs up through the partner's link in 30 days", (0.01, 0.02, 0.04), PCT1,
         "Research sheet (20 Nov 2026). P = median views on recent original posts + Discord members online, or + 20% of Telegram subscribers"),
        ("Activation a: share of sign-ups who trade (an FTD)", (0.061, 0.15, 0.3), PCT1,
         "Low: 7 of 115 Trojan referrals active (@0xSarimin, Mar 2026). Base: midpoint. High: campaign-boosted"),
        ("Swaps per FTD in its first month (v)", (286, 2000, 16400), MONEY,
         "Low: @0xSarimin's referred volume under $2K across 7 active traders. Base: midpoint, consistent with @0xRusdi's ~$3.4M referred on BasedBot. High: @AiSlimx, $547 a day on OKX Wallet Boost"),
        ("Swaps per FnF power trader in their first month", (2000, 16400, 50000), MONEY,
         "Low: trades like an average FTD. Base: @AiSlimx, $547 a day on OKX Wallet Boost (May 2026), the only measured OKX Wallet trader. High: between that and @Ponyin's ~$170K a month on Trojan"),
        ("FnF power traders start trading on day", (8, 4, 2), None, "Base: onboarded 1:1 in the partner's first 3 days"),
        ("Days in the partner's first month", (30, 30, 30), None, "Each partner's first 30 days live (the same basis as Task 2's first month)"),
        ("Reach of a typical post, accounts not yet measured (share of followers)", (0.2, 0.2, 0.2), PCT,
         "Median views divided by followers across the 11 accounts the research measured: 10% to 44%, median 20%"),
        ("Community benchmark: trading community, swaps a member a month", (300, 300, 300), MONEY,
         "15% of members active x $2,000 a month (research). A ceiling for context, never added to the forecast"),
        ("Community benchmark: NFT or airdrop community, swaps a member a month", (14, 14, 14), MONEY,
         "5% active x $286 a month (research)"),
        ("Net revenue rate", (0.0004, 0.0004, 0.0004), PCT2,
         "Brief (OKX DEX actually averaged ~0.09% in Oct 2026, DefiLlama)"),
    ]
    for i, (label, vals, fmt, basis) in enumerate(A):
        r = 5 + i
        cell(mi, r, 1, label)
        for j, v in enumerate(vals):
            cell(mi, r, 2 + j, v, fmt=fmt, inp=True)
        cell(mi, r, 5, basis)
    # row map
    S_, A_, V_, VF_, ST_, D_, REACH_, BT_, BN_, RT_ = 5, 6, 7, 8, 9, 10, 11, 12, 13, 14
    s = {k: f"{M}${c}${S_}" for k, c in zip("lbh", "BCD")}
    a = {k: f"{M}${c}${A_}" for k, c in zip("lbh", "BCD")}
    v = {k: f"{M}${c}${V_}" for k, c in zip("lbh", "BCD")}
    vf = {k: f"{M}${c}${VF_}" for k, c in zip("lbh", "BCD")}
    st = {k: f"{M}${c}${ST_}" for k, c in zip("lbh", "BCD")}
    DAYS, RATE = f"{M}$C${D_}", f"{M}$C${RT_}"
    RATE_L = f"$C${RT_}"  # same-sheet form

    put(mi, "A17", "B. Trade size by segment (turns swaps into trades a day)", F_SECTION, wrap=False)
    header(mi, 18, ["Segment", "Average FTD trade", "Average FnF trade", "Trades a day per FTD (base)", "Trades a day per FnF trader (base)", "Basis"])
    for i, (name, rs, ps, basis) in enumerate([(TREN, 25, 120, "~0.35 SOL buys at ~$72 SOL; FnF traders size up"),
                                               (LP, 50, 150, "Assumption: fewer, larger swaps when LPs rebalance")]):
        r = 19 + i
        cell(mi, r, 1, name)
        cell(mi, r, 2, rs, fmt=MONEY, inp=True)
        cell(mi, r, 3, ps, fmt=MONEY, inp=True)
        cell(mi, r, 4, f"=$C${V_}/$C${D_}/B{r}", fmt=NUM1)
        cell(mi, r, 5, f"=$C${VF_}/$C${D_}/C{r}", fmt=NUM1)
        cell(mi, r, 6, basis)
    SEG = f"{M}$A$19:$A$20"

    # C. products and fees
    put(mi, "A26", "C. OKX products and fees (published fee vs the brief's 0.04%)", F_SECTION, wrap=False)
    PR = L["Products & revenue"]
    hdr = [PR.cell(2, c).value for c in range(1, 10)]
    hdr[4] = "Swaps per active trader a month (research base for swaps; others assumed)"
    header(mi, 27, hdr)
    for i, r in enumerate(range(3, 13)):
        row = 28 + i
        for c in range(1, 10):
            vv = PR.cell(r, c).value
            if c == 5 and r in (3, 6):
                vv = f"=$C${V_}"
            elif c == 6:
                vv = f"=D{row}*E{row}"
            elif c == 7:
                vv = f"=E{row}*{RATE_L}"
            elif c == 8 and r == 9:
                vv = "Multi-chain traders moving funds between chains"
            elif c == 8 and r == 11:
                vv = "Not a target: fees go to a rival venue"
            fmt = MONEY2 if c in (6, 7) else (PCT2 if c == 4 or (c == 3 and isinstance(vv, (int, float))) else (MONEY if c == 5 else None))
            cell(mi, row, c, vv, fmt=fmt, inp=c in (4, 5) and vv not in (None, "") and not str(vv).startswith("="))
    MEME_FEE = f"{M}$D$28"
    header(mi, 39, ["OKX DEX actuals (DefiLlama, mid-Nov 2026)", "Value", "Note"])
    for i, r in enumerate(range(15, 18)):
        row = 40 + i
        cell(mi, row, 1, PR.cell(r, 1).value)
        cell(mi, row, 2, "=B41/B40" if r == 17 else PR.cell(r, 2).value, fmt=PCT2 if r == 17 else MONEY)
        cell(mi, row, 3, PR.cell(r, 3).value)

    # D. rival benchmark
    OF = L["Offers & support"]
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
    put(mi, "A53", "E. How the number grows (0.04%, before partner shares; each partner's first month, base case)", F_SECTION, wrap=False)
    header(mi, 54, ["Step", "Swaps", "Revenue (0.04%)", "Why it's realistic"])
    fnf_month = f"$C${VF_}*($C${D_}-$C${ST_}+1)/$C${D_}"
    steps = [
        ("1. Audience only: FTDs x $2,000 (no FnF power traders)", f"={T1}J{TOT}", "The research sheet's base case for the same 8"),
        ("2. Plus FnF power traders onboarded by day 4: our plan", f"={T1}M{TOT}", "FnF power traders trade from day 4, before any post"),
        ("3. Double the FnF power traders at the top 3", f"=B56+({T1}H37+{T1}H38+{T1}H39)*{fnf_month}",
         "Kevin's Jakarta circle, Bevan's Masterclass (29 paid seats), GuarEmperor's Emperor Journal"),
        ("4. Plus pilot B pushed to 3x with campaigns", f"=B57+3*{T2}$C$6", "B's growth deal: share brackets, milestone bonus and a funded trading competition"),
        ("5. $10K a month", f"=10000/{RATE_L}", "B at 3x, C at 2x and A at forecast (about $22M) plus our 8 with more FnF traders: a month 4-6 goal"),
    ]
    for i, (aa, b, d) in enumerate(steps):
        row = 55 + i
        cell(mi, row, 1, aa)
        cell(mi, row, 2, b, fmt=MONEY)
        cell(mi, row, 3, f"=B{row}*{RATE_L}", fmt=MONEY)
        cell(mi, row, 4, d)

    # F. brief checklist
    put(mi, "A61", "F. Brief checklist: where each request is answered", F_SECTION, wrap=False)
    checklist = [
        ("Format: up to 6 slides or 4 pages, plus a prospect spreadsheet", "Doc (4 pages) + 6 slides + this workbook"),
        ("Public information only; no contact, spend or private contacts", "All tabs; doc, method note"),
        ("Task 1: 8 potential partners", "Task 1, section A; doc, Task 1; slide 2"),
        ("Task 1: at least 6 whose main proposition is not airdrops", "Task 1, brief checks (row 21)"),
        ("Task 1: at least 3 partner types", "Task 1, column D (6 types)"),
        ("Task 1: at least 2 you could meet in person", "Task 1, column O (Kevin Susanto and Bevan, Jakarta)"),
        ("Task 1: link, audience and market, evidence, route, proposed partnership, main risk, priority with reason", "Task 1, section A"),
        ("Task 1: top 3 and why you'd approach them first", "Task 1, rows 4-8; doc, Task 1; slide 2"),
        ("Task 1: verified facts vs assumptions; follower count alone is not evidence", "Task 1, columns H and I, and section C (measured vs estimated reach)"),
        ("Task 2: first-month contribution for each partner, with the calculation", "Task 2, rows 23-35 and 40; doc, Task 2; slide 3"),
        ("Task 2: preferred pilot and reasoning", "Task 2, row 41"),
        ("Task 2: counter-offer or pilot structure", "Task 2, rows 6-20, 42 and 43"),
        ("Task 2: evidence before signing", "Task 2, row 44"),
        ("Task 2: chosen partner's economics at 50% of forecast", "Task 2, row 28 and row 45"),
        ("Task 2: exclude other costs, then say which you'd investigate", "Task 2, rows 46-47"),
        ("Task 2: pilot only one partner at first", "Task 2, row 41 (B first; C once B passes 2x; A later)"),
        ("Task 3: use your highest-priority real prospect", "Task 3 (Kevin Susanto); doc, Task 3; slide 4"),
        ("Task 3: first outreach message of 120 words or fewer", "Task 3, row 4 (word count in the label)"),
        ("Task 3: follow-up after no response", "Task 3, row 6"),
        ("Task 3: reply to 'another platform offers a higher commission'", "Task 3, row 7"),
        ("Task 3: activation plan for the first 14 days after agreement", "Task 3, rows 13-20"),
        ("Task 3: 3 most important first-meeting questions", "Task 3, rows 8-10"),
        ("Task 3: no unverified product, incentive or commercial promises", "Task 3, row 11"),
        ("Task 4: weekly targets for outreach, meetings, pilots and activated partners", "Task 4, section A; doc, Task 4; slide 5"),
        ("Task 4: one in-person prospecting day (where, whom, access, fallback)", "Task 4, section E; doc, Task 4"),
        ("Task 4: response to 20 messages, 2 replies, 0 meetings", "Task 4, section D (outreach reply rate); doc, Task 4; slide 5"),
        ("Task 4: results that mean continue, renegotiate or stop", "Task 4, section E (day 30); doc, Task 4"),
        ("Task 4: activated-partner definition", "Task 4, row 23 and section D; doc, Task 4; slide 5"),
        ("Task 4: commitment statement", "Doc, commitment; slide 6"),
        ("Cite the public sources you use", "Miscellaneous, section G; doc, sources"),
        ("Explain how you used AI and what you verified yourself", "Doc, method note; slide 6"),
        ("Be ready to adapt when assumptions change", "Miscellaneous, section A and Task 2, rows 6-9 (yellow inputs)"),
        ("Extra: the same deal terms for our real top 3", "Task 2, rows 50-55"),
        ("Extra: 100 prospects screened, with the reason each was cut", "Miscellaneous, section H"),
    ]
    header(mi, 62, ["#", "Brief request", "Where it's answered"] + (["Status"] if research else []))
    for i, (req, where) in enumerate(checklist):
        row = 63 + i
        cell(mi, row, 1, i + 1)
        cell(mi, row, 2, req)
        cell(mi, row, 3, where)
        if research:
            cell(mi, row, 4, "Done")
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
        ("Bevan's published wallet (Solscan)", "https://solscan.io/account/8DC3hpFtU3mEU3FUCtDfU72SmxTbrw4w35fyMYP8qGxg"),
        ("Narrative Labs (Masterclass, Discord, meetups)", "https://narrativelabs.id"),
        ("Bevan's weekly live sessions (X)", "https://x.com/bevanpramudt/status/2115727803245347257"),
        ("Kevin Susanto fomo PnL, 19-20 Nov 2026 (X)", "https://x.com/KevinSusanto/status/2123472525070102987"),
        ("Kevin Susanto fomo trades (tracker)", "https://fomoradar.io/traders/kevinsusanto"),
        ("Ponyin Trojan PnL card (X)", "https://x.com/Ponyin/status/2037746268039209077"),
        ("GuarEmperor member results (X)", "https://x.com/GuarEmperor/status/2115414023114256560"),
        ("GuarEmperor wallet lists (BasedBot, GMGN)", "https://github.com/ariefzzz5421/wallet-tracker-transfer"),
        ("Bevan 'The Kingslayer', 1 SOL to 10 SOL (Ave.ai)", "https://x.com/Jodski888/status/2121255854792188378"),
        ("Brandon Salim pinned fomo PnL (X)", "https://x.com/brandonsalim/status/2096005957692301790"),
        ("a79, Brandon Salim's trader group", "https://x.com/a79"),
        ("Degen Night speakers, 31 Oct 2026", "https://x.com/0xRememberUs/status/2116500473356411230"),
        ("DegenRex wins Clash of Trenchers (Ave.ai)", "https://x.com/AveaiGlobal/status/2120891736373518591"),
        ("Web3 Week Asia recap (W3W)", "https://x.com/W3WOfficial/status/2121498996955132404"),
        ("CVC Ventures team (X)", "https://x.com/CVC_Ventures"),
        ("ApexFNF, Ponyin's FnF (X)", "https://x.com/ApexFNF_io"),
        ("Tomket Lovers (X)", "https://x.com/TomketLovers"),
        ("Meteora LP Army Jakarta (462 went)", "https://luma.com/76fvdt2j"),
        ("MET Bali (671 went)", "https://luma.com/pkr27xr4"),
        ("W3AI Wednesday at JURA (weekly)", "https://luma.com/zkq937gq"),
        ("ICN: 10 X crypto influencers in Indonesia", "https://indonesiacrypto.network/blog/10-x-twitter-crypto-influencers-growing-the-indonesian-cryptosphere"),
        ("ICN community sizes (Mar 2026)", "https://indonesiacrypto.network/blog/crypto-community-indonesia"),
        ("POJK 6/2026 finfluencer certification", "https://economy.okezone.com/read/2026/07/08/320/3228975/ojk-wajibkan-sertifikasi-influencer-kripto-ini-tujuannya"),
        ("Law 4/2026 (OJK powers over offshore platforms)", "https://www.hsfkramer.com/ru/notes/indonesia/2026-posts/indonesia-reshapes-digital-asset-and-cryptocurrency-framework-with-2026-financial-services-omnibus-law"),
        ("Akademi Crypto inquiry halted (Oct 2026)", "https://justisio.com/berita/2026-10-26-polda-metro-hentikan-penyelidikan-laporan-dugaan-penipuan-kripto-timot"),
        ("OKX Alpha Traders Summit, Bali", "https://bitpinas.com/feature/okx-alpha-traders-summit/"),
        ("Rival referral rates", "pump.fun: PrimeXBT, Crypto Briefing; Axiom: affiliate guides (unconfirmed); fomo: Datawallet"),
        ("Indonesian X evidence", "@KevinSusanto fomo PnL; @bevanpramudt wallet; @GuarEmperor member results; @0xRusdi BasedBot dashboard; @Ponyin Trojan results; @brandonsalim fomo PnL; @degen2ex win; @cryptondo OKX Wallet use; @0xSarimin Trojan referrals; @AiSlimx OKX Wallet Boost"),
        ("Reach, FTD and volume rates", "Indonesia Crypto KOL & Partner Research, Volume proof and Shortlist deep-dive tabs (X API and Discord invite counts, 20 Nov 2026)"),
    ]
    header(mi, g0 + 1, ["Topic", "Source"])
    for i, (topic, url) in enumerate(sources):
        row = g0 + 2 + i
        cell(mi, row, 1, topic)
        c = cell(mi, row, 2, url)
        if url.startswith("http"):
            link(c, url)
    h0 = g0 + 2 + len(sources) + 1

    # H. longlist
    put(mi, f"A{h0}", "H. Longlist: 100 public Indonesian prospects screened", F_SECTION, wrap=False)
    put(mi, f"A{h0 + 1}", "Each prospect checked against the same filter: proof of trading, FnF power traders it can bring, rival flow to move, trencher fees.", wrap=False)
    LL = L["Longlist (100 screened)"]
    drop = {"Bitcoinist Indonesia", "Blockchainmedia.id", "CoinDesk Indonesia", "CoinSpeaker Indonesia", "Cryptonews Indonesia",
            "Investortrust.id", "Blockplay", "Crypto Indonesia (Discord)", "CryptoHack (Indonesia)", "Edwin Nazir / NQvers",
            "Indonesia Paham Bitcoin", "Komunitas BlockBali (Egy)", "Neko Group", "Xenorize", "BEM FEB UI (Biro Project)",
            "BINUS Blockchain and Crypto Club (BBCC)", "YESCON (IDBW @ UI)", "Bali Blockchain Weeks Campus Tour",
            "Bulan Literasi Kripto (OJK + ABI)", "Solana Village Bali"}
    rows = []
    for r in range(15, 115):
        rec = [LL.cell(r, c).value for c in range(1, 11)]
        name = rec[1]
        if name in drop:
            continue
        if name == "vemi.pramudya":
            rec = [None, next(q["name"] for q in partners if q["short"].startswith("GuarEmperor")), "Trading creator", "https://x.com/GuarEmperor", TREN, "Y", "In the 8",
                   "#4: FnF circle; members post results from his tools; 1.9% engagement", "Wave 2 (day 17)",
                   "https://x.com/GuarEmperor/status/2115414023114256560"]
        elif name == "Airdrop Finder (Telegram)":
            rec[7] = "Airdrop-led: fails the brief's main filter. Its co-founder Setya joins through the CVC Ventures network (#4), not this channel"
        elif name == "Hyperliquid Indonesia (community)":
            rec[6], rec[7], rec[8] = "Excluded by rule", "Rival perp venue: OKX can't send traffic there", ""
        elif name == "Superteam Indonesia":
            rec[6], rec[7], rec[8] = "Bench (month 2)", "Meetable weekly in Jakarta, but builder-led: no trading proof yet", "Month 2: co-hosted trader night once a partner proves demand"
        rows.append(rec)
    new = [
        ("Kevin Susanto", "Trading creator", "https://x.com/KevinSusanto", TREN, "Y", "In the 8",
         "#1: most FTDs, a verified top fomo trader, and the fastest yes through OKX's regional team", "Wave 1 (day 10)", "https://x.com/KevinSusanto/status/2123472525070102987"),
        ("Bevan Pramudito / Narrative Labs", "Trading educator", "https://x.com/bevanpramudt", TREN, "Y", "In the 8",
         "#2: verified wallet and a live 1-to-10 SOL run, a paid community he founded, GMGN flow, meetable in Jakarta", "Wave 1 (day 10)", "https://narrativelabs.id"),
        ("CVC Ventures network (0xRusdi + Tomket Lovers, Ponyin + ApexFNF, Setya)", "Trading community", "https://x.com/CVC_Ventures", TREN, "Y", "In the 8",
         "#3, through Ponyin: Trojan PnL, BasedBot referral book, ApexFNF and the Tomket clan (15 FnF traders)", "Wave 1 (day 10)", "https://x.com/Ponyin"),
        ("Brandon Salim + a79", "Trading creator", "https://x.com/brandonsalim", TREN, "Y", "In the 8",
         "#5: 14K median views, fomo code on almost every trade, founder of the a79 trader group", "Wave 2 (day 17), intro through Kevin", "https://x.com/brandonsalim/status/2096005957692301790"),
        ("Crypto Ndo", "Trading creator", "https://x.com/cryptondo", TREN, "Y", "In the 8",
         "#6: largest trencher reach on X; trades on OKX Wallet himself", "Wave 2 (day 17)", "https://x.com/cryptondo"),
        ("DegenRex", "Trading creator", "https://x.com/degen2ex", TREN, "Y", "In the 8",
         "#7: won Clash of Trenchers live (1 SOL to 10 SOL); multichain", "Wave 3 (day 24)", "https://x.com/AveaiGlobal/status/2120891736373518591"),
        ("Meteora LP Army", "Onchain trading community", "https://x.com/EvilPanda", LP, "Y", "In the 8",
         "#8: largest in-person onchain crowds (462 and 671)", "Wave 3 (day 24): co-hosted night", "https://luma.com/76fvdt2j"),
        ("Stress Capital (Orpheuzkaze)", "Trading community", "https://x.com/StressCapitals", TREN, "Y", "Bench (month 2)",
         "30.7K-member clan Discord, 2,453 online, but no trading proof yet", "Month 2: clan competition if clan leads show they trade", "https://discord.gg/stresscapitals"),
        ("TwipsX", "Trading creator", "https://x.com/TwipsX", TREN, "Y", "Bench (month 2)",
         "60K-view BasedBot promo, but no proof its readers trade", "Month 2 if its promo quality checks out", "https://x.com/TwipsX"),
        ("Nicholas Jason Santoso", "Finance creator", "https://instagram.com/nicholasjasons", "Beginners", "N", "Bench (month 2)",
         "~15M Instagram views in 30 days (self-reported); no trading proof; airdrop community", "Month 2: CEX-beginner and tokenized-stocks test", "https://nicholasjasons.vercel.app"),
        ("Tequila / Lil Stars", "Ecosystem community", "https://x.com/0xtequilaa", "NFT holders", "N", "Bench (month 2)",
         "128K-member Discord but 1.2% online; NFT and Monad audience", "Month 2: holder-only trading competition", "https://x.com/Lilstarrrs"),
        ("mage (@magersih)", "Trading creator", "https://x.com/magersih", TREN, "Y", "Bench (month 2)",
         "Strong execution content (Telegram 2,271); moved his main wallet from OKX to Bitget Wallet on 11 Nov", "Month 2 win-back: ask what made him leave", "https://x.com/magersih/status/2120438997705224345"),
        ("VIBI (@vibixuee)", "Finance creator", "https://x.com/vibixuee", "Beginners", "N", "Bench (month 2)",
         "Beginner onboarding content, 2.6% engagement; no trading proof", "Month 2: onboarding Reel", "https://linktr.ee/VIBI"),
        ("Friday7th / Hedgin", "Onchain tool", "https://x.com/Friday_SOL", LP, "Y", "Bench (month 2)",
         "Co-hosts Meteora LP Army; DLMM bot journal (+1.946 SOL over 155 trades)", "Through the Meteora night: vault rebalances on the DEX API", "https://x.com/Friday_SOL/status/2122600631651164614"),
        ("All Good (@mudacumasekali)", "Event", "https://x.com/mudacumasekali", TREN, "Y", "Access channel",
         "Caster and connector in Kevin, Bevan and Brandon's circle; Emperor Journal member; no group of his own (~715 median views)", "Intros to GuarEmperor and Brandon; MC for an OKX DEX competition", "https://x.com/Jodski888/status/2121255854792188378"),
        ("Luna (@iamretardio)", "Trading creator", "https://x.com/iamretardio", TREN, "Y", "Screened out",
         "Grows the Cove trading bot (rival tool tie); 6.1K followers", "", "https://x.com/iamretardio/status/2123200836331585876"),
        ("Limitmaxi", "Trading creator", "https://x.com/limitmaxi", TREN, "Y", "Screened out",
         "Clash of Trenchers podium (2nd or 3rd: organiser posts differ); ~440 followers", "Guest trader only", "https://x.com/aveai_id/status/2120950749580677430"),
        ("melon (@melonbuskk)", "Builder", "https://x.com/melonbuskk", TREN, "N", "Screened out",
         "Low activity; access to the EnclaviumFNF group not verified", "", "https://x.com/melonbuskk"),
        ("celinejoo", "Trading creator", "https://x.com/celinejoo__", "Beginners", "N", "Screened out",
         "~900 X followers; newer trader", "", "https://x.com/celinejoo__"),
        ("zeraa (@0xzeraa)", "Ecosystem community", "https://x.com/0xzeraa", "NFT holders", "N", "Screened out",
         "Team role at Lil Stars; reach not measured", "", "https://x.com/0xzeraa"),
    ]
    for rec in new:
        rows.append([None, *rec])
    assert len(rows) == 100, len(rows)
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
        for c, vv in enumerate(rec, 1):
            cc = cell(mi, row, c, vv)
            if c in (4, 10) and str(vv).startswith("http"):
                link(cc, vv)
    widths(mi, {"A": 44, "B": 30, "C": 30, "D": 30, "E": 34, "F": 16, "G": 18, "H": 50, "I": 40, "J": 44})
    mi.freeze_panes = "A3"

    # ============================================================ Task 1
    put(t1, "A1", "Task 1: 8 partners beyond airdrop communities", F_TITLE, wrap=False)
    put(t1, "A2", "Each partner's first month, base case: FTDs = reach P x 2% sign up x 15% trade; swaps = FTDs x $2,000 + FnF power traders x $16,400 from day 4. Revenue at the brief's 0.04%. Yellow cells are inputs.", wrap=False)
    put(t1, "A4", "Why these 3 first", F_BOLD, wrap=False)
    put(t1, "B4", f'="Filter, in order: 1 Proof they already trade → 2 FnF power traders they bring ("&TEXT(B48,"0%")&" of swaps) → 3 Rival flow we move to OKX → 4 FTDs, then trades a day"', F_BOLD, wrap=False)
    for i, p in enumerate(partners[:3]):
        r = N0 + i
        put(t1, f"A{5 + i}", f"{i + 1}. {p['short']}", F_BOLD, wrap=False)
        put(t1, f"B{5 + i}", f'="{p["line"]} · Trencher: "&TEXT(F{r},"0")&" FTDs, "&TEXT(I{r},"0")&" trades a day, $"&TEXT(M{r}/1000,"0")&"K swaps"', wrap=False)
    put(t1, "A8", "Engagement alone isn't enough", F_BOLD, wrap=False)
    put(t1, "B8", "GuarEmperor (#4) has the best engagement of the big accounts, but his circle's trading is not verified yet and his feed mixes WL and mint posts, so he starts in wave 2. Indonesia's biggest crypto names (100 screened, Miscellaneous section H) show no onchain proof.", wrap=False)

    put(t1, "A10", "A. The 8 partners (all fields the brief asks for)", F_SECTION, wrap=False)
    header(t1, 11, ["#", "Partner", "Public link", "Partner type", "Priority", "Priority reason", "Audience and market",
                    "Evidence of fit (verified)", "Assumptions to verify", "Acquisition route", "Proposed partnership",
                    "Main risk", "FnF / power traders they bring", "Rival tool their traders use now", "Meetable in person",
                    "Airdrop-led (main proposition)"])
    for i, p in enumerate(partners):
        r = 12 + i
        vals = [i + 1, p["name"], p["link"], p["type"], f'=IF(J{26 + i}<=3,"High",IF(J{26 + i}<=7,"Medium","Low"))', p["reason"],
                p["audience"], p["evidence"], p["assumptions"], p["route"], p["proposal"], p["risk"], p["fnf"], p["vampire"],
                p["in_person"], p["airdrop"]]
        for c, vv in enumerate(vals, 1):
            cc = cell(t1, r, c, vv)
            if c == 3:
                link(cc, str(vv).split(" ; ")[0])
    cell(t1, 21, 1, "Brief checks", bold=True)
    cell(t1, 21, 2, "Partners")
    cell(t1, 21, 3, "=COUNTA(B12:B19)")
    cell(t1, 21, 4, "Not airdrop-led (6+ needed)")
    cell(t1, 21, 5, '=COUNTIF(P12:P19,"No*")')
    cell(t1, 21, 6, "Meetable in person (2+ needed)")
    cell(t1, 21, 7, '=COUNTIF(O12:O19,"Yes")')
    cell(t1, 21, 8, "Partner types (3+ needed): 6 = a large KOL, a trader-educator with a paid community, alpha callers with tools communities, a trader network with clans, a celebrity trader with his own group, an LP community")

    put(t1, "A23", "B. Ranking (1-5 per criterion; yellow = inputs)", F_SECTION, wrap=False)
    header(t1, 24, ["Partner", "Proof of real trading", "FTD potential", "FnF / power-trader access", "Rival flow to move",
                    "Reachability", "Segment value", "Risk (5 = low)", "Weighted score", "Rank"])
    cell(t1, 25, 1, "Weight", bold=True)
    for c, w in enumerate(WEIGHTS, 2):
        cell(t1, 25, c, w, fmt=PCT, inp=True)
    cell(t1, 25, 9, "=SUM(B25:H25)", fmt=PCT)
    for i, p in enumerate(partners):
        r = 26 + i
        cell(t1, r, 1, f"=B{12 + i}")
        for c, sc in enumerate(p["scores"], 2):
            cell(t1, r, c, sc, inp=True)
        cell(t1, r, 9, f"=SUMPRODUCT(B{r}:H{r},$B$25:$H$25)", fmt="0.00")
        cell(t1, r, 10, f"=RANK(I{r},$I$26:$I$33)")

    put(t1, "A35", "C. Each partner's first month (research method: FTDs = P x s x a; swaps = FTDs x v; plus FnF power traders). Revenue at the brief's 0.04%", F_SECTION, wrap=False)
    header(t1, 36, ["Partner", "Segment", "Reach pool P (people a month)", "Reach measured or estimated", "FTDs (low)", "FTDs (base)",
                    "FTDs (high)", "FnF power traders", "Trades a day (base, once all are trading)", "FTD swaps (base)",
                    "FnF power-trader swaps (base)", "Swaps (low)", "Swaps (base)", "Swaps (high)", "Revenue (0.04%, base)",
                    "Partner share", "Paid to partner", "OKX keeps", "Owned community (members)", "Benchmark swaps a member a month",
                    "Community benchmark a month (context only)", "Reach basis"])
    for i, p in enumerate(partners):
        r = N0 + i
        seg = f"MATCH($B{r},{SEG},0)"
        tpd_r, tpd_f = f"INDEX({M}$D$19:$D$20,{seg})", f"INDEX({M}$E$19:$E$20,{seg})"
        live = lambda k: f"({DAYS}-{st[k]}+1)/{DAYS}"
        cell(t1, r, 1, f"=B{12 + i}")
        cell(t1, r, 2, p["segment"], inp=True)
        cell(t1, r, 3, p["pool"], fmt=NUM0, inp=True)
        cell(t1, r, 4, p["basis"].split(":")[0].split(" (")[0])
        cell(t1, r, 5, f"=C{r}*{s['l']}*{a['l']}", fmt=NUM1)
        cell(t1, r, 6, f"=C{r}*{s['b']}*{a['b']}", fmt=NUM1)
        cell(t1, r, 7, f"=C{r}*{s['h']}*{a['h']}", fmt=NUM1)
        cell(t1, r, 8, p["ptraders"], inp=True)
        cell(t1, r, 9, f"=F{r}*{tpd_r}+H{r}*{tpd_f}", fmt=NUM1)
        cell(t1, r, 10, f"=F{r}*{v['b']}", fmt=MONEY)
        cell(t1, r, 11, f"=H{r}*{vf['b']}*{live('b')}", fmt=MONEY)
        cell(t1, r, 12, f"=E{r}*{v['l']}+H{r}*{vf['l']}*{live('l')}", fmt=MONEY)
        cell(t1, r, 13, f"=J{r}+K{r}", fmt=MONEY)
        cell(t1, r, 14, f"=G{r}*{v['h']}+H{r}*{vf['h']}*{live('h')}", fmt=MONEY)
        cell(t1, r, 15, f"=M{r}*{RATE}", fmt=MONEY2)
        cell(t1, r, 16, p["share"], fmt=PCT, inp=True)
        cell(t1, r, 17, f"=O{r}*P{r}", fmt=MONEY2)
        cell(t1, r, 18, f"=O{r}-Q{r}", fmt=MONEY2)
        cell(t1, r, 19, p["members"], fmt=NUM0, inp=p["members"] is not None)
        bench = f"{M}$C${BT_}" if p["per_member"] == "trade" else f"{M}$C${BN_}"
        cell(t1, r, 20, f"={bench}", fmt=MONEY)
        cell(t1, r, 21, f'=IF(S{r}="","",S{r}*T{r})', fmt=MONEY)
        cell(t1, r, 22, p["basis"])
    cell(t1, TOT, 1, "Total, 8 partners", bold=True)
    for c in list(range(3, 4)) + list(range(5, 16)) + [17, 18]:
        L_ = col(c)
        cell(t1, TOT, c, f"=SUM({L_}{N0}:{L_}{N1})", fmt=t1.cell(N0, c).number_format, bold=True)
    cell(t1, TOT, 21, "Not summed: communities overlap")
    checks = [
        ("FnF power traders' share of swaps (base)", f"=K{TOT}/M{TOT}", PCT),
        ("Swaps if no FnF power traders arrive (FTDs only, base)", f"=J{TOT}", MONEY),
        ("Revenue in that case (0.04%)", f"=J{TOT}*{RATE}", MONEY),
        ("Largest single partner's share of FTDs (base)", f"=MAX(F{N0}:F{N1})/F{TOT}", PCT),
        ("Revenue range, low to high (0.04%)", f'="$"&TEXT(L{TOT}*{RATE},"#,##0")&" to $"&TEXT(N{TOT}*{RATE},"#,##0")&" (base $"&TEXT(O{TOT},"#,##0")&")"', None),
        ("Published-fee upside before partner shares (trenchers 0.5%; LP at the blend in C53)",
         f'=SUMIF(B{N0}:B{N1},"{TREN}",M{N0}:M{N1})*{MEME_FEE}+SUMIF(B{N0}:B{N1},"{LP}",M{N0}:M{N1})*C53', MONEY),
    ]
    put(t1, "A47", "Checks", F_BOLD, wrap=False)
    for i, (aa, f, fmt) in enumerate(checks):
        cell(t1, 48 + i, 1, aa)
        cell(t1, 48 + i, 2, f, fmt=fmt)
    cell(t1, 53, 3, 0.0015, fmt=PCT2, inp=True)
    put(t1, "A55", "Where prospects come from: X (KOL posts, PnL and referral screenshots, public wallet lists; median views via the X API), "
                   "Solscan wallet reads, Discord invite counts, Telegram previews, Luma event pages, rival ambassador and campaign lists "
                   "(Bitget Wallet, BasedBot, Trojan, fomo, pump.fun). Hyperliquid Indonesia and Bael calls were screened out"
                   " as rival-venue and rival-wallet ties.", wrap=False)
    widths(t1, {"A": 30, "B": 30, "C": 30, "D": 26, "E": 12, "F": 34, "G": 30, "H": 46, "I": 34, "J": 26, "K": 44, "L": 34,
                "M": 34, "N": 30, "O": 14, "P": 16, "Q": 14, "R": 14, "S": 14, "T": 14, "U": 18, "V": 48})
    t1.freeze_panes = "C12"

    # ============================================================ Task 2
    put(t2, "A1", "Task 2: pilot B first, then C, then A, on a growth deal instead of a retainer (net revenue rate 0.04%, from the brief)", F_TITLE, wrap=False)
    put(t2, "A2", "Month-1 contribution = revenue at 0.04% - retainer - revenue share - milestone bonus - campaign spend. Brackets are marginal: each rate applies only to the volume inside its bracket. Yellow cells are inputs.", wrap=False)
    put(t2, "A4", "A. Their asks and our growth deal", F_SECTION, wrap=False)
    header(t2, 5, ["Input", "A", "B", "C", "How it works"])
    LT2 = L["Task 2 numbers"]
    vol = [LT2.cell(r, 3).value for r in (4, 5, 6)]
    fee = [LT2.cell(r, 5).value for r in (4, 5, 6)]
    ask = [LT2.cell(r, 6).value for r in (4, 5, 6)]
    rows_in = [
        ("Monthly volume forecast (brief)", vol, MONEY, "Brief; profiles: A airdrop community, B repeat onchain traders, C trading educator (volume unverified)"),
        ("Their ask: fixed fee", fee, MONEY, "Paid once at the start of the pilot (brief)"),
        ("Their ask: revenue share", ask, PCT, "Brief"),
        ("Our offer: retainer at signing", [0, 0, 500], MONEY, "C only, and far below its $3,000 ask"),
        ("Our offer: retainer when forecast is reached", [0, 0, 500], MONEY, "C's second $500 is paid only once its volume reaches forecast"),
        ("Our offer: base revenue share (volume up to forecast)", [0.2, 0.3, 0.15], PCT, "B starts below its 35% ask; C above its 10%; A keeps its 20%"),
        ("Our offer: share step-up per bracket above forecast", [0, 0.05, 0.05], PCT, "Brackets: 100-200%, 200-300% and above 300% of forecast"),
        ("Our offer: milestone cash bonus (share of revenue at the milestone reached)", [0, 0.1, 0.1], PCT, "Paid in USDT within 7 days at 100%, 200% and 300% of forecast; each payment tops up to this share"),
        ("Our offer: launch campaign (OKX-funded)", [0, 250, 300], MONEY, "B: trading competition prize pool; C: co-hosted event"),
        ("Our offer: campaign reinvestment (share of revenue above forecast)", [0, 0.1, 0.1], PCT, "Funds more competitions and events as volume grows"),
        ("Pilot order", [3, 1, 2], None, "B first; C once B passes 2x its forecast; A later, on revenue share only"),
    ]
    for i, (lab, vals, fmt, note) in enumerate(rows_in):
        r = 6 + i
        cell(t2, r, 1, lab, bold=(i == 0))
        for j, vv in enumerate(vals):
            cell(t2, r, 2 + j, vv, fmt=fmt, inp=True)
        cell(t2, r, 5, note)
    for i, (lab, vv) in enumerate((("Milestone 1 (share of forecast)", 1.0), ("Milestone 2", 2.0), ("Milestone 3", 3.0))):
        cell(t2, 18 + i, 1, lab)
        cell(t2, 18 + i, 2, vv, fmt=PCT, inp=True)
    # input rows: 6 volume, 7 fee, 8 ask share, 9 retainer, 10 retainer at forecast, 11 base share, 12 step, 13 bonus,
    #             14 launch, 15 reinvest, 16 order; milestones B18:B20
    PC = {"A": "B", "B": "C", "C": "D"}

    put(t2, "A22", "B. Month-1 contribution at 50%, 100%, 200% and 300% of forecast", F_SECTION, wrap=False)
    header(t2, 23, ["Partner", "Volume vs forecast", "Volume", "Revenue (0.04%)", "Volume up to 100%", "Volume 100-200%",
                    "Volume 200-300%", "Volume above 300%", "Retainer", "Revenue share paid", "Milestone bonus", "Launch campaign",
                    "Campaign reinvestment", "1. Month 1 at our offer", "Partner earns (our offer)", "1. Month 1 at their ask",
                    "Partner earns (their ask)"])
    SC = {}
    r = 24
    for pn in "ABC":
        X = PC[pn]
        for mult in (0.5, 1, 2, 3):
            F = f"${X}$6"
            cell(t2, r, 1, pn, bold=True)
            cell(t2, r, 2, mult, fmt=PCT, inp=True)
            cell(t2, r, 3, f"={F}*B{r}", fmt=MONEY)
            cell(t2, r, 4, f"=C{r}*{RATE}", fmt=MONEY)
            cell(t2, r, 5, f"=MIN(C{r},{F}*$B$18)", fmt=MONEY)
            cell(t2, r, 6, f"=MIN(MAX(C{r}-{F}*$B$18,0),{F}*($B$19-$B$18))", fmt=MONEY)
            cell(t2, r, 7, f"=MIN(MAX(C{r}-{F}*$B$19,0),{F}*($B$20-$B$19))", fmt=MONEY)
            cell(t2, r, 8, f"=MAX(C{r}-{F}*$B$20,0)", fmt=MONEY)
            cell(t2, r, 9, f"=${X}$9+IF(C{r}>={F}*$B$18,${X}$10,0)", fmt=MONEY)
            cell(t2, r, 10, f"={RATE}*(E{r}*${X}$11+F{r}*(${X}$11+${X}$12)+G{r}*(${X}$11+2*${X}$12)+H{r}*(${X}$11+3*${X}$12))", fmt=MONEY)
            cell(t2, r, 11, f"=${X}$13*{RATE}*IF(C{r}>={F}*$B$20,{F}*$B$20,IF(C{r}>={F}*$B$19,{F}*$B$19,IF(C{r}>={F}*$B$18,{F}*$B$18,0)))", fmt=MONEY)
            cell(t2, r, 12, f"=${X}$14", fmt=MONEY)
            cell(t2, r, 13, f"=${X}$15*MAX(D{r}-{F}*{RATE},0)", fmt=MONEY)
            cell(t2, r, 14, f"=D{r}-I{r}-J{r}-K{r}-L{r}-M{r}", fmt=MONEY)
            cell(t2, r, 15, f"=I{r}+J{r}+K{r}", fmt=MONEY)
            cell(t2, r, 16, f"=D{r}-${X}$7-D{r}*${X}$8", fmt=MONEY)
            cell(t2, r, 17, f"=${X}$7+D{r}*${X}$8", fmt=MONEY)
            SC[(pn, mult)] = r
            r += 1
    # r is now 36
    rB1, rB2, rB05, rC05, rC1 = SC[("B", 1)], SC[("B", 2)], SC[("B", 0.5)], SC[("C", 0.5)], SC[("C", 1)]
    cell(t2, 37, 1, "B's campaign pays for itself above (x forecast)", bold=True)
    cell(t2, 37, 2, f"=1+(P{rB1}-N{rB1})/($C$6*{RATE}*(1-$C$11-$C$12-$C$15))", fmt='0.00"x"')
    cell(t2, 37, 3, "Volume at which B's month 1 on our deal matches B's own ask at forecast; below it, the campaign spend has not paid back yet")
    t2.merge_cells(start_row=37, start_column=3, end_row=37, end_column=12)
    cell(t2, 38, 1, "Largest C retainer that stays positive at 50%", bold=True)
    cell(t2, 38, 2, f"=D{rC05}-J{rC05}-K{rC05}-L{rC05}-M{rC05}", fmt=MONEY)
    cell(t2, 38, 3, "So C's retainer is $500 at signing, with the second $500 paid only once C reaches forecast")
    t2.merge_cells(start_row=38, start_column=3, end_row=38, end_column=12)

    answers = [
        ("1. First-month contribution", f'="At their asks: A $"&TEXT(-P{SC[("A", 1)]},"#,##0")&" short ($1,600 - $2,000 - $320), B +$"&TEXT(P{rB1},"#,##0")&" ($800 - $280), C $"&TEXT(-P{rC1},"#,##0")&" short ($2,400 - $3,000 - $240). On our growth deal at forecast: B +$"&TEXT(N{rB1},"#,##0")&" after a $250 campaign, C +$"&TEXT(N{rC1},"#,##0")&" after a $1,000 retainer and a $300 event, A +$"&TEXT(N{SC[("A", 1)]},"#,##0")&" on share only."'),
        ("2. Preferred pilot", "B. It has the clearest evidence (repeat onchain traders), asks for no fixed fee and stays positive at half volume, so it is the one with no downside. I would rather spend on campaigns for B to grow its volume and repeat trading: push B to 2-3x its forecast, then start C to expand once B passes 2x (C could bring volume but has to prove it first), and let A follow later on share only, since an airdrop audience is not our target and tends to arrive with campaigns on its own."),
        ("3. Counter-offer: a growth deal", "No retainer for B, and a small one for C ($500 at signing, $500 more only at forecast) instead of $3,000. Revenue share in brackets that rise as volume grows (B: 30% up to forecast, 35%, 40% and 45% in each bracket above; C: 15%, then 20%, 25%, 30%). A cash bonus at 100%, 200% and 300% of forecast, paid within 7 days and worth 10% of the revenue at that milestone. OKX-funded campaigns: a trading competition for B, a co-hosted event for C, plus 10% of revenue above forecast reinvested in more of them. Every row of section B stays positive."),
        ("Making it more attractive at no fixed cost", "Fast payouts (weekly share, bonus within 7 days); a rate ratchet (a bracket reached is kept as next month's starting rate for one month, if retention holds); a network override (5% of the revenue from partners they introduce, for 3 months); member perks paid from the campaign pool (competitions, fee discounts, early access to OKX Boost and X Launch); OKX speakers and co-branded events (need OKX approval)."),
        ("4. Evidence before signing", "Past numbers: 90 days of referral or bot dashboards (volume, active traders, retention by week). Proof of users: community size and activity (Discord or Telegram analytics) and audience location. Trading history: wallet addresses of the most active members, which we check onchain for unique wallets, trades, repeat days and wash-trade patterns. Past partnership results and a reference. For C, the dashboards behind its $6M estimate; for A, the share of wallets that trade rather than only claim airdrops."),
        ("5. If volume reaches only 50%", f'="B stays positive (+$"&TEXT(N{rB05},"#,##0")&" on our deal after the campaign, +$"&TEXT(P{rB05},"#,##0")&" on its own ask) because there is no fixed fee, and the bracket share falls to 30% automatically. Then find out why: the market (compare with overall memecoin DEX volume), the incentive (did activity drop when the competition ended), or the group itself (active wallets and members falling). The first two are normal swings: adjust the campaign and keep going. If the group has stopped carrying its weight, renegotiate to the base bracket or stop."'),
        ("Campaign cost per community: is it worth it?", f'="Each community gets its own campaign budget, and the rule is simple: B\'s $250 campaign has paid for itself once B passes "&TEXT(B37,"0.0")&"x its forecast. If B is not on track for that by week 3, the campaign stops and B continues on share only."'),
        ("Other costs to investigate", "Campaign budgets per community (prize pools, events), bonus payouts, follower fee discounts (0-20%), gas subsidies, content fees, fraud screening and tracking tools, compliance review (POJK 6/2026, Law 4/2026), translation and BD time."),
    ]
    for i, (aa, b) in enumerate(answers):
        cell(t2, 40 + i, 1, aa, bold=True)
        cell(t2, 40 + i, 2, b)
        t2.merge_cells(start_row=40 + i, start_column=2, end_row=40 + i, end_column=14)
        t2.row_dimensions[40 + i].height = 60

    put(t2, "A50", "C. Our real top 3 on growth-deal terms (each partner's first month, base case, 0.04%)", F_SECTION, wrap=False)
    header(t2, 51, ["Partner", "Swaps, first month (base)", "Revenue (0.04%)", "Base revenue share", "Milestone bonus at forecast",
                    "Month-1 contribution", "At 50% of forecast", "Partner earns", "OKX keeps at the published 0.5% fee", "Evidence before signing"])
    for kk in range(3):
        r, src = 52 + kk, N0 + kk
        cell(t2, r, 1, f"={T1}A{src}")
        cell(t2, r, 2, f"={T1}M{src}", fmt=MONEY)
        cell(t2, r, 3, f"=B{r}*{RATE}", fmt=MONEY2)
        cell(t2, r, 4, f"={T1}P{src}", fmt=PCT)
        cell(t2, r, 5, f"=C{r}*$C$13", fmt=MONEY2)
        cell(t2, r, 6, f"=C{r}-C{r}*D{r}-E{r}", fmt=MONEY2)
        cell(t2, r, 7, f"=C{r}/2*(1-D{r})", fmt=MONEY2)
        cell(t2, r, 8, f"=C{r}*D{r}+E{r}", fmt=MONEY2)
        cell(t2, r, 9, f"=B{r}*{MEME_FEE}*(1-D{r}-$C$13)", fmt=MONEY2)
        cell(t2, r, 10, partners[kk]["check"])
    cell(t2, 55, 1, "Top 3 total", bold=True)
    for cc in "BCEFGHI":
        cell(t2, 55, "ABCDEFGHIJ".index(cc) + 1, f"=SUM({cc}52:{cc}54)", fmt=t2[f"{cc}52"].number_format, bold=True)
    widths(t2, {"A": 34, "B": 16, "C": 16, "D": 16, "E": 16, "F": 16, "G": 16, "H": 16, "I": 14, "J": 16, "K": 14, "L": 14,
                "M": 14, "N": 16, "O": 16, "P": 16, "Q": 16})

    # ============================================================ Task 3
    put(t3, "A1", "Task 3: win and activate Kevin Susanto, our #1 prospect", F_TITLE, wrap=False)
    header(t3, 3, ["Item", "Text"] + (["Word count"] if research else []))
    for i, (aa, b) in enumerate(TASK3):
        r = 4 + i
        cell(t3, r, 1, aa, bold=True)
        cell(t3, r, 2, b)
        if research and r <= 7:
            cell(t3, r, 3, f'=LEN(TRIM(B{r}))-LEN(SUBSTITUTE(TRIM(B{r})," ",""))+1')
    t3["A4"].value = '="First message (X DM), "&(LEN(TRIM(B4))-LEN(SUBSTITUTE(TRIM(B4)," ",""))+1)&" words"'
    put(t3, "A13", "14-day activation plan (day 1 = agreement signed)", F_SECTION, wrap=False)
    header(t3, 14, ["Days", "What happens", "Target"])
    for i, (aa, b, c) in enumerate(PLAN14):
        cell(t3, 15 + i, 1, aa, bold=True)
        cell(t3, 15 + i, 2, b)
        cell(t3, 15 + i, 3, c)
    put(t3, "A22", "Offer to Kevin Susanto: a growth deal (published OKX DEX referral fee share; brackets, bonus and campaign need OKX approval)", F_SECTION, wrap=False)
    header(t3, 23, ["Term", "Value", "Note"])
    p0 = partners[0]
    offer = [("Retainer", "None", None, False, ""),
             ("Fee share up to milestone 1", 0.30, PCT, True, "Inside OKX's published DEX referral tiers (up to 50%)"),
             ("Fee share between milestones 1 and 2", 0.35, PCT, True, "Marginal: applies only to the volume in this bracket"),
             ("Fee share above milestone 2", 0.40, PCT, True, ""),
             ("Milestone cash bonus (share of swap fees at the milestone reached)", 0.10, PCT, True, "Paid in USDT within 7 days; each payment tops up to this share"),
             ("Milestone 1: referred swaps in a month", 250000, MONEY, True, ""),
             ("Milestone 2", 500000, MONEY, True, ""),
             ("Milestone 3", 1000000, MONEY, True, ""),
             ("Launch campaign", "Co-hosted live session and a 7-day trading competition for his Telegram", None, False, "Prize pool needs OKX approval"),
             ("Follower discount", "Up to 10%, from the partner's share", None, False, ""),
             ("Swap fee his followers pay", "0.5% on memecoins (published OKX DEX fee)", None, False, ""),
             ("What he needs most", p0["need"], None, False, ""), ("Support from us", p0["support"], None, False, ""),
             ("He earns in his first month: fee share (published fee)",
              f"={MEME_FEE}*(MIN({T1}M{N0},B29)*B25+MIN(MAX({T1}M{N0}-B29,0),B30-B29)*B26+MAX({T1}M{N0}-B30,0)*B27)", MONEY2, False, ""),
             ("He earns in his first month: milestone bonus",
              f"=IF({T1}M{N0}>=B31,B31,IF({T1}M{N0}>=B30,B30,IF({T1}M{N0}>=B29,B29,0)))*{MEME_FEE}*B28", MONEY2, False, ""),
             ("Rules for all partners", OF.cell(10, 1).value.replace("Rules for all: ", ""), None, False, "")]
    for i, (aa, b, fmt, inp, note) in enumerate(offer):
        cell(t3, 24 + i, 1, aa, bold=True)
        cell(t3, 24 + i, 2, b, fmt=fmt, inp=inp)
        if note:
            cell(t3, 24 + i, 3, note)
    put(t3, "A41", "Support every partner gets", F_SECTION, wrap=False)
    for i, r in enumerate(range(26, 32)):
        cell(t3, 42 + i, 1, f"{i + 1}.")
        cell(t3, 42 + i, 2, OF.cell(r, 1).value)
    widths(t3, {"A": 36, "B": 110, "C": 40})

    # ============================================================ Task 4
    DP = L["30-day plan"]
    put(t4, "A1", "Task 4: first 30 days (calendar days 1-30 of the job)", F_TITLE, wrap=False)
    put(t4, "A2", "Main metrics: FTDs, then trades a day; swaps and revenue follow. Partners go live in three waves: days 10, 17 and 24.", wrap=False)
    put(t4, "A4", "A. Weekly targets", F_SECTION, wrap=False)
    header(t4, 5, ["Target", "Week 1 (days 1-7)", "Week 2 (days 8-14)", "Week 3 (days 15-21)", "Week 4 (days 22-30)", "Day 30", "How it's set"])
    R0, R1 = 25, 32
    weekly = [
        ("Prospects scored", 25, 20, 15, 10, "=SUM(B6:E6)", "Score everyone before messaging"),
        ("Outreach messages (warm intro first)", 25, 20, 15, 10, "=SUM(B7:E7)", "100% of scored"),
        ("Replies", 8, 6, 5, 3, "=SUM(B8:E8)", "~30% of messages"),
        ("Meetings (call or in person)", 3, 6, 5, 4, "=SUM(B9:E9)", "~80% of replies; the Jakarta day (Kevin Susanto, Bevan) is in week 2"),
        ("Pilots live (running total, our 8)", 0, f'=COUNTIF($C${R0}:$C${R1},"<=14")', f'=COUNTIF($C${R0}:$C${R1},"<=21")',
         f'=COUNTIF($C${R0}:$C${R1},"<=30")', "=E10",
         "Waves: the top 3 (Kevin Susanto, Bevan, Ponyin with the CVC network) on day 10; GuarEmperor, Brandon Salim and Crypto Ndo on day 17; DegenRex and the Meteora night on day 24. If one of the top 3 says no in week 1, GuarEmperor moves into wave 1"),
        ("Pilot B, then C, then A (the brief's hypotheticals)", "Signing B", "B live from day 10", "B live; competition running", "B live", "B live", "B first; C starts once B passes 2x its forecast (month 2); A later on share only. Counted separately because the volumes are the brief's estimates"),
        ("Activated partners (running total, our 8)", 0, 0, 0, f'=COUNTIF(R{R0}:R{R1},"Yes")', "=E12", "FTDs >= 10 and trades a day >= 50 by day 30 (section B)"),
        ("FTDs (running total)", 0, f"=SUM(H{R0}:H{R1})", f"=SUM(I{R0}:I{R1})", f"=SUM(J{R0}:J{R1})", "=E13", "Task 1 FTDs x share of each partner's month live; event FTDs on the night"),
        ("Trades a day (last day of the week)", 0, f"=SUM(K{R0}:K{R1})", f"=SUM(L{R0}:L{R1})", f"=SUM(M{R0}:M{R1})", "=E14", "FTDs so far x trades a day per FTD + FnF power traders x trades a day each"),
        ("Swaps from our 8 (running total)", 0, f"=SUM(N{R0}:N{R1})", f"=SUM(O{R0}:O{R1})", f"=SUM(P{R0}:P{R1})", "=E15", "Calendar days: FTDs arrive through the month, so this trails Task 1's first-month cohort"),
        ("Swaps from the pilots (running total)", 0, "=SUM(N34:N36)", "=SUM(O34:O36)", "=SUM(P34:P36)", "=E16", "B at forecast ($2M a month) from day 10; C and A start after day 30"),
        ("Revenue at 0.04% (running total)", f"=(B15+B16)*{RATE}", f"=(C15+C16)*{RATE}", f"=(D15+D16)*{RATE}", f"=(E15+E16)*{RATE}", "=E17", "Swaps x 0.04%"),
        ("Revenue after partner shares (running total)", 0,
         f"=(SUMPRODUCT(N{R0}:N{R1},1-$G${R0}:$G${R1})+SUMPRODUCT(N34:N36,1-$G$34:$G$36))*{RATE}",
         f"=(SUMPRODUCT(O{R0}:O{R1},1-$G${R0}:$G${R1})+SUMPRODUCT(O34:O36,1-$G$34:$G$36))*{RATE}",
         f"=(SUMPRODUCT(P{R0}:P{R1},1-$G${R0}:$G${R1})+SUMPRODUCT(P34:P36,1-$G$34:$G$36))*{RATE}", "=E18",
         "Revenue x (1 - each partner's share; pilots at their own shares)"),
        ("Monthly run-rate on day 30 (our 8 + pilot B)", "", "", "", f"=(SUM(Q{R0}:Q{R1})+SUM(Q34:Q36))*30", "=E19",
         f'="Swaps a day on day 30 x 30 -> $"&TEXT(E19*{RATE},"#,##0")&" a month at 0.04%"'),
    ]
    for i, rowvals in enumerate(weekly):
        r = 6 + i
        for c, vv in enumerate(rowvals, 1):
            fmt = None
            if r in (13, 14):
                fmt = NUM1
            elif r in (15, 16, 19):
                fmt = MONEY
            elif r in (17, 18):
                fmt = MONEY2
            cell(t4, r, c, vv, fmt=fmt if 2 <= c <= 6 else None, inp=(r <= 9 and 2 <= c <= 5), bold=(c == 1))

    put(t4, "A22", "B. Calendar ramp behind the weekly targets (links to Task 1; posted FTDs arrive evenly through the partner's month, event FTDs on the night)", F_SECTION, wrap=False)
    header(t4, 24, ["Partner", "Event night?", "Go-live day", "Live days by day 14", "Live days by day 21", "Live days by day 30",
                    "Partner share", "FTDs by day 14", "FTDs by day 21", "FTDs by day 30", "Trades a day, day 14",
                    "Trades a day, day 21", "Trades a day, day 30", "Swaps by day 14", "Swaps by day 21", "Swaps by day 30",
                    "Swaps a day on day 30", "Activated by day 30?"])
    for i, p in enumerate(partners):
        r, t = R0 + i, N0 + i
        seg = f"MATCH({T1}$B{t},{SEG},0)"
        tpd_r, tpd_f = f"INDEX({M}$D$19:$D$20,{seg})", f"INDEX({M}$E$19:$E$20,{seg})"
        ftd, ptn, start, month = f"{T1}$F{t}", f"{T1}$H{t}", st["b"], DAYS
        vday, vfday = f"{v['b']}/{month}", f"{vf['b']}/{month}"
        cell(t4, r, 1, f"={T1}A{t}")
        cell(t4, r, 2, p["event"], inp=True)
        cell(t4, r, 3, p["go_live"], inp=True)
        for k, d in enumerate((14, 21, 30)):
            cell(t4, r, 4 + k, f"=MAX(0,{d}-$C{r})")
        cell(t4, r, 7, f"={T1}P{t}", fmt=PCT)
        for k, lv in enumerate("DEF"):
            L_ = f"{lv}{r}"
            cell(t4, r, 8 + k, f'=IF({L_}=0,0,IF($B{r}="Yes",{ftd},{ftd}*MIN({L_},{month})/{month}))', fmt=NUM1)
            fcol = "HIJ"[k]
            cell(t4, r, 11 + k, f"=IF({L_}=0,0,{fcol}{r}*{tpd_r}+IF({L_}>={start},{ptn}*{tpd_f},0))", fmt=NUM1)
            cell(t4, r, 14 + k, f'=IF({L_}=0,0,IF($B{r}="Yes",{ftd}*{vday}*{L_},{ftd}*{vday}*{L_}^2/(2*{month}))+{ptn}*{vfday}*MAX(0,{L_}-{start}+1))', fmt=MONEY)
        cell(t4, r, 17, f"=IF(F{r}=0,0,J{r}*{vday}+IF(F{r}>={start},{ptn}*{vfday},0))", fmt=MONEY)
        cell(t4, r, 18, f'=IF(AND(J{r}>=10,M{r}>=50),"Yes","No")')
    cell(t4, 33, 1, "Total, our 8", bold=True)
    for c in range(8, 18):
        cell(t4, 33, c, f"=SUM({col(c)}{R0}:{col(c)}{R1})", fmt=t4.cell(R0, c).number_format, bold=True)
    for i, (nm, src, live) in enumerate((("Pilot B (brief)", "C", 10), ("Pilot C (brief)", "D", 45), ("Pilot A (brief)", "B", 60))):
        r = 34 + i
        cell(t4, r, 1, nm)
        cell(t4, r, 2, "No")
        cell(t4, r, 3, live, inp=True)
        for k, d in enumerate((14, 21, 30)):
            cell(t4, r, 4 + k, f"=MAX(0,{d}-C{r})")
            cell(t4, r, 14 + k, f"={T2}${src}$6/{DAYS}*{'DEF'[k]}{r}", fmt=MONEY)
        cell(t4, r, 7, f"={T2}${src}$11", fmt=PCT)
        cell(t4, r, 17, f"=IF(F{r}=0,0,{T2}${src}$6/{DAYS})", fmt=MONEY)
    put(t4, "A23", "Activated partner: at least 10 FTDs and at least 50 trades a day within 30 days live, at least 30% of FTDs still "
                   "trading after 14 days and under 5% flagged wallets. The ramp checks the first two; retention and flags come from the dashboard.", wrap=False)

    put(t4, "A37", "C. What each partner delivers in its first month", F_SECTION, wrap=False)
    header(t4, 38, ["Partner", "Deliverables", "FnF power traders", "FTD target (base)", "Trades a day (base, once all are trading)", "Swaps (first month, base)"])
    for i, p in enumerate(partners):
        r, t = 39 + i, N0 + i
        cell(t4, r, 1, f"={T1}A{t}")
        cell(t4, r, 2, p["deliverables"])
        cell(t4, r, 3, f"={T1}H{t}")
        cell(t4, r, 4, f"={T1}F{t}", fmt=NUM1)
        cell(t4, r, 5, f"={T1}I{t}", fmt=NUM1)
        cell(t4, r, 6, f"={T1}M{t}", fmt=MONEY)

    put(t4, "A48", "D. What we monitor, and what we change when it slips", F_SECTION, wrap=False)
    header(t4, 49, [DP.cell(35, c).value for c in range(1, 8)])
    for i, r in enumerate(range(36, 45)):
        row = 50 + i
        vals = [DP.cell(r, c).value for c in range(1, 8)]
        if vals[0] == "Power traders active":
            vals[0], vals[1] = "FnF power traders active", "Introduced FnF power traders trading this week"
            vals[3] = "5 per partner by day 4 (3 for Kevin Susanto and DegenRex; 15 across Ponyin's network)"
        elif vals[0] == "Swaps per power trader":
            vals[0], vals[1], vals[3], vals[4] = "Swaps per FnF power trader", "Swap volume a day per FnF power trader", ">= $547 a day (the OKX Wallet Boost benchmark)", "Under $250"
        elif vals[0] == "Outreach reply rate":
            vals[5] = ("Connections of connections: ask live partners and their circles (Kevin Susanto, Bevan, the Degen Night hosts) "
                       "for direct intros to the decision-maker, skip the public inbox, and lead with the offer itself: fee share plus "
                       "a milestone cash bonus paid within 7 days, no retainer to negotiate")
        elif vals[0] == "Compliance gate":
            vals[5] = "Hold paid posts; private 1:1 FnF onboarding continues"
        for c, vv in enumerate(vals, 1):
            cell(t4, row, c, vv, bold=(c == 1))

    put(t4, "A60", "E. Operating cadence, in-person day and day-30 decisions", F_SECTION, wrap=False)
    header(t4, 61, ["When", "What", "Output"])
    cadence = [(DP.cell(r, 1).value, DP.cell(r, 2).value, DP.cell(r, 3).value) for r in range(47, 53)]
    cadence[0] = (cadence[0][0], "FTDs, trades a day and FnF power traders for every live partner", cadence[0][2])
    cadence[3] = ("Week 2, Wednesday",
                  "In-person day in Jakarta, set on a Wednesday so a weekly event is always on. 10:00 Kevin Susanto, set up by OKX's "
                  "regional marketing team, which already works with him. 13:00 Narrative Labs with Bevan: sign, then onboard his first "
                  "Masterclass traders 1:1 on site (set up by DM a week ahead). 19:00 the next LP Army night if one is listed on Luma, "
                  "otherwise W3AI Wednesday at JURA Kemanggisan (weekly), to meet the next wave. Fallback: same-day video call, and the "
                  "evening event regardless", "Meetings held even if one cancels")
    cadence[4] = ("Day 14", "Mid-point: partners under 10 FTDs or 50 trades a day get the FnF power-trader track again or pause", "Time reallocated")
    cadence[5] = ("Day 30", "Decision per partner: continue (activated, under 5% flagged), renegotiate (FnF trade but under 10 FTDs or 50 trades a day), stop (under 5 FTDs or over 10% flagged); month-2 plan to manager", "Decisions and month-2 plan")
    for i, (aa, b, c) in enumerate(cadence):
        row = 62 + i
        cell(t4, row, 1, aa, bold=True)
        cell(t4, row, 2, b)
        cell(t4, row, 3, c)
    widths(t4, {"A": 34, "B": 60, "C": 16, "D": 18, "E": 18, "F": 18, "G": 40, "H": 12, "I": 12, "J": 12, "K": 12, "L": 12,
                "M": 12, "N": 14, "O": 14, "P": 14, "Q": 14, "R": 14})
    t4.freeze_panes = "B6"

    # ============================================================ research-only tabs
    if research:
        rm = wb.create_sheet("Research merge")
        DD = R["Shortlist deep-dive"]
        n_dd = sum(1 for rr in range(2, DD.max_row + 1) if DD.cell(rr, 2).value)
        put(rm, "A1", f"Research merge: the research sheet's {n_dd} accounts, what we took and why (research copy only)", F_TITLE, wrap=False)
        put(rm, "A2", "Source: Indonesia Crypto KOL & Partner Research (Shortlist deep-dive and Volume proof tabs, 20-21 Nov 2026). Our model uses the deep-dive's reach pool P and rates.", wrap=False)
        put(rm, "A4", "A. Shortlist deep-dive (as delivered) and our decision", F_SECTION, wrap=False)
        hd = ["Rank (base volume)", "Account", "Founder / lead of a group?", "Pool P", "FTD low", "FTD base", "FTD high",
              "Volume low", "Volume base", "Volume high", "Net revenue base (0.04%)", "Research priority", "Our decision", "Why"]
        header(rm, 5, hd)
        decide = {
            "Kevin": ("In the 8 (#1)", "Most FTDs, a verified top fomo trader, and the fastest yes through OKX's regional team"),
            "GuarEmperor": ("In the 8 (#4)", "FnF circle and tools Discord; we add 20% of his ~16.5K Telegram to P (from the written research)"),
            "Brandon": ("In the 8 (#5)", "Second-largest measured reach; posts a trading referral code on almost every trade; founder of a79"),
            "Bevan": ("In the 8 (#2)", "Verified wallet and a live 1-to-10 SOL run, a paid community he founded, meetable in Jakarta"),
            "Nicholas": ("Bench (month 2)", "FTD source with no trading proof; airdrop community"),
            "Orpheu": ("Bench (month 2), as Stress Capital", "Largest owned clan Discord we could measure, but no trading proof yet"),
            "Ponyin": ("In the 8 (#3, leads the CVC Ventures network)", "Verified Trojan PnL card; brings ApexFNF, his FnF, and opens the CVC team"),
            "Tequila": ("Bench (month 2)", "128K Discord but 1.2% online; NFT audience"),
            "VIBI": ("Bench (month 2)", "Beginner onboarding; no trading proof"),
            "DegenRex": ("In the 8 (#7)", "Verified live win; we leave out the 1.5x boost so every partner uses one formula (P 1,697)"),
            "Luna": ("Screened out", "Grows the Cove bot: rival tool tie"),
            "mage": ("Bench (month 2)", "Moved his main wallet from OKX to Bitget Wallet on 11 Nov: win-back later"),
            "All Good": ("Access channel", "Caster and connector in Kevin, Bevan and Brandon's circle; Emperor Journal member; no group of his own"),
            "Friday7th": ("Via Meteora LP Army (#8)", "Co-hosts LP Army; his reach sits inside the Meteora estimate"),
            "celinejoo": ("Screened out", "~900 followers"),
            "zeraa": ("Screened out", "Team role only; reach not measured"),
            "Bryan": ("Removed", "Leads Hyperliquid Indonesia, a rival perp venue: OKX can't send traffic there"),
            "Limitmaxi": ("Screened out", "~440 followers; guest trader only"),
            "edb": ("Screened out", "Reach not measured"),
            "melon": ("Screened out", "Low activity"),
            "marvellalexanderr": ("Not assessed", "Account not identified"),
            "reyk1cv_": ("Not assessed", "Account not identified"),
        }
        row = 6
        for rr in range(2, DD.max_row + 1):
            vals = [DD.cell(rr, c).value for c in (1, 2, 6, 9, 10, 11, 12, 13, 14, 15, 16, 22)]
            if not vals[1]:
                continue
            key = next(k for k in decide if k.lower() in str(vals[1]).lower())
            for c, vv in enumerate(vals, 1):
                fmt = NUM0 if c in (4, 5, 6, 7) else (MONEY if c in (8, 9, 10) else (MONEY2 if c == 11 else None))
                cell(rm, row, c, vv, fmt=fmt)
            cell(rm, row, 13, decide[key][0], bold=True)
            cell(rm, row, 14, decide[key][1])
            row += 1
        row += 1
        put(rm, f"A{row}", "B. Reconciliation notes", F_SECTION, wrap=False)
        notes = [
            ("Ranking", "Kevin Susanto #1, Bevan #2, Ponyin with the CVC network #3, GuarEmperor #4, as you set it. The weights follow: proof 15%, FTD potential 25%, FnF 20%, rival flow 10%, reachability 15%, segment 5%, risk 10%."),
            ("CVC Ventures network", "One partner row for 0xRusdi (+ Tomket Lovers), Ponyin (+ ApexFNF) and Setya, as you asked. Team roles come from your X check; CVC_Ventures and ApexFNF_io didn't show up in web search. Reach adds four overlapping audiences, so it is an upper bound."),
            ("Tomket Lovers", "On Bitget Wallet Indonesia's Superfans list (CEX KOL map tab): clear non-exclusivity before signing."),
            ("DegenRex's 1.5x boost", "The deep-dive multiplies his reach by 1.5 for the win (P 2,546, base $15K). We leave it out so all 8 use one formula (P 1,697, about $10K before FnF); the win counts in his proof score instead."),
            ("Brandon Salim", "Reach measured (median 14,342 over 25 posts). a79's size is not public, so a79 adds FnF traders (5), not reach."),
            ("Bevan's pool P", "Volume proof counts all 9,238 Discord members (P 17,494, base $105K); the deep-dive counts only the 984 online (P 9,240, base $55K). We use the deep-dive: online members are the ones a post reaches."),
            ("Kevin Susanto's pool P", "Volume proof counts all 2,860 Telegram subscribers (P 44,108); the deep-dive counts 20% (P 41,820). We use the deep-dive."),
            ("GuarEmperor's pool P", "The deep-dive left Telegram out (P 15,885). The written research found ~16.5K on his 'Enjoy being poor' channel, so we add 20% (P 19,185). Snapshot from a third-party viewer: verify on t.me."),
            ("Written research vs the sheet", "The written shortlist sized FTDs from visits x conversion (Bevan 40 FTDs, $112K; Ponyin 30, $144K; Kevin 28, $120K). The sheet uses one formula for every account. We use the sheet so every partner is sized the same way."),
            ("Timing", "The research's volume is a cohort's first month of trading. Task 4's calendar ramp counts calendar days, so FTDs who arrive late in the month trade into month 2 and day-30 swaps trail Task 1."),
            ("FnF power traders", "The research leaves out partners' own trading. We add only the traders each partner introduces, at the research's high per-trader rate ($16,400 a month, @AiSlimx on OKX Wallet Boost); the partner's own trading is evidence, not forecast volume."),
            ("Kevin Susanto's OKX link", "He already holds okx.ac/join/VVIP7777. Route through OKX's regional marketing team and confirm which DEX volume would be new."),
            ("Reach for partners the research didn't measure", "0xRusdi, Setya, Crypto Ndo and MeteoraIDN use 20% of followers, the median views-to-followers ratio across the 11 accounts the research measured. Measure their median views before sending."),
            ("Bryan.HL", "Removed: he leads Hyperliquid Indonesia, a rival perp venue."),
        ]
        header(rm, row + 1, ["Topic", "Note"])
        for i, (aa, b) in enumerate(notes):
            cell(rm, row + 2 + i, 1, aa, bold=True)
            cell(rm, row + 2 + i, 2, b)
        row = row + 2 + len(notes) + 1
        put(rm, f"A{row}", "C. Volume proof tab (as delivered)", F_SECTION, wrap=False)
        VP = R["Volume proof"]
        header(rm, row + 1, [VP.cell(1, c).value for c in (2, 5, 6, 8, 9, 10, 11, 12, 14, 19)])
        rr_out = row + 2
        for rr in range(2, 13):
            if not VP.cell(rr, 2).value:
                continue
            vals = [VP.cell(rr, c).value for c in (2, 5, 6, 8, 9, 10, 11, 12, 14, 19)]
            for c, vv in enumerate(vals, 1):
                cell(rm, rr_out, c, vv, fmt=NUM0 if c in (2, 3, 8) else (MONEY if c == 9 else None))
            rr_out += 1
        widths(rm, {"A": 26, "B": 44, "C": 30, "D": 12, "E": 10, "F": 10, "G": 10, "H": 14, "I": 14, "J": 14, "K": 14, "L": 14, "M": 30, "N": 60})

        rv = wb.create_sheet("Checks and history")
        put(rv, "A1", "Checks and history (research copy only)", F_TITLE, wrap=False)
        put(rv, "A3", "A. Verify before sending", F_SECTION, wrap=False)
        header(rv, 4, ["#", "Check", "Where it shows"])
        todo = [
            ("Re-open the 8 partner links and re-check follower counts, median views and quoted posts on-platform", "Task 1, columns C and H; section C"),
            ("Kevin Susanto: with OKX's regional team, what okx.ac/join/VVIP7777 already earns, who owns it, and how DEX tracking stays separate", "Task 1 row 12; Task 3"),
            ("CVC Ventures network: confirm the team (Setya, Ponyin, 0xRusdi), ApexFNF and Tomket clan sizes, who signs, and Tomket's Bitget Wallet Superfans status", "Task 1 row 15"),
            ("Bevan: which wallet he trades from now (the published one is idle since 30 Jul)", "Task 1 row 13"),
            ("Brandon Salim: a79 member count and fomo referral results; ask Kevin for the intro", "Task 1 row 16"),
            ("Measure median views for 0xRusdi, Setya, Crypto Ndo and MeteoraIDN (reach is estimated at 20% of followers)", "Task 1 section C, column D"),
            ("GuarEmperor: confirm his FnF circle and the ~16.5K Telegram count on t.me", "Task 1 row 14; section C"),
            ("OKX DEX referral tiers: confirm the live tier table (regional help pages differ)", "Task 3 offer; deck footers"),
            ("Legal: promotion of OKX DEX in Indonesia under Law 4/2026; partner POJK 6/2026 certification status (Brandon Salim included)", "Task 4, compliance gate"),
        ]
        for i, (aa, b) in enumerate(todo):
            cell(rv, 5 + i, 1, i + 1)
            cell(rv, 5 + i, 2, aa)
            cell(rv, 5 + i, 3, b)
        row = 5 + len(todo) + 1
        put(rv, f"A{row}", "B. Partner changes after review", F_SECTION, wrap=False)
        header(rv, row + 1, ["Change", "Why"])
        changes = [
            ("Hyperliquid Indonesia replaced by GuarEmperor and his FnF circle", "Hyperliquid is a rival venue: OKX can't send traffic there"),
            ("Bael calls replaced by Superteam Indonesia (since benched)", "Bael is a Bitget Wallet Indonesia ambassador"),
            ("Four cryptokolz.com-only prospects removed from the longlist", "Promotional directory; replaced by four licensed exchanges screened out as competitors"),
            ("Research merge: Bevan, Kevin Susanto and Stress Capital in; TwipsX and Superteam benched", "Measured reach and verified trading from the research sheet"),
            ("Ranking set to Kevin Susanto #1, Bevan #2, GuarEmperor #3", "Your call: FTDs first, then trades a day; weights rebalanced to match"),
            ("0xRusdi (+ Tomket Lovers), Ponyin (+ ApexFNF) and Setya merged into one CVC Ventures network", "One network deal with the most FnF traders (15)"),
            ("Brandon Salim + a79 and DegenRex in; Stress Capital to the bench", "Measured reach and verified trading; Stress Capital has no trading proof yet"),
            ("All Good (@mudacumasekali) added as an access channel", "Connector and event caster, no group of his own"),
            ("Top 3 set to Kevin Susanto, Bevan and Ponyin (with access to the CVC circle and Setya)", "Ponyin's circle includes the CVC team and the airdrop traders; GuarEmperor moves to #4"),
            ("Only Kevin Susanto and Bevan counted as meetable in person", "The two you can realistically meet in Jakarta"),
            ("Task 2 reworked: B first on a growth deal, then C, then A", "Brackets, milestone bonus and funded campaigns for B; a $500 + $500 retainer and higher share for C; A on share only later"),
            ("Write-up moved to a Google Doc with a 6-slide companion deck", "The doc carries the detail (4 pages); slides carry the main points"),
            ("Volume model replaced by the research method plus an FnF line", "FTDs = P x s x a, swaps = FTDs x v (low/base/high); FnF at $16,400 a month each, down from $50,400"),
            ("Bryan.HL removed", "Hyperliquid Indonesia lead (rival venue)"),
            ("20 generic screened-out longlist rows swapped for the research's named accounts", "Media sites, campus clubs and size-unknown groups; the count stays at 100"),
        ]
        for i, (aa, b) in enumerate(changes):
            cell(rv, row + 2 + i, 1, aa)
            cell(rv, row + 2 + i, 2, b)
        row = row + 2 + len(changes) + 1
        put(rv, f"A{row}", "C. Review findings (from REVIEW.md)", F_SECTION, wrap=False)
        header(rv, row + 1, ["#", "Where", "What was wrong", "Fix"])
        r = row + 2
        for line in REVIEW.read_text().splitlines():
            m = re.match(r"\| (E\d+) \| (.*?) \| (.*?) \| (.*?) \|", line)
            if m:
                for c, vv in enumerate(m.groups(), 1):
                    cell(rv, r, c, vv.replace("**", ""))
                r += 1
        widths(rv, {"A": 40, "B": 70, "C": 70, "D": 70})
    return wb


# ---------------------------------------------------------------- Task 3 texts (Kevin Susanto, our #1)
TASK3 = [
    ("First message (X DM)",
     "Kevin, +$53,566 in one day and #3 on fomo's 24-hour board: your readers watch every trade like that. I'm Fedric, building "
     "OKX's DEX partnerships in Indonesia, and I'd like your followers trading on OKX Wallet through your own link. What I'm "
     "proposing for you: a share of every swap fee they pay, a cash bonus within 7 days each time they pass a volume milestone, and "
     "a co-hosted live trading session for your Telegram. For them: lower fees through your code. It runs alongside your existing "
     "OKX link. If it's interesting, reply here and I'll send a one-page offer; coffee in Jakarta this week, and your link can be "
     "live within 48 hours."),
    ("First message, Bahasa Indonesia version",
     "Kevin, +$53.566 dalam sehari dan #3 di papan 24 jam fomo: pembaca kamu memperhatikan setiap trade seperti itu. Saya Fedric, "
     "membangun kemitraan DEX OKX di Indonesia, dan saya ingin follower kamu trading di OKX Wallet lewat link kamu sendiri. Yang saya "
     "usulkan untuk kamu: bagian dari setiap biaya swap yang mereka bayar, bonus tunai dalam 7 hari setiap kali mereka melewati target "
     "volume, dan sesi live trading bersama untuk Telegram kamu. Untuk mereka: biaya lebih rendah lewat kode kamu. Ini berjalan "
     "bersama link OKX yang sudah kamu punya. Kalau tertarik, balas di sini dan saya kirim penawaran satu halaman; ngopi di Jakarta "
     "minggu ini, dan link kamu bisa aktif dalam 48 jam."),
    ("Follow-up (day 4, no reply)",
     "Hi Kevin, bumping this in case it got buried. I'm closing the first wave of Indonesian DEX partners this week, and the launch "
     "terms (the milestone cash bonus and a co-hosted session for your Telegram) are reserved for that wave. You were first on my "
     "list, so I've held a slot for you. Worth 15 minutes before Friday?"),
    ("Reply to 'another platform pays more'",
     "Fair question. When a platform pays more, there is usually a reason: it's new and buying market share, the rate drops when "
     "the promo ends, it's paid on less volume, or the payout itself is uncertain. Ours may start lower, but it's OKX: stable, paid "
     "on the real volume your followers trade, and it steps up at each milestone with a cash bonus on top. If the numbers grow and "
     "we work well together, the rate and support are open to discussion (bisa dibahas). Let's run 30 days and let the data set "
     "the next terms."),
    ("Meeting question 1",
     "How many of your followers trade onchain every week, and can you show it: fomo referral stats, a wallet list or your Telegram activity?"),
    ("Meeting question 2",
     "Who are the 3 traders in your circle who'd move first, and what would get them to switch: fees, speed or rewards?"),
    ("Meeting question 3",
     "What do your okx.ac link and the Lighter role cover, and what would it take for OKX DEX to be your main onchain recommendation?"),
    ("No unverified promises",
     "Every term in these messages is framed as a proposal. The fee share (up to 50% by tier; confirm the live tier table) and the "
     "follower discount are published OKX DEX referral features. The cash bonus, the co-hosted session and the launch-wave terms "
     "need OKX approval before they become terms, and OKX's regional team confirms how his existing OKX link is tracked before any post."),
]
_k = partners[0]
_f4, _t4, _, _ft = ramp(_k, 4)
_f7, _t7, _, _ = ramp(_k, 7)
_f10, _t10, _, _ = ramp(_k, 10)
_f14, _t14, _s14, _ = ramp(_k, 14)
PLAN14 = [  # targets follow the Task 4 ramp for Kevin Susanto (base case)
    ("Days 1-2", "Set-up: separate DEX link and follower discount, tracked apart from his okx.ac link (agreed with OKX's regional team); fee-share brackets, milestones and payout wallet in writing; compliance check (POJK 6/2026); shared dashboard", "Link live; terms signed"),
    ("Days 2-4", "Activate his circle: 3 traders onboarded 1:1 through the full flow (wallet, funding, first Meme Mode trade, Track alerts); log every drop-off", f"3 FnF traders trading by day 4 (~{_ft:.0f} trades a day)"),
    ("Day 5", "Launch the campaign: co-hosted live trading session on KevinSusantoAlpha and a 7-day trading competition (screened wallets only)", "Session held; first FTDs"),
    ("Days 6-7", "Push the flow: 2 X threads with the link, a pinned Bahasa and English guide, and an FAQ for first-time DEX users", f"~{_f7:.0f} FTDs by day 7"),
    ("Days 8-10", "Fix the funnel: check each step (click, wallet, deposit, first trade, second trade), fix the weakest, and send UI/UX feedback to the product team", f"~{_f10:.0f} FTDs, ~{_t10:.0f} trades a day by day 10"),
    ("Days 11-14", "Competition results and a second session; review with Kevin: FTDs, trades a day, repeat traders and flags; agree days 15-30 and the next bracket", f"~{_f14:.0f} FTDs and ~{_t14:.0f} trades a day by day 14"),
]

for research, name in ((False, "OKX_DEX_BD_Affiliate_Indonesia_Prospect_Spreadsheet.xlsx"), (True, "OKX_research_and_checks.xlsx")):
    wb = build(research)
    wb.save(OUTDIR / name)
    print("saved", name)
