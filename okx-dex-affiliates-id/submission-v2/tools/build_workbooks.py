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


# ---------------------------------------------------------------- v6: plain-language fields, FnF circles, brief-table summaries
_V6 = {
    "Kevin Susanto": dict(
        mult=2, why_reach="A typical post gets ~41K views, plus ~570 active members in his 2,860-member Telegram",
        why_fnf="Traders around a top-3 fomo trader; his own tracked fomo trades total ~$402K, so his circle trades at 2x the standard",
        s_aud="Southeast Asian and Indonesian crypto traders; 169K on X; Jakarta",
        s_evid="#3 on fomo's 24-hour profit board (+$53,566, 19-20 Nov); ~$402K across 99 tracked fomo trades",
        s_route="OKX's regional marketing team, which already works with him; then in person in Jakarta",
        s_prop="Fee share that rises with volume, cash bonus at each milestone, co-hosted live session",
        s_risk="Already has an OKX referral link, so tracking must stay separate; advises Lighter, a rival exchange", type="Large creator (KOL) and active trader",
        s_prio="High: the most first-time traders and the fastest yes"),
    "Bevan / Narrative Labs": dict(
        mult=1, why_reach="~7.5K views a post on his account and ~700 on Narrative Labs', plus ~980 Discord members online",
        why_fnf="Paid Masterclass traders; Bevan's own wallet averaged about $1,000 of swaps a day, the standard rate",
        s_aud="Intermediate and advanced Indonesian memecoin traders; 9,238-member Discord; Jakarta",
        s_evid="Published wallet: ~1,080 SOL swapped May-July, +374 SOL; turned 1 SOL into 10 live at Web3 Week Asia",
        s_route="Direct message, then meet at a Narrative Labs session in Jakarta",
        s_prop="Fee share, a co-branded workshop, a trade button for his Narrative Alpha tool",
        s_risk="Already shares GMGN and Mobee referral codes, so our terms must beat them", type="Trader-educator with a paid community",
        s_prio="High: proven trader we can meet"),
    "Ponyin + CVC network": dict(
        mult=1.5, why_reach="Ponyin ~4.6K views a post; 0xRusdi ~5.2K and Setya ~10.5K (about 20% of followers); Tomket ~4.8K views on its promotions",
        why_fnf="ApexFNF trades around Ponyin (~$170K a month on Trojan, 2x); 0xRusdi's referral book and the Tomket clan at the standard: 1.5x blended",
        s_aud="Indonesian Solana memecoin traders on trading bots; 22.7K on X plus the CVC team's audiences",
        s_evid="Trojan profit-and-loss card: ~$170K of volume in March; 0xRusdi's $8,463 of BasedBot referral pay",
        s_route="Direct message to Ponyin, who opens the CVC Ventures team",
        s_prop="One network deal: a code per member, 15 friends-and-family traders, a network trading competition",
        s_risk="Four people to align; Tomket Lovers also promotes Bitget Wallet", type="Trader network with FnF clans",
        s_prio="High: the most friends-and-family traders from one introduction"),
    "GuarEmperor + FnF": dict(
        mult=1, why_reach="~15.9K views a post, plus ~3.3K active members in his ~16.5K-member Telegram",
        why_fnf="His circle and Emperor Journal members trade from his wallet lists: the standard rate",
        s_aud="Indonesian and global memecoin traders on new chains; 94K on X",
        s_evid="Active trader who calls token buys: $PRL at ~$0.60, then $2.10",
        s_route="Direct message; introduction through All Good, an Emperor Journal member",
        s_prop="Fee share; OKX DEX trading access for his Discord tools; alpha calls executed on OKX Wallet",
        s_risk="His feed mixes alpha calls with whitelist and mint posts", type="Alpha caller with a tools community",
        s_prio="Medium: trader and caller; wave 2"),
    "Brandon Salim + a79": dict(
        mult=1, why_reach="~14.3K views a post (median of his last 25)",
        why_fnf="a79 is an application-only trader group; Brandon says he trades small, so the standard rate",
        s_aud="An actor's mass audience plus stock and crypto traders; 187K on X; Jakarta",
        s_evid="fomo +$32,185 in one day (pinned); posts his trading code on almost every trade",
        s_route="Introduction through Kevin Susanto",
        s_prop="Fee share in the 'trade with me' format he already runs; an a79 group code",
        s_risk="Celebrity reach may not trade on-chain", type="Celebrity trader with his own trader group",
        s_prio="Medium: big reach and a habit of sharing his code"),
    "Crypto Ndo": dict(
        mult=1, why_reach="~11.9K: about 20% of his 59.3K followers see a typical post",
        why_fnf="Traders in his circle: the standard rate",
        s_aud="Indonesian memecoin and Solana traders; 59.3K on X",
        s_evid="Says he trades only on OKX Wallet; posts fomo referral links",
        s_route="Direct message on X",
        s_prop="Fee share, Meme Mode trade threads, live X Spaces",
        s_risk="Runs a memecoin launchpad, so posts need clear disclosure", type="Memecoin trading creator",
        s_prio="Medium: the largest memecoin-trader reach, already an OKX Wallet user"),
    "DegenRex": dict(
        mult=1, why_reach="~1.3K views a post, plus ~370 active members in his 1,845-member Telegram",
        why_fnf="Traders in his Telegram call channel: the standard rate",
        s_aud="Indonesian memecoin traders on BNB Chain, Base and Solana; 1,845 on Telegram",
        s_evid="Won Clash of Trenchers live in Jakarta: 1 SOL to 10 SOL in 20 minutes",
        s_route="Direct message; Ave Indonesia's organisers",
        s_prop="'OKX DEX champion' live sessions and a title-defence competition",
        s_risk="Small reach; Ave.ai and four.meme already back him", type="Champion trader with a call channel",
        s_prio="Medium: live trading proof"),
    "Meteora LP Army": dict(
        mult=1, why_reach="~2.7K: about 20% of MeteoraIDN's 13.4K followers, plus ~300 at one community night",
        why_fnf="Large liquidity providers from the community: the standard rate",
        s_aud="Solana liquidity providers in Jakarta and Bali",
        s_evid="462 and 671 attendees at two community nights",
        s_route="Direct message to the hosts (Evil Panda, Friday7th)",
        s_prop="Co-hosted community night with wallets funded on site",
        s_risk="Liquidity rewards can attract farmers; Bitget Wallet already co-hosts", type="Liquidity-provider community (events)",
        s_prio="Low: LPs swap less often"),
}
_SHORT = {'Kevin Susanto': {'s_aud': 'Indonesian and regional traders; 169K on X; Jakarta', 's_evid': "#3 on fomo's 24-hour profit board (+$53,566, 19-20 Nov)", 's_route': "OKX's regional marketing team; then meet in Jakarta", 's_prop': 'Rising fee share, milestone bonus, co-hosted live session', 's_risk': 'Has an OKX referral link (keep tracking apart); advises Lighter', 's_prio': 'High: most FTDs, fastest yes'}, 'Bevan / Narrative Labs': {'s_aud': 'Advanced memecoin traders; 9,238-member Discord; Jakarta', 's_evid': 'Wallet: ~1,080 SOL swapped May-July, +374 SOL; 1 SOL to 10 live', 's_route': 'Direct message, then a Narrative Labs session in Jakarta', 's_prop': 'Fee share, co-branded workshop, a trade button in his tool', 's_risk': 'Shares GMGN and Mobee codes: our terms must beat them', 's_prio': 'High: proven trader we can meet'}, 'Ponyin + CVC network': {'s_aud': "Solana bot traders; 22.7K on X plus the CVC team's audiences", 's_evid': "Trojan card: ~$170K in March; 0xRusdi's $8,463 referral pay", 's_route': 'Direct message to Ponyin, who opens the CVC Ventures team', 's_prop': 'One network deal: a code per member, 15 FnF traders, a competition', 's_risk': 'Four people to align; Tomket Lovers also promotes Bitget Wallet', 's_prio': 'High: most FnF traders, one intro'}, 'GuarEmperor + FnF': {'s_aud': 'Memecoin traders on new chains; 94K on X', 's_evid': 'Active trader who calls token buys: $PRL at ~$0.60, then $2.10', 's_route': 'Introduction through All Good, an Emperor Journal member', 's_prop': 'Fee share; OKX DEX access for his Discord tools', 's_risk': 'Feed mixes alpha calls with whitelist and mint posts', 's_prio': 'Medium: trader and caller; wave 2'}, 'Brandon Salim + a79': {'s_aud': "An actor's mass audience plus retail traders; 187K on X", 's_evid': 'fomo +$32,185 in one day; posts his code on almost every trade', 's_route': 'Introduction through Kevin Susanto', 's_prop': "Fee share in his 'trade with me' format; an a79 group code", 's_risk': 'Celebrity reach may not trade on-chain', 's_prio': 'Medium: reach plus a trading code'}, 'Crypto Ndo': {'s_aud': 'Memecoin and Solana traders; 59.3K on X', 's_evid': 'Trades only on OKX Wallet; posts fomo referral links', 's_route': 'Direct message on X', 's_prop': 'Fee share, Meme Mode trade threads, live X Spaces', 's_risk': 'Runs a memecoin launchpad: posts need clear disclosure', 's_prio': 'Medium: reach; uses OKX Wallet'}, 'DegenRex': {'s_aud': 'Multichain memecoin traders; 1,845 on Telegram', 's_evid': 'Won Clash of Trenchers live: 1 SOL to 10 in 20 minutes', 's_route': "Direct message; Ave Indonesia's organisers", 's_prop': 'Champion live sessions and a title-defence competition', 's_risk': 'Small reach; Ave.ai and four.meme already back him', 's_prio': 'Medium: live trading proof'}, 'Meteora LP Army': {'s_aud': 'Solana liquidity providers in Jakarta and Bali', 's_evid': '462 and 671 attendees at two community nights', 's_route': 'Direct message to the hosts (Evil Panda, Friday7th)', 's_prop': 'Co-hosted community night, wallets funded on site', 's_risk': 'Liquidity rewards attract farmers; Bitget Wallet co-hosts', 's_prio': 'Low: LPs swap less often'}}
for _p in partners:
    _p.update(_V6[_p["short"]])
    _p.update(_SHORT[_p["short"]])
# rolling go-live: partners go live the day they sign
for _p, _d in zip(partners, (8, 8, 8, 14, 14, 14, 20, 20)):
    _p["go_live"] = _d
_ge = next(q for q in partners if q["short"].startswith("GuarEmperor"))
_ge["reason"] = "An active trader who calls token buys, with a circle that trades from his tools; second wave"
_ge["evidence"] = ("Active trader who calls token buys: called $PRL on 11 Sep 2026 at about $0.60, $2.10 by mid-October (3.5x, his chart); "
                   "posted almost $10K of profit on the trade (14 Oct). Members post gains from his tools (GMGN +$802, +$387); his "
                   "smart-money wallet lists are formatted for BasedBot and GMGN import. X 94,421, median 15,885 views a post; "
                   "'Enjoy being poor' Telegram ~16.5K (third-party snapshot)")
_PLAIN = (("PnL", "profit and loss"), ("24h", "24-hour"), (" X DM", " X direct message"), ("X DM", "X direct message"),
          ("DM,", "direct message,"), ("DM;", "direct message;"), (" DM", " direct message"), ("WL", "whitelist"),
          ("CT degens", "crypto-Twitter traders"), ("SEA ", "Southeast Asian "), ("BSC", "BNB Chain"), ("DEX API", "DEX trading interface (API)"),
          ("MEV", "front-running"), ("DLMM", "liquidity-pool"), ("1:1", "one-to-one"), ("onchain", "on-chain"), ("Onchain", "On-chain"),
          ("W3W", "Web3 Week Asia"), ("CEX", "exchange"), ("POJK 6/2026 finfluencer certification", "finfluencer certification (OJK rule 6/2026)"),
          ("POJK 6/2026", "OJK rule 6/2026"), ("degen circle", "trading circle"), ("trencher KOLs", "trading creators"), ("trenchers", "memecoin traders"),
          ("Trencher", "Memecoin trader"), ("trencher", "memecoin trader"))


def plain(t):
    for _a, _b in _PLAIN:
        t = t.replace(_a, _b)
    return (t.replace("emecoin memecoin", "emecoin").replace("on-chain on-chain", "on-chain")
            .replace("Clash of Memecoin traders", "Clash of Trenchers").replace("Clash of memecoin traders", "Clash of Trenchers"))


for _p in partners:  # plain words for the submission
    for _k in ("name", "evidence", "assumptions", "proposal", "risk", "reason", "route", "audience", "fnf", "vampire", "deliverables", "support", "need", "check", "airdrop"):
        if isinstance(_p.get(_k), str):
            _p[_k] = plain(_p[_k])

FNF_START = 4


def ramp(p, L, s=0.02, a=0.15, v=2000, fnf_day=1000, start=FNF_START, days=30, tr=25, tf=120):
    """Task 4 ramp for a posted partner (base case), L days after go-live: FTDs, trades a day, swaps, FnF trades a day."""
    ftd_total = p["pool"] * s * a
    ftd = ftd_total * min(L, days) / days
    daily = fnf_day * p["mult"]
    fnf_tpd = p["ptraders"] * daily / tf if L >= start else 0
    tpd = ftd * v / days / tr + fnf_tpd
    swaps = ftd_total * v / days * L * L / (2 * days) + p["ptraders"] * daily * max(0, L - start + 1)
    return ftd, tpd, swaps, fnf_tpd



# ---------------------------------------------------------------- v6b: the user's edited write-up (22 Nov 2026)
_BY = {p["short"]: p for p in partners}
_UPD = {
    "Kevin Susanto": dict(
        type="Trading Educator", trading_type="Large-size memecoin trader on fomo",
        s_route="Direct introduction from colleagues and a special long-term partnership offer",
        s_prop="Rising fee share, VIP-style tiers, co-hosted live session",
        s_risk="Advises several competitor brands", s_prio="High: most FTDs with good-quality users",
        why_fnf="Traders around a top-3 fomo trader, at the standard $1,000 a day",
        scores=(5, 5, 4, 4, 5, 5, 3)),
    "Bevan / Narrative Labs": dict(
        type="Trading Community", trading_type="Narrative-driven Solana memecoin trader",
        s_aud="Advanced memecoin traders; 21K on X, 9,238-member Discord",
        s_route="Direct message, then a Narrative Labs session for the partnership",
        s_prop="Fee share, co-branded workshop, connected trade button in his tool",
        s_risk="Shares GMGN and Mobee codes as well",
        pool=25000,
        why_reach="Guesstimate: ~7.5K views a post on X, ~980 Discord members online, plus his Instagram, TikTok and YouTube audiences",
        scores=(5, 4, 5, 4, 5, 5, 3)),
    "Ponyin + CVC network": dict(
        short="0xRusdi + CVC network",
        name="0xRusdi (@0xRusdi) + CVC Ventures network: Ponyin (+ ApexFNF), Setya, Tomket Lovers",
        link="https://x.com/0xRusdi ; https://x.com/CVC_Ventures",
        type="Trading Community", trading_type="Telegram trading-bot traders",
        s_aud="Solana bot traders; 25.8K on X plus the CVC team's audiences",
        s_evid="BasedBot referral book: ~$3.4M referred, $8,463 earned",
        s_route="Direct introduction from colleagues, who opens the CVC Ventures team",
        s_prop="One community network deal: a code per member, 15 FnF traders, a competition",
        s_risk="Four people to align; Tomket Lovers also promotes Bitget Wallet",
        s_prio="High: most FnF traders, one intro",
        evidence=("0xRusdi: $8,463 of BasedBot referral earnings at the 25% tier (dashboard, 26 Apr 2026), on ~$3.4M of referred swaps; "
                  "Ponyin: Trojan profit-and-loss card, March 2026, ~$170K of volume (1.13K SOL bought, 1.24K sold); "
                  "Tomket Lovers: fomo clan of 13-15 active traders; Setya: X 52,661; CVC Ventures team: 0xRusdi, Ponyin, Setya"),
        route="Direct introduction from colleagues to 0xRusdi, who opens the CVC Ventures team (Ponyin, Setya), ApexFNF and the Tomket clan; meetable in Jakarta",
        reason="0xRusdi's referral book is proof his followers trade through his link, and one introduction opens 15 FnF power traders",
        why_reach="0xRusdi ~5.2K and Setya ~10.5K (about 20% of followers); Ponyin ~4.6K views a post; Tomket ~4.8K views on its promotions",
        why_fnf="5 from 0xRusdi's BasedBot referral book, 5 from ApexFNF (Ponyin's circle), 5 from the Tomket clan",
        check=("0xRusdi's BasedBot referral export (volume, active traders); Ponyin's Trojan dashboard or wallet history; ApexFNF member "
               "wallets; Tomket clan size; who signs for the network"),
        in_person="Yes", scores=(5, 4, 5, 5, 5, 5, 1)),
    "GuarEmperor + FnF": dict(
        type="Trading Community + Onchain Tools Creator", trading_type="Early-entry caller and smart-money tracker",
        s_route="Introduction through Emperor Journal members", s_risk="Feed mixes alpha calls with NFT whitelist and mint posts",
        scores=(4, 4, 5, 4, 4, 4, 3)),
    "Brandon Salim + a79": dict(
        type="Celebrity Influencer", trading_type="Social 'trade with me' memecoin trader",
        s_route="Direct message on Telegram", s_prop="Fee share, Meme Mode trade promotion",
        s_risk="Celebrity reach may be too busy to focus on being an affiliate",
        scores=(4, 5, 4, 5, 4, 5, 2)),
    "Crypto Ndo": dict(type="Trading Community", trading_type="Memecoin trader on OKX Wallet",
                       s_evid="Existing trades on OKX Wallet; posts fomo referral links"),
    "DegenRex": dict(type="Trading Community", trading_type="Fast multichain memecoin scalper",
                     s_route="Direct introduction from colleagues", s_prop="Champion live sessions and a title-defense competition"),
    "Meteora LP Army": dict(type="Onchain Analytics", trading_type="Solana liquidity providers (LPs)",
                            scores=(2, 1, 4, 3, 5, 3, 4)),
}
for _k, _u in _UPD.items():
    _BY[_k].update(_u)
for _p in partners:
    _p["mult"] = 1  # no circle multiplier: every FnF power trader swaps the standard amount
# the new scores put Brandon Salim (4.20) ahead of GuarEmperor (4.10)
_i, _j = partners.index(_BY["GuarEmperor + FnF"]), partners.index(_BY["Brandon Salim + a79"])
partners[_i], partners[_j] = partners[_j], partners[_i]
for _p, _d in zip(partners, (8, 8, 8, 14, 14, 14, 20, 20)):
    _p["go_live"] = _d
assert [p["short"] for p in partners][:5] == ["Kevin Susanto", "Bevan / Narrative Labs", "0xRusdi + CVC network", "Brandon Salim + a79", "GuarEmperor + FnF"]


# ---------------------------------------------------------------- v6c: Laba Lab replaces Meteora LP Army (23 Nov 2026)
_LABA = dict(
    short="Michael Wyann / Laba Lab", name="Michael Wyann Santoso / Laba Lab",
    link="https://laba-lab.com ; https://linkedin.com/in/michaelwyannsantoso",
    type="Trading Educator + Trading Community", trading_type="Market-analysis trader (macro, price action, on-chain flows)",
    audience="Indonesian retail and intermediate traders; YouTube 121K, Telegram channel 10.8K and group 6.9K, Discord 4.6K; Jakarta",
    evidence=("3 years of weekly market updates without a gap (self-reported); paid membership tier and bootcamps (laba-lab.com/membership); "
              "named in Indonesia Crypto Network's 12 Trusted Crypto Influencers (Feb 2024); YouTube 121K, Telegram channel 10.8K and group "
              "6.9K, Discord 4.6K, TikTok 44K (self-reported on LinkedIn, read 19 Nov 2026)"),
    assumptions="How many members trade on-chain rather than exchange futures; follower counts are self-reported; his existing exchange partnerships",
    route="Direct message to Michael on LinkedIn (he says he is open to Web3 partnerships); Laba Lab site contact; Discord admins",
    proposal=("An OKX DEX and Wallet module in his bootcamp; a co-branded on-chain week with live sessions; fee share on a tracked OKX DEX "
              "referral code; 5 paid members onboarded one-to-one in days 1-3"),
    risk="His audience may prefer exchange futures to on-chain swaps; existing exchange partnerships",
    reason="The largest paid trading-education community we can reach, with members who pay to learn trading",
    vampire="Exchange partners (not named publicly)",
    fnf="5 paid members from his membership and bootcamps, onboarded first",
    in_person="No", airdrop="No", segment=TREN, pool=14000, ptraders=5, share=0.3, members=None, per_member="trade",
    basis="Guesstimate: YouTube views plus active Telegram and Discord members",
    deliverables=("Days 1-3: 5 paid members onboarded one-to-one; then an OKX DEX module in his bootcamp and a co-branded on-chain week "
                  "with two live sessions"),
    need="A product his paid members can't get elsewhere, and income on the trading he already teaches",
    support="Bootcamp module and guide in Bahasa; co-hosted live sessions; wallet-level referral dashboard",
    check="Membership size and how many members trade on-chain; YouTube analytics; existing exchange deals",
    line="", scores=(3, 4, 4, 3, 5, 4, 3), go_live=20, event="No", mult=1,
    why_reach=("Guesstimate: ~10K views on a typical YouTube video, plus about 20% of his Telegram channel (10.8K) and group (6.9K), "
               "plus ~460 Discord members online"),
    why_fnf="Paid members from his membership and bootcamps, at the standard $1,000 a day",
    s_aud="Retail and intermediate traders; 121K on YouTube; Jakarta",
    s_evid="3 years of weekly market updates; paid membership and bootcamps",
    s_route="Direct message to Michael on LinkedIn (open to partnerships)",
    s_prop="OKX DEX module in his bootcamp, co-branded on-chain week, fee share",
    s_risk="Audience may prefer exchange futures to on-chain swaps",
    s_prio="Medium: big paid community, Jakarta",
)
partners[:] = [q for q in partners if q["short"] != "Meteora LP Army"]
partners.insert(5, _LABA)  # score 3.80: above Crypto Ndo (3.50) and DegenRex (3.40)
for _p, _d in zip(partners, (8, 8, 8, 14, 14, 14, 20, 20)):
    _p["go_live"] = _d
assert len(partners) == 8 and partners[5]["short"] == "Michael Wyann / Laba Lab"


# ================================================================ v6 body: brief titles, answers on top, plain words
import math

from layout_v6 import MI, T1 as L1, T2 as L2, T3 as L3, T4 as L4, GT, CASES

TREN_OLD, LP_OLD = TREN, LP
TREN, LP = "Memecoin traders", "Liquidity providers (LPs)"
for _p in partners:
    _p["segment"] = TREN if _p["segment"] == TREN_OLD else LP

TERMS = ("Terms: FTD = first-time trader (a referred user's first trade through the partner's link). FnF = friends and family: "
         "the active power traders a partner brings personally. DEX = decentralised exchange (OKX DEX, traded in OKX Wallet). "
         "KOL = key opinion leader (crypto creator). LP = liquidity provider. SOL = Solana's token. USDT = Tether's US-dollar stablecoin.")
F_NOTE = Font(name="Calibri", size=10, italic=True, color="5B667A")


class X(str):
    """A formula expression inside fx(); plain strings become quoted text."""


NL = X("CHAR(10)")


def fx(*parts):
    out = []
    for q in parts:
        if isinstance(q, X):
            out.append(str(q))
        else:
            s_ = str(q)
            for i in range(0, len(s_), 200):  # Excel caps a text constant in a formula at 255 characters
                out.append('"' + s_[i:i + 200].replace('"', '""') + '"')
    return "=" + "&".join(out)


def tx(ref, f="#,##0"):
    return X(f'TEXT({ref},"{f}")')


def td(ref):
    return X(f'TEXT({ref},"$#,##0;-$#,##0")')


def tk(ref):
    return X(f'TEXT(({ref})/1000,"$#,##0")&"K"')


def tm(ref):
    return X(f'"$"&TEXT(ROUND(({ref})/1000000,2),"General")&"M"')


def pc(ref):
    return X(f'TEXT({ref},"0%")')


def flat(parts):
    if isinstance(parts, str):
        return [parts]
    return list(parts)


def bullets(items):
    out = []
    for i, b in enumerate(items):
        if i:
            out.append(NL)
        out.append("• ")
        out.extend(flat(b))
    return out


def value(parts):
    """(cell value, plain-text estimate): a formula when any part is an expression, else plain text with line breaks."""
    parts = flat(parts)
    est = "".join("\n" if q == NL else ("000000" if isinstance(q, X) else q) for q in parts)
    if any(isinstance(q, X) and q != NL for q in parts):
        return fx(*parts), est
    return est, est


def lines(text, width):
    per = max(8, width * 1.15)
    return sum(max(1, math.ceil(len(l) / per)) for l in text.split("\n"))


def mcell(ws, row, c0, c1, v, bold=False, fill=None):
    c = ws.cell(row, c0, v)
    c.font = F_BOLD if bold else F_BODY
    c.alignment = WRAP
    if fill:
        c.fill = fill
    for cc in range(c0, c1 + 1):
        ws.cell(row, cc).border = BOX
    if c1 > c0:
        ws.merge_cells(start_row=row, start_column=c0, end_row=row, end_column=c1)
    return c


def colw(ws, c0, c1):
    return sum(ws.column_dimensions[col(c)].width or 10 for c in range(c0, c1 + 1))


def answers(ws, row, items, b_span, c_span, head=("Question", "One-sentence answer", "Details")):
    """items: (question, answer parts, list of detail bullets). Spans are (first col, last col) for the merged cells."""
    for (c0, c1), h in zip(((1, 1), b_span, c_span), head):
        mcell(ws, row, c0, c1, h, bold=True, fill=FILL_HEAD)
    wa, wb_, wc = colw(ws, 1, 1), colw(ws, *b_span), colw(ws, *c_span)
    for i, (q, a_, d_) in enumerate(items):
        r = row + 1 + i
        va, ea = value(a_)
        vd, ed = value(bullets(d_)) if d_ else ("", "")
        mcell(ws, r, 1, 1, q, bold=True)
        mcell(ws, r, *b_span, va)
        mcell(ws, r, *c_span, vd)
        ws.row_dimensions[r].height = 15 * max(lines(q, wa), lines(ea, wb_), lines(ed, wc), 1) + 4


def WC(r):
    t_ = f'TRIM(SUBSTITUTE(C{r},CHAR(10)," "))'
    return f'(LEN({t_})-LEN(SUBSTITUTE({t_}," ",""))+1)'


def summary(ws, items):
    put(ws, "A2", "Executive summary", F_SECTION, wrap=False)
    for i, it in enumerate(items):
        v_, _ = value(["• "] + flat(it))
        put(ws, f"A{3 + i}", v_, wrap=False)
    put(ws, f"A{3 + len(items)}", TERMS, F_NOTE, wrap=False)


def build(research=False):
    wb = Workbook()
    t1 = wb.active
    t1.title = "Task 1"
    t2, t3, t4, mi = (wb.create_sheet(n) for n in ("Task 2", "Task 3", "Task 4", "Miscellaneous"))
    M, T1, T2 = "Miscellaneous!", "'Task 1'!", "'Task 2'!"
    S_, A_, V_, VF_, ST_, D_, REACH_, RT_ = (MI[k] for k in ("S", "A", "V", "VF", "ST", "D", "REACH", "RT"))
    s = {k: f"{M}${c}${S_}" for k, c in zip("lbh", "BCD")}
    a = {k: f"{M}${c}${A_}" for k, c in zip("lbh", "BCD")}
    v = {k: f"{M}${c}${V_}" for k, c in zip("lbh", "BCD")}
    vf = {k: f"{M}${c}${VF_}" for k, c in zip("lbh", "BCD")}
    st = {k: f"{M}${c}${ST_}" for k, c in zip("lbh", "BCD")}
    DAYS, RATE = f"{M}$C${D_}", f"{M}$C${RT_}"
    RATE_L = f"$C${RT_}"
    G0 = MI["SEG0"]
    SEG, TPD_R, FSIZE = f"{M}$A${G0}:$A${G0 + 1}", f"{M}$D${G0}:$D${G0 + 1}", f"{M}$C${G0}:$C${G0 + 1}"
    P0_ = MI["PROD0"]
    MEME_FEE = f"{M}$D${P0_}"
    N0, TOT = L1["N0"], L1["TOT"]
    N1 = N0 + 7

    # ============================================================ Miscellaneous
    widths(mi, {"A": 50, "B": 28, "C": 28, "D": 28, "E": 70, "F": 18, "G": 18, "H": 30, "I": 46})
    put(mi, "A1", "Miscellaneous: assumptions, products and fees, growth path, sources and resources", F_TITLE, wrap=False)
    summary(mi, [
        "Every forecast number in this workbook comes from the yellow inputs in section A: change one and every tab recalculates.",
        ["Reach is measured (median views, members online); sign-up, trade and swap rates come from three Indonesian referral books; FnF power traders swap ",
         td(f"$C${VF_}"), " a day each."],
        ["Revenue uses the brief's ", tx(f"$C${RT_}", "0.00%"), " net rate; at OKX's published ", tx(f"$D${P0_}", "0.0%"),
         " memecoin fee the same swaps earn ", tx(f"$D${P0_}/$C${RT_}", "0.0"), "x more before partner shares."],
    ])
    put(mi, "A8", "A. Assumptions (yellow cells are inputs)", F_SECTION, wrap=False)
    header(mi, 9, ["Assumption", "Low", "Base", "High", "Why (data)"])
    A = [
        ("Sign-up rate: share of people reached who sign up through the partner's link in 30 days", (0.01, 0.02, 0.04), PCT1,
         "The research sheet's rate for Indonesian crypto creators: about 1 in 50 people who see a partner's posts in a month signs up. Low and high halve and double it"),
        ("Activation: share of sign-ups who make a first trade (become FTDs)", (0.061, 0.15, 0.3), PCT1,
         "Low: 7 of 115 referrals traded on @0xSarimin's Trojan dashboard (March 2026). Base: the midpoint. High: with a launch campaign running"),
        ("Swaps per FTD in their first month", (286, 2000, 16400), MONEY,
         "Low: @0xSarimin's 7 active referrals averaged $286. Base: in line with @0xRusdi's ~$3.4M referred on BasedBot. High: @AiSlimx's $547 a day on OKX Wallet Boost"),
        ("Swaps a day per FnF power trader, standard circle", (250, 1000, 2000), MONEY,
         "Base: Bevan's published wallet swapped ~14 SOL a day (about $1,000) from May to July 2026, about 2x @AiSlimx's $547 a day on OKX Wallet. Low: a quarter of that; high: double"),
        ("FnF power traders start trading on the partner's day", (8, 4, 2), None, "Base: onboarded one-to-one in the partner's first 3 days"),
        ("Days in the partner's first month", (30, 30, 30), None, "Each partner's first 30 days live"),
        ("Reach of a typical post, accounts not measured (share of followers)", (0.2, 0.2, 0.2), PCT,
         "Median views ÷ followers across the 11 accounts the research measured: 10% to 44%, median 20%"),
        ("Net revenue rate", (0.0004, 0.0004, 0.0004), PCT2, "The brief's rate. OKX DEX averaged ~0.09% of its volume in October 2026 (DefiLlama)"),
    ]
    for i, (label, vals, fmt, basis) in enumerate(A):
        r = S_ + i
        cell(mi, r, 1, label)
        for j, vv in enumerate(vals):
            cell(mi, r, 2 + j, vv, fmt=fmt, inp=True)
        cell(mi, r, 5, basis)
    assert S_ + len(A) - 1 == RT_

    put(mi, f"A{G0 - 2}", "B. Trade size by segment (turns swaps into trades a day)", F_SECTION, wrap=False)
    header(mi, G0 - 1, ["Segment", "Average trade, FTD", "Average trade, FnF power trader", "Trades a day per FTD (base)",
                        "Trades a day per FnF power trader (standard circle, base)", "Why (data)"])
    for i, (name, rs, ps, basis) in enumerate([(TREN, 25, 120, "A typical memecoin buy is ~0.35 SOL (about $25 at ~$72 a SOL); power traders size up"),
                                               (LP, 50, 150, "Guesstimate: fewer, larger swaps when liquidity providers rebalance")]):
        r = G0 + i
        cell(mi, r, 1, name)
        cell(mi, r, 2, rs, fmt=MONEY, inp=True)
        cell(mi, r, 3, ps, fmt=MONEY, inp=True)
        cell(mi, r, 4, f"=$C${V_}/$C${D_}/B{r}", fmt=NUM1)
        cell(mi, r, 5, f"=$C${VF_}/C{r}", fmt=NUM1)
        cell(mi, r, 6, basis)

    put(mi, f"A{P0_ - 2}", "C. OKX products and fees (published fee vs the brief's 0.04%)", F_SECTION, wrap=False)
    PR = L["Products & revenue"]
    hdr = ["Product or feature", "What it does", "OKX fee (published)", "Fee rate used", "Swaps per active trader a month",
           "Revenue per trader a month (published fee)", "Revenue per trader a month (0.04%)", "Best-fit audience", "Source"]
    header(mi, P0_ - 1, hdr)
    prod_fix = (("MEV protection", "front-running protection"), ("KOL wallets", "creator (KOL) wallets"), ("DEX API (B2B)", "DEX trading interface (API) for businesses"),
                ("CEX beginners (today's play)", "Beginners from centralised exchanges"), ("DeFi / LP users", LP), ("Memecoin trenchers", TREN),
                ("Perps via Hyperliquid (dApp connect)", "Perpetual futures via Hyperliquid (app connect)"), ("Perp DEX traders funding Hyperliquid", "Perpetual-futures traders funding Hyperliquid"),
                ("B2B integrations, not affiliates", "Business integrations, not affiliates"), ("ETFs", "exchange-traded funds"), ("Onchain", "On-chain"),
                ("; confirm today's rate", ""), ("e.g. USDC to Arbitrum for Hyperliquid", "e.g. US-dollar stablecoins to Arbitrum"),
                ("Not a target unless OKX adds perps", "Not a target: fees go to a rival venue"), ("Trade perps", "Trade perpetual futures"), ("API charges no fee of its own", "The interface charges no fee of its own"))
    for i, r in enumerate(range(3, 13)):
        row = P0_ + i
        for c in range(1, 10):
            vv = PR.cell(r, c).value
            if isinstance(vv, str):
                for x_, y_ in prod_fix:
                    vv = vv.replace(x_, y_)
            if c == 5 and r in (3, 6):
                vv = f"=$C${V_}"
            elif c == 6:
                vv = f"=D{row}*E{row}"
            elif c == 7:
                vv = f"=E{row}*{RATE_L}"
            elif c == 8 and r == 9:
                vv = "Traders moving funds between chains"
            fmt = MONEY2 if c in (6, 7) else (PCT2 if c == 4 or (c == 3 and isinstance(vv, (int, float))) else (MONEY if c == 5 else None))
            cc = cell(mi, row, c, vv, fmt=fmt, inp=c in (4, 5) and vv not in (None, "") and not str(vv).startswith("="))
            if c == 9 and str(vv).startswith("http"):
                link(cc, str(vv).split(" ; ")[0])
    DL = MI["DL0"]
    header(mi, DL - 1, ["OKX DEX actuals (DefiLlama, mid-November 2026)", "Value", "Note"])
    dl_notes = ["DefiLlama's aggregator figure; OKX's own swap page shows $7.251B for 30 days", "September 2026: $3.11M",
                "Low because most volume is fee-free trading-interface flow and major tokens; memecoin swaps in the app pay far more"]
    for i, r in enumerate(range(15, 18)):
        row = DL + i
        cell(mi, row, 1, PR.cell(r, 1).value.replace("Fees, Oct 2026", "Fees, October 2026"))
        cell(mi, row, 2, f"=B{DL + 1}/B{DL}" if r == 17 else PR.cell(r, 2).value, fmt=PCT2 if r == 17 else MONEY)
        cell(mi, row, 3, dl_notes[i])

    GR = MI["GROW0"]
    put(mi, f"A{GR - 2}", "D. How the number grows (0.04%, before partner shares; each partner's first month, base case)", F_SECTION, wrap=False)
    header(mi, GR - 1, ["Step", "Swaps", "Revenue (0.04%)", "Why it's realistic"])
    steps = [
        ("1. Audience only: FTDs × $2,000 (no FnF power traders)", f"={T1}M{TOT}", "The research sheet's base case for the same 8"),
        ("2. Plus FnF power traders from day 4: our plan", f"={T1}P{TOT}", "Power traders trade from their partner's day 4, before any post lands"),
        ("3. Double the FnF power traders at the top 3", f"=B{GR + 1}+{T1}N{N0}+{T1}N{N0 + 1}+{T1}N{N0 + 2}",
         "Kevin's circle, Bevan's Masterclass (29 paid seats) and 0xRusdi's CVC network can each name more"),
        ("4. Plus pilot C at its forecast", f"=B{GR + 2}+{T2}$D${L2['VOL']}", "C's growth deal: no fixed fee, its $3,000 paid only once its volume covers it"),
        ("5. $10K a month", f"=10000/{RATE_L}", "C and A at 2x their forecasts and B at forecast (about $22M), plus our 8 with more power traders: a month 4-6 goal"),
    ]
    for i, (aa, b, d) in enumerate(steps):
        row = GR + i
        cell(mi, row, 1, aa)
        cell(mi, row, 2, b, fmt=MONEY)
        cell(mi, row, 3, f"=B{row}*{RATE_L}", fmt=MONEY)
        cell(mi, row, 4, d)

    SR = MI["SRC0"]
    put(mi, f"A{SR - 2}", "E. Sources and resources (every link checked 20-21 November 2026)", F_SECTION, wrap=False)
    header(mi, SR - 1, ["Group", "What it shows", "Link"])
    row = SR
    for grp, items in SOURCES:
        for topic, url in items:
            cell(mi, row, 1, grp)
            cell(mi, row, 2, topic)
            c_ = mcell(mi, row, 3, 5, url)
            link(c_, url)
            row += 1

    # ============================================================ Task 1
    A0, RK0, CHK = L1["A0"], L1["RK0"], L1["CHK"]
    widths(t1, {"A": 30, "B": 32, "C": 30, "D": 36, "E": 40, "F": 32, "G": 40, "H": 34, "I": 12, "J": 36, "K": 30, "L": 50,
                "M": 40, "N": 36, "O": 30, "P": 14, "Q": 20, "R": 14, "S": 12, "T": 14, "U": 14})
    put(t1, "A1", "Task 1: Find partners beyond airdrop communities", F_TITLE, wrap=False)
    top3 = f"SUM(F{N0}:F{N0 + 2})"
    summary(t1, [
        ["8 partners whose audiences already trade on-chain: ", tx(f"E{CHK}"), " of 8 not airdrop-led, ", tx(f"I{CHK}"),
         " partner types, ", tx(f"G{CHK}"), " meetable in person (Kevin Susanto, Bevan and 0xRusdi, all in Jakarta)."],
        ["Top 3 first: Kevin Susanto, Bevan / Narrative Labs and 0xRusdi with the CVC Ventures network bring ", tx(top3), " of the ",
         tx(f"F{TOT}"), " FTDs and ", tm(f"SUM(P{N0}:P{N0 + 2})"), " of the ", tm(f"P{TOT}"), " first-month swaps."],
        ["All 8 in their first month, base case: ", tx(f"F{TOT}"), " FTDs and ", tx(f"H{TOT}"), " FnF power traders, ",
         tm(f"P{TOT}"), " of swaps, ", td(f"R{TOT}"), " of revenue at the brief's 0.04%."],
    ])
    put(t1, f"A{L1['ANS0'] - 2}", "Answers", F_SECTION, wrap=False)
    ans1 = [
        ("Which 8 partners?",
         "Eight Indonesian traders and trading communities whose followers already trade memecoins on-chain, with every field the brief asks for in section A.",
         [[tx(f"E{CHK}"), " of 8 are not airdrop-led (6 needed); ", tx(f"I{CHK}"), " partner types (3 needed); ", tx(f"G{CHK}"), " meetable in person (2 needed)"],
          "Ranked on proof of trading, FTD potential, FnF power traders and rival flow we can move (section B)",
          "Sized one way for all 8: people reached × 2% sign up × 15% trade, plus their FnF power traders (section C)"]),
        ("Top 3, and why approach them first?",
         "Kevin Susanto, Bevan and 0xRusdi: each has public proof that he trades and brings his own circle of power traders, so they deliver the most volume soonest.",
         [["Kevin Susanto: #3 on fomo's 24-hour profit board; ", tx(f"F{N0}"), " FTDs, the most of the 8; he already knows OKX and is meetable in Jakarta"],
          ["Bevan: his published wallet swapped ~1,080 SOL in May-July 2026 (+374 SOL); a paid Masterclass and a 9,238-member Discord; meetable in Jakarta"],
          ["0xRusdi: ~$3.4M of swaps referred on BasedBot ($8,463 earned); one introduction opens the CVC team (Ponyin, Setya), ApexFNF and the Tomket clan: ",
           tx(f"H{N0 + 2}"), " FnF power traders, ", tk(f"P{N0 + 2}"), " of swaps; meetable in Jakarta"]]),
        ("Why not Brandon Salim or GuarEmperor in the top 3?",
         "Both trade and bring reach, but Brandon's celebrity audience may not trade on-chain and GuarEmperor's feed mixes his calls with whitelist and mint posts, so they start in the second wave.",
         ["The top 3 show stronger public proof: a leaderboard rank (Kevin Susanto), a wallet anyone can read (Bevan), a referral book (0xRusdi)"]),
        ("How do we pick the FnF power traders?",
         "Each partner names their most active traders, and we check every wallet on-chain first: trading on most days, at about $1,000 a day.",
         ["How many: what a partner can onboard one-to-one in 3 days: 3 for a solo trader, 5 for a community lead, 15 across 0xRusdi's network",
          ["How big: ", td(vf["b"]), " a day for every power trader (Bevan's own wallet averaged that)"],
          ["Checked weekly: under ", td(vf["l"]), " a day gets a call; a trader who stops is replaced from the partner's list"]]),
        ("How many people does each partner reach?",
         "The people who see a partner's typical post or are active in their community in a month: measured views plus members online, or 20% of followers where views are not public.",
         [["Largest: Kevin Susanto, ", tx(f"C{N0}"), " people; smallest: DegenRex, ", tx(f"C{N0 + 6}"), " (section C gives the reason for each)"],
          "Of those, 2% sign up and 15% of sign-ups trade: the research sheet's base rates (Miscellaneous, section A)"]),
    ]
    answers(t1, L1["ANS0"] - 1, ans1, (2, 4), (5, 8))

    put(t1, f"A{A0 - 2}", "A. The 8 partners: the brief's fields first, then more detail", F_SECTION, wrap=False)
    header(t1, A0 - 1, ["#", "Partner", "Public link", "Audience and market", "Evidence of fit", "Acquisition route", "Proposed partnership",
                        "Main risk", "Priority", "Reason for priority", "Partner type", "Evidence in full (with dates)", "Assumptions to verify",
                        "FnF power traders they bring", "Rival tool their traders use now", "Meetable in person", "Airdrop-led (main proposition)"])
    for i, p in enumerate(partners):
        r = A0 + i
        vals = [i + 1, p["name"], p["link"], p["s_aud"], p["s_evid"], p["s_route"], p["s_prop"], p["s_risk"],
                f'=IF(J{RK0 + i}<=3,"High",IF(J{RK0 + i}<=7,"Medium","Low"))', p["s_prio"].split(": ", 1)[1], p["type"],
                "\n".join("• " + x.strip() for x in p["evidence"].split("; ")), p["assumptions"], p["fnf"], p["vampire"], p["in_person"], p["airdrop"]]
        for c, vv in enumerate(vals, 1):
            cc = cell(t1, r, c, vv)
            if c == 3:
                link(cc, str(vv).split(" ; ")[0])
    A1_ = A0 + 7
    cell(t1, CHK, 1, "Brief checks", bold=True)
    cell(t1, CHK, 2, "Partners")
    cell(t1, CHK, 3, f"=COUNTA(B{A0}:B{A1_})")
    cell(t1, CHK, 4, "Not airdrop-led (6+ needed)")
    cell(t1, CHK, 5, f'=COUNTIF(Q{A0}:Q{A1_},"No*")')
    cell(t1, CHK, 6, "Meetable in person (2+ needed)")
    cell(t1, CHK, 7, f'=COUNTIF(P{A0}:P{A1_},"Yes")')
    cell(t1, CHK, 8, "Partner types (3+ needed)")
    cell(t1, CHK, 9, f"=SUMPRODUCT(1/COUNTIF(K{A0}:K{A1_},K{A0}:K{A1_}))", fmt=NUM0)

    put(t1, f"A{RK0 - 3}", "B. Ranking (1-5 per criterion; yellow = inputs)", F_SECTION, wrap=False)
    header(t1, RK0 - 2, ["Partner", "Proof of real trading", "FTD potential", "FnF power-trader access", "Rival flow to move",
                         "Reachability", "Segment value", "Risk (5 = low)", "Weighted score", "Rank"])
    cell(t1, RK0 - 1, 1, "Weight", bold=True)
    for c, w in enumerate(WEIGHTS, 2):
        cell(t1, RK0 - 1, c, w, fmt=PCT, inp=True)
    cell(t1, RK0 - 1, 9, f"=SUM(B{RK0 - 1}:H{RK0 - 1})", fmt=PCT)
    for i, p in enumerate(partners):
        r = RK0 + i
        cell(t1, r, 1, f"=B{A0 + i}")
        for c, sc in enumerate(p["scores"], 2):
            cell(t1, r, c, sc, inp=True)
        cell(t1, r, 9, f"=SUMPRODUCT(B{r}:H{r},$B${RK0 - 1}:$H${RK0 - 1})", fmt="0.00")
        cell(t1, r, 10, f"=RANK(I{r},$I${RK0}:$I${RK0 + 7})")

    put(t1, f"A{N0 - 2}", "C. Each partner's first month: FTDs = people reached × sign-up × activation; swaps = FTDs × swaps per FTD + FnF power traders × swaps a day × days trading", F_SECTION, wrap=False)
    header(t1, N0 - 1, ["Partner", "Segment", "People reached a month", "Why this many people", "FTDs (low)", "FTDs (base)", "FTDs (high)",
                        "FnF power traders", "Trading type", "FnF swaps a day each (base)", "Why these FnF power traders",
                        "Trades a day (base, once all are trading)", "Swaps from FTDs (base)", "Swaps from FnF power traders (base)",
                        "Swaps (low)", "Swaps (base)", "Swaps (high)", "Revenue (0.04%, base)", "Partner share", "Paid to partner", "OKX keeps"])
    for i, p in enumerate(partners):
        r = N0 + i
        seg = f"MATCH($B{r},{SEG},0)"
        live = lambda k: f"({DAYS}-{st[k]}+1)"
        cell(t1, r, 1, f"=B{A0 + i}")
        cell(t1, r, 2, p["segment"], inp=True)
        cell(t1, r, 3, p["pool"], fmt=NUM0, inp=True)
        cell(t1, r, 4, p["why_reach"])
        cell(t1, r, 5, f"=C{r}*{s['l']}*{a['l']}", fmt=NUM0)
        cell(t1, r, 6, f"=C{r}*{s['b']}*{a['b']}", fmt=NUM0)
        cell(t1, r, 7, f"=C{r}*{s['h']}*{a['h']}", fmt=NUM0)
        cell(t1, r, 8, p["ptraders"], inp=True)
        cell(t1, r, 9, p["trading_type"])
        cell(t1, r, 10, f"={vf['b']}", fmt=MONEY)
        cell(t1, r, 11, p["why_fnf"])
        cell(t1, r, 12, f"=F{r}*INDEX({TPD_R},{seg})+H{r}*J{r}/INDEX({FSIZE},{seg})", fmt=NUM0)
        cell(t1, r, 13, f"=F{r}*{v['b']}", fmt=MONEY)
        cell(t1, r, 14, f"=H{r}*J{r}*{live('b')}", fmt=MONEY)
        cell(t1, r, 15, f"=E{r}*{v['l']}+H{r}*{vf['l']}*{live('l')}", fmt=MONEY)
        cell(t1, r, 16, f"=M{r}+N{r}", fmt=MONEY)
        cell(t1, r, 17, f"=G{r}*{v['h']}+H{r}*{vf['h']}*{live('h')}", fmt=MONEY)
        cell(t1, r, 18, f"=P{r}*{RATE}", fmt=MONEY2)
        cell(t1, r, 19, p["share"], fmt=PCT, inp=True)
        cell(t1, r, 20, f"=R{r}*S{r}", fmt=MONEY2)
        cell(t1, r, 21, f"=R{r}-T{r}", fmt=MONEY2)
    cell(t1, TOT, 1, "Total, 8 partners", bold=True)
    for c in [3, 5, 6, 7, 8, 12, 13, 14, 15, 16, 17, 18, 20, 21]:
        L_ = col(c)
        cell(t1, TOT, c, f"=SUM({L_}{N0}:{L_}{N1})", fmt=t1.cell(N0, c).number_format, bold=True)
    SRC1 = L1["SRC"]
    put(t1, f"A{SRC1}", "Source:", F_BOLD, wrap=False)
    for i, b in enumerate([
            "X: partner posts, profit-and-loss and referral screenshots, public wallet lists; follower counts and median views through the X API (20 November 2026)",
            "Solscan: Bevan's published wallet (swaps, volume and profit, May-July 2026)",
            "Discord invite counts and Telegram channel previews: community size and members online",
            "Event organiser posts: Web3 Week Asia, Clash of Trenchers and Degen Night",
            "Rival programmes: Bitget Wallet ambassador lists and the BasedBot, Trojan, fomo and pump.fun referral pages",
            "Screened out: Hyperliquid Indonesia (a rival venue) and Bael calls (a Bitget Wallet ambassador)",
            "Every link: Miscellaneous, section E"]):
        put(t1, f"A{SRC1 + 1 + i}", "• " + b, wrap=False)

    # ============================================================ Task 2
    IK = {k: L2[k] for k in ("VOL", "FEE", "ASK", "BR1", "BR2", "BR3", "S1", "S2", "S3", "S4", "MS1", "MS2", "MS3", "MS4",
                             "PY1", "PY2", "PY3", "PY4", "ORDER")}
    SC0, TOP0, MINSH = L2["SC0"], L2["TOP0"], L2["MINSH"]
    VOL, BASE = IK["VOL"], IK["S1"]
    widths(t2, {"A": 34, **{col(c): 15 for c in range(2, 18)}})
    put(t2, "A1", "Task 2: Make a commercial decision", F_TITLE, wrap=False)
    PC = {"A": "B", "B": "C", "C": "D"}
    SC = {}
    rr_ = SC0
    for pn in "ABC":
        for cs in CASES:
            SC[(pn, cs)] = rr_
            rr_ += 1
    r = lambda pn, j: SC[(pn, CASES[j])]
    inp = lambda key, pn: f"${PC[pn]}${IK[key]}"
    summary(t2, [
        ["Pilot C first, once its volume is confirmed: on our deal OKX makes ", td(f"J{r('C', 1)}"), " at forecast instead of losing ",
         td(f"-F{r('C', 1)}"), " on C's own ask."],
        ["Every deal pays a smaller share that steps up with volume, plus milestone payouts, so OKX keeps more than half of the revenue in every case (lowest: ",
         pc(f"B{MINSH}"), ")."],
        ["B gets the steepest steps: ", pc(inp("S1", "B")), " at its forecast, rising to ", pc(inp("S4", "B")),
         " above ", tm(inp("BR3", "B")), ", so pushing volume pays."],
    ])

    def miles(pn):
        out = []
        for k in "1234":
            out += [td(inp("PY" + k, pn)), " at ", tm(inp("MS" + k, pn)), ", " if k != "4" else ""]
        return out

    def steps(pn):
        return [pc(inp("S1", pn)), " to ", tm(inp("BR1", pn)), ", ", pc(inp("S2", pn)), " to ", tm(inp("BR2", pn)), ", ",
                pc(inp("S3", pn)), " to ", tm(inp("BR3", pn)), ", ", pc(inp("S4", pn)), " above"]

    def deal(pn):
        same = f"AND({inp('S1', pn)}={inp('S2', pn)},{inp('S2', pn)}={inp('S3', pn)},{inp('S3', pn)}={inp('S4', pn)})"
        return [X(f'IF({same},TEXT({inp("S1", pn)},"0%")&" share",' + fx(*steps(pn))[1:] + ')'), "; milestone payouts: "] + miles(pn)

    who = {"A": "A · airdrop community", "B": "B · repeat on-chain traders", "C": "C · trading educator"}
    ans2 = [
        ("1. First-month contribution (at forecast)",
         ["At their own asks only B makes money; on our deals all three do, and OKX keeps ", pc(f"MIN(K{r('A', 1)},K{r('B', 1)},K{r('C', 1)})"),
          " or more of the revenue."],
         {pn: ["Their ask: ", td(f"D{r(pn, 1)}"), X(f'IF({inp("FEE", pn)}>0," − "&TEXT({inp("FEE", pn)},"$#,##0")&" fee","")'), " − ", td(f"D{r(pn, 1)}*{inp('ASK', pn)}"), " share = ",
               td(f"F{r(pn, 1)}"), ". Our deal: ", td(f"D{r(pn, 1)}"), " − ", td(f"G{r(pn, 1)}"), " share − ", td(f"H{r(pn, 1)}"),
               " milestone = ", td(f"J{r(pn, 1)}")] for pn in "ABC"}),
        ("2. Preferred pilot",
         ["C, once its 90-day dashboards confirm the ", tm(inp("VOL", "C")), ": the highest volume of the three."],
         {"A": "Later, with a campaign that turns its airdrop users into traders",
          "B": "The fallback if C's numbers don't hold: positive at any volume, but small",
          "C": "Pilot first: its members trade, and it brings the most volume once confirmed"}),
        ("3. Counter-offer: the growth deal",
         "A smaller share that steps up with volume, plus a payout each time monthly volume passes a milestone; no fixed fee.",
         {pn: deal(pn) for pn in "ABC"}),
        ("Why bring more volume",
         "Every milestone pays more per dollar of volume, so the partner earns far more by pushing past its forecast.",
         {pn: ["Earns ", td(f"I{r(pn, 1)}"), " at forecast, ", td(f"I{r(pn, 2)}"), " at $10M and ", td(f"I{r(pn, 3)}"),
               " at $20M; OKX keeps ", pc(f"K{r(pn, 3)}"), " at $20M"] for pn in "ABC"}),
        ("4. Evidence before signing",
         "90 days of real numbers: referral or bot dashboards, community analytics, and the wallets of their most active members, checked on-chain.",
         {"A": "Share of wallets that trade, not just claim airdrops; retention after past campaigns",
          "B": "90 days of referral dashboards and its top wallets' trading history",
          "C": ["The dashboards behind its ", td(inp("VOL", "C")), " monthly estimate, past partnership results and a reference"]}),
        ("5. If volume reaches only 50% of forecast",
         ["C at ", tm(f"C{r('C', 0)}"), " still makes OKX ", td(f"J{r('C', 0)}"), " on our deal, against a ", td(f"-F{r('C', 0)}"),
          " loss on its own ask: no milestone is reached, so nothing extra is paid."],
         {pn: ["Our deal: ", td(f"J{r(pn, 0)}"), " for OKX (its ask: ", td(f"F{r(pn, 0)}"), ")"] for pn in "ABC"}),
        ("Other costs, and what we do about them",
         "Month 1 carries no other costs; after the pilot we add only two kinds, each paid for by the volume it brings.",
         {"ALL": ["Campaigns: a trading competition or live session for each partner past its first milestone, capped at 10% of the revenue above it. "
                  "Fee discounts: a lower swap fee for the named FnF power traders each partner onboards, and a lower fee for a partner's whole "
                  "community once it passes a monthly volume level (for example $10M)"]}),
    ]
    a0 = L2["ANS0"]
    spans = {"Q": (1, 1), "S": (2, 5), "A": (6, 9), "B": (10, 13), "C": (14, 17)}
    for key, h in (("Q", "Question"), ("S", "One-sentence answer"), ("A", who["A"]), ("B", who["B"]), ("C", who["C"])):
        mcell(t2, a0 - 1, *spans[key], h, bold=True, fill=FILL_HEAD)
    put(t2, f"A{a0 - 2}", "Answers", F_SECTION, wrap=False)
    for i, (q, s1, per) in enumerate(ans2):
        rw = a0 + i
        mcell(t2, rw, 1, 1, q, bold=True)
        vs, es = value(s1)
        mcell(t2, rw, *spans["S"], vs)
        hts = [lines(q, 34), lines(es, 60)]
        if "ALL" in per:
            va, ea = value(per["ALL"])
            mcell(t2, rw, spans["A"][0], spans["C"][1], va)
            hts.append(lines(ea, 180))
        else:
            for pn in "ABC":
                va, ea = value(per[pn])
                mcell(t2, rw, *spans[pn], va)
                hts.append(lines(ea, 60))
        t2.row_dimensions[rw].height = 15 * max(hts) + 4
    assert a0 + len(ans2) - 1 == a0 + L2["NANS"] - 1

    put(t2, f"A{VOL - 2}", "A. Their asks and our growth deal (yellow cells are inputs)", F_SECTION, wrap=False)
    header(t2, VOL - 1, ["Input", "A", "B", "C", "How it works"])
    LT2 = L["Task 2 numbers"]
    vol = [LT2.cell(rr, 3).value for rr in (4, 5, 6)]
    fee = [LT2.cell(rr, 5).value for rr in (4, 5, 6)]
    ask = [LT2.cell(rr, 6).value for rr in (4, 5, 6)]
    rows_in = [
        ("Monthly volume forecast (brief)", vol, MONEY, "Brief: A airdrop community, B repeat on-chain traders, C trading educator (volume unverified)"),
        ("Their ask: fixed fee", fee, MONEY, "Brief"),
        ("Their ask: revenue share", ask, PCT, "Brief"),
        ("Share bracket 1 ends at (monthly swaps)", [4e6, 2e6, 6e6], MONEY, "Brackets are marginal: each rate applies only to the volume inside it"),
        ("Share bracket 2 ends at", [7e6, 5e6, 10e6], MONEY, ""),
        ("Share bracket 3 ends at", [10e6, 10e6, 15e6], MONEY, ""),
        ("Share in bracket 1", [0.15, 0.20, 0.10], PCT, "Below each partner's own ask"),
        ("Share in bracket 2", [0.15, 0.25, 0.10], PCT, "B steps up as it grows"),
        ("Share in bracket 3", [0.15, 0.30, 0.10], PCT, ""),
        ("Share above bracket 3", [0.15, 0.35, 0.10], PCT, "B reaches its own 35% ask only above $10M"),
        ("Milestone 1 (monthly swaps)", [4e6, 2e6, 6e6], MONEY, "A payout in each month the milestone is reached; the highest one reached counts"),
        ("Milestone 2", [7e6, 5e6, 10e6], MONEY, ""),
        ("Milestone 3", [10e6, 10e6, 15e6], MONEY, ""),
        ("Milestone 4", [15e6, 15e6, 20e6], MONEY, ""),
        ("Payout at milestone 1", [300, 100, 500], MONEY, "Sized so OKX keeps more than half of the revenue at every milestone"),
        ("Payout at milestone 2", [700, 250, 1000], MONEY, ""),
        ("Payout at milestone 3", [1200, 500, 2000], MONEY, ""),
        ("Payout at milestone 4", [2000, 800, 3000], MONEY, "C's $3,000 ask, paid once it brings $20M in a month"),
        ("Pilot order", [3, 2, 1], None, "C first once its volume is confirmed; B as the fallback; A later with a campaign"),
    ]
    for i, (lab, vals, fmt, note) in enumerate(rows_in):
        rw = VOL + i
        cell(t2, rw, 1, lab, bold=(i == 0))
        for j, vv in enumerate(vals):
            cell(t2, rw, 2 + j, vv, fmt=fmt, inp=True)
        cell(t2, rw, 5, note)
    assert VOL + len(rows_in) - 1 == IK["ORDER"]

    put(t2, f"A{SC0 - 2}", "B. Month 1 at four volumes: their ask against our growth deal", F_SECTION, wrap=False)
    header(t2, SC0 - 1, ["Partner", "Case", "Volume", "Revenue (0.04%)", "Partner earns on its ask", "Month 1 for OKX on its ask",
                         "Our deal: share paid", "Our deal: milestone payout", "Partner earns on our deal", "Month 1 for OKX on our deal",
                         "OKX's share of revenue (our deal)"])
    for pn in "ABC":
        X_ = PC[pn]
        F = f"${X_}${VOL}"
        g = lambda key: f"${X_}${IK[key]}"
        vols = [f"={F}*0.5", f"={F}", "=10000000", "=20000000"]
        for j, cs in enumerate(CASES):
            rw = SC[(pn, cs)]
            cell(t2, rw, 1, pn, bold=True)
            cell(t2, rw, 2, cs)
            cell(t2, rw, 3, vols[j], fmt=MONEY)
            cell(t2, rw, 4, f"=C{rw}*{RATE}", fmt=MONEY)
            cell(t2, rw, 5, f"={g('FEE')}+D{rw}*{g('ASK')}", fmt=MONEY)
            cell(t2, rw, 6, f"=D{rw}-E{rw}", fmt=MONEY)
            v_ = f"C{rw}"
            cell(t2, rw, 7, f"={RATE}*({g('S1')}*MIN({v_},{g('BR1')})+{g('S2')}*MIN(MAX({v_}-{g('BR1')},0),{g('BR2')}-{g('BR1')})"
                            f"+{g('S3')}*MIN(MAX({v_}-{g('BR2')},0),{g('BR3')}-{g('BR2')})+{g('S4')}*MAX({v_}-{g('BR3')},0))", fmt=MONEY)
            cell(t2, rw, 8, f"=IF({v_}>={g('MS4')},{g('PY4')},IF({v_}>={g('MS3')},{g('PY3')},IF({v_}>={g('MS2')},{g('PY2')},IF({v_}>={g('MS1')},{g('PY1')},0))))", fmt=MONEY)
            cell(t2, rw, 9, f"=G{rw}+H{rw}", fmt=MONEY)
            cell(t2, rw, 10, f"=D{rw}-I{rw}", fmt=MONEY)
            cell(t2, rw, 11, f"=J{rw}/D{rw}", fmt=PCT)
    cell(t2, MINSH, 1, "Lowest OKX share of revenue, any partner, any case", bold=True)
    cell(t2, MINSH, 2, f"=MIN(K{SC0}:K{SC0 + 11})", fmt=PCT)
    mcell(t2, MINSH, 3, 11, "Must stay above 50%: the milestone payouts are sized to keep it there")

    put(t2, f"A{TOP0 - 2}", "C. Our real top 3 on simple revenue share (each partner's first month, base case, 0.04%)", F_SECTION, wrap=False)
    header(t2, TOP0 - 1, ["Partner", "Swaps, first month (base)", "Revenue (0.04%)", "Revenue share", "Month-1 contribution",
                          "At 50% of forecast", "Partner earns", "OKX keeps at the published 0.5% fee", "Evidence before signing"])
    for kk in range(3):
        rw, src = TOP0 + kk, N0 + kk
        cell(t2, rw, 1, f"={T1}A{src}")
        cell(t2, rw, 2, f"={T1}P{src}", fmt=MONEY)
        cell(t2, rw, 3, f"=B{rw}*{RATE}", fmt=MONEY2)
        cell(t2, rw, 4, f"={T1}S{src}", fmt=PCT)
        cell(t2, rw, 5, f"=C{rw}*(1-D{rw})", fmt=MONEY2)
        cell(t2, rw, 6, f"=C{rw}/2*(1-D{rw})", fmt=MONEY2)
        cell(t2, rw, 7, f"=C{rw}*D{rw}", fmt=MONEY2)
        cell(t2, rw, 8, f"=B{rw}*{MEME_FEE}*(1-D{rw})", fmt=MONEY2)
        mcell(t2, rw, 9, 17, partners[kk]["check"])
        t2.row_dimensions[rw].height = 32
    cell(t2, TOP0 + 3, 1, "Top 3 total", bold=True)
    for cc in "BCEFGH":
        cell(t2, TOP0 + 3, "ABCDEFGHI".index(cc) + 1, f"=SUM({cc}{TOP0}:{cc}{TOP0 + 2})", fmt=t2[f"{cc}{TOP0}"].number_format, bold=True)

    # ============================================================ Task 3
    widths(t3, {"A": 34, "B": 56, "C": 110, "D": 12})
    put(t3, "A1", "Task 3: Show how you would win and activate a partner", F_TITLE, wrap=False)
    OF0, E1, ET, KT = L3["OFFER0"], L3["EARN1"], L3["EARNT"], L3["KEEPT"]
    summary(t3, [
        "Partner: Kevin Susanto, our #1, with the most FTDs of the 8 and the fastest yes: he already knows OKX.",
        ["Target: ", tm(f"B{OF0 + 3}"), " of monthly swaps through his link, worth ", td(f"B{OF0 + 2}"),
         " of fees at OKX's 0.5% memecoin fee; from there his own VIP-style tiers start."],
        [f"By his day 14: about {_f14:.0f} FTDs and {_t14:.0f} trades a day; at the target he earns ", td(f"B{ET}"),
         " a month and OKX keeps ", td(f"B{KT}"), "."],
    ])
    I0 = L3["ITEMS0"]
    put(t3, f"A{I0 - 2}", "Answers", F_SECTION, wrap=False)
    header(t3, I0 - 1, ["Item", "In one sentence", "Text"] + (["Words"] if research else []))
    for i, (aa, one, b) in enumerate(TASK3):
        r = I0 + i
        cell(t3, r, 1, aa, bold=True)
        cell(t3, r, 2, one)
        cell(t3, r, 3, b)
        if research and i <= 3:
            cell(t3, r, 4, "=" + WC(r))
        t3.row_dimensions[r].height = 15 * max(lines(b, 110), lines(one, 56)) + 4
    t3[f"A{I0}"].value = '="First message (X direct message), "&' + WC(I0) + '&" words"'
    PL0 = L3["PLAN0"]
    put(t3, f"A{PL0 - 3}", "14-day activation plan (day 1 = agreement signed)", F_SECTION, wrap=False)
    put(t3, f"A{L3['PLAN1']}", "In one sentence: days 1-4 set up the link and onboard his power traders one-to-one, day 5 launches the campaign, and days 6-14 push the flow and fix the weakest funnel step.", wrap=False)
    header(t3, PL0 - 1, ["Days", "Target", "What happens"])
    for i, (aa, b, c) in enumerate(PLAN14):
        cell(t3, PL0 + i, 1, aa, bold=True)
        cell(t3, PL0 + i, 2, c)
        cell(t3, PL0 + i, 3, b)
        t3.row_dimensions[PL0 + i].height = 15 * lines(b, 110) + 4
    put(t3, f"A{OF0 - 2}", "Offer to Kevin Susanto: OKX DEX's standard referral share now, his own VIP-style tiers once his link earns OKX $10,000 of fees in a month", F_SECTION, wrap=False)
    header(t3, OF0 - 1, ["Term", "Value", "Note"])
    p0 = partners[0]
    P_K = f"{T1}P{N0}"
    o = {k: f"B{OF0 + j}" for j, k in enumerate(("ret", "std", "fees", "sw1", "s1", "v2", "s2", "v3", "s3"))}
    share = lambda vol: f"IF({vol}>={o['v3']},{o['s3']},IF({vol}>={o['v2']},{o['s2']},IF({vol}>={o['sw1']},{o['s1']},{o['std']})))"
    offer = [("Retainer", "None", None, False, "Nothing fixed: he earns on what his followers trade"),
             ("Standard share until his tiers start", 0.20, PCT, True, "OKX DEX's default referral tier; the published table rises to 30% at $1M of monthly referee volume"),
             ("His tiers start at fees in a month", 10000, MONEY, True, "Like OKX's VIP tiers: a personal structure once his link proves itself"),
             ("Monthly swaps that earn that, at the published fee", f"=B{OF0 + 2}/{MEME_FEE}", MONEY, False, "Memecoin swaps at OKX DEX's published 0.5% fee"),
             ("Tier 1 share (from that volume)", 0.40, PCT, True, "Above the 30-35% the standard table pays at that volume"),
             ("Tier 2: monthly swaps", 5000000, MONEY, True, ""),
             ("Tier 2 share", 0.45, PCT, True, ""),
             ("Tier 3: monthly swaps", 10000000, MONEY, True, ""),
             ("Tier 3 share", 0.50, PCT, True, "OKX DEX's top published tier, which the standard table reaches only at $20M"),
             ("Launch campaign", "Co-hosted live session and a 7-day trading competition for his Telegram", None, False, "Prize pool from the campaign budget"),
             ("Follower discount", "Up to 10%, from his share", None, False, ""),
             ("Swap fee his followers pay", "0.5% on memecoins (published OKX DEX fee)", None, False, ""),
             ("What he needs most", p0["need"], None, False, ""),
             ("Support from us", p0["support"], None, False, ""),
             ("He earns in month one (base case)", f"={MEME_FEE}*{P_K}*{share(P_K)}", MONEY2, False, "His first-month swaps from Task 1 at the share they qualify for"),
             ("He earns at the target", f"={MEME_FEE}*{o['sw1']}*{share(o['sw1'])}", MONEY2, False, "A month at the target volume"),
             ("OKX keeps at the target", f"=B{OF0 + 2}-B{ET}", MONEY2, False, "Fees at the target minus his share"),
             ("Rules for all partners", "Paid monthly; 30-day clawback for wash trading, self-referral and clusters of fake accounts (sybils); no fixed fees until 90 days of data support them; every rate sits inside OKX DEX's referral tiers (20-50%) and is approved by OKX", None, False, "")]
    assert OF0 + 14 == E1 and OF0 + 15 == ET and OF0 + 16 == KT
    for i, (aa, bb, fmt, inp, note) in enumerate(offer):
        cell(t3, OF0 + i, 1, aa, bold=True)
        cell(t3, OF0 + i, 2, bb, fmt=fmt, inp=inp)
        cell(t3, OF0 + i, 3, note)
    SU0 = L3["SUP0"]
    put(t3, f"A{SU0 - 1}", "Support every partner gets", F_SECTION, wrap=False)
    for i, r in enumerate(range(26, 32)):
        cell(t3, SU0 + i, 1, f"{i + 1}.")
        cell(t3, SU0 + i, 2, OF.cell(r, 1).value.replace("POJK 6/2026 finfluencer rules", "finfluencer rules (OJK rule 6/2026)"))

    # ============================================================ Task 4
    R0, RT4, PP0 = L4["R0"], L4["RTOT"], L4["P0"]
    R1 = R0 + 7
    W = {k: L4[k] for k in ("FOCUS", "W0", "LIVE", "ACT", "FTD", "TPD", "SW8", "REV", "REVN", "RUN", "PIL0")}
    PIL = W["PIL0"]
    widths(t4, {"A": 34, "B": 24, "C": 24, "D": 24, "E": 24, "F": 14, "G": 46, **{col(c): 13 for c in range(8, 23)}})
    put(t4, "A1", "Task 4: Demonstrate practical execution and persistence", F_TITLE, wrap=False)
    summary(t4, [
        ["By day 30: ", tx(f"E{W['LIVE']}"), " partners live, ", tx(f"E{W['ACT']}"), " activated, about ", tx(f"E{W['FTD']}"),
         " FTDs, ", tx(f"E{W['TPD']}"), " trades a day and ", tk(f"E{W['SW8']}"), " of swaps from our 8 partners."],
        "Week 1 signs the first partners; weeks 2-3 are the major push (competitions, live sessions, every partner live by day 20); week 4 reviews and makes a last push. Day-by-day plan: the Task 4 Gantt tab.",
        ["Pilot C runs as its own track: evidence checked in days 1-9, live from day 10, ", tm(f"E{PIL + 1}"), " of swaps by day 30."],
    ])
    put(t4, f"A{L4['ANS0'] - 2}", "Answers", F_SECTION, wrap=False)
    ans4 = [
        ("Weekly targets: outreach, meetings, pilots and activated partners",
         [tx(f"F{W['W0'] + 1}"), " messages, about ", tx(f"F{W['W0'] + 2}"), " replies and ", tx(f"F{W['W0'] + 3}"), " meetings in 30 days, for ",
          tx(f"E{W['LIVE']}"), " partners live and ", tx(f"E{W['ACT']}"), " activated by day 30; pilot C is its own track (section B)."],
         ["Week 1: prospect, pitch and sign the first partners; set up links, tracking and the first campaigns",
          "Weeks 2-3, the major push: partners go live as they sign, each with a trading competition and a live session; power traders onboarded one-to-one",
          "Week 4, review and last push: second live sessions, competition finals, a decision per partner on day 30"]),
        ("One in-person prospecting day",
         "Jakarta, once the first partners sign: meet Kevin Susanto, Bevan and 0xRusdi, onboard their power traders in person, then a community trading event that evening.",
         ["Where: Jakarta, home to all three meetable partners",
          "Whom: Kevin Susanto, Bevan, 0xRusdi and their power traders",
          "Access: warm introductions a week ahead (colleagues; a Narrative Labs session)",
          "Fallback: a same-day video call; onboarding continues one-to-one online"]),
        ("20 messages, 2 replies, 0 meetings: what next?",
         "Stop cold messages and go through connections of connections: ask live partners and their circles for direct introductions to the decision-maker, with a clear offer up front.",
         ["Ask Kevin, Bevan and 0xRusdi's circle for introductions",
          "Lead with the offer: a share of every swap and VIP-style tiers as volume grows, nothing to negotiate up front",
          "Follow up with a booking link; test two versions on the next 10"]),
        ("Continue, renegotiate or stop at day 30",
         "Continue if activated with clean wallets, renegotiate if the power traders trade but the audience doesn't, stop if under 5 FTDs or too many flagged wallets.",
         ["Continue: activated (section C) and under 5% flagged wallets",
          "Renegotiate: the FnF power traders trade, but under 10 FTDs or 50 trades a day",
          "Stop: under 5 FTDs, or over 10% flagged wallets"]),
        ("Activated partner",
         "At least 10 FTDs and 50 trades a day within 30 days live, with at least 30% of FTDs still trading after 14 days and under 5% flagged wallets.",
         ["Flagged wallets: wash trades, self-referral and clusters of fake accounts (sybils)",
          "The ramp (section C) checks the first two each week; retention and flags come from the referral dashboard"]),
        ("How do the numbers land inside 30 days?",
         "Sign the first partners in week 1 through warm routes, so their power traders trade within days of go-live and bring volume before the audience arrives.",
         ["FnF power traders start on their partner's day 4; posted FTDs build through the rest of the month",
          "Weeks 2-3 are the major push: a competition and a live session with every partner as it goes live",
          ["Week 4 reviews and pushes once more; result by day 30: ", tk(f"E{W['SW8']}"), " of swaps and ", td(f"E{W['REV']}"),
           " of revenue at 0.04% from our 8"]]),
    ]
    answers(t4, L4["ANS0"] - 1, ans4, (2, 5), (6, 12))

    put(t4, f"A{W['FOCUS'] - 2}", "A. Weekly targets for our 8 partners (rolling: partners go live as they sign)", F_SECTION, wrap=False)
    header(t4, W["FOCUS"] - 1, ["Target", "Week 1 (days 1-7)", "Week 2 (days 8-14)", "Week 3 (days 15-21)", "Week 4 (days 22-30)", "Day 30", "How it's set"])
    focus = ["Prospect, pitch and sign the first partners; set up links and tracking",
             "Major push: first partners live, power traders onboarded, competitions start",
             "Major push: every partner live by day 20; fix the weakest funnel step",
             "Review and last push: second sessions, finals, day-30 decisions"]
    cell(t4, W["FOCUS"], 1, "Focus", bold=True)
    for k, f_ in enumerate(focus):
        cell(t4, W["FOCUS"], 2 + k, f_)
    cell(t4, W["FOCUS"], 6, "")
    cell(t4, W["FOCUS"], 7, "Weeks 2-3 carry the volume; week 4 turns it into decisions")
    t4.row_dimensions[W["FOCUS"]].height = 76
    w0 = W["W0"]
    gl = f"$C${R0}:$C${R1}"
    act = lambda f_, t_: f"=SUMPRODUCT(({f_}{R0}:{f_}{R1}>=10)*({t_}{R0}:{t_}{R1}>=50))"
    rev_after = lambda c_: f"=SUMPRODUCT({c_}{R0}:{c_}{R1},1-$H${R0}:$H${R1})*{RATE}"
    weekly = [
        ("Prospects scored", 25, 20, 15, 10, f"=SUM(B{w0}:E{w0})", "Score everyone before messaging"),
        ("Outreach messages (warm introduction first)", 25, 20, 15, 10, f"=SUM(B{w0 + 1}:E{w0 + 1})", "Every scored prospect"),
        ("Replies", 8, 6, 5, 3, f"=SUM(B{w0 + 2}:E{w0 + 2})", "About 30% of messages"),
        ("Meetings (call or in person)", 3, 6, 5, 4, f"=SUM(B{w0 + 3}:E{w0 + 3})", "About 80% of replies, including one in-person day in Jakarta once the first partners sign"),
        ("Partners live (running total)", f'=COUNTIF({gl},"<=7")', f'=COUNTIF({gl},"<=14")', f'=COUNTIF({gl},"<=21")',
         f'=COUNTIF({gl},"<=30")', f"=E{W['LIVE']}", "Each partner goes live once it signs (go-live days in section C)"),
        ("Activated partners (running total)", act("I", "M"), act("J", "N"), act("K", "O"), act("L", "P"), f"=E{W['ACT']}",
         "At least 10 FTDs and 50 trades a day (section C)"),
        ("FTDs (running total)", f"=SUM(I{R0}:I{R1})", f"=SUM(J{R0}:J{R1})", f"=SUM(K{R0}:K{R1})", f"=SUM(L{R0}:L{R1})", f"=E{W['FTD']}",
         "Task 1 FTDs × share of each partner's month live"),
        ("Trades a day (last day of the week)", f"=SUM(M{R0}:M{R1})", f"=SUM(N{R0}:N{R1})", f"=SUM(O{R0}:O{R1})", f"=SUM(P{R0}:P{R1})",
         f"=E{W['TPD']}", "FTDs so far × trades a day each, plus FnF power traders × trades a day each"),
        ("Swaps (running total)", f"=SUM(Q{R0}:Q{R1})", f"=SUM(R{R0}:R{R1})", f"=SUM(S{R0}:S{R1})", f"=SUM(T{R0}:T{R1})", f"=E{W['SW8']}",
         "Calendar days: FTDs arrive through the month, so this trails Task 1's first-month view"),
        ("Revenue at 0.04% (running total)", *[f"={c_}{W['SW8']}*{RATE}" for c_ in "BCDE"], f"=E{W['REV']}", "Swaps × 0.04%"),
        ("Revenue after partner shares (running total)", rev_after("Q"), rev_after("R"), rev_after("S"), rev_after("T"), f"=E{W['REVN']}",
         "Revenue × (1 − each partner's share)"),
        ("Monthly run-rate on day 30", "", "", "", f"=SUM(U{R0}:U{R1})*30", f"=E{W['RUN']}",
         f'="Swaps a day on day 30 × 30, so $"&TEXT(E{W["RUN"]}*{RATE},"#,##0")&" a month at 0.04%"'),
    ]
    assert w0 + len(weekly) - 1 == W["RUN"]
    for i, rowvals in enumerate(weekly):
        r = w0 + i
        for c, vv in enumerate(rowvals, 1):
            fmt = None
            if r in (W["FTD"], W["TPD"], W["LIVE"], W["ACT"]):
                fmt = NUM0
            elif r in (W["SW8"], W["RUN"]):
                fmt = MONEY
            elif r in (W["REV"], W["REVN"]):
                fmt = MONEY2
            cell(t4, r, c, vv, fmt=fmt if 2 <= c <= 6 else None, inp=(r <= w0 + 3 and 2 <= c <= 5), bold=(c == 1))

    put(t4, f"A{PIL - 2}", "B. Pilot C: a separate track (the brief's hypothetical, on the Task 2 growth deal)", F_SECTION, wrap=False)
    header(t4, PIL - 1, ["Pilot C", "Week 1 (days 1-7)", "Week 2 (days 8-14)", "Week 3 (days 15-21)", "Week 4 (days 22-30)", "Day 30", "How it's set"])
    pil = [("Status", "Checking its 90-day dashboards and top wallets", "Live from day 10 once confirmed", "Live; competition with its members",
            "Live; day-30 review against its first milestone", "Live", "B is the fallback if C's numbers don't hold; A comes later with a campaign"),
           ("Swaps (running total)", *[f"={c_}{PP0}" for c_ in "QRST"], f"=E{PIL + 1}", "C's forecast ($6M a month) spread evenly from day 10"),
           ("Revenue at 0.04% (running total)", *[f"={c_}{PIL + 1}*{RATE}" for c_ in "BCDE"], f"=E{PIL + 2}", "Swaps × 0.04%"),
           ("Revenue after its share (running total)", *[f"={c_}{PIL + 2}*(1-$H${PP0})" for c_ in "BCDE"], f"=E{PIL + 3}",
            "Its bracket-1 share; milestone payouts are monthly, see Task 2")]
    for i, rowvals in enumerate(pil):
        r = PIL + i
        for c, vv in enumerate(rowvals, 1):
            cell(t4, r, c, vv, fmt=(MONEY if i == 1 else MONEY2) if (i and 2 <= c <= 6) else None, bold=(c == 1))
    t4.row_dimensions[PIL].height = 46

    put(t4, f"A{R0 - 3}", "C. Calendar ramp behind the weekly targets (links to Task 1; posted FTDs arrive evenly through the partner's month, event FTDs on the night)", F_SECTION, wrap=False)
    put(t4, f"A{R0 - 2}", "Activated partner: at least 10 FTDs and at least 50 trades a day within 30 days live, at least 30% of FTDs still trading after 14 days "
                          "and under 5% flagged wallets. The ramp checks the first two; retention and flags come from the dashboard.", wrap=False)
    header(t4, R0 - 1, ["Partner", "Event night?", "Go-live day", "Live days by day 7", "Live days by day 14", "Live days by day 21", "Live days by day 30",
                        "Partner share", "FTDs by day 7", "FTDs by day 14", "FTDs by day 21", "FTDs by day 30", "Trades a day, day 7",
                        "Trades a day, day 14", "Trades a day, day 21", "Trades a day, day 30", "Swaps by day 7", "Swaps by day 14",
                        "Swaps by day 21", "Swaps by day 30", "Swaps a day on day 30", "Activated by day 30?"])
    for i, p in enumerate(partners):
        r, t = R0 + i, N0 + i
        seg = f"MATCH({T1}$B{t},{SEG},0)"
        tpd_r, fsz = f"INDEX({TPD_R},{seg})", f"INDEX({FSIZE},{seg})"
        ftd, ptn, fday, start, month = f"{T1}$F{t}", f"{T1}$H{t}", f"{T1}$J{t}", st["b"], DAYS
        vday = f"{v['b']}/{month}"
        cell(t4, r, 1, f"={T1}A{t}")
        cell(t4, r, 2, p["event"], inp=True)
        cell(t4, r, 3, p["go_live"], inp=True)
        for k, d in enumerate((7, 14, 21, 30)):
            cell(t4, r, 4 + k, f"=MAX(0,{d}-$C{r})")
        cell(t4, r, 8, f"={T1}S{t}", fmt=PCT)
        for k, lv in enumerate("DEFG"):
            L_ = f"{lv}{r}"
            fcol = "IJKL"[k]
            cell(t4, r, 9 + k, f'=IF({L_}=0,0,IF($B{r}="Yes",{ftd},{ftd}*MIN({L_},{month})/{month}))', fmt=NUM1)
            cell(t4, r, 13 + k, f"=IF({L_}=0,0,{fcol}{r}*{tpd_r}+IF({L_}>={start},{ptn}*{fday}/{fsz},0))", fmt=NUM1)
            cell(t4, r, 17 + k, f'=IF({L_}=0,0,IF($B{r}="Yes",{ftd}*{vday}*{L_},{ftd}*{vday}*{L_}^2/(2*{month}))+{ptn}*{fday}*MAX(0,{L_}-{start}+1))', fmt=MONEY)
        cell(t4, r, 21, f"=IF(G{r}=0,0,L{r}*{vday}+IF(G{r}>={start},{ptn}*{fday},0))", fmt=MONEY)
        cell(t4, r, 22, f'=IF(AND(L{r}>=10,P{r}>=50),"Yes","No")')
    cell(t4, RT4, 1, "Total, our 8", bold=True)
    for c in range(9, 22):
        cell(t4, RT4, c, f"=SUM({col(c)}{R0}:{col(c)}{R1})", fmt=t4.cell(R0, c).number_format, bold=True)
    for i, (nm, src, live) in enumerate((("Pilot C (brief)", "D", 10), ("Pilot B (brief)", "C", 45), ("Pilot A (brief)", "B", 60))):
        r = PP0 + i
        cell(t4, r, 1, nm)
        cell(t4, r, 2, "No")
        cell(t4, r, 3, live, inp=True)
        for k, d in enumerate((7, 14, 21, 30)):
            cell(t4, r, 4 + k, f"=MAX(0,{d}-C{r})")
            cell(t4, r, 17 + k, f"={T2}${src}${VOL}/{DAYS}*{'DEFG'[k]}{r}", fmt=MONEY)
        cell(t4, r, 8, f"={T2}${src}${L2['S1']}", fmt=PCT)
        cell(t4, r, 21, f"=IF(G{r}=0,0,{T2}${src}${VOL}/{DAYS})", fmt=MONEY)

    D0 = L4["DEL0"]
    put(t4, f"A{D0 - 2}", "D. What each partner delivers in its first month", F_SECTION, wrap=False)
    header(t4, D0 - 1, ["Partner", "Deliverables", "FnF power traders", "FTD target (base)", "Trades a day (base, once all are trading)", "Swaps (first month, base)"])
    for i, p in enumerate(partners):
        r, t = D0 + i, N0 + i
        cell(t4, r, 1, f"={T1}A{t}")
        cell(t4, r, 2, p["deliverables"])
        cell(t4, r, 3, f"={T1}H{t}")
        cell(t4, r, 4, f"={T1}F{t}", fmt=NUM0)
        cell(t4, r, 5, f"={T1}L{t}", fmt=NUM0)
        cell(t4, r, 6, f"={T1}P{t}", fmt=MONEY)

    MO = L4["MON0"]
    put(t4, f"A{MO - 2}", "E. What we monitor, and what we change when it slips", F_SECTION, wrap=False)
    header(t4, MO - 1, ["Metric", "Definition", "Source", "Target", "Red flag", "What we change", "How often"])
    mon = [
        ("FTDs", "Referred users making their first trade through the partner's link", "OKX referral dashboard; on-chain", "25% of the 30-day target by day 7, 50% by day 14",
         "Under half that pace for a week", "Follower discount (0-20%), a live wallet-funding session, a pinned Bahasa guide", "Daily"),
        ("Trades a day", "Swaps by referred wallets per day (7-day average)", "OKX referral dashboard", "At least 50 per partner by day 14", "Under 30",
         "Meme Mode presets session; Track alerts room; daily trade threads", "Daily"),
        ("FnF power traders active", "Introduced power traders trading this week", "Dashboard and the partner's list", "3-5 per partner by day 4 (15 across 0xRusdi's network)",
         "Under 3 by day 7", "Partner names more; one-to-one onboarding the same day", "Daily"),
        ("Swaps per FnF power trader", "Swap volume a day per power trader", "OKX referral dashboard", "About $1,000 a day", "Under $250",
         "Check chain and token coverage; send issues to OKX product", "Weekly"),
        ("Retention", "Share of FTDs still trading in days 15-30", "On-chain and dashboard", "At least 30%", "Under 20%", "Weekly trading room; review the partner's share", "Weekly"),
        ("Flagged wallets", "Wash trading, self-referral, clusters of fake accounts (sybils)", "Fraud screening", "Under 5%", "Over 10%", "Pause payouts; 30-day clawback", "Weekly"),
        ("Revenue per business-development hour", "Partner revenue ÷ hours spent on that partner", "Time log and dashboard", "Rank partners", "Bottom 2 for 2 weeks",
         "Move time to the best partner type", "Weekly"),
        ("Outreach reply rate", "Replies ÷ messages", "Pipeline sheet", "At least 30%", "20 messages, 2 replies, 0 meetings",
         "Connections of connections: direct introductions from live partners and their circles, with the offer up front (Answers)", "Weekly"),
        ("Compliance gate", "Partner's finfluencer certification (OJK rule 6/2026) checked; OKX legal sign-off on promotion in Indonesia", "Partner and OKX legal",
         "Before any paid post", "A paid post goes out without it", "Hold paid posts; private one-to-one onboarding continues", "Once per partner"),
    ]
    for i, vals in enumerate(mon):
        for c, vv in enumerate(vals, 1):
            cell(t4, MO + i, c, vv, bold=(c == 1))
    CA = L4["CAD0"]
    put(t4, f"A{CA - 2}", "F. Operating cadence", F_SECTION, wrap=False)
    header(t4, CA - 1, ["When", "What", "Output"])
    cad = [("Daily, 15 minutes", "FTDs, trades a day and FnF power traders for every live partner", "Red metrics flagged with one fix each"),
           ("Monday", "Partner review: scorecard to each partner, one fix agreed", "Weekly scorecards"),
           ("Thursday", "Pipeline review: score new prospects, plan next week's outreach", "Next week's target list"),
           ("As partners sign", "In-person day in Jakarta (Answers): onboard power traders in person, then a community event", "Partners onboarded in person"),
           ("Day 14", "Mid-point: partners under 10 FTDs or 50 trades a day get the power-trader track again or pause", "Time reallocated"),
           ("Day 30", "Decision per partner: continue, renegotiate or stop (Answers); month-2 plan to manager", "Decisions and month-2 plan")]
    for i, (aa, b, c) in enumerate(cad):
        cell(t4, CA + i, 1, aa, bold=True)
        cell(t4, CA + i, 2, b)
        cell(t4, CA + i, 3, c)

    # ============================================================ Task 4 Gantt (day-by-day plan and KPIs)
    from openpyxl.formatting.rule import FormulaRule
    gt = wb.create_sheet("Task 4 Gantt", index=4)
    G0, K0 = GT["G0"], GT["KPI0"]
    T4s = "'Task 4'!"
    widths(gt, {"A": 52, "B": 10, "C": 40, "D": 7, "E": 7, **{col(c): 3.6 for c in range(6, 36)}})
    put(gt, "A1", "Task 4: Gantt chart and KPIs by day", F_TITLE, wrap=False)
    summary(gt, [
        "Our 8 partners: sign the top 3 in week 1, go live as they sign (3 by day 8, 6 by day 14, all 8 by day 20), push in weeks 2-3, review in week 4.",
        "Pilot C is its own track: evidence checked in days 1-9, live from day 10.",
        ["By day 30: ", tx(f"{T4s}E{L4['LIVE']}"), " partners live, ", tx(f"{T4s}E{L4['ACT']}"), " activated, ", tx(f"{T4s}E{L4['FTD']}"),
         " FTDs and ", tk(f"{T4s}E{L4['SW8']}"), " of swaps from our 8; ", tm(f"{T4s}E{L4['PIL0'] + 1}"), " from pilot C."],
    ])
    put(gt, f"A{G0 - 2}", "A. Activities by day (bars recolour when you change a start or end day)", F_SECTION, wrap=False)
    header(gt, G0 - 1, ["Activity", "Track", "KPI", "Start day", "End day"] + [str(d) for d in range(1, 31)])
    for c in range(6, 36):
        gt.cell(G0 - 1, c).value = c - 5
        gt.cell(G0 - 1, c).alignment = Alignment(horizontal="center")
        gt.cell(G0 - 1, c).fill = PatternFill("solid", fgColor=("E2E6EC" if ((c - 6) // 7) % 2 == 0 else "F3F3F3"))
    top3_fnf = f"SUM({T1}H{N0}:H{N0 + 2})"
    kpis = {
        "msgs": [tx(f"{T4s}F{L4['W0'] + 1}"), " messages, ", tx(f"{T4s}F{L4['W0'] + 2}"), " replies"],
        "meet": [tx(f"{T4s}F{L4['W0'] + 3}"), " meetings"],
        "fnf3": [tx(top3_fnf), " power traders trading by day 12"],
        "live6": [tx(f"{T4s}C{L4['LIVE']}"), " partners live by day 14"],
        "live8": [tx(f"{T4s}D{L4['LIVE']}"), " partners live by day 21"],
        "act": [tx(f"{T4s}E{L4['ACT']}"), " partners activated by day 30"],
        "pil": ["On pace for ", tm(f"{T2}$D${L2['VOL']}"), " a month; ", tm(f"{T4s}E{L4['PIL0'] + 1}"), " of swaps by day 30"],
    }
    for i, (name, track, kpi, s0, s1) in enumerate(GANTT):
        rw = G0 + i
        cell(gt, rw, 1, name, bold=False)
        cell(gt, rw, 2, track)
        v_, _ = value(kpis[kpi] if kpi in kpis else kpi)
        cell(gt, rw, 3, v_)
        cell(gt, rw, 4, s0, inp=True)
        cell(gt, rw, 5, s1, inp=True)
        for c in range(6, 36):
            gt.cell(rw, c).border = BOX
        gt.row_dimensions[rw].height = 30
    g1 = G0 + len(GANTT) - 1
    rng = f"F{G0}:AI{g1}"
    gt.conditional_formatting.add(rng, FormulaRule(formula=[f'AND(F${G0 - 1}>=$D{G0},F${G0 - 1}<=$E{G0},$B{G0}="Pilot C")'],
                                                   fill=PatternFill("solid", fgColor="8FB0FF"), stopIfTrue=True))
    gt.conditional_formatting.add(rng, FormulaRule(formula=[f"AND(F${G0 - 1}>=$D{G0},F${G0 - 1}<=$E{G0})"],
                                                   fill=PatternFill("solid", fgColor="2C5BE0")))
    assert g1 < K0 - 2

    put(gt, f"A{K0 - 2}", "B. KPIs by day: our 8 partners, then pilot C (all linked to Task 4)", F_SECTION, wrap=False)
    header(gt, K0 - 1, ["KPI", "Day 7", "Day 14", "Day 21", "Day 30", "Activities that drive it"])
    R0g, R1g = L4["R0"], L4["R0"] + 7
    fnf_on = lambda c_: f"=SUMPRODUCT(({T4s}{c_}{R0g}:{c_}{R1g}>={st['b']})*{T1}H{N0}:H{N0 + 7})"
    krows = [
        ("Partners live", [f"={T4s}{c_}{L4['LIVE']}" for c_ in "BCDE"], NUM0, "Signing the top 3 in week 1; the next 3 and last 2 as they sign"),
        ("FnF power traders trading", [fnf_on(c_) for c_ in "DEFG"], NUM0, "One-to-one onboarding in each partner's first 3 days"),
        ("FTDs (running total)", [f"={T4s}{c_}{L4['FTD']}" for c_ in "BCDE"], NUM0, "Partner posts, live sessions and competitions"),
        ("Trades a day", [f"={T4s}{c_}{L4['TPD']}" for c_ in "BCDE"], NUM0, "Power traders plus FTDs who keep trading"),
        ("Activated partners", [f"={T4s}{c_}{L4['ACT']}" for c_ in "BCDE"], NUM0, "10+ FTDs and 50+ trades a day"),
        ("Swaps (running total)", [f"={T4s}{c_}{L4['SW8']}" for c_ in "BCDE"], MONEY, "All of the above"),
        ("Revenue at 0.04% (running total)", [f"={T4s}{c_}{L4['REV']}" for c_ in "BCDE"], MONEY, "Swaps × 0.04%"),
        ("Pilot C: swaps (running total)", [f"={T4s}{c_}{L4['PIL0'] + 1}" for c_ in "BCDE"], MONEY, "Its own track: evidence check, then live from day 10"),
        ("Pilot C: revenue at 0.04%", [f"={T4s}{c_}{L4['PIL0'] + 2}" for c_ in "BCDE"], MONEY, "Swaps × 0.04%"),
    ]
    for i, (lab, vals, fmt, why) in enumerate(krows):
        rw = K0 + i
        cell(gt, rw, 1, lab, bold=True)
        for j, vv in enumerate(vals):
            cell(gt, rw, 2 + j, vv, fmt=fmt)
        cell(gt, rw, 6, why)
        gt.merge_cells(start_row=rw, start_column=6, end_row=rw, end_column=35)
    gt.freeze_panes = f"F{G0}"

    if research:
        research_tabs(wb, partners)
        checklist_longlist(wb)
    return wb

def research_tabs(wb, partners):
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
        "GuarEmperor": ("In the 8 (#5)", "FnF circle and tools Discord; we add 20% of his ~16.5K Telegram to P (from the written research)"),
        "Brandon": ("In the 8 (#4)", "Second-largest measured reach; posts a trading referral code on almost every trade; founder of a79"),
        "Bevan": ("In the 8 (#2)", "Verified wallet and a live 1-to-10 SOL run, a paid community he founded, meetable in Jakarta"),
        "Nicholas": ("Bench (month 2)", "FTD source with no trading proof; airdrop community"),
        "Orpheu": ("Bench (month 2), as Stress Capital", "Largest owned clan Discord we could measure, but no trading proof yet"),
        "Ponyin": ("In the 8 (#3, inside 0xRusdi's CVC Ventures network)", "Verified Trojan PnL card; brings ApexFNF, his FnF, and opens the CVC team"),
        "Tequila": ("Bench (month 2)", "128K Discord but 1.2% online; NFT audience"),
        "VIBI": ("Bench (month 2)", "Beginner onboarding; no trading proof"),
        "DegenRex": ("In the 8 (#7)", "Verified live win; we leave out the 1.5x boost so every partner uses one formula (P 1,697)"),
        "Luna": ("Screened out", "Grows the Cove bot: rival tool tie"),
        "mage": ("Bench (month 2)", "Moved his main wallet from OKX to Bitget Wallet on 11 Nov: win-back later"),
        "All Good": ("Access channel", "Caster and connector in Kevin, Bevan and Brandon's circle; Emperor Journal member; no group of his own"),
        "Friday7th": ("Bench (month 2)", "Co-hosts Meteora LP Army, now benched for Laba Lab"),
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
        ("Ranking", "Kevin Susanto #1, Bevan #2, 0xRusdi with the CVC network #3, Brandon Salim #4, GuarEmperor #5, as you set it. The weights follow: proof 15%, FTD potential 25%, FnF 20%, rival flow 10%, reachability 15%, segment 5%, risk 10%."),
        ("CVC Ventures network", "One partner row for 0xRusdi (+ Tomket Lovers), Ponyin (+ ApexFNF) and Setya, as you asked. Team roles come from your X check; CVC_Ventures and ApexFNF_io didn't show up in web search. Reach adds four overlapping audiences, so it is an upper bound."),
        ("Tomket Lovers", "On Bitget Wallet Indonesia's Superfans list (CEX KOL map tab): clear non-exclusivity before signing."),
        ("DegenRex's 1.5x boost", "The deep-dive multiplies his reach by 1.5 for the win (P 2,546, base $15K). We leave it out so all 8 use one formula (P 1,697, about $10K before FnF); the win counts in his proof score instead."),
        ("Brandon Salim", "Reach measured (median 14,342 over 25 posts). a79's size is not public, so a79 adds FnF traders (5), not reach."),
        ("Bevan's pool P", "Volume proof counts all 9,238 Discord members (P 17,494, base $105K); the deep-dive counts only the 984 online (P 9,240, base $55K). We use the deep-dive: online members are the ones a post reaches."),
        ("Kevin Susanto's pool P", "Volume proof counts all 2,860 Telegram subscribers (P 44,108); the deep-dive counts 20% (P 41,820). We use the deep-dive."),
        ("GuarEmperor's pool P", "The deep-dive left Telegram out (P 15,885). The written research found ~16.5K on his 'Enjoy being poor' channel, so we add 20% (P 19,185). Snapshot from a third-party viewer: verify on t.me."),
        ("Written research vs the sheet", "The written shortlist sized FTDs from visits x conversion (Bevan 40 FTDs, $112K; Ponyin 30, $144K; Kevin 28, $120K). The sheet uses one formula for every account. We use the sheet so every partner is sized the same way."),
        ("Timing", "The research's volume is a cohort's first month of trading. Task 4's calendar ramp counts calendar days, so FTDs who arrive late in the month trade into month 2 and day-30 swaps trail Task 1."),
        ("FnF power traders", "The research leaves out partners' own trading. We add only the traders each partner introduces, at $1,000 a day standard (Bevan's published wallet averaged that), 2x for Kevin Susanto's whale circle and 1.5x for Ponyin's network; the partner's own trading is evidence, not forecast volume."),
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
        ("Re-open the 8 partner links and re-check follower counts, median views and quoted posts on-platform", "Task 1, section A and section C"),
        ("Kevin Susanto: with OKX's regional team, what okx.ac/join/VVIP7777 already earns, who owns it, and how DEX tracking stays separate", "Task 1, section A; Task 3"),
        ("CVC Ventures network: confirm the team (Setya, Ponyin, 0xRusdi), ApexFNF and Tomket clan sizes, who signs, and Tomket's Bitget Wallet Superfans status", "Task 1, section A"),
        ("Bevan: which wallet he trades from now (the published one is idle since 30 Jul)", "Task 1, section A"),
        ("Brandon Salim: a79 member count and fomo referral results; ask Kevin for the intro", "Task 1, section A"),
        ("Measure median views for 0xRusdi, Setya, Crypto Ndo and MeteoraIDN (reach is estimated at 20% of followers)", "Task 1, section C (why this many people)"),
        ("GuarEmperor: confirm his FnF circle and the ~16.5K Telegram count on t.me", "Task 1, section A and section C"),
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
        ("FnF power traders at $1,000 a day standard, 2x for Kevin Susanto's whale circle, 1.5x for Ponyin's network", "Rounded, and sized by circle; replaces $16,400 a month each"),
        ("Answers on top of each tab (one-sentence answer, then bullets) and an executive summary on every tab", "Easier to read; titles match the brief"),
        ("Task 1 section A in the brief's field order, with short texts; full evidence kept in the next columns", "The brief's table, readable at a glance"),
        ("Rolling go-live: 3 partners on day 8, 3 on day 14, 2 on day 20; pilot B on day 8", "Weeks 2-3 as the major push, week 4 as review and last push"),
        ("Brief checklist and the 100-prospect longlist moved to this research workbook", "Not needed in the submission"),
        ("Write-up delivered as a Word file", "Google Docs dropped"),
        ("Meteora LP Army replaced by Michael Wyann Santoso / Laba Lab (#6)", "More FTDs and volume: a paid trading-education community; liquidity providers swap less often"),
        ("Task 2: smaller step-up shares plus milestone payouts; OKX keeps over 50% of revenue in every case", "Your call: no $3,000 paid on volume that may not arrive"),
        ("Task 4: pilot C split out as its own track; Gantt tab added", "Clearer plan and KPIs by day"),
        ("#3 changed to 0xRusdi + CVC network; Brandon Salim #4, GuarEmperor #5", "Your edited write-up; your new scores put Brandon (4.20) ahead of GuarEmperor (4.10)"),
        ("FnF circle multiplier removed: every power trader at $1,000 a day; Bevan's reach set to 25,000 (guesstimate)", "Your call"),
        ("Task 2: pilot C first on a growth deal (share as asked, fee as a bonus at a volume threshold, 50% above); B on its own terms", "Push the highest volume; keep B simple"),
        ("Task 3: Kevin's personal tiers start at $10,000 of monthly fees ($2M of swaps), VIP-style", "Higher targets, modelled on OKX's VIP and referral tiers"),
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


def checklist_longlist(wb):
    ws = wb.create_sheet("Checklist and longlist")
    put(ws, "A1", "Brief checklist and the 100-prospect longlist (research copy only)", F_TITLE, wrap=False)
    put(ws, "A3", "A. Brief checklist: where each request is answered", F_SECTION, wrap=False)
    checklist = [
        ("Format: up to 6 slides or 4 pages, plus a prospect spreadsheet", "Write-up (4 pages plus a sources appendix), 6 slides, the submission workbook"),
        ("Public information only", "All tabs"),
        ("Task 1: 8 potential partners", "Task 1, section A"),
        ("Task 1: at least 6 whose main proposition is not airdrops", "Task 1, brief checks"),
        ("Task 1: at least 3 partner types", "Task 1, brief checks"),
        ("Task 1: at least 2 you could meet in person", "Task 1, brief checks (Kevin Susanto and Bevan)"),
        ("Task 1: link, audience and market, evidence, route, proposed partnership, main risk, priority with reason", "Task 1, section A, columns B-J"),
        ("Task 1: top 3 and why you'd approach them first", "Task 1, Answers"),
        ("Task 1: verified facts vs assumptions; follower count alone is not evidence", "Task 1, section A (evidence, assumptions to verify) and section C (why this many people)"),
        ("Task 2: first-month contribution for each partner, with the calculation", "Task 2, Answers 1 and section B"),
        ("Task 2: preferred pilot and reasoning", "Task 2, Answers 2"),
        ("Task 2: counter-offer or pilot structure", "Task 2, Answers 3 and section A"),
        ("Task 2: evidence before signing", "Task 2, Answers 4"),
        ("Task 2: chosen partner's economics at 50% of forecast", "Task 2, Answers 5 and section B"),
        ("Task 2: exclude other costs, then say which you'd investigate", "Task 2, Answers"),
        ("Task 3: use your highest-priority real prospect", "Task 3 (Kevin Susanto)"),
        ("Task 3: first outreach message of 120 words or fewer", "Task 3, Answers (word count in the label)"),
        ("Task 3: follow-up after no response", "Task 3, Answers"),
        ("Task 3: reply to 'another platform offers a higher commission'", "Task 3, Answers"),
        ("Task 3: activation plan for the first 14 days after agreement", "Task 3, activation plan"),
        ("Task 3: 3 most important first-meeting questions", "Task 3, Answers"),
        ("Task 3: no unverified product, incentive or commercial promises", "Task 3, Answers"),
        ("Task 4: weekly targets for outreach, meetings, pilots and activated partners", "Task 4, section A"),
        ("Task 4: one in-person prospecting day (where, whom, access, fallback)", "Task 4, Answers"),
        ("Task 4: response to 20 messages, 2 replies, 0 meetings", "Task 4, Answers"),
        ("Task 4: results that mean continue, renegotiate or stop", "Task 4, Answers"),
        ("Task 4: activated-partner definition", "Task 4, Answers and section B"),
        ("Task 4: commitment statement", "Write-up, commitment; slide 6"),
        ("Cite the public sources you use", "Miscellaneous, section E; write-up appendix"),
        ("Explain how you used AI and what you verified yourself", "Write-up, how I worked; slide 6"),
        ("Be ready to adapt when assumptions change", "Miscellaneous, section A and Task 2, section A (yellow inputs)"),
    ]
    header(ws, 4, ["#", "Brief request", "Where it's answered", "Status"])
    for i, (req, where) in enumerate(checklist):
        cell(ws, 5 + i, 1, i + 1)
        cell(ws, 5 + i, 2, req)
        cell(ws, 5 + i, 3, where)
        cell(ws, 5 + i, 4, "Done")
    h0 = 5 + len(checklist) + 1
    # H. longlist
    put(ws, f"A{h0}", "H. Longlist: 100 public Indonesian prospects screened", F_SECTION, wrap=False)
    put(ws, f"A{h0 + 1}", "Each prospect checked against the same filter: proof of trading, FnF power traders it can bring, rival flow to move, trencher fees.", wrap=False)
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
                   "#4: FnF circle; members post results from his tools; 1.9% engagement", "Wave 2 (day 14)",
                   "https://x.com/GuarEmperor/status/2115414023114256560"]
        elif name == "Airdrop Finder (Telegram)":
            rec[7] = "Airdrop-led: fails the brief's main filter. Its co-founder Setya joins through the CVC Ventures network (#3), not this channel"
        elif name == "Hyperliquid Indonesia (community)":
            rec[6], rec[7], rec[8] = "Excluded by rule", "Rival perp venue: OKX can't send traffic there", ""
        elif name == "Superteam Indonesia":
            rec[6], rec[7], rec[8] = "Bench (month 2)", "Meetable weekly in Jakarta, but builder-led: no trading proof yet", "Month 2: co-hosted trader night once a partner proves demand"
        rows.append(rec)
    new = [
        ("Kevin Susanto", "Trading creator", "https://x.com/KevinSusanto", TREN, "Y", "In the 8",
         "#1: most FTDs, a verified top fomo trader, and the fastest yes through OKX's regional team", "Wave 1 (day 8)", "https://x.com/KevinSusanto/status/2123472525070102987"),
        ("Bevan Pramudito / Narrative Labs", "Trading educator", "https://x.com/bevanpramudt", TREN, "Y", "In the 8",
         "#2: verified wallet and a live 1-to-10 SOL run, a paid community he founded, GMGN flow, meetable in Jakarta", "Wave 1 (day 8)", "https://narrativelabs.id"),
        ("Ponyin + CVC Ventures network (ApexFNF, 0xRusdi + Tomket Lovers, Setya)", "Trading community", "https://x.com/CVC_Ventures", TREN, "Y", "In the 8",
         "#3, through 0xRusdi: BasedBot referral book, Ponyin's Trojan card, ApexFNF and the Tomket clan (15 FnF traders)", "Wave 1 (day 8)", "https://x.com/Ponyin"),
        ("Brandon Salim + a79", "Trading creator", "https://x.com/brandonsalim", TREN, "Y", "In the 8",
         "#5: 14K median views, fomo code on almost every trade, founder of the a79 trader group", "Wave 2 (day 14), intro through Kevin", "https://x.com/brandonsalim/status/2096005957692301790"),
        ("Crypto Ndo", "Trading creator", "https://x.com/cryptondo", TREN, "Y", "In the 8",
         "#6: largest memecoin-trader reach on X; trades on OKX Wallet himself", "Wave 2 (day 14)", "https://x.com/cryptondo"),
        ("DegenRex", "Trading creator", "https://x.com/degen2ex", TREN, "Y", "In the 8",
         "#7: won Clash of Trenchers live (1 SOL to 10 SOL); multichain", "Wave 3 (day 20)", "https://x.com/AveaiGlobal/status/2120891736373518591"),
        ("Michael Wyann Santoso / Laba Lab", "Trading educator", "https://laba-lab.com", TREN, "Y", "In the 8",
         "#6: paid trading membership and bootcamps; YouTube 121K, Telegram 10.8K + 6.9K; Jakarta", "Wave 2 (day 14)", "https://laba-lab.com/membership"),
        ("Meteora LP Army", "Onchain trading community", "https://x.com/EvilPanda", LP, "Y", "Bench (month 2)",
         "Largest in-person on-chain crowds (462 and 671), but liquidity providers swap less often", "Month 2: co-hosted night", "https://luma.com/76fvdt2j"),
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
        ("zeraa (@0xzeraa)", "Ecosystem community", "https://x.com/0xzeraa", "NFT holders", "N", "Screened out",
         "Team role at Lil Stars; reach not measured", "", "https://x.com/0xzeraa"),
    ]
    for rec in new:
        rows.append([None, *rec])
    assert len(rows) == 100, len(rows)
    order = {"In the 8": 0, "Bench (month 2)": 1, "Access channel": 2, "Lead to verify": 3, "Excluded by rule": 4, "Legal hold": 5, "Screened out": 6}
    rows.sort(key=lambda x: (order[x[6]], str(x[1]).lower()))
    header(ws, h0 + 3, ["Screen result", "Count"])
    for i, res in enumerate(order):
        cell(ws, h0 + 4 + i, 1, res)
        cell(ws, h0 + 4 + i, 2, f'=COUNTIF($G${h0 + 14}:$G${h0 + 113},A{h0 + 4 + i})')
    cell(ws, h0 + 11, 1, "Total", bold=True)
    cell(ws, h0 + 11, 2, f"=SUM(B{h0 + 4}:B{h0 + 10})", bold=True)
    header(ws, h0 + 13, ["#", "Prospect", "Type", "Public link", "Segment", "Onchain users? (Y/N)", "Screen result", "Why", "Use in month 1-2", "Source"])
    for i, rec in enumerate(rows):
        row = h0 + 14 + i
        rec[0] = i + 1
        for c, vv in enumerate(rec, 1):
            cc = cell(ws, row, c, vv)
            if c in (4, 10) and str(vv).startswith("http"):
                link(cc, vv)
    widths(ws, {"A": 30, "B": 50, "C": 50, "D": 30, "E": 26, "F": 16, "G": 18, "H": 50, "I": 40, "J": 44})


# ================================================================ v6 data used by build()
OF = L["Offers & support"]

SOURCES = [
    ("OKX product and fees", [
        ("OKX DEX fee schedule", "https://web3.okx.com/dex-fees"),
        ("OKX DEX interface-fee notice (effective 12 May 2026)", "https://web3.okx.com/help/update-on-okx-dex-interface-fee"),
        ("OKX DEX referral programme", "https://web3.okx.com/help/whats-dex-referral-program"),
        ("OKX DEX referral FAQ (tiers, invitee discount)", "https://www.okx.com/help/dex-referral-faq"),
        ("OKX VIP tiers (rate adjustment FAQ)", "https://www.okx.com/help/vip-rate-adjustment-faq"),
        ("OKX Wallet Meme Mode", "https://web3.okx.com/learn/okx-meme-mode"),
        ("OKX Boost and X Launch", "https://web3.okx.com/help/okx-wallet-announcement-okx-boosts-x-launch-project-zama-zama"),
        ("OKX tokenized stocks", "https://www.okx.com/en-us/help/tokenized-stocks-faq"),
        ("OKX DEX trading interface (API) fees", "https://web3.okx.com/build/docs/waas/dex-api-addfee"),
        ("OKX DEX volume and fees (DefiLlama)", "https://defillama.com/protocol/okx-dex"),
        ("OKX Alpha Traders Summit, Bali (BitPinas)", "https://bitpinas.com/feature/okx-alpha-traders-summit/"),
    ]),
    ("Partners", [
        ("Kevin Susanto on X", "https://x.com/KevinSusanto"),
        ("Kevin Susanto's Telegram channel", "https://t.me/KevinSusantoAlpha"),
        ("Bevan Pramudito on X", "https://x.com/bevanpramudt"),
        ("Narrative Labs (Masterclass, Discord, meetups)", "https://narrativelabs.id"),
        ("Ponyin on X", "https://x.com/Ponyin"),
        ("CVC Ventures team on X", "https://x.com/CVC_Ventures"),
        ("ApexFNF, Ponyin's friends-and-family group", "https://x.com/ApexFNF_io"),
        ("0xRusdi on X", "https://x.com/0xRusdi"),
        ("Setya on X", "https://x.com/setyamickala"),
        ("Tomket Lovers on X", "https://x.com/TomketLovers"),
        ("GuarEmperor on X", "https://x.com/GuarEmperor"),
        ("Emperor Journal on X", "https://x.com/emperorjournal_"),
        ("Brandon Salim on X", "https://x.com/brandonsalim"),
        ("a79, Brandon Salim's trader group", "https://x.com/a79"),
        ("Crypto Ndo on X", "https://x.com/cryptondo"),
        ("DegenRex on X", "https://x.com/degen2ex"),
        ("DegenRex's Telegram channel", "https://t.me/degenrex"),
        ("Laba Lab (Michael Wyann Santoso)", "https://laba-lab.com"),
        ("Laba Lab membership and bootcamps", "https://laba-lab.com/membership"),
        ("Michael Wyann Santoso on LinkedIn", "https://linkedin.com/in/michaelwyannsantoso"),
    ]),
    ("Evidence", [
        ("Kevin Susanto: fomo 24-hour profit board, 19-20 November 2026", "https://x.com/KevinSusanto/status/2123472525070102987"),
        ("Kevin Susanto: fomo trades tracker", "https://fomoradar.io/traders/kevinsusanto"),
        ("Bevan: published wallet (Solscan)", "https://solscan.io/account/8DC3hpFtU3mEU3FUCtDfU72SmxTbrw4w35fyMYP8qGxg"),
        ("Bevan: weekly live sessions", "https://x.com/bevanpramudt/status/2115727803245347257"),
        ("Bevan: 'The Kingslayer', 1 SOL to 10 SOL live", "https://x.com/Jodski888/status/2121255854792188378"),
        ("Ponyin: Trojan profit-and-loss card", "https://x.com/Ponyin/status/2037746268039209077"),
        ("Tomket Lovers: trading-clan promotion", "https://x.com/TomketLovers/status/2119817955244663215"),
        ("GuarEmperor: $PRL call and profit post, 14 October 2026", "https://x.com/GuarEmperor/status/2110335710851117278"),
        ("GuarEmperor: member results from his tools", "https://x.com/GuarEmperor/status/2115414023114256560"),
        ("GuarEmperor: public wallet lists", "https://github.com/ariefzzz5421/wallet-tracker-transfer"),
        ("Brandon Salim: pinned fomo profit post", "https://x.com/brandonsalim/status/2096005957692301790"),
        ("DegenRex wins Clash of Trenchers (Ave.ai)", "https://x.com/AveaiGlobal/status/2120891736373518591"),
        ("Indonesia Crypto Network: 12 trusted crypto influencers (Laba Lab)", "https://indonesiacrypto.network/blog/12-trusted-crypto-influencers-in-indonesia-shaping-the-industry"),
        ("Degen Night speakers, 31 October 2026", "https://x.com/0xRememberUs/status/2116500473356411230"),
        ("Web3 Week Asia recap", "https://x.com/W3WOfficial/status/2121498996955132404"),
    ]),
    ("Benchmarks", [
        ("@0xSarimin: Trojan referral results (low case)", "https://x.com/0xSarimin"),
        ("@0xRusdi: BasedBot referral earnings (base case)", "https://x.com/0xRusdi"),
        ("@AiSlimx: $547 a day on OKX Wallet Boost (high case)", "https://x.com/AiSlimx"),
    ]),
    ("Market", [
        ("Indonesia Crypto Network: 10 X crypto influencers in Indonesia", "https://indonesiacrypto.network/blog/10-x-twitter-crypto-influencers-growing-the-indonesian-cryptosphere"),
        ("Indonesia Crypto Network: crypto community sizes (March 2026)", "https://indonesiacrypto.network/blog/crypto-community-indonesia"),
    ]),
    ("Tools", [
        ("X API: follower counts and median views on recent posts", "https://developer.x.com"),
        ("Solscan: wallet swaps and profit", "https://solscan.io"),
        ("DefiLlama: OKX DEX volume and fees", "https://defillama.com"),
        ("Discord invite counts: members and members online", "https://discord.com/developers/docs/resources/invite"),
        ("Telegram public channel previews: member counts", "https://t.me/s/KevinSusantoAlpha"),
        ("Claude: research, automation and fact-checks", "https://claude.ai"),
    ]),
]

TASK3 = [
    ("First message (X direct message)",
     "A warm, personal opener from someone who follows his trading, the offer in two sentences, and an easy next step.",
     "Hi Kevin, I'm Fedric from OKX's on-chain team in Indonesia. I've been following your trades on fomo for a while, and I'd love to build "
     "something with you for the people who trade alongside you.\n\n"
     "We'd give you a share of every swap your followers make on OKX Wallet, and once your link really takes off, we'd build a structure "
     "around you, much like OKX's VIP tiers. It runs alongside your current OKX link, so nothing changes on your side.\n\n"
     "Would you be open to a coffee in Jakarta or a quick call? Here's my calendar (Calendly link), or message me on Telegram (handle)."),
    ("First message, Bahasa Indonesia",
     "The same message in Bahasa Indonesia.",
     "Hai Kevin, saya Fedric dari tim on-chain OKX di Indonesia. Saya sudah cukup lama mengikuti trade kamu di fomo, dan saya ingin "
     "membangun sesuatu bersama kamu untuk orang-orang yang trading bersama kamu.\n\n"
     "Kami akan memberi kamu bagian dari setiap swap yang dilakukan follower kamu di OKX Wallet, dan begitu link kamu benar-benar berkembang, "
     "kami akan menyusun struktur khusus untuk kamu, mirip tingkatan VIP di OKX. Ini berjalan bersama link OKX kamu yang sekarang, jadi "
     "tidak ada yang berubah di sisi kamu.\n\n"
     "Apakah kamu terbuka untuk ngopi di Jakarta atau call singkat? Ini kalender saya (link Calendly), atau kirim pesan ke Telegram saya (handle)."),
    ("Follow-up (day 4, no reply)",
     "A short, friendly nudge with a real reason to reply now and a booking link.",
     "Hi Kevin, bumping this in case it got buried. We're bringing on our first group of Indonesian partners this week, and I'd really like "
     "you to be part of it: the launch includes a co-hosted live session for your Telegram. I've kept a spot for you. Grab any time here "
     "(Calendly link), or message me on Telegram (handle)."),
    ("Reply to 'another platform pays more'",
     "Listen, ask what they offer, and come back within 48 hours with the best we can build around him.",
     "Totally understand, and thanks for being open with me. I'd really like to make this work for you, so let me see what we can do. "
     "Could you share roughly what they're offering? I'll take it to our team and come back within 48 hours with the best we can put "
     "together, built around how your community actually trades. And whatever we land on runs alongside what you already have, so you "
     "never have to choose."),
    ("First-meeting question 1", "His needs: what would make this worth his time.",
     "What would make this partnership genuinely worth your time: extra income, tools for your community, or something else?"),
    ("First-meeting question 2", "His community: where they get stuck, and what we can provide.",
     "Where do your followers get stuck when they trade on-chain today, and what could we provide to help them: guides, live sessions or direct support?"),
    ("First-meeting question 3", "The lightest format that works, so it costs least to run.",
     "What does your community respond to best, and what's the lightest way to run this together: a pinned link, a weekly live session or a trading competition?"),
    ("No unverified promises", "Every term is a proposal; only published OKX DEX referral features are stated as fact.",
     "Every term in these messages is a proposal. The referral share and follower discount are published OKX DEX referral features "
     "(20% by default, up to 50% by tier). His personal tiers above $10,000 of monthly fees, the co-hosted session and the launch terms are "
     "offered subject to OKX approval, and OKX's regional marketing team confirms how his existing OKX link is tracked before any post."),
]
_k = partners[0]
_f4, _t4, _, _ft = ramp(_k, 4)
_f7, _t7, _, _ = ramp(_k, 7)
_f10, _t10, _, _ = ramp(_k, 10)
_f14, _t14, _s14, _ = ramp(_k, 14)
PLAN14 = [  # targets follow the Task 4 ramp for Kevin Susanto (base case)
    ("Days 1-2", "Set-up: a separate OKX DEX referral link and follower discount, tracked apart from his okx.ac link; share, tier trigger and payout wallet in writing; shared dashboard", "Link live; terms signed"),
    ("Days 2-4", "Onboard his power traders one-to-one through the full user flow: wallet set-up, funding, first Meme Mode trade, Track alerts; log every drop-off",
     f"{_k['ptraders']} FnF power traders trading by day 4 (~{_ft:.0f} trades a day)"),
    ("Day 5", "Launch the campaign: a co-hosted live trading session on his Telegram and a 7-day trading competition (screened wallets only)", "Session held; first sign-ups"),
    ("Days 6-7", "Push the flow: 2 X threads with the link, a pinned Bahasa and English guide, and an FAQ for first-time on-chain traders", f"~{_f7:.0f} FTDs by day 7"),
    ("Days 8-10", "Fix the funnel: check each step (click, wallet, deposit, first and second trade), fix the weakest, send product feedback", f"~{_f10:.0f} FTDs and ~{_t10:.0f} trades a day by day 10"),
    ("Days 11-14", "Competition results and a second live session; review FTDs, trades a day and flagged wallets with Kevin; agree days 15-30", f"~{_f14:.0f} FTDs and ~{_t14:.0f} trades a day by day 14"),
]

GANTT = [  # (activity, track, KPI key or text, start day, end day)
    ("Score and message prospects, warm introductions first", "Our 8", "msgs", 1, 21),
    ("Meetings and pitches", "Our 8", "meet", 2, 24),
    ("Sign the top 3: Kevin Susanto, Bevan, 0xRusdi", "Our 8", "3 signed by day 7", 3, 7),
    ("Top 3 live: onboard their FnF power traders one-to-one", "Our 8", "fnf3", 8, 11),
    ("In-person day in Jakarta: power traders onboarded in person", "Our 8", "Meetings held, traders set up", 10, 10),
    ("Next 3 sign and go live: Brandon Salim, GuarEmperor, Laba Lab", "Our 8", "live6", 8, 14),
    ("Last 2 sign and go live: Crypto Ndo, DegenRex", "Our 8", "live8", 15, 20),
    ("Major push: a trading competition and a live session with each live partner", "Our 8", "One of each per partner", 12, 28),
    ("Fix the weakest funnel step; product feedback", "Our 8", "Drop-off logged at each step", 15, 21),
    ("Review and last push: second sessions, competition finals", "Our 8", "act", 22, 29),
    ("Day-30 decision per partner and month-2 plan", "Our 8", "Continue, renegotiate or stop", 30, 30),
    ("Pilot C: check its 90-day dashboards and top wallets", "Pilot C", "Volume confirmed before go-live", 1, 9),
    ("Pilot C: live on the growth deal", "Pilot C", "pil", 10, 30),
]



for research, name in ((False, "OKX_DEX_BD_Affiliate_Indonesia_Prospect_Spreadsheet.xlsx"), (True, "OKX_research_and_checks.xlsx")):
    wb = build(research)
    wb.save(OUTDIR / name)
    print("saved", name)
