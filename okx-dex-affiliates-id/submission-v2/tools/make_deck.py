"""Rewrite the six slides of deck_v2.html from the recalculated submission workbook (research merge, Bevan as #1).

Usage: python make_deck.py <deck_v2.html> <submission.xlsx>
"""
import html
import re
import sys
from pathlib import Path

from openpyxl import load_workbook

deck, xlsx = Path(sys.argv[1]), Path(sys.argv[2])
wb = load_workbook(xlsx, data_only=True)
t1, t2, t3, t4, mi = (wb[n] for n in ("Task 1", "Task 2", "Task 3", "Task 4", "Miscellaneous"))
E = html.escape


def k(x):  # $129K
    return f"${x / 1000:,.0f}K"


def m(x):  # $1.12M
    return f"${x / 1e6:.2f}M" if x < 1e7 else f"${x / 1e6:.1f}M"


def usd(x, d=0):
    return f"${x:,.{d}f}"


def n0(x):
    return f"{x:,.0f}"


# ---------------------------------------------------------------- numbers
P = []
for i in range(8):
    r = 37 + i
    P.append(dict(ftd=t1.cell(r, 6).value, fnf=t1.cell(r, 8).value, tpd=t1.cell(r, 9).value, swaps=t1.cell(r, 13).value,
                  rev=t1.cell(r, 15).value, low=t1.cell(r, 12).value, high=t1.cell(r, 14).value, pool=t1.cell(r, 3).value,
                  ftd_sw=t1.cell(r, 10).value, fnf_sw=t1.cell(r, 11).value, score=t1.cell(26 + i, 9).value,
                  prio=t1.cell(12 + i, 5).value))
T = dict(ftd=t1["F45"].value, fnf=t1["H45"].value, tpd=t1["I45"].value, swaps=t1["M45"].value, rev=t1["O45"].value,
         keeps=t1["R45"].value, low=t1["L45"].value, high=t1["N45"].value, ftd_sw=t1["J45"].value)
fnf_share, nofnf, nofnf_rev, top_ftd_share, pub_up = (t1[f"B{r}"].value for r in (48, 49, 50, 51, 53))
low_rev, high_rev = T["low"] * 0.0004, T["high"] * 0.0004
c24, c25, c26, c27, g27 = (t2[c].value for c in ("F24", "F25", "F26", "F27", "G27"))
bev_earn, bev_earn50, breakeven, bev_pub_keep = t2["H24"].value, t2["I24"].value, t2["J24"].value, t2["K24"].value
bev_pub_earn = t3["B31"].value
wk = {c: [t4[f"{c}{r}"].value for r in (10, 12, 13, 14, 15)] for c in "BCDE"}
rev30, rev30_after, runrate, b_swaps = t4["E17"].value, t4["E18"].value, t4["E19"].value, t4["E16"].value
grow = [(mi[f"B{r}"].value, mi[f"C{r}"].value) for r in range(55, 60)]
bev = P[0]
bev_fnf_share = bev["fnf_sw"] / bev["swaps"]
fnf_day = mi["C8"].value / 30
fnf_tpd = t1["H37"].value * mi["E19"].value  # Bevan's 5 FnF traders' trades a day
assert abs(T["ftd"] - sum(p["ftd"] for p in P)) < 1e-6

names = ["Bevan / Narrative Labs", "0xRusdi", "Ponyin", "Crypto Ndo", "Kevin Susanto", "GuarEmperor + FnF", "Meteora LP Army", "Stress Capital"]
types = ["Trader-educator", "Bot referrer", "Trader KOL", "Trench creator", "Large KOL", "Alpha caller", "LP community", "Trading clan"]
proof = ["Wallet +374 SOL net, May–Jul", "$8.5K BasedBot referral earnings", "Trojan PnL: ~$170K in March", "Uses OKX Wallet; 59.3K X",
         "fomo top-3 trader, 19–20 Nov", "1.9% engagement; member PnL posts", "462 + 671 at two meetups", "30.7K Discord, 2,453 online"]
contact = ["Jakarta", "X DM", "X DM", "X DM", "Jakarta", "X DM", "Events", "X DM"]

s = deck.read_text()


def replace_slide(nr, new):
    global s
    a = s.index(f'<section class="slide-wrap" id="slide-{nr}">')
    b = s.index("</section>", a) + len("</section>")
    s = s[:a] + new.strip("\n") + s[b:]


# ================================================================ slide 1
replace_slide(1, f"""
<section class="slide-wrap" id="slide-1">
  <article class="slide" aria-label="Slide 1 of 6"><div class="frame">
    <header>
      <div class="kicker">Task 1 · Who we target and how we measure it</div>
      <h2>Trenchers and their FnF power traders earn OKX the most</h2>
    </header>
    <div class="s-body">
      <div class="tscroll"><table class="t">
        <thead><tr><th>OKX product</th><th>OKX fee<sup>1</sup></th><th class="num">Swaps/trader a month<sup>2</sup></th><th class="num">Revenue/trader, OKX fee<sup>1</sup></th><th>Best audience</th></tr></thead>
        <tbody>
          <tr class="hl"><td>Meme Mode + Track</td><td>0.5%</td><td class="num">$2K · FnF $16.4K</td><td class="num">$10 · FnF $82</td><td>Memecoin trenchers</td></tr>
          <tr><td>Swaps + Bridge</td><td>0.1–0.25%</td><td class="num">$2K · FnF $16.4K</td><td class="num">$3 · FnF $25</td><td>DeFi / LP users</td></tr>
          <tr><td>Boost + tokenized stocks</td><td>0.25–0.5%</td><td class="num">$600</td><td class="num">$1.80</td><td>Beginners, stock investors</td></tr>
          <tr><td>Easy-mode majors</td><td>0.1%</td><td class="num">$300</td><td class="num">$0.30</td><td>Beginners today</td></tr>
        </tbody>
      </table></div>
      <div class="steps">
        <div class="steprow">
          <div class="step"><b>Reach P</b>views + members</div>
          <div class="step"><b>× 2%</b>sign up<sup>2</sup></div>
          <div class="step"><b>× 15%</b>trade: 1 FTD<sup>2</sup></div>
          <div class="step"><b>$2K a month</b>swaps per FTD<sup>2</sup></div>
          <div class="step"><b>{n0(T['ftd'])} FTDs</b>from our 8</div>
          <div class="step"><b>$0.80</b>per FTD at 0.04%</div>
        </div>
        <div class="steprow">
          <div class="step"><b>1 FnF power trader</b>from the circle</div>
          <div class="step"><b>Days 1–3</b>onboarded 1:1</div>
          <div class="step"><b>$547 a day</b>OKX Wallet trader<sup>2</sup></div>
          <div class="step"><b>$16.4K a month</b>of swaps</div>
          <div class="step"><b>= 8 FTDs</b>of volume</div>
          <div class="step"><b>$6.56</b>at 0.04%</div>
        </div>
        <div class="banner">So every partner onboards its FnF power traders in days 1–3, then posts for FTDs; we track FTDs and trades a day</div>
      </div>
    </div>
    <p class="foot"><sup>1</sup> OKX DEX fee schedule (web3.okx.com/dex-fees) and fee notice effective 12 May 2026 (Group 1 or 2 against other tokens: 0.5%); tokenized stocks at the Group 2 rate since 7 Sep 2026 · <sup>2</sup> Research sheet base rates (20 Nov 2026), tested in the pilot; FnF: @AiSlimx on OKX Wallet Boost, May 2026</p>
  </div></article>
  <details class="notes">
    <summary>Speaker notes · slide 1</summary>
    <div class="body">
      <p>What OKX Wallet and DEX offer, and what each earns. OKX DEX charges an interface fee by token group: 0.1% between majors and stablecoins (Group 1), 0.25% for mid-caps (Group 2), 0.5% when Group 1 or 2 is swapped against any other token, which covers most memecoin buys with SOL or USDC (OKX notice, effective 12 May 2026; swaps between two "other" tokens stay at 0%). Meme Mode itself launched fee-free in July 2024, so the 0.5% is recent and worth confirming on the day. Tokenized stocks (xStocks) moved to the Group 2 rate on 7 Sep 2026. Earn carries no interface fee, and perps opened on a rival venue through OKX Wallet's dApp connection pay that venue, not OKX, so perp communities are out of scope.</p>
      <p>How we size every partner (the research sheet's method, the same formula for all 8): reach pool P = median views on the partner's recent original posts plus Discord members online (or 20% of Telegram subscribers). 2% of P sign up through the link and 15% of those trade, so 0.3% of P become FTDs (first-time depositors who trade). Each FTD swaps about $2,000 in its first month. Low and high cases use 1% × 6.1% × $286 (Trojan referral data, @0xSarimin) and 4% × 30% × $16,400 (a campaign-boosted OKX Wallet trader). On top, each partner introduces FnF power traders (its friends-and-family trading circle, clan or referral book) in days 1-3; each trades about $547 a day, the only measured OKX Wallet figure (@AiSlimx, OKX Wallet Boost, May 2026), so $16.4K a month, the volume of 8 FTDs.</p>
      <p>Revenue per trader at OKX's published fee: an FTD trencher $10 a month and an FnF power trencher about $82; DeFi swaps at a ~0.15% blend $3 and $25; Boost and tokenized stocks $600 × ~0.3% = $1.80; Easy-mode majors $300 × 0.1% = $0.30. So most BD time goes to trenchers, and every partner onboards its FnF power traders before posting for FTDs: they carry {fnf_share:.0%} of our 8 partners' swaps. Fee upside: OKX DEX earned about $4.4M in Oct 2026 (DefiLlama), an average take of about 0.09%; at published fees our 8 partners' first-month swaps ({m(T['swaps'])}) would generate about ${pub_up / 1000:.1f}K before the partner's share, against {usd(T['rev'])} at the brief's 0.04%. Detail: spreadsheet, Task 1 (section C) and Miscellaneous tabs.</p>
    </div>
  </details>
</section>
""")

# ================================================================ slide 2
rows = ""
for i in range(8):
    p = P[i]
    cls = ' class="hl"' if i < 3 else ""
    rows += (f'          <tr{cls}><td>{i + 1}</td><td>{E(names[i])}</td><td>{types[i]}</td><td>{E(proof[i])}</td><td>{contact[i]}</td>'
             f'<td>{p["prio"]}</td><td class="num">{p["fnf"]}</td><td class="num">{p["ftd"]:.0f}</td><td class="num">{p["tpd"]:.0f}</td>'
             f'<td class="num">{k(p["swaps"])}</td><td class="num">{usd(p["rev"])}</td></tr>\n')
rows += (f'          <tr><td></td><td>Total</td><td></td><td></td><td></td><td></td><td class="num">{T["fnf"]}</td><td class="num">{T["ftd"]:.0f}</td>'
         f'<td class="num">{T["tpd"]:.0f}</td><td class="num">{m(T["swaps"])}</td><td class="num">{usd(T["rev"])}</td></tr>\n')
picks = [
    ("Wallet: +374 SOL net, May–Jul", "His paid Masterclass (29 seats)", "GMGN (code BEVANP) → OKX"),
    ("$8,463 BasedBot referral pay", "His BasedBot referral book", "BasedBot users → OKX"),
    ("Trojan PnL: ~$170K in March", "His trading circle", "Trojan users → OKX"),
]
pick_html = ""
for i in range(3):
    p = P[i]
    pr, fn, fl = picks[i]
    pick_html += f"""          <div class="pick">
            <div class="hd"><b>{i + 1} · {E(names[i])}</b><span class="score">Score {p['score']:.2f}</span></div>
            <dl>
              <dt><i>1</i>Proof</dt><dd>{E(pr)}</dd>
              <dt><i>2</i>FnF circle</dt><dd>{E(fn)}</dd>
              <dt><i>3</i>Flow to move</dt><dd>{E(fl)}</dd>
            </dl>
            <div class="kpi"><i>4</i>Trencher · {p['ftd']:.0f} FTDs · {p['tpd']:.0f} trades a day · {k(p['swaps'])}</div>
          </div>
"""
top3_ftd = sum(P[i]["ftd"] for i in range(3))
ranges = ", ".join(f"{names[i].replace(' + FnF', '')} {k(P[i]['low'])}-{m(P[i]['high'])}" for i in range(8))
scores = ", ".join(f"{names[i].replace(' + FnF', '')} {P[i]['score']:.2f}" for i in range(8))
replace_slide(2, f"""
<section class="slide-wrap" id="slide-2">
  <article class="slide" aria-label="Slide 2 of 6"><div class="frame">
    <header>
      <div class="kicker">Task 1 · 8 partners · Each partner's first month</div>
      <h2>8 partners, ranked on proof of real trading</h2>
    </header>
    <div class="s-body">
      <div class="tscroll"><table class="t">
        <thead><tr><th>#</th><th>Partner</th><th>Type</th><th>Verified proof<sup>1</sup></th><th>Contact</th><th>Priority</th><th class="num">FnF<sup>2</sup></th><th class="num">FTDs<sup>2</sup></th><th class="num">Trades/day<sup>2</sup></th><th class="num">Swaps<sup>2</sup></th><th class="num">Revenue<sup>3</sup></th></tr></thead>
        <tbody>
{rows}        </tbody>
      </table></div>
      <div class="why">
        <div class="chain" aria-label="How the top 3 were picked, in order">
          <span class="lbl">Why these 3 first</span>
          <span class="c"><i>1</i>Proof they already trade</span><span class="arr">→</span>
          <span class="c"><i>2</i>FnF traders they bring (~{fnf_share:.0%} of swaps)</span><span class="arr">→</span>
          <span class="c"><i>3</i>Rival flow we move to OKX</span><span class="arr">→</span>
          <span class="c"><i>4</i>Trencher fees: 0.5%</span>
        </div>
        <div class="top3">
{pick_html}        </div>
        <p class="tail"><b>Reach alone isn't enough:</b> Kevin Susanto (#5) brings {P[4]['ftd']:.0f} FTDs, more than the top 3 together ({top3_ftd:.0f}), but already holds an OKX referral link and has no FnF circle on record.</p>
      </div>
    </div>
    <p class="foot"><sup>1</sup> Solscan (Bevan's published wallet); X posts by @0xRusdi, @Ponyin, @cryptondo, @KevinSusanto, @GuarEmperor (2026); Discord counts, 20 Nov 2026; Luma · <sup>2</sup> FTDs = reach × 2% sign up × 15% trade (research sheet); FnF = power traders each partner brings, $16.4K a month each from day 4; trades/day once all trade · <sup>3</sup> Swaps × 0.04%, before partner shares (OKX keeps {usd(T['keeps'])}) · Score: Task 1 tab</p>
  </div></article>
  <details class="notes">
    <summary>Speaker notes · slide 2</summary>
    <div class="body">
      <p>All figures are each partner's first month, base case: FTDs, trades a day once everyone is trading, swaps, and revenue at the brief's 0.04% before the partner's share (after shares, OKX keeps {usd(T['keeps'])} of the {usd(T['rev'])}). How the numbers are built, the same way for all 8 (the research sheet's method): reach pool P = median views on recent original posts plus Discord members online, or plus 20% of Telegram subscribers; FTDs = P × 2% who sign up × 15% who trade; each FTD swaps about $2,000 in its first month. Five partners' reach is measured (X API and Discord counts, 20 Nov 2026); 0xRusdi, Crypto Ndo and MeteoraIDN use 20% of followers, the median ratio across the 11 accounts the research measured, until we measure them. On top come the FnF power traders each partner introduces in days 1-3 (5 each; 3 for Stress Capital, 2 for Kevin Susanto, who has no circle on record), at about $547 a day each from day 4, the one measured OKX Wallet figure. FnF power traders carry {fnf_share:.0%} of swaps; without them the 8 bring {k(nofnf)} ({usd(nofnf_rev)}), so the FnF ask is a first-call gate, not a hope. Ranges (low to high): {ranges}; total {k(T['low'])}-{m(T['high'])} ({usd(low_rev)}-{usd(high_rev)}). The high case assumes every FTD trades like a campaign-boosted power user, so we plan on the base.</p>
      <p>Ranking: a 7-criterion score (proof of real trading 20%, FTD potential 20%, FnF power-trader access 15%, rival flow to move 15%, reachability 10%, segment value 10%, risk 10%): {scores}.</p>
      <p><b>Why these 3 first: the logic, in order.</b> 1. Proof they already trade: public numbers, not follower counts (the brief says follower count alone isn't evidence). 2. FnF power traders they can bring: the partner's friends-and-family trading circle carries {fnf_share:.0%} of swaps, so a partner with a paid class, a referral book or a trading circle beats a bigger audience. 3. Rival flow we can move: their traders already use GMGN, BasedBot or Trojan, so we move existing volume to OKX instead of creating new traders from scratch. 4. Trencher fees: memecoin swaps pay OKX 0.5%, the highest fee. Then we measure FTDs first and trades a day second.</p>
      <p>How each passes. Bevan / Narrative Labs is #1, our best prospect: the wallet he publishes on narrativelabs.id shows 91 swaps from 14 May to 30 Jul 2026, 352.6 SOL bought and 726.8 SOL sold, +374 SOL net (read on Solscan); he founded Narrative Labs, a paid Masterclass (29 of 40 promo seats sold) with a 9,238-member Discord, 984 online; his members trade on GMGN through his code BEVANP; and he runs member meetups in Jakarta, so we can onboard his FnF traders in a room. He is the only partner that passes all four filters and can be met in person. 0xRusdi's BasedBot dashboard shows $8,463 of referral earnings at the 25% tier, the only public proof on our list that a partner's referred traders actually trade (at a 1% bot fee, about $3.4M referred; we confirm on the first call). Ponyin's Trojan PnL card shows 1.13K SOL bought and 1.24K SOL sold in March 2026, about $170K of volume: a proven heavy trader whose circle trades on Trojan.</p>
      <p>Why Kevin Susanto is #5, not top 3: his measured reach (41,248 median views a post) gives {P[4]['ftd']:.0f} FTDs, {top_ftd_share:.0%} of the 8 partners' total and more than the top 3 together. But he already holds an OKX referral link (okx.ac/join/VVIP7777), so some of his flow may already count for OKX; he has no FnF circle on record; his engagement is low (0.84%); and he advises Lighter, a rival perp DEX. We route him through OKX's regional team in wave 2.</p>
      <p>Priority reasons for the rest: Crypto Ndo, the largest trencher reach on X (59.3K) and an OKX Wallet user himself, wave 1 if a top-3 partner says no; GuarEmperor, the best engagement among the big accounts (1.9%) and an FnF circle that trades from his tools, but a feed with WL and mint posts; Meteora LP Army, the largest in-person onchain crowds (462 at LP Army Jakarta, 671 at MET Bali) and meetable; Stress Capital, a 30,663-member clan Discord (2,453 online) run by Orpheuzkaze, Low until clan leads show they trade. Removed or benched after the research merge: Hyperliquid Indonesia (a rival perp venue), mage (moved his main wallet to Bitget Wallet on 11 Nov), TwipsX (no proof its readers trade), Superteam Indonesia (builder-led) and Tomket Lovers (runs Bitget Wallet promos).</p>
      <p>Brief checks: 8 partners, 6 not airdrop-led (GuarEmperor and Stress Capital lean partly airdrop and NFT), 5 partner types (a trader-educator with a paid community: Bevan; X trading creators: 0xRusdi, Ponyin, Crypto Ndo, GuarEmperor; a large KOL: Kevin Susanto; an LP community: Meteora LP Army; a trading clan: Stress Capital), 3 meetable in person (Bevan and Kevin Susanto in Jakarta, Meteora LP Army at its nights). Links, audience and market, proposed partnership and main risk for all 8: spreadsheet, Task 1 tab.</p>
    </div>
  </details>
</section>
""")

# ================================================================ slide 3 (Task 2): notes paragraph on the real top 3
old_p = s[s.index("<p>Same test on our real top 3"):]
old_p = old_p[:old_p.index("</p>") + 4]
s = s.replace(old_p, (f"<p>Same test on our real top 3 (spreadsheet, Task 2 tab, rows 22-33): month-1 contribution after their 30% share is "
                      f"{usd(c24)} (Bevan), {usd(c25)} (0xRusdi) and {usd(c26)} (Ponyin), {usd(c27)} together; at 50% of forecast it halves to "
                      f"{usd(g27)} but stays positive because there is no fixed fee. Even a $500 fixed fee would need about ${breakeven / 1e6:.1f}M "
                      f"of swaps to break even against roughly {k(sum(P[i]['swaps'] for i in range(3)) / 3)} forecast per partner, so real partners get "
                      f"revenue share only until 90 days of data say otherwise.</p>"))

# ================================================================ slide 4 (Task 3 on Bevan)
T3 = {t3.cell(r, 1).value: t3.cell(r, 2).value for r in range(4, 12)}
msg = t3["B4"].value
wc = len(msg.split())
fu = T3["Follow-up (day 4, no reply)"]
obj = T3["Reply to 'another platform pays more'"]
q = [T3[f"Meeting question {i}"] for i in (1, 2, 3)]
replace_slide(4, f"""
<section class="slide-wrap" id="slide-4">
  <article class="slide" aria-label="Slide 4 of 6"><div class="frame">
    <header>
      <div class="kicker">Task 3 · Win and activate · Worked on our #1 prospect</div>
      <h2>One playbook to win and activate any partner</h2>
    </header>
    <div class="s-body">
      <div class="tscroll"><table class="t">
        <thead><tr><th>Stage</th><th>Days (A1 = agreement)</th><th>Playbook (any partner)</th><th>Worked example: Bevan / Narrative Labs, our #1</th></tr></thead>
        <tbody>
          <tr><td>Open</td><td class="nowrap">Job days 1–4</td><td>DM citing their public proof; follow-up day 4</td><td>Cites his published wallet: +374 SOL net<sup>1</sup></td></tr>
          <tr><td>First call</td><td class="nowrap">Job days 5–7</td><td>Ask for wallet proof + 5 FnF intros</td><td>Current wallet, Discord stats; terms inside OKX's tiers<sup>3</sup></td></tr>
          <tr><td>Set up</td><td class="nowrap">A1–A2</td><td>Link + discount, Bahasa guide, compliance check</td><td>Meme Mode + Track guide; certification status</td></tr>
          <tr class="hl"><td>FnF power traders</td><td class="nowrap">A2–A4</td><td>Onboard 5 FnF power traders 1:1 before any post</td><td>5 Masterclass traders, in Jakarta: ~{fnf_tpd:.0f} trades a day<sup>2</sup></td></tr>
          <tr><td>Posts for FTDs</td><td class="nowrap">A5–A7</td><td>A live session + posts with the referral link</td><td>Discord workshop + 2 posts: 6+ FTDs by A7<sup>2</sup></td></tr>
          <tr><td>Review</td><td class="nowrap">A8–A14</td><td>Daily FTD + trade check; fix the weakest</td><td>~13 FTDs, ~57 trades a day by A14<sup>2</sup></td></tr>
        </tbody>
      </table></div>
      <div class="cards three">
        <div class="card blue"><h3>Why one playbook</h3><ul><li>Proof first: their own numbers</li><li>FnF power traders: ~{fnf_share:.0%} of swaps</li><li>Same targets: FTDs, trades a day</li></ul></div>
        <div class="card green"><h3>#1's first month: {k(bev['swaps'])}</h3><p class="soft">{bev['ftd']:.0f} FTDs, ~{bev['tpd']:.0f} trades a day<sup>2</sup> · he earns {usd(bev_pub_earn)} (30% of the 0.5% fee)<sup>3</sup> · OKX keeps {usd(bev_pub_keep)} at the published fee, or {usd(c24)} at the brief's 0.04% after his share</p></div>
        <div class="card amber"><h3>If a step falls short</h3><ul><li>FnF traders &lt;3 → 5 more intros</li><li>FTDs &lt;4 by A7 → live set-up</li><li>Trades &lt;50 a day → presets room</li></ul></div>
      </div>
    </div>
    <p class="foot"><sup>1</sup> Solscan: wallet 8DC3hp…qGxg, published on narrativelabs.id; 91 swaps, 14 May–30 Jul 2026 · <sup>2</sup> Research base: reach {n0(bev['pool'])} × 2% × 15% = {bev['ftd']:.0f} FTDs at $2K a month; FnF at $16.4K a month from A4 · <sup>3</sup> OKX DEX fees: 0.5% on memecoins; referral share up to 50% by tier (regional help pages list different tier tables; confirm the live rate)</p>
  </div></article>
  <details class="notes">
    <summary>Speaker notes · slide 4 (includes the outreach texts)</summary>
    <div class="body">
      <p>Why the slide is a playbook with a named example: the brief asks for the outreach message, follow-up, objection reply, 14-day activation plan and 3 meeting questions on our highest-priority real prospect from Task 1, so those are written for Bevan Pramudito of Narrative Labs (below and in the spreadsheet's Task 3 tab). The left column is what we repeat for every partner: open with their own public proof, ask for wallet or dashboard proof and 5 FnF intros on the first call, set up the link with a follower discount and a compliance check, onboard FnF power traders 1:1 before any post ({bev_fnf_share:.0%} of Bevan's forecast swaps), then run a live session and posts for FTDs and review FTDs and trades a day daily. A1 to A14 is the 14-day activation plan (A1 = the day he agrees; in the 30-day calendar that is about job day 10).</p>
      <p>Compliance order matters here. POJK 6/2026 requires certification for anyone publishing crypto information and makes the firm answerable for its marketing partners, so public posts wait for his certification status and OKX legal's sign-off. Private 1:1 onboarding of 5 Masterclass traders is not a public post, so it can likely start first (confirm with legal); that's another reason FnF power traders go before posts.</p>
      <p>Why Bevan is #1: he is the only prospect that passes all four filters and can be met in person. Proof: his published wallet shows 91 swaps, about 1,080 SOL traded and +374 SOL net from May to July 2026. FnF: he founded Narrative Labs, a paid Masterclass with a 9,238-member Discord (984 online). Flow: his members trade on GMGN through his code BEVANP. Fees: memecoins, 0.5%. Route: X DM, then a Narrative Labs session in Jakarta. Risks: the published wallet has been idle since 30 Jul, so we ask which wallet he trades from now; and he already earns on GMGN and Mobee codes, so the pitch is non-exclusive.</p>
      <p><b>First message ({wc} words):</b></p>
      <blockquote>{E(msg)}</blockquote>
      <p>Bahasa version: spreadsheet, Task 3 tab.</p>
      <p><b>Follow-up on day 4:</b></p>
      <blockquote>{E(fu)}</blockquote>
      <p><b>If he says another platform pays more:</b></p>
      <blockquote>{E(obj)}</blockquote>
      <p><b>Meeting questions:</b> 1. {E(q[0])} 2. {E(q[1])} 3. {E(q[2])}</p>
      <p><b>Targets:</b> 5 FnF power traders trading by A4 (~{fnf_tpd:.0f} trades a day); 6+ FTDs by A7; ~9 FTDs and ~47 trades a day by A10; ~13 FTDs and ~57 trades a day by A14; for the month, {bev['ftd']:.0f} FTDs, ~{bev['tpd']:.0f} trades a day once all are trading and ~{k(bev['swaps'])} of swaps ({usd(bev['rev'])} at 0.04% before his share, {usd(c24)} after). No rates, rewards or features are promised beyond OKX's published referral programme; a Narrative Alpha trade button on the DEX API is scoped, not promised.</p>
    </div>
  </details>
</section>
""")

# ================================================================ slide 5 (Task 4)
live = [f"{wk[c][0]} + B" if c != "B" else "0" for c in "BCDE"]
replace_slide(5, f"""
<section class="slide-wrap" id="slide-5">
  <article class="slide" aria-label="Slide 5 of 6"><div class="frame">
    <header>
      <div class="kicker">Task 4 · First 30 days</div>
      <h2>30 days: target, sign, activate, review</h2>
    </header>
    <div class="s-body">
      <div class="tscroll"><table class="t">
        <thead><tr><th>Week</th><th>Focus</th><th class="num">Outreach</th><th class="num">Meetings</th><th class="num">Live<sup>1</sup></th><th class="num">Activated</th><th class="num">FTDs<sup>2</sup></th><th class="num">Trades/day<sup>2</sup></th><th class="num">Swaps<sup>2</sup></th></tr></thead>
        <tbody>
          <tr><td class="nowrap">1 · days 1–7</td><td>Score 25 prospects; message all</td><td class="num">25</td><td class="num">3</td><td class="num">0</td><td class="num">0</td><td class="num">0</td><td class="num">0</td><td class="num">$0</td></tr>
          <tr><td class="nowrap">2 · days 8–14</td><td>Sign top 3 and B; Jakarta day</td><td class="num">20</td><td class="num">6</td><td class="num">{live[1]}</td><td class="num">{wk['C'][1]}</td><td class="num">{wk['C'][2]:.0f}</td><td class="num">{wk['C'][3]:.0f}</td><td class="num">{k(wk['C'][4])}</td></tr>
          <tr><td class="nowrap">3 · days 15–21</td><td>3 more live: Ndo, Kevin, GE</td><td class="num">15</td><td class="num">5</td><td class="num">{live[2]}</td><td class="num">{wk['D'][1]}</td><td class="num">{wk['D'][2]:.0f}</td><td class="num">{wk['D'][3]:.0f}</td><td class="num">{k(wk['D'][4])}</td></tr>
          <tr class="hl"><td class="nowrap">4 · days 22–30</td><td>Meteora night; review; decide</td><td class="num">10</td><td class="num">4</td><td class="num">{live[3]}</td><td class="num">{wk['E'][1]}</td><td class="num">{wk['E'][2]:.0f}</td><td class="num">{wk['E'][3]:.0f}</td><td class="num">{k(wk['E'][4])}</td></tr>
        </tbody>
      </table></div>
      <div class="cards three">
        <div class="card blue"><h3>In-person day, Jakarta (Wed)</h3><ul><li>10:00 Bevan: sign, onboard FnF</li><li>13:00 Kevin Susanto (OKX intro)</li><li>19:00 LP Army night, else JURA's weekly W3AI Wednesday</li></ul></div>
        <div class="card white"><h3>If 20 DMs get 0 meetings</h3><ul><li>Warm intros: OKX's Bali alumni, hosts, admins</li><li>One observation + one ask; A/B test</li><li>Attend 2 rival-hosted events</li></ul></div>
        <div class="card amber"><h3>Monitor and adjust</h3><ul><li>Daily: FTDs, trades, FnF traders</li><li>Behind pace → live set-up, presets</li><li>No paid post before certification</li><li>&gt;10% flagged → pause payouts</li></ul></div>
      </div>
      <div class="card green wide-act">
        <p><b>Activated:</b> ≥10 FTDs and ≥50 trades a day within 30 days live, ≥30% of FTDs still trading after 14 days, &lt;5% flagged wallets.</p>
        <p><b>Day 30:</b> continue if activated · renegotiate if FnF trade but FTDs or trades fall short · stop if &lt;5 FTDs or &gt;10% flagged.</p>
      </div>
    </div>
    <p class="foot"><sup>1</sup> Assumption: our 8 go live in waves on days 10, 17 and 24; B is the brief's hypothetical partner, live from day 10 · <sup>2</sup> Our 8 only, calendar days (FTDs arrive through each partner's month), from the spreadsheet's Task 4 tab, section B; B adds {m(b_swaps)} of swaps by day 30; revenue {usd(rev30)} by day 30, {usd(rev30_after)} after partner shares</p>
  </div></article>
  <details class="notes">
    <summary>Speaker notes · slide 5</summary>
    <div class="body">
      <p>Weekly targets (calendar days of the job): week 1 score 25 prospects and message all of them, top 3 first (25 messages, 3 meetings); week 2 sign the top 3 (Bevan / Narrative Labs, 0xRusdi, Ponyin) and Partner B, onboard their FnF power traders, Jakarta day (20 messages, 6 meetings, 3 pilots live plus B); week 3 Crypto Ndo, Kevin Susanto and GuarEmperor go live (15 messages, 5 meetings, 6 live); week 4 the Meteora community night and Stress Capital go live, review and day-30 decisions (10 messages, 4 meetings, 8 live, {wk['E'][1]} activated: Bevan, 0xRusdi, Crypto Ndo, Kevin Susanto and GuarEmperor; Ponyin lands just under 10 FTDs by day 30). If one of the top 3 says no in week 1, Crypto Ndo moves into wave 1. FTDs and trades a day are running totals at the end of each week; swaps are our 8 only, on calendar days, so they trail Task 1's first-month figures: FTDs arrive through each partner's month and keep trading into month 2. Day 30: {k(wk['E'][4])} of swaps from our 8 plus {m(b_swaps)} from B, or {usd(rev30)} at 0.04% ({usd(rev30_after)} after partner shares); with B the run-rate on day 30 is about {m(runrate)} a month, or {usd(runrate * 0.0004)} a month at 0.04%.</p>
      <p>In-person day (week 2, Wednesday, Jakarta). Wednesday because JURA in Kemanggisan runs W3AI Wednesday every week, so the evening has an event even if everything else cancels. 10:00 Narrative Labs with Bevan: sign, then onboard his first Masterclass traders 1:1 on site (set up by DM a week ahead). 13:00 coffee with Kevin Susanto, arranged through OKX's regional marketing team, which already works with him. 15:00 a Jakarta-based trench creator from the pipeline. 19:00 the next LP Army night if one is listed on Luma (meet Evil Panda and Friday7th about the co-hosted night); otherwise W3AI Wednesday at JURA (check Luma that week). Fallback if a meeting falls through: a same-day video call, the next Jakarta prospect on the list, and the evening event regardless.</p>
      <p>If 20 tailored messages bring 2 replies and no meetings after two weeks: check whether the targets, the channel or the message is wrong; switch to warm intros, starting with OKX's own network (traders and community leaders from OKX's Bali Alpha Traders Summit and Alphas Connect co-hosts), then event hosts, community admins and mutuals; rewrite to one specific observation plus one clear ask; test two versions on the next 10 messages; attend two rival-hosted events in person (for example a MetaMask Community Builder Night in Jakarta).</p>
      <p>Monitoring: a 15-minute daily check of FTDs, trades a day and FnF power traders for every live partner; a Monday scorecard to each partner with one agreed fix; a Thursday pipeline review. Fixes: FTDs behind pace (under half the target for a week) get the follower discount and a live wallet-funding session; under 50 trades a day by day 14 gets a Meme Mode presets session and a Track alerts room; under 3 FnF power traders by day 7 means asking for 5 more intros; under $250 a day per FnF trader means checking chain and token coverage; over 10% flagged wallets pauses payouts with a 30-day clawback. Compliance gate: no paid public post goes out before the partner's POJK 6/2026 certification status is checked and OKX legal has signed off on promotion in Indonesia.</p>
      <p>Activated partner: at least 10 FTDs and at least 50 trades a day within 30 days live, at least 30% of FTDs still trading after 14 days, and under 5% flagged wallets; swaps follow from the first two. Day 30, per partner: continue if activated; renegotiate if its FnF power traders trade but FTDs or trades a day fall short; stop if under 5 FTDs or over 10% flagged.</p>
    </div>
  </details>
</section>
""")

# ================================================================ slide 6
g = grow
replace_slide(6, f"""
<section class="slide-wrap" id="slide-6">
  <article class="slide" aria-label="Slide 6 of 6"><div class="frame">
    <header>
      <div class="kicker">Sources · Assumptions · Growth · Commitment</div>
      <h2>What's verified, what's assumed, how it grows</h2>
    </header>
    <div class="s-body">
      <div class="s6">
        <div class="card green"><h3>Verified · public sources</h3><ul><li>Wallets and PnL posts (Solscan, X)</li><li>Discord, Telegram and meetup sizes</li><li>OKX DEX fees, referral tiers, no perp DEX</li><li>OKX DEX fees, Oct 2026 (DefiLlama)</li><li>Rival rates: 2 confirmed, Axiom not</li><li>POJK 6/2026, Law 4/2026 (OJK)</li></ul></div>
        <div class="card amber"><h3>Assumed · the pilot tests these</h3><ul><li>FTDs = reach × 2% × 15%</li><li>$2K of swaps a month per FTD</li><li>FnF at $16.4K a month (none: {k(nofnf)})</li><li>OKX legal clears promotion in Indonesia</li><li>Revenue at the brief's 0.04%</li></ul></div>
        <div class="card blue"><h3>How it grows (0.04%, before shares)<sup>1</sup></h3><ul><li>Audience only: {k(g[0][0])} → {usd(g[0][1])}</li><li>Plus FnF by day 4: {m(g[1][0])} → {usd(g[1][1])}</li><li>10 FnF at the top 3: {m(g[2][0])} → {usd(g[2][1])}</li><li>Plus Partner B: {m(g[3][0])} → {usd(g[3][1])}</li><li>$10K = $25M: a month 4–6 goal</li></ul></div>
        <div class="card white wide"><h3>How I used AI</h3><p>Claude searched public sources, screened 100+ Indonesian prospects, pulled figures, built the model and drafted wording. I re-checked: <span class="ph">[what you verified yourself]</span>. Public information only; no one was contacted.</p></div>
        <div class="card blue wide"><h3>Commitment</h3><ul><li>Earliest start: <span class="ph">[date]</span></li><li>Full-time: <span class="ph">[yes / from date]</span></li><li>Other commitments: <span class="ph">[none / list]</span></li><li>In person: <span class="ph">[yes: Jakarta, Bali, others as needed]</span></li></ul></div>
      </div>
    </div>
    <p class="foot">Sources: OKX DEX fee schedule, May 2026 fee notice and referral programme (web3.okx.com); DefiLlama; Solscan; Discord; Luma; ICN; OJK (POJK 6/2026); X posts listed in the spreadsheet · <sup>1</sup> Each partner's first month (base); steps and inputs in the Miscellaneous tab, sections A and E</p>
  </div></article>
  <details class="notes">
    <summary>Speaker notes · slide 6 (includes source links)</summary>
    <div class="body">
      <p>How it grows, step by step, at the same 0.04% before partner shares (each partner's first month, base case): the audience alone, sized the research sheet's way (reach × 2% × 15% × $2,000), gives {k(g[0][0])} of swaps ({usd(g[0][1])}); onboarding FnF power traders first, by day 4, gives {m(g[1][0])} ({usd(g[1][1])}), because they carry {fnf_share:.0%} of swaps; asking the top 3 (Bevan's Masterclass, 0xRusdi's referral book, Ponyin's circle) for 10 FnF power traders instead of 5 gives {m(g[2][0])} ({usd(g[2][1])}); adding the Partner B pilot ($2M a month) gives {m(g[3][0])} ({usd(g[3][1])}). $10K a month needs $25M of swaps: A and C on revenue share only if their 90-day data holds ($10M), plus more partners at this profile, a month 4-6 goal rather than month 1. The downside is just as clear: at the research's low rates (1% sign up, 6.1% trade, $286 each, FnF trading like an average FTD) the 8 bring about {k(T['low'])} ({usd(low_rev)}), so FnF intros are the first thing we test on every call.</p>
      <p>Why "OKX legal clears promotion" is on the assumed list: Law 4/2026 (in force 17 June 2026) lets OJK suspend or block non-compliant crypto activity, including offshore platforms serving Indonesians, and POJK 6/2026 makes firms answerable for their marketing partners' content. I found no OJK licence for OKX on public lists. OKX Wallet and DEX are non-custodial, which may be treated differently, but that is legal's call, not mine, and I'd ask for it in week 1.</p>
      <p>Source links. OKX DEX fees: web3.okx.com/dex-fees; fee notice effective 12 May 2026: web3.okx.com/help/update-on-okx-dex-interface-fee. OKX DEX referral programme: web3.okx.com/help/whats-dex-referral-program (fee share up to 50%, 0-20% invitee discount). Meme Mode launch (Jul 2024): okx.com/learn/meme-mode. OKX DEX fees by month: DefiLlama. Bevan's published wallet: solscan.io/account/8DC3hpFtU3mEU3FUCtDfU72SmxTbrw4w35fyMYP8qGxg; Narrative Labs: narrativelabs.id. Kevin Susanto fomo PnL: x.com/KevinSusanto/status/2123472525070102987. Ponyin Trojan PnL card: x.com/Ponyin/status/2037746268039209077. GuarEmperor member results: x.com/GuarEmperor/status/2115414023114256560. Stress Capital: discord.gg/stresscapitals. LP Army Jakarta: luma.com/76fvdt2j. MET Bali: luma.com/pkr27xr4. W3AI Wednesday at JURA: luma.com/zkq937gq. OKX Alpha Traders Summit, Bali: bitpinas.com/feature/okx-alpha-traders-summit. ICN community sizes (Mar 2026): indonesiacrypto.network/blog/crypto-community-indonesia. Reach, FTD and volume rates: Indonesia Crypto KOL &amp; Partner Research sheet (Volume proof and Shortlist deep-dive tabs, X API and Discord counts, 20 Nov 2026). Indonesian X evidence (links in the spreadsheet): @0xRusdi BasedBot dashboard, @0xSarimin Trojan referrals, @AiSlimx OKX Wallet Boost, @Ponyin Trojan results, @cryptondo OKX Wallet use, @KevinSusanto fomo PnL, @GuarEmperor member results. Rival referral rates: pump.fun (PrimeXBT, Crypto Briefing), Axiom (affiliate guides, unconfirmed), fomo (Datawallet). POJK 6/2026 finfluencer rules: ojk.go.id; coverage: okezone.com (8 Jul 2026). Law 4/2026: hsfkramer.com (2026 omnibus law note). Akademi Crypto inquiry halted: justisio.com (26 Oct 2026). Assumptions change in the spreadsheet's yellow cells and every estimate recalculates. Fill the bracketed fields before sending.</p>
    </div>
  </details>
</section>
""")

assert "—" not in s, "em dash"
deck.write_text(s)
print("deck written; swaps", T["swaps"], "rev", T["rev"])
