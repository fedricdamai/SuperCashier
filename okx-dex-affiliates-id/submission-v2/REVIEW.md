# Second opinion: OKX DEX BD Affiliate (Indonesia) submission

**Verdict:** the strategy is sharper than mine and the maths is clean. Every value in the review brief's section 6 reproduces, and the workbook recalculates with 0 errors across 626 formulas. Ten things needed fixing before sending: mostly labels and evidence framing, not numbers. One real gap: Indonesian compliance risk barely appears, and a panel in Jakarta will ask about it.

Files in this folder:

| File | What it is |
|---|---|
| `deck_slides_v2.md` | Your 6 slides with every fix below applied |
| `OKX_prospect_spreadsheet_v2.xlsx` | Your workbook: same model, same numbers, 20 text cells fixed, a `Longlist (100 screened)` tab, and live formulas for every new deck figure (Task 2 rows 16-20, Partner plan row 49, 30-day plan rows 33 and 44) |
| `REVIEW.md` | This file |

## 1. What I checked

| Check | Result |
|---|---|
| Task 2: A, B, C at their terms and ours, break-evens, 50% cases | All correct (A $720 short, B +$520, C $840 short; $6.25M and $8.33M break-even) |
| Partner plan: FTDs, retail and power swaps, trades a day, scenarios | All 8 rows and totals reproduce ($1,768,690; $707.48; 90.3% power share) |
| 30-day calendar ramp | Day 14, 21 and 30 totals reproduce ($741,788; $65,139 a day; $3.95M run-rate) |
| Ranking scores | All 8 weighted scores reproduce (4.30 to 2.75) |
| Growth steps and published-fee upside | Reproduce ($1.12M, $2.45M, $4.45M; $7,748) |
| Product claims | Meme Mode is real (launched Jul 2024). The 0.1 / 0.25 / 0.5% tiers match OKX's fee notice effective 12 May 2026 |
| Word counts | First message 116 after the edit below (limit 120) |

## 2. Errors: fix before sending (all fixed in v2 unless marked)

| # | Where | What's wrong | Fix |
|---|---|---|---|
| E1 | Slide 4, card "#1's first 30 days" | "OKX keeps $832, or $95 at the brief's 0.04%" mixes bases. $832 is after his 30% share; $95 is before it. At 0.04% OKX keeps $66.53 (Partner plan P26) | "OKX keeps $832 at the published fee, or $67 at the brief's 0.04% after his share" |
| E2 | Slide 2 "Verified proof"; Shortlist F8, B2; Partner plan W26, E54 | "$3.4M+ referred" sits under *verified*, but it's derived: it assumes BasedBot's fee is 1% and all $8,463 was earned at 25%. Only the $8,463 is verified. Also missing its "$" in F8 and B2 | Slide shows "$8.5K BasedBot referral earnings". The $3.4M moves to notes and to Shortlist K8 (assumptions) as an estimate |
| E3 | Slide 2 footer ² | "FTDs = 1% of post readers ... + 5 power traders each" says FTDs include power traders. They don't: 0xRusdi's 18 = 1% × 1,800 | "each partner's 5 power traders are counted separately" |
| E4 | Brief checklist D8; Shortlist B5; slide 2 notes | "8 partner types" overclaims. Bot referrer, trench creator, trader KOL and promo creator are all X creators. A reviewer will count them as one type | "4 partner types by the brief's categories": X trading creators, a trading clan, event-led onchain communities, a Telegram call channel |
| E5 | Task 3 B7 (objection reply) and B11; slide 3 footer ², slide 4 footer ³ | "shares 20-50% of the swap fee by tier" is told to a prospect as fact. OKX's own regional help pages disagree on the tier table (one lists tiers starting at 10%). The brief bans unverified terms in Task 3 | "up to 50% of the swap fee, depending on tier" (true on every page). Internal tabs still say 20-50%; fine with a "confirm" note |
| E6 | Slide 2, slide 5, slide 6: every revenue figure | Revenue is shown before partner payouts, while the brief's contribution formula subtracts them. Slide 6's "$1,780" includes B's $800 of revenue; slide 3 shows B's contribution as $600. A reviewer comparing the two will stop there | Label every figure "before partner share" and add the after-share number: $707 → $508 (slide 2), $830 → $612 (slide 5) |
| E7 | Slide 4 table | Mixes job days (1-7) and activation days ("Days 8-21 are the 14-day activation plan") while the targets use activation days 4, 7, 10 and 14 | Rows after agreement now read A1-A2, A2-A4, A5-A7, A8-A14 (A1 = agreement, about job day 10) |
| E8 | Slide 6, Verified card | "Rival referral rates (3 tools)" is listed as verified, but Offers & support marks Axiom as unconfirmed | "Rival rates: 2 confirmed, Axiom not" |
| E9 | Partner plan T25 (**not changed**) | Hyperliquid ID's audience is 100, but the only public figure is ~62 going on its sold-out Luma meetup (luma.com/lmx2s2ut, checked Nov 2026). 100 sits above the evidence, unlike Meteora's 300, which sits below its 462 | v2 relabels W25 and the slide notes: "100 assumes a bigger co-hosted room". If you'd rather match the evidence, set T25 to 60: FTDs 263.4 → 255.4, swaps $1,768,690 → $1,764,210, revenue $707 → $706, day-30 calendar swaps $741,788 → $740,828. I left the model alone so the deck and sheet still tie |
| E10 | Partner plan E13 (**not changed**) | "@AiSlimx (~$8K a month, OKX Wallet Boost)" vs slide 1's "~$550 a day". $550 × 30 = ~$16.5K | State the active days behind $8K, or use one figure in both places. I didn't know which, so I left it |

## 3. Suggestions: worth doing (applied in v2 unless marked)

**S1. Add a compliance gate. This is the biggest gap.** Two 2026 rules matter to this exact job:
- **POJK 6/2026** requires certification for anyone publishing crypto information, and makes the firm answerable for its marketing partners' content.
- **Law 4/2026** (in force 17 Jun 2026) lets OJK suspend or block non-compliant crypto activity, including offshore platforms serving Indonesians. I found no OJK licence for OKX on public lists.

Your deck lists compliance only as a cost. v2 adds:
- "No paid post before certification" (slide 5)
- "OKX legal clears promotion" as an assumption (slide 6)
- One line on slide 4 notes: power-trader-first sequencing is also compliance-resilient. 1:1 private onboarding isn't a public post, so ~90% of swaps don't wait on certification. Confirm that with legal.

Sources: [Okezone, POJK 6/2026](https://economy.okezone.com/read/2026/07/08/320/3228975/ojk-wajibkan-sertifikasi-influencer-kripto-ini-tujuannya), [HSF Kramer, Law 4/2026](https://www.hsfkramer.com/ru/notes/indonesia/2026-posts/indonesia-reshapes-digital-asset-and-cryptocurrency-framework-with-2026-financial-services-omnibus-law).

**S2. Show the floor.** Without power traders, the 8 bring $171,640 of swaps ($69). Saying this yourself turns the model's biggest assumption into a deliberate first-call gate. Added to the slide 2 and slide 6 notes, plus the Assumed card.

**S3. Show why the big names lost.** My 100-prospect screen found Indonesia's largest crypto creators teach CEX, futures or forex trading with no onchain proof:
- Andy Senjaya, ~695K YouTube ([socialcounts](https://socialcounts.org/youtube-live-subscriber-count/UCF7JTH4DJibstTeJEjZe5DA))
- Felicia Putri Tjiasaka, ~1.3M followers

The largest paid academy, Akademi Crypto, is on legal hold: police halted the inquiry in Oct 2026 and a special review was requested ([Justisio](https://justisio.com/berita/2026-10-26-polda-metro-hentikan-penyelidikan-laporan-dugaan-penipuan-kripto-timot)).

One line on slide 2 plus the new Longlist tab turn "we picked 8" into "we screened 100 and these 8 survived". That speaks straight to the 30% prospect-quality weight.

**S4. Make the in-person day impossible to empty.** It moves to a **Wednesday**, because JURA Kemanggisan runs W3AI Wednesday every week ([Luma](https://luma.com/zkq937gq)). Other changes:
- 16:00 is now Superteam Hub Jakarta, run with Indonesia Crypto Hub ([Luma](https://luma.com/2fpz2q5p)).
- Hyperliquid ID's hosts overlap with Web3 Week Asia's organisers. Alex Rumondor co-founded it ([detik](https://inet.detik.com/cyberlife/d-8236202/web3-week-asia-cetak-rekor-baru-indonesia-disebut-power-house-blockchain)), and Yosia Sophia hosted its UI edition, so the 10:00 coffee reaches two communities.

Updated on slide 5 and in 30-day plan A50/B50.

**S5. Use OKX's own network for warm intros.** OKX already has a footprint here:
- It ran the Alpha Traders Summit in Bali, Jan 2025, with 200+ traders, affiliates and community leaders ([BitPinas](https://bitpinas.com/feature/okx-alpha-traders-summit/)).
- It hosted Alphas Connect Bali with Lisk, Sui and IDNFT.
- It sponsored Coinfest 2024.

This doesn't break your "no existing OKX partners" rule, because they become intro sources, not the 8. It's the strongest answer to "20 messages, 0 meetings". Updated on slide 5 and in 30-day plan F43.

**S6. Answer "the chosen partner's economics" from both sides.** The brief's wording can mean B's economics too. B's payout halves from $200 to $100 on our terms, and from $280 to $140 on its ask. Also note that +$600 becomes +$520 if the step-up hits in month 1. Added to slide 3.

**S7. Cite the fee notice, not just the schedule.** OKX's notice effective 12 May 2026 raised Group 1 or 2 against other tokens to 0.5%. Swaps between two "other" tokens stay at 0%. Meme Mode launched fee-free in 2024 ([OKX notice](https://web3.okx.com/help/update-on-okx-dex-interface-fee), [Meme Mode launch](https://www.fintechfutures.com/press-releases/okx-wallet-launches-meme-mode-for-instant-self-custodial-memecoin-swaps)). Added to slide 1.

**S8. Name the condition that would change your Hyperliquid call.** Hyperliquid lets wallets add a "builder code" fee; Phantom has earned ~$20.6M that way ([Solana Compass](https://solanacompass.com/news/phantom-tops-hyperliquids-builder-program-with-206m-in-fees)). I found no OKX builder code. If OKX adds one, Hyperliquid ID jumps. This is a ready answer for "adapt when assumptions change". Added to slide 1 notes.

**S9. Plan for a top-3 "no".** Week 1 has 3 meetings and wave 1 needs all 3 to sign. v2 adds: "if one says no, TwipsX moves into wave 1" (slide 5 notes).

**S10. Tone down the title.** "Leading DEX partnerships for OKX" overstates an affiliate-BD role. Now "working on DEX partnerships for OKX" (Task 3 B4; Bahasa B5 "menangani"). Word count is 116.

**S11 (not changed): Products & revenue, Meme Mode row.** It lists "limit orders". OKX's own explainer called limit orders "coming soon" at the time. Confirm before you say it on a call.

**S12 (not changed): Offers & support, section B.** Items like "24h support chat", "wallet-level dashboard" and "alerts built on the DEX API" read as commitments. Some already say "if OKX agrees"; tag the rest "proposed, needs OKX sign-off".

**S13: Small fixes, all applied:**
- Slide 2 notes now match the 30-day plan tab ("under $10"; Stress Capital added).
- Meteora's proof now names both events (462 at LP Army Jakarta, 671 at MET Bali) instead of implying both were LP Army.

## 4. Where my research and yours disagreed

- **I ranked Hyperliquid ID #1; you ranked it Low. You're right.** Perps opened through OKX Wallet pay Hyperliquid, so only bridge and spot swaps count. My scoring rewarded trading intent without asking what OKX actually earns.
- **My top 8 were mostly big education creators. Your filter was better.** Proof of onchain trading plus rival flow to move picks partners who bring volume OKX can bill.
- **The lists overlap on one name (Hyperliquid ID), and that's useful.** Yours is the volume layer: who trades. Mine is the access layer: venues, events, organisers, compliance channels, and a month-2 bench (Andy Senjaya, Michael Wyann, Hendy Tanuardy, Belvin Tannadi for the CEX-beginner and xStocks tests). v2 keeps your 8 and uses mine as support.

## 5. Questions the panel is likely to ask

1. **"Your 8 real partners make $707 in a month. Hypothetical B alone makes $800. Why hire for this?"**
   Month 1 buys leading indicators: 263 FTDs, ~800 trades a day, 3 activated partners. The same swaps earn ~$7.7K at OKX's published fees. At the brief's rate, the path to $10K a month is month 4-6.
2. **"What if partners don't bring power traders?"**
   Then it's $172K and $69. That's why 5 intros are asked for on the first call, and a partner that can't provide them drops to a later wave.
3. **"Can OKX legally pay Indonesian KOLs right now?"**
   Not assumed. Legal sign-off comes in week 1. Private power-trader onboarding goes first, and public posts wait for certification.
4. **"Why cut B from 35% to 25%?"**
   We don't. We tie their 35% to proof: 50 wallets trading on 3+ days.
5. **"Why spend your best in-person slot on a Low-priority partner?"**
   It's the confirmed meeting. The hosts also run Web3 Week Asia. Bridge flows into Hyperliquid route through OKX. And the brief requires partners you can meet in person.

## 6. Still yours to fill

- Slide 6 commitment card: start date, full-time, other commitments, in person.
- Slide 6 AI card: I drafted the first sentence. "I re-checked: [ ]" has to be what you checked yourself. Re-opening the 8 partner links and the OKX fee and referral pages would cover it.
- Whether to apply E9 (Hyperliquid 100 → 60) and E10 (AiSlimx $8K).
