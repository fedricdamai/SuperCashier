"""Build the OKX DEX Indonesia affiliate prospect list (100 rows, scored).

Run:  python build_prospects.py
Out:  prospects.xlsx (live formulas) and prospects.csv (values, same math)

Scores are 1-5 inputs. The weighted score and priority are formulas in the
workbook and are recomputed here only to write the CSV.
"""
import csv
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

OUT = Path(__file__).parent

WEIGHTS = {"V": 0.30, "T": 0.25, "O": 0.15, "R": 0.15, "K": 0.15}
HIGH, MED = 68, 52

# ---------------------------------------------------------------- sources
S = {
    "socialcounts_andy": "https://socialcounts.org/youtube-live-subscriber-count/UCF7JTH4DJibstTeJEjZe5DA",
    "heepsy_yt": "https://www.heepsy.com/top-youtube/crypto/indonesia",
    "heepsy_tt": "https://www.heepsy.com/en/find-influencers/tiktok/crypto/indonesia",
    "heepsy_tt_es": "https://www.heepsy.com/es/top-tiktok/crypto/indonesia",
    "heepsy_tt_2025": "https://www.heepsy.com/top-tiktok/crypto/indonesia",
    "heepsy_ig": "https://www.heepsy.com/top-instagram/crypto/indonesia",
    "cryptonews_kol": "https://cryptonews.com/id/news/best-crypto-influencer-accounts/",
    "ajmarketing": "https://www.ajmarketing.io/feed/top-11-crypto-influencers-in-indonesia",
    "ajaib_kol": "https://ajaib.co.id/influencer-crypto-daftar-wajib-follow",
    "bybit_id": "https://cointelegraph.com/news/crypto-exchange-bybit-launches-in-indonesia-after-nobi-acquisition",
    "btse_id": "https://crypto.news/btse-debuts-regulated-crypto-trading-platform-in-indonesia/",
    "pintu_web3": "https://coindesk.com/business/2024/05/29/crypto-wallet-firm-fordefi-expands-in-indonesia-powering-pintu-web3-self-custodial-wallet",
    "icex": "https://www.dealstreetasia.com/stories/indonesia-crypto-exchange-launches-integrated-platform-under-ojk-framework-477889",
    "dupoin": "https://www.dupoin.co.id/insights/market-analysis/79167/",
    "idx_wyann": "https://indodax.com/academy/michael-wyann-bitcoin-tren-kripto/",
    "idx_hendy": "https://indodax.com/academy/hendy-tanuardy",
    "idx_kevin": "https://indodax.com/academy/kevin-sailly-narasi-kripto/",
    "idx_angga": "https://indodax.com/academy/?p=71248",
    "idx_andreas": "https://indodax.com/academy/andreas-tobing-influencer-kripto-gen-z/",
    "idx_felicia": "https://indodax.com/academy/felicia-putri-tjiasaka-top-influencer-crypto/",
    "idx_samuel": "https://indodax.com/academy/?p=71386",
    "idx_belvin": "https://indodax.com/academy/?p=71321",
    "idx_solil": "https://indodax.com/academy/profil-soliloquizephar/",
    "idx_joko": "https://indodax.com/academy/joko-crypto-ekosistem-kripto-indonesia/",
    "idx_jonas": "https://indodax.com/academy/jonas-sunandar-cmif/",
    "idx_komunitas": "https://indodax.com/academy/komunitas-trading-kripto-indonesia/",
    "idx_komunitas2": "https://indodax.com/academy/komunitas-kripto-indonesia-nambah-wawasan/",
    "cmif2025": "https://icn-preview.vercel.app/blog/coinvestasi-most-impactful-figures-2025-recognizing-the-leaders-shaping-indonesias-crypto-industry",
    "liputan_cmif2024": "https://www.liputan6.com/tekno/read/5864272/12-tokoh-paling-berpengaruh-di-industri-kripto-dan-web3-indonesia-pada-2024",
    "akademi_profile": "https://www.investortrust.id/financial/91043/profil-akademi-crypto-dan-pusaran-kasus-hukum-timothy-ronald",
    "akademi_halt": "https://justisio.com/berita/2026-10-26-polda-metro-hentikan-penyelidikan-laporan-dugaan-penipuan-kripto-timot",
    "akademi_jan": "https://kabar24.bisnis.com/read/20260112/16/1943517/polisi-selidiki-kasus-penipuan-trading-kripto-yang-diduga-terkait-timothy-ronald",
    "akademi_may": "https://www.rctiplus.com/news/detail/berita-utama/5378710/polisi-periksa-timothy-ronald-dan-kalimasada-di-kasus-dugaan-penipuan-investasi-kripto",
    "ternak": "https://teknologi.id/business/kisah-sukses-raymond-chin-dalam-mendirikan-platform-investasi-online-ternak-uang",
    "linkedin_edwin": "https://id.linkedin.com/in/edwinnazir",
    "discord_ci": "https://discord.me/cryptoindonesia",
    "yellow_signals": "https://yellow.com/id/news/9-grup-telegram-terbaik-untuk-wawasan-perdagangan-crypto",
    "hl_meetup": "https://luma.com/lmx2s2ut",
    "hl_indodax": "https://www.coingabbar.com/en/crypto-listings/hyperliquid-hype-crypto-listing-on-indodax-exchange",
    "pump80": "https://pluang.com/en/news-feed/aktivitas-pumpfun-turun-80-persen-trader-tinggalkan-memecoin-drag-fee-solana",
    "xenorize": "https://luma.com/91lre04f",
    "btc_id_meetups": "https://luma.com/u0qik6v3",
    "btc_house_bali": "https://www.eventbrite.com/e/bitcoin-meetup-at-the-bitcoin-house-bali-tickets-859907956847",
    "bitdevs": "https://www.meetup.com/bitcoin-builders-association-indonesia/events/312784975/",
    "stacks": "https://www.meetup.com/stacks-indonesia/events/308499540",
    "yap_panel": "https://lu.ma/jfmaptpi",
    "coinfest2019": "https://www.liputan6.com/tekno/read/3910141/coinfest-2019-edukasi-masyarakat-soal-bitcoin-dan-aset-digital",
    "cryptosiast_ws": "https://luma.com/yc2bac9g",
    "id_web3_hack": "https://luma.com/pcc699dv",
    "superteam_showcase": "https://luma.com/7gbaanxq",
    "superteam_hub": "https://luma.com/2fpz2q5p",
    "solana_village": "https://luma.com/ia8qk9nx",
    "superteam_hackerhouse": "https://luma.com/pmsbnaog",
    "ethjkt": "https://luma.com/ibtz3tvm",
    "eth_genesis": "https://vercel.lu.ma/jr95ihyx",
    "blockdevid": "https://id.linkedin.com/company/blockdevid",
    "base_ws": "https://luma.com/s3x75do7",
    "base_bukber": "https://lu.ma/3d5g9g68",
    "lisk_ignit3": "https://cryptoslate.com/press-releases/lisk-and-gerakan-nasional-1000-startup-digital-unveil-ignit3/",
    "okx_alphas_bali": "https://luma.com/81m8jgct",
    "monad_blitz": "https://luma.com/ksxks0qo",
    "monad_101": "https://lu.ma/tx8jk6fg",
    "sui_bootcamp": "https://luma.com/uog22okf",
    "near_legion": "https://luma.com/cqjjxsy8",
    "odx": "https://lu.ma/ctbz3h2j",
    "jura": "https://luma.com/zkq937gq",
    "eth_meetup_pintu": "https://voi.id/en/amp/376438",
    "idrx_cg": "https://coingecko.com/en/coins/idrx",
    "idrx_50m": "https://investortrust.id/market/70706/transaksi-idrx-capai-rp-816-miliar-proyekstablecoin-berbasis-rupiah-kian-diminati",
    "xellar": "https://mwm.ai/apps/xellar-crypto-wallet/1671215861",
    "satu": "https://crypto.news/satu-wallet-launches-all-in-one-crypto-super-app-with-satu-token-on-solana/",
    "bisa": "https://bisa-wallet.vercel.app/",
    "tokenomy": "https://monotykamary.github.io/consulting/case-study/tokenomy",
    "blockvest": "https://inet.detik.com/business/d-7123103/mengenal-blockvest-platform-pemersatu-industri-kripto-di-indonesia",
    "cryptowave": "https://investortrust.id/market/72136/literasi-aset-digital-meningkat-cryptowave-disambut-positif-pengguna",
    "coinvestasi_partner": "https://coinvestasi.com/partner",
    "coinvestasi_bisnis": "https://market.bisnis.com/read/20210727/94/1422410/coinvestasicom-tegaskan-sebagai-portal-informasi-resmi-blockchain-dan-kripto",
    "coindesk_id": "https://www.coindesk.com/id/about",
    "bic_cfx": "https://beincrypto.com/beincrypto-indonesia-cfx-crypto-journalism-partnership",
    "cryptonews_about": "https://cryptonews.com/id/about-us/",
    "bcm_tg": "https://t.me/s/indo/110",
    "bitcoinist_id": "https://bitcoinist.com/daftar-akun-yang-menawarkan-crypto-signal-terbaik/",
    "coinspeaker_id": "https://www.coinspeaker.com/id/3-meme-coin-solana-bisa-naik-1000x-2/",
    "investortrust": "https://investortrust.id/",
    "coinfest2026": "https://bitpinas.com/feature/coinfest-asia-2026-scheduled/",
    "coinfest2025": "https://investortrust.id/market/77065/10-ribu-orang-padati-coinfest-asia-2025-di-bali",
    "okx_coinfest": "https://www.fintechfutures.com/press-releases/okx-sponsors-coinfest-asia-2024-to-showcase-web3-innovations",
    "idbw": "https://coingape.com/press-releases/indonesia-blockchain-week-idbw-2026-flags-rwa-tokenization-and-high-level-enterprise-networking-as-major-business-priorities/",
    "web3weekasia": "https://inet.detik.com/cyberlife/d-8236202/web3-week-asia-cetak-rekor-baru-indonesia-disebut-power-house-blockchain",
    "abi": "https://investortrust.id/market/107497/abi-dorong-literasi-kripto-untuk-perkuat-perlindungan-konsumen-digital",
    "aftech": "https://investortrust.id/financial/120152/pojk-62026-berlaku-aftech-dan-iarfc-kembangkan-program-sertifikasi-verified-finfluencer",
    "pojk6": "https://economy.okezone.com/read/2026/07/08/320/3228975/ojk-wajibkan-sertifikasi-influencer-kripto-ini-tujuannya",
    "ekraf_seabw": "https://ekraf.go.id/news-en/indonesia-supports-national-web3-talent-through-protocol-camp-and-southeast-asia-blockchain-week-2026",
    "bali_campus": "https://luma.com/ndwylksj",
    "blk2026": "https://investortrust.id/financial/99254/lewat-bulan-literasi-kripto-2026-ojk-dan-abi-fokus-edukasi-risiko",
    "okx_alpha_summit": "https://bitpinas.com/feature/okx-alpha-traders-summit/",
    "metamask_jkt": "https://luma.com/cbnjakartajuly2026",
    "cto_event": "https://luma.com/f1v1scbg",
    "binus": "https://student-activity.binus.ac.id/bbcc",
    "petra": "https://investortrust.id/market/102849/upbit-indonesia-perkuat-literasi-blockchain-lewat-web3-education-center-di-universitas-kristen-petra",
    "ui_feb": "https://lu.ma/yi6mwz6n",
    "yescon": "https://luma.com/usnq6w2q",
    "upj": "https://www.investortrust.id/market/87254/pintu-dan-blockvest-perkuat-literasi-ai-dan-blockchain-bagi-mahasiswa-universitas-pembangunan-jaya",
    "airdrop_finder": "https://ejournal.up45.ac.id/index.php/JHCJ/article/download/2183/1363/9048",
    "techdropper": "https://t.me/s/techdropper/9098",
    "trackingdrops": "https://t.me/s/trackingdrops/5916",
    "cnesia": "https://t.me/s/cnesia112/20674",
    "stockbit": "https://www.jawapos.com/finance/2603070331/masyarakat-indonesia-makin-melek-investasi-transaksi-saham-di-stockbit-melonjak-dari-rp-17-triliun-jadi-rp-107-triliun-dalam-setahun",
}


def src(*keys):
    return " | ".join(S[k] for k in keys)


# ---------------------------------------------------------------- prospects
# scores = (V volume, T trading intent, O onchain readiness, R reachability, K risk safety)
ROWS = []


def P(name, ptype, link, base, inperson, airdrop, audience, evidence,
      assumptions, route, proposal, risk, scores, ver, reason, sources, flag=""):
    ROWS.append(dict(name=name, type=ptype, link=link, base=base,
                     inperson=inperson, airdrop=airdrop, audience=audience,
                     evidence=evidence, assumptions=assumptions, route=route,
                     proposal=proposal, risk=risk, scores=scores, ver=ver,
                     reason=reason, sources=sources, flag=flag))


CERT = "Must hold OJK finfluencer certification (POJK 6/2026) before any paid content."

# ---- Trading educators and creators
P("Andy Senjaya", "Trading educator", "https://www.youtube.com/channel/UCF7JTH4DJibstTeJEjZe5DA",
  "Indonesia (city not confirmed)", "Likely", "N",
  "Indonesian retail traders learning technical analysis across crypto, stocks and forex.",
  "~695K YouTube subscribers and 227 videos (socialcounts, Jul 2026). Heepsy lists him as a technical trader covering crypto, stocks and forex; ~133-144K on TikTok.",
  "Audience trades mostly on CEX futures and forex brokers, not DEX. May hold existing broker or exchange sponsorships.",
  "Business email on YouTube About page; follow up by Instagram DM; ask for a 30-minute call.",
  "Co-produced 'onchain trading for TA traders' series (OKX Wallet + DEX: slippage, limit orders, gas) with his own DEX referral code. Rev share on invitees' UI fees per OKX DEX Referral program; any retainer only after a 30-day volume read.",
  CERT + " Forex/stock viewers may not convert to onchain; possible exchange exclusivity.",
  (5, 5, 2, 3, 3), "Partial",
  "Largest trader-education audience with a third-party count; content is trading, not giveaways.",
  src("socialcounts_andy", "heepsy_yt", "heepsy_tt"), "Top 3")

P("Michael Wyann", "Trading educator", S["idx_wyann"], "Indonesia", "Likely", "N",
  "Crypto-only YouTube audience following Bitcoin trends and chart analysis.",
  "Indodax Academy profiles him as an independent crypto voice. Cryptonews Indonesia's 2026 creator list puts the channel near 100K subscribers with frequent uploads.",
  "Subscriber count and cadence come from a commercial listicle, not checked on YouTube. No exclusive exchange deal assumed.",
  "YouTube About business contact; X DM; intro via Coinvestasi community team.",
  "Monthly onchain market walkthrough using OKX DEX and Wallet tools; custom referral code; rev share. Pilot: one video plus one live session.",
  CERT + " An independent stance may mean he declines paid promotion.",
  (4, 5, 3, 3, 4), "Partial",
  "Crypto-only trading content; mid-size channel is easier to close and measure inside 30 days.",
  src("idx_wyann", "cryptonews_kol"), "Top 8")

P("Hendy Tanuardy", "Trading educator", S["idx_hendy"], "Indonesia", "Likely", "N",
  "Young traders learning strategy, risk management and trader psychology.",
  "Indodax Academy profile: content centres on trading strategy, risk management and trader psychology, with a technical-analysis focus.",
  "Audience size not verified in this pass. Viewers mostly trade on CEX.",
  "Instagram or TikTok DM, then business email.",
  "Risk-managed onchain trading module (position sizing, slippage, MEV protection) with referral code; rev share.",
  CERT + " Audience size unverified.",
  (3, 5, 2, 3, 4), "Partial",
  "Discipline-focused content attracts repeat traders, the retention signal we need.",
  src("idx_hendy"), "Top 8")

P("Kevin Sailly", "Analytics & news creator", S["idx_kevin"], "Indonesia", "Likely", "N",
  "Daily crypto news followers and swing traders.",
  "Indodax profile: mixes technical analysis with project narratives. Cryptonews ID 2026 list: daily-news channel, 15-20 minute videos, self-described trader and market analyst.",
  "Audience size unknown.",
  "YouTube business email; X DM.",
  "Sponsored 'token of the week onchain' segment with a live OKX DEX swap and referral code.",
  CERT + " News viewers may watch more than they trade.",
  (3, 4, 2, 3, 4), "Partial",
  "Daily touchpoint with traders; conversion less certain than pure TA educators.",
  src("idx_kevin", "cryptonews_kol"))

P("Angga Andinata / BelajarCrypto", "Trading educator", "https://www.youtube.com/channel/UCbkKwnkYEY_5JCMPU9IEGxg",
  "Indonesia", "Likely", "N",
  "Beginner long-term crypto investors.",
  "Indodax profile: founder of BelajarCrypto; long-term education approach, explicitly not trading signals.",
  "Audience buys and holds on CEX.",
  "BelajarCrypto contact page; YouTube business email.",
  "Self-custody and wallet-safety series featuring OKX Wallet; referral code for a first onchain swap.",
  "Buy-and-hold audience creates little repeat volume.",
  (3, 2, 2, 3, 4), "Partial",
  "Credible educator, but holders do not trade often.",
  src("idx_angga"))

P("Andreas Tobing", "Trading educator", S["idx_andreas"], "Jakarta", "Y", "N",
  "Gen Z crypto learners on TikTok, YouTube and Instagram; podcast listeners; campus audiences.",
  "TikTok @andreasrtobing 143K+ followers (Indodax, Jul 2026). Co-founder of Web3 Week Asia (Jakarta, Nov 2025). CMIF 2025 honorable mention as founder of Banteng Lari Podcast.",
  "Content is education-first, not active trading.",
  "Meet at Web3 Week Asia or Jakarta community events; TikTok or Instagram DM.",
  "Podcast episode plus campus segment on onchain trading basics; referral code; co-hosted Jakarta meetup.",
  CERT + " Education audience converts slowly.",
  (3, 3, 3, 5, 4), "Verified",
  "Meetable in person and runs his own events; moderate trading intent.",
  src("idx_andreas", "web3weekasia", "cmif2025"), "Top 8")

P("Felicia Putri Tjiasaka", "Finance creator", S["idx_felicia"], "Jakarta (assumed)", "Likely", "N",
  "Gen Z beginner investors across stocks, funds and crypto.",
  "~1.3M followers (Ajaib list). Co-founder and CPO of Ternak Uang. Indodax notes no official Indodax collaboration yet.",
  "Small share of audience uses self-custody wallets.",
  "Talent management via business email on profile.",
  "One self-custody explainer at most; not a volume partner.",
  "Beginner mass audience, high fee expectations. Reputational check needed: a Ternak Uang co-founder was subject of a 2026 police inquiry, since halted.",
  (4, 2, 1, 2, 2), "Verified",
  "Reach without trading intent.",
  src("idx_felicia", "ajaib_kol", "ternak"))

P("Samuel Christ", "Finance creator", S["idx_samuel"], "Indonesia", "Likely", "N",
  "General finance and business viewers.",
  "Cryptonews ID 2026 list: channel covers finance and business generally, not crypto specifically. Indodax lists him among 2025 crypto influencers.",
  "Low onchain share.",
  "YouTube business email.",
  "Awareness only.",
  "Off-target audience.",
  (2, 2, 1, 3, 4), "Partial",
  "General finance audience.",
  src("idx_samuel", "cryptonews_kol"))

P("Belvin Tannadi", "Trading educator", S["idx_belvin"], "Indonesia", "Likely", "N",
  "Stock traders moving into crypto.",
  "Indodax Academy profile frames him as a bridge between stock and crypto knowledge.",
  "Audience trades stocks actively but is new to wallets.",
  "Instagram DM or business email.",
  "'From IDX to onchain' series with referral code.",
  CERT + " Stock audience lacks self-custody habits.",
  (3, 4, 2, 3, 4), "Partial",
  "Active traders, but a wallet onboarding gap.",
  src("idx_belvin"))

P("Andre Rizky", "Trading educator", S["ajmarketing"], "Indonesia", "Unknown", "N",
  "Forex and crypto technical traders.",
  "Listed as a forex and crypto technical trader posting 10-minute trading tips (AJ Marketing, 2025).",
  "Current activity and audience size not checked.",
  "Instagram DM.",
  "Test one short tutorial with referral code.",
  "Likely forex broker ties; data is dated.",
  (3, 4, 2, 2, 3), "Lead",
  "Trading-led but unverified.",
  src("ajmarketing"))

P("Edwin Nazir / NQvers", "Trading community", S["linkedin_edwin"], "Indonesia", "Unknown", "N",
  "Members of the NQvers trading community.",
  "LinkedIn profile self-describes him as a trading educator and NQvers founder.",
  "Community size and asset focus unknown; the name suggests Nasdaq futures.",
  "LinkedIn message.",
  "Pilot one onchain crypto session for members.",
  "May be index futures, not crypto.",
  (2, 4, 2, 3, 3), "Lead",
  "Trader community, asset fit unconfirmed.",
  src("linkedin_edwin"))

P("Timothy Ronald / Akademi Crypto", "Trading educator", S["akademi_profile"], "Jakarta", "Y", "N",
  "Paid Discord members (Rp17M per year per Timothy Ronald) plus a large YouTube following.",
  "Founded Dec 2022 with Kalimasada; Discord-centred learning platform. Reported to Polda Metro Jaya in Jan 2026 over alleged investment fraud; police halted the inquiry in Oct 2026 after finding no criminal event; a complainant asked for a special case review.",
  "Member count not public.",
  "Hold until the special review is resolved.",
  "None during hold.",
  "Legal and reputational overhang; members allege losses.",
  (5, 5, 3, 2, 1), "Verified",
  "Highest raw volume potential on the list, blocked by the risk gate.",
  src("akademi_profile", "akademi_jan", "akademi_halt"))

P("Kalimasada", "Analytics & news creator", S["dupoin"], "Jakarta", "Y", "N",
  "Indonesian crypto traders following the Akademi Crypto co-founder.",
  "Akademi Crypto co-founder; named in the same 2026 police reports and examined in May 2026. Named among local 'Crypto Bros' KOLs (Dupoin).",
  "Personal audience size not verified.",
  "Hold.",
  "None during hold.",
  "Same legal overhang as Akademi Crypto.",
  (4, 5, 3, 2, 1), "Verified",
  "Strong trading audience, blocked by risk gate.",
  src("dupoin", "akademi_may", "akademi_halt"))

P("Gabriel Rey", "Analytics & news creator", S["dupoin"], "Indonesia", "Unknown", "N",
  "Indonesian crypto Twitter followers.",
  "Named with Timothy Ronald, Kalimasada and Suli as a local 'Crypto Bros' KOL (Dupoin).",
  "Handle, size and content not verified.",
  "X DM after verification.",
  "Referral pilot if content is trading-led.",
  "Unverified; reputational link to the Akademi cluster.",
  (3, 4, 3, 2, 3), "Lead",
  "Possible trader audience, unverified.",
  src("dupoin"))

P("Suli", "Onchain trader KOL", S["dupoin"], "Indonesia", "Unknown", "N",
  "Indonesian degen and memecoin traders.",
  "Dupoin describes Suli as the 'degen' side of the local KOL cluster: an aggressive trader hunting alpha.",
  "Trades onchain; X handle and wallet history not confirmed.",
  "Verify handle and wallet history on Solscan/Bubblemaps, then X DM.",
  "Referral code for onchain followers; rev share only, no fixed fee.",
  "Memecoin KOLs can front-run calls; identity unverified.",
  (3, 5, 5, 2, 3), "Lead",
  "Best onchain fit in the KOL cluster, but unverified so capped at Medium.",
  src("dupoin"))

P("Soliloquizephar", "Analytics & news creator", S["idx_solil"], "Indonesia", "Unknown", "N",
  "Indonesian crypto followers.",
  "Profiled by Indodax Academy as an Indonesian crypto influencer.",
  "Content focus and size unknown.",
  "Social DM.",
  "Hold until content is reviewed.",
  "Unknown fit.",
  (2, 3, 3, 2, 3), "Lead",
  "Too little evidence.",
  src("idx_solil"))

P("Joko Crypto", "Analytics & news creator", S["idx_joko"], "Indonesia", "Unknown", "N",
  "Indonesian crypto followers.",
  "Indodax Academy profile on his contribution to the local crypto ecosystem.",
  "Content focus and size unknown.",
  "Social DM.",
  "Hold until content is reviewed.",
  "Unknown fit.",
  (2, 3, 2, 2, 3), "Lead",
  "Too little evidence.",
  src("idx_joko"))

P("Sarjana Crypto", "Analytics & news creator", S["ajmarketing"], "Indonesia", "Unknown", "N",
  "Indonesian crypto Twitter readers.",
  "~111K followers on X, described as a go-to source for crypto info in Indonesia (AJ Marketing, 2025).",
  "Follower count dated; trading share unknown.",
  "X DM.",
  "Thread series on onchain execution with referral code.",
  "Dated data.",
  (3, 3, 3, 2, 3), "Lead",
  "Sizeable X reach, unverified.",
  src("ajmarketing"))

P("Jonas Sunandar", "Builder", S["idx_jonas"], "Indonesia", "Likely", "N",
  "Builders and industry peers.",
  "CMIF 2025 honorable mention as a blockchain-based builder.",
  "Not a trader-facing creator.",
  "Coinfest or IDBW networking.",
  "Integration conversation, not affiliate.",
  "Low trader reach.",
  (1, 2, 4, 3, 4), "Partial",
  "Builder, not a volume source.",
  src("idx_jonas", "cmif2025"))

P("Raymond Chin / Ternak Uang", "Finance creator", S["ternak"], "Jakarta", "Likely", "N",
  "Young Indonesians learning general investing.",
  "CEO and co-founder of Ternak Uang (with Timothy Ronald and Felicia Tjiasaka). YouTube ~1.2M, TikTok ~1.4M (Gizmologi, undated). No source describes him as a crypto educator.",
  "Audience mostly stocks and funds.",
  "Business email.",
  "Not a priority.",
  "Off-target; reputational adjacency to the Akademi case.",
  (4, 2, 1, 2, 2), "Verified",
  "Reach without onchain intent.",
  src("ternak"))

P("titikmulai", "Analytics & news creator", S["heepsy_tt_es"], "Indonesia", "Unknown", "N",
  "Indonesian crypto TikTok viewers.",
  "Ranked first on Heepsy's Indonesia crypto TikTok list (Jul 2026, ranked by relevance, not size).",
  "Follower count and content not checked.",
  "TikTok DM.",
  "Short-form onchain explainer test.",
  "Unknown audience quality.",
  (3, 3, 2, 2, 3), "Lead",
  "Relevance-ranked lead; needs a look at the content.",
  src("heepsy_tt_es"))

P("vemi.pramudya", "Analytics & news creator", S["heepsy_tt_2025"], "Indonesia", "Unknown", "N",
  "Indonesian crypto TikTok viewers.",
  "Heepsy (Apr 2025): bio 'Crypto & BlockChain Enthusiast', 0.3% engagement, ~123 average likes.",
  "Small audience.",
  "TikTok DM.",
  "Not a priority.",
  "Low reach.",
  (1, 2, 2, 2, 3), "Lead",
  "Too small.",
  src("heepsy_tt_2025"))

P("Jo SaKe | Sunny_Gaming TV", "Analytics & news creator", S["heepsy_ig"], "Indonesia", "Unknown", "N",
  "Crypto and gaming followers on Instagram.",
  "Heepsy Indonesia crypto Instagram list (Jul 2026): 17.7K followers.",
  "Gaming-heavy audience.",
  "Instagram DM.",
  "Not a priority.",
  "Off-target mix.",
  (1, 2, 2, 2, 3), "Lead",
  "Small and mixed audience.",
  src("heepsy_ig"))

# ---- Trading communities, signal channels
P("Cryptoiz", "Trading community", S["ajmarketing"], "Indonesia", "Likely", "N",
  "Indonesian crypto learners and traders on Telegram, Instagram and X.",
  "Education platform with 10K+ Telegram members, 18K Instagram, 28K X (AJ Marketing, 2025).",
  "Telegram group still active; figures dated.",
  "Telegram admin DM.",
  "Weekly onchain trading clinic in the group; referral code; rev share.",
  "Dated numbers; Telegram group activity unverified.",
  (2, 4, 3, 4, 4), "Partial",
  "Community format suits repeat engagement.",
  src("ajmarketing"))

P("Tradting", "Trading community", S["idx_komunitas"], "Indonesia", "Likely", "N",
  "Active crypto traders on WhatsApp and Discord.",
  "Indodax: solid community base on WhatsApp and Discord with daily market discussion.",
  "Size unknown; source is an exchange blog.",
  "Discord admin DM.",
  "Code-tracked referral and a monthly onchain trading review session.",
  "WhatsApp is hard to measure; attribution must be by code.",
  (2, 5, 3, 3, 4), "Lead",
  "Daily trading talk is the strongest intent signal; size unknown.",
  src("idx_komunitas"))

P("CryptoGalaxy", "Trading community", S["idx_komunitas"], "Indonesia", "Likely", "N",
  "Discord-based crypto traders.",
  "Indodax: active on Discord with real-time chat, community events and occasional offline meetups.",
  "Size unknown.",
  "Discord admin DM; attend a kopdar.",
  "Discord trading channel sponsorship with referral code.",
  "Unverified activity level.",
  (2, 4, 3, 3, 4), "Lead",
  "Trader community with offline meetups.",
  src("idx_komunitas"))

P("CryptoHack (Indonesia)", "Trading community", S["idx_komunitas"], "Indonesia", "Y", "N",
  "Crypto community attending offline talkshows.",
  "Indodax: active online with frequent offline talkshows featuring long-time Indonesian crypto figures.",
  "Name collides with a global cryptography site; local group needs confirmation.",
  "Attend a talkshow; organiser DM.",
  "Sponsor a talkshow on onchain execution.",
  "Identity ambiguity.",
  (2, 3, 2, 4, 4), "Lead",
  "In-person access; trading depth unclear.",
  src("idx_komunitas"))

P("Blockplay", "Trading community", S["idx_komunitas2"], "Indonesia", "Y", "N",
  "Crypto learners at public education events.",
  "Indodax: brings crypto education into public spaces with regular offline gatherings.",
  "Size unknown.",
  "Attend a gathering.",
  "Co-hosted gathering with wallet onboarding desk.",
  "Beginner skew.",
  (2, 3, 2, 4, 4), "Lead",
  "In-person access; beginner audience.",
  src("idx_komunitas2"))

P("Crypto Indonesia (Discord)", "Trading community", S["discord_ci"], "Indonesia", "N", "N",
  "Indonesian crypto buyers and sellers.",
  "Public Discord listing for Indonesian crypto buy/sell discussion; member count not shown.",
  "Activity unknown.",
  "Discord admin DM.",
  "Not a priority.",
  "OTC buy/sell servers carry scam risk.",
  (2, 3, 2, 2, 3), "Lead",
  "Unclear quality.",
  src("discord_ci"))

P("Crypto News Indonesia (Telegram)", "Signal/call channel", S["yellow_signals"], "Indonesia", "N", "N",
  "Signal followers on Telegram.",
  "Appears in Indonesian-language lists of crypto signal Telegram groups.",
  "Signal accuracy and membership unknown.",
  "Telegram admin DM.",
  "Rev-share-only referral; no fixed fee.",
  "Signal channels inflate claims; possible Telegram restrictions in Indonesia.",
  (3, 4, 2, 2, 2), "Lead",
  "Trader intent, weak trust signals.",
  src("yellow_signals"))

P("Bybit Indonesia (formerly NOBI)", "Licensed exchange", S["bybit_id"], "Jakarta", "Y", "N",
  "Indonesian retail exchange users.",
  "Bybit took a majority stake in local platform NOBI in 2025 and relaunched it as Bybit Indonesia under the local licence (Cointelegraph).",
  "Competes for the same trader.",
  "N/A",
  "None.",
  "Direct competitor.",
  (2, 3, 1, 2, 1), "Verified",
  "Licensed exchange: competes for the same trader.",
  src("bybit_id"))

P("BTSE Indonesia", "Licensed exchange", S["btse_id"], "Jakarta", "Y", "N",
  "Indonesian retail exchange users.",
  "BTSE Group launched a regulated crypto trading platform in Indonesia (crypto.news).",
  "Competes for the same trader.",
  "N/A",
  "None.",
  "Direct competitor.",
  (2, 3, 1, 2, 1), "Verified",
  "Licensed exchange: competes for the same trader.",
  src("btse_id"))

P("Pintu Web3 Wallet", "Licensed exchange", S["pintu_web3"], "Jakarta", "Y", "N",
  "Pintu app users moving onchain.",
  "Pintu launched a self-custodial Web3 wallet with DEX and dApp access, built on Fordefi (CoinDesk, May 2024).",
  "Competes directly with OKX Wallet for Indonesian onchain users.",
  "N/A",
  "None.",
  "Direct competitor.",
  (3, 3, 4, 2, 1), "Verified",
  "Competing wallet from a licensed exchange.",
  src("pintu_web3"))

P("Indonesia Crypto Exchange (ICEX / CFX)", "Licensed exchange", S["icex"], "Jakarta", "Y", "N",
  "Licensed traders and members of the national exchange.",
  "Launched an integrated platform under the OJK framework (DealStreetAsia).",
  "Market infrastructure, not a distribution partner.",
  "N/A",
  "None.",
  "Regulated market operator.",
  (1, 2, 1, 2, 1), "Verified",
  "Market infrastructure, not an affiliate.",
  src("icex"))

P("Hyperliquid Indonesia (community)", "Onchain trading community", S["hl_meetup"], "Jakarta (SCBD)", "Y", "N",
  "Indonesian onchain perps traders and DeFi users.",
  "Ran an approval-gated Jakarta SCBD meetup for traders with 62 registered (Luma). HYPE is now listed on Indodax, a sign of local demand. Pump.fun activity fell ~80% by Jun 2026 as traders moved to perps (Pluang).",
  "Organiser is a community group, not Hyperliquid Labs. Members already hold self-custody wallets.",
  "Contact the Luma host; attend the next meetup; offer to co-host.",
  "Co-hosted 'spot + perps onchain' meetup; OKX DEX referral code for spot swaps and bridging; rev share; leaderboard prizes paid only from earned rev share.",
  "Loyalty to Hyperliquid; small headcount; check OKX DEX product overlap with perps before pitching.",
  (4, 5, 5, 4, 3), "Partial",
  "Already trades onchain with leverage, so each referred wallet carries high volume. Meetable in Jakarta.",
  src("hl_meetup", "hl_indodax", "pump80"), "Top 3")

P("Xenorize", "Trading community", S["xenorize"], "Jakarta", "Y", "N",
  "Jakarta crypto community.",
  "Hosted a Jakarta community meetup listed on Luma (Oct 2025).",
  "Focus and size unknown.",
  "Luma host contact.",
  "Co-host a trader night.",
  "Unknown fit.",
  (2, 3, 3, 4, 3), "Lead",
  "In-person lead; content unknown.",
  src("xenorize"))

P("Indonesia Paham Bitcoin", "Trading community", S["cmif2025"], "Indonesia", "Y", "N",
  "Bitcoin-first Indonesian crypto community.",
  "Coinvestasi's CMIF 2025 write-up credits community lead Niki with growing the Coinvestasi and Indonesia Paham Bitcoin communities to 80,000+ active members.",
  "Bitcoin-first culture; activity of the 80K not audited.",
  "Via Coinvestasi community team (Niki).",
  "BTC-onchain education (wrapped BTC) with referral code.",
  "Low DEX appetite in a BTC-first community.",
  (3, 3, 2, 4, 4), "Partial",
  "Large and organised, but BTC-first.",
  src("cmif2025"))

P("Bitcoin Indonesia Community (city meetups)", "Bitcoin community", S["btc_id_meetups"], "Bandung Barat, Malang, Yogyakarta, Magelang, Bogor, Jakarta Barat", "Y", "N",
  "Bitcoin holders across Java.",
  "Free monthly meetups in 6+ cities (Jul-Aug 2026); listings state Bitcoin only, with no altcoin, NFT or token promotion.",
  "Will not host DEX content.",
  "Luma host contact.",
  "None; observe only.",
  "Explicitly rejects token promotion.",
  (2, 1, 2, 5, 4), "Verified",
  "Wrong fit by the community's own rules.",
  src("btc_id_meetups"))

P("Bitcoin House Bali (Bitcoin Indonesia)", "Bitcoin community", S["btc_house_bali"], "Canggu, Bali", "Y", "N",
  "Bitcoiners and expats in Bali.",
  "Free monthly Bitcoin meetup in Canggu run by Bitcoin Indonesia, in English and Bahasa.",
  "Bitcoin-only culture.",
  "Attend meetup.",
  "None.",
  "BTC-only.",
  (1, 1, 2, 5, 4), "Verified",
  "Wrong fit.",
  src("btc_house_bali"))

P("BitDevs Bali / Bitcoin Builders Association Indonesia", "Bitcoin community", S["bitdevs"], "Bali", "Y", "N",
  "Bitcoin developers.",
  "Free monthly BitDevs Bali meetups via Bitcoin Builders Association Indonesia.",
  "Developer audience.",
  "Attend meetup.",
  "None.",
  "BTC-only developers.",
  (1, 1, 3, 5, 4), "Verified",
  "Wrong fit.",
  src("bitdevs"))

P("Stacks Indonesia", "Bitcoin community", S["stacks"], "Bali", "Y", "N",
  "Bitcoin L2 enthusiasts in Bali.",
  "Runs monthly Bitcoin walks and meetups in Bali.",
  "Small group.",
  "Meetup contact.",
  "None.",
  "Niche ecosystem.",
  (1, 2, 3, 4, 4), "Verified",
  "Small, niche.",
  src("stacks"))

P("Forum Crypto Indonesia (Yan Suprich)", "Trading community", S["yap_panel"], "Indonesia", "Y", "N",
  "Long-running Indonesian crypto forum community.",
  "Yan Suprich of Forum Crypto Indonesia spoke on a Coinfest Asia 2025 side-event panel. Forum Bitcoin Indonesia partnered on Coinfest 2019 city roadshows.",
  "Current activity level unknown.",
  "Meet at Coinfest side events; forum admin contact.",
  "Forum section on onchain trading with referral code.",
  "Possibly ageing audience.",
  (2, 3, 3, 4, 4), "Partial",
  "Established community, reachable.",
  src("yap_panel", "coinfest2019"))

P("Cryptosiast Jakarta", "Trading community", S["cryptosiast_ws"], "Jakarta", "Y", "N",
  "Jakarta crypto community running Web3 workshops.",
  "Hosted 'From Zero to Web3 Hero' as an official Coinfest Asia 2025 side event.",
  "Beginner-heavy.",
  "Luma host contact; Jakarta meetup.",
  "Co-run a hands-on onchain swap workshop with referral code.",
  "Beginners trade small.",
  (2, 3, 3, 5, 4), "Verified",
  "Reachable organiser with workshop format.",
  src("cryptosiast_ws"))

P("Cryptosiast Bali (Gus Tut)", "Trading community", S["yap_panel"], "Bali", "Y", "N",
  "Bali crypto community.",
  "Gus Tut of Cryptosiast Bali spoke on a SocialFi/yapping panel; the group co-taught the Coinfest 2025 workshop.",
  "SocialFi focus may skew to points farming.",
  "Meet in Bali.",
  "Co-run a workshop with referral code.",
  "Incentive-farming lean.",
  (2, 3, 3, 5, 4), "Partial",
  "Reachable; farming risk to watch.",
  src("yap_panel", "cryptosiast_ws"))

P("Komunitas CryptoIndo Jogja", "Trading community", S["cryptosiast_ws"], "Yogyakarta", "Y", "N",
  "Yogyakarta crypto community.",
  "Co-taught the Coinfest Asia 2025 community workshop.",
  "Size unknown.",
  "Organiser contact via Coinfest network.",
  "Campus-city trader meetup with referral code.",
  "Unknown size.",
  (2, 3, 3, 4, 4), "Partial",
  "Second-city foothold.",
  src("cryptosiast_ws"))

P("Neko Group", "Trading community", S["cryptosiast_ws"], "Indonesia", "Unknown", "N",
  "Unknown.",
  "Listed as co-instructor at the Coinfest 2025 community workshop.",
  "Nature of the group unknown.",
  "Via Cryptosiast.",
  "Hold.",
  "Unknown fit.",
  (1, 2, 3, 3, 3), "Lead",
  "Too little evidence.",
  src("cryptosiast_ws"))

P("Komunitas BlockBali (Egy)", "Trading community", S["yap_panel"], "Bali", "Y", "N",
  "Bali crypto community.",
  "Egy of BlockBali spoke on a Coinfest Asia 2025 side-event panel.",
  "Size unknown.",
  "Meet in Bali.",
  "Co-host a trader meetup.",
  "Unknown size.",
  (2, 3, 3, 4, 4), "Lead",
  "Reachable Bali lead.",
  src("yap_panel"))

P("Dev Web3 Jogja", "Ecosystem community", S["id_web3_hack"], "Yogyakarta", "Y", "N",
  "Web3 developers and students in Yogyakarta.",
  "Co-organiser of the 2026 Indonesia Web3 Hackathon with Binance Academy, BNB Chain and Coinvestasi; co-taught the Coinfest 2025 workshop.",
  "Developer audience.",
  "Hackathon organiser contact.",
  "OKX DEX API track at a future hackathon.",
  "Binance-linked programme.",
  (1, 2, 4, 4, 3), "Verified",
  "Builders, not traders.",
  src("id_web3_hack", "cryptosiast_ws"))

# ---- Ecosystem communities and venues
P("Superteam Indonesia", "Ecosystem community", S["superteam_showcase"], "Jakarta and Bali", "Y", "N",
  "Solana builders, users and students; campus classes nationwide.",
  "Runs Superteam Hub Jakarta (with Indonesia Crypto Hub), ecosystem showcases, Campus Weekend Classes, Bali Solana Village and a 2026 Hacker House. Its Country Lead was named a CMIF 2025 most impactful figure, credited with 19 initiatives across ministries and agencies.",
  "Skews to builders more than traders; may avoid exchange-branded deals.",
  "Biweekly community call; visit Superteam Hub Jakarta; Country Lead via LinkedIn.",
  "Co-hosted 'trade onchain safely on Solana' workshop at the Hub; OKX DEX API for hackathon teams; referral code for attendees.",
  "Builder-first culture; Solana Foundation neutrality.",
  (3, 3, 5, 5, 4), "Verified",
  "Lowest-risk in-person door to onchain-native users.",
  src("superteam_showcase", "superteam_hub", "superteam_hackerhouse", "cmif2025"), "Top 3")

P("Indonesia Crypto Hub (ICH)", "Venue", S["superteam_hub"], "Jakarta", "Y", "N",
  "Jakarta builders and community members using the hub.",
  "Operates 'Superteam Hub Jakarta by ICH' and co-hosts Buildstation co-working (31 attendees at one session).",
  "Foot traffic modest.",
  "Visit the hub; Luma host contact.",
  "Monthly OKX onchain trader night at the hub.",
  "Small sessions.",
  (2, 3, 4, 5, 4), "Verified",
  "Physical base for repeat events.",
  src("superteam_hub"))

P("ETHJKT (Ethereum Jakarta)", "Ecosystem community", S["ethjkt"], "Jakarta", "Y", "N",
  "Ethereum builders and users in Jakarta.",
  "Ran Ethereum Jakarta Hackathon (Aug 2025), a Devcon satellite and regular meetups; community lead Revo Arya; co-hosts with MetaMask.",
  "Builder skew.",
  "Meetup attendance; community lead contact.",
  "L2 trading workshop with OKX DEX aggregator demo.",
  "Close MetaMask ties (competing wallet).",
  (2, 2, 5, 5, 3), "Verified",
  "Onchain-ready but builder-led.",
  src("ethjkt", "eth_genesis"))

P("BlockDevId", "Ecosystem community", S["blockdevid"], "Indonesia", "Y", "N",
  "Indonesian Web3 developers.",
  "Developer community and Superteam Indonesia community partner; associated with Indodax CEO William Sutanto (CMIF 2024).",
  "Exchange-adjacent.",
  "Event networking.",
  "None beyond hackathon API.",
  "Indodax affiliation.",
  (1, 2, 4, 4, 2), "Verified",
  "Competitor-adjacent builders.",
  src("blockdevid", "liputan_cmif2024"))

P("Base Indonesia", "Ecosystem community", S["base_ws"], "Jakarta", "Y", "N",
  "Base builders and users in Jakarta.",
  "Ran Base Indonesia Workshop (Jan-Feb 2026, JURA, 50 seats), Base Builders Meetup Jakarta and Base(d) in Jakarta: Bukber with IDRX.",
  "Builder skew.",
  "Luma host contact.",
  "Base onchain trading session with OKX DEX; referral code.",
  "Builder-first.",
  (2, 3, 5, 4, 4), "Verified",
  "Onchain-ready, reachable.",
  src("base_ws", "base_bukber"))

P("Lisk Indonesia (LiskSpark / Ignit3)", "Ecosystem community", S["lisk_ignit3"], "Jakarta and Bali", "Y", "N",
  "Government-backed Web3 startup cohort.",
  "Ignit3 programme with Gerakan Nasional 1000 Startup Digital (unveiled Coinfest 2024); LiskSpark incubator. Lisk co-organised OKX's Alphas Connect Bali.",
  "Startups, not traders.",
  "Warm intro via OKX SEA team that co-ran Alphas Connect.",
  "OKX DEX integration for cohort apps.",
  "Low direct trader volume.",
  (2, 2, 4, 5, 4), "Verified",
  "Warm relationship; startup audience.",
  src("lisk_ignit3", "okx_alphas_bali"))

P("Monad Indonesia (Foundation events)", "Ecosystem community", S["monad_blitz"], "Jakarta and Yogyakarta", "Y", "N",
  "Monad developers and users.",
  "Monad Foundation ran Monad Blitz Jakarta (Aug 2026, KOMDIGI office, $2K prizes), Monad Blitz Jogja (Apr 2026) and Monad 101 Jakarta.",
  "Local audience overlaps with testnet and airdrop farmers.",
  "Luma host contact.",
  "Trading session at next Blitz.",
  "Farming-heavy audience possible.",
  (2, 2, 5, 4, 4), "Verified",
  "Onchain-ready; intent unclear.",
  src("monad_blitz", "monad_101"))

P("Sui Indonesia", "Ecosystem community", S["sui_bootcamp"], "Jakarta", "Y", "N",
  "Sui builders and users.",
  "Sui Jakarta Bootcamp (Dec 2025, JURA); Sui Indonesia Telegram group.",
  "OKX DEX Sui coverage to confirm.",
  "Telegram group admin.",
  "Sui trading session with referral code.",
  "Builder skew.",
  (2, 2, 4, 4, 4), "Verified",
  "Onchain-ready; small.",
  src("sui_bootcamp"))

P("NEAR Legion Jakarta", "Ecosystem community", S["near_legion"], "Jakarta", "Y", "N",
  "NEAR community members.",
  "Ran a NEAR Legion Jakarta meetup (Luma).",
  "Small.",
  "Luma host.",
  "None.",
  "Low volume.",
  (1, 2, 3, 4, 4), "Verified",
  "Small ecosystem group.",
  src("near_legion"))

P("ODX Labs Jakarta", "Ecosystem community", S["odx"], "Jakarta", "Y", "N",
  "Unknown.",
  "Hosted a Jakarta community meetup (Luma).",
  "Focus unknown.",
  "Luma host.",
  "Hold.",
  "Unknown fit.",
  (1, 2, 3, 3, 3), "Lead",
  "Too little evidence.",
  src("odx"))

P("JURA (Kemanggisan)", "Venue", S["jura"], "Jakarta Barat", "Y", "N",
  "Web3 and AI builders using the co-working space.",
  "Hosts weekly W3AI Wednesday co-working and Base, Sui and Superteam sessions.",
  "Venue, not an audience owner.",
  "Walk in; Luma host.",
  "Venue for OKX trader nights.",
  "None material.",
  (1, 2, 3, 5, 5), "Verified",
  "Useful venue, no volume on its own.",
  src("jura"))

P("Ethereum Meetup Indonesia", "Ecosystem community", S["eth_meetup_pintu"], "Multi-city", "Y", "N",
  "Ethereum developers.",
  "Run by PINTU (licensed exchange) with roadshows to three cities.",
  "Pintu controls the agenda.",
  "N/A",
  "None.",
  "Competitor-run.",
  (1, 2, 4, 4, 2), "Verified",
  "Competitor-run.",
  src("eth_meetup_pintu"))

P("IDNFT", "Ecosystem community", S["okx_alphas_bali"], "Bali", "Y", "N",
  "Indonesian NFT and Web3 community.",
  "Co-organised OKX Alphas Connect Bali with OKX SEA, Tether, Lisk and Sui.",
  "NFT-first audience.",
  "Warm intro via OKX SEA team.",
  "Co-host next Alphas event in Indonesia.",
  "NFT audience trades less.",
  (2, 2, 3, 4, 4), "Partial",
  "Warm, but NFT-first.",
  src("okx_alphas_bali"))

P("Solana Village Bali", "Event", S["solana_village"], "Bali", "Y", "N",
  "Solana builders and users in Bali.",
  "Co-hosted by Superteam Indonesia and Solana Foundation in Bali.",
  "Recurring edition not confirmed.",
  "Via Superteam Indonesia.",
  "Trader session inside the village programme.",
  "Event-only reach.",
  (2, 3, 5, 4, 4), "Verified",
  "Good access point, limited volume.",
  src("solana_village"))

# ---- Onchain tools
P("IDRX (rupiah stablecoin)", "Onchain tool", S["idrx_cg"], "Jakarta", "Y", "N",
  "Indonesians moving rupiah onchain; users on Base, BNB Chain, Polygon, Lisk and Solana.",
  "Rupiah stablecoin by PT IDRX Indo Inovasi; company says US$50M+ all-time onchain transactions (Jun 2025). Most active pair IDRX/USDC on PancakeSwap (Base), ~$12K 24h volume (CoinGecko snapshot). Co-hosted Base(d) in Jakarta.",
  "Liquidity is thin; integration appetite unknown.",
  "Founders via Base Indonesia events; LinkedIn.",
  "IDRX routing and 'rupiah to any token' flows on OKX DEX; co-marketing to IDRX holders; referral code inside IDRX channels.",
  "Low current liquidity; regulatory scope of stablecoin promotion.",
  (3, 3, 5, 4, 4), "Verified",
  "Removes the rupiah on-ramp friction that blocks onchain trading.",
  src("idrx_cg", "idrx_50m", "base_bukber"), "Top 8")

P("Xellar", "Onchain tool", S["monad_101"], "Indonesia", "Y", "N",
  "Indonesian apps using embedded wallets.",
  "PT Xellar Teknologi Indonesia: wallet-as-a-service with embedded, hot and cold wallets; supported Monad 101 Jakarta.",
  "Client apps and user numbers unknown.",
  "Founder contact via events.",
  "Embed OKX DEX API swaps in Xellar-powered apps; fee share on routed volume.",
  "B2B sales cycle longer than 30 days.",
  (2, 2, 5, 4, 4), "Partial",
  "Distribution through partner apps.",
  src("monad_101", "xellar"))

P("Satu Wallet", "Onchain tool", S["satu"], "Tanjungpinang, Riau Islands", "Unknown", "N",
  "Indonesian self-custody wallet users.",
  "Non-custodial wallet launched Mar 2026 on Google Play and App Store by PT Bintan Teknologi Kreatif; $SATU token on Solana.",
  "User base not disclosed.",
  "Company contact.",
  "OKX DEX API as swap engine; fee share.",
  "Own token creates incentive conflicts.",
  (2, 3, 5, 3, 3), "Partial",
  "Onchain users, unknown scale.",
  src("satu"))

P("BISA Wallet", "Onchain tool", S["bisa"], "Indonesia", "Unknown", "N",
  "Bahasa-first Web3 app users.",
  "Self-described Bahasa-first Web3 app for portfolio, token and NFT trading; claims unverified.",
  "May be pre-launch.",
  "Site contact.",
  "Hold.",
  "Unverified product.",
  (1, 3, 4, 2, 2), "Lead",
  "Unverified.",
  src("bisa"))

P("Tokenomy", "Onchain tool", S["tokenomy"], "Jakarta", "Y", "N",
  "Indodax-linked investors.",
  "Started 2017 as Indodax's innovation arm; founder Nathanael Christian (CMIF 2024).",
  "Competitor alignment.",
  "N/A",
  "None.",
  "Competitor-affiliated.",
  (1, 2, 3, 2, 2), "Partial",
  "Competitor-affiliated.",
  src("tokenomy", "liputan_cmif2024"))

P("Blockvest", "Onchain tool", S["blockvest"], "Jakarta", "Y", "N",
  "Indonesian crypto industry players and students.",
  "Industry mapping and events platform (founder Isybel Harto); campus programmes with Pintu.",
  "Event-led.",
  "Founder contact.",
  "Event sponsorship only.",
  "Pintu ties.",
  (2, 2, 2, 4, 3), "Partial",
  "Events, not traders.",
  src("blockvest", "upj"))

P("CryptoWave", "Analytics & news creator", S["cryptowave"], "Indonesia", "Unknown", "N",
  "Crypto news and education app users.",
  "News and education app reported to have hundreds of thousands of users (press).",
  "User count is a company claim.",
  "Company contact.",
  "In-app onchain guide with referral code.",
  "Unaudited users.",
  (3, 3, 2, 3, 4), "Partial",
  "App distribution; intent unclear.",
  src("cryptowave"))

# ---- Media
P("Coinvestasi / Indonesia Crypto Network", "Media", S["coinvestasi_partner"], "Jakarta and Bali", "Y", "N",
  "Indonesia's largest crypto media audience; runs Coinfest Asia, CMIF and Web3 University Tour; also operates CoinDesk Indonesia.",
  "Founded 2017 as a crypto information and education portal; public partner and advertising pages. With Indonesia Paham Bitcoin, its communities have 80K+ active members (CMIF 2025 write-up).",
  "Will accept a performance deal rather than flat sponsorship.",
  "coinvestasi.com/partner form; meet at Coinfest or Jakarta office.",
  "Performance-based 'onchain trading guide' hub with tracked referral codes; trader track at Coinfest 2027.",
  "Commercial ties with Binance Academy and Tokocrypto; prefers paid media.",
  (4, 3, 3, 4, 3), "Verified",
  "Biggest reach with a trading-literate audience; deal shape is the risk.",
  src("coinvestasi_partner", "coinvestasi_bisnis", "cmif2025"), "Top 8")

P("CoinDesk Indonesia", "Media", S["coindesk_id"], "Jakarta", "Y", "N",
  "Indonesian crypto news readers.",
  "Launched Aug 2022 at Coinfest; Managing Director Glenn Ardi; operated within Indonesia Crypto Network.",
  "Editorial rules limit affiliate formats.",
  "Via ICN commercial team.",
  "Sponsored market reports with referral code.",
  "Editorial independence.",
  (3, 3, 3, 4, 4), "Verified",
  "Trusted reach; format limits.",
  src("coindesk_id"))

P("BeInCrypto Indonesia", "Media", S["bic_cfx"], "Indonesia", "Unknown", "N",
  "Indonesian crypto news readers.",
  "Editorial partnership with CFX (OJK-regulated exchange) since Jul 2025; one writer has 5,000+ articles.",
  "CFX partnership may block competing promotion.",
  "Editorial or ads team.",
  "Display and content test.",
  "Conflict with CFX partnership.",
  (3, 3, 2, 3, 2), "Verified",
  "Reach, but conflicted.",
  src("bic_cfx"))

P("Cryptonews Indonesia", "Media", S["cryptonews_about"], "Indonesia", "Y", "N",
  "Indonesian crypto readers.",
  "Publishes Indonesian listicles on creators, signals and airdrops; attends Coinfest.",
  "Readers skew to presale and airdrop content.",
  "Ads team; Coinfest.",
  "Sponsored guide with referral code.",
  "Presale advertising dilutes trust.",
  (3, 3, 2, 3, 3), "Verified",
  "Reach with mixed intent.",
  src("cryptonews_about", "cryptonews_kol"))

P("Blockchainmedia.id", "Media", S["bcm_tg"], "Indonesia", "Unknown", "N",
  "Indonesian crypto news readers.",
  "Cited as a source in an Indonesian crypto news Telegram digest.",
  "Audience size unknown.",
  "Editorial contact.",
  "Hold.",
  "Unknown reach.",
  (2, 2, 2, 3, 4), "Lead",
  "Too little evidence.",
  src("bcm_tg"))

P("Bitcoinist Indonesia", "Media", S["bitcoinist_id"], "Indonesia", "N", "N",
  "Indonesian crypto readers.",
  "Indonesian listicles mixed with presale press releases.",
  "Presale-heavy readership.",
  "Ads team.",
  "None.",
  "Presale promotion association.",
  (3, 2, 2, 3, 2), "Partial",
  "Low-trust content mix.",
  src("bitcoinist_id"))

P("CoinSpeaker Indonesia", "Media", S["coinspeaker_id"], "Indonesia", "N", "N",
  "Young Indonesian memecoin readers.",
  "Publishes Solana memecoin '1000x' content aimed at young Indonesian investors.",
  "Readers chase pumps.",
  "Ads team.",
  "None.",
  "Hype content.",
  (2, 3, 3, 3, 2), "Partial",
  "Hype-led.",
  src("coinspeaker_id"))

P("Investortrust.id", "Media", S["investortrust"], "Jakarta", "Y", "N",
  "Mainstream finance readers.",
  "Mainstream financial outlet with dense crypto coverage (cited across this research).",
  "Low onchain share.",
  "Ads team.",
  "PR only.",
  "Off-target.",
  (2, 2, 1, 3, 4), "Partial",
  "PR value, not volume.",
  src("investortrust"))

# ---- Events, associations, channels
P("Coinfest Asia", "Event", S["coinfest2026"], "Bali", "Y", "N",
  "Asia-wide crypto attendees; heavy Indonesian community presence.",
  "2025 drew ~10,000 attendees from 90 countries; 2026 ran Aug 20-21 at Melasti Beach with ~5,000 per day expected. OKX sponsored the 2024 edition.",
  "2027 edition timing.",
  "ICN partnership team.",
  "Trader side event; recruit partners on site.",
  "Event cost versus measurable volume.",
  (2, 3, 4, 4, 4), "Verified",
  "Recruiting venue, not a volume source.",
  src("coinfest2026", "coinfest2025", "okx_coinfest"))

P("Indonesia Blockchain Week (IDBW)", "Event", S["idbw"], "Jakarta (JICC)", "Y", "N",
  "Institutional and enterprise Web3 attendees.",
  "6th edition Aug 12-13, 2026 at JICC; 12,000+ attendees per press release; chairman Adytia Raflein.",
  "Institutional skew.",
  "Organiser contact.",
  "Networking only.",
  "Low retail-trader share.",
  (2, 2, 3, 4, 4), "Verified",
  "Institutional focus.",
  src("idbw"))

P("Web3 Week Asia", "Event", S["web3weekasia"], "Jakarta", "Y", "N",
  "Students and young Web3 professionals.",
  "Founded by Alex J. Rumondor, Andreas Tobing, Jodi Kalim and Muhammad Rofiq; annual event Nov 20-21, 2025 in Jakarta; campus programmes at Binus, UI, Untar and UGM.",
  "2026 edition not confirmed.",
  "Via Andreas Tobing.",
  "Campus onchain trading track.",
  "Student budgets.",
  (2, 2, 3, 5, 4), "Verified",
  "Reachable, low volume.",
  src("web3weekasia"))

P("ABI / Aspakrindo", "Association", S["abi"], "Jakarta", "Y", "N",
  "Licensed Indonesian crypto businesses.",
  "Industry association with merged leadership with Aspakrindo (chair Robby Bun); runs Bulan Literasi Kripto with OJK; launched a consumer complaint portal in Apr 2026.",
  "Members are licensed exchanges.",
  "Formal meeting request.",
  "Compliance relationship only.",
  "Members compete with OKX.",
  (1, 1, 1, 3, 3), "Verified",
  "Regulatory relationship, not distribution.",
  src("abi"))

P("AFTECH x IARFC Verified Finfluencer programme", "Association", S["aftech"], "Jakarta", "Y", "N",
  "Creators seeking POJK 6/2026 certification.",
  "Certification programme built on POJK 6/2026, which requires finfluencers covering crypto to hold competency certification.",
  "Cohort list may not be public.",
  "AFTECH secretariat.",
  "Source certified creators from the cohort; make certification a contract condition.",
  "Certification is not a licence to give personal advice.",
  (3, 3, 2, 4, 5), "Verified",
  "Pre-vetted creator pool that solves the compliance gate.",
  src("aftech", "pojk6"))

P("Southeast Asia Blockchain Week / Protocol Camp", "Event", S["ekraf_seabw"], "Indonesia", "Y", "N",
  "Web3 talent and builders.",
  "Supported by the Ministry of Creative Economy as a national Web3 talent programme (2026).",
  "Builder audience.",
  "Organiser contact.",
  "None.",
  "Low trader share.",
  (1, 2, 3, 4, 4), "Verified",
  "Talent, not traders.",
  src("ekraf_seabw"))

P("Bali Blockchain Weeks Campus Tour", "Event", S["bali_campus"], "Bali", "Y", "N",
  "Students.",
  "Campus tour listing on Luma.",
  "Details thin.",
  "Luma host.",
  "None.",
  "Low volume.",
  (1, 2, 2, 4, 4), "Partial",
  "Students.",
  src("bali_campus"))

P("Web3 University Tour / Indonesia Web3 Hackathon", "Event", S["id_web3_hack"], "Multi-city", "Y", "N",
  "Students and builders.",
  "Run by Coinvestasi with Binance Academy, BNB Chain and Dev Web3 Jogja.",
  "Binance-aligned.",
  "N/A",
  "None.",
  "Competitor-aligned.",
  (1, 2, 3, 4, 2), "Verified",
  "Competitor-aligned.",
  src("id_web3_hack"))

P("Bulan Literasi Kripto (OJK + ABI)", "Event", S["blk2026"], "Jakarta", "Y", "N",
  "Public crypto literacy audience.",
  "OJK and ABI's 2026 crypto literacy month, with emphasis on influencers.",
  "Regulator-led.",
  "Observe.",
  "None.",
  "Regulator setting.",
  (1, 1, 1, 3, 4), "Verified",
  "Not a commercial channel.",
  src("blk2026"))

P("OKX Alphas Indonesia alumni (internal)", "Internal channel", S["okx_alpha_summit"], "Bali", "Y", "N",
  "Traders, affiliates and community leaders from OKX's Bali events.",
  "OKX Alpha Traders Summit in Bali (Jan 22, 2025) gathered 200+ traders, affiliates and community leaders; OKX SEA and OKX Alphas later hosted Alphas Connect Bali.",
  "Attendee list exists internally; part of it is Indonesian.",
  "Internal CRM; re-contact in week 1.",
  "Re-activate under DEX referral; first access to co-hosted Indonesian meetups.",
  "Prior relationship quality unknown; may be on competitor programmes.",
  (4, 5, 4, 5, 4), "Verified",
  "Warmest leads; listed for completeness, excluded from the external top 8.",
  src("okx_alpha_summit", "okx_alphas_bali"), "Internal")

P("MetaMask Community Builder Night Jakarta", "Event", S["metamask_jkt"], "Jakarta", "Y", "N",
  "Jakarta Ethereum users.",
  "Recurring Jakarta nights co-hosted by MetaMask and ETHJKT (Mar and Jul 2026).",
  "Competitor wallet series.",
  "Attend as guest.",
  "Meet attendees only.",
  "Competitor-run.",
  (1, 2, 5, 4, 2), "Verified",
  "Competitor event; useful only for scouting.",
  src("metamask_jkt"))

P("CTO 'Community Take Over' side event (Coinfest 2025)", "Onchain trading community", S["cto_event"], "Bali", "Y", "N",
  "Memecoin community-takeover participants.",
  "Coinfest Asia 2025 (IN)Side event themed on memecoin community takeovers.",
  "Host group and size unknown.",
  "Luma host contact.",
  "Rev-share-only referral if hosts are traders.",
  "Memecoin abuse risk.",
  (2, 4, 5, 3, 2), "Lead",
  "Onchain-native, unverified.",
  src("cto_event"))

P("Yap + Network side event (SocialFi)", "Airdrop community", S["yap_panel"], "Bali", "Y", "Y",
  "SocialFi and post-to-earn participants.",
  "Coinfest 2025 side event on SocialFi and 'yapping' (post-to-earn).",
  "Audience farms points.",
  "N/A",
  "None.",
  "Incentive farming.",
  (2, 2, 3, 3, 2), "Partial",
  "Incentive-led audience.",
  src("yap_panel"))

# ---- Campus
P("BINUS Blockchain and Crypto Club (BBCC)", "Campus club", S["binus"], "Jakarta", "Y", "N",
  "BINUS students.",
  "Described as Indonesia's first university-level blockchain and crypto club.",
  "Low capital per student.",
  "Club officers via student activity page.",
  "Campus onchain trading workshop.",
  "Small tickets; age checks.",
  (1, 2, 2, 5, 4), "Verified",
  "Reachable, low volume.",
  src("binus"))

P("Universitas Kristen Petra WEC Blockchain Lab", "Campus club", S["petra"], "Surabaya", "Y", "N",
  "Petra students.",
  "Web3 Education Center launched with Upbit Indonesia (May 2026).",
  "Upbit-aligned.",
  "N/A",
  "None.",
  "Competitor-aligned.",
  (1, 2, 2, 4, 2), "Verified",
  "Competitor-aligned.",
  src("petra"))

P("BEM FEB UI (Biro Project)", "Campus club", S["ui_feb"], "Depok", "Y", "N",
  "UI economics students.",
  "Hosted Web3Week Asia x Universitas Indonesia.",
  "Event-driven.",
  "BEM contact.",
  "Campus talk.",
  "Low volume.",
  (1, 2, 1, 4, 4), "Partial",
  "Students.",
  src("ui_feb"))

P("YESCON (IDBW @ UI)", "Campus club", S["yescon"], "Depok", "Y", "N",
  "UI students.",
  "Ran an Indonesia Blockchain Week session at UI.",
  "Event-driven.",
  "Organiser contact.",
  "Campus talk.",
  "Low volume.",
  (1, 2, 2, 4, 4), "Partial",
  "Students.",
  src("yescon"))

P("Universitas Pembangunan Jaya Web3 club (planned)", "Campus club", S["upj"], "Tangerang Selatan", "Y", "N",
  "UPJ informatics students.",
  "Planned IT Club Web3 and Blockchain; programmes with Pintu and Blockvest.",
  "Pintu-aligned.",
  "N/A",
  "None.",
  "Competitor-aligned.",
  (1, 1, 1, 4, 2), "Verified",
  "Competitor-aligned, not yet formed.",
  src("upj"))

# ---- Airdrop-led (kept for contrast)
P("Airdrop Finder (Telegram)", "Airdrop community", S["airdrop_finder"], "Indonesia", "N", "Y",
  "Indonesian airdrop hunters.",
  "Indonesian airdrop Telegram group examined in a legal journal study.",
  "Large but farming-driven.",
  "N/A",
  "None.",
  "Sybil and wash volume.",
  (3, 1, 4, 2, 2), "Partial",
  "Airdrop-led: excluded by gate.",
  src("airdrop_finder"))

P("Tech Dropper (Telegram)", "Airdrop community", S["techdropper"], "Indonesia", "N", "Y",
  "Airdrop task followers.",
  "Telegram channel posting airdrop task lists.",
  "Indonesian share unknown.",
  "N/A",
  "None.",
  "Sybil and wash volume.",
  (2, 1, 4, 2, 2), "Partial",
  "Airdrop-led: excluded by gate.",
  src("techdropper"))

P("Tracking Drops (Telegram)", "Airdrop community", S["trackingdrops"], "Indonesia", "N", "Y",
  "Airdrop task followers.",
  "Telegram channel posting airdrop task lists.",
  "Indonesian share unknown.",
  "N/A",
  "None.",
  "Sybil and wash volume.",
  (2, 1, 4, 2, 2), "Partial",
  "Airdrop-led: excluded by gate.",
  src("trackingdrops"))

P("cnesia112 (Telegram)", "Airdrop community", S["cnesia"], "Indonesia", "N", "Y",
  "Airdrop followers.",
  "Telegram channel posting airdrop announcements.",
  "Ownership unknown.",
  "N/A",
  "None.",
  "Sybil and wash volume.",
  (2, 1, 3, 1, 2), "Lead",
  "Airdrop-led: excluded by gate.",
  src("cnesia"))

# ---- Adjacent trading platforms
P("Stockbit (community)", "Trading community", S["stockbit"], "Jakarta", "Y", "N",
  "Indonesian stock traders.",
  "Stock transactions reached Rp107T in Feb 2026; Bibit and Stockbit serve 10M+ users, ~90% millennial and Gen Z; built-in investor discussion feeds.",
  "Licensed securities broker will not promote offshore onchain trading.",
  "N/A",
  "None.",
  "Regulated competitor for wallet share.",
  (2, 4, 1, 1, 2), "Verified",
  "Real traders, unreachable for this product.",
  src("stockbit"))

assert len(ROWS) == 100, len(ROWS)


# ---------------------------------------------------------------- scoring (CSV mirror of sheet formulas)
def score(s):
    v, t, o, r, k = s
    raw = (v * WEIGHTS["V"] + t * WEIGHTS["T"] + o * WEIGHTS["O"] + r * WEIGHTS["R"] + k * WEIGHTS["K"]) / 5 * 100
    return int(raw + 0.5)


def priority(row):
    sc = score(row["scores"])
    k = row["scores"][4]
    if k <= 1 or row["airdrop"] == "Y":
        return "Low"
    if sc >= HIGH and row["ver"] != "Lead" and k >= 3:
        return "High"
    return "Medium" if sc >= MED else "Low"


order = {"High": 0, "Medium": 1, "Low": 2}
ROWS.sort(key=lambda r: (order[priority(r)], -score(r["scores"]), r["name"]))

# Top 8 external picks (assessment Task 1), in approach order
TOP8 = ["Hyperliquid Indonesia (community)", "Andy Senjaya", "Superteam Indonesia",
        "Michael Wyann", "IDRX (rupiah stablecoin)", "Hendy Tanuardy",
        "Coinvestasi / Indonesia Crypto Network", "Andreas Tobing"]

# ---------------------------------------------------------------- workbook
ARIAL = "Arial"
HDR_FILL = PatternFill("solid", fgColor="1F2937")
HDR_FONT = Font(name=ARIAL, bold=True, color="FFFFFF", size=10)
BODY = Font(name=ARIAL, size=9)
INPUT = Font(name=ARIAL, size=9, color="0000FF")
BOLD = Font(name=ARIAL, size=10, bold=True)
TITLE = Font(name=ARIAL, size=14, bold=True)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="top")
THIN = Side(style="thin", color="D1D5DB")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
YELLOW = PatternFill("solid", fgColor="FFFF00")

wb = Workbook()

# --- Method sheet
m = wb.active
m.title = "Method"
m["A1"] = "OKX DEX Affiliate Prospects: Indonesia (scoring method)"
m["A1"].font = TITLE
m["A3"] = "Scoring weights (edit blue cells; total must equal 100%)"
m["A3"].font = BOLD
hdr = ["Criterion", "Weight", "Scores 5 when", "Scores 1 when"]
for i, h in enumerate(hdr, 1):
    c = m.cell(row=4, column=i, value=h)
    c.font, c.fill, c.alignment = HDR_FONT, HDR_FILL, WRAP
crit = [
    ("V  Volume potential", WEIGHTS["V"], "Large audience of people who already trade weekly", "Tiny or non-trading audience"),
    ("T  Trading intent", WEIGHTS["T"], "Content and chat are about trades, setups, execution", "Giveaways, airdrops, buy-and-hold or general finance"),
    ("O  Onchain readiness", WEIGHTS["O"], "Members already use self-custody wallets and DEXs", "CEX-only or stock-only users"),
    ("R  Reachability", WEIGHTS["R"], "Decision-maker reachable in 30 days, meetable in person", "No known route, or competitor-controlled"),
    ("K  Risk safety", WEIGHTS["K"], "No legal, reputational or conflict issues found", "Active legal overhang or severe abuse risk"),
]
for r, (name, w, hi, lo) in enumerate(crit, 5):
    m.cell(row=r, column=1, value=name).font = BODY
    c = m.cell(row=r, column=2, value=w)
    c.font, c.number_format, c.fill = INPUT, "0%", YELLOW
    m.cell(row=r, column=3, value=hi).font = BODY
    m.cell(row=r, column=4, value=lo).font = BODY
m["A10"] = "Total"
m["A10"].font = BOLD
m["B10"] = "=SUM(B5:B9)"
m["B10"].number_format = "0%"
m["B10"].font = BOLD

m["A12"] = "Priority thresholds (score 0-100)"
m["A12"].font = BOLD
m["A13"] = "High at or above"
m["B13"] = HIGH
m["A14"] = "Medium at or above"
m["B14"] = MED
for ref in ("B13", "B14"):
    m[ref].font, m[ref].fill = INPUT, YELLOW
for ref in ("A13", "A14"):
    m[ref].font = BODY

m["A16"] = "Gates applied after the score"
m["A16"].font = BOLD
gates = [
    "1. Risk safety of 1 (active legal overhang) or an airdrop-led audience forces Low, whatever the score.",
    "2. High requires Verification of Partial or Verified and Risk safety of 3 or more. Leads and Risk safety 2 cap at Medium.",
    "3. Score = (V x wV + T x wT + O x wO + R x wR + K x wK) / 5 x 100, rounded.",
]
for i, g in enumerate(gates, 17):
    m.cell(row=i, column=1, value=g).font = BODY

m["A21"] = "Verification legend"
m["A21"].font = BOLD
legend = [
    ("Verified", "Key facts found in at least one independent or primary public source during research (Nov 2026)."),
    ("Partial", "Facts come from a single commercial or exchange-written source (e.g. Indodax Academy, listicles, press releases)."),
    ("Lead", "Name found in a directory or panel listing only; audience and content not yet checked."),
]
for i, (k, v) in enumerate(legend, 22):
    m.cell(row=i, column=1, value=k).font = BOLD
    m.cell(row=i, column=2, value=v).font = BODY

m["A26"] = "How to use"
m["A26"].font = BOLD
howto = [
    "Edit the five 1-5 scores (blue) on the Prospects sheet; Score, Priority and Rank recalculate.",
    "Columns I (Evidence) hold sourced facts; column J (Assumptions) holds what still needs checking. Follower counts must be re-checked on-platform before outreach.",
    "Licensed exchanges (Bybit Indonesia, BTSE Indonesia, Pintu, ICEX/CFX) are listed only to show they were screened: they compete for the same trader.",
]
for i, t in enumerate(howto, 27):
    m.cell(row=i, column=1, value=t).font = BODY

m.column_dimensions["A"].width = 28
m.column_dimensions["B"].width = 12
m.column_dimensions["C"].width = 52
m.column_dimensions["D"].width = 48
for row in m.iter_rows(min_row=5, max_row=9, max_col=4):
    for c in row:
        c.alignment = WRAP
        c.border = BORDER

# --- Prospects sheet
ws = wb.create_sheet("Prospects")
COLS = [
    ("ID", 6), ("Partner", 26), ("Partner type", 18), ("Public link", 30), ("Base / market", 18),
    ("In person?", 9), ("Airdrop-led?", 9), ("Audience and market", 34), ("Evidence of fit (sourced)", 50),
    ("Assumptions (not yet verified)", 34), ("Acquisition route", 30), ("Proposed partnership", 42),
    ("Main risk", 34), ("V Volume (1-5)", 8), ("T Trading intent (1-5)", 8), ("O Onchain ready (1-5)", 8),
    ("R Reach (1-5)", 8), ("K Risk safety (1-5)", 8), ("Verification", 10), ("Score (0-100)", 9),
    ("Priority", 9), ("Rank", 6), ("Shortlist", 9), ("Priority reason", 40), ("Sources", 60),
]
for i, (h, w) in enumerate(COLS, 1):
    c = ws.cell(row=1, column=i, value=h)
    c.font, c.fill, c.alignment, c.border = HDR_FONT, HDR_FILL, WRAP, BORDER
    ws.column_dimensions[get_column_letter(i)].width = w
ws.row_dimensions[1].height = 42

N = len(ROWS)
last = N + 1
for idx, row in enumerate(ROWS, start=1):
    r = idx + 1
    pid = f"P{idx:03d}"
    row["id"] = pid
    shortlist = "Top 3" if row["flag"] == "Top 3" else ("Top 8" if row["name"] in TOP8 else row["flag"] if row["flag"] == "Internal" else "")
    row["shortlist"] = shortlist
    vals = [pid, row["name"], row["type"], row["link"], row["base"], row["inperson"], row["airdrop"],
            row["audience"], row["evidence"], row["assumptions"], row["route"], row["proposal"], row["risk"]]
    for ci, v in enumerate(vals, 1):
        c = ws.cell(row=r, column=ci, value=v)
        c.font, c.alignment, c.border = BODY, WRAP, BORDER
    ws.cell(row=r, column=4).hyperlink = row["link"]
    for ci, s in enumerate(row["scores"], 14):
        c = ws.cell(row=r, column=ci, value=s)
        c.font, c.alignment, c.border = INPUT, CENTER, BORDER
    c = ws.cell(row=r, column=19, value=row["ver"])
    c.font, c.alignment, c.border = BODY, CENTER, BORDER
    ws.cell(row=r, column=20, value=(
        f"=ROUND((N{r}*Method!$B$5+O{r}*Method!$B$6+P{r}*Method!$B$7"
        f"+Q{r}*Method!$B$8+R{r}*Method!$B$9)/5*100,0)"))
    ws.cell(row=r, column=21, value=(
        f'=IF(OR(R{r}<=1,G{r}="Y"),"Low",IF(AND(T{r}>=Method!$B$13,S{r}<>"Lead",R{r}>=3),"High",'
        f'IF(T{r}>=Method!$B$14,"Medium","Low")))'))
    ws.cell(row=r, column=22, value=f"=RANK(T{r},$T$2:$T${last})")
    for ci in (20, 21, 22):
        c = ws.cell(row=r, column=ci)
        c.font, c.alignment, c.border = BOLD if ci == 21 else BODY, CENTER, BORDER
    ws.cell(row=r, column=20).number_format = "0"
    for ci, v in ((23, shortlist), (24, row["reason"]), (25, row["sources"])):
        c = ws.cell(row=r, column=ci, value=v)
        c.font, c.alignment, c.border = BODY, WRAP, BORDER
    ws.row_dimensions[r].height = 96

ws.freeze_panes = "C2"
ws.auto_filter.ref = f"A1:Y{last}"
green = PatternFill("solid", fgColor="C6EFCE")
amber = PatternFill("solid", fgColor="FFEB9C")
grey = PatternFill("solid", fgColor="E5E7EB")
ws.conditional_formatting.add(f"U2:U{last}", CellIsRule(operator="equal", formula=['"High"'], fill=green))
ws.conditional_formatting.add(f"U2:U{last}", CellIsRule(operator="equal", formula=['"Medium"'], fill=amber))
ws.conditional_formatting.add(f"U2:U{last}", CellIsRule(operator="equal", formula=['"Low"'], fill=grey))
ws.conditional_formatting.add(f"T2:T{last}", ColorScaleRule(start_type="num", start_value=30, start_color="F8696B",
                                                          mid_type="num", mid_value=55, mid_color="FFEB84",
                                                          end_type="num", end_value=85, end_color="63BE7B"))

# --- Shortlist sheet
sl = wb.create_sheet("Shortlist")
sl["A1"] = "Top 8 external prospects (approach order) and why the first 3 go first"
sl["A1"].font = TITLE
sl["A2"] = ("Fields pull live from the Prospects sheet by ID. All 8 are non-airdrop; types span onchain trading community, "
            "trading educators, ecosystem community, onchain tool and media; 5 are meetable in person.")
sl["A2"].font = BODY
SL_COLS = [("Order", 6), ("ID", 7), ("Partner", 26), ("Type", 18), ("Link", 28), ("Audience and market", 34),
           ("Evidence of fit", 48), ("Acquisition route", 30), ("Proposed partnership", 40), ("Main risk", 32),
           ("Score", 7), ("Priority", 9), ("Why approach first (top 3)", 54)]
for i, (h, w) in enumerate(SL_COLS, 1):
    c = sl.cell(row=4, column=i, value=h)
    c.font, c.fill, c.alignment, c.border = HDR_FONT, HDR_FILL, WRAP, BORDER
    sl.column_dimensions[get_column_letter(i)].width = w
WHY = {
    "Hyperliquid Indonesia (community)": "Members already trade with leverage from self-custody wallets, so we skip wallet onboarding, the step where most conversions die. One Jakarta meetup puts us in a room with 60+ active onchain traders in week 2.",
    "Andy Senjaya": "Largest trader-education audience with a third-party count (~695K YouTube, Jul 2026). Illustration only: 0.5% converting is ~3,500 wallets, more than any community on the list. Certification check comes first.",
    "Superteam Indonesia": "Lowest-risk in-person door: a staffed Jakarta hub, a named Country Lead and weekly events. Solana users are already onchain, and a safety-first trading workshop fits their education mandate.",
}
id_by_name = {r["name"]: r["id"] for r in ROWS}
for o, name in enumerate(TOP8, 1):
    r = o + 4
    sl.cell(row=r, column=1, value=o)
    sl.cell(row=r, column=2, value=id_by_name[name])
    lookups = {3: "B", 4: "C", 5: "D", 6: "H", 7: "I", 8: "K", 9: "L", 10: "M", 11: "T", 12: "U"}
    for ci, col in lookups.items():
        sl.cell(row=r, column=ci, value=f"=INDEX(Prospects!${col}$2:${col}${last},MATCH($B{r},Prospects!$A$2:$A${last},0))")
    sl.cell(row=r, column=13, value=WHY.get(name, ""))
    for ci in range(1, 14):
        c = sl.cell(row=r, column=ci)
        c.font, c.alignment, c.border = BODY, WRAP, BORDER
    sl.row_dimensions[r].height = 110
sl.freeze_panes = "D5"

# --- Summary sheet
sm = wb.create_sheet("Summary")
sm["A1"] = "Summary (live counts)"
sm["A1"].font = TITLE
sm["A3"], sm["B3"] = "Priority", "Count"
for c in (sm["A3"], sm["B3"]):
    c.font, c.fill = HDR_FONT, HDR_FILL
for i, p in enumerate(["High", "Medium", "Low"], 4):
    sm.cell(row=i, column=1, value=p).font = BODY
    sm.cell(row=i, column=2, value=f'=COUNTIF(Prospects!$U$2:$U${last},A{i})').font = BODY
sm["A7"] = "Total"
sm["A7"].font = BOLD
sm["B7"] = "=SUM(B4:B6)"
sm["B7"].font = BOLD
sm["A9"] = "Meetable in person (Y)"
sm["B9"] = f'=COUNTIF(Prospects!$F$2:$F${last},"Y")'
sm["A10"] = "Airdrop-led (Y)"
sm["B10"] = f'=COUNTIF(Prospects!$G$2:$G${last},"Y")'
sm["A11"] = "Verified or Partial"
sm["B11"] = f'=COUNTIF(Prospects!$S$2:$S${last},"Verified")+COUNTIF(Prospects!$S$2:$S${last},"Partial")'
sm["A12"] = "Leads (need checking)"
sm["B12"] = f'=COUNTIF(Prospects!$S$2:$S${last},"Lead")'
for i in range(9, 13):
    sm.cell(row=i, column=1).font = BODY
    sm.cell(row=i, column=2).font = BODY

types = sorted({r["type"] for r in ROWS})
sm["A14"], sm["B14"], sm["C14"], sm["D14"], sm["E14"] = "Partner type", "High", "Medium", "Low", "Total"
for c in ("A14", "B14", "C14", "D14", "E14"):
    sm[c].font, sm[c].fill = HDR_FONT, HDR_FILL
for i, t in enumerate(types, 15):
    sm.cell(row=i, column=1, value=t).font = BODY
    for ci, p in ((2, "High"), (3, "Medium"), (4, "Low")):
        sm.cell(row=i, column=ci, value=f'=COUNTIFS(Prospects!$C$2:$C${last},$A{i},Prospects!$U$2:$U${last},"{p}")').font = BODY
    sm.cell(row=i, column=5, value=f"=SUM(B{i}:D{i})").font = BOLD
sm.column_dimensions["A"].width = 30
for col in "BCDE":
    sm.column_dimensions[col].width = 10

# --- Context sources sheet
cx = wb.create_sheet("Context")
cx["A1"] = "Market and regulatory context used in scoring"
cx["A1"].font = TITLE
CTX = [
    ("Finfluencer rules", "POJK 6/2026 requires competency certification for influencers giving crypto information; firms are responsible for their marketing partners' content; sanctions include account blocking and content removal.", S["pojk6"]),
    ("Certification route", "AFTECH and IARFC run a Verified Finfluencer programme built on POJK 6/2026.", S["aftech"]),
    ("Offshore platforms", "Law No. 4 of 2026 (in force 17 Jun 2026) lets OJK suspend or block non-compliant crypto activity, including offshore platforms serving Indonesians. No OJK licence for OKX was found in this research: legal sign-off needed before paid promotion.", "https://www.hsfkramer.com/ru/notes/indonesia/2026-posts/indonesia-reshapes-digital-asset-and-cryptocurrency-framework-with-2026-financial-services-omnibus-law"),
    ("Licensed list", "OJK whitelist of licensed crypto platforms (29 in Dec 2025).", "https://ojk.go.id/en/berita-dan-kegiatan/siaran-pers/Pages/OJK-Licensed-Registered-Digital-Financial-Assets-and-Crypto-Assets-Trading-Operators-Whitelist.aspx"),
    ("Trader behaviour", "Pump.fun activity fell ~80% by Jun 2026 as traders moved from memecoins to perps.", S["pump80"]),
    ("Memecoin demand", "Indonesia consistently in CoinGecko's top 10 countries for memecoin interest (2025).", "https://investortrust.id/market/89050/laporan-coingecko-indonesia-konsisten-masuk-10-besar-negara-dengan-minat-memecoin-tertinggi"),
    ("OKX DEX referral", "Referrers earn a share of invitees' UI service fees, paid to the OKX Web3 wallet; tier tables differ by region page, so confirm in the dashboard.", "https://web3.okx.com/help/whats-dex-referral-program"),
    ("OKX Web3 Affiliate", "KOL programme paying commissions and/or retainer fees for promoting OKX Web3 products (terms dated May 2026).", "https://www.okx.com/help/okx-web3-affiliate-program"),
    ("OKX local footprint", "OKX sponsored Coinfest Asia 2024 and ran the Alpha Traders Summit in Bali (Jan 2025, 200+ traders and affiliates).", S["okx_alpha_summit"]),
]
for i, h in enumerate(["Topic", "Finding", "Source"], 1):
    c = cx.cell(row=3, column=i, value=h)
    c.font, c.fill = HDR_FONT, HDR_FILL
for i, (t, f, u) in enumerate(CTX, 4):
    for ci, v in enumerate((t, f, u), 1):
        c = cx.cell(row=i, column=ci, value=v)
        c.font, c.alignment = BODY, WRAP
    cx.cell(row=i, column=3).hyperlink = u
    cx.row_dimensions[i].height = 48
cx.column_dimensions["A"].width = 20
cx.column_dimensions["B"].width = 80
cx.column_dimensions["C"].width = 60

wb.move_sheet("Prospects", offset=-1)
wb.active = 0
wb.save(OUT / "prospects.xlsx")

# ---------------------------------------------------------------- CSV
with open(OUT / "prospects.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow([h for h, _ in COLS])
    ranks = sorted((score(r["scores"]) for r in ROWS), reverse=True)
    for r in ROWS:
        sc = score(r["scores"])
        w.writerow([r["id"], r["name"], r["type"], r["link"], r["base"], r["inperson"], r["airdrop"],
                    r["audience"], r["evidence"], r["assumptions"], r["route"], r["proposal"], r["risk"],
                    *r["scores"], r["ver"], sc, priority(r), ranks.index(sc) + 1, r["shortlist"],
                    r["reason"], r["sources"]])

from collections import Counter
print("rows:", len(ROWS))
print("priority:", Counter(priority(r) for r in ROWS))
for r in ROWS[:15]:
    print(r["id"], score(r["scores"]), priority(r), r["shortlist"], r["name"])
