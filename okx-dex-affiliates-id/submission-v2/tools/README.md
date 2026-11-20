# Rebuild tools

Run from `okx-dex-affiliates-id/submission-v2/`:

```bash
python tools/build_workbooks.py tools/legacy_model.xlsx tools/research_sheet.xlsx REVIEW.md .
```

Writes both workbooks (submission and research). Partner inputs (reach pool P, FnF power traders, share, scores, go-live day and all partner text) are in the `partners` list at the top of `build_workbooks.py`; the rates (sign-up, activation, swaps per FTD and per FnF trader, low/base/high) are yellow cells in the Miscellaneous tab, section A. `research_sheet.xlsx` is an export of the Indonesia Crypto KOL & Partner Research sheet, used for the Research merge tab. Open each workbook in Excel once so formulas recalculate.

```bash
python tools/make_deck.py deck_v2.html OKX_DEX_BD_Affiliate_Indonesia_Prospect_Spreadsheet.xlsx
```

Rewrites the six slides of `deck_v2.html` from the recalculated submission workbook, so every number on a slide matches a cell. Recalculate the workbook first (Excel, or LibreOffice headless).

```bash
python tools/html_to_md.py deck_v2.html deck_slides_v2.md
```

Regenerates the deck text from the HTML deck.
