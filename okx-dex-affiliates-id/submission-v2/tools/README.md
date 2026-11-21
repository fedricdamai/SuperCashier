# Rebuild tools

Run from `okx-dex-affiliates-id/submission-v2/`.

1. Workbooks (submission and research):

```bash
python tools/build_workbooks.py tools/legacy_model.xlsx tools/research_sheet.xlsx REVIEW.md .
```

Partner inputs are in the `partners` list at the top of `build_workbooks.py`; rates are yellow cells in the Miscellaneous tab (section A) and the Task 2 deal terms are yellow cells in the Task 2 tab (rows 6-20). Recalculate once (open in Excel, or LibreOffice headless) before the next steps, which read calculated values.

2. Write-up (.docx, 4 pages):

```bash
python tools/make_writeup.py OKX_DEX_BD_Affiliate_Indonesia_Prospect_Spreadsheet.xlsx /tmp/writeup.json
node tools/build_writeup.js /tmp/writeup.json OKX_DEX_BD_Affiliate_Indonesia_Writeup.docx 11
```

3. Slides (.pptx, 6 slides; needs `npm install pptxgenjs@3.12.0`):

```bash
python tools/slides_data.py OKX_DEX_BD_Affiliate_Indonesia_Prospect_Spreadsheet.xlsx /tmp/slides.json
node tools/build_slides.js /tmp/slides.json OKX_DEX_BD_Affiliate_Indonesia_Slides.pptx
```

`research_sheet.xlsx` is an export of the Indonesia Crypto KOL & Partner Research sheet, used for the Research merge tab.
