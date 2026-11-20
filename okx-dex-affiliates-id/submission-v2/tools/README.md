# Rebuild tools

Run from `okx-dex-affiliates-id/submission-v2/`:

```bash
python tools/build_workbooks.py tools/legacy_model.xlsx ../prospects.csv REVIEW.md .
```

Writes both workbooks (submission and research). Partner inputs for GuarEmperor and Superteam Indonesia are in the `partners.insert(...)` and `partners.append(...)` blocks of `build_workbooks.py`; the other six partners' text comes from `legacy_model.xlsx`. Open each workbook in Excel once so formulas recalculate.

```bash
python tools/html_to_md.py deck_v2.html deck_slides_v2.md
```

Regenerates the deck text from the HTML deck.
