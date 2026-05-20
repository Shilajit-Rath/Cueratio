# CI Benchmarking → CODIFY Converter

Automated conversion of insurance Critical Illness benchmarking workbooks into the
CODIFY COMPARISON structure. Drop a benchmarking file in, get a codified
multi-sheet workbook out.

## What you get

- **`codify.py`** — command-line script. Run it from any terminal.
- **`app.py`** — drag-and-drop web app. Run locally, open in browser, upload, download.

Both share the same conversion engine. The web app is just a UI wrapper over the script.

## Install

Requires Python 3.10 or newer.

```bash
pip install -r requirements.txt
```

That installs `openpyxl` (Excel handling) and `flask` (web app).

## Usage — Command line

```bash
# Auto-name output: writes "<input>_Codified.xlsx" next to the input
python codify.py path/to/Philippines_CI_Benchmarking.xlsx

# Explicit output path
python codify.py input.xlsx output.xlsx
```

Output:
```
→ Reading:  Philippines_CI_Benchmarking.xlsx
→ Writing:  Philippines_CI_Benchmarking_Codified.xlsx

✓ Done.
  Companies: 10
  Products:  11
  CI conditions: 231
  Output sheets: 38
```

## Usage — Web app

```bash
python app.py
```

Open `http://127.0.0.1:5000` in your browser. Drag the file onto the drop zone,
click **Convert**, and the codified file downloads automatically.

To share on a network, set host:
```bash
# Edit the last line of app.py to:
# app.run(host='0.0.0.0', port=5000)
```

## What's auto-detected

The script doesn't depend on exact sheet names or row positions. It looks for:

1. **Feature Benchmarking sheet:** any sheet (except the Conditions sheet) where
   a cell contains the word `Company` near another cell containing `Products`,
   with product names spread across columns to the right.

2. **Condition Benchmarking sheet:** any sheet with a header column called
   `Condition Name` (plus optionally `Major/Minor`, `Stage`, `Condition Category`).

If detection fails, the script reports a clear error rather than guessing.

## Output structure

| Sheet | Purpose | Source-coverage |
|-------|---------|----------------|
| `00_Index` | Summary with tab-color legend and per-sheet ratings | – |
| `01_Company_Master` | Company IDs and short names | Full |
| `02_Product_Master` | Product IDs | Full |
| `03_Variant_Master` | Variants (one V01 per product) | Full |
| `04_Entry_Age` | Min/max ages with day/month/year buckets | Good |
| `05_Policy_Term` | Term value + unit | Good |
| `06_Coverage` | Coverage max age, renewal eligibility | Good |
| `07A–07E` | Waiting periods | Standard parsed; specifics flagged as NA |
| `08_Policy_Year_Benefit_Amt` | Max SI | Partial |
| `09–27` | Hospital/IPD/OPD benefits | NA (not in CI source) |
| `28_Critical_Illness_Benefits` | Core CI codification | **Strong** |
| `28A_CI_Conditions_Coverage` | Long-format coverage matrix | Full |
| `29_Misc_&_Extra_Benefits` | Death/TPD/Waiver/Hospital Income/etc. | **Strong** |
| `30–32` | Deductible / Copay / NCB | NA |

Every codified sheet keeps a `raw_text` column with the original free-text value,
so you can always trace back to the source.

## Customising / extending

- **Add a new short-name mapping:** edit `COMPANY_SHORT_MAP` at the top of `codify.py`.
- **Change which feature keys feed which output column:** look at the
  `feat(p['features'], 'Key1', 'Key2', ...)` calls. They accept multiple candidate
  keys so you can support naming variations.
- **Add a new benefit row to sheet 29:** append a tuple to `EXTRA_BENEFITS`:
  `('Display Name', 'Description Key', 'Amount Key', 'Conditions Key')`.

## Troubleshooting

**"Could not auto-detect a Feature Benchmarking sheet"**
The source needs literal `Company` and `Products` labels somewhere in the first
few rows of cells A–D. Add them if missing, or contact the maintainer to extend
the heuristic.

**"No products were detected"**
The label cells were found but the company/product rows are empty. Check that
product names sit in the columns immediately to the right of the labels.

**Companies showing wrong short names**
The fallback heuristic trims `Inc.`, `Ltd`, `(...)` etc. For better names, add an
entry to `COMPANY_SHORT_MAP` (lowercase key).

## License

Provided as-is for internal use.
