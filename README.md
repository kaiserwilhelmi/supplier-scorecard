# Supplier Scorecard Dashboard

Interactive scorecard that ranks suppliers on a weighted composite of on-time delivery, fill rate, defect rate, invoice accuracy and price vs. contract. It flags high-spend, low-performing suppliers and tracks score trends over time.

**Data:** fictional office-supply distributor, 12 suppliers, ~2,000 purchase orders over 12 months (generated, seeded for reproducibility). Visitors can score their own suppliers by filling in the downloadable template (`template/`, .xlsx or .csv) and uploading it; the file is read in the browser only.

## Files
- `index.html` – the dashboard (one self-contained file; host anywhere, e.g. GitHub Pages)
- `generate_data.py` – generates PO-level sample data and monthly supplier rollups
- `score.py` – scoring logic in Python, used to independently verify the dashboard's numbers (`python score.py 6`)
- `template.html` + `build.py` – dashboard source; `python build.py` embeds the data into `index.html`
- `data/purchase_orders.csv`, `data/supplier_monthly.csv`
- `template/supplier-scorecard-template.xlsx` – upload template: Data tab (exact columns + 9 labeled SAMPLE rows, input validation), Instructions (every field, unit and upload rule), KPI definitions (formulas, weights, floors/targets, rationale, worked example) and Sources
- `template/supplier-scorecard-template.csv` – the same Data columns and sample rows as plain CSV
- `make_template.py` – builds both template files from the column list and the weights in `score.py`
- `tests/roundtrip.py` + `tests/browser_upload.js` – round-trip test: uploads the template, a filled-in copy and broken files through the real page in headless Chromium and checks scores against `score.py` and the error messages

## Scoring
Each KPI earns 0–100 points on a straight line from a floor to a target, then is weighted: on-time 30%, fill rate 20%, defects 20%, invoice accuracy 15%, price variance 15%. Tiers: Preferred ≥ 85, Approved 70–85, Watch 55–70, Probation < 55.

## Upload validation
Uploads are checked before scoring and every problem is listed with its spreadsheet row: missing columns, blank or non-numeric values, negative numbers, fractional counts, counts above their base (e.g. `on_time_pos` > `pos`), months that aren't months, duplicate supplier + month rows, and suppliers whose totals are zero. Months may be typed as 2026-06, 2026-06-01, 6/1/2026, Jun 2026 or be real Excel dates. A warning appears if template SAMPLE rows are still in the file.

## Sources for the KPI definitions
APICS (2017) SCOR v12, Perfect Order Fulfillment components; Mena, van Hoek & Christopher (2021) *Leading Procurement Strategy*, 3rd ed., Kogan Page, Table 3.2 and pp. 70-71; Weber, Current & Benton (1991) *EJOR* 50(1); Ho, Xu & Dey (2010) *EJOR* 202(1); Gunasekaran, Patel & McGaughey (2004) *IJPE* 87(3); Dickson (1966) *Journal of Purchasing* 2(1); Monczka et al. (2021) *Purchasing and Supply Chain Management*, 7th ed.; Datar & Rajan (2021) *Horngren's Cost Accounting*, 17th ed. Full references are on the template's Sources tab. Weights, floors, targets and tier cut-offs are this tool's defaults, not published benchmarks.

## Rebuild and test
```
python generate_data.py && python make_template.py && python build.py
python tests/roundtrip.py path/to/xlsx.full.min.js   # SheetJS 0.18.5, the version the page loads
```
