# Supplier Scorecard Dashboard

Interactive scorecard that ranks suppliers on a weighted composite of on-time delivery, fill rate, defect rate, invoice accuracy and price vs. agreed price. Visitors can set their own weights (or turn KPIs off) on the page. It flags high-spend, low-performing suppliers and tracks score trends over time.

**Data:** fictional office-supply distributor, 12 suppliers, ~2,000 purchase orders over 12 months (generated, seeded for reproducibility). Visitors can score their own suppliers by downloading the template (.xlsx or .csv, generated in the browser to match their weights), filling it in and uploading it; the file is read in the browser only.

## Files
- `index.html` – the dashboard (one self-contained file; host anywhere, e.g. GitHub Pages)
- `generate_data.py` – generates PO-level sample data and monthly supplier rollups
- `score.py` – scoring logic in Python, used to independently verify the dashboard's numbers (`python score.py 6`)
- `template.html` + `build.py` – dashboard source; `python build.py` embeds the data into `index.html`
- `data/purchase_orders.csv`, `data/supplier_monthly.csv`
- `tests/roundtrip.py` + `tests/browser_upload.js` – round-trip test in headless Chromium: sets weights on the page, downloads default and custom templates, checks them in Python (valid xlsx, columns, weights, worked example), then uploads them plus filled-in and broken files and compares the page's scores with `score.py`

## Scoring
Each KPI earns 0–100 points on a straight line from a floor to a target, then is weighted. Default weights: on-time 30%, fill rate 20%, defects 20%, invoice accuracy 15%, price vs. agreed 15% (changeable on the page). Tiers: Preferred ≥ 85, Approved 70–85, Watch 55–70, Probation < 55.

## Weights and the template
The weights table in "How the score works" sets each KPI's weight (must total 100%) or turns it off; scores update right away and the choice is remembered in the browser. "Download template" builds the workbook on the spot (own small .xlsx writer, no library): the Data tab holds only the columns the enabled KPIs need plus 9 labeled SAMPLE rows; Instructions explains each column and unit; KPI definitions shows the chosen weights, formulas, sources and a worked example; Settings stores the weights, which the page reads back when the .xlsx is uploaded. `agreed_value` (qty × contract, quote or PO price; older files' `contract_value` is accepted) is optional: price is measured only on rows that have it, and a supplier with none is scored on its other KPIs, reweighted to 100%.

## Upload validation
Uploads are checked before scoring and every problem is listed with its spreadsheet row: missing columns, blank or non-numeric values, negative numbers, fractional counts, counts above their base (e.g. `on_time_pos` > `pos`), months that aren't months, duplicate supplier + month rows, and suppliers whose totals are zero. Months may be typed as 2026-06, 2026-06-01, 6/1/2026, Jun 2026 or be real Excel dates. A warning appears if template SAMPLE rows are still in the file.

## Sources for the KPI definitions
APICS (2017) SCOR v12, Perfect Order Fulfillment components; Mena, van Hoek & Christopher (2021) *Leading Procurement Strategy*, 3rd ed., Kogan Page, Table 3.2 and pp. 70-71; Weber, Current & Benton (1991) *EJOR* 50(1); Ho, Xu & Dey (2010) *EJOR* 202(1); Gunasekaran, Patel & McGaughey (2004) *IJPE* 87(3); Dickson (1966) *Journal of Purchasing* 2(1); Monczka et al. (2021) *Purchasing and Supply Chain Management*, 7th ed.; Datar & Rajan (2021) *Horngren's Cost Accounting*, 17th ed. Full references are on the template's Sources tab. Weights, floors, targets and tier cut-offs are this tool's defaults, not published benchmarks.

## Rebuild and test
```
python generate_data.py && python build.py
python score.py 6 --weights otd=40,fill=0,quality=30,invoice=10,price=20   # check any weights
python tests/roundtrip.py path/to/xlsx.full.min.js   # SheetJS 0.18.5, the version the page loads
```
