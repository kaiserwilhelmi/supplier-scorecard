# Supplier Scorecard Dashboard

Interactive scorecard that ranks suppliers on a weighted composite of on-time delivery, fill rate, defect rate, invoice accuracy and price vs. contract. It flags high-spend, low-performing suppliers and tracks score trends over time.

**Data:** fictional office-supply distributor, 12 suppliers, ~2,000 purchase orders over 12 months (generated, seeded for reproducibility). Visitors can paste their own CSV in the same format.

## Files
- `index.html` – the dashboard (one self-contained file; host anywhere, e.g. GitHub Pages)
- `generate_data.py` – generates PO-level sample data and monthly supplier rollups
- `score.py` – scoring logic in Python, used to independently verify the dashboard's numbers (`python score.py 6`)
- `template.html` + `build.py` – dashboard source; `python build.py` embeds the data into `index.html`
- `data/purchase_orders.csv`, `data/supplier_monthly.csv`

## Scoring
Each KPI earns 0–100 points on a straight line from a floor to a target, then is weighted: on-time 30%, fill rate 20%, defects 20%, invoice accuracy 15%, price variance 15%. Tiers: Preferred ≥ 85, Approved 70–85, Watch 55–70, Probation < 55.

## Rebuild
```
python generate_data.py && python build.py
```
