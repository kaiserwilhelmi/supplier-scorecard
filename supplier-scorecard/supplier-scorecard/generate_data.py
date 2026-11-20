"""
Generate sample purchase-order data for a fictional office-supply distributor
and roll it up into monthly supplier metrics for the scorecard dashboard.

All suppliers and numbers are fictional. Seeded so results are reproducible.

Outputs:
  data/purchase_orders.csv  - one row per PO (the raw data)
  data/supplier_monthly.csv - one row per supplier per month (what the dashboard reads)
"""
import csv
import random
from datetime import date
from pathlib import Path

random.seed(42)
OUT = Path(__file__).parent / "data"
OUT.mkdir(exist_ok=True)

MONTHS = [date(2025, 12, 1)] + [date(2026, m, 1) for m in range(1, 12)]

# name, category, avg monthly spend, base on-time %, base fill %, base defect %,
# base invoice accuracy %, base price variance % (vs contract), lead time days,
# last field: optional (start_month_index, end-state otd, fill, defect, inv, pv) -
# metrics move linearly from base to end-state from that month onward
SUPPLIERS = [
    ("Harbor Line Paper",          "Paper",          21000, 96.5, 98.8, 0.4, 99.2, 0.2,  4, None),
    ("Clearwater Ink & Toner",     "Toner & Ink",    16500, 97.0, 97.5, 0.9, 98.5, 0.6,  5, (5, 84.0, 94.0, 1.6, 97.0, 2.6)),
    ("Mesa Office Furniture",      "Furniture",      12500, 88.0, 95.0, 1.8, 96.5, 1.5, 14, None),
    ("Granite Seating Co.",        "Furniture",       7200, 91.0, 96.0, 1.2, 97.8, 0.9, 12, None),
    ("Northgate Janitorial",       "Janitorial",      6100, 93.5, 97.0, 0.8, 99.0, 0.4,  3, None),
    ("Brightmark Writing Goods",   "Writing",         4800, 98.0, 99.2, 0.3, 99.6, -0.3, 4, None),
    ("Summit Tech Accessories",    "Technology",      9400, 82.0, 92.5, 2.4, 95.0, 2.2,  7, (0, 95.5, 97.5, 0.8, 98.8, 0.6)),
    ("Redwood Breakroom Supply",   "Breakroom",       3900, 94.0, 96.5, 0.7, 98.8, 0.1,  3, None),
    ("Ironleaf Filing Systems",    "Filing",          3100, 90.0, 94.0, 1.4, 97.0, 1.1,  6, None),
    ("Bayview Packaging",          "Packaging",       5300, 95.5, 98.0, 0.6, 99.1, 0.3,  4, None),
    ("Sunridge Labels & Tape",     "Packaging",       2200, 86.0, 93.0, 2.0, 96.0, 1.8,  6, None),
    ("Coastal Toner Remanufacturing", "Toner & Ink",  4300, 92.0, 95.5, 1.6, 97.5, -1.2, 6, None),
]

SEASON = {12: 0.85, 1: 1.15, 2: 1.0, 3: 1.05, 4: 0.95, 5: 0.95,
          6: 0.9, 7: 1.1, 8: 1.25, 9: 1.05, 10: 0.95, 11: 0.9}

def clamp(x, lo, hi):
    return max(lo, min(hi, x))

po_rows, monthly_rows = [], []
po_num = 40100

for (name, cat, spend, otd0, fill0, defect0, inv0, pv0, lead, trend) in SUPPLIERS:
    for i, m in enumerate(MONTHS):
        month_spend_target = spend * SEASON[m.month] * random.uniform(0.85, 1.15)
        n_pos = max(4, round(month_spend_target / random.uniform(450, 750)))
        otd, fill, defect, inv, pv = otd0, fill0, defect0, inv0, pv0
        if trend:
            start, *end = trend
            prog = clamp((i - start) / (len(MONTHS) - 1 - start), 0, 1)
            otd, fill, defect, inv, pv = [b + (e - b) * prog for b, e in
                                          zip((otd0, fill0, defect0, inv0, pv0), end)]
        p_otd = clamp(otd + random.gauss(0, 1.5), 55, 99.8) / 100
        agg = dict(spend=0.0, pos=0, on_time=0, lines=0, lines_filled=0,
                   units=0, defective=0, invoices=0, invoices_ok=0,
                   contract_value=0.0, lead_days=0)
        for _ in range(n_pos):
            po_num += 1
            amount = month_spend_target / n_pos * random.uniform(0.6, 1.4)
            lines = random.randint(3, 18)
            filled = sum(random.random() < clamp(fill + random.gauss(0, 1), 80, 100) / 100
                         for _ in range(lines))
            units = lines * random.randint(5, 40)
            defective = sum(1 for _ in range(units) if random.random() < defect / 100)
            on_time = random.random() < p_otd
            lead_days = max(1, round(random.gauss(lead, lead * 0.2)) + (0 if on_time else random.randint(1, 5)))
            invoice_ok = random.random() < clamp(inv + random.gauss(0, 0.8), 85, 100) / 100
            var = random.gauss(pv, 0.8)
            contract = amount / (1 + var / 100)
            po_rows.append(dict(po_number=f"PO-{po_num}", supplier=name, category=cat,
                                month=m.isoformat()[:7], amount=round(amount, 2),
                                contract_value=round(contract, 2), lines_ordered=lines,
                                lines_filled=filled, units_received=units,
                                units_defective=defective, on_time=int(on_time),
                                lead_time_days=lead_days, invoice_accurate=int(invoice_ok)))
            agg["spend"] += amount; agg["pos"] += 1; agg["on_time"] += on_time
            agg["lines"] += lines; agg["lines_filled"] += filled
            agg["units"] += units; agg["defective"] += defective
            agg["invoices"] += 1; agg["invoices_ok"] += invoice_ok
            agg["contract_value"] += contract; agg["lead_days"] += lead_days
        monthly_rows.append(dict(supplier=name, category=cat, month=m.isoformat()[:7],
                                 spend=round(agg["spend"], 2),
                                 contract_value=round(agg["contract_value"], 2),
                                 pos=agg["pos"], on_time_pos=agg["on_time"],
                                 lines_ordered=agg["lines"], lines_filled=agg["lines_filled"],
                                 units_received=agg["units"], units_defective=agg["defective"],
                                 invoices=agg["invoices"], invoices_accurate=agg["invoices_ok"],
                                 lead_time_days_total=agg["lead_days"]))

for fname, rows in (("purchase_orders.csv", po_rows), ("supplier_monthly.csv", monthly_rows)):
    with open(OUT / fname, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader(); w.writerows(rows)

print(f"{len(po_rows)} POs, {len(monthly_rows)} supplier-months, "
      f"total spend ${sum(r['amount'] for r in po_rows):,.0f}")
