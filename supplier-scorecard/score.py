"""
Supplier scorecard scoring - same rules the dashboard uses, so results can be
checked independently. Usage: python score.py [months]   (default 12)

Each KPI is scored 0-100 on a straight line between a floor (0 pts) and a
target (100 pts), then weighted into a composite score.
"""
import csv
import sys
from collections import defaultdict
from pathlib import Path

KPIS = {  # key: (label, weight, floor, target)  - lower-is-better when floor > target
    "otd":     ("On-time delivery %",   0.30, 80.0, 97.0),
    "fill":    ("Fill rate %",          0.20, 88.0, 99.0),
    "quality": ("Defect rate %",        0.20,  3.0,  0.3),
    "invoice": ("Invoice accuracy %",   0.15, 92.0, 99.5),
    "price":   ("Price variance %",     0.15,  3.0,  0.0),
}
TIERS = [(85, "Preferred"), (70, "Approved"), (55, "Watch"), (0, "Probation")]

def kpi_points(value, floor, target):
    pts = (value - floor) / (target - floor) * 100
    return max(0.0, min(100.0, pts))

def metrics(rows):
    s = lambda k: sum(float(r[k]) for r in rows)
    return {
        "spend": s("spend"),
        "otd": s("on_time_pos") / s("pos") * 100,
        "fill": s("lines_filled") / s("lines_ordered") * 100,
        "quality": s("units_defective") / s("units_received") * 100,
        "invoice": s("invoices_accurate") / s("invoices") * 100,
        "price": (s("spend") / s("contract_value") - 1) * 100,
        "lead": s("lead_time_days_total") / s("pos"),
    }

def score(m):
    return sum(w * kpi_points(m[k], f, t) for k, (_, w, f, t) in KPIS.items())

def tier(sc):
    return next(name for cut, name in TIERS if sc >= cut)

if __name__ == "__main__":
    n = int(sys.argv[1]) if len(sys.argv) > 1 else 12
    rows = list(csv.DictReader(open(Path(__file__).parent / "data/supplier_monthly.csv")))
    months = sorted({r["month"] for r in rows})[-n:]
    by = defaultdict(list)
    for r in rows:
        if r["month"] in months:
            by[r["supplier"]].append(r)
    out = sorted(((sup, metrics(rs)) for sup, rs in by.items()), key=lambda x: -score(x[1]))
    print(f"Last {n} months: {months[0]} to {months[-1]}")
    for sup, m in out:
        sc = score(m)
        print(f"{sup:32s} {sc:5.1f} {tier(sc):10s} spend ${m['spend']:>9,.0f}  "
              f"OTD {m['otd']:5.1f}  fill {m['fill']:5.1f}  def {m['quality']:4.2f}  "
              f"inv {m['invoice']:5.1f}  pv {m['price']:5.2f}  lead {m['lead']:4.1f}")
