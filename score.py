"""
Supplier scorecard scoring - same rules the dashboard uses, so results can be
checked independently.

Usage: python score.py [months] [--weights otd=30,fill=20,quality=20,invoice=15,price=15] [--file data.csv]

Each KPI is scored 0-100 on a straight line between a floor (0 pts) and a
target (100 pts), then weighted into a composite score. A weight of 0 turns a KPI
off. agreed_value is optional: price is measured only on rows that have it, and a
supplier with none is scored on its other KPIs, reweighted to 100%.
"""
import csv
import sys
from collections import defaultdict
from pathlib import Path

KPIS = {  # key: (label, default weight, floor, target)  - lower-is-better when floor > target
    "otd":     ("On-time delivery %",  0.30, 80.0, 97.0),
    "fill":    ("Fill rate %",         0.20, 88.0, 99.0),
    "quality": ("Defect rate %",       0.20,  3.0,  0.3),
    "invoice": ("Invoice accuracy %",  0.15, 92.0, 99.5),
    "price":   ("Price vs. agreed %",  0.15,  3.0,  0.0),
}
DEFAULT_WEIGHTS = {k: v[1] for k, v in KPIS.items()}
TIERS = [(85, "Preferred"), (70, "Approved"), (55, "Watch"), (0, "Probation")]


def num(v):
    if v is None or (isinstance(v, str) and v.strip() == ""):
        return None
    return float(str(v).replace(",", "").replace("$", ""))


def kpi_points(value, floor, target):
    pts = (value - floor) / (target - floor) * 100
    return max(0.0, min(100.0, pts))


def metrics(rows, weights=None):
    w = DEFAULT_WEIGHTS if weights is None else weights

    def s(k):
        return sum(num(r[k]) for r in rows)

    m = {"spend": s("spend"), "pos": s("pos")}
    if w.get("otd"):
        m["otd"] = s("on_time_pos") / s("pos") * 100
    if w.get("fill"):
        m["fill"] = s("lines_filled") / s("lines_ordered") * 100
    if w.get("quality"):
        m["quality"] = s("units_defective") / s("units_received") * 100
    if w.get("invoice"):
        m["invoice"] = s("invoices_accurate") / s("invoices") * 100
    if w.get("price"):
        priced = [r for r in rows if num(r.get("agreed_value")) is not None]
        agreed = sum(num(r["agreed_value"]) for r in priced)
        if agreed > 0:
            m["price"] = (sum(num(r["spend"]) for r in priced) / agreed - 1) * 100
    leads = [num(r.get("lead_time_days_total")) for r in rows]
    if all(v is not None for v in leads):
        m["lead"] = sum(leads) / m["pos"]
    return m


def score(m, weights=None):
    w = DEFAULT_WEIGHTS if weights is None else weights
    used = [k for k in KPIS if w.get(k) and k in m]
    total = sum(w[k] for k in used)
    return sum(w[k] / total * kpi_points(m[k], KPIS[k][2], KPIS[k][3]) for k in used)


def tier(sc):
    return next(name for cut, name in TIERS if sc >= cut)


def scores_by_supplier(rows, months=12, weights=None):
    ms = sorted({r["month"] for r in rows})[-months:]
    by = defaultdict(list)
    for r in rows:
        if r["month"] in ms:
            by[r["supplier"]].append(r)
    return {sup: score(metrics(rs, weights), weights) for sup, rs in by.items()}


def parse_weights(text):
    w = {k: 0.0 for k in KPIS}
    for part in text.split(","):
        k, v = part.split("=")
        w[k.strip()] = float(v) / 100
    assert abs(sum(w.values()) - 1) < 1e-9, "weights must total 100"
    return w


if __name__ == "__main__":
    args = sys.argv[1:]
    weights = parse_weights(args[args.index("--weights") + 1]) if "--weights" in args else DEFAULT_WEIGHTS
    path = Path(args[args.index("--file") + 1]) if "--file" in args else Path(__file__).parent / "data/supplier_monthly.csv"
    n = int(args[0]) if args and args[0].isdigit() else 12
    rows = list(csv.DictReader(open(path, encoding="utf-8-sig")))
    months = sorted({r["month"] for r in rows})[-n:]
    print(f"Last {n} months: {months[0]} to {months[-1]}   weights: {weights}")
    by = defaultdict(list)
    for r in rows:
        if r["month"] in months:
            by[r["supplier"]].append(r)
    out = sorted(((sup, metrics(rs, weights)) for sup, rs in by.items()), key=lambda x: -score(x[1], weights))
    for sup, m in out:
        sc = score(m, weights)
        kp = "  ".join(f"{k} {m[k]:6.2f}" for k in KPIS if k in m)
        print(f"{sup:32s} {sc:5.1f} {tier(sc):10s} spend ${m['spend']:>9,.0f}  {kp}")
