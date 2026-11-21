"""
Round-trip check for the upload template.

1. Python: the template's Data columns equal the columns the page's parser requires (read from template.html),
   and the 'Worked example' on the KPI sheet equals score.py on the sample rows.
2. Browser: builds filled-in and broken fixtures, uploads each one through the real page in headless Chromium
   (tests/browser_upload.js) and checks the page's scores against score.py, and that bad files are rejected
   with row-level messages.

Usage: python tests/roundtrip.py <path to xlsx.full.min.js 0.18.5>
"""
import csv
import datetime as dt
import json
import re
import subprocess
import sys
import tempfile
from collections import defaultdict
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
from make_template import COLUMNS, SAMPLE  # noqa: E402
from score import metrics, score, tier  # noqa: E402

XLSX = ROOT / "template/supplier-scorecard-template.xlsx"
CSV = ROOT / "template/supplier-scorecard-template.csv"
fails = []


def check(cond, what):
    print(("PASS " if cond else "FAIL ") + what)
    if not cond:
        fails.append(what)


def py_scores(rows, months):
    ms = sorted({r["month"] for r in rows})[-months:]
    by = defaultdict(list)
    for r in rows:
        if r["month"] in ms:
            by[r["supplier"]].append(r)
    return {s: round(score(metrics(rs)), 1) for s, rs in by.items()}


# ---------- 1. static checks
html = (ROOT / "template.html").read_text()
num = json.loads(re.search(r"const NUM = (\[.*?\]);", html).group(1))
required = ["supplier", "category", "month", *num]
wb = load_workbook(XLSX)
check(wb.sheetnames[0] == "Data", "xlsx: first sheet is 'Data'")
ws = wb["Data"]
head = [c.value for c in ws[1] if c.value is not None]
check(head == required, "xlsx: Data header equals the parser's required columns, in order")
with open(CSV, newline="") as f:
    crow = list(csv.reader(f))
check(crow[0] == required, "csv: header equals the parser's required columns")
xrows = [dict(zip(head, [c.value for c in r])) for r in ws.iter_rows(min_row=2) if r[0].value]
crows = [dict(zip(crow[0], r)) for r in crow[1:]]
check(len(xrows) == len(crows) == len(SAMPLE) == 9, "xlsx and csv each hold the 9 sample rows")
check(all(r["supplier"].startswith("SAMPLE - ") for r in xrows + crows), "every sample row is labeled 'SAMPLE - '")
check(all(str(x[k]) == str(c[k]) or float(x[k]) == float(c[k]) for x, c in zip(xrows, crows) for k in required), "xlsx and csv rows are identical")
check(all(ws.cell(i, 3).number_format == "@" for i in range(2, 1002)), "month column pre-formatted as text (rows 2-1001)")
kp = wb["KPI definitions"]
weights = [kp.cell(i, 3).value for i in range(5, 10)]
check(abs(sum(weights) - 1) < 1e-9, f"KPI weights sum to 100% ({weights})")
ex = {}
for row in kp.iter_rows(values_only=True):
    if row[0] and str(row[0]).startswith("SAMPLE - ") and isinstance(row[1], (int, float)):
        ex[row[0]] = (row[1], row[2])
expect = {s: (v, tier(v)) for s, v in py_scores(xrows, 3).items()}
check(ex == expect, f"worked example on KPI sheet equals score.py {expect}")

# ---------- 2. fixtures
tmp = Path(tempfile.mkdtemp())
rest = list(csv.DictReader(open(ROOT / "data/mock_restaurant_suppliers.csv")))

# C: filled-in xlsx, done the way a user would: open the template, delete SAMPLE rows, type own rows.
#    Some months typed so Excel turned them into real dates (shown as Jun-26), to prove those still load.
wb = load_workbook(XLSX)
ws = wb["Data"]
ws.delete_rows(2, len(SAMPLE))
for i, r in enumerate(rest, 2):
    for j, k in enumerate(COLUMNS, 1):
        v = r[k]
        if k not in ("supplier", "category", "month"):
            v = float(v) if "." in v else int(v)
        ws.cell(i, j, v)
    if i % 5 == 0:
        y, m = map(int, r["month"].split("-"))
        c = ws.cell(i, 3, dt.datetime(y, m, 1))
        c.number_format = "mmm-yy"
filled_xlsx = tmp / "C_filled_template.xlsx"
wb.save(filled_xlsx)

# D: same data saved by Excel as CSV (UTF-8 BOM, CRLF, quoted thousands, header in Title Case with spaces)
filled_csv = tmp / "D_filled_excel.csv"
with open(filled_csv, "w", newline="", encoding="utf-8-sig") as f:
    w = csv.writer(f, lineterminator="\r\n")
    w.writerow([k.replace("_", " ").title() for k in COLUMNS])
    for r in rest:
        w.writerow([f"{float(r[k]):,.2f}" if k in ("spend", "contract_value") else r[k] for k in COLUMNS])

# E: broken file (bad number, impossible count, bad month, duplicate row)
wb = load_workbook(XLSX)
ws = wb["Data"]
ws["D3"] = "about 13k"     # row 3 spend
ws["G4"] = 30              # row 4 on_time_pos 30 > pos 24? (row 4 pos is 22) -> error
ws["C6"] = "Q2 2026"       # row 6 month
for j in range(1, 15):     # row 10 duplicates row 9
    ws.cell(10, j, ws.cell(9, j).value)
bad_xlsx = tmp / "E_errors.xlsx"
wb.save(bad_xlsx)

# F: wrong headers
bad_csv = tmp / "F_wrong_headers.csv"
bad_csv.write_text("vendor,month,spend\nAcme,2026-06,100\n")

files = [XLSX, CSV, filled_xlsx, filled_csv, bad_xlsx, bad_csv]
out = json.loads(subprocess.run(["node", str(ROOT / "tests/browser_upload.js"), str(ROOT), sys.argv[1], *map(str, files)],
                                capture_output=True, text=True, check=True).stdout)
res = {r["file"]: r for r in out["results"]}
check(all(v == 200 for v in out["downloads"].values()), f"both download links resolve {out['downloads']}")
check(out["detailsOpen"], "'Score your own suppliers' button opens the upload panel")
check(not out["pageErrors"], f"no page errors {out['pageErrors']}")


def shown(rows):
    return {r["name"]: float(r["cells"][2]) for r in rows}


for name, data in (("supplier-scorecard-template.xlsx", xrows), ("supplier-scorecard-template.csv", crows)):
    r = res[name]
    check("ok" in r["msg"]["cls"] and "9 rows, 3 suppliers, 3 months" in r["msg"]["text"], f"{name}: loads ({r['msg']['text'].splitlines()[0]})")
    check("still the template's SAMPLE data" in r["msg"]["text"], f"{name}: page warns the SAMPLE rows are still there")
    check(shown(r["m3"]) == py_scores(data, 3), f"{name}: page scores equal score.py {shown(r['m3'])}")
    check(r["badge"] == "Your data", f"{name}: badge switches to 'Your data'")

for name in ("C_filled_template.xlsx", "D_filled_excel.csv"):
    r = res[name]
    check("ok" in r["msg"]["cls"] and f"{len(rest)} rows" in r["msg"]["text"] and "SAMPLE" not in r["msg"]["text"], f"{name}: loads ({r['msg']['text']})")
    for m in (3, 6):
        check(shown(r[f"m{m}"]) == py_scores(rest, m), f"{name}: last {m} months scores equal score.py")

e = res["E_errors.xlsx"]["msg"]
check("err" in e["cls"], "E_errors.xlsx: rejected")
for frag in ('Row 3 (SAMPLE - Northwind Fasteners): spend "about 13k" isn\'t a number',
             "Row 4 (SAMPLE - Northwind Fasteners): on_time_pos (30) is more than pos (22)",
             'Row 6 (SAMPLE - Bayview Packaging): month "Q2 2026" isn\'t a month',
             "Row 10 (SAMPLE - Ridgeway Industrial): duplicate of row 9"):
    check(frag in e["text"], f"E_errors.xlsx: message lists '{frag}'")
f = res["F_wrong_headers.csv"]["msg"]
check("err" in f["cls"] and "Missing columns: supplier, category, contract_value" in f["text"], f"F_wrong_headers.csv: rejected ({f['text'][:80]}...)")

print(f"\n{'ALL PASS' if not fails else str(len(fails)) + ' FAILED'}")
sys.exit(1 if fails else 0)
