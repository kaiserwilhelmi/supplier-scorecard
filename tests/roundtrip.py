"""
Round-trip test for the downloadable template and the upload parser.

Step 0 checks that the committed index.html is a fresh build, so a template change that was
never rebuilt (no favicon or link preview on the live page, Nov 2026) fails the test.

The template is generated in the browser from the weights chosen on the page, so this test
drives the real page in headless Chromium (tests/browser_upload.js):
  1. Download the default template and a custom-weights template (fill rate off). Check the files
     in Python: valid zip/XML, columns match the enabled KPIs, weights on the KPI and Settings tabs,
     worked example equals score.py.
  2. Build filled-in and broken files from those downloads, upload each through the page, and check
     the page's scores against score.py and the error messages.

Usage: python tests/roundtrip.py   (needs Node with Playwright, and openpyxl)
"""
import csv
import datetime as dt
import json
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.dom import minidom

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
from score import DEFAULT_WEIGHTS, KPIS, scores_by_supplier, tier  # noqa: E402

TMP = Path(tempfile.mkdtemp())
CUSTOM = {"otd": 0.40, "fill": 0.0, "quality": 0.30, "invoice": 0.10, "price": 0.20}
BASE = ["supplier", "category", "month", "spend", "pos"]
KPI_COLS = {"otd": ["on_time_pos"], "fill": ["lines_ordered", "lines_filled"], "quality": ["units_received", "units_defective"],
            "invoice": ["invoices", "invoices_accurate"], "price": ["agreed_value"]}
ORDER = ["supplier", "category", "month", "spend", "pos", "agreed_value", "on_time_pos", "lines_ordered", "lines_filled",
         "units_received", "units_defective", "invoices", "invoices_accurate", "lead_time_days_total"]
fails = []


def check(cond, what):
    print(("PASS " if cond else "FAIL ") + what)
    if not cond:
        fails.append(what)


def browser(*args):
    r = subprocess.run(["node", str(ROOT / "tests/browser_upload.js"), *args], capture_output=True, text=True)
    if r.returncode:
        sys.exit(r.stderr)
    return json.loads(r.stdout)


def pct(w):
    return {k: round(v * 100) for k, v in w.items()}


def rounded(d):
    return {k: round(v, 1) for k, v in d.items()}


def expected_cols(w):
    on = {c for k, cs in KPI_COLS.items() if w[k] > 0 for c in cs}
    return [c for c in ORDER if c in BASE or c == "lead_time_days_total" or c in on]


def read_data(path):
    ws = load_workbook(path)["Data"]
    head = [c.value for c in ws[1]]
    return head, [dict(zip(head, [c.value for c in r])) for r in ws.iter_rows(min_row=2) if r[0].value]


# ---------------- 0. the committed index.html is a fresh build of template.html + data
from build import build_html  # noqa: E402
check((ROOT / "index.html").read_text() == build_html()[0], "index.html matches a fresh `python build.py` (rebuild and commit it if not)")

# ---------------- 0b. the page is on the site theme (same rule as tools/check_theme.py in the site repo)
import re  # noqa: E402
page = (ROOT / "template.html").read_text()
check(bool(re.search(r'<link rel="stylesheet" href="/theme\.css">', page)), "template.html links /theme.css")
check('<script src="/theme.js" defer></script>' in page, "template.html loads /theme.js")
check(page.index('href="/theme.css"') < page.index("<style>"), "/theme.css comes before the page's own styles")
check('<header class="dw-site">' in page and 'class="dw-skip"' in page, "template.html has the shared header and skip link")
for old in ["#161a1f", "#1b2026", "#303741", "#262c34", "#ebe8e1", "#aeb3b9", "#8f969f", "#9cc0dc", "156,192,220", "Source Serif"]:
    check(old.lower().replace(" ", "") not in page.lower().replace(" ", ""), f"template.html no longer uses {old}")

# ---------------- 1. downloads
d = browser("download", str(ROOT), str(TMP))
check(not d["errors"], f"no page errors {d['errors']}")
check(d["detailsOpen"], "'Score your own suppliers' button opens the upload panel")
check(d["names"] == ["supplier-scorecard-template.xlsx", "supplier-scorecard-template.csv"], f"download file names {d['names']}")
check(d["defaultWeights"] == pct(DEFAULT_WEIGHTS) and "research-based defaults" in d["defaultStatus"], "page starts on the research-based default weights")
check(d["afterUncheck"]["fill"] == 0 and sum(d["afterUncheck"].values()) == 100, f"unticking fill rate rebalances the rest to 100 {d['afterUncheck']}")
check("must add up to 100" in d["invalidStatus"] and d["invalidFixVisible"], f"a total other than 100 is flagged ({d['invalidStatus']})")
check(d["invalidScoresUnchanged"], "an invalid total doesn't change the scores")
check(sum(d["afterFix"].values()) == 100, f"'Make total 100%' fixes the total {d['afterFix']}")
check(d["customWeights"] == pct(CUSTOM) and "your weights" in d["customStatus"], f"custom weights applied {d['customWeights']}")
check("Fill rate" not in d["customHeaders"], "turned-off KPI disappears from the scorecard table")
check("Fill rate" not in d["customDlNote"].split("(off")[0] and "off: Fill rate" in d["customDlNote"], f"download note lists the weights ({d['customDlNote']})")
check(d["afterReload"] == pct(CUSTOM), "weights are remembered after a reload")
check(d["afterReset"] == pct(DEFAULT_WEIGHTS), "'Reset to research defaults' restores the defaults")
monthly = list(csv.DictReader(open(ROOT / "data/supplier_monthly.csv")))
check(d["defaultScoresSample"] == rounded(scores_by_supplier(monthly, 6)), "built-in sample data, default weights: page equals score.py (last 6 months)")
check(d["customScoresSample"] == rounded(scores_by_supplier(monthly, 6, CUSTOM)), "built-in sample data, custom weights: page equals score.py (last 6 months)")

for name, w in (("default", DEFAULT_WEIGHTS), ("custom", CUSTOM)):
    x = TMP / f"{name}.xlsx"
    with zipfile.ZipFile(x) as z:
        check(z.testzip() is None, f"{name}.xlsx: zip is valid")
        bad = []
        for n in z.namelist():
            if n.endswith(".xml") or n.endswith(".rels"):
                try:
                    minidom.parseString(z.read(n))
                except Exception as e:  # noqa: BLE001
                    bad.append(f"{n}: {e}")
        check(not bad, f"{name}.xlsx: every XML part is well-formed {bad}")
    wb = load_workbook(x)
    check(wb.sheetnames == ["Data", "Instructions", "KPI definitions", "Sources", "Settings"], f"{name}.xlsx: tabs {wb.sheetnames}")
    head, rows = read_data(x)
    check(head == expected_cols(w), f"{name}.xlsx: Data columns match the enabled KPIs {head}")
    check(len(rows) == 9 and all(r["supplier"].startswith("SAMPLE - ") for r in rows), f"{name}.xlsx: 9 labeled SAMPLE rows")
    with open(TMP / f"{name}.csv", newline="") as f:
        crow = list(csv.reader(f))
    check(crow[0] == head and len(crow) == 10, f"{name}.csv: same columns and rows as the .xlsx")
    check(all(str(a) == b or float(a) == float(b) for r, c in zip(rows, crow[1:]) for a, b in zip([r[k] for k in head], c)), f"{name}.csv: same values as the .xlsx")
    st = {r[0]: r[2] for r in wb["Settings"].iter_rows(min_row=4, values_only=True) if r[0] in KPIS}
    check(st == pct(w), f"{name}.xlsx: Settings tab holds the weights {st}")
    kp = wb["KPI definitions"]
    kw = {r[0].replace(" (off)", ""): r[2] for r in kp.iter_rows(values_only=True) if r[0] and isinstance(r[2], (int, float)) and r[1]}
    labels = {"otd": "On-time delivery", "fill": "Fill rate", "quality": "Defect rate", "invoice": "Invoice accuracy", "price": "Price vs. agreed"}
    check(all(abs(kw[labels[k]] - w[k]) < 1e-9 for k in KPIS), f"{name}.xlsx: KPI definitions tab shows the weights")
    ex = {r[0]: (r[1], r[2]) for r in kp.iter_rows(values_only=True) if r[0] and str(r[0]).startswith("SAMPLE - ") and isinstance(r[1], (int, float))}
    exp = {s: (round(v, 1), tier(v)) for s, v in scores_by_supplier(rows, 3, w).items()}
    check(ex == exp, f"{name}.xlsx: worked example equals score.py {exp}")

# ---------------- 2. upload fixtures
rest = list(csv.DictReader(open(ROOT / "data/mock_restaurant_suppliers.csv")))
for r in rest:  # make agreed_value optional in practice: one supplier never has it, another only some months
    if r["supplier"] == rest[0]["supplier"]:
        r["agreed_value"] = ""
suppliers = sorted({r["supplier"] for r in rest})
second = [s for s in suppliers if s != rest[0]["supplier"]][0]
for i, r in enumerate([r for r in rest if r["supplier"] == second]):
    if i % 2:
        r["agreed_value"] = ""

# C: filled-in default template, the way a user would: delete SAMPLE rows, type own rows; some months became Excel dates
wb = load_workbook(TMP / "default.xlsx")
ws = wb["Data"]
head = [c.value for c in ws[1]]
ws.delete_rows(2, 9)
for i, r in enumerate(rest, 2):
    for j, k in enumerate(head, 1):
        v = r[k]
        if k not in ("supplier", "category", "month"):
            v = None if v == "" else (float(v) if "." in v else int(v))
        ws.cell(i, j, v)
    if i % 5 == 0:
        y, m = map(int, r["month"].split("-"))
        c = ws.cell(i, 3, dt.datetime(y, m, 1))
        c.number_format = "mmm-yy"
filled_xlsx = TMP / "C_filled_template.xlsx"
wb.save(filled_xlsx)

# D: same data as Excel saves CSV, using the OLD header name contract_value, Title Case headers, BOM, CRLF, quoted thousands
filled_csv = TMP / "D_old_header_excel.csv"
old = [("contract_value" if k == "agreed_value" else k) for k in ORDER]
with open(filled_csv, "w", newline="", encoding="utf-8-sig") as f:
    w = csv.writer(f, lineterminator="\r\n")
    w.writerow([k.replace("_", " ").title() for k in old])
    for r in rest:
        w.writerow([(f"{float(r[k]):,.2f}" if r[k] != "" else "") if k in ("spend", "agreed_value") else r[k] for k in ORDER])

# E: broken file built from the default template
wb = load_workbook(TMP / "default.xlsx")
ws = wb["Data"]
L = {c.value: c.column_letter for c in ws[1]}
ws[f"{L['spend']}3"] = "about 13k"
ws[f"{L['on_time_pos']}4"] = 30           # pos on row 4 is 22
ws[f"{L['month']}6"] = "Q2 2026"
ws[f"{L['agreed_value']}7"] = "n/a"
for j in range(1, len(L) + 1):            # row 10 duplicates row 9
    ws.cell(10, j, ws.cell(9, j).value)
bad_xlsx = TMP / "E_errors.xlsx"
wb.save(bad_xlsx)

# F: wrong headers
bad_csv = TMP / "F_wrong_headers.csv"
bad_csv.write_text("vendor,month,spend\nAcme,2026-06,100\n")

# H: custom template with its Settings weights edited so they don't total 100
wb = load_workbook(TMP / "custom.xlsx")
wb["Settings"]["C4"] = 55
bad_settings = TMP / "H_bad_settings.xlsx"
wb.save(bad_settings)

files = [TMP / "default.xlsx", TMP / "default.csv", TMP / "custom.xlsx", TMP / "custom.csv", filled_xlsx, filled_csv, bad_xlsx, bad_csv, bad_settings]
u = browser("upload", str(ROOT), *map(str, files))
check(not u["errors"], f"no page errors during uploads {u['errors']}")
res = {r["file"]: r for r in u["results"]}
_, sample_rows = read_data(TMP / "default.xlsx")

for name in ("default.xlsx", "default.csv"):
    r = res[name]
    check("ok" in r["msg"]["cls"] and "9 rows, 3 suppliers, 3 months" in r["msg"]["text"], f"{name}: loads ({r['msg']['text'].splitlines()[0]})")
    check("still the template's SAMPLE data" in r["msg"]["text"], f"{name}: warns the SAMPLE rows are still there")
    check(r["m3"] == rounded(scores_by_supplier(sample_rows, 3)), f"{name}: page scores equal score.py {r['m3']}")
    check(r["badge"] == "Your data", f"{name}: badge switches to 'Your data'")

r = res["custom.xlsx"]
check("ok" in r["msg"]["cls"] and "Applied the weights saved in this template" in r["msg"]["text"], "custom.xlsx: page reads the weights back from the Settings tab")
check(r["weights"] == pct(CUSTOM), f"custom.xlsx: page weights now {r['weights']}")
check(r["m3"] == rounded(scores_by_supplier(sample_rows, 3, CUSTOM)), f"custom.xlsx: page scores equal score.py with the custom weights {r['m3']}")
check("Fill rate" not in r["headers"], "custom.xlsx: fill rate column hidden")

r = res["custom.csv"]
check("err" in r["msg"]["cls"] and "Fill rate is turned on but the file has no lines_ordered or lines_filled column" in r["msg"]["text"],
      f"custom.csv under default weights: explains the missing fill-rate columns ({r['msg']['text'][:90]}...)")

for name in ("C_filled_template.xlsx", "D_old_header_excel.csv"):
    r = res[name]
    check("ok" in r["msg"]["cls"] and f"{len(rest)} rows" in r["msg"]["text"] and "SAMPLE" not in r["msg"]["text"], f"{name}: loads")
    check("1 supplier has no agreed_value" in r["msg"]["text"], f"{name}: says which suppliers have no agreed price")
    for m in (3, 6):
        check(r[f"m{m}"] == rounded(scores_by_supplier(rest, m)), f"{name}: last {m} months scores equal score.py (blank agreed_value handled)")

e = res["E_errors.xlsx"]["msg"]
check("err" in e["cls"], "E_errors.xlsx: rejected")
for frag in ('Row 3 (SAMPLE - Northwind Fasteners): spend "about 13k" isn\'t a number',
             "Row 4 (SAMPLE - Northwind Fasteners): on_time_pos (30) is more than pos (22)",
             'Row 6 (SAMPLE - Bayview Packaging): month "Q2 2026" isn\'t a month',
             'Row 7 (SAMPLE - Bayview Packaging): agreed_value "n/a" isn\'t a number',
             "Row 10 (SAMPLE - Ridgeway Industrial): duplicate of row 9"):
    check(frag in e["text"], f"E_errors.xlsx: lists '{frag}'")
f = res["F_wrong_headers.csv"]["msg"]
check("err" in f["cls"] and "Missing columns: supplier, category, pos" in f["text"], f"F_wrong_headers.csv: rejected ({f['text'][:70]}...)")
h = res["H_bad_settings.xlsx"]
check("don't add up to 100%" in h["msg"]["text"] and h["weights"] == pct(DEFAULT_WEIGHTS), "H_bad_settings.xlsx: bad Settings weights are ignored, with a note")

print(f"\n{'ALL PASS' if not fails else str(len(fails)) + ' FAILED'}")
sys.exit(1 if fails else 0)
