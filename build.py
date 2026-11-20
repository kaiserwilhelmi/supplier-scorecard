"""Embed data/supplier_monthly.csv into template.html -> index.html (one self-contained file)."""
import csv
import json
from pathlib import Path

here = Path(__file__).parent
rows = list(csv.DictReader(open(here / "data/supplier_monthly.csv")))
for r in rows:
    for k, v in r.items():
        if k not in ("supplier", "category", "month"):
            r[k] = float(v) if "." in v else int(v)
html = (here / "template.html").read_text().replace("/*__DATA__*/[]", json.dumps(rows, separators=(",", ":")))
(here / "index.html").write_text(html)
print(f"index.html written ({len(html) // 1024} KB, {len(rows)} rows)")
