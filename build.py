"""Embed data/supplier_monthly.csv into template.html -> index.html (one self-contained file)."""
import csv
import json
from pathlib import Path

here = Path(__file__).parent


def build_html():
    """The page as build.py writes it. tests/roundtrip.py compares this with the committed index.html."""
    rows = list(csv.DictReader(open(here / "data/supplier_monthly.csv")))
    for r in rows:
        for k, v in r.items():
            if k not in ("supplier", "category", "month"):
                r[k] = float(v) if "." in v else int(v)
    return (here / "template.html").read_text().replace("/*__DATA__*/[]", json.dumps(rows, separators=(",", ":"))), len(rows)


if __name__ == "__main__":
    html, n = build_html()
    (here / "index.html").write_text(html)
    print(f"index.html written ({len(html) // 1024} KB, {n} rows)")
