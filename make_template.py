"""
Build the downloadable upload template:
  template/supplier-scorecard-template.xlsx  (Data + Instructions + KPI definitions + Sources)
  template/supplier-scorecard-template.csv   (same Data columns and sample rows)

Column names, weights, floors and targets are imported from score.py, so the template
can never drift from the scoring the dashboard uses.   Usage: python make_template.py
"""
import csv
from collections import defaultdict
from pathlib import Path

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from score import KPIS, TIERS, metrics, score, tier

HERE = Path(__file__).parent
OUT = HERE / "template"
DATA_ROWS = 1000  # rows pre-formatted for input

# ---------------------------------------------------------------- columns
# (column, type, unit, required rule, what to enter, example)
FIELDS = [
    ("supplier", "text", "-", "Required. Same spelling every month.",
     "Supplier name. One row per supplier per month.", "Acme Fasteners"),
    ("category", "text", "-", "Required.",
     "Spend category the supplier belongs to. Drives the category filter.", "Fasteners"),
    ("month", "text", "YYYY-MM", "Required. YYYY-MM (2026-06). Also accepted: 2026-06-01, 6/1/2026, Jun 2026.",
     "Calendar month the deliveries were due in.", "2026-06"),
    ("spend", "number", "USD", "Required, >= 0.",
     "Amount actually invoiced/paid for the month's POs.", "12,480.50"),
    ("contract_value", "number", "USD", "Required, > 0 in total per supplier.",
     "What the same quantities would have cost at contracted or quoted prices (qty x contract price).", "12,300.00"),
    ("pos", "whole number", "count", "Required, >= 0; total per supplier > 0.",
     "Purchase orders (deliveries) due in the month.", "24"),
    ("on_time_pos", "whole number", "count", "Required, 0 to pos.",
     "POs received by the promised (supplier-confirmed) date. Pick one late-day tolerance and keep it constant.", "22"),
    ("lines_ordered", "whole number", "count", "Required, >= 0; total per supplier > 0.",
     "PO lines due in the month.", "180"),
    ("lines_filled", "whole number", "count", "Required, 0 to lines_ordered.",
     "PO lines received in full quantity (no short-ship or backorder).", "176"),
    ("units_received", "whole number", "count", "Required, >= 0; total per supplier > 0.",
     "Units received (or inspected) in the month.", "5,400"),
    ("units_defective", "whole number", "count", "Required, 0 to units_received.",
     "Units rejected at receiving/inspection or returned for a quality reason. PPM = units_defective / units_received x 1,000,000.", "11"),
    ("invoices", "whole number", "count", "Required, >= 0; total per supplier > 0.",
     "Supplier invoices received for the month's POs.", "24"),
    ("invoices_accurate", "whole number", "count", "Required, 0 to invoices.",
     "Invoices that passed three-way match (PO, goods receipt, invoice) first time, with no correction or credit note.", "23"),
    ("lead_time_days_total", "number", "days", "Required, >= 0.",
     "Sum over the month's POs of (receipt date - PO date), in days. The tool shows the average (total / pos). Shown, not scored.", "168"),
]
COLUMNS = [f[0] for f in FIELDS]
COUNTS = [f[0] for f in FIELDS if f[1] == "whole number"]
LIMITS = {"on_time_pos": "pos", "lines_filled": "lines_ordered",
          "units_defective": "units_received", "invoices_accurate": "invoices"}

# ---------------------------------------------------------------- sample rows (clearly labeled)
# Three fictional suppliers x three months, chosen to land in different tiers.
SAMPLE = [
    # supplier, category, month, spend, contract, pos, on_time, lines_o, lines_f, units_r, units_d, inv, inv_ok, lead_total
    ("SAMPLE - Northwind Fasteners", "Fasteners", "2026-04", 12480.50, 12415.20, 24, 24, 180, 179, 5400, 3, 24, 24, 120),
    ("SAMPLE - Northwind Fasteners", "Fasteners", "2026-05", 13105.75, 13040.10, 26, 25, 196, 195, 5880, 4, 26, 26, 128),
    ("SAMPLE - Northwind Fasteners", "Fasteners", "2026-06", 11890.00, 11850.40, 22, 22, 165, 164, 4950, 2, 22, 22, 108),
    ("SAMPLE - Bayview Packaging", "Packaging", "2026-04", 8650.00, 8480.40, 18, 16, 90, 86, 2700, 30, 18, 17, 126),
    ("SAMPLE - Bayview Packaging", "Packaging", "2026-05", 9120.30, 8920.00, 19, 17, 95, 91, 2850, 31, 19, 18, 135),
    ("SAMPLE - Bayview Packaging", "Packaging", "2026-06", 8840.60, 8660.25, 17, 16, 85, 82, 2550, 25, 17, 16, 117),
    ("SAMPLE - Ridgeway Industrial", "MRO", "2026-04", 21300.00, 20280.00, 30, 25, 240, 214, 3600, 95, 30, 27, 300),
    ("SAMPLE - Ridgeway Industrial", "MRO", "2026-05", 22650.40, 21540.00, 32, 26, 256, 230, 3840, 104, 32, 29, 330),
    ("SAMPLE - Ridgeway Industrial", "MRO", "2026-06", 20980.10, 19975.50, 29, 25, 232, 210, 3480, 88, 29, 27, 285),
]

KPI_TEXT = {  # key: (formula from template columns, direction, why a buyer cares, sources)
    "otd": ("on_time_pos / pos x 100", "Higher is better",
            "Late deliveries drive expediting, safety stock and line stoppages.",
            "APICS (2017) SCOR v12, 'delivery performance to customer commit date' (component of Perfect Order Fulfillment, RL.1.1); "
            "Mena et al. (2021) Table 3.2, p. 49 ('on-time, in-full deliveries'); Gunasekaran et al. (2004)."),
    "fill": ("lines_filled / lines_ordered x 100", "Higher is better",
             "Short shipments and backorders create shortages even when the truck is on time.",
             "APICS (2017) SCOR v12, '% of orders delivered in full' (component of RL.1.1); Mena et al. (2021) Table 3.2, p. 49 (OTIF)."),
    "quality": ("units_defective / units_received x 100  (x 10,000 = PPM)", "Lower is better",
                "Rejects cost inspection, rework, returns and production risk.",
                "Mena et al. (2021) Table 3.2, p. 49 ('quality (PPM)'); APICS (2017) SCOR v12, 'perfect condition' (component of RL.1.1); "
                "Ho et al. (2010): quality is the most-used supplier evaluation criterion."),
    "invoice": ("invoices_accurate / invoices x 100", "Higher is better",
                "Invoice errors block three-way match, delay payment and add AP and buyer workload.",
                "APICS (2017) SCOR v12, 'documentation accuracy' (component of RL.1.1); Monczka et al. (2021) on the procure-to-pay cycle."),
    "price": ("(spend / contract_value - 1) x 100", "Lower is better (0% = paid contract price)",
              "Shows leakage between the negotiated price and what was actually paid.",
              "Purchase price variance, Datar & Rajan (2021); Weber et al. (1991): net price is the most frequently cited vendor criterion."),
}

SOURCES = [
    ("APICS (2017)", "APICS. (2017). Supply Chain Operations Reference Model (SCOR), Version 12.0. Chicago, IL: APICS (now ASCM). "
     "Perfect Order Fulfillment (RL.1.1) and its components: % of orders delivered in full, delivery performance to customer commit date, "
     "documentation accuracy, perfect condition.", "Standards body (ASCM)"),
    ("Datar & Rajan (2021)", "Datar, S. M., & Rajan, M. V. (2021). Horngren's Cost Accounting: A Managerial Emphasis (17th ed.). Pearson. "
     "Price variances.", "Textbook"),
    ("Dickson (1966)", "Dickson, G. W. (1966). An analysis of vendor selection systems and decisions. Journal of Purchasing, 2(1), 5-17.",
     "Peer-reviewed"),
    ("Gunasekaran et al. (2004)", "Gunasekaran, A., Patel, C., & McGaughey, R. E. (2004). A framework for supply chain performance measurement. "
     "International Journal of Production Economics, 87(3), 333-347. https://doi.org/10.1016/j.ijpe.2003.08.003", "Peer-reviewed"),
    ("Ho et al. (2010)", "Ho, W., Xu, X., & Dey, P. K. (2010). Multi-criteria decision making approaches for supplier evaluation and selection: "
     "A literature review. European Journal of Operational Research, 202(1), 16-24. https://doi.org/10.1016/j.ejor.2009.05.009", "Peer-reviewed"),
    ("Mena et al. (2021)", "Mena, C., van Hoek, R., & Christopher, M. (2021). Leading Procurement Strategy: Driving Value Through the Supply Chain "
     "(3rd ed.). Kogan Page. Table 3.2 (p. 49), balanced scorecard for procurement; pp. 70-71, rules for supplier dashboards.", "Textbook"),
    ("Monczka et al. (2021)", "Monczka, R. M., Handfield, R. B., Giunipero, L. C., & Patterson, J. L. (2021). Purchasing and Supply Chain "
     "Management (7th ed.). Cengage. Supplier performance measurement and the weighted-point evaluation method.", "Textbook"),
    ("Weber et al. (1991)", "Weber, C. A., Current, J. R., & Benton, W. C. (1991). Vendor selection criteria and methods. European Journal of "
     "Operational Research, 50(1), 2-18. https://doi.org/10.1016/0377-2217(91)90033-R", "Peer-reviewed"),
]

# ---------------------------------------------------------------- styles
NAVY, INK, GREY, LINE = "1F3A5F", "1A1A1A", "6B7280", "D0D5DD"
HEAD = Font(bold=True, color="FFFFFF")
HEAD_FILL = PatternFill("solid", fgColor=NAVY)
SAMPLE_FILL = PatternFill("solid", fgColor="FFF4D6")
TITLE = Font(bold=True, size=14, color=INK)
BOLD = Font(bold=True, color=INK)
NOTE = Font(italic=True, color=GREY)
WRAP = Alignment(wrap_text=True, vertical="top")
THIN = Border(bottom=Side(style="thin", color=LINE))


def header(ws, row, labels, widths):
    for j, (lab, w) in enumerate(zip(labels, widths), 1):
        c = ws.cell(row, j, lab)
        c.font, c.fill, c.alignment = HEAD, HEAD_FILL, Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[get_column_letter(j)].width = w


def sample_scores():
    by = defaultdict(list)
    for r in SAMPLE:
        by[r[0]].append(dict(zip(COLUMNS, r)))
    out = []
    for sup, rs in by.items():
        m = metrics(rs)
        out.append((sup, m, score(m), tier(score(m))))
    return sorted(out, key=lambda x: -x[2])


def build_xlsx(path):
    wb = Workbook()

    # ---- Data (first sheet: the dashboard reads the sheet named "Data", else the first sheet)
    ws = wb.active
    ws.title = "Data"
    widths = [32, 14, 10, 12, 15, 7, 12, 13, 12, 14, 15, 9, 16, 19]
    header(ws, 1, COLUMNS, widths)
    for j, f in enumerate(FIELDS, 1):
        ws.cell(1, j).comment = Comment(f"{f[4]}\nUnit: {f[2]}. {f[3]}", "Template")
    for i, r in enumerate(SAMPLE, 2):
        for j, v in enumerate(r, 1):
            c = ws.cell(i, j, v)
            c.fill = SAMPLE_FILL
    for i in range(2, DATA_ROWS + 2):
        ws.cell(i, 3).number_format = "@"  # month stays text: Excel won't turn 2026-06 into a date
        for j in (4, 5):
            ws.cell(i, j).number_format = "#,##0.00"
        for j in range(6, 14):
            ws.cell(i, j).number_format = "#,##0"
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = f"A1:N{len(SAMPLE) + 1}"

    last = DATA_ROWS + 1
    col = {k: get_column_letter(j) for j, k in enumerate(COLUMNS, 1)}
    dv_money = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True,
                              showErrorMessage=True, errorTitle="Not a valid amount", error="Enter a number of 0 or more (USD, no text).")
    dv_money.add(f"D2:E{last}")
    dv_money.add(f"N2:N{last}")
    ws.add_data_validation(dv_money)
    for k in ("pos", "lines_ordered", "units_received", "invoices"):
        dv = DataValidation(type="whole", operator="greaterThanOrEqual", formula1="0", allow_blank=True,
                            showErrorMessage=True, errorTitle="Whole number needed", error=f"{k} must be a whole number of 0 or more.")
        dv.add(f"{col[k]}2:{col[k]}{last}")
        ws.add_data_validation(dv)
    for k, lim in LIMITS.items():
        dv = DataValidation(type="custom", formula1=f"AND(INT({col[k]}2)={col[k]}2,{col[k]}2>=0,{col[k]}2<={col[lim]}2)",
                            allow_blank=True, showErrorMessage=True, errorTitle="Check this count",
                            error=f"{k} must be a whole number between 0 and {lim} on the same row.")
        dv.add(f"{col[k]}2:{col[k]}{last}")
        ws.add_data_validation(dv)
    dv_m = DataValidation(type="custom", formula1='AND(LEN(C2)=7,MID(C2,5,1)="-")', allow_blank=True,
                          showErrorMessage=True, showInputMessage=True, promptTitle="Month", prompt="Type as YYYY-MM, e.g. 2026-06",
                          errorStyle="warning", errorTitle="Month format", error="Use YYYY-MM, e.g. 2026-06.")
    dv_m.add(f"C2:C{last}")
    ws.add_data_validation(dv_m)

    # ---- Instructions
    ins = wb.create_sheet("Instructions")
    ins["A1"] = "Supplier Scorecard upload template"
    ins["A1"].font = TITLE
    steps = [
        "1. Go to the Data tab. The yellow rows starting with 'SAMPLE -' are fictional examples. Delete them before adding your own data.",
        "2. Add one row per supplier per month. Keep the header row (row 1) exactly as it is; extra columns to the right are ignored.",
        "3. Enter counts for the month, not percentages. The dashboard calculates the KPIs and weights them by volume.",
        "4. Save as .xlsx, or File > Save As > CSV (Data tab only). Then open dawsonwilhelmi.com/supplier-scorecard/, open 'Score your own suppliers', and drop the file in.",
        "5. If something is wrong, the page lists each problem with its row number. Nothing is uploaded: the file is read in your browser.",
        "Tip: load this template unchanged first. The page should show the three SAMPLE suppliers with the scores in the 'Worked example' on the KPI definitions tab.",
    ]
    for i, s in enumerate(steps, 3):
        ins.cell(i, 1, s).alignment = WRAP
        ins.merge_cells(start_row=i, start_column=1, end_row=i, end_column=6)
        ins.row_dimensions[i].height = 30
    r0 = len(steps) + 4
    ins.cell(r0, 1, "Columns").font = BOLD
    header(ins, r0 + 1, ["Column", "Type", "Unit", "What to enter", "Rule checked on upload", "Example"], [24, 13, 10, 62, 44, 18])
    for i, (name, typ, unit, rule, what, ex) in enumerate(FIELDS, r0 + 2):
        for j, v in enumerate((name, typ, unit, what, rule, ex), 1):
            c = ins.cell(i, j, v)
            c.alignment, c.border = WRAP, THIN
        ins.cell(i, 1).font = Font(bold=True, name="Consolas", color=INK)
        ins.row_dimensions[i].height = 32
    r1 = r0 + 2 + len(FIELDS) + 1
    notes = [
        "Data quality notes",
        "- Count each PO, line and invoice once, in the month its delivery was due, so on-time and fill rate use the same base.",
        "- 'On time' means received by the date the supplier confirmed. Decide the late-day tolerance once (e.g. 0 days) and keep it.",
        "- Use the same supplier name every month (one vendor ID per name). Duplicate supplier + month rows are rejected.",
        "- Money in one currency (USD assumed). Leave out tax and freight unless the contract price includes them.",
    ]
    for i, s in enumerate(notes, r1):
        c = ins.cell(i, 1, s)
        c.font = BOLD if i == r1 else Font(color=INK)
        ins.merge_cells(start_row=i, start_column=1, end_row=i, end_column=6)

    # ---- KPI definitions
    kp = wb.create_sheet("KPI definitions")
    kp["A1"] = "How the score is calculated"
    kp["A1"].font = TITLE
    kp["A2"] = ("Each KPI earns 0-100 points on a straight line from its floor (0 pts) to its target (100 pts), capped at both ends. "
                "The composite score is the weighted sum (weighted-point method; Monczka et al., 2021). KPIs are calculated from summed counts "
                "over the period, so larger months carry more weight.")
    kp["A2"].alignment = WRAP
    kp.merge_cells("A2:H2")
    kp.row_dimensions[2].height = 45
    header(kp, 4, ["KPI", "Formula (template columns)", "Weight", "Floor (0 pts)", "Target (100 pts)", "Direction", "Why it matters", "Definition source"],
           [30, 34, 9, 12, 14, 18, 40, 70])
    for i, (k, (label, w, f, t)) in enumerate(KPIS.items(), 5):
        formula, direction, why, src = KPI_TEXT[k]
        vals = (label, formula, w, f / 100, t / 100, direction, why, src)
        for j, v in enumerate(vals, 1):
            c = kp.cell(i, j, v)
            c.alignment, c.border = WRAP, THIN
        kp.cell(i, 3).number_format = "0%"
        kp.cell(i, 4).number_format = kp.cell(i, 5).number_format = "0.0%"
        kp.row_dimensions[i].height = 62
    r = 5 + len(KPIS)
    kp.cell(r, 1, "Total").font = BOLD
    kp.cell(r, 3, f"=SUM(C5:C{r - 1})").number_format = "0%"
    kp.cell(r + 1, 1, "Average lead time = lead_time_days_total / pos (days). Shown for context, not scored.").font = NOTE
    tiers = " | ".join(f"{name} {'>= ' + str(cut) if cut else '< ' + str(TIERS[-2][0])}" for cut, name in TIERS)
    kp.cell(r + 2, 1, f"Tiers: {tiers}").font = BOLD

    r = r + 4
    kp.cell(r, 1, "Why these KPIs and weights").font = BOLD
    why = [
        "- Quality, delivery and price are the criteria most used to evaluate suppliers in the research literature "
        "(Dickson, 1966; Weber et al., 1991; Ho et al., 2010).",
        "- On-time, in-full, damage-free and correctly documented are the four parts of the SCOR 'perfect order' (APICS, 2017); "
        "this scorecard measures each one separately so a buyer can see which part fails.",
        "- Delivery (on-time 30% + fill 20%) and quality (20%) carry 70% because they are what stops operations; price (15%) is tracked "
        "against contract because negotiated price is already set at award; invoice accuracy (15%) reflects procure-to-pay workload.",
        "- The weights, floors, targets and tier cut-offs are this tool's defaults for a distribution business, not published benchmarks. "
        "Set them to the service levels in your own supplier contracts, and keep the KPI set short and stable (Mena et al., 2021, pp. 70-71).",
    ]
    for i, s in enumerate(why, r + 1):
        c = kp.cell(i, 1, s)
        c.alignment = WRAP
        kp.merge_cells(start_row=i, start_column=1, end_row=i, end_column=8)
        kp.row_dimensions[i].height = 32

    r = r + len(why) + 2
    kp.cell(r, 1, "Worked example: the SAMPLE rows on the Data tab (Apr-Jun 2026). Upload the template unchanged and the page should show these numbers.").font = BOLD
    labels = ["Supplier", "Score", "Tier", "On-time %", "Fill rate %", "Defect rate %", "PPM", "Invoice accuracy %", "Price vs contract %", "Avg lead time (days)", "Spend"]
    for j, lab in enumerate(labels, 1):
        c = kp.cell(r + 1, j, lab)
        c.font, c.fill, c.alignment = HEAD, HEAD_FILL, Alignment(wrap_text=True)
    for i, (sup, m, sc, tr) in enumerate(sample_scores(), r + 2):
        vals = [sup, round(sc, 1), tr, round(m["otd"], 1), round(m["fill"], 1), round(m["quality"], 2), round(m["quality"] * 10000),
                round(m["invoice"], 1), round(m["price"], 2), round(m["lead"], 1), round(m["spend"], 2)]
        for j, v in enumerate(vals, 1):
            kp.cell(i, j, v).border = THIN
        kp.cell(i, 11).number_format = "$#,##0"
    for j in range(9, 12):
        kp.column_dimensions[get_column_letter(j)].width = 16

    # ---- Sources
    so = wb.create_sheet("Sources")
    so["A1"] = "Sources"
    so["A1"].font = TITLE
    header(so, 3, ["Short cite", "Full reference", "Type"], [24, 120, 20])
    for i, (short, full, typ) in enumerate(SOURCES, 4):
        for j, v in enumerate((short, full, typ), 1):
            c = so.cell(i, j, v)
            c.alignment, c.border = WRAP, THIN
        so.row_dimensions[i].height = 34
    so.cell(5 + len(SOURCES), 1, "Sample rows are fictional. Template by Dawson Wilhelmi, dawsonwilhelmi.com/supplier-scorecard/").font = NOTE

    wb.properties.title = "Supplier Scorecard upload template"
    wb.properties.creator = "Dawson Wilhelmi"
    wb.save(path)


def build_csv(path):
    with open(path, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(COLUMNS)
        w.writerows(SAMPLE)


if __name__ == "__main__":
    OUT.mkdir(exist_ok=True)
    build_xlsx(OUT / "supplier-scorecard-template.xlsx")
    build_csv(OUT / "supplier-scorecard-template.csv")
    for sup, m, sc, tr in sample_scores():
        print(f"{sup:32s} {sc:5.1f} {tr}")
    print("template/ written")
