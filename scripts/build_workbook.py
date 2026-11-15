"""Build the workbook: Raw Data, Cleaned Data, Calculations, KPI Definitions, Dashboard.

Scope: item-level KPIs only. Order-level KPIs stay disabled until the within-order consistency
check in clean.py (cleaning_log.json -> order_consistency) shows no varying fields.
"""
import argparse
import sys
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

sys.path.insert(0, str(Path(__file__).parent))
from clean import clean, load  # noqa: E402
from config import COLS, EXCLUDED_STATUSES, YEARS  # noqa: E402

CD = "'Cleaned Data'"
KD = "'KPI Definitions'"
# Cleaned Data column layout (letters used in formulas)
CLEAN_COLS = ["order_id", "item_id", "order_date", "mode", "region", "status", "delivery_status",
              "risk", "ship_real", "ship_sched", "qty"]
DERIVED = ["Month", "Lateness_Days", "Eligible", "Calc_Late"]
LET = {k: L(i + 1) for i, k in enumerate(CLEAN_COLS + DERIVED)}
BOLD = Font(bold=True)

KPI_ROWS = [
    ("Calculated lateness days", "Lateness_Days = Days for shipping (real) - Days for shipment (scheduled)",
     "Order item", "Both fields numeric", "n/a", "Blank if either field is missing; never imputed"),
    ("Eligible item", "Eligible = 1 when Lateness_Days is numeric AND Order Status is non-blank and not in the exclusion list (column H)",
     "Order item", "See definition", "n/a", "Missing status => not eligible"),
    ("Calculated late-shipping rate", "COUNT(Calc_Late=1) / COUNT(Eligible=1); Calc_Late = 1 when Lateness_Days > 0",
     "Order item", "Eligible = 1", "Eligible items in group", "Rows outside eligibility are excluded and counted separately"),
    ("Mean lateness days", "AVERAGE(Lateness_Days) over eligible items (early shipments are negative)",
     "Order item", "Eligible = 1", "Eligible items in group", "Blank if no eligible items"),
    ("Recorded late-delivery-risk rate", "SUM(Late_delivery_risk) / COUNT(numeric Late_delivery_risk). Source-supplied flag, NOT calculated lateness",
     "Order item", "Late_delivery_risk numeric", "Rows with numeric flag", "Rows with non-numeric flag are excluded"),
    ("Disagreements", "Eligible items where Calc_Late and Late_delivery_risk differ (0/1 mismatch)",
     "Order item", "Eligible = 1 and flag numeric", "n/a (count)", "Reconciles the two measures; they are never merged"),
    ("Fill rate", "NOT IMPLEMENTED: no delivered/fulfilled quantity field is known to exist (only Order Item Quantity)",
     "n/a", "n/a", "n/a", "Omitted until a suitable field is confirmed in the real header"),
    ("Order-level KPIs", "NOT IMPLEMENTED: gated on within-order consistency check (cleaning_log.json)",
     "Order", "n/a", "n/a", "Item-level grain used throughout"),
]


def rng(key, n):
    c = LET[key]
    return f"{CD}!${c}$2:${c}${n + 1}"


def write_raw(wb, raw_subset):
    ws = wb.create_sheet("Raw Data")
    ws.append(list(raw_subset.columns))
    for row in raw_subset.itertuples(index=False):
        ws.append(list(row))
    ws.freeze_panes = "A2"


def write_cleaned(wb, df):
    ws = wb.create_sheet("Cleaned Data")
    ws.append([COLS[k] for k in CLEAN_COLS] + DERIVED)
    n_ex = len(EXCLUDED_STATUSES)
    for i, rec in enumerate(df[CLEAN_COLS].itertuples(index=False), start=2):
        vals = [None if pd.isna(v) else (v.to_pydatetime() if isinstance(v, pd.Timestamp) else v) for v in rec]
        c = LET
        vals += [
            f'=TEXT({c["order_date"]}{i},"yyyy-mm")',
            f'=IF(AND(ISNUMBER({c["ship_real"]}{i}),ISNUMBER({c["ship_sched"]}{i})),{c["ship_real"]}{i}-{c["ship_sched"]}{i},"")',
            f'=IF(AND(ISNUMBER({c["Lateness_Days"]}{i}),{c["status"]}{i}<>"",'
            f'COUNTIF({KD}!$H$2:$H${n_ex + 1},{c["status"]}{i})=0),1,0)',
            f'=IF({c["Eligible"]}{i}=1,IF({c["Lateness_Days"]}{i}>0,1,0),"")',
        ]
        ws.append(vals)
        ws.cell(row=i, column=3).number_format = "yyyy-mm-dd"
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{L(len(CLEAN_COLS) + len(DERIVED))}{len(df) + 1}"


def write_kpi_defs(wb):
    ws = wb.create_sheet("KPI Definitions")
    ws.append(["KPI", "Definition", "Grain", "Eligibility", "Denominator", "Missing-value handling"])
    for r in KPI_ROWS:
        ws.append(list(r))
    ws["H1"] = "Excluded Order Status values"
    for i, st in enumerate(EXCLUDED_STATUSES, start=2):
        ws.cell(row=i, column=8, value=st)
    return ws


def write_calcs(wb, df):
    ws = wb.create_sheet("Calculations")
    n = len(df)
    hdr = ["Group", "Items", "Eligible items", "Calc late items", "Calc late rate", "Mean lateness days",
           "Rows w/ numeric risk flag", "Recorded risk rate", "Calc vs recorded disagreements"]
    blocks, row = {}, 1
    for title, key in (("By shipping mode", "mode"), ("By order region", "region"), ("By order month", "Month")):
        ws.cell(row=row, column=1, value=title).font = BOLD
        row += 1
        for j, h in enumerate(hdr, 1):
            ws.cell(row=row, column=j, value=h).font = BOLD
        first = row + 1
        if key == "Month":
            groups = sorted(df["order_date"].dt.strftime("%Y-%m").unique())
        else:
            groups = sorted(x for x in df[key].dropna().unique() if x != "")
        for g in ["All"] + groups:
            row += 1
            crit = '"*"' if g == "All" else f"$A{row}"
            ws.cell(row=row, column=1, value=g)
            d, el, cl, rk, ri, sh = (rng(key, n), rng("Eligible", n), rng("Calc_Late", n),
                                     rng("risk", n), rng("risk", n), rng("Lateness_Days", n))
            f = {
                2: f"=COUNTIFS({d},{crit})",
                3: f"=COUNTIFS({d},{crit},{el},1)",
                4: f"=COUNTIFS({d},{crit},{cl},1)",
                5: f'=IF(C{row}=0,"",D{row}/C{row})',
                6: f'=IF(C{row}=0,"",AVERAGEIFS({sh},{d},{crit},{el},1))',
                7: f'=COUNTIFS({d},{crit},{ri},">=0")',
                8: f'=IF(G{row}=0,"",SUMIFS({rk},{d},{crit})/G{row})',
                9: f"=COUNTIFS({d},{crit},{el},1,{cl},1,{rk},0)+COUNTIFS({d},{crit},{el},1,{cl},0,{rk},1)",
            }
            for col, fx in f.items():
                ws.cell(row=row, column=col, value=fx)
            for col in (5, 8):
                ws.cell(row=row, column=col).number_format = "0.0%"
            ws.cell(row=row, column=6).number_format = "0.00"
        blocks[key] = (first, row)
        row += 3
    return ws, blocks


def write_dashboard(wb, df, calc, blocks):
    ws = wb.create_sheet("Dashboard", 0)
    n = len(df)
    ws["A1"] = f"Shipping performance, orders {YEARS[0]}-{YEARS[1]} (item-level grain)"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"], ws["A3"] = "Shipping mode", "Order region"
    ws["B2"], ws["B3"] = "All", "All"
    modes = ",".join(["All"] + sorted(x for x in df["mode"].dropna().unique() if x != ""))
    regions = ",".join(["All"] + sorted(x for x in df["region"].dropna().unique() if x != ""))
    for cell, lst in (("B2", modes), ("B3", regions)):
        dv = DataValidation(type="list", formula1=f'"{lst}"', allow_blank=False)
        ws.add_data_validation(dv)
        dv.add(cell)
    m, r = '$B$2', '$B$3'
    mc, rc = f'IF({m}="All","*",{m})', f'IF({r}="All","*",{r})'
    d, e, el, cl, sh, rk = (rng("mode", n), rng("region", n), rng("Eligible", n), rng("Calc_Late", n),
                            rng("Lateness_Days", n), rng("risk", n))
    base = f"{d},{mc},{e},{rc}"
    rows = [
        ("Eligible items", f"=COUNTIFS({base},{el},1)", "0"),
        ("Calculated late-shipping rate", f'=IF(B5=0,"",COUNTIFS({base},{cl},1)/B5)', "0.0%"),
        ("Mean lateness days (real - scheduled)", f'=IF(B5=0,"",AVERAGEIFS({sh},{base},{el},1))', "0.00"),
        ("Recorded Late_delivery_risk rate (source flag)",
         f'=IF(COUNTIFS({base},{rk},">=0")=0,"",SUMIFS({rk},{base})/COUNTIFS({base},{rk},">=0"))', "0.0%"),
        ("Eligible items where calc and recorded flag disagree",
         f"=COUNTIFS({base},{el},1,{cl},1,{rk},0)+COUNTIFS({base},{el},1,{cl},0,{rk},1)", "0"),
    ]
    ws["A4"] = "KPI (filtered by selectors above)"
    ws["A4"].font = BOLD
    for i, (lab, fx, fmt) in enumerate(rows, start=5):
        ws.cell(row=i, column=1, value=lab)
        ws.cell(row=i, column=2, value=fx).number_format = fmt
    ws["A11"] = ("Calculated lateness and the recorded Late_delivery_risk flag are different measures and are "
                 "shown separately. Charts below ignore the selectors (all data).")
    ws.column_dimensions["A"].width = 52
    ws.column_dimensions["B"].width = 18

    f0, l0 = blocks["mode"]
    bar = BarChart()
    bar.title = "Late rate by shipping mode: calculated vs recorded flag"
    for col in (5, 8):
        bar.add_data(Reference(calc, min_col=col, min_row=f0 - 1, max_row=l0), titles_from_data=True)
    bar.set_categories(Reference(calc, min_col=1, min_row=f0 + 1, max_row=l0))
    bar.series[0].val.numRef.f = f"'Calculations'!$E${f0 + 1}:$E${l0}"
    bar.series[1].val.numRef.f = f"'Calculations'!$H${f0 + 1}:$H${l0}"
    bar.y_axis.title, bar.x_axis.title = "Rate", "Shipping mode"
    bar.y_axis.delete = bar.x_axis.delete = False
    bar.width, bar.height = 18, 8
    ws.add_chart(bar, "D2")

    f1, l1 = blocks["Month"]
    line = LineChart()
    line.title = "Calculated late-shipping rate by order month"
    line.add_data(Reference(calc, min_col=5, min_row=f1 + 1, max_row=l1), titles_from_data=False)
    line.set_categories(Reference(calc, min_col=1, min_row=f1 + 1, max_row=l1))
    line.legend = None
    line.y_axis.title, line.x_axis.title = "Rate", "Month"
    line.y_axis.delete = line.x_axis.delete = False
    line.width, line.height = 18, 8
    ws.add_chart(line, "D19")


def build(raw_subset, df, out_path):
    wb = Workbook()
    wb.remove(wb.active)
    write_raw(wb, raw_subset)
    write_cleaned(wb, df)
    write_kpi_defs(wb)
    calc, blocks = write_calcs(wb, df)
    write_dashboard(wb, df, calc, blocks)
    wb.move_sheet("Dashboard", offset=0)
    wb.save(out_path)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("raw_csv")
    ap.add_argument("--encoding", default="latin-1")
    ap.add_argument("--out", default="build/SupplyChain_Analytics.xlsx")
    a = ap.parse_args()
    raw = load(a.raw_csv, a.encoding)
    df, _ = clean(raw)
    Path(a.out).parent.mkdir(parents=True, exist_ok=True)
    build(raw.loc[df.index, list(COLS.values())], df, a.out)
    print("wrote", a.out)


if __name__ == "__main__":
    main()
