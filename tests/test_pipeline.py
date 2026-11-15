"""Pipeline test on SYNTHETIC data: real LibreOffice recalculation vs independent pandas calculation."""
import shutil
import subprocess
from pathlib import Path

import pandas as pd
import pytest
from openpyxl import load_workbook

from scripts.build_workbook import build
from scripts.clean import clean
from scripts.config import COLS, EXCLUDED_STATUSES
from tests.synthetic import make_raw

pytestmark = pytest.mark.skipif(shutil.which("soffice") is None, reason="LibreOffice not installed")


@pytest.fixture(scope="module")
def recalculated(tmp_path_factory):
    d = tmp_path_factory.mktemp("wb")
    raw = make_raw()
    df, log = clean(raw)
    src = d / "in.xlsx"
    build(raw.loc[df.index, list(COLS.values())], df, src)
    out = d / "out"
    subprocess.run(["soffice", "--headless", "--convert-to", "xlsx", "--outdir", str(out), str(src)],
                   check=True, capture_output=True, timeout=300)
    return df, log, load_workbook(out / "in.xlsx", data_only=True), load_workbook(src)


def expected(df):
    """Independent pandas implementation of the KPI rules (does not reuse build_workbook)."""
    lat = df["ship_real"] - df["ship_sched"]
    elig = lat.notna() & df["status"].ne("") & ~df["status"].isin(EXCLUDED_STATUSES)
    late = lat > 0
    return lat, elig, late


def test_scope_and_log(recalculated):
    df, log, _, _ = recalculated
    assert set(df["order_date"].dt.year) <= {2017, 2018}
    assert log["raw_rows"] == log["in_scope_rows"] + log["rows_outside_years_or_bad_date"]


def test_sheets_present(recalculated):
    _, _, vals, _ = recalculated
    assert vals.sheetnames[0] == "Dashboard"
    assert {"Raw Data", "Cleaned Data", "Calculations", "KPI Definitions"} <= set(vals.sheetnames)


def test_no_error_cells(recalculated):
    _, _, vals, _ = recalculated
    errs = [(ws.title, c.coordinate, c.value) for ws in vals for row in ws.iter_rows() for c in row
            if isinstance(c.value, str) and c.value.startswith("#")]
    assert errs == []


def test_calculations_match_pandas(recalculated):
    df, _, vals, _ = recalculated
    lat, elig, late = expected(df)
    ws = vals["Calculations"]
    checked = 0
    for r in range(1, ws.max_row + 1):
        g = ws.cell(r, 1).value
        if g is None or ws.cell(r, 2).value is None or g in ("Group",):
            continue
        if g == "All":
            m = pd.Series(True, index=df.index)
        else:
            m = (df["mode"] == g) | (df["region"] == g) | (df["order_date"].dt.strftime("%Y-%m") == g)
        e = m & elig
        assert ws.cell(r, 2).value == m.sum()
        assert ws.cell(r, 3).value == e.sum()
        assert ws.cell(r, 4).value == (e & late).sum()
        if e.sum():
            assert ws.cell(r, 5).value == pytest.approx((e & late).sum() / e.sum())
            assert ws.cell(r, 6).value == pytest.approx(lat[e].mean())
        has = m & df["risk"].notna()
        assert ws.cell(r, 7).value == has.sum()
        if has.sum():
            assert ws.cell(r, 8).value == pytest.approx(df.loc[has, "risk"].mean())
        dis = e & df["risk"].notna() & (late.astype(int) != df["risk"])
        assert ws.cell(r, 9).value == dis.sum()
        checked += 1
    assert checked >= 10


def test_dashboard_default_matches_all(recalculated):
    df, _, vals, _ = recalculated
    _, elig, late = expected(df)
    ws = vals["Dashboard"]
    assert ws["B5"].value == elig.sum()
    assert ws["B6"].value == pytest.approx((elig & late).sum() / elig.sum())


def test_dashboard_has_dropdowns_and_charts(recalculated):
    _, _, _, formulas = recalculated
    ws = formulas["Dashboard"]
    assert len(ws.data_validations.dataValidation) == 2
    assert len(ws._charts) == 2


def test_dashboard_selectors_filter(recalculated, tmp_path):
    df, _, _, formulas = recalculated
    formulas["Dashboard"]["B2"], formulas["Dashboard"]["B3"] = "First Class", "West"
    src = tmp_path / "sel.xlsx"
    formulas.save(src)
    subprocess.run(["soffice", "--headless", "--convert-to", "xlsx", "--outdir", str(tmp_path / "o"), str(src)],
                   check=True, capture_output=True, timeout=300)
    ws = load_workbook(tmp_path / "o" / "sel.xlsx", data_only=True)["Dashboard"]
    lat, elig, late = expected(df)
    e = elig & (df["mode"] == "First Class") & (df["region"] == "West")
    assert ws["B5"].value == e.sum() > 0
    assert ws["B6"].value == pytest.approx((e & late).sum() / e.sum())
