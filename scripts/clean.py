"""Reproducible cleaning: raw CSV -> in-scope (2017-2018) frame + cleaning log. The raw CSV is never modified."""
import argparse
import json
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).parent))
from config import COLS, ORDER_CONSTANT, YEARS  # noqa: E402


def load(path, encoding="latin-1"):
    return pd.read_csv(path, encoding=encoding, dtype=str, keep_default_na=False)


def clean(raw):
    """Returns (in_scope, log). Values stay as raw text except typed fields; nothing is imputed."""
    missing = [v for v in COLS.values() if v not in raw.columns]
    if missing:
        raise KeyError(f"Columns not in file (update scripts/config.py): {missing}")
    log = {"raw_rows": len(raw)}
    df = raw[list(COLS.values())].copy()
    df.columns = list(COLS.keys())

    dates = pd.to_datetime(df["order_date"], errors="coerce")
    log["unparseable_order_date"] = int(dates.isna().sum())
    df["order_date"] = dates
    df = df[df["order_date"].dt.year.isin(YEARS)]
    log["rows_outside_years_or_bad_date"] = log["raw_rows"] - len(df)
    log["in_scope_rows"] = len(df)

    for c in ("ship_real", "ship_sched", "risk", "qty"):
        df[c] = pd.to_numeric(df[c], errors="coerce")
        log[f"non_numeric_{c}"] = int(df[c].isna().sum())

    log["duplicate_item_ids"] = int(df["item_id"].duplicated().sum())
    log["order_consistency"] = order_consistency(df)
    return df, log  # index kept so Raw Data rows can be selected from the untouched raw frame


def order_consistency(df):
    """Count orders with >1 distinct value per field that is assumed order-level."""
    g = df.groupby("order_id")
    out = {"orders": int(g.ngroups), "items_per_order_max": int(g.size().max()) if len(df) else 0}
    for c in ORDER_CONSTANT:
        out[f"orders_with_varying_{c}"] = int((g[c].nunique(dropna=False) > 1).sum())
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("raw_csv")
    ap.add_argument("--encoding", default="latin-1")
    ap.add_argument("--out", default="data/clean")
    a = ap.parse_args()
    df, log = clean(load(a.raw_csv, a.encoding))
    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    df.to_csv(out / "in_scope.csv", index=False)
    (out / "cleaning_log.json").write_text(json.dumps(log, indent=2))
    print(json.dumps(log, indent=2))


if __name__ == "__main__":
    main()
