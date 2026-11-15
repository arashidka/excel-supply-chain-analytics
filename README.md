# excel-supply-chain-analytics
Excel-based supply chain delivery and fulfillment performance analysis using the DataCo SMART Supply Chain dataset

> **Status: scaffolding only.** The source CSVs have not been obtained, so no real workbook, results or
> screenshots exist yet. The pipeline has been tested only on clearly labelled synthetic data (`tests/`).

## Data
See [DATA_PROVENANCE.md](DATA_PROVENANCE.md) (licence **unverified**). Download `DataCoSupplyChainDataset.csv` and
`DescriptionDataCoSupplyChain.csv` from https://data.mendeley.com/datasets/8gx2fvg2k6/5 into `data/raw/`
(excluded from Git). The original files are never modified.

## Rebuild
```
pip install -r requirements.txt
python scripts/checksums.py data/raw/*.csv        # record in DATA_PROVENANCE.md
python scripts/clean.py data/raw/DataCoSupplyChainDataset.csv
python scripts/build_workbook.py data/raw/DataCoSupplyChainDataset.csv
python -m pytest tests                            # synthetic data; needs LibreOffice (soffice)
```
Column names in `scripts/config.py` marked UNVERIFIED must be checked against the real header; `clean.py` fails if any is missing.

## Method (planned, item-level grain)
- Scope: orders dated 2017–2018 (date-field semantics to be verified on the real file).
- Calculated lateness = `Days for shipping (real)` − `Days for shipment (scheduled)`; eligibility excludes missing values
  and excluded statuses. The source `Late_delivery_risk` flag is reported separately and reconciled, never merged.
- Order-level KPIs are disabled until the within-order consistency check passes. Fill rate is not implemented
  (no delivered-quantity field known). Definitions: `KPI Definitions` sheet.
- Dashboard uses formula-driven tables, native charts and data-validation dropdowns (no PivotTables/slicers).
