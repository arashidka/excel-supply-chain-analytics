"""Column mapping and KPI rules.

Column names marked UNVERIFIED come from memory of the public dataset, not from the real header.
scripts/clean.py fails loudly if any are missing, so they are confirmed on first real run.
"""

COLS = {
    "order_id": "Order Id",                         # UNVERIFIED
    "item_id": "Order Item Id",                     # UNVERIFIED
    "order_date": "order date (DateOrders)",        # UNVERIFIED; semantics must be checked
    "ship_real": "Days for shipping (real)",        # seen in search snippet
    "ship_sched": "Days for shipment (scheduled)",  # seen in search snippet
    "status": "Order Status",                       # UNVERIFIED
    "delivery_status": "Delivery Status",           # seen in search snippet
    "risk": "Late_delivery_risk",                   # seen in search snippet
    "mode": "Shipping Mode",                        # UNVERIFIED
    "region": "Order Region",                       # UNVERIFIED
    "qty": "Order Item Quantity",                   # seen in search snippet
}

YEARS = (2017, 2018)

# Statuses excluded from the on-time denominator. Values UNVERIFIED; confirm against profile output.
EXCLUDED_STATUSES = ["CANCELED", "SUSPECTED_FRAUD"]

# Fields that must be constant within an order before any order-level KPI is computed.
ORDER_CONSTANT = ["order_date", "mode", "ship_real", "ship_sched", "risk", "delivery_status"]
