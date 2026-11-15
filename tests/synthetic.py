"""SYNTHETIC test data only. Not real DataCo data; never use for findings or screenshots in deliverables."""
import random

import pandas as pd

from scripts.config import COLS

MODES = ["Standard Class", "First Class", "Second Class", "Same Day"]
REGIONS = ["West", "East", "Central"]


def make_raw(n=400, seed=7):
    rnd = random.Random(seed)
    rows = []
    for i in range(n):
        oid = i // 2 + 1  # two items per order
        year = rnd.choice([2016, 2017, 2017, 2018, 2018])
        sched = rnd.choice([0, 1, 2, 4])
        real = sched + rnd.choice([-1, 0, 0, 1, 2])
        status = rnd.choice(["COMPLETE", "COMPLETE", "PENDING", "CANCELED", "SUSPECTED_FRAUD", ""])
        rows.append({
            COLS["order_id"]: str(oid), COLS["item_id"]: str(i + 1),
            COLS["order_date"]: f"{rnd.randint(1, 12)}/{rnd.randint(1, 28)}/{year} 10:00",
            COLS["ship_real"]: "" if rnd.random() < 0.05 else str(max(real, 0)),
            COLS["ship_sched"]: "" if rnd.random() < 0.03 else str(sched),
            COLS["status"]: status, COLS["delivery_status"]: "Late delivery",
            COLS["risk"]: rnd.choice(["0", "1", "1", ""]),
            COLS["mode"]: rnd.choice(MODES), COLS["region"]: rnd.choice(REGIONS),
            COLS["qty"]: str(rnd.randint(1, 5)),
        })
    return pd.DataFrame(rows)
