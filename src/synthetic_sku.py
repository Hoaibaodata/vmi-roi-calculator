"""
Synthetic SKU master for VMI ROI analysis.

Generates 200 SKUs with diverse profiles (demand, price, lead time, criticality)
to demonstrate VMI candidate ranking. Real deployment loads from ERP master.
"""
from __future__ import annotations

import numpy as np
import pandas as pd

DEFAULT_SEED = 42
NUM_SKUS = 200


def generate_synthetic_sku_master(num_skus: int = NUM_SKUS, seed: int = DEFAULT_SEED) -> pd.DataFrame:
    """Generate diverse SKU portfolio with realistic VMI candidacy patterns.

    Distribution:
    - 10% high-volume strategic (likely VMI candidates)
    - 30% medium-volume standard
    - 60% low-volume long-tail
    """
    rng = np.random.default_rng(seed)
    rows = []
    categories = ["Electronics", "Mechanical", "Plastics", "Chemicals", "Packaging"]
    abc_classes = ["A", "B", "C"]
    abc_weights = [0.10, 0.30, 0.60]
    xyz_classes = ["X", "Y", "Z"]
    xyz_weights = [0.40, 0.40, 0.20]

    for i in range(1, num_skus + 1):
        abc = rng.choice(abc_classes, p=abc_weights)
        xyz = rng.choice(xyz_classes, p=xyz_weights)

        # Demand scale based on ABC class
        if abc == "A":
            avg_weekly_demand = float(rng.uniform(500, 5000))
            unit_price = float(rng.uniform(10, 100))
        elif abc == "B":
            avg_weekly_demand = float(rng.uniform(100, 500))
            unit_price = float(rng.uniform(5, 50))
        else:
            avg_weekly_demand = float(rng.uniform(5, 100))
            unit_price = float(rng.uniform(1, 20))

        # Variability based on XYZ class
        if xyz == "X":
            cv = float(rng.uniform(0.1, 0.4))  # stable
        elif xyz == "Y":
            cv = float(rng.uniform(0.4, 0.8))
        else:
            cv = float(rng.uniform(0.8, 2.0))  # erratic
        sigma_weekly_demand = avg_weekly_demand * cv

        # Lead time (longer = more VMI benefit)
        lt_days = int(rng.choice([7, 14, 21, 28, 35], p=[0.1, 0.3, 0.25, 0.25, 0.1]))
        # Sigma LT (slip variability)
        sigma_lt_days = float(lt_days * rng.uniform(0.1, 0.3))

        # Criticality: A-class with high consumption rate = more likely critical
        is_critical = abc == "A" and avg_weekly_demand > 1000

        rows.append({
            "sku_id": f"SKU{i:05d}",
            "category": rng.choice(categories),
            "abc_class": abc,
            "xyz_class": xyz,
            "abc_xyz": f"{abc}{xyz}",
            "avg_weekly_demand": avg_weekly_demand,
            "sigma_weekly_demand": sigma_weekly_demand,
            "unit_price": unit_price,
            "lead_time_days": lt_days,
            "sigma_lead_time_days": sigma_lt_days,
            "is_critical": is_critical,
            "annual_volume": avg_weekly_demand * 52,
            "annual_revenue": avg_weekly_demand * 52 * unit_price,
        })

    return pd.DataFrame(rows)
