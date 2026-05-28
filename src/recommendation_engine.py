"""VMI candidate ranking + recommendation engine."""
from __future__ import annotations

import pandas as pd

from .cost_model import DEFAULTS, compute_status_quo_cost, compute_vmi_cost


def compute_vmi_roi(sku_master: pd.DataFrame, params: dict | None = None) -> pd.DataFrame:
    """For each SKU, compare status quo vs VMI annual cost."""
    p = {**DEFAULTS, **(params or {})}
    rows = []

    for _, sku in sku_master.iterrows():
        sq = compute_status_quo_cost(
            avg_weekly_demand=sku["avg_weekly_demand"],
            sigma_weekly_demand=sku["sigma_weekly_demand"],
            unit_price=sku["unit_price"],
            lt_days=sku["lead_time_days"],
            sigma_lt_days=sku["sigma_lead_time_days"],
            params=p,
        )
        vmi = compute_vmi_cost(
            avg_weekly_demand=sku["avg_weekly_demand"],
            sigma_weekly_demand=sku["sigma_weekly_demand"],
            unit_price=sku["unit_price"],
            lt_days_original=sku["lead_time_days"],
            sigma_lt_days=sku["sigma_lead_time_days"],
            params=p,
        )

        savings = sq.annual_total_cost - vmi.annual_total_cost
        payback_months = (
            p["vmi_setup_cost"] / (savings / 12) if savings > 0 else float("inf")
        )

        rows.append({
            "sku_id": sku["sku_id"],
            "category": sku["category"],
            "abc_xyz": sku["abc_xyz"],
            "annual_volume": sku["annual_volume"],
            "annual_revenue": sku["annual_revenue"],
            "lead_time_days": sku["lead_time_days"],
            "status_quo_total_cost": round(sq.annual_total_cost, 2),
            "status_quo_holding": round(sq.annual_holding_cost, 2),
            "status_quo_stockout": round(sq.annual_stockout_cost, 2),
            "vmi_total_cost": round(vmi.annual_total_cost, 2),
            "vmi_holding": round(vmi.annual_holding_cost, 2),
            "vmi_fee": round(vmi.annual_vmi_fee, 2),
            "annual_savings": round(savings, 2),
            "payback_months": round(payback_months, 1) if payback_months != float("inf") else None,
            "is_critical": sku["is_critical"],
        })

    return pd.DataFrame(rows)


def recommend_vmi_candidates(roi_df: pd.DataFrame, top_n: int = 20) -> pd.DataFrame:
    """Rank SKUs by VMI candidacy.

    Criteria (in order):
    1. Positive annual_savings (must save money)
    2. Payback <= 24 months (sensible ROI window)
    3. Stable demand (X or Y class)
    4. Long base LT (more benefit from reduction)
    5. High volume / value
    """
    candidates = roi_df.copy()
    candidates["xyz"] = candidates["abc_xyz"].str[1]
    candidates["abc"] = candidates["abc_xyz"].str[0]

    # Filter: positive savings, reasonable payback, stable demand
    eligible = candidates[
        (candidates["annual_savings"] > 0)
        & (candidates["payback_months"].notna())
        & (candidates["payback_months"] <= 24)
        & (candidates["xyz"].isin(["X", "Y"]))
    ].copy()

    # Score = weighted savings + payback inverse + criticality
    eligible["candidacy_score"] = (
        eligible["annual_savings"] / 1000.0
        - eligible["payback_months"] * 10
        + eligible["is_critical"].astype(int) * 500
        + (eligible["abc"] == "A").astype(int) * 300
    )

    eligible = eligible.sort_values("candidacy_score", ascending=False).head(top_n)

    eligible["recommendation"] = "STRONG_VMI_CANDIDATE"
    eligible.loc[eligible["payback_months"] > 12, "recommendation"] = "MODERATE_VMI_CANDIDATE"
    eligible.loc[eligible["payback_months"] > 18, "recommendation"] = "WEAK_VMI_CANDIDATE"

    return eligible[
        [
            "sku_id",
            "category",
            "abc_xyz",
            "annual_volume",
            "annual_revenue",
            "lead_time_days",
            "status_quo_total_cost",
            "vmi_total_cost",
            "annual_savings",
            "payback_months",
            "candidacy_score",
            "recommendation",
            "is_critical",
        ]
    ].reset_index(drop=True)
