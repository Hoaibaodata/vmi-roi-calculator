"""
Cost model for VMI ROI analysis.

Industry-typical parameters (override per-engagement):
- Holding cost rate: 20-30% of unit value per year
- Stockout cost: lost margin + expedite freight + customer penalty
- Order placement cost: $50-200 per PO
- VMI consignment fee: 3-8% of unit value held
"""
from __future__ import annotations

import math
from dataclasses import dataclass

from scipy.stats import norm

# Default cost parameters (industry-typical, override per-engagement)
DEFAULTS = {
    "holding_rate_annual": 0.25,           # 25% of unit price per year
    "stockout_cost_multiplier": 3.0,        # stockout cost = 3x unit margin
    "margin_rate": 0.30,                    # 30% gross margin
    "order_placement_cost": 100.0,          # $100 per PO administrative
    "vmi_fee_rate_annual": 0.06,            # 6% of value held by vendor per year
    "vmi_setup_cost": 5000.0,               # $5k one-time to negotiate + onboard
    "vmi_lt_days": 7,                       # VMI reduces LT to 7 days
    "service_level": 0.95,                  # 95% — Z = 1.65
    "review_period_weeks": 1,               # weekly review cycle
}


@dataclass
class CostBreakdown:
    safety_stock_units: float
    cycle_stock_units: float
    avg_inventory_units: float
    annual_holding_cost: float
    annual_stockout_cost: float
    annual_ordering_cost: float
    annual_vmi_fee: float
    annual_total_cost: float


def safety_stock(
    service_level: float,
    sigma_weekly_demand: float,
    lt_weeks: float,
    sigma_lt_weeks: float,
    avg_weekly_demand: float,
) -> float:
    """Classical Z full variability formula (industry tier-1 standard)."""
    z = norm.ppf(service_level)
    variance = lt_weeks * (sigma_weekly_demand ** 2) + (avg_weekly_demand ** 2) * (sigma_lt_weeks ** 2)
    return z * math.sqrt(max(variance, 0.0))


def compute_status_quo_cost(
    avg_weekly_demand: float,
    sigma_weekly_demand: float,
    unit_price: float,
    lt_days: float,
    sigma_lt_days: float,
    params: dict | None = None,
) -> CostBreakdown:
    """Cost under standard PO (longer LT, higher SS).

    Cycle stock assumed = LT_weeks * avg_weekly_demand (Q/2 lot = LT/2 weeks of supply average).
    Average inventory = cycle_stock + safety_stock.
    """
    p = {**DEFAULTS, **(params or {})}
    lt_weeks = lt_days / 7.0
    sigma_lt_weeks = sigma_lt_days / 7.0

    ss = safety_stock(p["service_level"], sigma_weekly_demand, lt_weeks, sigma_lt_weeks, avg_weekly_demand)
    cycle_stock = lt_weeks * avg_weekly_demand / 2.0  # average over reorder cycle
    avg_inventory = ss + cycle_stock

    annual_holding = avg_inventory * unit_price * p["holding_rate_annual"]

    # Stockout cost: probability × expected lost units × cost per unit
    # P(stockout) approx = 1 - service_level
    annual_demand = avg_weekly_demand * 52
    expected_stockout_units_per_year = annual_demand * (1 - p["service_level"]) * 0.1  # 10% of risk materializes
    stockout_cost_per_unit = unit_price * p["margin_rate"] * p["stockout_cost_multiplier"]
    annual_stockout = expected_stockout_units_per_year * stockout_cost_per_unit

    # Ordering: assume order once per LT cycle
    orders_per_year = 52 / max(lt_weeks, 1)
    annual_ordering = orders_per_year * p["order_placement_cost"]

    total = annual_holding + annual_stockout + annual_ordering

    return CostBreakdown(
        safety_stock_units=ss,
        cycle_stock_units=cycle_stock,
        avg_inventory_units=avg_inventory,
        annual_holding_cost=annual_holding,
        annual_stockout_cost=annual_stockout,
        annual_ordering_cost=annual_ordering,
        annual_vmi_fee=0.0,
        annual_total_cost=total,
    )


def compute_vmi_cost(
    avg_weekly_demand: float,
    sigma_weekly_demand: float,
    unit_price: float,
    lt_days_original: float,
    sigma_lt_days: float,
    params: dict | None = None,
) -> CostBreakdown:
    """Cost under VMI (shorter LT, lower SS, but pay VMI fee)."""
    p = {**DEFAULTS, **(params or {})}
    lt_weeks_vmi = p["vmi_lt_days"] / 7.0
    sigma_lt_weeks_vmi = sigma_lt_days / 7.0 * 0.3  # reduced slip variability under VMI

    ss = safety_stock(p["service_level"], sigma_weekly_demand, lt_weeks_vmi, sigma_lt_weeks_vmi, avg_weekly_demand)
    cycle_stock = lt_weeks_vmi * avg_weekly_demand / 2.0
    avg_inventory_buyer = ss + cycle_stock

    annual_holding = avg_inventory_buyer * unit_price * p["holding_rate_annual"]

    # Stockout much reduced under VMI (fast replenishment)
    annual_demand = avg_weekly_demand * 52
    expected_stockout = annual_demand * (1 - p["service_level"]) * 0.02  # 2% of original risk
    stockout_cost_per_unit = unit_price * p["margin_rate"] * p["stockout_cost_multiplier"]
    annual_stockout = expected_stockout * stockout_cost_per_unit

    # Ordering: fewer/shorter cycles under VMI (call-off based)
    annual_ordering = (52 / max(lt_weeks_vmi, 1)) * p["order_placement_cost"] * 0.3  # 30% admin reduction

    # VMI fee: vendor holds ~4 weeks worth of stock for buyer
    vmi_stock_at_vendor = 4 * avg_weekly_demand
    annual_vmi_fee = vmi_stock_at_vendor * unit_price * p["vmi_fee_rate_annual"]

    total = annual_holding + annual_stockout + annual_ordering + annual_vmi_fee

    return CostBreakdown(
        safety_stock_units=ss,
        cycle_stock_units=cycle_stock,
        avg_inventory_units=avg_inventory_buyer,
        annual_holding_cost=annual_holding,
        annual_stockout_cost=annual_stockout,
        annual_ordering_cost=annual_ordering,
        annual_vmi_fee=annual_vmi_fee,
        annual_total_cost=total,
    )
