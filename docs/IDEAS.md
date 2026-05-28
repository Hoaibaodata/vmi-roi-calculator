# VMI ROI Calculator — Tech Ideas

## Holding cost components
1. **Capital cost** (cost of money tied up): WACC × unit_price
2. **Storage cost** (warehouse space): per cubic meter × volume per unit
3. **Insurance + tax**: typically 1-2% of inventory value
4. **Obsolescence risk**: probability × loss if obsolete

Total typically 20-30% of unit price per year for retail/manufacturing.

## Stockout cost
- **Lost sale margin**: probability_to_substitute × margin
- **Expedite freight**: difference between air vs sea
- **Customer penalty**: contract OTIF penalty clauses

## VMI fee structure (industry-typical)
- 3-8% of unit price held
- Some vendors: flat monthly fee
- Hybrid: base fee + per-unit pickup fee

## Decision criteria (from MRP-INSIGHTS field knowledge)
Friend at Milwaukee plant mentions: "VMI cho những con top demand". Translation: VMI only justifies for SKU with:
- High volume (high holding cost savings)
- Stable demand (X or Y class)
- Long base LT (>= 14 days, since VMI reduces to 7 days)
- Critical to production (stop-line risk)

## ROI calculation template

```
annual_status_quo_cost = 
    holding_cost(safety_stock_current * unit_price) +
    expected_stockout_cost +
    order_placement_cost * orders_per_year

annual_vmi_cost = 
    holding_cost(safety_stock_reduced * unit_price) +  // less SS because LT shorter
    vmi_fee +
    expected_stockout_cost_reduced  // fewer stockouts

annual_savings = annual_status_quo_cost - annual_vmi_cost
roi_payback_months = vmi_setup_cost / (annual_savings / 12)
```

## Visualization
- Waterfall chart: status quo cost → VMI savings → VMI new cost
- Scatter plot: candidates by (annual_volume, savings_potential)
- Sensitivity tornado: which parameter affects ROI most?
