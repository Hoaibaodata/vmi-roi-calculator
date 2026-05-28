# VMI / Building Stock Program ROI Calculator

Quantify the financial case for **Vendor Managed Inventory (VMI) / Consignment programs** SKU-by-SKU. Ranks candidates by annual savings, payback period, and demand stability.

> Companion to [inventory-optimization-portfolio](https://github.com/Hoaibaodata/inventory-optimization-portfolio).

## What VMI does

Under a Building Stock Program (VMI / Consignment):
1. Buyer pre-pays vendor to produce stock
2. Vendor holds inventory at their facility (not buyer's warehouse)
3. When buyer demands → vendor ships within **7 days** (vs typical 14-30 day LT)
4. Buyer pays consignment fee (~6% per year of value held by vendor)

Trade-off: pay holding cost to vendor, but **dramatically reduce buyer's safety stock + stockout risk** because LT compresses.

## Why this matters

VMI typically justifies for top 5-10% of SKU — but identifying them manually takes hours. This tool ranks all SKU by VMI candidacy automatically. CFO-level metric: working capital reduction.

## Output (verified end-to-end on 200 synthetic SKU)

```
Total SKUs analyzed:                  200
SKUs with positive VMI savings:       174 (87%)
Total potential annual savings:       $2,478,161
Top 20 candidates annual savings:     $939,557
```

Top 10 candidates (all A-class strategic SKU with long original LT):

```
sku_id   category    abc_xyz  annual_revenue  LT(d)  annual_savings  payback(mo)  recommendation
SKU109   Plastics    AX       $16.2M          28     $199,708         0.3         STRONG_VMI
SKU176   Plastics    AX       $11.2M          28     $136,929         0.4         STRONG_VMI
SKU195   Plastics    AX       $12.2M          28     $134,086         0.4         STRONG_VMI
SKU169   Packaging   AY       $9.7M           28     $115,171         0.5         STRONG_VMI
SKU113   Electronics AY       $6.9M           28     $77,614          0.8         STRONG_VMI
```

→ Payback under 1 year on top 5 SKU. Setup cost $5k → break-even in **first month** for top candidates.

## How the calculator works

### Cost model — Status quo

```
safety_stock     = Z × √(LT × σ_d² + d² × σ_LT²)            # classical full-variability
cycle_stock      = LT_weeks × avg_demand / 2                # average over reorder cycle
avg_inventory    = safety_stock + cycle_stock
holding_cost     = avg_inventory × unit_price × 0.25         # 25% holding rate
stockout_cost    = annual_demand × (1-SL) × 0.1 × margin × 3 # 10% of risk materializes
ordering_cost    = orders_per_year × $100
```

### Cost model — VMI

```
safety_stock_vmi    = Z × √(7d × σ_d² + d² × σ_LT_vmi²)     # LT reduced to 7 days
                                                              # sigma_LT reduced to 30% (vendor-managed = less slip)
holding_cost_vmi    = avg_inventory_vmi × unit_price × 0.25
stockout_cost_vmi   = annual_demand × (1-SL) × 0.02 × ...    # 2% of original risk
ordering_cost_vmi   = orders × $100 × 0.3                    # 70% admin reduction (call-off only)
vmi_fee             = 4_weeks_stock × unit_price × 0.06       # 6% annual fee on consigned stock
```

### ROI ranking

```
annual_savings    = status_quo_total - vmi_total
payback_months    = $5000 setup / (savings / 12)
candidacy_score   = savings/1000 - payback*10 + critical*500 + A_class*300
```

Filters out: stockouts negative savings, payback > 24 months, Z-class (erratic demand).

## Setup

```powershell
cd "D:\HOC_TAP\vmi-roi-calculator"
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python run_vmi_analysis.py
```

Runs in ~10 seconds. No external data needed.

## Parameters tunable in `src/cost_model.py:DEFAULTS`

| Parameter | Default | Industry range |
|---|---|---|
| `holding_rate_annual` | 25% | 15-30% |
| `stockout_cost_multiplier` | 3.0× | 2-10× |
| `margin_rate` | 30% | varies by industry |
| `order_placement_cost` | $100 | $50-200 |
| `vmi_fee_rate_annual` | 6% | 3-8% |
| `vmi_setup_cost` | $5,000 | $2-50k |
| `vmi_lt_days` | 7 | 5-14 |
| `service_level` | 95% | 90-99% |

Override by passing `params=` dict to `compute_vmi_roi()`.

## File structure

```
vmi-roi-calculator/
├── data/sku_master.csv         (generated synthetic, gitignored)
├── exports/
│   ├── vmi_roi_full.csv        (all 200 SKU with cost breakdown)
│   └── vmi_top_candidates.csv  (top 20 ranked)
├── src/
│   ├── synthetic_sku.py        # 200 SKU with diverse ABC-XYZ profiles
│   ├── cost_model.py           # status quo + VMI cost formulas
│   └── recommendation_engine.py # ranking + recommendation
├── docs/
│   ├── ROADMAP.md
│   └── IDEAS.md
├── requirements.txt
└── run_vmi_analysis.py
```

## Roadmap

See [docs/ROADMAP.md](docs/ROADMAP.md). Phase 1 (current) = deterministic cost model. Phase 2+: Monte Carlo simulation of demand uncertainty, parameter sensitivity tornado charts, interactive Streamlit calculator.

## License

MIT
