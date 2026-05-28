# VMI ROI Calculator — Roadmap

## Phase 1: Cost model
- Holding cost rate (typical 15-25% of unit price per year)
- Stockout cost (lost revenue + expedite freight + customer penalty)
- Order placement cost (PO admin overhead)
- VMI consignment fee (% of unit price held by vendor)

## Phase 2: Status quo simulation
- For each SKU, simulate full year using current LT + safety stock policy
- Compute: holding cost, stockout events, expedite events, total cost

## Phase 3: VMI simulation
- Reduce LT to 7 days (or configurable)
- Recompute safety stock under reduced LT
- Add VMI consignment fee
- Total cost comparison

## Phase 4: Sensitivity + Monte Carlo
- Parameter sweeps (what if VMI fee = 5% vs 10% vs 15%?)
- Monte Carlo: random demand + slip scenarios

## Phase 5: Recommendation engine
- Rank SKU by potential savings
- Filter by: demand stability (XYZ class), volume (ABC class), unit price
- Output: top N candidates for VMI negotiation
