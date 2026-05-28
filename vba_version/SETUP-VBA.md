# VMI ROI Calculator — VBA Version Setup

> Excel-only edition. UDFs callable as Excel formulas. Tunable parameters trong sheet `Parameters`.

## Bước 1 — Generate starter

```powershell
cd "D:\HOC_TAP\vmi-roi-calculator\vba_version"
python generate_starter_xlsx.py
```

→ Tạo `VMI_ROI.xlsx` với 6 sheets pre-populated.

## Bước 2 — Save As .xlsm

`VMI_ROI.xlsx` → Save As → **Excel Macro-Enabled Workbook (.xlsm)** → `VMI_ROI.xlsm`.

## Bước 3 — Import 5 VBA modules

`Alt + F11` → File → Import File → 5 .bas trong `vba_modules/`:
- `M0_Helpers.bas`
- `M1_CostFormulas.bas` (UDFs)
- `M2_ROIEngine.bas`
- `M3_TopCandidates.bas`
- `M4_SampleData.bas`
- `M5_Dashboard.bas`

## Bước 4 — Tạo 5 nút trên Dashboard

| Caption | Macro |
|---|---|
| 🚀 **Run Full Analysis** | `Action_RunFullAnalysis` |
| 📥 Generate Sample SKU | `Action_GenerateSample` |
| 💰 Compute ROI | `Action_ComputeROI` |
| 🏆 Top Candidates | `Action_TopCandidates` |
| 🔄 Reset Outputs | `Action_ResetOutputs` |

## Bước 5 — Chạy thử

Bấm **Run Full Analysis** → Verify:

- Sheet `SKU_Master`: 50 SKUs với ABC-XYZ profiles
- Sheet `ROI_Output`: 50 dòng với status quo cost, VMI cost, savings, recommendation
- Sheet `Top_Candidates`: top 20 sorted by candidacy score

## UDFs Available (gõ trong Excel cell)

```excel
=SafetyStock(service_level, sigma_demand, lt_weeks, avg_demand, sigma_lt_weeks)
=StatusQuoCost(avg_demand, sigma_demand, unit_price, lt_days, sigma_lt_days)
=VMICost(avg_demand, sigma_demand, unit_price, sigma_lt_days)
=VMISavings(avg_demand, sigma_demand, unit_price, lt_days, sigma_lt_days)
=VMIPaybackMonths(5000, annual_savings)
=VMIRecommendation(payback_months, "AX")
```

→ Bạn có thể paste data thực rồi gõ formulas thay vì bấm nút!

## Edit parameters

Sheet `Parameters` — edit cột B values:
- `holding_rate_annual`: 0.25 → đổi sang 0.20 nếu công ty bạn rate thấp hơn
- `vmi_fee_rate_annual`: 0.06 → tăng nếu vendor đòi cao
- `vmi_setup_cost`: $5000 → adjust theo negotiation
- `service_level`: 0.95 → cao hơn = SS to hơn = status quo cost cao = savings to hơn

Re-run "Compute ROI" sau khi edit.

## Patterns demonstrated

| Module | VBA pattern |
|---|---|
| M0 | `GetParameter()` UDF reading from sheet |
| M1 | UDFs callable from cells, `NormSInv()` worksheet function |
| M2 | Loop SKU master, compute, write, conditional formatting |
| M3 | Filter + sort + ranking, Bubble sort via Dictionary |
| M4 | Random seed via `Randomize n`, weighted random distribution |
| M5 | Action router, ScreenUpdating optimization |
