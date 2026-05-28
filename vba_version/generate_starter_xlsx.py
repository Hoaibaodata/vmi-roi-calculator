"""Generate VMI_ROI.xlsx starter — sheets + headers + Parameters + Dashboard."""
from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill


HEADER_FONT = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
TITLE_FONT = Font(name="Segoe UI", size=18, bold=True, color="1F4E79")


def write_headers(ws, headers):
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 22


def main() -> None:
    wb = Workbook()
    wb.remove(wb.active)

    # Dashboard
    ws_d = wb.create_sheet("Dashboard")
    ws_d["A1"] = "VMI ROI Calculator — Dashboard"
    ws_d["A1"].font = TITLE_FONT
    ws_d.merge_cells("A1:H1")
    workflow = [
        "",
        "Workflow:",
        "  BƯỚC 1: Generate Sample (50 SKU diverse ABC-XYZ profiles)",
        "  BƯỚC 2: Compute ROI (status quo vs VMI cost per SKU)",
        "  BƯỚC 3: Top Candidates (ranked by candidacy score)",
        "",
        "Hoặc bấm 'Run Full Analysis' để chạy hết một mạch.",
        "",
        "Tunable parameters trong sheet 'Parameters' (holding rate, MOV, service level...).",
    ]
    for i, line in enumerate(workflow, start=3):
        ws_d.cell(row=i, column=1, value=line).font = Font(name="Segoe UI", size=11, bold=(i in (4, 6, 9)))
        ws_d.merge_cells(start_row=i, start_column=1, end_row=i, end_column=8)
    ws_d.column_dimensions["A"].width = 80

    # SKU_Master
    ws_m = wb.create_sheet("SKU_Master")
    write_headers(ws_m, [
        "sku_id", "category", "abc_class", "xyz_class", "abc_xyz",
        "avg_weekly_demand", "sigma_weekly_demand", "unit_price",
        "lead_time_days", "sigma_lead_time_days", "is_critical",
    ])

    # Parameters (editable)
    ws_p = wb.create_sheet("Parameters")
    write_headers(ws_p, ["parameter_name", "value", "description"])
    params = [
        ("holding_rate_annual", 0.25, "Cost of holding inventory per year (% of unit value)"),
        ("stockout_cost_multiplier", 3.0, "Stockout cost = N x margin"),
        ("margin_rate", 0.30, "Gross margin rate"),
        ("order_placement_cost", 100.0, "Cost per PO ($)"),
        ("vmi_fee_rate_annual", 0.06, "VMI consignment fee (% of consigned stock value)"),
        ("vmi_setup_cost", 5000.0, "One-time VMI program setup cost ($)"),
        ("vmi_lt_days", 7, "Reduced LT under VMI (days)"),
        ("service_level", 0.95, "Target service level (Z-score = NORM.S.INV(.95) = 1.645)"),
    ]
    for i, (n, v, d) in enumerate(params, start=2):
        ws_p.cell(row=i, column=1, value=n)
        ws_p.cell(row=i, column=2, value=v)
        ws_p.cell(row=i, column=3, value=d)
    ws_p.column_dimensions["A"].width = 28
    ws_p.column_dimensions["B"].width = 12
    ws_p.column_dimensions["C"].width = 55

    # Empty output sheets
    ws_roi = wb.create_sheet("ROI_Output")
    write_headers(ws_roi, [
        "sku_id", "category", "abc_xyz", "annual_volume", "annual_revenue",
        "lead_time_days", "status_quo_cost", "vmi_cost",
        "annual_savings", "payback_months", "recommendation",
    ])

    ws_c = wb.create_sheet("Top_Candidates")
    write_headers(ws_c, [
        "rank", "sku_id", "category", "abc_xyz",
        "annual_revenue", "lead_time_days", "annual_savings",
        "payback_months", "recommendation", "candidacy_score",
    ])

    ws_log = wb.create_sheet("Log")
    write_headers(ws_log, ["Timestamp", "Level", "Source", "Message"])

    out = Path(__file__).resolve().parent / "VMI_ROI.xlsx"
    wb.save(out)
    print(f"[OK] Tao xong: {out}")
    print()
    print("Buoc tiep theo:")
    print("  1. Mo VMI_ROI.xlsx > Save As Macro-Enabled (.xlsm)")
    print("  2. Alt+F11 > Import 5 .bas files trong vba_modules/")
    print("  3. Tao 5 nut tren Dashboard:")
    print("     - Action_RunFullAnalysis")
    print("     - Action_GenerateSample")
    print("     - Action_ComputeROI")
    print("     - Action_TopCandidates")
    print("     - Action_ResetOutputs")
    print("  4. Bam 'Run Full Analysis' de test")


if __name__ == "__main__":
    main()
