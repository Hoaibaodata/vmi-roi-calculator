"""Entry point — generate SKU master, compute VMI ROI per SKU, rank candidates."""
from __future__ import annotations

from pathlib import Path

import pandas as pd

from src.recommendation_engine import compute_vmi_roi, recommend_vmi_candidates
from src.synthetic_sku import generate_synthetic_sku_master


def main() -> None:
    project_root = Path(__file__).resolve().parent
    data_dir = project_root / "data"
    export_dir = project_root / "exports"
    data_dir.mkdir(parents=True, exist_ok=True)
    export_dir.mkdir(parents=True, exist_ok=True)

    print("Generating synthetic SKU master (200 SKUs)...")
    sku_master = generate_synthetic_sku_master()
    master_path = data_dir / "sku_master.csv"
    sku_master.to_csv(master_path, index=False)
    print(f"  Saved: {master_path}")

    print("\nComputing VMI ROI for each SKU...")
    roi_df = compute_vmi_roi(sku_master)
    roi_path = export_dir / "vmi_roi_full.csv"
    roi_df.to_csv(roi_path, index=False)
    print(f"  Full ROI table: {roi_path}")

    print("\nRanking VMI candidates...")
    candidates = recommend_vmi_candidates(roi_df, top_n=20)
    cand_path = export_dir / "vmi_top_candidates.csv"
    candidates.to_csv(cand_path, index=False)
    print(f"  Top candidates: {cand_path}")

    # Summary
    print("\n" + "=" * 80)
    print("VMI ROI ANALYSIS SUMMARY")
    print("=" * 80)
    print(f"Total SKUs analyzed: {len(roi_df)}")
    positive_savings = (roi_df["annual_savings"] > 0).sum()
    print(f"SKUs with positive VMI savings: {positive_savings} ({positive_savings/len(roi_df)*100:.0f}%)")
    print(f"Total potential annual savings: ${roi_df[roi_df['annual_savings'] > 0]['annual_savings'].sum():,.0f}")
    print(f"Top 20 candidates annual savings: ${candidates['annual_savings'].sum():,.0f}")

    print("\n" + "=" * 80)
    print("TOP 10 VMI CANDIDATES")
    print("=" * 80)
    display_cols = [
        "sku_id", "category", "abc_xyz", "annual_revenue", "lead_time_days",
        "annual_savings", "payback_months", "recommendation"
    ]
    print(candidates[display_cols].head(10).to_string(index=False))
    print("=" * 80)


if __name__ == "__main__":
    main()
