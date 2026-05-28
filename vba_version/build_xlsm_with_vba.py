"""Auto-build VMI_ROI.xlsm with VBA pre-imported (xlwings + Excel COM).

PREREQS: Excel installed + Trust Center setting + pip install xlwings.
Run generate_starter_xlsx.py FIRST.
"""
from __future__ import annotations

import sys
from pathlib import Path


def main() -> None:
    try:
        import xlwings as xw
    except ImportError:
        print("ERROR: xlwings not installed. Run: pip install xlwings")
        sys.exit(1)

    here = Path(__file__).resolve().parent
    xlsx_path = here / "VMI_ROI.xlsx"
    xlsm_path = here / "VMI_ROI.xlsm"
    bas_dir = here / "vba_modules"

    if not xlsx_path.exists():
        print(f"ERROR: {xlsx_path.name} not found. Run generate_starter_xlsx.py first.")
        sys.exit(1)

    bas_files = sorted(bas_dir.glob("*.bas"))

    print("Opening Excel...")
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    try:
        print(f"Loading {xlsx_path.name}...")
        wb = app.books.open(str(xlsx_path))
        if xlsm_path.exists():
            xlsm_path.unlink()
        print(f"Saving as .xlsm...")
        wb.api.SaveAs(str(xlsm_path), FileFormat=52)
        print(f"Importing {len(bas_files)} VBA modules...")
        for bas_file in bas_files:
            try:
                wb.api.VBProject.VBComponents.Import(str(bas_file))
                print(f"  OK  {bas_file.name}")
            except Exception as e:
                print(f"  FAIL {bas_file.name}: {e}")
                print("If 'programmatic access not trusted': enable in Trust Center > Macro Settings.")
                raise
        wb.save()
        wb.close()
        print(f"\n[OK] Built: {xlsm_path}")
        print("Add buttons per SETUP-VBA.md Buoc 4.")
    finally:
        app.quit()


if __name__ == "__main__":
    main()
