import sys
import os
import csv
import argparse

try:
    import win32com.client
    import pythoncom
except ImportError:
    print("ERROR: pywin32 not installed. Run: pip install pywin32")
    sys.exit(1)

EXCEL_ERR = -2146826273


def safe_float(val):
    if val is None or val == EXCEL_ERR:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    return None


def main():
    parser = argparse.ArgumentParser(description="Export Excel tree data to data.csv")
    parser.add_argument("--project", default="Projects/7631-creditview",
                        help="Project directory (default: Projects/7631-creditview)")
    parser.add_argument("--excel", default=None, help="Override Excel file path")
    args = parser.parse_args()

    project_dir = os.path.abspath(args.project)
    if args.excel:
        excel_path = os.path.abspath(args.excel)
    else:
        excel_path = os.path.join(project_dir, "7631 Creditview Rd.xlsx")

    csv_path = os.path.join(project_dir, "data.csv")

    print(f"Reading Excel: {excel_path}")
    pythoncom.CoInitialize()
    xl = win32com.client.Dispatch("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    wb = xl.Workbooks.Open(excel_path)
    ws = wb.Sheets("Sheet1")

    # Find last data row — single COM call via UsedRange
    used = ws.UsedRange
    last_row = used.Row + used.Rows.Count - 1

    if last_row < 3:
        wb.Close(False)
        print("No data found in Excel.")
        return

    # Bulk read cols A–R (1–18) in one COM call
    data = ws.Range(ws.Cells(3, 1), ws.Cells(last_row, 18)).Value
    wb.Close(False)

    # Single-row result is a flat tuple — normalise
    if last_row == 3:
        data = (data,)

    rows_written = 0
    rows_skipped = 0
    with open(csv_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(['tree_num', 'direction', 'tpz_m', 'tpz_circle_m', 'cx', 'cy', 'trunk_mm'])
        for row in data:
            tree_num_raw = row[0]   # col A
            if tree_num_raw is None:
                continue
            direction    = row[8]   # col I
            tpz_m_raw    = row[9]   # col J  — label display value (may have formula errors)
            tpz_circ_raw = row[10]  # col K  — clean TPZ radius m for circle sizing
            cx_raw       = row[15]  # col P  — Center X AI pts
            cy_raw       = row[16]  # col Q  — Center Y AI pts
            trunk_raw    = row[17]  # col R  — Trunk 1:500 mm

            tree_num = (str(int(tree_num_raw)) if isinstance(tree_num_raw, (int, float))
                        else str(tree_num_raw).strip())
            direction_str = str(direction).strip() if direction else ''

            tpz_m        = safe_float(tpz_m_raw)
            tpz_circle_m = safe_float(tpz_circ_raw)
            trunk_mm     = safe_float(trunk_raw)
            cx           = safe_float(cx_raw)
            cy           = safe_float(cy_raw)

            if cx is None or cy is None:
                writer.writerow([tree_num, direction_str,
                                 f'{tpz_m:.2f}' if tpz_m is not None else '',
                                 f'{tpz_circle_m:.2f}' if tpz_circle_m is not None else '',
                                 '', '',
                                 f'{trunk_mm:.2f}' if trunk_mm is not None else ''])
                rows_skipped += 1
                continue

            # Apply LogLog → AI artboard transform if coords are outside artboard bounds.
            # Artboard: left=-684, top=468, right=1908, bottom=-1260
            if cy > 468 or cy < -1260 or cx < -684 or cx > 1908:
                cx, cy = -684 + cy, 468 - cx

            writer.writerow([
                tree_num, direction_str,
                f'{tpz_m:.2f}'        if tpz_m        is not None else '',
                f'{tpz_circle_m:.2f}' if tpz_circle_m is not None else '',
                f'{cx:.2f}', f'{cy:.2f}',
                f'{trunk_mm:.2f}'     if trunk_mm     is not None else '',
            ])
            rows_written += 1

    print(f"  Wrote {rows_written} trees (+ {rows_skipped} without coords) to {csv_path}")


if __name__ == "__main__":
    main()
