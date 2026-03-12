"""Read Excel Master Sheet → write data.csv for place_tpz.py"""
import sys, os, csv

try:
    import win32com.client, pythoncom
except ImportError:
    print("ERROR: pywin32 not installed."); sys.exit(1)

EXCEL_PATH = r"C:\Projects\arborist-plans\Projects\scarlett\Excel Master Sheet.xlsx"
CSV_PATH   = r"C:\Projects\arborist-plans\Projects\scarlett\data.csv"
EXCEL_ERR  = -2146826273


def safe_float(val):
    if val is None or val == EXCEL_ERR:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    return None


def main():
    print(f"Reading: {EXCEL_PATH}")
    pythoncom.CoInitialize()
    xl = win32com.client.Dispatch("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    wb = xl.Workbooks.Open(EXCEL_PATH)
    ws = wb.Sheets("Inventory")

    used = ws.UsedRange
    last_row = used.Row + used.Rows.Count - 1

    # Headers at row 2, data from row 3
    # Cols: 1=TREE#, 8=Direction, 9=TPZ(m), 13=1:500(mm diam), 14=cx, 15=cy, 16=cx_phase1, 17=cy_phase1
    data = ws.Range(ws.Cells(3, 1), ws.Cells(last_row, 17)).Value
    wb.Close(False)

    if last_row < 3:
        print("No data rows found.")
        return

    if last_row == 3:
        data = (data,)

    written = skipped = 0
    with open(CSV_PATH, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(['tree_num', 'direction', 'tpz_m', 'tpz_mm', 'cx', 'cy', 'cx_phase1', 'cy_phase1'])
        for row in data:
            if row[0] is None:
                continue
            tree_num  = str(int(row[0]))
            direction = str(row[7]).strip() if row[7] else ''
            tpz_m     = safe_float(row[8])   # col 9 — TPZ radius m
            tpz_mm    = safe_float(row[12])  # col 13 — 1:500 diameter mm (formula result)
            cx        = safe_float(row[13])  # col 14
            cy        = safe_float(row[14])  # col 15
            cx_phase1 = safe_float(row[15])  # col 16
            cy_phase1 = safe_float(row[16])  # col 17

            writer.writerow([
                tree_num, direction,
                f'{tpz_m:.2f}'      if tpz_m      is not None else '',
                f'{tpz_mm:.2f}'     if tpz_mm     is not None else '',
                f'{cx:.2f}'         if cx          is not None else '',
                f'{cy:.2f}'         if cy          is not None else '',
                f'{cx_phase1:.4f}'  if cx_phase1   is not None else '',
                f'{cy_phase1:.4f}'  if cy_phase1   is not None else '',
            ])
            if cx is not None and cy is not None:
                written += 1
            else:
                skipped += 1

    print(f"Wrote {written} trees with coords (+{skipped} without) to {CSV_PATH}")


if __name__ == "__main__":
    main()
