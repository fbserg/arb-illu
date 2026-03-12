import os, csv, tempfile

try:
    import win32com.client, pythoncom
except ImportError:
    import sys
    print("ERROR: pywin32 not installed."); sys.exit(1)

DATA_PATH  = r"C:\Projects\arborist-plans\Projects\scarlett\data.csv"
EXCEL_PATH = r"C:\Projects\arborist-plans\Projects\scarlett\Excel Master Sheet.xlsx"
PLAN_W, PLAN_H = 2384, 3370


def transform(cx, cy):
    """PLAN.ai coords → template coords (90° CCW rotation around artboard centre)."""
    return (PLAN_W + PLAN_H) / 2 - cy, cx + (PLAN_H - PLAN_W) / 2


def normalize_dir(d):
    """Collapse Excel variants like 'Remove (1)' → 'Remove'."""
    dl = d.lower()
    if dl.startswith('remove') or dl.startswith('removal'):
        return 'Remove'
    if dl.startswith('injur'):
        return 'Injury'
    if dl.startswith('retain'):
        return 'Retain'
    if dl.startswith('protect'):
        return 'Protect'
    return d


def check_excel_staleness():
    if os.path.exists(EXCEL_PATH) and os.path.getmtime(EXCEL_PATH) > os.path.getmtime(DATA_PATH):
        print("WARNING: Excel is newer than data.csv — run: python Projects/scarlett/export_data.py")


def run_jsx(jsx_code):
    pythoncom.CoInitialize()
    tmp = tempfile.NamedTemporaryFile(suffix=".jsx", delete=False, mode="w", encoding="utf-8")
    tmp.write(jsx_code); path = tmp.name; tmp.close()
    try:
        ai = win32com.client.GetActiveObject("Illustrator.Application")
        result = ai.DoJavaScriptFile(path)
    finally:
        os.unlink(path)
    result = str(result) if result is not None else ""
    if not result.strip().startswith("{"):
        raise RuntimeError("JS returned: " + result)
    return result


def load_trees():
    """Read data.csv, apply coord transform + direction normalisation, return list of tree dicts."""
    trees = []
    with open(DATA_PATH, newline='', encoding='utf-8') as f:
        for row in csv.DictReader(f):
            cx = float(row['cx']) if row['cx'] else None
            cy = float(row['cy']) if row['cy'] else None
            if cx is None or cy is None:
                continue
            cx, cy = transform(cx, cy)
            trees.append({
                'num':    row['tree_num'],
                'dir':    normalize_dir(row['direction']),
                'tpz_m':  float(row['tpz_m'])  if row['tpz_m']  else None,
                'tpz_mm': float(row['tpz_mm']) if row['tpz_mm'] else None,
                'cx':     cx,
                'cy':     cy,
            })
    return trees
