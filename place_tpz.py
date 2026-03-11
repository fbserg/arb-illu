import sys
import os
import json
import argparse
import tempfile

try:
    import win32com.client
    import pythoncom
except ImportError:
    print("ERROR: pywin32 not installed. Run: pip install pywin32")
    sys.exit(1)

EXCEL_PATH = r"C:\Projects\arborist-plans\Projects\7631-creditview\7631 Creditview Rd.xlsx"

JSX_BODY = r"""
(function() {
    var doc;
    try { doc = app.activeDocument; } catch(e) { return '{"error":"No document open"}'; }

    var PT_PER_MM = 72.0 / 25.4;

    // Colour helpers
    function cmyk(c, m, y, k) {
        var col = new CMYKColor();
        col.cyan = c; col.magenta = m; col.yellow = y; col.black = k;
        return col;
    }
    var GREEN  = cmyk(70, 0,  100, 0);
    var ORANGE = cmyk(0,  62, 100, 0);

    // Find or create a layer by name under parent
    function findOrCreateLayer(parent, name) {
        var coll = parent.layers;
        for (var i = 0; i < coll.length; i++) {
            if (coll[i].name === name) return coll[i];
        }
        var l = parent.layers.add();
        l.name = name;
        return l;
    }

    // Clear a layer using only typed collections (never pageItems)
    function clearLayer(layer) {
        while (layer.pathItems.length > 0) layer.pathItems[0].remove();
        while (layer.textFrames.length > 0) layer.textFrames[0].remove();
        while (layer.groupItems.length > 0) layer.groupItems[0].remove();
    }

    // Get artboard bounds for out-of-bounds check
    var ab = doc.artboards[0].artboardRect;
    var abLeft = ab[0], abTop = ab[1], abRight = ab[2], abBottom = ab[3];

    // Find or create layers
    var tpzLayer    = findOrCreateLayer(doc, 'TPZs');
    var trunkLayer  = findOrCreateLayer(tpzLayer, 'Trunks');

    tpzLayer.visible  = true; tpzLayer.locked  = false;
    trunkLayer.visible = true; trunkLayer.locked = false;

    // Clear existing circles (idempotent re-run)
    clearLayer(tpzLayer);
    clearLayer(trunkLayer);

    var tpzPlaced   = 0;
    var trunksPlaced = 0;
    var skipped = [];
    var errors  = [];

    for (var ti = 0; ti < TREES.length; ti++) {
        var tree = TREES[ti];
        var cx = tree.cx, cy = tree.cy;
        var tpzMm   = tree.tpz_mm;
        var trunkMm = tree.trunk_mm;
        var dir     = tree.dir;

        if (tpzMm === null || tpzMm <= 0) {
            skipped.push({ num: tree.num, reason: "no TPZ size" });
            continue;
        }

        var tpzDiam = tpzMm * PT_PER_MM;
        var tpzRad  = tpzDiam / 2;

        // Bounds check
        if (cx < abLeft || cx > abRight || cy > abTop || cy < abBottom) {
            skipped.push({ num: tree.num, reason: "outside artboard" });
            continue;
        }

        var color = (dir === 'Protect') ? GREEN : ORANGE;

        // TPZ circle: ellipse(top, left, width, height) in AI coords
        try {
            var circle = tpzLayer.pathItems.ellipse(cy + tpzRad, cx - tpzRad, tpzDiam, tpzDiam);
            circle.filled  = false;
            circle.stroked = true;
            circle.strokeColor = color;
            circle.strokeWidth = 1.5;

            if (dir === 'Injury') {
                circle.strokeDashes = [4, 2];
            } else {
                circle.strokeDashes = [];
            }

            if (dir === 'Removal') {
                // Two diagonal X-lines through center
                var r = tpzRad * 0.707; // sin(45)
                var line1 = tpzLayer.pathItems.add();
                line1.setEntirePath([[cx - r, cy - r], [cx + r, cy + r]]);
                line1.filled  = false;
                line1.stroked = true;
                line1.strokeColor = ORANGE;
                line1.strokeWidth = 1.0;

                var line2 = tpzLayer.pathItems.add();
                line2.setEntirePath([[cx - r, cy + r], [cx + r, cy - r]]);
                line2.filled  = false;
                line2.stroked = true;
                line2.strokeColor = ORANGE;
                line2.strokeWidth = 1.0;
            }

            tpzPlaced++;
        } catch(e) {
            errors.push({ num: tree.num, error: e.toString() });
            continue;
        }

        // Trunk circle (if trunk_mm present)
        if (trunkMm !== null && trunkMm > 0) {
            try {
                var tDiam = trunkMm * PT_PER_MM;
                var tRad  = tDiam / 2;
                var trunk = trunkLayer.pathItems.ellipse(cy + tRad, cx - tRad, tDiam, tDiam);

                var fillColor = (dir === 'Protect') ? cmyk(70, 0, 100, 0) : cmyk(0, 62, 100, 0);
                trunk.filled     = true;
                trunk.fillColor  = fillColor;
                trunk.opacity    = 50;
                trunk.stroked    = false;

                trunksPlaced++;
            } catch(e) {
                errors.push({ num: tree.num, error: "trunk: " + e.toString() });
            }
        }
    }

    return JSON.stringify({
        tpz_placed:    tpzPlaced,
        trunks_placed: trunksPlaced,
        skipped:       skipped,
        errors:        errors
    });
})();
"""


def read_excel(path):
    pythoncom.CoInitialize()
    xl = win32com.client.Dispatch("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    wb = xl.Workbooks.Open(os.path.abspath(path))
    ws = wb.Sheets("Sheet1")

    trees = []
    row = 3
    while True:
        tree_num_raw = ws.Cells(row, 1).Value   # col A: tree number
        if tree_num_raw is None:
            break
        direction = ws.Cells(row, 9).Value      # col I: Protect/Injury/Removal
        tpz_raw   = ws.Cells(row, 14).Value     # col N: TPZ diameter (mm)
        cx_raw    = ws.Cells(row, 15).Value     # col O: center x (pt)
        cy_raw    = ws.Cells(row, 16).Value     # col P: center y (pt)
        trunk_raw = ws.Cells(row, 17).Value     # col Q: trunk diameter (mm)

        tpz_mm   = float(tpz_raw)   if isinstance(tpz_raw,   (int, float)) and tpz_raw  > 0 else None
        trunk_mm = float(trunk_raw) if isinstance(trunk_raw, (int, float)) and trunk_raw > 0 else None

        if not isinstance(cx_raw, (int, float)) or not isinstance(cy_raw, (int, float)):
            row += 1
            continue

        trees.append({
            "num":      str(int(tree_num_raw)) if isinstance(tree_num_raw, (int, float)) else str(tree_num_raw).strip(),
            "dir":      str(direction) if direction else "",
            "tpz_mm":   tpz_mm,
            "trunk_mm": trunk_mm,
            "cx":       float(cx_raw),
            "cy":       float(cy_raw),
        })
        row += 1

    wb.Close(False)
    return trees


def run_jsx(jsx_code):
    pythoncom.CoInitialize()
    tmp = tempfile.NamedTemporaryFile(suffix=".jsx", delete=False, mode="w", encoding="utf-8")
    tmp.write(jsx_code)
    tmp_path = tmp.name
    tmp.close()
    try:
        illustrator = win32com.client.GetActiveObject("Illustrator.Application")
        result = illustrator.DoJavaScriptFile(tmp_path)
    finally:
        os.unlink(tmp_path)
    result = str(result) if result is not None else ""
    if not result.strip().startswith("{"):
        raise RuntimeError("JS returned unexpected output:\n" + result)
    return result


def main():
    parser = argparse.ArgumentParser(description="Place TPZ circles and trunk indicators in Illustrator")
    parser.add_argument("--limit", type=int, default=None, help="Process only first N trees")
    args = parser.parse_args()

    print(f"Reading Excel: {EXCEL_PATH}")
    trees = read_excel(EXCEL_PATH)
    print(f"  {len(trees)} trees loaded")

    if args.limit is not None:
        trees = trees[:args.limit]
        print(f"  Limited to first {len(trees)} trees")

    jsx = "var TREES = " + json.dumps(trees) + ";\n" + JSX_BODY

    print("Placing TPZ circles and trunks in Illustrator...")
    raw = run_jsx(jsx)
    data = json.loads(raw)

    if "error" in data:
        print(f"ERROR: {data['error']}")
        sys.exit(1)

    print(f"TPZ circles placed:  {data['tpz_placed']}")
    print(f"Trunk circles placed: {data['trunks_placed']}")
    if data.get("skipped"):
        print(f"Skipped ({len(data['skipped'])}):")
        for s in data["skipped"]:
            print(f"  Tree #{s['num']}: {s['reason']}")
    if data.get("errors"):
        print(f"Errors ({len(data['errors'])}):")
        for e in data["errors"]:
            print(f"  Tree #{e['num']}: {e['error']}")


if __name__ == "__main__":
    main()
