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

    var fontName = "Arial-BoldMT", fontSize = 5;

    // isCircle helper (from review.py)
    function isCircle(item) {
        if (item.typename !== 'PathItem') return false;
        var pp = item.pathPoints;
        if (pp.length < 4 || pp.length > 5) return false;
        var gb = item.geometricBounds;
        var w = Math.abs(gb[2] - gb[0]);
        var h = Math.abs(gb[1] - gb[3]);
        if (w <= 0 || h <= 0) return false;
        return Math.abs(w - h) / Math.max(w, h) < 0.15;
    }

    // directPathItems helper (from review.py)
    function directPathItems(layer) {
        var result = [];
        var items = layer.pathItems;
        for (var i = 0; i < items.length; i++) {
            var parentName = '';
            try { parentName = items[i].parent.name; } catch(e) {}
            if (parentName === layer.name) result.push(items[i]);
        }
        return result;
    }

    // 2 — Gather circle bounds from TPZs layer, expanded by 4pt margin
    var tpzLayer = null;
    for (var li = 0; li < doc.layers.length; li++) {
        if (doc.layers[li].name === 'TPZs') { tpzLayer = doc.layers[li]; break; }
    }
    if (!tpzLayer) return '{"error":"TPZs layer not found"}';

    var circleObstacles = [];
    var allDirect = directPathItems(tpzLayer);
    for (var i = 0; i < allDirect.length; i++) {
        if (isCircle(allDirect[i])) {
            var cg = allDirect[i].geometricBounds;
            circleObstacles.push([cg[0]-4, cg[1]+4, cg[2]+4, cg[3]-4]);
        }
    }

    // 3 — Labels layer: find or create, clear contents
    var labelsLayer = null;
    for (var li = 0; li < doc.layers.length; li++) {
        if (doc.layers[li].name === 'Labels') { labelsLayer = doc.layers[li]; break; }
    }
    if (labelsLayer) {
        while (labelsLayer.pageItems.length > 0) labelsLayer.pageItems[0].remove();
    } else {
        labelsLayer = doc.layers.add();
        labelsLayer.name = 'Labels';
    }
    labelsLayer.visible = true;
    labelsLayer.locked = false;
    labelsLayer.zOrder(ZOrderMethod.BRINGTOFRONT);

    // Colour helpers
    function cmyk(c, m, y, k) {
        var col = new CMYKColor();
        col.cyan = c; col.magenta = m; col.yellow = y; col.black = k;
        return col;
    }
    var GREEN  = cmyk(70, 0,  100, 0);
    var ORANGE = cmyk(0,  62, 100, 0);
    var BLACK  = cmyk(0,  0,  0,   100);
    var WHITE  = cmyk(0,  0,  0,   0);

    // AABB overlap — AI coords: [left, top, right, bottom] where top > bottom
    function overlaps(a, b) {
        return a[0] < b[2] && a[2] > b[0] && a[3] < b[1] && a[1] > b[3];
    }

    // Place a temp text frame at (lx, ly), read bounds, remove
    function measureLabel(contents, lx, ly, fName, fSize, just) {
        var tf = labelsLayer.textFrames.add();
        tf.contents = contents;
        var ca = tf.textRange.characterAttributes;
        try { ca.textFont = app.textFonts.getByName(fName); } catch(e) {}
        ca.size = fSize;
        ca.fillColor = BLACK;
        tf.textRange.paragraphAttributes.justification = just;
        tf.position = [lx, ly];
        var gb = tf.geometricBounds;
        tf.remove();
        return gb;
    }

    // rotate(45) rotates around the frame CENTER, not the anchor.
    // Pre-compute pre-rotation position so that after rotate(45) the text
    // baseline-left anchor lands exactly at (cx, cy).
    // If cxOff/cyOff = bb-center offset from anchor (measured at origin),
    // then: px = cx - cxOff*(1-s) - cyOff*s
    //        py = cy + cxOff*s    - cyOff*(1-s)   where s = sin(45) = cos(45)
    var S45 = Math.SQRT2 / 2;  // sin(45) = cos(45) ≈ 0.7071
    var placed = 0;
    var placedBounds = [];

    for (var ti = 0; ti < TREES.length; ti++) {
        var tree = TREES[ti];
        var cx = tree.cx, cy = tree.cy;

        // Compact label: "#1 Pro 12.0m" / "#1 Inj 10.0m" / "#3 Rem"
        var abbr = tree.dir === 'Protect' ? 'Pro' : tree.dir === 'Injury' ? 'Inj' : 'Rem';
        var contents;
        if (tree.tpz_m === null) {
            contents = '#' + tree.num + ' ' + abbr;
        } else {
            var tpzStr = (Math.round(tree.tpz_m * 10) / 10).toFixed(1);
            contents = '#' + tree.num + ' ' + abbr + ' ' + tpzStr + 'm';
        }

        var color = (tree.dir === 'Protect') ? GREEN : ORANGE;

        // Measure at origin to get bb-center offset from anchor
        var gbM = measureLabel(contents, 0, 0, fontName, fontSize, Justification.LEFT);
        var cxOff = (gbM[0] + gbM[2]) / 2;
        var cyOff = (gbM[1] + gbM[3]) / 2;

        // Offset so bg's start corner lands at trunk dot edge (dot radius=1, pad=0.5)
        var anchorX = cx + (1 + 0.5) * S45;
        var anchorY = cy + (1 + 0.5) * S45;
        // Pre-rotation position so post-rotate(45) text anchor lands at (anchorX, anchorY)
        var px = anchorX - cxOff * (1 - S45) - cyOff * S45;
        var py = anchorY + cxOff * S45       - cyOff * (1 - S45);

        // Text frame
        var tf = labelsLayer.textFrames.add();
        tf.contents = contents;
        var ca = tf.textRange.characterAttributes;
        try { ca.textFont = app.textFonts.getByName(fontName); } catch(e) {}
        ca.size = fontSize;
        ca.fillColor = BLACK;
        tf.textRange.paragraphAttributes.justification = Justification.LEFT;
        tf.position = [px, py];

        // Background rect — read bounds BEFORE rotating so it fits tightly
        var finalGb = tf.geometricBounds;
        var pad = 0.5;
        var bg = labelsLayer.pathItems.rectangle(
            finalGb[1] + pad, finalGb[0] - pad,
            (finalGb[2] - finalGb[0]) + pad * 2,
            (finalGb[1] - finalGb[3]) + pad * 2
        );
        bg.filled = true;
        bg.fillColor = WHITE;
        bg.stroked = true;
        bg.strokeColor = BLACK;
        bg.strokeWidth = 0.25;

        // Rotate both around their shared centre
        tf.rotate(45);
        bg.rotate(45);
        bg.zOrder(ZOrderMethod.SENDBACKWARD);

        // Dot placed LAST so it sits on top of the bg rect
        var dot = labelsLayer.pathItems.ellipse(cy + 1, cx - 1, 2, 2);
        dot.filled = true;
        dot.fillColor = color;
        dot.stroked = false;

        // Group dot + tf + bg into one selectable label unit
        var grp = labelsLayer.groupItems.add();
        dot.moveToEnd(grp);
        tf.moveToEnd(grp);
        bg.moveToEnd(grp);

        placedBounds.push(finalGb);
        placed++;
    }

    return '{"placed":' + placed + '}';
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
        tree_num_raw = ws.Cells(row, 1).Value   # col A
        if tree_num_raw is None:
            break
        direction = ws.Cells(row, 9).Value      # col I
        tpz_raw   = ws.Cells(row, 10).Value     # col J
        size_raw  = ws.Cells(row, 14).Value     # col N
        cx_raw    = ws.Cells(row, 15).Value     # col O
        cy_raw    = ws.Cells(row, 16).Value     # col P

        tpz_m   = float(tpz_raw)  if isinstance(tpz_raw,  (int, float)) else None
        size_mm = float(size_raw) if isinstance(size_raw, (int, float)) and size_raw > 0 else None

        if not isinstance(cx_raw, (int, float)) or not isinstance(cy_raw, (int, float)):
            row += 1
            continue

        trees.append({
            "num":    str(int(tree_num_raw)) if isinstance(tree_num_raw, (int, float)) else str(tree_num_raw).strip(),
            "dir":    str(direction) if direction else "",
            "tpz_m":  tpz_m,
            "size_mm": size_mm,
            "cx":     float(cx_raw),
            "cy":     float(cy_raw),
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
    parser = argparse.ArgumentParser(description="Place TPZ labels in Illustrator")
    parser.add_argument("--limit", type=int, default=None, help="Process only first N trees")
    args = parser.parse_args()

    print(f"Reading Excel: {EXCEL_PATH}")
    trees = read_excel(EXCEL_PATH)
    print(f"  {len(trees)} trees loaded")

    if args.limit is not None:
        trees = trees[:args.limit]
        print(f"  Limited to first {len(trees)} trees")

    jsx = "var TREES = " + json.dumps(trees) + ";\n" + JSX_BODY

    print("Placing labels in Illustrator...")
    raw = run_jsx(jsx)
    data = json.loads(raw)

    if "error" in data:
        print(f"ERROR: {data['error']}")
        sys.exit(1)

    print(f"Placed {data['placed']} labels.")


if __name__ == "__main__":
    main()
