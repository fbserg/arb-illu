import sys
import os
import csv
import json
import argparse
import tempfile

try:
    import win32com.client
    import pythoncom
except ImportError:
    print("ERROR: pywin32 not installed. Run: pip install pywin32")
    sys.exit(1)

CSV_PATH  = r"C:\Projects\arborist-plans\Projects\7631-creditview\data.csv"
XLSX_PATH = r"C:\Projects\arborist-plans\Projects\7631-creditview\7631 Creditview Rd.xlsx"

JSX_BODY = r"""
(function() {
    var doc;
    try { doc = app.activeDocument; } catch(e) { return '{"error":"No document open"}'; }

    var fontName = "Arial-BoldMT", fontSize = 5;

    // 1 — Gather circle bounds from TPZs layer (circles are in groups)
    var tpzLayer = null;
    for (var li = 0; li < doc.layers.length; li++) {
        if (doc.layers[li].name === 'TPZs') { tpzLayer = doc.layers[li]; break; }
    }
    if (!tpzLayer) return '{"error":"TPZs layer not found"}';

    // 2 — Labels layer: find or create, clear contents
    var labelsLayer = null;
    for (var li = 0; li < doc.layers.length; li++) {
        if (doc.layers[li].name === 'Labels') { labelsLayer = doc.layers[li]; break; }
    }
    if (labelsLayer) {
        while (labelsLayer.groupItems.length > 0) labelsLayer.groupItems[0].remove();
        while (labelsLayer.pathItems.length  > 0) labelsLayer.pathItems[0].remove();
        while (labelsLayer.textFrames.length > 0) labelsLayer.textFrames[0].remove();
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

    for (var ti = 0; ti < TREES.length; ti++) {
        var tree = TREES[ti];
        var cx = tree.cx, cy = tree.cy;
        var dir = tree.dir;

        // Retain is visually identical to Protect
        var isProtect = (dir === 'Protect' || dir === 'Retain');
        var abbr  = isProtect ? 'Pro' : dir === 'Injury' ? 'Inj' : 'Rem';
        var color = isProtect ? GREEN : ORANGE;

        // Compact label: "#1 Pro 12.0m" / "#1 Inj 10.0m" / "#3 Rem"
        var contents;
        if (tree.tpz_m === null) {
            contents = '#' + tree.num + ' ' + abbr;
        } else {
            var tpzStr = (Math.round(tree.tpz_m * 10) / 10).toFixed(1);
            contents = '#' + tree.num + ' ' + abbr + ' ' + tpzStr + 'm';
        }

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

        placed++;
    }

    return '{"placed":' + placed + '}';
})();
"""


def read_csv(path):
    if os.path.exists(XLSX_PATH) and os.path.getmtime(XLSX_PATH) > os.path.getmtime(path):
        print("WARNING: Excel is newer than data.csv — run: python export_data.py")
    trees = []
    with open(path, newline='', encoding='utf-8') as f:
        for row in csv.DictReader(f):
            cx = float(row['cx']) if row['cx'] else None
            cy = float(row['cy']) if row['cy'] else None
            if cx is None or cy is None:
                continue
            tpz_m = float(row['tpz_m']) if row['tpz_m'] else None
            trees.append({
                'num':   row['tree_num'],
                'dir':   row['direction'],
                'tpz_m': tpz_m,
                'cx':    cx,
                'cy':    cy,
            })
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

    print(f"Reading CSV: {CSV_PATH}")
    trees = read_csv(CSV_PATH)
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
