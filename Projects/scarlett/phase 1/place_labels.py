import sys, os, csv, json, argparse, tempfile

try:
    import win32com.client, pythoncom
except ImportError:
    print("ERROR: pywin32 not installed."); sys.exit(1)

DATA_PATH  = r"C:\Projects\arborist-plans\Projects\scarlett\data.csv"
EXCEL_PATH = r"C:\Projects\arborist-plans\Projects\scarlett\Excel Master Sheet.xlsx"

# PLAN.ai artboard dims — used for coord transform
PLAN_W, PLAN_H = 2384, 3370

def transform(cx, cy):
    """PLAN.ai coords → template coords (90° CCW rotation around artboard centre)."""
    return (PLAN_W + PLAN_H) / 2 - cy, cx + (PLAN_H - PLAN_W) / 2

JSX_BODY = r"""
(function() {
    var doc;
    try { doc = app.activeDocument; } catch(e) { return '{"error":"No document open"}'; }

    var fontName = "Arial-BoldMT", fontSize = 5;

    function findOrCreateLayer(parent, name) {
        var coll = parent.layers;
        for (var i = 0; i < coll.length; i++) if (coll[i].name === name) return coll[i];
        var l = parent.layers.add(); l.name = name; return l;
    }

    var labelsLayer = findOrCreateLayer(doc, 'Labels');
    if (CLEAR_FIRST) {
        while (labelsLayer.groupItems.length > 0) labelsLayer.groupItems[0].remove();
        while (labelsLayer.pathItems.length  > 0) labelsLayer.pathItems[0].remove();
        while (labelsLayer.textFrames.length > 0) labelsLayer.textFrames[0].remove();
    }
    labelsLayer.visible = true;
    labelsLayer.locked = false;
    labelsLayer.zOrder(ZOrderMethod.BRINGTOFRONT);
    doc.activeLayer = labelsLayer;

    function cmyk(c, m, y, k) {
        var col = new CMYKColor();
        col.cyan = c; col.magenta = m; col.yellow = y; col.black = k;
        return col;
    }
    var GREEN  = cmyk(70, 0,  100, 0);
    var ORANGE = cmyk(0,  62, 100, 0);
    var BLACK  = cmyk(0,  0,  0,   100);
    var WHITE  = cmyk(0,  0,  0,   0);

    function measureLabel(contents, lx, ly) {
        var tf = labelsLayer.textFrames.add();
        tf.contents = contents;
        var ca = tf.textRange.characterAttributes;
        try { ca.textFont = app.textFonts.getByName(fontName); } catch(e) {}
        ca.size = fontSize;
        ca.fillColor = BLACK;
        tf.textRange.paragraphAttributes.justification = Justification.LEFT;
        tf.position = [lx, ly];
        var gb = tf.geometricBounds;
        tf.remove();
        return gb;
    }

    var S45 = Math.SQRT2 / 2;
    var placed = 0;

    for (var ti = 0; ti < TREES.length; ti++) {
        var tree = TREES[ti];
        var cx = tree.cx, cy = tree.cy;
        var dir = tree.dir;

        var isProtect = (dir === 'Protect' || dir === 'Retain');
        var abbr  = isProtect ? 'Pro' : (dir === 'Injury' ? 'Inj' : 'Rem');
        var color = isProtect ? GREEN : ORANGE;

        var contents;
        if (tree.tpz_m === null) {
            contents = '#' + tree.num + ' ' + abbr;
        } else {
            var tpzStr = (Math.round(tree.tpz_m * 10) / 10).toFixed(1);
            contents = '#' + tree.num + ' ' + abbr + ' ' + tpzStr + 'm';
        }

        var gbM = measureLabel(contents, 0, 0);
        var cxOff = (gbM[0] + gbM[2]) / 2;
        var cyOff = (gbM[1] + gbM[3]) / 2;

        var anchorX = cx - (1 + 0.5) * S45;
        var anchorY = cy - (1 + 0.5) * S45;
        var px = anchorX - cxOff * (1 - S45) - cyOff * S45;
        var py = anchorY + cxOff * S45       - cyOff * (1 - S45);

        var tf = labelsLayer.textFrames.add();
        tf.contents = contents;
        var ca = tf.textRange.characterAttributes;
        try { ca.textFont = app.textFonts.getByName(fontName); } catch(e) {}
        ca.size = fontSize;
        ca.fillColor = BLACK;
        tf.textRange.paragraphAttributes.justification = Justification.LEFT;
        tf.position = [px, py];

        var finalGb = tf.geometricBounds;
        var pad = 0.5;
        var bg = labelsLayer.pathItems.rectangle(
            finalGb[1] + pad, finalGb[0] - pad,
            (finalGb[2] - finalGb[0]) + pad * 2,
            (finalGb[1] - finalGb[3]) + pad * 2
        );
        bg.filled = true; bg.fillColor = WHITE;
        bg.stroked = true; bg.strokeColor = BLACK; bg.strokeWidth = 0.25;

        tf.rotate(45);
        bg.rotate(45);
        bg.zOrder(ZOrderMethod.SENDBACKWARD);

        var dot = labelsLayer.pathItems.ellipse(cy + 1, cx - 1, 2, 2);
        dot.filled = true; dot.fillColor = color; dot.stroked = false;

        var grp = labelsLayer.groupItems.add();
        grp.name = 'Tree ' + tree.num;
        dot.moveToEnd(grp);
        tf.moveToEnd(grp);
        bg.moveToEnd(grp);

        placed++;
    }

    return '{"placed":' + placed + '}';
})();
"""


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--offset", type=int, default=0)
    parser.add_argument("--limit",  type=int, default=50)
    parser.add_argument("--all",    action="store_true")
    args = parser.parse_args()

    if os.path.exists(EXCEL_PATH) and os.path.getmtime(EXCEL_PATH) > os.path.getmtime(DATA_PATH):
        print("WARNING: Excel is newer than data.csv — run: python Projects/scarlett/export_data.py")

    trees = []
    with open(DATA_PATH, newline='', encoding='utf-8') as f:
        for row in csv.DictReader(f):
            cx = float(row['cx']) if row['cx'] else None
            cy = float(row['cy']) if row['cy'] else None
            if cx is None or cy is None:
                continue
            cx, cy = transform(cx, cy)
            trees.append({
                'num':   row['tree_num'],
                'dir':   row['direction'],
                'tpz_m': float(row['tpz_m']) if row['tpz_m'] else None,
                'cx':    cx,
                'cy':    cy,
            })

    trees = trees[args.offset:]
    if not args.all:
        trees = trees[:args.limit]

    clear_first = (args.offset == 0)
    print(f"Placing {len(trees)} labels (offset={args.offset}, {'clearing' if clear_first else 'appending'})")

    jsx = (
        "var CLEAR_FIRST = " + ("true" if clear_first else "false") + ";\n"
        "var TREES = " + json.dumps(trees) + ";\n"
        + JSX_BODY
    )

    pythoncom.CoInitialize()
    tmp = tempfile.NamedTemporaryFile(suffix=".jsx", delete=False, mode="w", encoding="utf-8")
    tmp.write(jsx); path = tmp.name; tmp.close()
    try:
        ai = win32com.client.GetActiveObject("Illustrator.Application")
        raw = ai.DoJavaScriptFile(path)
    finally:
        os.unlink(path)

    raw = str(raw) if raw is not None else ""
    if not raw.strip().startswith("{"):
        raise RuntimeError("JS returned: " + raw)
    data = json.loads(raw)

    if "error" in data:
        print(f"ERROR: {data['error']}"); sys.exit(1)
    print(f"Placed {data['placed']} labels.")

    if not args.all and len(trees) == args.limit:
        print(f"\nNext batch: python Projects/scarlett/place_labels.py --offset {args.offset + args.limit}")


if __name__ == "__main__":
    main()
