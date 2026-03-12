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

# CLEAR_FIRST and TREES are injected by Python before this body runs.
JSX_BODY = r"""
(function() {
    var doc;
    try { doc = app.activeDocument; } catch(e) { return '{"error":"No document open"}'; }

    var PT_PER_MM = 72.0 / 25.4;

    function cmyk(c, m, y, k) {
        var col = new CMYKColor();
        col.cyan = c; col.magenta = m; col.yellow = y; col.black = k;
        return col;
    }
    var GREEN  = cmyk(70, 0,  100, 0);
    var ORANGE = cmyk(0,  62, 100, 0);

    function findOrCreateLayer(parent, name) {
        var coll = parent.layers;
        for (var i = 0; i < coll.length; i++) {
            if (coll[i].name === name) return coll[i];
        }
        var l = parent.layers.add();
        l.name = name;
        return l;
    }

    var ab = doc.artboards[0].artboardRect;
    var abLeft = ab[0], abTop = ab[1], abRight = ab[2], abBottom = ab[3];

    var tpzLayer   = findOrCreateLayer(doc, 'TPZs');
    tpzLayer.visible = true; tpzLayer.locked = false;
    doc.activeLayer = tpzLayer;

    if (CLEAR_FIRST) {
        // Groups clear their children in one shot — much faster than path-by-path removal.
        // Remove sublayers first so pathItems is non-recursive before the safety sweep.
        for (var di = tpzLayer.layers.length - 1; di >= 0; di--) {
            tpzLayer.layers[di].remove();
        }
        while (tpzLayer.groupItems.length > 0) tpzLayer.groupItems[0].remove();
        while (tpzLayer.pathItems.length  > 0) tpzLayer.pathItems[0].remove();
    }

    var tpzPlaced    = 0;
    var trunksPlaced = 0;
    var skipped = [];
    var errors  = [];

    for (var ti = 0; ti < TREES.length; ti++) {
        var tree = TREES[ti];
        var cx = tree.cx, cy = tree.cy;
        var tpzMm   = tree.tpz_mm;
        var trunkMm = tree.trunk_mm;
        var dir     = tree.dir;

        // Portrait artboard = PDF-native space; convert from landscape artboard coords
        // Inverse of: cx = -684 + y_pdf, cy = 468 - x_pdf
        if ((abTop - abBottom) > (abRight - abLeft)) {
            var x_pdf = 468 - cy;
            var y_pdf = cx + 684;
            cx = x_pdf;
            cy = y_pdf;
        }

        if (tpzMm === null || tpzMm <= 0) {
            skipped.push({ num: tree.num, reason: "no TPZ size" });
            continue;
        }

        var tpzDiam = tpzMm * PT_PER_MM;
        var tpzRad  = tpzDiam / 2;

        if (cx < abLeft || cx > abRight || cy > abTop || cy < abBottom) {
            skipped.push({ num: tree.num, reason: "outside artboard" });
            continue;
        }

        // Retain is visually identical to Protect (green, solid, no X)
        var color = (dir === 'Protect' || dir === 'Retain') ? GREEN : ORANGE;

        try {
            // Named group per tree — allows targeted select/delete/move
            var grp = tpzLayer.groupItems.add();
            grp.name = 'Tree ' + tree.num;

            // Create items on the layer, then move into group (last moved = on top)
            var circle = tpzLayer.pathItems.ellipse(cy + tpzRad, cx - tpzRad, tpzDiam, tpzDiam);
            circle.filled       = false;
            circle.stroked      = true;
            circle.strokeColor  = color;
            circle.strokeWidth  = 0.84;
            circle.strokeDashes = (dir === 'Injury') ? [5] : [];
            circle.move(grp, ElementPlacement.PLACEATBEGINNING);

            if (dir === 'Remove' || dir === 'Removal') {
                var r = tpzRad * 0.707;
                var line1 = tpzLayer.pathItems.add();
                line1.setEntirePath([[cx - r, cy - r], [cx + r, cy + r]]);
                line1.filled = false; line1.stroked = true;
                line1.strokeColor = ORANGE; line1.strokeWidth = 0.84;
                line1.move(grp, ElementPlacement.PLACEATBEGINNING);

                var line2 = tpzLayer.pathItems.add();
                line2.setEntirePath([[cx - r, cy + r], [cx + r, cy - r]]);
                line2.filled = false; line2.stroked = true;
                line2.strokeColor = ORANGE; line2.strokeWidth = 0.84;
                line2.move(grp, ElementPlacement.PLACEATBEGINNING);
            }

            tpzPlaced++;

            // Trunk dot inside the same group (moved last = sits on top)
            if (trunkMm !== null && trunkMm > 0) {
                var tDiam = trunkMm * PT_PER_MM;
                var tRad  = tDiam / 2;
                var trunk = tpzLayer.pathItems.ellipse(cy + tRad, cx - tRad, tDiam, tDiam);
                var fillColor = (dir === 'Protect' || dir === 'Retain') ? cmyk(70,0,100,0) : cmyk(0,62,100,0);
                trunk.filled    = true;
                trunk.fillColor = fillColor;
                trunk.opacity   = 50;
                trunk.stroked   = false;
                trunk.move(grp, ElementPlacement.PLACEATBEGINNING);
                trunksPlaced++;
            }
        } catch(e) {
            errors.push({ num: tree.num, error: e.toString() });
        }
    }

    var skJson = '[';
    for (var si = 0; si < skipped.length; si++) {
        if (si > 0) skJson += ',';
        skJson += '{"num":"' + skipped[si].num + '","reason":"' + skipped[si].reason.replace(/"/g,'\\"') + '"}';
    }
    skJson += ']';
    var erJson = '[';
    for (var ei = 0; ei < errors.length; ei++) {
        if (ei > 0) erJson += ',';
        erJson += '{"num":"' + errors[ei].num + '","error":"' + errors[ei].error.replace(/"/g,'\\"') + '"}';
    }
    erJson += ']';
    return '{"tpz_placed":' + tpzPlaced + ',"trunks_placed":' + trunksPlaced +
        ',"skipped":' + skJson + ',"errors":' + erJson + '}';
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
            tpz_circle_m = float(row['tpz_circle_m']) if row['tpz_circle_m'] else None
            tpz_mm = tpz_circle_m * 4.0 if tpz_circle_m else None
            trunk_mm = float(row['trunk_mm']) if row['trunk_mm'] else 2.0
            trees.append({
                'num':      row['tree_num'],
                'dir':      row['direction'],
                'tpz_mm':   tpz_mm,
                'trunk_mm': trunk_mm,
                'cx':       cx,
                'cy':       cy,
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
    parser = argparse.ArgumentParser(description="Place TPZ circles in Illustrator (10 trees at a time recommended)")
    parser.add_argument("--offset", type=int, default=0,   help="Skip first N trees (for batching)")
    parser.add_argument("--limit",  type=int, default=10,  help="Process at most N trees (default 10)")
    parser.add_argument("--all",    action="store_true",   help="Process all trees in one shot (overrides --limit)")
    args = parser.parse_args()

    print(f"Reading CSV: {CSV_PATH}")
    trees = read_csv(CSV_PATH)
    print(f"  {len(trees)} trees loaded")

    trees = trees[args.offset:]
    if not args.all:
        trees = trees[:args.limit]

    clear_first = (args.offset == 0)

    print(f"  Placing trees {args.offset + 1}–{args.offset + len(trees)} "
          f"({'clearing layer first' if clear_first else 'appending'})")

    jsx = (
        "var TREES = " + json.dumps(trees) + ";\n" +
        "var CLEAR_FIRST = " + ("true" if clear_first else "false") + ";\n" +
        JSX_BODY
    )

    print("Sending to Illustrator...")
    raw = run_jsx(jsx)
    data = json.loads(raw)

    if "error" in data:
        print(f"ERROR: {data['error']}")
        sys.exit(1)

    print(f"TPZ circles placed:   {data['tpz_placed']}")
    print(f"Trunk circles placed: {data['trunks_placed']}")
    if data.get("skipped"):
        print(f"Skipped ({len(data['skipped'])}):")
        for s in data["skipped"]:
            print(f"  Tree #{s['num']}: {s['reason']}")
    if data.get("errors"):
        print(f"Errors ({len(data['errors'])}):")
        for e in data["errors"]:
            print(f"  Tree #{e['num']}: {e['error']}")

    if not args.all and len(trees) == args.limit:
        print(f"\nNext batch: python place_tpz.py --offset {args.offset + args.limit}")


if __name__ == "__main__":
    main()
