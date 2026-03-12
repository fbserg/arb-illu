import sys, os, csv, json, argparse, tempfile

try:
    import win32com.client, pythoncom
except ImportError:
    print("ERROR: pywin32 not installed."); sys.exit(1)

DATA_PATH   = r"C:\Projects\arborist-plans\Projects\scarlett\data.csv"
EXCEL_PATH  = r"C:\Projects\arborist-plans\Projects\scarlett\Excel Master Sheet.xlsx"

PLAN_W, PLAN_H = 2384, 3370

def transform(cx, cy):
    """PLAN.ai coords → template coords (90° CCW rotation around artboard centre)."""
    return (PLAN_W + PLAN_H) / 2 - cy, cx + (PLAN_H - PLAN_W) / 2

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
        for (var i = 0; i < coll.length; i++) if (coll[i].name === name) return coll[i];
        var l = parent.layers.add(); l.name = name; return l;
    }

    var tpzLayer = findOrCreateLayer(doc, 'TPZs');
    tpzLayer.visible = true; tpzLayer.locked = false;
    tpzLayer.zOrder(ZOrderMethod.SENDTOBACK);
    doc.activeLayer = tpzLayer;

    if (CLEAR_FIRST) {
        for (var di = tpzLayer.layers.length - 1; di >= 0; di--) tpzLayer.layers[di].remove();
        while (tpzLayer.groupItems.length > 0) tpzLayer.groupItems[0].remove();
        while (tpzLayer.pathItems.length  > 0) tpzLayer.pathItems[0].remove();
    }

    var tpzPlaced = 0;
    var skipped = [];
    var errors = [];

    for (var ti = 0; ti < TREES.length; ti++) {
        var tree = TREES[ti];
        var cx = tree.cx, cy = tree.cy;
        var tpzMm = tree.tpz_mm;
        var dir   = tree.dir;

        if (!tpzMm || tpzMm <= 0) { skipped.push({num: tree.num, reason: 'no TPZ'}); continue; }

        var tpzDiam = tpzMm * PT_PER_MM;
        var tpzRad  = tpzDiam / 2;
        var color = (dir === 'Protect' || dir === 'Retain') ? GREEN : ORANGE;

        try {
            var grp = tpzLayer.groupItems.add();
            grp.name = 'Tree ' + tree.num;

            var circle = tpzLayer.pathItems.ellipse(cy + tpzRad, cx - tpzRad, tpzDiam, tpzDiam);
            circle.filled = false; circle.stroked = true;
            circle.strokeColor = color; circle.strokeWidth = 0.84;
            circle.strokeDashes = (dir === 'Injury') ? [5] : [];
            circle.move(grp, ElementPlacement.PLACEATBEGINNING);

            if (dir === 'Remove' || dir === 'Removal') {
                var r = tpzRad * 0.707;
                var l1 = tpzLayer.pathItems.add();
                l1.setEntirePath([[cx - r, cy - r], [cx + r, cy + r]]);
                l1.filled = false; l1.stroked = true;
                l1.strokeColor = ORANGE; l1.strokeWidth = 0.84;
                l1.move(grp, ElementPlacement.PLACEATBEGINNING);
                var l2 = tpzLayer.pathItems.add();
                l2.setEntirePath([[cx - r, cy + r], [cx + r, cy - r]]);
                l2.filled = false; l2.stroked = true;
                l2.strokeColor = ORANGE; l2.strokeWidth = 0.84;
                l2.move(grp, ElementPlacement.PLACEATBEGINNING);
            }

            tpzPlaced++;
        } catch(e) {
            errors.push({num: tree.num, error: e.toString()});
        }
    }

    var sk = '[';
    for (var si = 0; si < skipped.length; si++) {
        if (si) sk += ',';
        sk += '{"num":"' + skipped[si].num + '","reason":"' + skipped[si].reason + '"}';
    }
    sk += ']';
    var er = '[';
    for (var ei = 0; ei < errors.length; ei++) {
        if (ei) er += ',';
        er += '{"num":"' + errors[ei].num + '","error":"' + errors[ei].error.replace(/"/g,'\\"') + '"}';
    }
    er += ']';
    return '{"tpz_placed":' + tpzPlaced + ',"skipped":' + sk + ',"errors":' + er + '}';
})();
"""


def load_data(limit, offset):
    if os.path.exists(EXCEL_PATH) and os.path.getmtime(EXCEL_PATH) > os.path.getmtime(DATA_PATH):
        print("WARNING: Excel is newer than data.csv — run: python Projects/scarlett/export_data.py")

    trees = []
    with open(DATA_PATH, newline='', encoding='utf-8') as f:
        for row in csv.DictReader(f):
            cx  = float(row['cx'])  if row['cx']  else None
            cy  = float(row['cy'])  if row['cy']  else None
            if cx is None or cy is None:
                continue
            tpz_mm = float(row['tpz_mm']) if row['tpz_mm'] else None
            cx, cy = transform(cx, cy)
            trees.append({
                'num':    row['tree_num'],
                'cx':     cx,
                'cy':     cy,
                'dir':    row['direction'],
                'tpz_mm': tpz_mm,
            })

    trees = trees[offset:]
    if limit is not None:
        trees = trees[:limit]
    return trees


def run_jsx(jsx_code):
    pythoncom.CoInitialize()
    tmp = tempfile.NamedTemporaryFile(suffix=".jsx", delete=False, mode="w", encoding="utf-8")
    tmp.write(jsx_code); tmp_path = tmp.name; tmp.close()
    try:
        ai = win32com.client.GetActiveObject("Illustrator.Application")
        result = ai.DoJavaScriptFile(tmp_path)
    finally:
        os.unlink(tmp_path)
    result = str(result) if result is not None else ""
    if not result.strip().startswith("{"):
        raise RuntimeError("JS returned: " + result)
    return result


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--offset", type=int, default=0)
    parser.add_argument("--limit",  type=int, default=50)
    parser.add_argument("--all",    action="store_true")
    args = parser.parse_args()

    limit = None if args.all else args.limit
    trees = load_data(limit, args.offset)
    clear_first = (args.offset == 0)
    print(f"Placing {len(trees)} trees (offset={args.offset}, {'clearing layer' if clear_first else 'appending'})")

    jsx = (
        "var TREES = " + json.dumps(trees) + ";\n"
        "var CLEAR_FIRST = " + ("true" if clear_first else "false") + ";\n"
        + JSX_BODY
    )

    print("Sending to Illustrator...")
    raw = run_jsx(jsx)
    data = json.loads(raw)

    if "error" in data:
        print(f"ERROR: {data['error']}"); sys.exit(1)

    print(f"Placed:  {data['tpz_placed']}")
    if data.get("skipped"):
        print(f"Skipped: {len(data['skipped'])}")
    if data.get("errors"):
        print(f"Errors:  {len(data['errors'])}")
        for e in data["errors"][:5]:
            print(f"  {e['num']}: {e['error']}")

    if not args.all and len(trees) == args.limit:
        print(f"\nNext batch: python Projects/scarlett/place_tpz.py --offset {args.offset + args.limit}")


if __name__ == "__main__":
    main()
