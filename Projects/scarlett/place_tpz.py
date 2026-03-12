import sys, json, argparse
from _utils import check_excel_staleness, run_jsx, load_trees

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


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--offset", type=int, default=0)
    parser.add_argument("--limit",  type=int, default=50)
    parser.add_argument("--all",    action="store_true")
    args = parser.parse_args()

    check_excel_staleness()
    trees = load_trees()
    trees = trees[args.offset:]
    if not args.all:
        trees = trees[:args.limit]

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
