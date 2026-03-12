"""Re-place circles + labels for trees missing from the Labels layer.

Workflow:
  1. Edit direction in Excel, run export_data.py
  2. Delete stale labels from Illustrator
  3. python Projects/scarlett/repair.py
"""
import sys, os, csv, json, tempfile

try:
    import win32com.client, pythoncom
except ImportError:
    print("ERROR: pywin32 not installed."); sys.exit(1)

DATA_PATH  = r"C:\Projects\arborist-plans\Projects\scarlett\data.csv"
EXCEL_PATH = r"C:\Projects\arborist-plans\Projects\scarlett\Excel Master Sheet.xlsx"
PLAN_W, PLAN_H = 2384, 3370


def transform(cx, cy):
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


QUERY_JSX = r"""
(function() {
    var doc;
    try { doc = app.activeDocument; } catch(e) { return '{"error":"No document open"}'; }
    var labelsLayer = null;
    for (var i = 0; i < doc.layers.length; i++) {
        if (doc.layers[i].name === 'Labels') { labelsLayer = doc.layers[i]; break; }
    }
    if (!labelsLayer) return '{"groups":[]}';
    var names = '[';
    for (var j = 0; j < labelsLayer.groupItems.length; j++) {
        if (j > 0) names += ',';
        names += '"' + labelsLayer.groupItems[j].name + '"';
    }
    names += ']';
    return '{"groups":' + names + '}';
})();
"""

PLACE_JSX = r"""
(function() {
    var doc;
    try { doc = app.activeDocument; } catch(e) { return '{"error":"No document open"}'; }

    var PT_PER_MM = 72.0 / 25.4;
    var fontName = "Arial-BoldMT", fontSize = 5;
    var S45 = Math.SQRT2 / 2;

    function cmyk(c, m, y, k) {
        var col = new CMYKColor();
        col.cyan = c; col.magenta = m; col.yellow = y; col.black = k;
        return col;
    }
    var GREEN  = cmyk(70, 0,  100, 0);
    var ORANGE = cmyk(0,  62, 100, 0);
    var BLACK  = cmyk(0,  0,  0,   100);
    var WHITE  = cmyk(0,  0,  0,   0);

    function getOrCreateLayer(name) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === name) return doc.layers[i];
        }
        var l = doc.layers.add(); l.name = name; return l;
    }

    var tpzLayer    = getOrCreateLayer('TPZs');
    var labelsLayer = getOrCreateLayer('Labels');
    tpzLayer.visible    = true; tpzLayer.locked    = false;
    labelsLayer.visible = true; labelsLayer.locked = false;

    var placed_tpz = 0, placed_lbl = 0, errors = [];

    for (var ti = 0; ti < TREES.length; ti++) {
        var tree = TREES[ti];
        var cx = tree.cx, cy = tree.cy;
        var dir = tree.dir;
        var groupName = 'Tree ' + tree.num;

        // Remove stale groups from both layers (silent if absent)
        try { tpzLayer.groupItems.getByName(groupName).remove(); }    catch(e) {}
        try { labelsLayer.groupItems.getByName(groupName).remove(); } catch(e) {}

        var color = (dir === 'Protect' || dir === 'Retain') ? GREEN : ORANGE;

        // ── TPZ circle ──────────────────────────────────────────────────────
        if (tree.tpz_mm && tree.tpz_mm > 0) {
            try {
                doc.activeLayer = tpzLayer;
                var tpzDiam = tree.tpz_mm * PT_PER_MM;
                var tpzRad  = tpzDiam / 2;

                var grp = tpzLayer.groupItems.add();
                grp.name = groupName;

                var circle = tpzLayer.pathItems.ellipse(cy + tpzRad, cx - tpzRad, tpzDiam, tpzDiam);
                circle.filled = false; circle.stroked = true;
                circle.strokeColor = color; circle.strokeWidth = 0.84;
                circle.strokeDashes = (dir === 'Injury') ? [5] : [];
                circle.move(grp, ElementPlacement.PLACEATBEGINNING);

                if (dir === 'Remove') {
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
                placed_tpz++;
            } catch(e) {
                errors.push('TPZ ' + tree.num + ': ' + e.toString());
            }
        }

        // ── Label ────────────────────────────────────────────────────────────
        try {
            doc.activeLayer = labelsLayer;

            var isProtect = (dir === 'Protect' || dir === 'Retain');
            var abbr = isProtect ? 'Pro' : (dir === 'Injury' ? 'Inj' : 'Rem');
            var dotColor = isProtect ? GREEN : ORANGE;

            var contents;
            if (tree.tpz_m === null) {
                contents = '#' + tree.num + ' ' + abbr;
            } else {
                var tpzStr = (Math.round(tree.tpz_m * 10) / 10).toFixed(1);
                contents = '#' + tree.num + ' ' + abbr + ' ' + tpzStr + 'm';
            }

            // Measure at origin to get bb offsets
            var mtf = labelsLayer.textFrames.add();
            mtf.contents = contents;
            var mca = mtf.textRange.characterAttributes;
            try { mca.textFont = app.textFonts.getByName(fontName); } catch(e2) {}
            mca.size = fontSize;
            mca.fillColor = BLACK;
            mtf.textRange.paragraphAttributes.justification = Justification.LEFT;
            mtf.position = [0, 0];
            var gbM = mtf.geometricBounds;
            mtf.remove();

            var cxOff = (gbM[0] + gbM[2]) / 2;
            var cyOff = (gbM[1] + gbM[3]) / 2;
            var anchorX = cx - (1 + 0.5) * S45;
            var anchorY = cy - (1 + 0.5) * S45;
            var px = anchorX - cxOff * (1 - S45) - cyOff * S45;
            var py = anchorY + cxOff * S45       - cyOff * (1 - S45);

            var tf = labelsLayer.textFrames.add();
            tf.contents = contents;
            var ca = tf.textRange.characterAttributes;
            try { ca.textFont = app.textFonts.getByName(fontName); } catch(e2) {}
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
            dot.filled = true; dot.fillColor = dotColor; dot.stroked = false;

            var lgrp = labelsLayer.groupItems.add();
            lgrp.name = groupName;
            dot.moveToEnd(lgrp);
            tf.moveToEnd(lgrp);
            bg.moveToEnd(lgrp);

            placed_lbl++;
        } catch(e) {
            errors.push('Label ' + tree.num + ': ' + e.toString());
        }
    }

    var er = '[';
    for (var ei = 0; ei < errors.length; ei++) {
        if (ei) er += ',';
        er += '"' + errors[ei].replace(/\\/g,'\\\\').replace(/"/g,'\\"') + '"';
    }
    er += ']';
    return '{"tpz_placed":' + placed_tpz + ',"labels_placed":' + placed_lbl + ',"errors":' + er + '}';
})();
"""


def main():
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--trees", help="Comma-separated tree nums to force-repair (bypasses missing detection)")
    args = parser.parse_args()

    if os.path.exists(EXCEL_PATH) and os.path.getmtime(EXCEL_PATH) > os.path.getmtime(DATA_PATH):
        print("WARNING: Excel is newer than data.csv — run: python Projects/scarlett/export_data.py")

    # Load data.csv
    all_trees = []
    with open(DATA_PATH, newline='', encoding='utf-8') as f:
        for row in csv.DictReader(f):
            cx = float(row['cx']) if row['cx'] else None
            cy = float(row['cy']) if row['cy'] else None
            if cx is None or cy is None:
                continue
            cx, cy = transform(cx, cy)
            all_trees.append({
                'num':    row['tree_num'],
                'dir':    normalize_dir(row['direction']),
                'tpz_m':  float(row['tpz_m'])  if row['tpz_m']  else None,
                'tpz_mm': float(row['tpz_mm']) if row['tpz_mm'] else None,
                'cx':     cx,
                'cy':     cy,
            })

    if args.trees:
        # Force-repair specific trees regardless of what's in Illustrator
        force_set = set(t.strip() for t in args.trees.split(","))
        to_repair = [t for t in all_trees if t['num'] in force_set]
    else:
        # Step 1: find which labels exist in Illustrator
        print("Querying Labels layer...")
        raw = run_jsx(QUERY_JSX)
        qdata = json.loads(raw)
        if "error" in qdata:
            print(f"ERROR: {qdata['error']}"); sys.exit(1)

        existing = set()
        for name in qdata.get("groups", []):
            if name.startswith("Tree "):
                existing.add(name[5:])

        to_repair = [t for t in all_trees if t['num'] not in existing]

    if not to_repair:
        print("Nothing to repair — all labels present."); return

    print(f"Repairing {len(to_repair)} trees: {', '.join(t['num'] for t in to_repair)}")

    # Step 3: place circles + labels
    jsx = "var TREES = " + json.dumps(to_repair) + ";\n" + PLACE_JSX
    raw = run_jsx(jsx)
    result = json.loads(raw)

    if "error" in result:
        print(f"ERROR: {result['error']}"); sys.exit(1)

    print(f"Placed: {result['tpz_placed']} circles, {result['labels_placed']} labels.")
    if result.get("errors"):
        for e in result["errors"]:
            print(f"  ERROR: {e}")


if __name__ == "__main__":
    main()
