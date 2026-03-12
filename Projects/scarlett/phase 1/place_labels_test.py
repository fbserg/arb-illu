"""Quick test: place a text label at every coord in coords.csv — no circles."""
import csv, json, os, sys, tempfile

try:
    import win32com.client, pythoncom
except ImportError:
    print("ERROR: pywin32 not installed."); sys.exit(1)

COORDS_PATH = r"C:\Projects\arborist-plans\Projects\scarlett\coords.csv"

JSX = r"""
(function() {
    var doc;
    try { doc = app.activeDocument; } catch(e) { return '{"error":"No document open"}'; }

    function findOrCreateLayer(parent, name) {
        var coll = parent.layers;
        for (var i = 0; i < coll.length; i++) if (coll[i].name === name) return coll[i];
        var l = parent.layers.add(); l.name = name; return l;
    }

    var lyr = findOrCreateLayer(doc, 'Labels Test');
    lyr.visible = true; lyr.locked = false;
    doc.activeLayer = lyr;

    // Clear layer
    while (lyr.textFrames.length > 0) lyr.textFrames[0].remove();
    while (lyr.pathItems.length  > 0) lyr.pathItems[0].remove();

    var placed = 0;
    for (var i = 0; i < TREES.length; i++) {
        var t = TREES[i];
        var tf = lyr.textFrames.add();
        tf.contents = t.label;
        tf.top  = t.cy;
        tf.left = t.cx;
        tf.textRange.characterAttributes.size = 5;
        placed++;
    }
    return '{"placed":' + placed + '}';
})();
"""


def main():
    trees = []
    with open(COORDS_PATH, newline='', encoding='utf-8') as f:
        for row in csv.DictReader(f):
            t = row['t_num'].strip()
            h = row['hash_num'].strip()
            label = ('T' + t) if t else ('#' + h)
            trees.append({'label': label, 'cx': float(row['cx']), 'cy': float(row['cy'])})

    print(f"Placing {len(trees)} labels...")
    jsx = "var TREES = " + json.dumps(trees) + ";\n" + JSX

    pythoncom.CoInitialize()
    tmp = tempfile.NamedTemporaryFile(suffix=".jsx", delete=False, mode="w", encoding="utf-8")
    tmp.write(jsx); path = tmp.name; tmp.close()
    try:
        ai = win32com.client.GetActiveObject("Illustrator.Application")
        raw = ai.DoJavaScriptFile(path)
    finally:
        os.unlink(path)

    import json as j
    data = j.loads(raw)
    if "error" in data:
        print(f"ERROR: {data['error']}"); sys.exit(1)
    print(f"Done: {data['placed']} labels on 'Labels Test' layer")


if __name__ == "__main__":
    main()
