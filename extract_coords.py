import sys
import os
import json
import tempfile

try:
    import win32com.client
    import pythoncom
except ImportError:
    print("ERROR: pywin32 not installed. Run: pip install pywin32")
    sys.exit(1)

EXCEL_PATH = r"C:\Projects\arborist-plans\Projects\7631-creditview\7631 Creditview Rd.xlsx"

JSX = r"""
(function() {
    var doc;
    try { doc = app.activeDocument; } catch(e) { return JSON.stringify({error: "No document open"}); }

    var dimLayer = null;
    for (var li = 0; li < doc.layers.length; li++) {
        if (doc.layers[li].name === 'Dimensions') { dimLayer = doc.layers[li]; break; }
    }
    if (!dimLayer) return JSON.stringify({error: "Dimensions layer not found"});

    function dist(ax, ay, bx, by) {
        var dx = ax - bx, dy = ay - by;
        return Math.sqrt(dx*dx + dy*dy);
    }

    // Collect labels: {num, cx, cy}
    var labels = [];
    var tfs = dimLayer.textFrames;
    for (var i = 0; i < tfs.length; i++) {
        var contents = tfs[i].contents;
        if (!contents.match(/^#\d+$/)) continue;
        var gb = tfs[i].geometricBounds; // [left, top, right, bottom]
        labels.push({
            num: String(parseInt(contents.slice(1), 10)),
            cx:  (gb[0] + gb[2]) / 2,
            cy:  (gb[1] + gb[3]) / 2
        });
    }

    // Collect open leader lines with 3 or 4 points
    var coords = [];
    var paths = dimLayer.pathItems;
    for (var i = 0; i < paths.length; i++) {
        var path = paths[i];
        if (path.closed) continue;
        var nPts = path.pathPoints.length;
        if (nPts < 3 || nPts > 4) continue;

        var ep0 = path.pathPoints[0].anchor; // [x, y]
        var ep2 = path.pathPoints[nPts - 1].anchor; // last point

        // Find nearest label to each outer endpoint
        var best0 = null, best0d = 80;
        var best2 = null, best2d = 80;
        for (var j = 0; j < labels.length; j++) {
            var d0 = dist(ep0[0], ep0[1], labels[j].cx, labels[j].cy);
            var d2 = dist(ep2[0], ep2[1], labels[j].cx, labels[j].cy);
            if (d0 < best0d) { best0d = d0; best0 = labels[j]; }
            if (d2 < best2d) { best2d = d2; best2 = labels[j]; }
        }

        // Skip if both or neither endpoints match a label
        if (best0 && best2) continue; // dimension line between two labels
        if (!best0 && !best2) continue; // not a tree leader

        if (best0) {
            // ep0 is near the label → ep2 is the circle centre
            coords.push({num: best0.num, cx: ep2[0], cy: ep2[1]});
        } else {
            // ep2 is near the label → ep0 is the circle centre
            coords.push({num: best2.num, cx: ep0[0], cy: ep0[1]});
        }
    }

    var parts = [];
    for (var i = 0; i < coords.length; i++) {
        parts.push('{"num":"' + coords[i].num + '","cx":' + coords[i].cx + ',"cy":' + coords[i].cy + '}');
    }
    return '{"coords":[' + parts.join(',') + '],"label_count":' + labels.length + '}';
})();
"""


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
    print("Running ExtendScript to extract leader-line endpoints...")
    raw = run_jsx(JSX)
    data = json.loads(raw)

    if "error" in data:
        print(f"ERROR: {data['error']}")
        sys.exit(1)

    coords = data["coords"]
    print(f"  Found {len(coords)} leader lines matched to labels ({data['label_count']} labels total)")

    # Build lookup: tree_num_str → (cx, cy)
    found = {}
    for c in coords:
        found[c["num"]] = (c["cx"], c["cy"])

    # Open Excel and fill missing O/P
    pythoncom.CoInitialize()
    xl = win32com.client.Dispatch("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    wb = xl.Workbooks.Open(os.path.abspath(EXCEL_PATH))
    ws = wb.Sheets("Sheet1")

    filled = []
    skipped = []
    not_found = []

    row = 3
    while True:
        tree_num_raw = ws.Cells(row, 1).Value  # col A
        if tree_num_raw is None:
            break
        tree_num = str(int(tree_num_raw)) if isinstance(tree_num_raw, (int, float)) else str(tree_num_raw).strip()

        has_cx = isinstance(ws.Cells(row, 15).Value, (int, float))  # col O
        if has_cx:
            skipped.append(tree_num)
        elif tree_num in found:
            cx, cy = found[tree_num]
            ws.Cells(row, 15).Value = cx
            ws.Cells(row, 16).Value = cy
            filled.append(tree_num)
        else:
            not_found.append(tree_num)

        row += 1

    if filled:
        wb.Save()
    wb.Close(False)

    print(f"\nFilled {len(filled)} trees: {', '.join(filled)}" if filled else "\nNo new coords written.")
    if not_found:
        print(f"Not found in Illustrator: {', '.join(not_found)}  (check leader lines)")
    if skipped:
        print(f"Already had coords (skipped): {', '.join(skipped)}")


if __name__ == "__main__":
    main()
