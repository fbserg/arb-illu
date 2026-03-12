import csv
import os
import sys
import tempfile
import time

import pythoncom
import win32com.client

JSX = r"""
(function() {
    var doc = app.activeDocument;
    var lyr = doc.layers.getByName('Layer 1');
    var lines = [];

    // Pass 1: collect ALL labels (T-XX and #NNN) into one pool
    var allLabels = [];
    for (var i = 0; i < lyr.textFrames.length; i++) {
        var tf = lyr.textFrames[i];
        var s = tf.contents;
        var gb = tf.geometricBounds; // [left, top, right, bottom]
        var tcx = (gb[0] + gb[2]) / 2;
        var tcy = (gb[1] + gb[3]) / 2;
        if (/^T-\d+$/.test(s)) {
            allLabels.push({ltype: 'T', num: parseInt(s.substring(2), 10), cx: tcx, cy: tcy});
        } else if (/^#\d+$/.test(s)) {
            allLabels.push({ltype: 'H', num: parseInt(s.substring(1), 10), cx: tcx, cy: tcy});
        }
    }

    // Pass 2: build circles array from stroke-only square-ish pathItems
    var circles = [];
    for (var j = 0; j < lyr.pathItems.length; j++) {
        var pi = lyr.pathItems[j];

        if (pi.filled) continue;
        if (!pi.stroked) continue;

        var gb2 = pi.geometricBounds; // [left, top, right, bottom]
        var w = gb2[2] - gb2[0];
        var h = gb2[1] - gb2[3]; // top > bottom in AI coords
        if (w < 5 || w > 100) continue;
        if (Math.abs(w - h) > 2) continue;

        var cx = (gb2[0] + gb2[2]) / 2;
        var cy = (gb2[1] + gb2[3]) / 2;

        var sc = pi.strokeColor;
        var colName = 'other';
        if (sc.typename === 'RGBColor') {
            var r = sc.red; var g = sc.green; var b = sc.blue;
            if (r > 200 && g > 100 && g < 180 && b < 30) {
                colName = 'orange';
            } else if (r < 50 && g > 80 && b > 150) {
                colName = 'blue';
            }
        }

        circles.push({cx: cx, cy: cy, color: colName});
    }

    // Pass 3: for each label find its nearest circle — 1 label : 1 circle
    for (var li = 0; li < allLabels.length; li++) {
        var lbl = allLabels[li];
        var bestIdx = -1;
        var bestD2 = -1;
        for (var ci = 0; ci < circles.length; ci++) {
            var dx = circles[ci].cx - lbl.cx;
            var dy = circles[ci].cy - lbl.cy;
            var d2 = dx*dx + dy*dy;
            if (bestIdx === -1 || d2 < bestD2) { bestD2 = d2; bestIdx = ci; }
        }
        if (bestIdx === -1) continue;
        var c = circles[bestIdx];
        var tNum = (lbl.ltype === 'T') ? lbl.num : '';
        var hNum = (lbl.ltype === 'H') ? lbl.num : '';
        lines.push('' + tNum + '|' + hNum + '|' + c.cx + '|' + c.cy + '|' + c.color);
    }

    return lines.join('\n');
})();
"""


def get_illustrator(max_attempts=3, delay=0.3):
    last_err = RuntimeError("Illustrator not reachable")
    for attempt in range(max_attempts):
        try:
            pythoncom.CoInitialize()
            return win32com.client.GetActiveObject("Illustrator.Application")
        except Exception as e:
            last_err = e
            if attempt < max_attempts - 1:
                time.sleep(delay)
    raise last_err


def main():
    app = get_illustrator()
    with tempfile.NamedTemporaryFile(mode="w", suffix=".jsx", delete=False, encoding="utf-8") as f:
        f.write(JSX)
        jsx_path = f.name
    try:
        result = app.DoJavaScriptFile(jsx_path)
    finally:
        os.unlink(jsx_path)

    rows = []
    label_counts = {}  # label_key -> list of row indices (to detect duplicate claims)

    for line in result.strip().splitlines():
        line = line.strip()
        if not line:
            continue
        parts = line.split("|")
        if len(parts) != 5:
            print(f"WARN: unexpected line: {line!r}", file=sys.stderr)
            continue
        t_num, hash_num, cx, cy, color = parts
        label_key = ("T" + t_num) if t_num else ("H" + hash_num)
        rows.append({
            "t_num": t_num,
            "hash_num": hash_num,
            "cx": round(float(cx), 2),
            "cy": round(float(cy), 2),
            "color": color,
        })
        label_counts.setdefault(label_key, []).append(len(rows) - 1)

    duplicates = {k: v for k, v in label_counts.items() if len(v) > 1}
    no_label = [r for r in rows if not r["t_num"] and not r["hash_num"]]

    out_path = "Projects/scarlett/coords.csv"
    with open(out_path, "w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["t_num", "hash_num", "cx", "cy", "color"])
        writer.writeheader()
        writer.writerows(rows)

    t_count = sum(1 for r in rows if r["t_num"])
    h_count = sum(1 for r in rows if r["hash_num"])
    print(f"Total circles  : {len(rows)}")
    print(f"T-label        : {t_count}")
    print(f"#-label        : {h_count}")
    print(f"No label       : {len(no_label)}")
    print(f"Duplicate claims: {len(duplicates)}")
    if duplicates:
        for k, idxs in list(duplicates.items())[:10]:
            print(f"  {k}: {[rows[i] for i in idxs]}")
    print(f"Written        : {out_path}")


if __name__ == "__main__":
    main()
