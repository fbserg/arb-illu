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

import csv

CSV_PATH  = r"C:\Projects\arborist-plans\Projects\7631-creditview\data.csv"
XLSX_PATH = r"C:\Projects\arborist-plans\Projects\7631-creditview\7631 Creditview Rd.xlsx"


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
            size_mm = tpz_circle_m * 4.0 if tpz_circle_m else None
            trees.append({
                'num':     row['tree_num'],
                'dir':     row['direction'],
                'cx':      cx,
                'cy':      cy,
                'size_mm': size_mm,
            })
    return trees


JSX_BODY = r"""
(function() {
    var doc;
    try { doc = app.activeDocument; } catch(e) { return '{"error":"No document open"}'; }

    var ab = doc.artboards[0].artboardRect; // [left, top, right, bottom]
    var abLeft  = ab[0], abTop    = ab[1];
    var abRight = ab[2], abBottom = ab[3];

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

    function circleBounds(item) {
        var gb = item.geometricBounds;
        var w  = Math.abs(gb[2] - gb[0]);
        var r  = w / 2;
        return {
            cx:   (gb[0] + gb[2]) / 2,
            cy:   (gb[1] + gb[3]) / 2,
            r:    r,
            diam: w
        };
    }

    function colorType(item) {
        var c = item.strokeColor;
        if (!c || c.typename !== 'CMYKColor') return "unknown";
        if (c.cyan > 50 && c.yellow > 50) return "green";
        if (c.magenta > 40 && c.yellow > 80) return "orange";
        return "unknown";
    }

    function isDashed(item) {
        return !!(item.strokeDashes && item.strokeDashes.length > 0);
    }

    function dist2(x1, y1, x2, y2) {
        var dx = x1 - x2, dy = y1 - y2;
        return Math.sqrt(dx * dx + dy * dy);
    }

    // Find TPZs layer
    var tpzLayer = null;
    for (var li = 0; li < doc.layers.length; li++) {
        if (doc.layers[li].name === 'TPZs') { tpzLayer = doc.layers[li]; break; }
    }
    if (!tpzLayer) return '{"error":"TPZs layer not found"}';

    // pathItems is recursive — includes items inside groups.
    // TPZ circles are unfilled; trunk dots are filled. Use !item.filled to exclude trunks.
    var allItems = tpzLayer.pathItems;
    var circles = [], lines = [];
    for (var i = 0; i < allItems.length; i++) {
        var item = allItems[i];
        if (isCircle(item) && !item.filled) {
            circles.push(item);
        } else if (item.pathPoints.length === 2 && !item.closed) {
            lines.push(item);
        }
    }

    // Pre-compute bounds for all circles
    var cBounds = [];
    for (var i = 0; i < circles.length; i++) cBounds.push(circleBounds(circles[i]));

    // Greedy nearest-circle within tolerance; usedSet is mutated
    function findNearest(x, y, tol, usedSet) {
        var bestDist = tol + 1, bestIdx = -1;
        for (var i = 0; i < cBounds.length; i++) {
            if (usedSet[i]) continue;
            var d = dist2(x, y, cBounds[i].cx, cBounds[i].cy);
            if (d < bestDist) { bestDist = d; bestIdx = i; }
        }
        return bestIdx;
    }

    var checks = [];

    // --- Check 1: count_match ---
    var countOk = circles.length === EXPECTED.length;
    checks.push({
        name: "count_match",
        status: countOk ? "PASS" : "FAIL",
        message: "Found " + circles.length + " circles, expected " + EXPECTED.length
    });

    // --- Check 2: bounds_check ---
    var oob = [];
    var minY = Math.min(abTop, abBottom), maxY = Math.max(abTop, abBottom);
    for (var i = 0; i < circles.length; i++) {
        var cb = cBounds[i];
        if (cb.cx < abLeft || cb.cx > abRight || cb.cy < minY || cb.cy > maxY) {
            oob.push("circle@(" + Math.round(cb.cx) + "," + Math.round(cb.cy) + ")");
        }
    }
    checks.push({
        name: "bounds_check",
        status: oob.length === 0 ? "PASS" : "FAIL",
        message: oob.length === 0
            ? "All " + circles.length + " circles within artboard"
            : oob.length + " circle(s) outside artboard: " + oob.join(", ")
    });

    // --- Check 3: style_audit ---
    var styleFailures = [], styleUsed = {};
    for (var i = 0; i < EXPECTED.length; i++) {
        var t   = EXPECTED[i];
        var idx = findNearest(t.cx, t.cy, 50, styleUsed);
        if (idx === -1) continue;
        styleUsed[idx] = true;
        var ct   = colorType(circles[idx]);
        var dash = isDashed(circles[idx]);
        var dir  = t.dir;
        var ok   = false, expectedStyle = "";
        var actualStyle = ct + (dash ? "-dashed" : "-solid");
        if      (dir === "Protect") { expectedStyle = "green-solid";  ok = ct === "green"  && !dash; }
        else if (dir === "Injury")  { expectedStyle = "orange-dashed"; ok = ct === "orange" && dash; }
        else if (dir === "Remove")  { expectedStyle = "orange-solid"; ok = ct === "orange" && !dash; }
        if (!ok && expectedStyle) {
            styleFailures.push("Tree " + t.num + ": expected " + expectedStyle + ", got " + actualStyle);
        }
    }
    checks.push({
        name: "style_audit",
        status: styleFailures.length === 0 ? "PASS" : "FAIL",
        message: styleFailures.length === 0 ? "All styles match" : styleFailures.join("; ")
    });

    // --- Check 4: removal_x_check ---
    var xFailures = [];
    for (var i = 0; i < circles.length; i++) {
        if (colorType(circles[i]) !== "orange" || isDashed(circles[i])) continue;
        var cb  = cBounds[i];
        var tol = cb.r * 1.5;
        var matched = 0;
        for (var j = 0; j < lines.length; j++) {
            var pp = lines[j].pathPoints;
            var p0 = pp[0].anchor, p1 = pp[1].anchor;
            if (dist2(p0[0], p0[1], cb.cx, cb.cy) <= tol &&
                dist2(p1[0], p1[1], cb.cx, cb.cy) <= tol) {
                matched++;
            }
        }
        if (matched < 2) {
            xFailures.push("Circle@(" + Math.round(cb.cx) + "," + Math.round(cb.cy) + "): " + matched + " X-line(s) (need 2)");
        }
    }
    checks.push({
        name: "removal_x_check",
        status: xFailures.length === 0 ? "PASS" : "FAIL",
        message: xFailures.length === 0 ? "All removal circles have X-lines" : xFailures.join("; ")
    });

    // --- Check 5: size_sanity ---
    var sizeFailures = [], sizeUsed = {};
    for (var i = 0; i < EXPECTED.length; i++) {
        var t = EXPECTED[i];
        if (t.size_mm === null) continue;
        var idx = findNearest(t.cx, t.cy, 50, sizeUsed);
        if (idx === -1) continue;
        sizeUsed[idx] = true;
        var expectedDiam = t.size_mm * (72 / 25.4);
        var actualDiam   = cBounds[idx].diam;
        var diff = Math.abs(actualDiam - expectedDiam);
        if (diff > 2) {
            sizeFailures.push("Tree " + t.num + ": expected " + expectedDiam.toFixed(1) + "pt, got " + actualDiam.toFixed(1) + "pt (diff " + diff.toFixed(1) + "pt)");
        }
    }
    checks.push({
        name: "size_sanity",
        status: sizeFailures.length === 0 ? "PASS" : "FAIL",
        message: sizeFailures.length === 0 ? "All sizes within tolerance" : sizeFailures.join("; ")
    });

    // --- Check 6: orphan_missing ---
    var pairs = [];
    for (var ei = 0; ei < EXPECTED.length; ei++) {
        var t = EXPECTED[ei];
        for (var ci = 0; ci < cBounds.length; ci++) {
            var d = dist2(t.cx, t.cy, cBounds[ci].cx, cBounds[ci].cy);
            if (d <= 50) pairs.push({ei: ei, ci: ci, d: d});
        }
    }
    pairs.sort(function(a, b) { return a.d - b.d; });
    var matchedExp = {}, matchedCirc = {};
    for (var k = 0; k < pairs.length; k++) {
        var p = pairs[k];
        if (!matchedExp[p.ei] && !matchedCirc[p.ci]) {
            matchedExp[p.ei] = true;
            matchedCirc[p.ci] = true;
        }
    }
    var orphans = [], missing = [];
    for (var ci = 0; ci < circles.length; ci++) {
        if (!matchedCirc[ci]) orphans.push("circle@(" + Math.round(cBounds[ci].cx) + "," + Math.round(cBounds[ci].cy) + ")");
    }
    for (var ei = 0; ei < EXPECTED.length; ei++) {
        if (!matchedExp[ei]) missing.push("Tree " + EXPECTED[ei].num);
    }
    var omMsgs = [];
    if (orphans.length > 0) omMsgs.push("Orphan circles: " + orphans.join(", "));
    if (missing.length > 0) omMsgs.push("Missing trees: " + missing.join(", "));
    checks.push({
        name: "orphan_missing",
        status: (orphans.length + missing.length) === 0 ? "PASS" : "FAIL",
        message: omMsgs.length === 0 ? "All circles matched to expected trees" : omMsgs.join("; ")
    });

    // --- summary ---
    var nPass = 0, nFail = 0, nWarn = 0;
    for (var i = 0; i < checks.length; i++) {
        if      (checks[i].status === "PASS") nPass++;
        else if (checks[i].status === "FAIL") nFail++;
        else if (checks[i].status === "WARN") nWarn++;
    }
    var chJson = '[';
    for (var ci = 0; ci < checks.length; ci++) {
        if (ci > 0) chJson += ',';
        var ch = checks[ci];
        chJson += '{"name":"' + ch.name + '","status":"' + ch.status + '","message":"' + ch.message.replace(/\\/g,'\\\\').replace(/"/g,'\\"') + '"}';
    }
    chJson += ']';
    return '{"checks":' + chJson + ',"summary":{"pass":' + nPass + ',"fail":' + nFail + ',"warn":' + nWarn + '}}';
})();
"""


def build_jsx(trees):
    return "var EXPECTED = " + json.dumps(trees) + ";\n" + JSX_BODY


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


def print_report(results):
    print("TPZ Review — 7631 Creditview Rd")
    print("=" * 56)
    for chk in results["checks"]:
        print(f"{chk['status']:<4}  {chk['name']:<20}  {chk['message']}")
    print("=" * 56)
    s = results["summary"]
    print(f"Result: {s['pass']} PASS, {s['fail']} FAIL, {s['warn']} WARN")
    return s["fail"] > 0


def main():
    print(f"Reading CSV: {CSV_PATH}")
    trees = read_csv(CSV_PATH)
    print(f"  {len(trees)} trees loaded")

    jsx = build_jsx(trees)
    print("Running review script in Illustrator...")
    raw = run_jsx(jsx)
    results = json.loads(raw)

    if "error" in results:
        print(f"ERROR: {results['error']}")
        sys.exit(1)

    has_failures = print_report(results)
    sys.exit(1 if has_failures else 0)


if __name__ == "__main__":
    main()
