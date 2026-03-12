# Usage: python remove_tree.py 4
#        python remove_tree.py 4 7 12
import sys, json, tempfile, os

try:
    import win32com.client, pythoncom
except ImportError:
    print("ERROR: pywin32 not installed"); sys.exit(1)

def run_jsx(code):
    pythoncom.CoInitialize()
    ai = win32com.client.GetActiveObject("Illustrator.Application")
    with tempfile.NamedTemporaryFile(suffix='.jsx', delete=False, mode='w', encoding='utf-8') as f:
        f.write(code); fname = f.name
    try:
        return ai.DoJavaScriptFile(fname)
    finally:
        os.unlink(fname)

JSX = r"""
(function() {
    var nums = TREE_NUMS;
    var doc = app.activeDocument;
    var removed = [], missing = [];

    function tryRemove(layerName) {
        var lyr;
        try { lyr = doc.layers.getByName(layerName); } catch(e) { return; }
        lyr.locked = false;
        for (var i = 0; i < nums.length; i++) {
            var name = 'Tree ' + nums[i];
            try {
                lyr.groupItems.getByName(name).remove();
                removed.push(name + ' (' + layerName + ')');
            } catch(e) {
                missing.push(name + ' (' + layerName + ')');
            }
        }
    }

    tryRemove('TPZs');
    tryRemove('Labels');

    function toArr(arr) {
        var out = '[';
        for (var j = 0; j < arr.length; j++) out += (j ? ',' : '') + '"' + arr[j] + '"';
        return out + ']';
    }
    return '{"removed":' + toArr(removed) + ',"missing":' + toArr(missing) + '}';
})();
"""

nums = [int(a) for a in sys.argv[1:]]
if not nums:
    print("Usage: python remove_tree.py <tree_num> [tree_num ...]")
    sys.exit(1)

code = JSX.replace('TREE_NUMS', json.dumps(nums))
result = json.loads(run_jsx(code))
for r in result['removed']: print(f"  Removed: {r}")
for m in result['missing']:  print(f"  Not found: {m}")
