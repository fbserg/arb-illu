# New Plans — Inserting a PDF Site Plan into an Illustrator File

Procedure for placing a civil engineering site plan PDF as a background reference layer
in an arborist Illustrator file. Run this at the start of each new project.

---

## Step 1: Gather file info

### PDF page size (MediaBox)
```python
with open('Site Plan.pdf', 'rb') as f:
    data = f.read(4096)
import re
print(re.findall(r'MediaBox\s*\[([^\]]+)\]', data.decode('latin-1')))
# e.g. ['0 0 2592 1728'] → 2592 pts wide × 1728 pts tall = 36"×24"
```

### PDF drawing scale
```bash
pdftotext "Site Plan.pdf" - | head -100
# Scan output for "1:200", "1:500", etc. in the title block area
```

### .ai artboard size
```
mcp__illustrator__query  →  "artboard": [x1, y1, x2, y2]
width  = x2 − x1
height = y1 − y2  (y-axis is inverted in Illustrator)
```

### .ai drawing scale
Ask user — it will be stated in the title block or known from project setup.

### PDF creation date (for layer name)
```python
with open('Site Plan.pdf', 'rb') as f:
    data = f.read(8192)
import re
print(re.findall(r'CreateDate.*?(\d{4}-\d{2}-\d{2})', data.decode('latin-1')))
# or look for /CreationDate(D:YYYYMMDD...) in the first 4KB
```

**Rule: Always confirm the scale of BOTH drawings before placing. Never assume they match.**

---

## Step 2: Calculate scale factor

```
scale_factor = pdf_drawing_scale / ai_drawing_scale
```

| PDF scale | .ai scale | scale_factor | Action |
|-----------|-----------|--------------|--------|
| 1:200     | 1:200     | 1.0          | No resize |
| 1:500     | 1:200     | 2.5          | Enlarge PDF ×2.5 |
| 1:200     | 1:500     | 0.4          | Shrink PDF ×0.4 |

---

## Step 3: ExtendScript via MCP `run`

```javascript
var doc = app.activeDocument;

// Suppress all dialogs (font warnings etc.) — MUST be first
// Missing-font dialogs block the COM connection and crash the MCP server
app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

// Save lock states, then unlock — restore at end (all in one run() call, no extra latency)
var lockStates = [];
for (var i = 0; i < doc.layers.length; i++) {
    lockStates.push({ name: doc.layers[i].name, locked: doc.layers[i].locked });
    doc.layers[i].locked = false;
}

// Create and position new layer BEFORE locking anything
var titleBlock = doc.layers.getByName("Title Block");
var newLayer = doc.layers.add();
newLayer.name = "Site Plan YYYY-MM-DD";  // ← use PDF creation date
newLayer.move(titleBlock, ElementPlacement.PLACEBEFORE);

// Place PDF directly on the layer — NOTE: use layer.placedItems.add(), NOT doc.placedItems.add()
var placed = newLayer.placedItems.add();
placed.file = new File("C:/Projects/arborist-plans/Projects/PROJECT/Site Plan.pdf");

// Resize if scale_factor != 1.0  (e.g. 75 for pdf 1:150 into ai 1:200)
placed.resize(scale_factor * 100, scale_factor * 100);

// Embed — makes the .ai self-contained, no external PDF dependency
placed.embed();

// Restore original lock states
for (var k = 0; k < lockStates.length; k++) {
    try { doc.layers.getByName(lockStates[k].name).locked = lockStates[k].locked; } catch(e) {}
}
app.userInteractionLevel = UserInteractionLevel.DISPLAYALERTS; // restore

'Done. Embedded. Bounds: ' + Math.round(placed.width) + 'x' + Math.round(placed.height) + ' pts on "' + newLayer.name + '"';
```

---

## Step 4: Manual alignment

PDF coordinate origin ≠ .ai coordinate origin. The user must drag the placed layer
to register the drawings:

1. Find a feature visible in **both** drawings — property corner, road edge, building outline
2. Drag the "Site Plan YYYY-MM-DD" layer so those features overlap
3. Once aligned, lock the layer in the Layers panel

---

## Verification

```
mcp__illustrator__query
```
Confirm the new layer appears in the layers list with the correct name.
User visually confirms the PDF is visible and the drawings align after dragging.

---

## Layer naming convention

`"Site Plan YYYY-MM-DD"` — always use the PDF creation date so layers are traceable
across revisions.

---

## SHX / AutoCAD font problem

Civil engineering PDFs often use AutoCAD SHX fonts (RomanS, Romantic, Simplex, etc.).
Illustrator does not have these fonts and renders each glyph as a rectangle placeholder.

### Fix: flatten fonts to vector outlines with Ghostscript

**One-time setup:** Install Ghostscript from https://github.com/ArtifexSoftware/ghostpdl-downloads/releases/latest (grab `gs####w64.exe`). The installer puts `gswin64c.exe` at `C:\Program Files\gs\gs10.x.x\bin\` — add that folder to your user PATH.

**Per-file:** run `flatten_pdf.py` (project root) on the original PDF:

```bash
# Font flattening only:
python flatten_pdf.py "Projects/71 Lloyd/Site Plan (1).pdf"

# Remove embedded raster images (site photos, aerial backgrounds):
python flatten_pdf.py "Projects/71 Lloyd/Site Plan (1).pdf" --no-images

# Crop to drawing area + remove images (recommended for clean imports):
python flatten_pdf.py "Projects/71 Lloyd/Site Plan (1).pdf" --no-images --crop 741 286 2001 1299
```

Output is always `<name> flat.pdf` in the same folder.

This uses `gswin64c -sDEVICE=pdfwrite -dNoOutputFonts` which converts every font glyph to bezier outline paths. Output is still fully vector (not raster) — scalable, no pixelation, same visual quality.

### Finding the crop rectangle

The easiest way to determine crop coordinates for a new plan:

1. Run `--no-images` first to produce a clean flat PDF
2. Open it in any PDF viewer that supports annotations (Bluebeam, Acrobat, etc.)
3. Draw a **rectangle comment** (square/rectangle annotation tool) around the drawing area you want to keep
4. Save the PDF, then run:

```python
from pypdf import PdfReader
p = PdfReader("file flat.pdf").pages[0]
for a in p['/Annots']:
    obj = a.get_object()
    if obj.get('/Subtype') in ('/Square', '/Rectangle'):
        print(obj['/Rect'])
# → e.g. [740.706, 286.267, 2000.89, 1299.38]
```

Use those four values as `--crop x0 y0 x1 y1`.

**Then import the `flat.pdf` file** (not the original) in Step 3.

---

## Notes on tools (Windows)

- Use **Git Bash `pdftotext`** or **Windows Python** for file inspection
- Do NOT route through WSL unless the tool is only available there
- File paths in ExtendScript: use forward slashes (`C:/Projects/...`)
