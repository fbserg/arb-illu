# Arborist Plans — Claude Map

Automated TPZ circle placement on arborist site plans using Illustrator MCP + Excel inventory data.

## What needs to be done

### 1. Read tree data from Excel
→ `docs/excel-workflow.md`
Key: use `win32com` (not openpyxl) to get computed values. Main data on **Sheet1**. Circle sizes in col N (mm), centerpoints in cols O/P.

### 2. Transform LogLog coordinates → Illustrator
→ `docs/coordinate-system.md`
Key: site plan is rotated 90° CW in Illustrator. X and Y axes swap + scale. Quick formula:
`cx = docW - y_pdf*(docW/docH)` / `cy = docH - x_pdf*(docH/docW)`

### 3. Place TPZ circles in Illustrator
→ `docs/circle-styles.md`
Key: all circles on **TPZs** layer. Protect=green solid, Injury=orange dashed, Removal=orange solid+X. Trunk sublayer inside TPZs.

### 4. Match existing circles to tree numbers
→ `docs/circle-styles.md#matching`
Key: find 3-pt elbow leader lines on Dimensions layer, match endpoints to labels and circle centers (dist < 80pt).

## MCP Setup
- Server: `illustrator-mcp/` — config in `.mcp.json`
- Stability fix already applied (`pythoncom.CoInitialize()`)
- `view` tool unreliable for screenshots — ask user to confirm visually
- Never `JSON.stringify()` Illustrator objects — freezes app
