# Arborist Plans — Claude Map

Automated TPZ circle placement on arborist site plans using Illustrator MCP + Excel inventory data.

## Hard Rules (never break these)
- NEVER use `doc.pageItems` or `layer.pageItems` on Site Plan or any layer with placed PDFs
- Always use typed collections: `layer.pathItems` / `layer.textFrames` / `layer.groupItems`
- NEVER `JSON.stringify()` Illustrator objects — freezes app

## What needs to be done

### 1. Read tree data from Excel
→ `docs/excel-workflow.md`
Key: use `win32com` (not openpyxl) to get computed values. Main data on **Sheet1**. Circle sizes in col N (mm), centerpoints in cols O/P, trunk sizes in col Q.

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
- COM retry already applied (3× retry with 1s delay)

## Placement workflow
1. Read + transform: `python place_tpz.py` — places TPZ circles + trunks in one pass
2. Labels: `python place_labels.py` — places tree number labels
3. Verify: `python review.py` — 6-point audit
4. Visual: ask user to confirm via `view`

## Verification (use query, not view)
- `query` tool → JSON snapshot of layer state, circle counts, positions
- Use `view` only for final human sign-off, not intermediate checks
- After each `run`, call `query` to confirm expected circle count before proceeding

## COM reliability
- `run` tool retries 3× before failing — transient errors should self-heal
- If 3 retries all fail: Illustrator may be unresponsive; ask user to check
