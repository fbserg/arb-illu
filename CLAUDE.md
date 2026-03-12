# Arborist Plans — Claude Map

Automated TPZ circle placement on arborist site plans using Illustrator MCP + Excel inventory data.

## Current Project
**Scarlett** — `Projects/scarlett/` — active, primary focus until complete.
- Working file: `Projects/scarlett/PLAN.ai`
- Inventory: `Projects/scarlett/inventory.xlsx`
- PDFs: `PLAN flat.pdf` (flattened, use for import), `site-plan.pdf` (reference)

## Hard Rules (never break these)
- NEVER use `doc.pageItems` or `layer.pageItems` — use typed: `layer.pathItems` / `layer.textFrames` / `layer.groupItems`
- NEVER `JSON.stringify()` Illustrator objects — freezes app
- NEVER clear a layer in a `pathItems` while loop if sublayers exist — remove sublayers first, then loop
- TPZ layer is named **"TPZs"** (not "Trees", not "TPZ")
- NEVER use chained ternaries in ExtendScript — it associates left-to-right. Always parenthesise: `a ? 'x' : (b ? 'y' : 'z')`
- After manual position nudges in Illustrator: query MCP `geometricBounds` → update data.csv → update Excel P/Q. Excel is the source of truth; skipping the Excel write means the next `export_data.py` clobbers the corrected coords.

## Workflow
```
python export_data.py           # whenever Excel is edited → regenerates data.csv
python place_tpz.py --limit 10  # place in batches; use --all only when confident
python place_labels.py
python review.py                # 6-point audit
python extract_coords.py        # first-time only: extract coords from Dimensions layer → Excel → CSV
```

## Docs (source of truth)
- **Excel columns + CSV schema**: → `docs/excel-workflow.md`
- **Coordinate transform**: → `docs/coordinate-system.md`
- **Circle styles + layer structure**: → `docs/circle-styles.md`

## MCP Setup
- Server: `illustrator-mcp/` — config in `.mcp.json`
- COM retry: 3× with 0.3s delay (auto-heals transient errors)
- Illustrator must be running with the document open

## Verification
- Use `query` tool for intermediate checks (JSON snapshot of layer state)
- `view` is unreliable (AppActivate doesn't reliably focus Illustrator) — ask user to check directly
- After each `run`, call `query` to confirm expected circle count
