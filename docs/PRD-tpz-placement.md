# PRD: Automated TPZ Circle Placement

## Problem
Arborists manually place Tree Protection Zone (TPZ) circles on Illustrator site plans — a repetitive, error-prone process requiring correct sizing, styling, and positioning for each tree type.

## Goal
Given a tree inventory Excel file and an Illustrator site plan, automatically place correctly sized and styled TPZ circles at each tree's location on the correct layer.

---

## Inputs
1. **Illustrator site plan** — open and active, PDF placed on artboard (1728 × 2592 pts)
2. **Excel workbook** (Sheet1) — tree inventory with TPZ calculations and LogLog coordinates
3. **LogLog coordinates** — X/Y PDF Pts captured from the site plan PDF

## Outputs
- **TPZs layer** in Illustrator with one circle per tree, correctly:
  - Positioned at tree center (coordinate-transformed from LogLog)
  - Sized per the 1:500 column in Excel
  - Styled per tree Direction (Protect / Injury / Remove)
- **Trunks sublayer** inside TPZs with a small filled circle at each tree center

---

## Circle Spec (see `circle-styles.md` for full detail)
| Type | Color | Style |
|---|---|---|
| Protect | Green CMYK(70,0,100,0) | Solid stroke |
| Injury | Orange CMYK(0,62,100,0) | Dashed stroke (5pt) |
| Removal | Orange CMYK(0,62,100,0) | Solid stroke + X through center |
| Trunk | Same as TPZ type | Filled, 50% opacity |

---

## Workflow Steps

### Step 1 — Read Excel
- Use `win32com` to get computed values from Sheet1
- Extract: Tree #, Direction, 1:500 (mm), Center X, Center Y, Trunk 1:500 (mm)
- Skip trees where 1:500 is missing/error (Remove trees with no TPZ formula)

### Step 2 — Transform coordinates
- Apply 90° CW rotation formula (see `coordinate-system.md`)
- Verify artboard dimensions match expected 1728 × 2592

### Step 3 — Place circles
- Clear or create target layer ("TPZs")
- For each tree with valid size + center: place circle with correct style
- For Removal trees: add X lines through center
- Create Trunks sublayer, place filled trunk circles

### Step 4 — Verify
- User visually confirms a sample of circles land on trees
- Spot-check sizes against known trees (e.g. large DBH should have visibly larger circle)

---

## Out of Scope (this phase)
- Automatically capturing LogLog coordinates (manual step)
- Placing tree number labels / leader lines
- Handling Remove trees with no calculated TPZ size
- Multi-group trees sharing a center point (currently both get a circle at the same location)

---

## Known Issues / Future Work
- **Remove trees**: TPZ formula returns N/A — need to decide circle size for removal markers
- **Multi-group centers**: trees sharing a single survey point get overlapping circles
- **Trunk scale**: at 1:500, all trunks fall below 2mm — future projects should drop the floor and use true scale for size differentiation
- **view tool**: Illustrator screenshot via MCP is unreliable (AppActivate doesn't focus window) — rely on user confirmation
