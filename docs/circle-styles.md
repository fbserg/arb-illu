# TPZ Circle Styles

## Layer Structure
```
TPZs              ← all TPZ circles
  └─ Trunks       ← trunk indicators (sublayer)
Dimensions        ← tree labels, leader lines, annotations (same layer)
Site Plan         ← placed PDF
Title Block
```

## Circle Types

| Type | Direction value | Stroke color | Stroke style | Fill |
|---|---|---|---|---|
| Protect | "Protect" | CMYK(70, 0, 100, 0) green | 0.84pt solid | None |
| Injury | "Injury" | CMYK(0, 62, 100, 0) orange | 0.84pt dashed (5pt) | None |
| Removal | "Remove" | CMYK(0, 62, 100, 0) orange | 0.84pt solid | None + X lines |
| Trunk | — | None | None | Same color as TPZ, 50% opacity |

## Removal X Lines
Same color and stroke width as the circle:
```
Line \:  (cx + r×0.707, cy + r×0.707)  →  (cx - r×0.707, cy - r×0.707)
Line /:  (cx - r×0.707, cy + r×0.707)  →  (cx + r×0.707, cy - r×0.707)
```

## Circle Sizing
- TPZ diameter: column N ("1:500") in Excel — already in mm at 1:500 scale
- Convert to Illustrator pts: `mm × (72 / 25.4)`
- Trunk diameter: column Q ("Trunk 1:500 (mm)") = `MAX(Multistem Calc / 50, min_floor)`
  - 7631 Creditview: fixed 2mm floor (all trees fell below 1:500 trunk size)
  - Future projects: use no floor, let scale differentiate naturally

## Matching Existing Circles to Tree Numbers {#matching}
Labels, leader lines, and dimensions share the same layer (e.g. "Dimensions").
Leader lines are **3-point elbow paths** (not straight 2-pt lines).

Algorithm:
1. Circles: `pathPoints.length 4–5`, `abs(w-h)/max(w,h) < 0.15`
2. Labels: textFrames matching `/^#\d+$/`
3. Leaders: **3- or 4-pt open paths** — use `pathPoints[0]` and `pathPoints[n-1]` as outer endpoints (don't hardcode index 2). 7631 Creditview had a mix of both styles.
4. Match threshold: dist < 80pt for both label end and circle end
