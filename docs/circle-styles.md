# TPZ Circle Styles

## Layer Structure
```
TPZs              ← all tree groups (no sublayers)
  Tree N          ← groupItem named "Tree N" (one per tree)
    trunk circle  ← filled ellipse, 50% opacity (bottom of group)
    X lines       ← (Remove trees only)
    TPZ circle    ← stroked ellipse, no fill (top of group)
Labels            ← tree number labels (place_labels.py)
Dimensions        ← original leader lines + annotations
Site Plan         ← placed PDF
Title Block
```

## Circle Types

| Type | Direction value | Stroke color | Stroke style | Fill |
|---|---|---|---|---|
| Protect | "Protect" | CMYK(70, 0, 100, 0) green | 0.84pt solid | None |
| Retain | "Retain" | CMYK(70, 0, 100, 0) green | 0.84pt solid | None |
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
- TPZ radius: col K ("TPZ Circle Radius (m)") in Excel
- Convert to diameter in mm: `radius_m × 4.0` (×2 for 1:500 scale, ×2 for diameter)
- Convert mm → Illustrator pts: `mm × (72 / 25.4)`
- Trunk diameter: col R ("Trunk 1:500 (mm)") — already in mm, convert to pts directly

## Matching Existing Circles to Tree Numbers {#matching}
Labels, leader lines, and dimensions share the same layer (e.g. "Dimensions").
Leader lines are **3- or 4-point elbow open paths** (not straight 2-pt lines).

Algorithm:
1. Circles: `pathPoints.length 4–5`, `abs(w-h)/max(w,h) < 0.15`
2. Labels: textFrames matching `/^#\d+$/`
3. Leaders: 3- or 4-pt open paths — use `pathPoints[0]` and `pathPoints[n-1]` as outer endpoints (don't hardcode index 2)
4. Match threshold: dist < 80pt for both label end and circle end
