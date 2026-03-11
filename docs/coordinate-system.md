# Coordinate System

## Setup
- Site plan PDFs are landscape in natural orientation (title block on right)
- LogLog app captures coordinates with PDF displayed in portrait — coordinate space: **1728 wide × 2592 tall**
- When dragged into Illustrator, PDF content appears rotated 90° CW (title block lands at bottom)

## 7631 Creditview artboard
The artboard is **landscape in document coordinates** (2592 wide × 1728 tall) and is
**offset** from the document origin:
```
artboardRect = [-684, 468, 1908, -1260]
  (left, top, right, bottom in Illustrator coordinates)
```

## Transform: LogLog PDF Pts → Illustrator artboard
```
cx = -684 + y_pdf
cy =  468 - x_pdf
```
Where `x_pdf` and `y_pdf` are the LogLog coordinates in PDF points.

Why there is no scale factor: the LogLog portrait coordinate space is 1728 wide × 2592 tall,
which matches the artboard's 1728 tall × 2592 wide exactly (just transposed by the 90° rotation).
So the mapping is 1:1 — only a translation and y-axis flip, no multiplication.

## Inverse: Illustrator artboard → LogLog PDF Pts
```
y_pdf = cx + 684
x_pdf = 468 - cy
```

## Placing an ellipse at a center point
```javascript
pathItems.ellipse(cy + radius, cx - radius, diameter, diameter)
```
(Illustrator `ellipse(top, left, width, height)` — top = cy + r, left = cx − r)

## Artboard bounds check
Valid Illustrator coords are within:
```
x: [-684, 1908]
y: [-1260, 468]
```

## Data provenance

Coordinates originate from the **LogLog web app**, which exports a CSV with `X %` and `Y %` columns.
These are converted to PDF points on import:
```
X (PDF Pts) = X% / 100 * 1728
Y (PDF Pts) = Y% / 100 * 2592
```
**X (PDF Pts) and Y (PDF Pts) are the ground truth** — they come directly from LogLog and should
never need to be recomputed. Only `cx`/`cy` (the Illustrator artboard coords) are derived from them.

The original `pins.csv` had wrong `cx`/`cy` because the legacy formula below was used instead of
the correct artboard-offset formula. The X/Y PDF Pts in that file were always correct.

## Validation
- Verified against items 41/42 (X_pdf=574.0, Y_pdf=499.5 → cx=-184.5, cy=-106.0) ✓
- Verified against item 49 (X_pdf=614.0, Y_pdf=1284.0 → cx=600.0, cy=-146.0) ✓

---

## Legacy formula (pre-7631 Creditview fix)

The original formula assumed a **portrait artboard at the document origin** (1728 × 2592, no offset).
This may apply if a project was set up with the PDF as the artboard origin:
```
cx = 1728 - y_pdf * (1728 / 2592)   →  cx = 1728 - y_pdf * 0.6667
cy = 2592 - x_pdf * (2592 / 1728)   →  cy = 2592 - x_pdf * 1.5
```
Inverse:
```
y_pdf = (1728 - cx) * 1.5
x_pdf = (2592 - cy) * 0.6667
```
If you need to recover PDF coords from old cx/cy values computed with this formula, use the inverse above.
