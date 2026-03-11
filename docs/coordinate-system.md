# Coordinate System

## Setup
- Site plan PDFs are landscape in natural orientation (title block on right)
- LogLog app captures coordinates with PDF displayed in portrait — coordinate space: **1728 wide × 2592 tall**
- When dragged into Illustrator, PDF content appears rotated 90° CW (title block lands at bottom)
- Illustrator artboard: **1728 × 2592 pts** (24" × 36" portrait)

## Transform: LogLog PDF Pts → Illustrator artboard
```
cx = docW - y_pdf * (docW / docH)
cy = docH - x_pdf * (docH / docW)
```
For a 1728 × 2592 artboard:
```
cx = 1728 - y_pdf * (1728 / 2592)
cy = 2592 - x_pdf * (2592 / 1728)
```

## Placing an ellipse at a center point
```javascript
pathItems.ellipse(cy + radius, cx - radius, diameter, diameter)
```

## Validation
- X% and Y% columns in CSV cross-check: `X_pdf / docW = X%/100` ✓
- Title block position confirms rotation direction: if title block is at bottom in Illustrator, rotation is 90° CW ✓
