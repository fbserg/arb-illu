# Excel Workflow

## Reading the File
Use `win32com` — openpyxl `data_only=True` returns None for array formulas:
```python
import win32com.client
xl = win32com.client.Dispatch('Excel.Application')
wb = xl.Workbooks.Open(r'absolute\path\to\file.xlsx')
ws = wb.Sheets('Sheet1')
value = ws.Cells(row, col).Value
wb.Close(False)
```

## Sheet1 Column Reference
| Col | Header | Notes |
|---|---|---|
| A | Tree # | Integer, primary key |
| B | Species | |
| C | Botanical Name | |
| D | DBH (cm) | May be string for multistem e.g. "32, 35, 48" |
| E | Condition Rating | |
| F | Comments | |
| G | Ownership Category | |
| H | Crown Radius (m) | |
| I | Direction | "Protect" / "Injury" / "Remove" |
| J | TPZ (m) | Array formula — "N/A" for Remove trees |
| K | Permit Requirement | |
| L | Multistem Calculation | Effective single-stem DBH equivalent — use this, not col D |
| M | Map TPZ Diameter (m) | |
| N | 1:500 | Circle diameter in mm at 1:500 scale |
| O | Center X (AI pts) | Added by Claude — Illustrator artboard X |
| P | Center Y (AI pts) | Added by Claude — Illustrator artboard Y |
| Q | Trunk 1:500 (mm) | Added by Claude — trunk circle diameter |

## Key Formulas
- N (1:500): `= (M × 1000) / 500` → mm on paper at 1:500 scale
- Q (Trunk): `= MAX(L / 50, floor)` — use col L (multistem calc), not raw DBH
- Convert mm → Illustrator pts: `mm × (72 / 25.4)`

## Notes
- Remove trees: TPZ = "N/A", columns M/N error out — no TPZ circle size available from formula
- Multistem DBH string in col D: use col L for calculations, it normalises to a single equivalent value
