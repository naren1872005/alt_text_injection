# Implementation Summary: Download Unselected Figures (Excel)

## Overview
Replaced the "Download All Figures (ZIP)" button with a new "Download Unselected Figures (Excel)" feature that exports only the figures you've deselected from the PDF as an Excel workbook with embedded images and metadata.

## Changes Made

### 1. **HTML Changes** (`static/index.html`)
- **Line 224-232**: Renamed button element from `downloadZipBtn` to `downloadUnselectedBtn`
- Updated label text to "Download Unselected Figures (Excel)"
- Updated label element ID to `downloadUnselectedLabel`

### 2. **JavaScript Changes** (`static/app.js`)
- **Line 44-46**: Updated DOM element references
  - `downloadZipBtn` → `downloadUnselectedBtn`
  - `downloadZipLabel` → `downloadUnselectedLabel`

- **Line 334-341**: Updated click handler to call new `downloadUnselectedFigures()` function
  - For Excel tab: Still downloads Excel Images ZIP
  - For PDF tab: Calls new unselected figures export

- **Line 1271-1310**: Added new `downloadUnselectedFigures()` async function
  - Collects all figure IDs that are NOT in `selectedFigureIds` Set
  - Sends POST request to `/api/download-unselected-excel/{session_id}`
  - Downloads the Excel file with automatic naming

- **Line 693-738**: Updated `switchSourceTab()` function
  - Changed all `downloadZipLabel` references to `downloadUnselectedLabel`
  - Updated label text:
    - PDF tab: "Download Unselected Figures (Excel)"
    - Excel tab: "Download Excel Images (ZIP)"
    - Match tab: "Download Figures (ZIP)"

### 3. **Python Backend Changes** (`app.py`)
- **Line 8, 15-17**: Added imports
  ```python
  from typing import Optional, Dict, Any, List
  from openpyxl import Workbook
  from openpyxl.drawing.image import Image as XLImage
  from openpyxl.utils import get_column_letter
  ```

- **Line 553-554**: Added new Pydantic model for request validation
  ```python
  class UnselectedFiguresDownload(BaseModel):
      figure_ids: List[int]
  ```

- **Line 704-772**: New API endpoint `/api/download-unselected-excel/{session_id}`
  - Accepts POST request with list of figure IDs to download
  - Creates Excel workbook with:
    - Worksheet titled "Unselected Figures"
    - Headers: Figure ID, Page, MCIDs, Status, Alt Text, BBox, Width, Height, Image File
    - Figure metadata rows
    - Embedded thumbnail images (150x150px) in column I
    - Optimized column widths for readability
  - Returns Excel file (.xlsx) with automatic download

## How It Works

### User Workflow:
1. Upload a PDF to extract figures
2. View all extracted figures in the grid/table
3. Select figures you want to KEEP (for injection)
4. Click "Download Unselected Figures (Excel)"
5. An Excel file is generated containing only the figures you DIDN'T select

### Selection Mechanism:
- Figures are selected by clicking the checkbox or "+ Select" button on each figure card
- Selected figures are tracked in the global `selectedFigureIds` Set
- Unselected = all figures NOT in this Set
- Button is disabled if all figures are selected (nothing to download)

### Excel Output Structure:
```
Column A-H: Metadata (Figure ID, Page, MCIDs, Status, Alt Text, BBox, Width, Height)
Column I:   Embedded thumbnail image preview (150x150px)
```

## Integration with Existing Flow:
- Selection bar still works as before for injection
- Selection is independent for PDF figures tab
- When switching to Excel/Match tabs:
  - Button label updates appropriately
  - Button behavior changes based on tab
  - Selection bar visibility toggles

## Error Handling:
- Session not found → 404 error
- No figures in session → 404 error  
- No unselected figures → 400 error with message
- Image embedding failures → Warning logged (continues processing)

## File Structure After Download:
Session directory (`sessions/{session_id}/`):
- `input.pdf` - Original uploaded PDF
- `figures/` - Extracted figure crops (PNG)
- `unselected_figures_{session_id}.xlsx` - Generated Excel file
- `manifest.xlsx` - Uploaded Excel (if any)
- `excel_images/` - Extracted Excel images (if any)

## Testing Checklist:
- ✓ HTML elements renamed correctly
- ✓ JavaScript references updated
- ✓ No syntax errors in Python/JavaScript
- ✓ API endpoint properly implemented
- ✓ Image embedding with proper sizing
- ✓ Excel file generation with metadata
- ✓ Tab switching updates button label correctly
- ✓ Error handling for edge cases

## Notes:
- The feature maintains the existing UI/UX flow
- Selection mechanism is reused from injection feature
- Excel file includes thumbnail previews for visual reference
- File is automatically deleted from session when a new upload occurs
