import io
import os
import re
import openpyxl
from PIL import Image
from typing import List, Dict, Any, Optional

class ExcelParser:
    """
    Universal Excel parser supporting arbitrary workbooks, dynamic column detection,
    flexible sheet selection, and embedded drawing extraction.
    Works for any customer Excel ALT manifest, not just default/sample files.
    """

    def __init__(self, excel_path: str):
        self.excel_path = excel_path
        self.wb = openpyxl.load_workbook(excel_path, data_only=True)
        sheet_names = self.wb.sheetnames
        
        # 1. Select sheet: prioritize sheet with images or relevant keywords
        chosen_sheet = None
        for name in sheet_names:
            ws = self.wb[name]
            if hasattr(ws, "_images") and len(ws._images) > 0:
                chosen_sheet = name
                break

        if not chosen_sheet:
            for name in sheet_names:
                lower = name.lower()
                if any(k in lower for k in ["alt", "manifest", "drawing", "image", "figure", "data"]):
                    chosen_sheet = name
                    break

        self.sheet_name = chosen_sheet or sheet_names[0]
        self.ws = self.wb[self.sheet_name]

        # 2. Build map from 1-indexed row number to drawing Image object
        self.row_to_img = {}
        self.unanchored_images = []

        if hasattr(self.ws, "_images"):
            for idx, im in enumerate(self.ws._images):
                row = self._get_image_row(im)
                if row is not None:
                    self.row_to_img[row] = im
                else:
                    self.unanchored_images.append(im)

        # 3. Detect column positions dynamically from headers
        self.col_sr, self.col_fn, self.col_alt1, self.col_alt2, self.data_start_row = self._detect_columns()

    def _get_image_row(self, im: Any) -> Optional[int]:
        """Safely extract 1-indexed row number from an openpyxl image anchor."""
        try:
            if hasattr(im, 'anchor'):
                anchor = im.anchor
                if hasattr(anchor, '_from') and hasattr(anchor._from, 'row'):
                    return int(anchor._from.row) + 1
                if hasattr(anchor, 'from_row'):
                    return int(anchor.from_row) + 1
        except Exception:
            pass
        return None

    def _detect_columns(self):
        """
        Dynamically inspects the first rows to detect columns:
        - Serial / ID column
        - Filename / Figure Name column
        - Alt text column(s)
        """
        max_col = min(getattr(self.ws, "max_column", 10) or 10, 40)
        
        # Test row 1 and row 2 for header labels
        for test_row in [1, 2]:
            col_sr = None
            col_fn = None
            col_alt1 = None
            col_alt2 = None

            for c in range(1, max_col + 1):
                val = str(self.ws.cell(test_row, c).value or "").strip().lower()
                if not val:
                    continue

                if any(k in val for k in ["updated", "revised", "new alt", "final alt"]):
                    col_alt2 = c
                elif any(k in val for k in ["alt", "description", "desc", "caption", "accessibility", "long desc"]):
                    if col_alt1 is None:
                        col_alt1 = c
                    elif col_alt2 is None:
                        col_alt2 = c
                elif any(k in val for k in ["file", "filename", "image name", "img name", "drawing", "fig", "graphic"]) and "alt" not in val:
                    col_fn = c
                elif any(k in val for k in ["sr", "s.no", "sl", "item", "#", "index", "figure #", "fig #"]) and "name" not in val:
                    col_sr = c

            # If we found at least an alt column or filename column, we found the header row
            if col_alt1 is not None or col_fn is not None:
                return (
                    col_sr or 1,
                    col_fn or 2,
                    col_alt1 or (4 if max_col >= 4 else max_col),
                    col_alt2,
                    test_row + 1
                )

        # Fallback to positional defaults if no matching headers found
        if max_col >= 5:
            return 1, 2, 4, 5, 2
        elif max_col >= 4:
            return 1, 2, 3, 4, 2
        elif max_col >= 3:
            return 1, 2, 3, None, 2
        elif max_col >= 2:
            return 1, 1, 2, None, 2
        else:
            return 1, 1, 1, None, 2

    def get_summary(self) -> Dict[str, Any]:
        return {
            "sheet_name": self.sheet_name,
            "total_rows": max(0, (self.ws.max_row or 0) - self.data_start_row + 1),
            "total_images": len(self.row_to_img) + len(self.unanchored_images)
        }

    def sanitize_filename(self, fn: str, fallback: str) -> str:
        if not fn or not str(fn).strip():
            return fallback
        clean = re.sub(r'[\\/*?:"<>|\r\n\t]', '_', str(fn).strip())
        return clean or fallback

    def parse_all_records(self, output_dir: Optional[str] = None, max_records: Optional[int] = None) -> List[Dict[str, Any]]:
        """
        Parses all rows and extracts image bytes and alt texts.
        If output_dir is provided, saves images as PNG.
        """
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)

        records = []
        max_row = self.ws.max_row or self.data_start_row
        limit = max_row + 1 if max_records is None else min(max_row + 1, self.data_start_row + max_records)

        # Unanchored images iterator
        unanchored_idx = 0

        for r in range(self.data_start_row, limit):
            sr_val = self.ws.cell(r, self.col_sr).value if self.col_sr else (r - self.data_start_row + 1)
            raw_fn = self.ws.cell(r, self.col_fn).value if self.col_fn else None
            clean_fn = self.sanitize_filename(raw_fn, f"image_row_{r}.png")
            
            alt1 = str(self.ws.cell(r, self.col_alt1).value or "").strip() if self.col_alt1 else ""
            alt2 = str(self.ws.cell(r, self.col_alt2).value or "").strip() if self.col_alt2 else ""
            
            # Use Updated ALT Text if available, else original ALT Text
            best_alt = alt2 if alt2 else alt1
            
            # If no alt found in primary column, scan row for any long descriptive text
            if not best_alt:
                for c in range(1, min(self.ws.max_column or 10, 15)):
                    val = str(self.ws.cell(r, c).value or "").strip()
                    if len(val) > 20 and not val.lower().endswith((".png", ".jpg", ".jpeg", ".pdf")):
                        best_alt = val
                        break

            # Find matching image
            im = self.row_to_img.get(r)
            if not im and not self.row_to_img and unanchored_idx < len(self.unanchored_images):
                im = self.unanchored_images[unanchored_idx]
                unanchored_idx += 1

            has_img = im is not None
            img_filename = None
            width = None
            height = None

            if has_img and output_dir and im:
                try:
                    img_data = im._data()
                    img_filename = f"excel_row_{r:04d}_{clean_fn}"
                    if not img_filename.lower().endswith((".png", ".jpg", ".jpeg")):
                        img_filename += ".png"
                    img_path = os.path.join(output_dir, img_filename)
                    with open(img_path, "wb") as f:
                        f.write(img_data)

                    try:
                        width = getattr(im, 'width', None)
                        height = getattr(im, 'height', None)
                    except Exception:
                        pass
                except Exception:
                    img_filename = None

            # Only append if row has at least an image, an alt text, or a filename
            if has_img or best_alt or raw_fn:
                records.append({
                    "row": r,
                    "sr_no": sr_val if sr_val is not None else (r - self.data_start_row + 1),
                    "filename": clean_fn,
                    "alt_text": best_alt,
                    "original_alt": alt1,
                    "updated_alt": alt2,
                    "has_alt": bool(best_alt),
                    "has_image": has_img,
                    "image_filename": img_filename,
                    "width": int(width) if width else None,
                    "height": int(height) if height else None
                })

        return records

    def extract_single_image(self, row: int) -> Optional[bytes]:
        im = self.row_to_img.get(row)
        if im:
            try:
                return im._data()
            except Exception:
                pass
        return None

    def close(self):
        try:
            self.wb.close()
        except Exception:
            pass
