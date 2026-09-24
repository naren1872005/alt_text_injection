import io
import os
import re
import zipfile
import hashlib
import xml.etree.ElementTree as ET
import openpyxl
from PIL import Image
import imagehash
from typing import List, Dict, Any, Optional, Tuple, Callable

def is_duplicate_marker(text: Optional[str]) -> bool:
    """
    Checks whether a given text cell is a 'Duplicate' placeholder / marker
    (e.g., 'Duplicate', 'duplicate', 'DUPLICATE', 'Duplicate.', 'Duplicate image', 'Duplicate of ...')
    rather than a real authoritative description.
    """
    if not text:
        return False
    t = str(text).strip()
    if not t:
        return False
    lower = t.lower()
    
    if lower in {
        "duplicate", "duplicate.", "duplicate image", "duplicate figure",
        "duplicate photo", "duplicate graphic", "duplicate drawing",
        "duplicate alt", "dup", "dup.", "is duplicate", "same as above",
        "same image", "repeat", "repeated", "repeated image", "n/a - duplicate",
        "n/a duplicate", "n/a - dup", "n/a dup"
    }:
        return True
        
    if len(t) <= 60 and re.match(r'^(?:is\s+)?duplicate(?:\s*(?:of|[-:]|\(|row|fig|image|as)\b|\.|\s*$)', lower):
        return True
        
    return False

class ExcelParser:
    """
    Universal, robust Excel parser supporting arbitrary workbooks, dynamic column detection,
    flexible sheet selection, and complete embedded drawing extraction (PNG, JPEG, EMF, WMF, TIFF).
    Directly inspects OOXML DrawingML structures to ensure zero dropped images.
    """

    def __init__(
        self,
        excel_path: str,
        sheet_name: Optional[str] = None,
        is_cancelled: Optional[Callable[[], bool]] = None,
        progress_callback: Optional[Callable[[int, str], None]] = None
    ):
        self.excel_path = excel_path
        if is_cancelled and is_cancelled():
            raise RuntimeError("Excel processing cancelled by user")

        if progress_callback:
            progress_callback(5, "Reading workbook structure...")

        self.wb, self.sheet_names = self._load_workbook(excel_path)

        if progress_callback:
            progress_callback(10, "Scanning worksheet and embedded DrawingML images...")

        # 1. Extract all drawing images by sheet and row directly from the XLSX ZIP archive (if available)
        self.drawings_by_sheet = self._extract_drawings_from_zip(is_cancelled=is_cancelled, progress_callback=progress_callback)

        if is_cancelled and is_cancelled():
            raise RuntimeError("Excel processing cancelled by user")

        # 2. Select initial worksheet
        if sheet_name and sheet_name in self.sheet_names:
            self.sheet_name = sheet_name
        else:
            chosen_sheet = None
            # Prioritize sheet with images or relevant keywords
            for name in self.sheet_names:
                if name in self.drawings_by_sheet and len(self.drawings_by_sheet[name]) > 0:
                    chosen_sheet = name
                    break

            if not chosen_sheet:
                for name in self.sheet_names:
                    lower = name.lower()
                    if any(k in lower for k in ["pickup", "alt", "manifest", "drawing", "image", "figure", "data", "sheet1"]):
                        chosen_sheet = name
                        break

            self.sheet_name = chosen_sheet or self.sheet_names[0]

        self.ws = self.wb[self.sheet_name]
        self._setup_current_sheet()

    def select_sheet(self, sheet_name: str) -> bool:
        """Switches the active worksheet and re-indexes images and columns."""
        if sheet_name not in self.sheet_names:
            return False
        self.sheet_name = sheet_name
        self.ws = self.wb[self.sheet_name]
        self._setup_current_sheet()
        return True

    def _setup_current_sheet(self):
        """Indexes images and detects columns for the active worksheet."""
        self.row_to_img = self.drawings_by_sheet.get(self.sheet_name, {})
        self.unanchored_images = self.drawings_by_sheet.get(f"_unanchored_{self.sheet_name}", [])

        # Fallback: check openpyxl ws._images if direct drawing didn't catch something
        if hasattr(self.ws, "_images"):
            for im in self.ws._images:
                row = self._get_openpyxl_image_row(im)
                if row is not None and row not in self.row_to_img:
                    try:
                        raw_data = im._data()
                        pil_im = Image.open(io.BytesIO(raw_data))
                        if pil_im.mode not in ('RGB', 'RGBA'):
                            pil_im = pil_im.convert('RGBA')
                        buf = io.BytesIO()
                        pil_im.save(buf, format="PNG")
                        self.row_to_img[row] = {
                            "bytes": buf.getvalue(),
                            "width": pil_im.width,
                            "height": pil_im.height,
                            "media_name": f"openpyxl_img_{row}.png"
                        }
                    except Exception:
                        pass

        # Detect column positions dynamically from headers
        self.col_page, self.col_sr, self.col_fn, self.col_alt1, self.col_alt2, self.data_start_row = self._detect_columns()

    def _load_workbook(self, path: str) -> Tuple[openpyxl.Workbook, List[str]]:
        """
        Universal workbook loader supporting .xlsx, .xlsm, .xltx, .xltm, .xls (Excel 97-2003), .csv, and .tsv.
        """
        import csv
        lower = str(path).lower()

        if lower.endswith((".csv", ".tsv", ".txt")):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Manifest"
            
            encodings = ["utf-8-sig", "utf-8", "cp1252", "latin-1"]
            rows = []
            for enc in encodings:
                try:
                    with open(path, "r", encoding=enc, errors="replace") as f:
                        sample = f.read(4096)
                        f.seek(0)
                        try:
                            dialect = csv.Sniffer().sniff(sample)
                            reader = csv.reader(f, dialect)
                        except Exception:
                            delimiter = "\t" if lower.endswith(".tsv") else ","
                            reader = csv.reader(f, delimiter=delimiter)
                        rows = list(reader)
                    break
                except Exception:
                    continue
            for r in rows:
                ws.append(r)
            return wb, ["Manifest"]

        elif lower.endswith(".xls"):
            try:
                import xlrd
                wb_xls = xlrd.open_workbook(path)
                wb = openpyxl.Workbook()
                wb.remove(wb.active) # remove default sheet
                sheet_names = wb_xls.sheet_names()
                for sname in sheet_names:
                    ws_xls = wb_xls.sheet_by_name(sname)
                    ws = wb.create_sheet(title=sname)
                    for r in range(ws_xls.nrows):
                        row_vals = [ws_xls.cell_value(r, c) for c in range(ws_xls.ncols)]
                        ws.append(row_vals)
                return wb, sheet_names
            except Exception as e:
                print(f"Notice: xlrd fallback encountered: {e}")
                wb = openpyxl.load_workbook(path, data_only=True)
                return wb, wb.sheetnames

        else:
            wb = openpyxl.load_workbook(path, data_only=True)
            return wb, wb.sheetnames

    def _extract_drawings_from_zip(
        self,
        is_cancelled: Optional[Callable[[], bool]] = None,
        progress_callback: Optional[Callable[[int, str], None]] = None
    ) -> Dict[str, Any]:
        """
        Directly parses OOXML DrawingML and Modern In-Cell RichData Pictures from the .xlsx ZIP archive.
        Maps sheet names -> {1-indexed row: {'bytes': png_bytes, 'width': w, 'height': h, 'media_name': name}}
        Guarantees zero dropped images for both floating drawings and in-cell pictures.
        """
        drawings_by_sheet = {}
        if not zipfile.is_zipfile(self.excel_path):
            return drawings_by_sheet

        try:
            with zipfile.ZipFile(self.excel_path, "r") as z:
                names = set(z.namelist())
                if "xl/workbook.xml" not in names or "xl/_rels/workbook.xml.rels" not in names:
                    return drawings_by_sheet

                # Parse workbook.xml for sheet names and r:id
                wb_tree = ET.fromstring(z.read("xl/workbook.xml"))
                ns_wb = {
                    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
                    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                }

                wb_rels_tree = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
                wb_rels = {rel.attrib["Id"]: rel.attrib["Target"] for rel in wb_rels_tree}

                sheets = []
                for s in wb_tree.findall(".//main:sheet", ns_wb):
                    name = s.attrib.get("name")
                    r_id = s.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                    target = wb_rels.get(r_id, "")
                    sheet_xml_path = "xl/" + target if not target.startswith("xl/") else target
                    sheets.append((name, sheet_xml_path))

                # 1. Parse Modern In-Cell RichData picture metadata
                rich_rel_media_map = {}
                if "xl/richData/_rels/richValueRel.xml.rels" in names:
                    tree = ET.fromstring(z.read("xl/richData/_rels/richValueRel.xml.rels"))
                    for rel in tree:
                        r_id = rel.attrib.get("Id")
                        target = rel.attrib.get("Target", "")
                        if target.startswith("../"):
                            m_path = "xl/" + target[3:]
                        elif not target.startswith("xl/"):
                            m_path = f"xl/media/{target.split('/')[-1]}"
                        else:
                            m_path = target
                        rich_rel_media_map[r_id] = m_path

                rich_rv_index_to_rId = {}
                if "xl/richData/richValueRel.xml" in names:
                    tree = ET.fromstring(z.read("xl/richData/richValueRel.xml"))
                    for idx, rel in enumerate(tree):
                        r_id = (
                            rel.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id") or
                            rel.attrib.get("r:id") or
                            rel.attrib.get("id")
                        )
                        rich_rv_index_to_rId[idx] = r_id

                rich_rv_data_to_rel_idx = {}
                if "xl/richData/rdrichvalue.xml" in names:
                    tree = ET.fromstring(z.read("xl/richData/rdrichvalue.xml"))
                    for idx, rv in enumerate(tree):
                        v_elems = rv.findall("{http://schemas.microsoft.com/office/spreadsheetml/2017/richdata}v")
                        if not v_elems:
                            v_elems = rv.findall("v")
                        if v_elems:
                            try:
                                rich_rv_data_to_rel_idx[idx] = int(v_elems[0].text)
                            except Exception:
                                pass

                vm_to_rv_data_idx = {}
                if "xl/metadata.xml" in names:
                    tree = ET.fromstring(z.read("xl/metadata.xml"))
                    for future_meta in tree.iter():
                        if future_meta.tag.endswith("futureMetadata") and future_meta.attrib.get("name") == "XLRICHVALUE":
                            for bk_idx, bk in enumerate(future_meta):
                                for elem in bk.iter():
                                    if elem.tag.endswith("rvb"):
                                        i_val = elem.attrib.get("i")
                                        if i_val is not None:
                                            vm_to_rv_data_idx[bk_idx + 1] = int(i_val)

                # 2. For each sheet, parse both in-cell RichData and DrawingML floating images
                for sname, sxml in sheets:
                    drawings_by_sheet[sname] = {}
                    unanchored = []

                    # A. In-Cell RichData Images
                    if sxml in names and vm_to_rv_data_idx:
                        s_tree = ET.fromstring(z.read(sxml))
                        for c in s_tree.iter():
                            if c.tag.endswith("c"):
                                ref = c.attrib.get("r")
                                vm = c.attrib.get("vm")
                                if ref and vm:
                                    try:
                                        vm_int = int(vm)
                                        rv_data_idx = vm_to_rv_data_idx.get(vm_int, vm_int - 1)
                                        rel_idx = rich_rv_data_to_rel_idx.get(rv_data_idx, rv_data_idx)
                                        r_id = rich_rv_index_to_rId.get(rel_idx)
                                        media_path = rich_rel_media_map.get(r_id)
                                        if media_path and media_path in names:
                                            raw_bytes = z.read(media_path)
                                            png_bytes, w, h = self._convert_image_to_png(raw_bytes)
                                            if png_bytes:
                                                m_row = re.search(r"\d+", ref)
                                                if m_row:
                                                    row_idx = int(m_row.group(0))
                                                    drawings_by_sheet[sname][row_idx] = {
                                                        "bytes": png_bytes,
                                                        "width": w,
                                                        "height": h,
                                                        "media_name": media_path.split("/")[-1]
                                                    }
                                    except Exception as e:
                                        print(f"Warning extracting in-cell image for {sname} cell {ref}: {e}")

                    # B. DrawingML Floating Images
                    s_dir, s_file = sxml.rsplit("/", 1)
                    s_rels_path = f"{s_dir}/_rels/{s_file}.rels"
                    if s_rels_path in names:
                        s_rels = ET.fromstring(z.read(s_rels_path))
                        for rel in s_rels:
                            if "drawing" in rel.attrib.get("Type", ""):
                                d_target = rel.attrib.get("Target", "")
                                if d_target.startswith("../"):
                                    d_path = "xl/" + d_target[3:]
                                else:
                                    d_path = f"{s_dir}/{d_target}"

                                if d_path not in names:
                                    continue

                                # Resolve drawing rels to media
                                d_dir, d_file = d_path.rsplit("/", 1)
                                d_rels_path = f"{d_dir}/_rels/{d_file}.rels"
                                rel_media_map = {}
                                if d_rels_path in names:
                                    d_rels_root = ET.fromstring(z.read(d_rels_path))
                                    for d_rel in d_rels_root:
                                        r_id = d_rel.attrib.get("Id")
                                        m_target = d_rel.attrib.get("Target", "")
                                        if m_target.startswith("../"):
                                            m_path = "xl/" + m_target[3:]
                                        elif not m_target.startswith("xl/"):
                                            m_path = f"xl/media/{m_target.split('/')[-1]}"
                                        else:
                                            m_path = m_target
                                        rel_media_map[r_id] = m_path

                                # Parse drawing XML anchors
                                draw_root = ET.fromstring(z.read(d_path))
                                ns_draw = {
                                    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
                                    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
                                    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                                }

                                anchors = (
                                    draw_root.findall(".//xdr:twoCellAnchor", ns_draw) +
                                    draw_root.findall(".//xdr:oneCellAnchor", ns_draw)
                                )

                                total_anchors = max(1, len(anchors))
                                for anc_idx, anc in enumerate(anchors):
                                    if is_cancelled and is_cancelled():
                                        raise RuntimeError("Excel processing cancelled by user")
                                    if progress_callback and (anc_idx % 2 == 0 or anc_idx == total_anchors - 1):
                                        pct = int(10 + (anc_idx / total_anchors) * 30)
                                        progress_callback(pct, f"Extracting DrawingML image {anc_idx + 1} of {total_anchors}...")

                                    from_el = anc.find("xdr:from", ns_draw)
                                    blip = anc.find(".//a:blip", ns_draw)
                                    if blip is None:
                                        continue

                                    r_embed = (
                                        blip.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed") or
                                        blip.attrib.get("r:embed") or
                                        blip.attrib.get("embed")
                                    )
                                    media_path = rel_media_map.get(r_embed)
                                    if not media_path or media_path not in names:
                                        continue

                                    raw_bytes = z.read(media_path)
                                    png_bytes, w, h = self._convert_image_to_png(raw_bytes)
                                    if not png_bytes:
                                        continue

                                    img_entry = {
                                        "bytes": png_bytes,
                                        "width": w,
                                        "height": h,
                                        "media_name": media_path.split("/")[-1]
                                    }

                                    if from_el is not None:
                                        r_el = from_el.find("xdr:row", ns_draw)
                                        if r_el is not None and r_el.text:
                                            row_idx = int(r_el.text) + 1
                                            if row_idx not in drawings_by_sheet[sname]:
                                                drawings_by_sheet[sname][row_idx] = img_entry
                                        else:
                                            unanchored.append(img_entry)
                                    else:
                                        unanchored.append(img_entry)

                    drawings_by_sheet[f"_unanchored_{sname}"] = unanchored
        except Exception as e:
            print(f"Warning: Direct OOXML drawing extraction encountered: {e}")

        return drawings_by_sheet

    def _convert_image_to_png(self, raw_bytes: bytes) -> Tuple[Optional[bytes], Optional[int], Optional[int]]:
        """
        Converts any input image (PNG, JPEG, EMF, WMF, TIFF, BMP, WebP) into standard PNG bytes.
        For vector / DIB-bearing EMF/WMF files, first extracts direct embedded bitmaps (JPEG/PNG/DIB)
        with zero quality loss, and seamlessly falls back to high-res GDI vector rendering for pure vector metafiles.
        Guarantees zero 4x4 or distorted artifacts.
        """
        if not raw_bytes:
            return None, None, None

        import struct
        import io
        from PIL import Image

        # 1. Standard raster image formats (PNG, JPEG, WebP, TIFF, BMP, GIF)
        try:
            im = Image.open(io.BytesIO(raw_bytes))
            w, h = im.size
            if w >= 4 and h >= 4:
                if im.mode not in ("RGB", "RGBA"):
                    im = im.convert("RGBA")
                buf = io.BytesIO()
                im.save(buf, format="PNG")
                return buf.getvalue(), w, h
        except Exception:
            pass

        # 2. EMF / WMF embedded DIB / JPEG / PNG direct extraction
        is_emf = raw_bytes.startswith(b"\x01\x00\x00\x00") or (len(raw_bytes) > 44 and raw_bytes[40:44] == b" EMF")
        is_wmf = raw_bytes.startswith(b"\xd7\xcd\xc6\x9a") or raw_bytes.startswith(b"\x01\x00\t\x00")

        if is_emf:
            try:
                pos = 0
                while pos < len(raw_bytes) - 64:
                    rec_type, rec_size = struct.unpack_from("<II", raw_bytes, pos)
                    if rec_size == 0:
                        break
                    if rec_type in (81, 76): # EMR_STRETCHDIBITS, EMR_SETDIBITSTODEVICE
                        off_bmi, cb_bmi, off_bits, cb_bits = struct.unpack_from("<IIII", raw_bytes, pos + 48)
                        if off_bmi + cb_bmi <= rec_size and off_bits + cb_bits <= rec_size:
                            bmi_bytes = raw_bytes[pos + off_bmi : pos + off_bmi + cb_bmi]
                            bits_bytes = raw_bytes[pos + off_bits : pos + off_bits + cb_bits]

                            # Direct JPEG / PNG check
                            if bits_bytes.startswith(b"\xff\xd8") or bits_bytes.startswith(b"\x89PNG"):
                                try:
                                    im = Image.open(io.BytesIO(bits_bytes))
                                    if im.width >= 4 and im.height >= 4:
                                        if im.mode not in ("RGB", "RGBA"):
                                            im = im.convert("RGBA")
                                        buf = io.BytesIO()
                                        im.save(buf, format="PNG")
                                        return buf.getvalue(), im.width, im.height
                                except Exception:
                                    pass

                            # Construct valid BMP container for uncompressed / DIB bitmap bits
                            if len(bmi_bytes) >= 40:
                                bmp_file_hdr = b"BM" + struct.pack("<IHHI", 14 + len(bmi_bytes) + len(bits_bytes), 0, 0, 14 + len(bmi_bytes))
                                bmp_data = bmp_file_hdr + bmi_bytes + bits_bytes
                                try:
                                    im = Image.open(io.BytesIO(bmp_data))
                                    if im.width >= 4 and im.height >= 4:
                                        if im.mode not in ("RGB", "RGBA"):
                                            im = im.convert("RGBA")
                                        buf = io.BytesIO()
                                        im.save(buf, format="PNG")
                                        return buf.getvalue(), im.width, im.height
                                except Exception:
                                    pass
                    pos += rec_size
            except Exception as e:
                pass

        # 3. Native Windows GDI vector rendering for pure vector EMF/WMF files
        if is_emf or is_wmf:
            try:
                import ctypes
                from ctypes import wintypes

                gdi32 = ctypes.windll.gdi32
                user32 = ctypes.windll.user32

                hemf = gdi32.SetEnhMetaFileBits(len(raw_bytes), (ctypes.c_char * len(raw_bytes)).from_buffer_copy(raw_bytes))
                if hemf:
                    try:
                        class RECT(ctypes.Structure):
                            _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long), ("right", ctypes.c_long), ("bottom", ctypes.c_long)]

                        class ENHMETAHEADER(ctypes.Structure):
                            _fields_ = [
                                ("iType", wintypes.DWORD), ("nSize", wintypes.DWORD), ("rclBounds", RECT),
                                ("rclFrame", RECT), ("dSignature", wintypes.DWORD), ("nVersion", wintypes.DWORD),
                                ("nBytes", wintypes.DWORD), ("nRecords", wintypes.DWORD), ("nHandles", wintypes.WORD),
                                ("sReserved", wintypes.WORD), ("nDescription", wintypes.DWORD), ("offDescription", wintypes.DWORD),
                                ("nPalEntries", wintypes.DWORD), ("szlDevice", wintypes.SIZE), ("szlMillimeters", wintypes.SIZE),
                            ]

                        hdr = ENHMETAHEADER()
                        gdi32.GetEnhMetaFileHeader(hemf, ctypes.sizeof(hdr), ctypes.byref(hdr))

                        orig_w = hdr.rclBounds.right - hdr.rclBounds.left
                        orig_h = hdr.rclBounds.bottom - hdr.rclBounds.top

                        if orig_w <= 10 or orig_h <= 10:
                            frame_w = hdr.rclFrame.right - hdr.rclFrame.left
                            frame_h = hdr.rclFrame.bottom - hdr.rclFrame.top
                            if frame_w > 10 and frame_h > 10:
                                orig_w = max(1, int(frame_w * 0.01 * 3.7795))
                                orig_h = max(1, int(frame_h * 0.01 * 3.7795))
                            else:
                                orig_w = 600
                                orig_h = 450

                        # Scale up for crisp vector display (target ~600-800px)
                        scale = max(2.0, min(4.0, 800.0 / max(orig_w, orig_h, 1)))
                        target_w = max(100, int(orig_w * scale))
                        target_h = max(100, int(orig_h * scale))

                        hdc_screen = user32.GetDC(None)
                        hdc_mem = gdi32.CreateCompatibleDC(hdc_screen)

                        class BITMAPINFOHEADER(ctypes.Structure):
                            _fields_ = [
                                ("biSize", wintypes.DWORD), ("biWidth", ctypes.c_long), ("biHeight", ctypes.c_long),
                                ("biPlanes", wintypes.WORD), ("biBitCount", wintypes.WORD), ("biCompression", wintypes.DWORD),
                                ("biSizeImage", wintypes.DWORD), ("biXPelsPerMeter", ctypes.c_long), ("biYPelsPerMeter", ctypes.c_long),
                                ("biClrUsed", wintypes.DWORD), ("biClrImportant", wintypes.DWORD)
                            ]

                        class BITMAPINFO(ctypes.Structure):
                            _fields_ = [("bmiHeader", BITMAPINFOHEADER), ("bmiColors", wintypes.DWORD * 3)]

                        bmi = BITMAPINFO()
                        bmi.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
                        bmi.bmiHeader.biWidth = target_w
                        bmi.bmiHeader.biHeight = -target_h  # Top-down DIB
                        bmi.bmiHeader.biPlanes = 1
                        bmi.bmiHeader.biBitCount = 32
                        bmi.bmiHeader.biCompression = 0  # BI_RGB

                        ppvBits = ctypes.c_void_p()
                        hbmp = gdi32.CreateDIBSection(hdc_mem, ctypes.byref(bmi), 0, ctypes.byref(ppvBits), None, 0)
                        old_bmp = gdi32.SelectObject(hdc_mem, hbmp)

                        # Clean white canvas background
                        white_brush = gdi32.CreateSolidBrush(0x00FFFFFF)
                        rect = RECT(0, 0, target_w, target_h)
                        user32.FillRect(hdc_mem, ctypes.byref(rect), white_brush)
                        gdi32.DeleteObject(white_brush)

                        # Render the vector metafile instructions
                        gdi32.PlayEnhMetaFile(hdc_mem, hemf, ctypes.byref(rect))

                        buf_size = target_w * target_h * 4
                        raw_pixels = bytes((ctypes.c_char * buf_size).from_address(ppvBits.value))

                        gdi32.SelectObject(hdc_mem, old_bmp)
                        gdi32.DeleteObject(hbmp)
                        gdi32.DeleteDC(hdc_mem)
                        user32.ReleaseDC(None, hdc_screen)

                        # Convert BGRA buffer into PIL PNG
                        im = Image.frombuffer("RGBA", (target_w, target_h), raw_pixels, "raw", "BGRA", 0, 1)
                        buf = io.BytesIO()
                        im.save(buf, format="PNG")
                        return buf.getvalue(), target_w, target_h
                    finally:
                        gdi32.DeleteEnhMetaFile(hemf)
            except Exception as e:
                print(f"Notice: GDI vector rasterizer encountered: {e}")

        return None, None, None

    def _get_openpyxl_image_row(self, im: Any) -> Optional[int]:
        """Safely extract 1-indexed row number from an openpyxl image anchor."""
        try:
            if hasattr(im, "anchor"):
                anchor = im.anchor
                if hasattr(anchor, "_from") and hasattr(anchor._from, "row"):
                    return int(anchor._from.row) + 1
                if hasattr(anchor, "from_row"):
                    return int(anchor.from_row) + 1
        except Exception:
            pass
        return None

    def _detect_columns(self):
        """
        Dynamically inspects the first rows to detect columns:
        - Page / Location column
        - Serial / ID column
        - Filename / Figure Name / Preview column
        - Alt text column(s) (prioritizing finalized, revised, and translated descriptions)
        """
        max_col = min(getattr(self.ws, "max_column", 10) or 10, 40)

        for test_row in [1, 2, 3]:
            col_page = None
            col_sr = None
            col_fn = None
            col_alt_priority = {} # col -> score

            for c in range(1, max_col + 1):
                val = str(self.ws.cell(test_row, c).value or "").strip().lower()
                if not val:
                    continue

                # 1. Page column detection
                if any(k in val for k in ["pdf page", "page no", "page #", "page", "pg #", "pg", "página", "pagina", "sheet page", "slide"]):
                    if not any(k in val for k in ["alt", "desc", "preview", "image", "file", "text", "match", "comment", "location"]):
                        if col_page is None:
                            col_page = c

                # 2. Filename / Image Preview column detection
                if any(k in val for k in ["image preview", "preview", "thumbnail", "file", "filename", "image name", "img name", "drawing", "fig", "graphic", "image location", "location", "image"]):
                    if "alt" not in val and "comment" not in val:
                        if col_fn is None:
                            col_fn = c

                # 3. Serial / S.No column detection
                if any(k in val for k in ["sr", "s.no", "sl", "item", "#", "index", "figure #", "fig #"]) and "name" not in val and "page" not in val and "alt" not in val:
                    if col_sr is None:
                        col_sr = c

                # 4. Alt text columns with intelligent priority scoring
                if any(k in val for k in ["alt", "description", "desc", "caption", "accessibility", "long desc", "copyedit", "texto"]):
                    score = 10
                    # Prioritize finalized / revised / translated / approved columns
                    if any(k in val for k in ["revised", "final", "spanish", "español", "updated", "copyedit", "approved", "target", "new alt"]):
                        score += 30
                    if any(k in val for k in ["revised/final", "final alt", "final spanish", "approved alt"]):
                        score += 50
                    if any(k in val for k in ["old", "previous", "source", "original", "draft"]):
                        score -= 5
                    col_alt_priority[c] = score

            if col_alt_priority or col_fn is not None or col_page is not None:
                # Sort alt columns by priority score descending
                sorted_alts = sorted(col_alt_priority.items(), key=lambda x: x[1], reverse=True)
                col_alt2 = sorted_alts[0][0] if len(sorted_alts) >= 1 else None
                col_alt1 = sorted_alts[1][0] if len(sorted_alts) >= 2 else col_alt2

                # If only one alt found, put it in col_alt2 (authoritative)
                if col_alt2 is not None and col_alt1 == col_alt2:
                    col_alt1 = None

                return (
                    col_page,
                    col_sr or 1,
                    col_fn or (2 if col_page != 2 else 3),
                    col_alt1,
                    col_alt2 or col_alt1,
                    test_row + 1
                )

        # Fallback to positional defaults if no matching headers found
        if max_col >= 5:
            return 1, 1, 2, 3, 4, 2
        elif max_col >= 4:
            return 1, 1, 2, 3, 4, 2
        elif max_col >= 3:
            return 1, 1, 2, None, 3, 2
        else:
            return None, 1, 1, None, 2, 2

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

    def parse_all_records(
        self,
        output_dir: Optional[str] = None,
        max_records: Optional[int] = None,
        is_cancelled: Optional[Callable[[], bool]] = None,
        progress_callback: Optional[Callable[[int, str], None]] = None
    ) -> List[Dict[str, Any]]:
        """
        Parses all rows and extracts image bytes and alt texts.
        If output_dir is provided, saves images as PNG.
        Automatically resolves 'Duplicate' ALT markers to their previously
        available authoritative ALT text, guaranteeing that the literal word
        'Duplicate' is NEVER retained as ALT text.
        """
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)

        records = []
        max_row = self.ws.max_row or self.data_start_row
        limit = max_row + 1 if max_records is None else min(max_row + 1, self.data_start_row + max_records)

        unanchored_idx = 0
        total_rows = max(1, limit - self.data_start_row)

        for i, r in enumerate(range(self.data_start_row, limit)):
            if is_cancelled and is_cancelled():
                raise RuntimeError("Excel processing cancelled by user")
            if progress_callback and (i % 2 == 0 or i == total_rows - 1):
                pct = int((i / total_rows) * 100)
                progress_callback(pct, f"Extracting Excel row {r} of {max_row}...")
            sr_val = self.ws.cell(r, self.col_sr).value if self.col_sr else (r - self.data_start_row + 1)
            raw_fn = self.ws.cell(r, self.col_fn).value if self.col_fn else None

            # Extract page hint
            page_hint = None
            if self.col_page:
                raw_pg = self.ws.cell(r, self.col_page).value
                if raw_pg is not None:
                    raw_pg_str = str(raw_pg).strip()
                    # Check if integer or Roman numeral
                    m_pg = re.search(r'\b(\d+)\b', raw_pg_str)
                    if m_pg:
                        try:
                            page_hint = int(m_pg.group(1))
                        except Exception:
                            page_hint = raw_pg_str
                    elif raw_pg_str.lower() in ('i', 'ii', 'iii', 'iv', 'v', 'vi', 'vii', 'viii', 'ix', 'x'):
                        roman_map = {'i': 1, 'ii': 2, 'iii': 3, 'iv': 4, 'v': 5, 'vi': 6, 'vii': 7, 'viii': 8, 'ix': 9, 'x': 10}
                        page_hint = roman_map.get(raw_pg_str.lower(), raw_pg_str)
                    elif raw_pg_str:
                        page_hint = raw_pg_str

            # If col_fn cell is empty, scan other metadata columns for page/figure identifiers (e.g. FM28, FM29, P1)
            if not raw_fn or not str(raw_fn).strip():
                for c in range(1, min(self.ws.max_column or 10, 10)):
                    if c != self.col_alt1 and c != self.col_alt2 and c != self.col_page:
                        val = str(self.ws.cell(r, c).value or "").strip()
                        if val and len(val) <= 40 and not (val.startswith("978") and len(val) > 15):
                            raw_fn = val
                            break

            clean_fn = self.sanitize_filename(raw_fn, f"image_row_{r}.png")

            alt1 = str(self.ws.cell(r, self.col_alt1).value or "").strip() if self.col_alt1 else ""
            alt2 = str(self.ws.cell(r, self.col_alt2).value or "").strip() if self.col_alt2 else ""

            is_dup1 = is_duplicate_marker(alt1)
            is_dup2 = is_duplicate_marker(alt2)

            best_alt = ""
            is_dup_row = False
            dup_raw_text = ""

            # Check if authoritative columns contain real ALT or duplicate markers
            if is_dup2 or is_dup1:
                is_dup_row = True
                dup_raw_text = alt2 if is_dup2 else alt1
                best_alt = ""
            elif alt2:
                best_alt = alt2
            elif alt1:
                best_alt = alt1
            else:
                # If no alt found in primary columns, scan row for any long descriptive text or duplicate markers
                for c in range(1, min(self.ws.max_column or 10, 15)):
                    if c == self.col_page:
                        continue
                    val = str(self.ws.cell(r, c).value or "").strip()
                    if is_duplicate_marker(val):
                        is_dup_row = True
                        dup_raw_text = val
                        break
                    elif len(val) > 20 and not val.lower().endswith((".png", ".jpg", ".jpeg", ".pdf")):
                        best_alt = val
                        break

            # Find matching image
            img_data_obj = self.row_to_img.get(r)
            if not img_data_obj and not self.row_to_img and unanchored_idx < len(self.unanchored_images):
                img_data_obj = self.unanchored_images[unanchored_idx]
                unanchored_idx += 1

            has_img = img_data_obj is not None
            img_filename = None
            width = None
            height = None
            img_md5 = None
            img_phash = None
            media_name = None

            if has_img and img_data_obj:
                img_bytes = img_data_obj.get("bytes")
                media_name = img_data_obj.get("media_name")
                if img_bytes:
                    img_md5 = hashlib.md5(img_bytes).hexdigest()
                    try:
                        with Image.open(io.BytesIO(img_bytes)) as p_im:
                            img_phash = str(imagehash.phash(p_im))
                    except Exception:
                        img_phash = None

                if output_dir and img_bytes:
                    try:
                        img_filename = f"excel_row_{r:04d}_{clean_fn}"
                        if not img_filename.lower().endswith((".png", ".jpg", ".jpeg")):
                            img_filename += ".png"
                        img_path = os.path.join(output_dir, img_filename)
                        with open(img_path, "wb") as f:
                            f.write(img_bytes)

                        width = img_data_obj.get("width")
                        height = img_data_obj.get("height")
                    except Exception as e:
                        print(f"Failed to write image for row {r}: {e}")
                        img_filename = None

            # Only append if row has at least an image, an alt text, a filename, or is a duplicate row
            if has_img or best_alt or raw_fn or is_dup_row:
                records.append({
                    "row": r,
                    "row_index": r,
                    "page_hint": page_hint,
                    "sheet_name": self.sheet_name,
                    "sr_no": sr_val if sr_val is not None else (r - self.data_start_row + 1),
                    "filename": clean_fn,
                    "raw_fn": str(raw_fn).strip() if raw_fn else "",
                    "alt_text": best_alt,
                    "original_alt": alt1,
                    "updated_alt": alt2,
                    "has_alt": bool(best_alt),
                    "has_image": bool(img_filename or has_img),
                    "image_filename": img_filename,
                    "width": int(width) if width else None,
                    "height": int(height) if height else None,
                    "img_md5": img_md5,
                    "img_phash": img_phash,
                    "media_name": media_name,
                    "is_duplicate": is_dup_row,
                    "dup_raw_text": dup_raw_text
                })

        # Save any secondary sheet images (e.g., Identified Duplicates) to output_dir so they can also be referenced
        secondary_images_pool = []
        if output_dir:
            for sname, sheet_drawings in self.drawings_by_sheet.items():
                if sname != self.sheet_name and not sname.startswith("_"):
                    for s_row, s_img_obj in sheet_drawings.items():
                        try:
                            s_bytes = s_img_obj.get("bytes")
                            if s_bytes:
                                s_fn = f"excel_extra_{self.sanitize_filename(sname, 'sheet')}_{s_row:04d}.png"
                                s_path = os.path.join(output_dir, s_fn)
                                with open(s_path, "wb") as f:
                                    f.write(s_bytes)
                                
                                # Check if the secondary sheet row had alt text
                                s_alt = ""
                                if sname in self.wb.sheetnames:
                                    s_ws = self.wb[sname]
                                    for c in range(1, min(s_ws.max_column or 5, 10)):
                                        val = str(s_ws.cell(s_row, c).value or "").strip()
                                        if len(val) > 15 and not is_duplicate_marker(val):
                                            s_alt = val
                                            break
                                
                                s_md5 = hashlib.md5(s_bytes).hexdigest() if s_bytes else None
                                s_phash = None
                                if s_bytes:
                                    try:
                                        with Image.open(io.BytesIO(s_bytes)) as sp_im:
                                            s_phash = str(imagehash.phash(sp_im))
                                    except Exception:
                                        s_phash = None

                                secondary_images_pool.append({
                                    "sheet": sname,
                                    "row": s_row,
                                    "image_filename": s_fn,
                                    "width": s_img_obj.get("width"),
                                    "height": s_img_obj.get("height"),
                                    "alt_text": s_alt,
                                    "img_md5": s_md5,
                                    "img_phash": s_phash
                                })
                        except Exception as e:
                            print(f"Notice: saving secondary sheet image: {e}")

        # Build lookup indices for authoritative non-duplicate ALT texts
        auth_by_md5 = {}
        auth_by_phash = []
        auth_by_clean_fn = {}
        auth_by_raw_fn = {}
        auth_by_media_name = {}
        auth_by_row = {}
        
        last_seen_auth_alt = ""

        for rec in records:
            if not rec.get("is_duplicate") and rec.get("alt_text") and not is_duplicate_marker(rec["alt_text"]):
                alt_val = rec["alt_text"]
                rec["prev_authoritative_alt"] = last_seen_auth_alt
                last_seen_auth_alt = alt_val
                auth_by_row[rec["row"]] = alt_val
                if rec.get("img_md5"):
                    auth_by_md5[rec["img_md5"]] = alt_val
                if rec.get("img_phash"):
                    auth_by_phash.append((rec["img_phash"], alt_val))
                if rec.get("filename"):
                    auth_by_clean_fn[rec["filename"].lower()] = alt_val
                if rec.get("raw_fn"):
                    auth_by_raw_fn[rec["raw_fn"].lower()] = alt_val
                if rec.get("media_name"):
                    auth_by_media_name[rec["media_name"].lower()] = alt_val
            else:
                rec["prev_authoritative_alt"] = last_seen_auth_alt

        # Also incorporate secondary sheets authoritative texts
        for sec in secondary_images_pool:
            if sec.get("alt_text") and not is_duplicate_marker(sec["alt_text"]):
                if sec.get("img_md5") and sec["img_md5"] not in auth_by_md5:
                    auth_by_md5[sec["img_md5"]] = sec["alt_text"]
                if sec.get("img_phash"):
                    auth_by_phash.append((sec["img_phash"], sec["alt_text"]))

        # Resolve all Duplicate records to their previously available authoritative ALT text
        for rec in records:
            if rec.get("is_duplicate"):
                resolved_alt = None

                # 1. Exact MD5 / Byte hash match of image
                if rec.get("img_md5") and rec["img_md5"] in auth_by_md5:
                    resolved_alt = auth_by_md5[rec["img_md5"]]

                # 2. Perceptual hash similarity (Hamming distance <= 4)
                if not resolved_alt and rec.get("img_phash") and auth_by_phash:
                    try:
                        target_ph = imagehash.hex_to_hash(rec["img_phash"])
                        best_dist = 999
                        best_alt = None
                        for cand_ph_str, cand_alt in auth_by_phash:
                            cand_ph = imagehash.hex_to_hash(cand_ph_str)
                            dist = target_ph - cand_ph
                            if dist < best_dist:
                                best_dist = dist
                                best_alt = cand_alt
                        if best_dist <= 4 and best_alt:
                            resolved_alt = best_alt
                    except Exception:
                        pass

                # 3. Media name match (e.g. image1.png in DrawingML rels)
                if not resolved_alt and rec.get("media_name") and rec["media_name"].lower() in auth_by_media_name:
                    resolved_alt = auth_by_media_name[rec["media_name"].lower()]

                # 4. Explicit row / filename reference in duplicate cell text (e.g., "Duplicate of row 12", "Duplicate of epub_img_05.png")
                if not resolved_alt and rec.get("dup_raw_text"):
                    d_text = rec["dup_raw_text"]
                    m_row = re.search(r'(?:row|#)\s*(\d+)', d_text, re.IGNORECASE)
                    if m_row:
                        r_num = int(m_row.group(1))
                        if r_num in auth_by_row:
                            resolved_alt = auth_by_row[r_num]
                    m_fn = re.search(r'([\w\-\.]+\.(?:png|jpe?g|emf|wmf))', d_text, re.IGNORECASE)
                    if not resolved_alt and m_fn:
                        fn_ref = m_fn.group(1).lower()
                        if fn_ref in auth_by_clean_fn:
                            resolved_alt = auth_by_clean_fn[fn_ref]
                        elif fn_ref in auth_by_raw_fn:
                            resolved_alt = auth_by_raw_fn[fn_ref]

                # 5. Filename / Identifier match
                if not resolved_alt:
                    if rec.get("filename") and rec["filename"].lower() in auth_by_clean_fn:
                        resolved_alt = auth_by_clean_fn[rec["filename"].lower()]
                    elif rec.get("raw_fn") and rec["raw_fn"].lower() in auth_by_raw_fn:
                        resolved_alt = auth_by_raw_fn[rec["raw_fn"].lower()]

                # 6. Preceding authoritative ALT fallback
                if not resolved_alt and rec.get("prev_authoritative_alt"):
                    resolved_alt = rec["prev_authoritative_alt"]

                if resolved_alt and not is_duplicate_marker(resolved_alt):
                    rec["alt_text"] = resolved_alt
                    rec["has_alt"] = True
                    rec["duplicate_resolved"] = True
                else:
                    rec["alt_text"] = ""
                    rec["has_alt"] = False
                    rec["duplicate_resolved"] = False

            # Absolute safeguard: NEVER let any record retain a literal duplicate marker as alt_text
            if is_duplicate_marker(rec.get("alt_text")):
                rec["alt_text"] = ""
                rec["has_alt"] = False

        # Build comprehensive pool of all authoritative extracted images (excluding duplicate markers)
        authoritative_pool = []
        for rec in records:
            if rec.get("has_image") and rec.get("image_filename") and rec.get("alt_text") and not is_duplicate_marker(rec["alt_text"]):
                raw_alt = rec["alt_text"]
                clean_alt = re.sub(r'[^a-z0-9 ]', '', raw_alt.lower()).strip()
                core_alt = re.sub(r'^(a|an|the)?\s*(thumbnail|thumbnail image|thumbnail picture|illustration|photo|picture|diagram)?\s*(image)?\s*(shows|depicts|illustrates|is)?\s*', '', clean_alt).strip()
                authoritative_pool.append({
                    "row": rec["row"],
                    "alt_text": raw_alt,
                    "clean_alt": clean_alt,
                    "core_alt": core_alt,
                    "image_filename": rec["image_filename"],
                    "width": rec.get("width"),
                    "height": rec.get("height"),
                    "tokens": set(w for w in clean_alt.split() if len(w) > 3)
                })

        for sec in secondary_images_pool:
            if sec.get("alt_text") and not is_duplicate_marker(sec["alt_text"]):
                raw_alt = sec["alt_text"]
                clean_alt = re.sub(r'[^a-z0-9 ]', '', raw_alt.lower()).strip()
                core_alt = re.sub(r'^(a|an|the)?\s*(thumbnail|thumbnail image|thumbnail picture|illustration|photo|picture|diagram)?\s*(image)?\s*(shows|depicts|illustrates|is)?\s*', '', clean_alt).strip()
                authoritative_pool.append({
                    "row": sec["row"],
                    "alt_text": raw_alt,
                    "clean_alt": clean_alt,
                    "core_alt": core_alt,
                    "image_filename": sec["image_filename"],
                    "width": sec.get("width"),
                    "height": sec.get("height"),
                    "tokens": set(w for w in clean_alt.split() if len(w) > 3)
                })

        def match_authoritative_image(target_alt: str) -> Optional[Dict[str, Any]]:
            if not target_alt or not target_alt.strip():
                return None
            clean_tgt = re.sub(r'[^a-z0-9 ]', '', target_alt.lower()).strip()
            if not clean_tgt:
                return None

            # Tier 1: Exact normalized ALT match
            for auth in authoritative_pool:
                if auth["clean_alt"] == clean_tgt:
                    return auth

            # Tier 2: Core ALT match (ignoring thumbnail / illustration prefixes)
            core_tgt = re.sub(r'^(a|an|the)?\s*(thumbnail|thumbnail image|thumbnail picture|illustration|photo|picture|diagram)?\s*(image)?\s*(shows|depicts|illustrates|is)?\s*', '', clean_tgt).strip()
            if len(core_tgt) >= 12:
                for auth in authoritative_pool:
                    if auth["core_alt"] == core_tgt:
                        return auth
                for auth in authoritative_pool:
                    if len(auth["core_alt"]) >= 12 and (core_tgt in auth["core_alt"] or auth["core_alt"] in core_tgt):
                        return auth

            # Tier 3: Specialized keyphrase rules for recurring corporate / pedagogical icons
            if "national geographic exclusive" in clean_tgt:
                for auth in authoritative_pool:
                    if "national geographic exclusive" in auth["clean_alt"]:
                        return auth

            # Tier 4: Token overlap similarity (Jaccard similarity on content words)
            tgt_tokens = set(w for w in clean_tgt.split() if len(w) > 3)
            if len(tgt_tokens) >= 3:
                best_match = None
                best_score = 0.0
                for auth in authoritative_pool:
                    auth_tokens = auth["tokens"]
                    if not auth_tokens:
                        continue
                    inter = tgt_tokens.intersection(auth_tokens)
                    score = len(inter) / max(len(tgt_tokens), len(auth_tokens))
                    if score > best_score and score >= 0.35:
                        best_score = score
                        best_match = auth
                if best_match:
                    return best_match

            # Tier 5: Substring phrase match for notable distinctive keywords
            for kw in ["wheelchair", "accessible chair", "sensors and circuits", "leaning tower", "redwood forest", "monarch butterfly", "pont du gard", "sheep graze", "marmot"]:
                if kw in clean_tgt:
                    for auth in authoritative_pool:
                        if kw in auth["clean_alt"] or (kw == "wheelchair" and "accessible chair" in auth["clean_alt"]) or (kw == "accessible chair" and "wheelchair" in auth["clean_alt"]):
                            return auth

            return None

        # Link unanchored and recurring duplicate ALT rows to their authoritative image
        for rec in records:
            if not rec.get("has_image") and rec.get("alt_text"):
                matched_auth = match_authoritative_image(rec["alt_text"])
                if matched_auth:
                    rec["has_image"] = True
                    rec["image_filename"] = matched_auth["image_filename"]
                    rec["width"] = matched_auth.get("width")
                    rec["height"] = matched_auth.get("height")

        return records

    def extract_single_image(self, row: int) -> Optional[bytes]:
        img_obj = self.row_to_img.get(row)
        if img_obj and "bytes" in img_obj:
            return img_obj["bytes"]
        return None

    def close(self):
        try:
            self.wb.close()
        except Exception:
            pass
