import io
import os
import re
import numpy as np
import pypdf
from pypdf.generic import NameObject, TextStringObject, DictionaryObject, ArrayObject, IndirectObject
import pymupdf
from PIL import Image
from typing import List, Dict, Any, Optional, Tuple

PAINT_OPERATORS = {
    "f", "F", "f*",
    "S", "s",
    "B", "B*",
    "b", "b*"
}

def _deref(obj):
    if isinstance(obj, IndirectObject):
        return obj.get_object()
    return obj

def decode_pdf_string(val: Any) -> Optional[str]:
    """
    Decodes PDF string objects (/Alt, /ActualText, /T) accurately,
    eliminating UTF-16 mis-decoding and CJK mojibake corruption.
    """
    if val is None:
        return None

    raw_bytes = None
    if hasattr(val, "original_bytes") and val.original_bytes:
        raw_bytes = val.original_bytes
    elif hasattr(val, "get_original_bytes"):
        try:
            raw_bytes = val.get_original_bytes()
        except Exception:
            pass
    elif isinstance(val, (bytes, bytearray)):
        raw_bytes = bytes(val)

    if raw_bytes:
        if raw_bytes.startswith(b'\xfe\xff'):
            try:
                txt = raw_bytes[2:].decode('utf-16-be').strip()
                if txt:
                    return txt
            except Exception:
                pass
        elif raw_bytes.startswith(b'\xff\xfe'):
            try:
                txt = raw_bytes[2:].decode('utf-16-le').strip()
                if txt:
                    return txt
            except Exception:
                pass
        elif raw_bytes.startswith(b'\xef\xbb\xbf'):
            try:
                txt = raw_bytes[3:].decode('utf-8').strip()
                if txt:
                    return txt
            except Exception:
                pass

        if len(raw_bytes) >= 4:
            if raw_bytes[0] == 0 and raw_bytes[2] == 0:
                try:
                    txt = raw_bytes.decode('utf-16-be').strip()
                    if txt:
                        return txt
                except Exception:
                    pass
            elif raw_bytes[1] == 0 and raw_bytes[3] == 0:
                try:
                    txt = raw_bytes.decode('utf-16-le').strip()
                    if txt:
                        return txt
                except Exception:
                    pass

        try:
            txt = raw_bytes.decode('utf-8').strip()
            if txt:
                return txt
        except Exception:
            pass
        try:
            txt = raw_bytes.decode('latin1').strip()
            if txt:
                return txt
        except Exception:
            pass

    s = str(val).strip("\x00 \t\r\n")
    if not s:
        return None

    cjk_count = sum(1 for c in s if 0x4E00 <= ord(c) <= 0x9FFF)
    if cjk_count > 0 and cjk_count >= max(1, len(s) * 0.15):
        reconstructed = bytearray()
        for c in s:
            val_ord = ord(c)
            b_low = val_ord & 0xFF
            b_high = (val_ord >> 8) & 0xFF
            if b_high != 0:
                reconstructed.append(b_low)
                reconstructed.append(b_high)
            else:
                reconstructed.append(b_low)
        
        for enc in ['utf-16-be', 'utf-16-le', 'utf-8', 'latin1']:
            try:
                dec = reconstructed.decode(enc).strip()
                if dec and sum(1 for c in dec if 0x4E00 <= ord(c) <= 0x9FFF) == 0:
                    return dec
            except Exception:
                pass

    return s

class FigureExtractor:
    """
    Extracts /Figure tags from a PDF's accessibility structure (StructTreeRoot),
    resolves their page numbers, Marked Content Identifiers (MCID), bounding boxes,
    and associated /XObject /Image resources.
    """

    def __init__(self, pdf_path: str):
        self.pdf_path = pdf_path
        self.doc = pymupdf.open(pdf_path)
        self.reader = pypdf.PdfReader(pdf_path)
        self.page_count = len(self.doc)
        
        # Map pypdf page object identities to 1-indexed page numbers
        self.page_obj_to_num = {}
        for idx, page in enumerate(self.reader.pages):
            ref = page.indirect_reference
            if ref:
                self.page_obj_to_num[(ref.idnum, ref.generation)] = idx + 1
                self.page_obj_to_num[ref.idnum] = idx + 1

    @property
    def has_struct_tree(self) -> bool:
        trailer = getattr(self.reader, "trailer", {}) or {}
        root = _deref(trailer.get("/Root") or getattr(self.reader, "root_object", {}))
        return bool(root and "/StructTreeRoot" in root)

    def _get_page_num(self, pg_elem, page_obj_to_num_map) -> Optional[int]:
        if pg_elem is None:
            return None
        if isinstance(pg_elem, IndirectObject):
            return page_obj_to_num_map.get((pg_elem.idnum, pg_elem.generation), page_obj_to_num_map.get(pg_elem.idnum, None))
        if hasattr(pg_elem, "indirect_reference") and pg_elem.indirect_reference:
            ref = pg_elem.indirect_reference
            return page_obj_to_num_map.get((ref.idnum, ref.generation), page_obj_to_num_map.get(ref.idnum, None))
        return None

    def _resolve_element_page_num(self, elem, page_obj_to_num_map, ancestor_pg: Optional[int] = None) -> Optional[int]:
        """
        Resolves page number for a structure element by checking:
        1. Direct /Pg on the element dictionary.
        2. /Pg inside /K dictionary (Marked Content Reference /MCR or Object Reference /OBJR).
        3. /Pg inside list of /K children.
        4. Inherited /Pg from enclosing ancestor structure elements in StructTree.
        """
        # 1. Direct /Pg on element
        pg = elem.get("/Pg")
        pg_num = self._get_page_num(pg, page_obj_to_num_map)
        if pg_num is not None:
            return pg_num

        # 2. Check /Pg inside /K
        k = _deref(elem.get("/K"))
        if isinstance(k, (dict, DictionaryObject)):
            k_pg = k.get("/Pg")
            pg_num = self._get_page_num(k_pg, page_obj_to_num_map)
            if pg_num is not None:
                return pg_num
        elif isinstance(k, (list, ArrayObject)):
            for item_raw in k:
                item = _deref(item_raw)
                if isinstance(item, (dict, DictionaryObject)):
                    k_pg = item.get("/Pg")
                    pg_num = self._get_page_num(k_pg, page_obj_to_num_map)
                    if pg_num is not None:
                        return pg_num

        # 3. Inherit from ancestor in StructTree
        return ancestor_pg

    def _extract_layout_bbox(self, elem) -> Optional[List[float]]:
        """
        Extracts authoritative Layout BBox [x0, y0, x1, y1] if specified in
        the element's /A attribute dictionary or direct /BBox key.
        """
        if "/A" in elem:
            a = _deref(elem["/A"])
            if isinstance(a, (dict, DictionaryObject)) and "/BBox" in a:
                bbox_raw = _deref(a["/BBox"])
                if isinstance(bbox_raw, (list, ArrayObject)) and len(bbox_raw) == 4:
                    try:
                        return [float(x) for x in bbox_raw]
                    except Exception:
                        pass
            elif isinstance(a, (list, ArrayObject)):
                for item_raw in a:
                    item = _deref(item_raw)
                    if isinstance(item, (dict, DictionaryObject)) and "/BBox" in item:
                        bbox_raw = _deref(item["/BBox"])
                        if isinstance(bbox_raw, (list, ArrayObject)) and len(bbox_raw) == 4:
                            try:
                                return [float(x) for x in bbox_raw]
                            except Exception:
                                pass
        if "/BBox" in elem:
            bbox_raw = _deref(elem["/BBox"])
            if isinstance(bbox_raw, (list, ArrayObject)) and len(bbox_raw) == 4:
                try:
                    return [float(x) for x in bbox_raw]
                except Exception:
                    pass
        return None

    def find_figure_tags(self, include_elem: bool = False, root_dict=None, page_map=None) -> List[Dict[str, Any]]:
        """
        Recursively traverse StructTreeRoot to find all elements where /S == '/Figure'.
        Only excludes a /Figure when it is semantically inside a /Table structure
        (e.g., /Table, /TR, /TH, /TD).
        Preserves all legitimate instructional figures nested inside lists (/L, /LI, /LBody),
        paragraphs (/P), headings (/H1-/H6), etc.
        """
        if root_dict is None:
            trailer = getattr(self.reader, "trailer", {}) or {}
            root = _deref(trailer.get("/Root") or getattr(self.reader, "root_object", {}))
        else:
            root = _deref(root_dict)

        if not root or "/StructTreeRoot" not in root:
            return []

        struct_root = _deref(root["/StructTreeRoot"])
        figures = []
        table_tags = {"/Table", "/TR", "/TH", "/TD", "Table", "TR", "TH", "TD"}
        current_page_map = page_map if page_map is not None else self.page_obj_to_num

        def walk(elem_raw, path="Root", ancestors=None, cur_pg=None):
            if ancestors is None:
                ancestors = []

            elem = _deref(elem_raw)
            if isinstance(elem, (dict, DictionaryObject)):
                s = elem.get("/S")
                s_str = str(s) if s is not None else ""
                
                # Check if current element defines /Pg to pass to descendants
                pg = elem.get("/Pg")
                pg_num_direct = self._get_page_num(pg, current_page_map)
                effective_pg = pg_num_direct if pg_num_direct is not None else cur_pg

                # Check if current element or any ancestor is a table tag
                in_table = any(a in table_tags for a in ancestors) or (s_str in table_tags)

                if s == "/Figure" or s_str == "/Figure":
                    if not in_table:
                        alt = elem.get("/Alt")
                        alt_str = decode_pdf_string(alt)
                        
                        pg_num = self._resolve_element_page_num(elem, current_page_map, effective_pg)
                        layout_bbox = self._extract_layout_bbox(elem)
                        
                        k = _deref(elem.get("/K"))
                        mcids = []
                        if isinstance(k, int):
                            mcids.append(k)
                        elif isinstance(k, (list, ArrayObject)):
                            for item_raw in k:
                                item = _deref(item_raw)
                                if isinstance(item, int):
                                    mcids.append(item)
                                elif isinstance(item, (dict, DictionaryObject)) and "/MCID" in item:
                                    mcids.append(int(item["/MCID"]))
                        elif isinstance(k, (dict, DictionaryObject)) and "/MCID" in k:
                            mcids.append(int(k["/MCID"]))

                        fig_entry = {
                            "figure_id": len(figures) + 1,
                            "path": path,
                            "page_number": pg_num,
                            "mcids": mcids,
                            "layout_bbox": layout_bbox,
                            "alt_text": alt_str,
                            "has_alt": bool(alt_str and alt_str.strip()),
                            "elem_keys": [str(k_name) for k_name in elem.keys()],
                            "title": decode_pdf_string(elem.get("/T")) if "/T" in elem else None
                        }
                        if include_elem:
                            fig_entry["_elem"] = elem
                        figures.append(fig_entry)

                current_ancestors = ancestors + [s_str]
                kids = _deref(elem.get("/K"))
                if isinstance(kids, (list, ArrayObject)):
                    for idx, kid in enumerate(kids):
                        walk(kid, f"{path}/K[{idx}]", current_ancestors, effective_pg)
                elif isinstance(kids, (dict, DictionaryObject)):
                    walk(kids, f"{path}/K", current_ancestors, effective_pg)

        walk(struct_root)
        return figures

    def inject_alt_texts(self, injections: Dict[int, str], output_pdf_path: str) -> int:
        """
        Injects alt texts into the PDF's StructTreeRoot elements where /S == '/Figure'.
        Consumes the exact same filtered figure list produced by find_figure_tags()
        to guarantee consistent ordering and figure_id alignment without re-walking independently.
        injections: dict mapping figure_id (1-indexed) -> alt_text string.
        Saves the modified accessible PDF to output_pdf_path.
        Returns the count of successfully injected figures.
        """
        writer = pypdf.PdfWriter(clone_from=self.pdf_path)
        root = _deref(writer.root_object)
        if not root or "/StructTreeRoot" not in root:
            return 0

        writer_page_map = {}
        for idx, page in enumerate(writer.pages):
            ref = page.indirect_reference
            if ref:
                writer_page_map[(ref.idnum, ref.generation)] = idx + 1
                writer_page_map[ref.idnum] = idx + 1

        figures = self.find_figure_tags(include_elem=True, root_dict=root, page_map=writer_page_map)
        injected_count = 0

        for fig in figures:
            fig_id = fig["figure_id"]
            if fig_id in injections:
                alt_to_inject = injections[fig_id]
                if alt_to_inject and str(alt_to_inject).strip().lower() not in {"duplicate", "duplicate.", "duplicate image", "dup", "dup."}:
                    fig["_elem"][NameObject("/Alt")] = TextStringObject(alt_to_inject.strip())
                    injected_count += 1

        with open(output_pdf_path, "wb") as f:
            writer.write(f)
        return injected_count

    def remove_alt_texts(self, figure_ids: Optional[List[int]], output_pdf_path: str) -> int:
        """
        Removes /Alt accessibility texts from the PDF's StructTreeRoot elements where /S == '/Figure'.
        If figure_ids is provided, removes /Alt from those specific figures.
        If figure_ids is None or empty, removes /Alt from all /Figure tags in the document.
        Also clears /ActualText, /E, and /A attribute dictionaries for complete cleanup.
        Saves the modified PDF to output_pdf_path.
        Returns the count of figures whose accessibility text was removed.
        """
        writer = pypdf.PdfWriter(clone_from=self.pdf_path)
        root = _deref(writer.root_object)
        if not root or "/StructTreeRoot" not in root:
            return 0

        writer_page_map = {}
        for idx, page in enumerate(writer.pages):
            ref = page.indirect_reference
            if ref:
                writer_page_map[(ref.idnum, ref.generation)] = idx + 1
                writer_page_map[ref.idnum] = idx + 1

        figures = self.find_figure_tags(include_elem=True, root_dict=root, page_map=writer_page_map)
        removed_count = 0
        target_ids = set(figure_ids) if figure_ids else None

        for fig in figures:
            fig_id = fig["figure_id"]
            if target_ids is None or fig_id in target_ids:
                elem = fig["_elem"]
                had_text = False

                # Remove standard accessibility text keys
                for key in ["/Alt", "/ActualText", "/E"]:
                    for k_variant in [key, NameObject(key)]:
                        if k_variant in elem:
                            del elem[k_variant]
                            had_text = True

                # Check and clean attribute dictionaries in /A
                if "/A" in elem:
                    a_val = _deref(elem["/A"])
                    if isinstance(a_val, (dict, DictionaryObject)):
                        for a_key in ["/Alt", "/ActualText"]:
                            for k_var in [a_key, NameObject(a_key)]:
                                if k_var in a_val:
                                    del a_val[k_var]
                                    had_text = True
                    elif isinstance(a_val, (list, ArrayObject)):
                        for item_raw in a_val:
                            item = _deref(item_raw)
                            if isinstance(item, (dict, DictionaryObject)):
                                for a_key in ["/Alt", "/ActualText"]:
                                    for k_var in [a_key, NameObject(a_key)]:
                                        if k_var in item:
                                            del item[k_var]
                                            had_text = True

                if had_text:
                    removed_count += 1

        with open(output_pdf_path, "wb") as f:
            writer.write(f)
        return removed_count

    def find_formula_tags(self, include_elem: bool = False, root_dict=None, page_map=None) -> List[Dict[str, Any]]:
        """
        Recursively traverse StructTreeRoot to find all mathematical formula elements
        where /S is /Formula, /Math, /MathType, etc.
        """
        if root_dict is None:
            trailer = getattr(self.reader, "trailer", {}) or {}
            root = _deref(trailer.get("/Root") or getattr(self.reader, "root_object", {}))
        else:
            root = _deref(root_dict)

        if not root or "/StructTreeRoot" not in root:
            return []

        struct_root = _deref(root["/StructTreeRoot"])
        formulas = []
        formula_tag_names = {"/Formula", "Formula", "/Math", "Math", "/MathType", "MathType"}
        current_page_map = page_map if page_map is not None else self.page_obj_to_num

        def walk(elem_raw, path="Root", ancestors=None, cur_pg=None):
            if ancestors is None:
                ancestors = []
            elem = _deref(elem_raw)
            if isinstance(elem, (dict, DictionaryObject)):
                s = elem.get("/S")
                s_str = str(s) if s is not None else ""

                pg = elem.get("/Pg")
                pg_num_direct = self._get_page_num(pg, current_page_map)
                effective_pg = pg_num_direct if pg_num_direct is not None else cur_pg

                if s_str in formula_tag_names or any(s_str == f"/{t}" for t in ["Formula", "Math", "MathType"]):
                    alt = elem.get("/Alt")
                    alt_str = decode_pdf_string(alt)
                    act = elem.get("/ActualText")
                    act_str = decode_pdf_string(act)

                    pg_num = self._resolve_element_page_num(elem, current_page_map, effective_pg)
                    layout_bbox = self._extract_layout_bbox(elem)

                    k = _deref(elem.get("/K"))
                    mcids = []
                    if isinstance(k, int):
                        mcids.append(k)
                    elif isinstance(k, (list, ArrayObject)):
                        for item_raw in k:
                            item = _deref(item_raw)
                            if isinstance(item, int):
                                mcids.append(item)
                            elif isinstance(item, (dict, DictionaryObject)) and "/MCID" in item:
                                mcids.append(int(item["/MCID"]))
                    elif isinstance(k, (dict, DictionaryObject)) and "/MCID" in k:
                        mcids.append(int(k["/MCID"]))

                    form_entry = {
                        "formula_id": len(formulas) + 1,
                        "path": path,
                        "page_number": pg_num,
                        "mcids": mcids,
                        "layout_bbox": layout_bbox,
                        "alt_text": alt_str,
                        "actual_text": act_str,
                        "has_alt": bool((alt_str and alt_str.strip()) or (act_str and act_str.strip())),
                        "elem_keys": [str(k_name) for k_name in elem.keys()],
                        "title": str(elem.get("/T")) if "/T" in elem else None
                    }
                    if include_elem:
                        form_entry["_elem"] = elem
                    formulas.append(form_entry)

                kids = _deref(elem.get("/K"))
                if isinstance(kids, (list, ArrayObject)):
                    for idx, kid in enumerate(kids):
                        walk(kid, f"{path}/K[{idx}]", ancestors + [s_str], effective_pg)
                elif isinstance(kids, (dict, DictionaryObject)):
                    walk(kids, f"{path}/K", ancestors + [s_str], effective_pg)

        walk(struct_root)
        return formulas

    def inject_formula_alt_texts(self, injections: Dict[int, str], output_pdf_path: str) -> int:
        """
        Injects alt texts into the PDF's StructTreeRoot elements where /S is a formula tag.
        injections: dict mapping formula_id (1-indexed) -> alt_text string.
        Saves the modified accessible PDF to output_pdf_path.
        Returns the count of successfully injected formulas.
        """
        writer = pypdf.PdfWriter(clone_from=self.pdf_path)
        root = _deref(writer.root_object)
        if not root or "/StructTreeRoot" not in root:
            return 0

        writer_page_map = {}
        for idx, page in enumerate(writer.pages):
            ref = page.indirect_reference
            if ref:
                writer_page_map[(ref.idnum, ref.generation)] = idx + 1
                writer_page_map[ref.idnum] = idx + 1

        formulas = self.find_formula_tags(include_elem=True, root_dict=root, page_map=writer_page_map)
        injected_count = 0

        for form in formulas:
            form_id = form["formula_id"]
            if form_id in injections:
                alt_to_inject = injections[form_id]
                if alt_to_inject and str(alt_to_inject).strip().lower() not in {"duplicate", "duplicate.", "duplicate image", "dup", "dup."}:
                    form["_elem"][NameObject("/Alt")] = TextStringObject(alt_to_inject.strip())
                    injected_count += 1

        with open(output_pdf_path, "wb") as f:
            writer.write(f)
        return injected_count

    def remove_formula_alt_texts(self, formula_ids: Optional[List[int]], output_pdf_path: str) -> int:
        """
        Removes /Alt and /ActualText from formula tags.
        """
        writer = pypdf.PdfWriter(clone_from=self.pdf_path)
        root = _deref(writer.root_object)
        if not root or "/StructTreeRoot" not in root:
            return 0

        writer_page_map = {}
        for idx, page in enumerate(writer.pages):
            ref = page.indirect_reference
            if ref:
                writer_page_map[(ref.idnum, ref.generation)] = idx + 1
                writer_page_map[ref.idnum] = idx + 1

        formulas = self.find_formula_tags(include_elem=True, root_dict=root, page_map=writer_page_map)
        removed_count = 0
        target_ids = set(formula_ids) if formula_ids else None

        for form in formulas:
            form_id = form["formula_id"]
            if target_ids is None or form_id in target_ids:
                elem = form["_elem"]
                had_text = False
                for key in ["/Alt", "/ActualText", "/E"]:
                    for k_var in [key, NameObject(key)]:
                        if k_var in elem:
                            del elem[k_var]
                            had_text = True
                if had_text:
                    removed_count += 1

        with open(output_pdf_path, "wb") as f:
            writer.write(f)
        return removed_count

    def parse_page_mcid_bboxes(self, page_index: int) -> Dict[int, List[float]]:
        """
        Parse the content stream of a page using a complete graphics state and text matrix interpreter.
        Computes exact bounding boxes [x0, y0, x1, y1] for all MCIDs (figures, formulas, etc.) in top-left coordinates.
        Handles text matrices (BT/ET, Tm, Td, T*), fonts (Tf), text strings (Tj, TJ), raster XObjects (Do),
        and vector path painting (m, l, re, c, paint operators f, S, B, etc.).
        """
        if page_index < 0 or page_index >= len(self.doc):
            return {}

        page = self.doc[page_index]
        cb = page.cropbox
        try:
            stream = page.read_contents().decode("latin1", errors="ignore")
        except Exception:
            return {}

        # Pre-fetch XObjects on page to resolve precise Form XObject bounding boxes
        xobj_map = {}
        try:
            for x in page.get_xobjects():
                xref, xname, parent_xref, rect = x
                if parent_xref == 0:
                    xobj_map[xname] = rect
        except Exception:
            pass

        init_clip = [cb.x0, cb.y0, cb.x1, cb.y1]
        gstate_stack = [{"ctm": np.eye(3), "clip": init_clip.copy()}]
        tm = np.eye(3)
        tlm = np.eye(3)
        font_size = 10.0
        current_path_pts = []
        mc_stack = []
        mcid_painted_boxes = {}

        token_pattern = re.compile(r'/[A-Za-z0-9_\.\-]+|<[0-9A-Fa-f\s]*>|\((?:\\.|[^)])*\)|\[(?:\\.|[^\]])*\]|[-+]?[0-9]*\.?[0-9]+(?:[eE][-+]?[0-9]+)?|[A-Za-z\*\']+')
        tokens = token_pattern.findall(stream)
        i = 0
        n = len(tokens)

        tc = 0.0
        tw = 0.0
        tz = 100.0
        ts = 0.0
        tl = 0.0

        def is_blank_or_space(tok_str: str) -> bool:
            clean = tok_str.strip("()<>[]").strip()
            return not clean or clean in ("0003", "0000", "20", "0020", "03", "0", " ")

        def _char_w(ch: str, current_fs: float) -> float:
            if ch in ' .,:;!\'"`|':
                return 0.28 * current_fs
            if ch in 'iljt1()[]{}':
                return 0.35 * current_fs
            if ch in 'mwMW@_':
                return 0.82 * current_fs
            return 0.55 * current_fs

        def _clean_pdf_string(s: str) -> str:
            if s.startswith('(') and s.endswith(')'):
                inner = s[1:-1]
                return re.sub(r'\\[0-7]{1,3}|\\[^0-7]', 'X', inner)
            if s.startswith('<') and s.endswith('>'):
                return s.strip('<> ')
            return s

        while i < n:
            tok = tokens[i]

            if tok == 'q':
                top = gstate_stack[-1]
                gstate_stack.append({"ctm": top["ctm"].copy(), "clip": top["clip"].copy()})
                current_path_pts.clear()
                i += 1
            elif tok == 'Q':
                if len(gstate_stack) > 1:
                    gstate_stack.pop()
                current_path_pts.clear()
                i += 1
            elif tok == 'cm':
                if i >= 6:
                    try:
                        vals = [float(tokens[i - 6 + j]) for j in range(6)]
                        m = np.array([[vals[0], vals[2], vals[4]],
                                      [vals[1], vals[3], vals[5]],
                                      [0, 0, 1]])
                        gstate_stack[-1]["ctm"] = gstate_stack[-1]["ctm"] @ m
                    except Exception:
                        pass
                i += 1
            elif tok == 'BDC':
                mcid = None
                if i >= 1:
                    dict_str = tokens[i - 1]
                    m_mcid = re.search(r'/MCID\s+(\d+)', dict_str)
                    if not m_mcid and i >= 3:
                        for lookback in range(1, min(6, i + 1)):
                            if tokens[i - lookback] == 'MCID' or '/MCID' in tokens[i - lookback]:
                                try:
                                    mcid = int(tokens[i - lookback + 1])
                                    break
                                except Exception:
                                    pass
                    elif m_mcid:
                        mcid = int(m_mcid.group(1))
                tag = tokens[i - 2].lstrip('/') if i >= 2 else 'Unknown'
                mc_stack.append((tag, mcid))
                current_path_pts.clear()
                i += 1
            elif tok == 'BMC':
                tag = tokens[i - 1].lstrip('/') if i >= 1 else 'Unknown'
                mc_stack.append((tag, None))
                current_path_pts.clear()
                i += 1
            elif tok == 'EMC':
                if mc_stack:
                    mc_stack.pop()
                current_path_pts.clear()
                i += 1
            elif tok == 'BT':
                tm = np.eye(3)
                tlm = np.eye(3)
                current_path_pts.clear()
                i += 1
            elif tok == 'ET':
                current_path_pts.clear()
                i += 1
            elif tok == 'Tf':
                if i >= 1:
                    try:
                        font_size = float(tokens[i - 1])
                    except Exception:
                        pass
                i += 1
            elif tok == 'Tc':
                if i >= 1:
                    try:
                        tc = float(tokens[i - 1])
                    except Exception:
                        pass
                i += 1
            elif tok == 'Tw':
                if i >= 1:
                    try:
                        tw = float(tokens[i - 1])
                    except Exception:
                        pass
                i += 1
            elif tok == 'Tz':
                if i >= 1:
                    try:
                        tz = float(tokens[i - 1])
                    except Exception:
                        pass
                i += 1
            elif tok == 'Ts':
                if i >= 1:
                    try:
                        ts = float(tokens[i - 1])
                    except Exception:
                        pass
                i += 1
            elif tok == 'TL':
                if i >= 1:
                    try:
                        tl = float(tokens[i - 1])
                    except Exception:
                        pass
                i += 1
            elif tok == 'Tm':
                if i >= 6:
                    try:
                        vals = [float(tokens[i - 6 + j]) for j in range(6)]
                        tm = np.array([[vals[0], vals[2], vals[4]],
                                       [vals[1], vals[3], vals[5]],
                                       [0, 0, 1]])
                        tlm = tm.copy()
                    except Exception:
                        pass
                i += 1
            elif tok in ('Td', 'TD'):
                if i >= 2:
                    try:
                        tx, ty = float(tokens[i - 2]), float(tokens[i - 1])
                        if tok == 'TD':
                            tl = -ty
                        t_trans = np.array([[1, 0, tx], [0, 1, ty], [0, 0, 1]])
                        tlm = tlm @ t_trans
                        tm = tlm.copy()
                    except Exception:
                        pass
                i += 1
            elif tok == 'T*':
                leading = tl if tl != 0.0 else font_size
                t_trans = np.array([[1, 0, 0], [0, 1, -leading], [0, 0, 1]])
                tlm = tlm @ t_trans
                tm = tlm.copy()
                i += 1
            elif tok in ('Tj', 'TJ', "'", '"'):
                fs = font_size if font_size > 0.001 else 10.0
                h_scale = (tz / 100.0) if tz > 0 else 1.0

                if tok == "'":
                    leading = tl if tl != 0.0 else fs
                    t_trans = np.array([[1, 0, 0], [0, 1, -leading], [0, 0, 1]])
                    tlm = tlm @ t_trans
                    tm = tlm.copy()
                elif tok == '"' and i >= 3:
                    try:
                        tw = float(tokens[i - 3])
                        tc = float(tokens[i - 2])
                    except Exception:
                        pass
                    leading = tl if tl != 0.0 else fs
                    t_trans = np.array([[1, 0, 0], [0, 1, -leading], [0, 0, 1]])
                    tlm = tlm @ t_trans
                    tm = tlm.copy()

                eff_m = gstate_stack[-1]["ctm"] @ tm
                act = gstate_stack[-1]["clip"]
                x_current = 0.0

                if tok == 'TJ' and i >= 1:
                    raw_arr = tokens[i - 1]
                    item_pat = re.compile(r'\((?:\\.|[^)])*\)|<[0-9A-Fa-f\s]*>|[-+]?[0-9]*\.?[0-9]+(?:[eE][-+]?[0-9]+)?')
                    content = raw_arr[1:-1] if raw_arr.startswith('[') and raw_arr.endswith(']') else raw_arr
                    items = item_pat.findall(content)
                    for item in items:
                        if item.startswith('(') or item.startswith('<'):
                            if not is_blank_or_space(item):
                                clean = _clean_pdf_string(item)
                                if item.startswith('<'):
                                    num_chars = max(1, len(clean) // 2)
                                    w_item = num_chars * 0.55 * fs
                                else:
                                    w_item = sum(_char_w(c, fs) for c in clean) + len(clean) * tc
                                w_item *= h_scale

                                p1 = eff_m @ np.array([x_current, -0.22 * fs + ts, 1.0])
                                p2 = eff_m @ np.array([x_current + w_item, -0.22 * fs + ts, 1.0])
                                p3 = eff_m @ np.array([x_current + w_item, 0.85 * fs + ts, 1.0])
                                p4 = eff_m @ np.array([x_current, 0.85 * fs + ts, 1.0])
                                pts = [p1, p2, p3, p4]

                                px0 = min(p[0] for p in pts)
                                py0 = min(p[1] for p in pts)
                                px1 = max(p[0] for p in pts)
                                py1 = max(p[1] for p in pts)

                                if px1 - px0 < page.rect.width * 0.95 and py1 - py0 < page.rect.height * 0.95:
                                    ix0 = max(px0, act[0])
                                    iy0 = max(py0, act[1])
                                    ix1 = min(px1, act[2])
                                    iy1 = min(py1, act[3])
                                    if ix1 > ix0 and iy1 > iy0:
                                        for mc in mc_stack:
                                            if mc[1] is not None:
                                                mcid_painted_boxes.setdefault(mc[1], []).append([ix0, iy0, ix1, iy1])
                                x_current += w_item
                            else:
                                x_current += (0.30 * fs + tw) * h_scale
                        else:
                            try:
                                adj = float(item)
                                x_current -= (adj / 1000.0) * fs * h_scale
                            except Exception:
                                pass
                elif tok in ('Tj', "'", '"') and i >= 1:
                    raw_str = tokens[i - 1]
                    if not is_blank_or_space(raw_str):
                        clean = _clean_pdf_string(raw_str)
                        if raw_str.startswith('<'):
                            num_chars = max(1, len(clean) // 2)
                            w_item = num_chars * 0.55 * fs
                        else:
                            w_item = sum(_char_w(c, fs) for c in clean) + len(clean) * tc
                        w_item *= h_scale

                        p1 = eff_m @ np.array([0.0, -0.22 * fs + ts, 1.0])
                        p2 = eff_m @ np.array([w_item, -0.22 * fs + ts, 1.0])
                        p3 = eff_m @ np.array([w_item, 0.85 * fs + ts, 1.0])
                        p4 = eff_m @ np.array([0.0, 0.85 * fs + ts, 1.0])
                        pts = [p1, p2, p3, p4]

                        px0 = min(p[0] for p in pts)
                        py0 = min(p[1] for p in pts)
                        px1 = max(p[0] for p in pts)
                        py1 = max(p[1] for p in pts)

                        if px1 - px0 < page.rect.width * 0.95 and py1 - py0 < page.rect.height * 0.95:
                            ix0 = max(px0, act[0])
                            iy0 = max(py0, act[1])
                            ix1 = min(px1, act[2])
                            iy1 = min(py1, act[3])
                            if ix1 > ix0 and iy1 > iy0:
                                for mc in mc_stack:
                                    if mc[1] is not None:
                                        mcid_painted_boxes.setdefault(mc[1], []).append([ix0, iy0, ix1, iy1])
                        x_current += w_item
                    else:
                        x_current += (0.30 * fs + tw) * h_scale

                if x_current != 0.0:
                    tm = tm @ np.array([[1.0, 0.0, x_current], [0.0, 1.0, 0.0], [0.0, 0.0, 1.0]])

                i += 1
            elif tok == 'Do':
                act = gstate_stack[-1]["clip"]
                xname = tokens[i - 1].lstrip('/') if i >= 1 else None
                used_box = False
                if xname and xname in xobj_map:
                    rect = xobj_map[xname]
                    if rect and rect[2] > rect[0] and rect[3] > rect[1]:
                        if not (rect[2] - rect[0] >= page.rect.width * 0.95 and rect[3] - rect[1] >= page.rect.height * 0.95):
                            ix0 = max(rect[0], act[0])
                            iy0 = max(rect[1], act[1])
                            ix1 = min(rect[2], act[2])
                            iy1 = min(rect[3], act[3])
                            if ix1 > ix0 and iy1 > iy0:
                                for mc in mc_stack:
                                    if mc[1] is not None:
                                        mcid_painted_boxes.setdefault(mc[1], []).append([ix0, iy0, ix1, iy1])
                                used_box = True
                if not used_box:
                    corners = []
                    for corner in [(0.0, 0.0), (1.0, 0.0), (1.0, 1.0), (0.0, 1.0)]:
                        pt = gstate_stack[-1]["ctm"] @ np.array([corner[0], corner[1], 1.0])
                        corners.append((pt[0], pt[1]))
                    ib = [min(p[0] for p in corners), min(p[1] for p in corners),
                          max(p[0] for p in corners), max(p[1] for p in corners)]
                    if ib[2] - ib[0] < page.rect.width * 0.95 or ib[3] - ib[1] < page.rect.height * 0.95:
                        ix0 = max(ib[0], act[0])
                        iy0 = max(ib[1], act[1])
                        ix1 = min(ib[2], act[2])
                        iy1 = min(ib[3], act[3])
                        if ix1 > ix0 and iy1 > iy0:
                            for mc in mc_stack:
                                if mc[1] is not None:
                                    mcid_painted_boxes.setdefault(mc[1], []).append([ix0, iy0, ix1, iy1])
                i += 1
            elif tok in ('m', 'l'):
                if i >= 2:
                    try:
                        x, y = float(tokens[i - 2]), float(tokens[i - 1])
                        pt = gstate_stack[-1]["ctm"] @ np.array([x, y, 1.0])
                        current_path_pts.append((pt[0], pt[1]))
                    except Exception:
                        pass
                i += 1
            elif tok == 're':
                if i >= 4:
                    try:
                        x, y, w, hbox = float(tokens[i - 4]), float(tokens[i - 3]), float(tokens[i - 2]), float(tokens[i - 1])
                        for corner in [(x, y), (x + w, y), (x + w, y + hbox), (x, y + hbox)]:
                            pt = gstate_stack[-1]["ctm"] @ np.array([corner[0], corner[1], 1.0])
                            current_path_pts.append((pt[0], pt[1]))
                    except Exception:
                        pass
                i += 1
            elif tok == 'c':
                if i >= 6:
                    try:
                        for k in [2, 4, 6]:
                            x, y = float(tokens[i - k]), float(tokens[i - k + 1])
                            pt = gstate_stack[-1]["ctm"] @ np.array([x, y, 1.0])
                            current_path_pts.append((pt[0], pt[1]))
                    except Exception:
                        pass
                i += 1
            elif tok in PAINT_OPERATORS:
                if current_path_pts:
                    px0 = min(p[0] for p in current_path_pts)
                    py0 = min(p[1] for p in current_path_pts)
                    px1 = max(p[0] for p in current_path_pts)
                    py1 = max(p[1] for p in current_path_pts)
                    if not (px1 - px0 >= page.rect.width * 0.95 and py1 - py0 >= page.rect.height * 0.95):
                        act = gstate_stack[-1]["clip"]
                        ix0 = max(px0, act[0])
                        iy0 = max(py0, act[1])
                        ix1 = min(px1, act[2])
                        iy1 = min(py1, act[3])
                        if ix1 > ix0 and iy1 > iy0:
                            for mc in mc_stack:
                                if mc[1] is not None:
                                    mcid_painted_boxes.setdefault(mc[1], []).append([ix0, iy0, ix1, iy1])
                    current_path_pts.clear()
                i += 1
            elif tok in ('n', 'W', 'W*'):
                current_path_pts.clear()
                i += 1
            else:
                i += 1

        final_bboxes = {}
        for mcid, boxes in mcid_painted_boxes.items():
            if boxes:
                valid_boxes = [b for b in boxes if not ((b[2]-b[0] >= page.rect.width * 0.85) and (b[3]-b[1] >= page.rect.height * 0.85))]
                small_boxes = [b for b in valid_boxes if (b[3] - b[1]) < page.rect.height * 0.65]
                use_boxes = small_boxes if small_boxes else (valid_boxes if valid_boxes else boxes)

                min_x = min(b[0] for b in use_boxes)
                min_y = min(b[1] for b in use_boxes)
                max_x = max(b[2] for b in use_boxes)
                max_y = max(b[3] for b in use_boxes)

                x0 = max(0.0, min_x - cb.x0)
                x1 = min(page.rect.width, max_x - cb.x0)
                y0 = max(0.0, cb.y1 - max_y)
                y1 = min(page.rect.height, cb.y1 - min_y)
                if x1 > x0 and y1 > y0:
                    final_bboxes[mcid] = [float(round(x0, 2)), float(round(y0, 2)), float(round(x1, 2)), float(round(y1, 2))]

        return final_bboxes

    def extract_figures_with_crops(self, output_dir: str, dpi: int = 150) -> List[Dict[str, Any]]:
        """
        Extracts all valid (non-table) /Figure tags, computes their bounding boxes, renders the high-res crops,
        and saves them as PNG files into output_dir.
        If a figure's bbox cannot be resolved, no crop is generated and bbox is set to None.
        """
        os.makedirs(output_dir, exist_ok=True)
        figures = self.find_figure_tags()
        
        page_renders = {}
        page_mcid_bboxes = {}

        results = []
        for fig in figures:
            pg_num = fig["page_number"]
            fig_id = fig["figure_id"]
            bbox = None

            if pg_num and 1 <= pg_num <= self.page_count:
                pg_idx = pg_num - 1
                if pg_idx not in page_mcid_bboxes:
                    page_mcid_bboxes[pg_idx] = self.parse_page_mcid_bboxes(pg_idx)
                
                bboxes_on_page = page_mcid_bboxes[pg_idx]
                
                matching_boxes = []
                for mcid in fig["mcids"]:
                    if mcid in bboxes_on_page:
                        matching_boxes.append(bboxes_on_page[mcid])
                
                if matching_boxes:
                    bbox = [
                        min(b[0] for b in matching_boxes),
                        min(b[1] for b in matching_boxes),
                        max(b[2] for b in matching_boxes),
                        max(b[3] for b in matching_boxes)
                    ]
                elif fig.get("layout_bbox"):
                    lb = fig["layout_bbox"]
                    page = self.doc[pg_idx]
                    rx0 = max(0.0, min(lb[0], lb[2]))
                    rx1 = min(page.rect.width, max(lb[0], lb[2]))
                    ry0 = max(0.0, page.rect.height - max(lb[1], lb[3]))
                    ry1 = min(page.rect.height, page.rect.height - min(lb[1], lb[3]))
                    if rx1 > rx0 and ry1 > ry0:
                        bbox = [float(round(rx0, 2)), float(round(ry0, 2)), float(round(rx1, 2)), float(round(ry1, 2))]

                page = self.doc[pg_idx]
                img_list = page.get_images(full=True)

                if bbox:
                    if pg_idx not in page_renders:
                        pix = page.get_pixmap(dpi=dpi)
                        page_renders[pg_idx] = Image.open(io.BytesIO(pix.tobytes("png")))

                    page_img = page_renders[pg_idx]
                    scale = dpi / 72.0

                    crop_filename = f"figure_{fig_id:03d}_page_{pg_num}.png"
                    crop_path = os.path.join(output_dir, crop_filename)

                    pad = 3.0
                    x0 = max(0, int((bbox[0] - pad) * scale))
                    y0 = max(0, int((bbox[1] - pad) * scale))
                    x1 = min(page_img.width, int((bbox[2] + pad) * scale))
                    y1 = min(page_img.height, int((bbox[3] + pad) * scale))
                    
                    if x1 > x0 + 10 and y1 > y0 + 10:
                        crop_img = page_img.crop((x0, y0, x1, y1))
                        crop_img.save(crop_path, "PNG")
                    else:
                        crop_img = page_img.crop((0, y0, page_img.width, min(page_img.height, y0 + int(300 * scale))))
                        crop_img.save(crop_path, "PNG")

                    bbox_text = ""
                    try:
                        rect = pymupdf.Rect(bbox[0], bbox[1], bbox[2], bbox[3])
                        bbox_text = self.doc[pg_idx].get_text("text", clip=rect).strip()
                    except Exception:
                        pass

                    fig_info = {
                        **fig,
                        "bbox": bbox,
                        "bbox_width": float(round(bbox[2] - bbox[0], 2)),
                        "bbox_height": float(round(bbox[3] - bbox[1], 2)),
                        "bbox_text": bbox_text,
                        "crop_filename": crop_filename,
                        "crop_path": crop_path,
                        "underlying_xobjects_count": len(img_list),
                        "status": "Has Alt" if fig["has_alt"] else "Missing Alt"
                    }
                    results.append(fig_info)
                else:
                    fig_info = {
                        **fig,
                        "bbox": None,
                        "bbox_width": 0.0,
                        "bbox_height": 0.0,
                        "bbox_text": "",
                        "crop_filename": None,
                        "crop_path": None,
                        "underlying_xobjects_count": len(img_list),
                        "status": "Has Alt" if fig["has_alt"] else "Missing Alt",
                        "crop_status": "no crop available"
                    }
                    results.append(fig_info)
            else:
                fig_info = {
                    **fig,
                    "bbox": None,
                    "bbox_width": 0.0,
                    "bbox_height": 0.0,
                    "bbox_text": "",
                    "crop_filename": None,
                    "crop_path": None,
                    "underlying_xobjects_count": 0,
                    "status": "Missing Alt",
                    "crop_status": "no crop available"
                }
                results.append(fig_info)

        return results

    def extract_formulas_with_crops(self, output_dir: str, dpi: int = 150) -> List[Dict[str, Any]]:
        """
        Extracts all mathematical /Formula tags, computes their bounding boxes, renders the high-res crops,
        and saves them as PNG files into output_dir.
        """
        os.makedirs(output_dir, exist_ok=True)
        formulas = self.find_formula_tags()
        
        page_renders = {}
        page_mcid_bboxes = {}

        results = []
        for form in formulas:
            pg_num = form["page_number"]
            form_id = form["formula_id"]
            bbox = None

            if pg_num and 1 <= pg_num <= self.page_count:
                pg_idx = pg_num - 1
                if pg_idx not in page_mcid_bboxes:
                    page_mcid_bboxes[pg_idx] = self.parse_page_mcid_bboxes(pg_idx)
                
                bboxes_on_page = page_mcid_bboxes[pg_idx]
                
                matching_boxes = []
                for mcid in form["mcids"]:
                    if mcid in bboxes_on_page:
                        matching_boxes.append(bboxes_on_page[mcid])
                
                if matching_boxes:
                    valid_matching = [b for b in matching_boxes if not ((b[2]-b[0] >= self.doc[pg_idx].rect.width * 0.85) and (b[3]-b[1] >= self.doc[pg_idx].rect.height * 0.85))]
                    matching_to_use = valid_matching if valid_matching else []
                    if matching_to_use:
                        bbox = [
                            min(b[0] for b in matching_to_use),
                            min(b[1] for b in matching_to_use),
                            max(b[2] for b in matching_to_use),
                            max(b[3] for b in matching_to_use)
                        ]
                    else:
                        bbox = None
                elif form.get("layout_bbox"):
                    lb = form["layout_bbox"]
                    page = self.doc[pg_idx]
                    rx0 = max(0.0, min(lb[0], lb[2]))
                    rx1 = min(page.rect.width, max(lb[0], lb[2]))
                    ry0 = max(0.0, page.rect.height - max(lb[1], lb[3]))
                    ry1 = min(page.rect.height, page.rect.height - min(lb[1], lb[3]))
                    if rx1 > rx0 and ry1 > ry0:
                        bbox = [float(round(rx0, 2)), float(round(ry0, 2)), float(round(rx1, 2)), float(round(ry1, 2))]

                if bbox and not (bbox[2] - bbox[0] >= self.doc[pg_idx].rect.width * 0.85 and bbox[3] - bbox[1] >= self.doc[pg_idx].rect.height * 0.85):
                    if pg_idx not in page_renders:
                        pix = self.doc[pg_idx].get_pixmap(dpi=dpi)
                        page_renders[pg_idx] = Image.open(io.BytesIO(pix.tobytes("png")))

                    page_img = page_renders[pg_idx]
                    scale = dpi / 72.0

                    crop_filename = f"formula_{form_id:04d}_page_{pg_num}.png"
                    crop_path = os.path.join(output_dir, crop_filename)

                    pad_x = 2.0
                    pad_y = 1.0
                    x0 = max(0, int((bbox[0] - pad_x) * scale))
                    y0 = max(0, int((bbox[1] - pad_y) * scale))
                    x1 = min(page_img.width, int((bbox[2] + pad_x) * scale))
                    y1 = min(page_img.height, int((bbox[3] + pad_y) * scale))
                    
                    if x1 > x0 + 4 and y1 > y0 + 4:
                        crop_img = page_img.crop((x0, y0, x1, y1))
                        crop_img.save(crop_path, "PNG")
                    else:
                        crop_filename = None
                        crop_path = None

                    bbox_text = ""
                    try:
                        rect = pymupdf.Rect(bbox[0], bbox[1], bbox[2], bbox[3])
                        bbox_text = self.doc[pg_idx].get_text("text", clip=rect).strip()
                    except Exception:
                        pass

                    form_info = {
                        **form,
                        "bbox": bbox,
                        "bbox_width": float(round(bbox[2] - bbox[0], 2)),
                        "bbox_height": float(round(bbox[3] - bbox[1], 2)),
                        "bbox_text": bbox_text,
                        "crop_filename": crop_filename,
                        "crop_path": crop_path,
                        "status": "Has Alt" if form["has_alt"] else "Missing Alt"
                    }
                    results.append(form_info)
                else:
                    form_info = {
                        **form,
                        "bbox": None,
                        "bbox_width": 0.0,
                        "bbox_height": 0.0,
                        "bbox_text": "",
                        "crop_filename": None,
                        "crop_path": None,
                        "status": "Has Alt" if form["has_alt"] else "Missing Alt",
                        "crop_status": "no crop available"
                    }
                    results.append(form_info)
            else:
                form_info = {
                    **form,
                    "bbox": None,
                    "bbox_width": 0.0,
                    "bbox_height": 0.0,
                    "bbox_text": "",
                    "crop_filename": None,
                    "crop_path": None,
                    "status": "Missing Alt",
                    "crop_status": "no crop available"
                }
                results.append(form_info)

        return results

    def close(self):
        self.doc.close()

