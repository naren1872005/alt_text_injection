import io
import os
import re
import numpy as np
import pikepdf
import pymupdf
from PIL import Image
from typing import List, Dict, Any, Optional, Tuple

PAINT_OPERATORS = {
    "f", "F", "f*",
    "S", "s",
    "B", "B*",
    "b", "b*"
}

class FigureExtractor:
    """
    Extracts /Figure tags from a PDF's accessibility structure (StructTreeRoot),
    resolves their page numbers, Marked Content Identifiers (MCID), bounding boxes,
    and associated /XObject /Image resources.
    """

    def __init__(self, pdf_path: str):
        self.pdf_path = pdf_path
        self.doc = pymupdf.open(pdf_path)
        self.pdoc = pikepdf.Pdf.open(pdf_path)
        self.page_count = len(self.doc)
        
        # Map pikepdf page object identities to 1-indexed page numbers
        self.page_obj_to_num = {}
        for idx, page_obj in enumerate(self.pdoc.pages):
            self.page_obj_to_num[page_obj.objgen] = idx + 1

    def find_figure_tags(self, include_elem: bool = False) -> List[Dict[str, Any]]:
        """
        Recursively traverse StructTreeRoot to find all elements where /S == '/Figure'.
        Only excludes a /Figure when it is semantically inside a /Table structure
        (e.g., /Table, /TR, /TH, /TD).
        Preserves all legitimate instructional figures nested inside lists (/L, /LI, /LBody),
        paragraphs (/P), headings (/H1-/H6), etc.
        """
        if "/StructTreeRoot" not in self.pdoc.Root:
            return []

        struct_root = self.pdoc.Root.StructTreeRoot
        figures = []
        table_tags = {"/Table", "/TR", "/TH", "/TD", "Table", "TR", "TH", "TD"}

        def walk(elem, path="Root", ancestors=None):
            if ancestors is None:
                ancestors = []

            if isinstance(elem, pikepdf.Dictionary):
                s = elem.get("/S")
                s_str = str(s) if s is not None else ""
                
                # Check if current element or any ancestor is a table tag
                in_table = any(a in table_tags for a in ancestors) or (s_str in table_tags)

                if s == "/Figure":
                    if not in_table:
                        alt = elem.get("/Alt")
                        alt_str = str(alt) if alt is not None else None
                        
                        pg = elem.get("/Pg")
                        pg_num = None
                        if pg:
                            pg_num = self.page_obj_to_num.get(pg.objgen, None)
                        
                        k = elem.get("/K")
                        mcids = []
                        if isinstance(k, int):
                            mcids.append(k)
                        elif isinstance(k, pikepdf.Array):
                            for item in k:
                                if isinstance(item, int):
                                    mcids.append(item)
                                elif isinstance(item, pikepdf.Dictionary) and "/MCID" in item:
                                    mcids.append(int(item["/MCID"]))
                        elif isinstance(k, pikepdf.Dictionary) and "/MCID" in k:
                            mcids.append(int(k["/MCID"]))

                        fig_entry = {
                            "figure_id": len(figures) + 1,
                            "path": path,
                            "page_number": pg_num,
                            "mcids": mcids,
                            "alt_text": alt_str,
                            "has_alt": bool(alt_str and alt_str.strip()),
                            "elem_keys": [str(k_name) for k_name in elem.keys()],
                            "title": str(elem.get("/T")) if "/T" in elem else None
                        }
                        if include_elem:
                            fig_entry["_elem"] = elem
                        figures.append(fig_entry)

                current_ancestors = ancestors + [s_str]
                kids = elem.get("/K")
                if isinstance(kids, pikepdf.Array):
                    for idx, kid in enumerate(kids):
                        walk(kid, f"{path}/K[{idx}]", current_ancestors)
                elif isinstance(kids, pikepdf.Dictionary):
                    walk(kids, f"{path}/K", current_ancestors)

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
        if "/StructTreeRoot" not in self.pdoc.Root:
            return 0

        figures = self.find_figure_tags(include_elem=True)
        injected_count = 0

        for fig in figures:
            fig_id = fig["figure_id"]
            if fig_id in injections:
                alt_to_inject = injections[fig_id]
                if alt_to_inject:
                    fig["_elem"]["/Alt"] = pikepdf.String(alt_to_inject.strip())
                    injected_count += 1

        self.pdoc.save(output_pdf_path)
        return injected_count

    def remove_alt_texts(self, figure_ids: Optional[List[int]], output_pdf_path: str) -> int:
        """
        Removes /Alt accessibility texts from the PDF's StructTreeRoot elements where /S == '/Figure'.
        If figure_ids is provided, removes /Alt from those specific figures.
        If figure_ids is None or empty, removes /Alt from all /Figure tags in the document.
        Saves the modified PDF to output_pdf_path.
        Returns the count of figures whose /Alt text was removed.
        """
        if "/StructTreeRoot" not in self.pdoc.Root:
            return 0

        figures = self.find_figure_tags(include_elem=True)
        removed_count = 0
        target_ids = set(figure_ids) if figure_ids else None

        for fig in figures:
            fig_id = fig["figure_id"]
            if target_ids is None or fig_id in target_ids:
                elem = fig["_elem"]
                if "/Alt" in elem:
                    del elem["/Alt"]
                    removed_count += 1

        self.pdoc.save(output_pdf_path)
        return removed_count

    def parse_page_mcid_bboxes(self, page_index: int) -> Dict[int, List[float]]:
        """
        Parse the content stream of a page using a graphics state stack (CTM)
        to compute the exact bounding box [x0, y0, x1, y1] for each MCID in top-left coords.
        Supports both vector paths (m, l, re, c) and raster images invoked via Do.
        """
        if page_index < 0 or page_index >= len(self.doc):
            return {}

        page = self.doc[page_index]
        h = page.rect.height
        try:
            stream = page.read_contents().decode("latin1", errors="ignore")
        except Exception:
            return {}

        pattern = re.compile(r'/Figure\s*<<\s*/MCID\s+(\d+)\s*>>\s*BDC(.*?)EMC', re.DOTALL)
        bboxes = {}

        cb = page.cropbox
        init_clip = [cb.x0, cb.y0, cb.x1, cb.y1]

        for match in pattern.finditer(stream):
            mcid = int(match.group(1))
            content = match.group(2)

            clip_stack = [init_clip.copy()]
            ctm_stack = [np.eye(3)]
            current_path_pts = []
            painted_boxes = []
            clipping_boxes = []

            tokens = content.split()
            i = 0
            n = len(tokens)
            while i < n:
                tok = tokens[i]
                if tok == 'q':
                    ctm_stack.append(ctm_stack[-1].copy())
                    clip_stack.append(clip_stack[-1].copy())
                    i += 1
                elif tok == 'Q':
                    if len(ctm_stack) > 1:
                        ctm_stack.pop()
                    if len(clip_stack) > 1:
                        clip_stack.pop()
                    i += 1
                elif tok == 'cm':
                    if i >= 6:
                        try:
                            vals = [float(tokens[i - 6 + j]) for j in range(6)]
                            m = np.array([[vals[0], vals[2], vals[4]],
                                          [vals[1], vals[3], vals[5]],
                                          [0, 0, 1]])
                            ctm_stack[-1] = ctm_stack[-1] @ m
                        except Exception:
                            pass
                    i += 1
                elif tok == 'Do':
                    # Image unit square (0,0), (1,0), (1,1), (0,1) transformed through current CTM
                    act = clip_stack[-1]
                    corners = []
                    for corner in [(0.0, 0.0), (1.0, 0.0), (1.0, 1.0), (0.0, 1.0)]:
                        pt = ctm_stack[-1] @ np.array([corner[0], corner[1], 1.0])
                        corners.append((pt[0], pt[1]))
                    ib = [min(p[0] for p in corners), min(p[1] for p in corners),
                          max(p[0] for p in corners), max(p[1] for p in corners)]
                    ix0 = max(ib[0], act[0])
                    iy0 = max(ib[1], act[1])
                    ix1 = min(ib[2], act[2])
                    iy1 = min(ib[3], act[3])
                    if ix1 > ix0 and iy1 > iy0:
                        painted_boxes.append([ix0, iy0, ix1, iy1])
                    i += 1
                elif tok in ('m', 'l'):
                    if i >= 2:
                        try:
                            x, y = float(tokens[i - 2]), float(tokens[i - 1])
                            pt = ctm_stack[-1] @ np.array([x, y, 1.0])
                            current_path_pts.append((pt[0], pt[1]))
                        except Exception:
                            pass
                    i += 1
                elif tok == 're':
                    if i >= 4:
                        try:
                            x, y, w, hbox = float(tokens[i - 4]), float(tokens[i - 3]), float(tokens[i - 2]), float(tokens[i - 1])
                            for corner in [(x, y), (x + w, y), (x, y + hbox), (x + w, y + hbox)]:
                                pt = ctm_stack[-1] @ np.array([corner[0], corner[1], 1.0])
                                current_path_pts.append((pt[0], pt[1]))
                        except Exception:
                            pass
                    i += 1
                elif tok == 'c':
                    if i >= 6:
                        try:
                            for k in [2, 4, 6]:
                                x, y = float(tokens[i - k]), float(tokens[i - k + 1])
                                pt = ctm_stack[-1] @ np.array([x, y, 1.0])
                                current_path_pts.append((pt[0], pt[1]))
                        except Exception:
                            pass
                    i += 1
                elif tok in ('W', 'W*'):
                    # Intersect active clipping box with current path bounding box
                    if current_path_pts:
                        px0 = min(p[0] for p in current_path_pts)
                        py0 = min(p[1] for p in current_path_pts)
                        px1 = max(p[0] for p in current_path_pts)
                        py1 = max(p[1] for p in current_path_pts)
                        act = clip_stack[-1]
                        clip_stack[-1] = [max(act[0], px0), max(act[1], py0), min(act[2], px1), min(act[3], py1)]
                    i += 1
                elif tok in PAINT_OPERATORS:
                    # Visible artwork: intersect path with active clipping box
                    if current_path_pts:
                        px0 = min(p[0] for p in current_path_pts)
                        py0 = min(p[1] for p in current_path_pts)
                        px1 = max(p[0] for p in current_path_pts)
                        py1 = max(p[1] for p in current_path_pts)
                        act = clip_stack[-1]
                        ix0 = max(px0, act[0])
                        iy0 = max(py0, act[1])
                        ix1 = min(px1, act[2])
                        iy1 = min(py1, act[3])
                        if ix1 > ix0 and iy1 > iy0:
                            painted_boxes.append([ix0, iy0, ix1, iy1])
                        current_path_pts.clear()
                    i += 1
                elif tok == 'n':
                    # Clipping-only path: retain fallback bounds and clear current path
                    if current_path_pts:
                        px0 = min(p[0] for p in current_path_pts)
                        py0 = min(p[1] for p in current_path_pts)
                        px1 = max(p[0] for p in current_path_pts)
                        py1 = max(p[1] for p in current_path_pts)
                        clipping_boxes.append([px0, py0, px1, py1])
                        current_path_pts.clear()
                    i += 1
                else:
                    i += 1

            boxes_to_use = painted_boxes if painted_boxes else clipping_boxes
            if boxes_to_use:
                min_x = min(b[0] for b in boxes_to_use)
                min_y = min(b[1] for b in boxes_to_use)
                max_x = max(b[2] for b in boxes_to_use)
                max_y = max(b[3] for b in boxes_to_use)

                # Page rendering origin relative to CropBox
                x0 = max(0.0, min_x - cb.x0)
                x1 = min(page.rect.width, max_x - cb.x0)
                y0 = max(0.0, cb.y1 - max_y)
                y1 = min(page.rect.height, cb.y1 - min_y)
                if x1 > x0 and y1 > y0:
                    bboxes[mcid] = [float(round(x0, 2)), float(round(y0, 2)), float(round(x1, 2)), float(round(y1, 2))]

        return bboxes

    def extract_figures_with_crops(self, output_dir: str, dpi: int = 150) -> List[Dict[str, Any]]:
        """
        Extracts all valid (non-table) /Figure tags, computes their bounding boxes, renders the high-res crops,
        and saves them as PNG files into output_dir.
        If a figure's bbox cannot be resolved, no crop is generated and bbox is set to None.
        """
        os.makedirs(output_dir, exist_ok=True)
        figures = self.find_figure_tags()
        
        # Cache page renders and parsed MCID bboxes
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
                
                # Check for matching MCID
                for mcid in fig["mcids"]:
                    if mcid in bboxes_on_page:
                        bbox = bboxes_on_page[mcid]
                        break

                page = self.doc[pg_idx]
                img_list = page.get_images(full=True)

                if bbox:
                    # Render page if needed
                    if pg_idx not in page_renders:
                        pix = page.get_pixmap(dpi=dpi)
                        page_renders[pg_idx] = Image.open(io.BytesIO(pix.tobytes("png")))

                    page_img = page_renders[pg_idx]
                    scale = dpi / 72.0

                    crop_filename = f"figure_{fig_id:03d}_page_{pg_num}.png"
                    crop_path = os.path.join(output_dir, crop_filename)

                    # Add a 3pt padding around diagram (avoids bleeding into adjacent text headers)
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

                    fig_info = {
                        **fig,
                        "bbox": bbox,
                        "bbox_width": float(round(bbox[2] - bbox[0], 2)),
                        "bbox_height": float(round(bbox[3] - bbox[1], 2)),
                        "crop_filename": crop_filename,
                        "crop_path": crop_path,
                        "underlying_xobjects_count": len(img_list),
                        "status": "Has Alt" if fig["has_alt"] else "Missing Alt"
                    }
                    results.append(fig_info)
                else:
                    # Bounding box could not be resolved - do not crop or substitute page region
                    fig_info = {
                        **fig,
                        "bbox": None,
                        "bbox_width": 0.0,
                        "bbox_height": 0.0,
                        "crop_filename": None,
                        "crop_path": None,
                        "underlying_xobjects_count": len(img_list),
                        "status": "Has Alt" if fig["has_alt"] else "Missing Alt",
                        "crop_status": "no crop available"
                    }
                    results.append(fig_info)
            else:
                # Page number could not be determined
                fig_info = {
                    **fig,
                    "bbox": None,
                    "bbox_width": 0.0,
                    "bbox_height": 0.0,
                    "crop_filename": None,
                    "crop_path": None,
                    "underlying_xobjects_count": 0,
                    "status": "Missing Alt",
                    "crop_status": "no crop available"
                }
                results.append(fig_info)

        return results

    def close(self):
        self.doc.close()
        self.pdoc.close()

