import os
import io
import re
from pathlib import Path
from PIL import Image
import numpy as np
import imagehash
import pymupdf
import cv2
from typing import List, Dict, Any, Optional

def normalize_image(img: Image.Image, threshold: int = 245) -> Image.Image:
    """
    Trims solid white or transparent margins from an image in memory.
    Preserves actual image content without modifying any files on disk.
    """
    if img.mode in ('RGBA', 'LA') or (img.mode == 'P' and 'transparency' in img.info):
        rgba = img.convert('RGBA')
        alpha = np.array(rgba)[:, :, 3]
        non_transparent = alpha > 10
        if np.any(non_transparent):
            coords = np.argwhere(non_transparent)
            y0, x0 = coords.min(axis=0)
            y1, x1 = coords.max(axis=0) + 1
            pad = 2
            x0 = max(0, x0 - pad)
            y0 = max(0, y0 - pad)
            x1 = min(img.width, x1 + pad)
            y1 = min(img.height, y1 + pad)
            if x1 > x0 + 4 and y1 > y0 + 4:
                img = img.crop((x0, y0, x1, y1))

    rgb = img.convert('RGB')
    arr = np.array(rgb)
    non_white = (arr[:, :, 0] < threshold) | (arr[:, :, 1] < threshold) | (arr[:, :, 2] < threshold)
    if np.any(non_white):
        coords = np.argwhere(non_white)
        y0, x0 = coords.min(axis=0)
        y1, x1 = coords.max(axis=0) + 1
        pad = 2
        x0 = max(0, x0 - pad)
        y0 = max(0, y0 - pad)
        x1 = min(img.width, x1 + pad)
        y1 = min(img.height, y1 + pad)
        if x1 > x0 + 4 and y1 > y0 + 4:
            return img.crop((x0, y0, x1, y1))
    return img

def extract_core_graphic(img: Image.Image, threshold: int = 245) -> Image.Image:
    """
    Derives an optional 'core graphic' representation for small distinctive graphics
    where unrelated peripheral decorative elements (such as horizontal dashes, rules,
    or decorative separators) might distort the visual comparison and aspect ratio.
    Does not permanently modify the image; used as a candidate visual representation.
    """
    try:
        rgb = img.convert('RGB')
        arr = np.array(rgb)
        mask = ((arr[:, :, 0] < threshold) | (arr[:, :, 1] < threshold) | (arr[:, :, 2] < threshold)).astype(np.uint8) * 255
        num_labels, labels, stats, centroids = cv2.connectedComponentsWithStats(mask, connectivity=8)
        if num_labels <= 2:
            return img

        components = []
        for i in range(1, num_labels):
            x, y, w, h, area = stats[i]
            components.append({"label": i, "x": x, "y": y, "w": w, "h": h, "area": area})

        components.sort(key=lambda c: c["area"], reverse=True)
        largest = components[0]

        has_peripheral = False
        valid_components = [largest]
        for c in components[1:]:
            is_small = c["area"] < largest["area"] * 0.25
            is_dash = (c["h"] <= 14 and c["w"] >= c["h"]) or (c["area"] < 200)
            is_lateral = abs(c["x"] - largest["x"]) > largest["w"] * 0.30
            if is_small and (is_dash or is_lateral):
                has_peripheral = True
            else:
                valid_components.append(c)

        if has_peripheral and len(valid_components) < len(components):
            vx0 = min(c["x"] for c in valid_components)
            vy0 = min(c["y"] for c in valid_components)
            vx1 = max(c["x"] + c["w"] for c in valid_components)
            vy1 = max(c["y"] + c["h"] for c in valid_components)
            pad = 2
            vx0 = max(0, vx0 - pad)
            vy0 = max(0, vy0 - pad)
            vx1 = min(img.width, vx1 + pad)
            vy1 = min(img.height, vy1 + pad)
            if vx1 > vx0 + 4 and vy1 > vy0 + 4:
                return img.crop((vx0, vy0, vx1, vy1))
    except Exception:
        pass
    return img

def compute_color_sim(im1: Image.Image, im2: Image.Image) -> float:
    """
    Computes a lightweight color histogram correlation (HSV) between two images.
    Used ONLY as an additional supporting signal for candidates with visual structural similarity.
    Never used alone to produce a match.
    """
    try:
        c1 = cv2.cvtColor(np.array(im1.convert('RGB')), cv2.COLOR_RGB2HSV)
        c2 = cv2.cvtColor(np.array(im2.convert('RGB')), cv2.COLOR_RGB2HSV)
        h1 = cv2.calcHist([c1], [0, 1], None, [30, 32], [0, 180, 0, 256])
        h2 = cv2.calcHist([c2], [0, 1], None, [30, 32], [0, 180, 0, 256])
        cv2.normalize(h1, h1, 0, 1, cv2.NORM_MINMAX)
        cv2.normalize(h2, h2, 0, 1, cv2.NORM_MINMAX)
        correl = cv2.compareHist(h1, h2, cv2.HISTCMP_CORREL)
        return max(0.0, float(correl))
    except Exception:
        return 0.0

class VisualMatcher:
    """
    Matches PDF figure crops against candidate Excel manifest images.
    Features:
      1. Document Scope Filter (detects and prioritizes active document images)
      2. White-margin image normalization (crops solid whitespace before comparison)
      3. Core-graphic normalization (removes decorative dashes/rules for small icons)
      4. Color as supporting signal (HSV histogram correlation, never primary)
      5. Aspect ratio handling (uses core-graphic aspect when evaluating core graphic)
      6. Multi-part figure matching (splits tall/composite figures into vertical components)
      7. Alt text fusion (combines authoritative alt texts for multi-part figures)
      8. Strict confidence gating (confidence < 0.50 => excel_match = None)
      9. Global optimal one-to-one assignment (no Excel row assigned to multiple figures)
      10. Duplicate figure mark detection (preserves repeated-image policy)
      11. Non-visual workflow preservation (sequential fallback if no images exist)
    """

    MATCH_THRESHOLD = 0.50

    def __init__(self, excel_records: List[Dict[str, Any]], excel_images_dir: str, pdf_filename: Optional[str] = None):
        self.excel_images_dir = excel_images_dir
        self.excel_records = excel_records
        with_images = [r for r in excel_records if r.get('has_image')]
        substantial = [r for r in with_images if (r.get('width') or 100) >= 30 and (r.get('height') or 100) >= 30]
        self.candidate_records = substantial if len(substantial) >= 5 else with_images

        self.pdf_filename = pdf_filename
        self.active_scope = self._detect_document_scope()

        self.excel_hashes = {}
        self._precompute_excel_hashes()

    def _detect_document_scope(self) -> Optional[str]:
        """
        Infers the dominant document scope from Excel image filenames and matches
        against PDF title/filename/text sample to filter out unrelated book manifests.
        """
        prefix_counts = {}
        for r in self.candidate_records:
            fn = r.get("filename") or ""
            pfx = re.sub(r'_img_\d+\.[a-zA-Z]+$', '', fn)
            if pfx:
                prefix_counts[pfx] = prefix_counts.get(pfx, 0) + 1

        if not prefix_counts or len(prefix_counts) <= 1:
            return list(prefix_counts.keys())[0] if prefix_counts else None

        sample_text = ""
        if self.pdf_filename:
            sample_text += " " + self.pdf_filename.lower()

        try:
            pdf_path_candidate = Path(self.excel_images_dir).parent / "input.pdf"
            if pdf_path_candidate.exists():
                doc = pymupdf.open(str(pdf_path_candidate))
                sample_text += " " + (doc.metadata.get("title", "") or "")
                if len(doc) > 0:
                    sample_text += " " + doc[0].get_text()[:400]
                doc.close()
        except Exception:
            pass

        sample_text = sample_text.lower()
        best_pfx = None
        best_overlap = -1
        for pfx in prefix_counts:
            tokens = [t.lower() for t in pfx.replace("epub_img_", "").split("_") if len(t) > 2]
            overlap = sum(1 for t in tokens if t in sample_text)
            if overlap > best_overlap:
                best_overlap = overlap
                best_pfx = pfx

        return best_pfx if best_overlap > 0 else None

    def _precompute_excel_hashes(self):
        """
        Precomputes normalized perceptual hashes, aspect ratios, and images for Excel candidates.
        """
        for rec in self.candidate_records:
            r = rec["row"]
            img_fn = rec.get("image_filename")
            orig_fn = rec.get("filename") or ""
            if not img_fn:
                continue
            img_path = os.path.join(self.excel_images_dir, img_fn)
            if not os.path.exists(img_path):
                continue

            rec_pfx = re.sub(r'_img_\d+\.[a-zA-Z]+$', '', orig_fn)
            in_scope = (self.active_scope is None) or (rec_pfx == self.active_scope)

            try:
                with Image.open(img_path) as orig_im:
                    norm_im = normalize_image(orig_im)
                    w, h = norm_im.size
                    asp = round(w / max(1, h), 2)
                    ph = imagehash.phash(norm_im)
                    dh = imagehash.dhash(norm_im)
                    self.excel_hashes[r] = {
                        "row": r,
                        "phash": ph,
                        "dhash": dh,
                        "aspect": asp,
                        "width": w,
                        "height": h,
                        "in_scope": in_scope,
                        "prefix": rec_pfx,
                        "rec": rec,
                        "norm_im": norm_im
                    }
            except Exception:
                pass

    def _score_image_pair(self, fig_norm: Image.Image, fig_core: Image.Image, eh: Dict[str, Any], is_sub_slice: bool = False) -> Dict[str, Any]:
        """
        Calculates visual similarity between a figure representation (evaluating both
        full normalized image and core-graphic representation) and an Excel candidate.
        """
        eh_ph = eh["phash"]
        eh_dh = eh["dhash"]
        eh_asp = eh["aspect"]
        eh_im = eh["norm_im"]

        # 1. Full normalized comparison
        ph_full = imagehash.phash(fig_norm)
        dh_full = imagehash.dhash(fig_norm)
        diff_p_full = int(ph_full - eh_ph)
        diff_d_full = int(dh_full - eh_dh)
        asp_full = fig_norm.width / max(1, fig_norm.height)
        asp_diff_full = abs(asp_full - eh_asp) / max(0.5, eh_asp)
        asp_pen_full = min(0.35, asp_diff_full * 0.25)
        hash_sim_full = max(0.0, 1.0 - (diff_p_full / 32.0)) * 0.6 + max(0.0, 1.0 - (diff_d_full / 32.0)) * 0.4
        score_full = max(0.0, hash_sim_full - asp_pen_full)

        best_score = score_full
        best_rep = "full"
        best_diff_p = diff_p_full
        best_diff_d = diff_d_full
        best_im = fig_norm

        # 2. Core graphic comparison (for small graphics with decorative dashes/rules)
        if fig_core is not fig_norm and not is_sub_slice:
            ph_core = imagehash.phash(fig_core)
            dh_core = imagehash.dhash(fig_core)
            diff_p_core = int(ph_core - eh_ph)
            diff_d_core = int(dh_core - eh_dh)
            asp_core = fig_core.width / max(1, fig_core.height)
            asp_diff_core = abs(asp_core - eh_asp) / max(0.5, eh_asp)
            asp_pen_core = min(0.35, asp_diff_core * 0.25)
            hash_sim_core = max(0.0, 1.0 - (diff_p_core / 32.0)) * 0.6 + max(0.0, 1.0 - (diff_d_core / 32.0)) * 0.4
            score_core = max(0.0, hash_sim_core - asp_pen_core)

            if score_core > best_score:
                best_score = score_core
                best_rep = "core_graphic"
                best_diff_p = diff_p_core
                best_diff_d = diff_d_core
                best_im = fig_core

        # 3. Optional color similarity boost (supports shape match, never matches on color alone)
        color_sim = 0.0
        if best_score >= 0.40:
            color_sim = compute_color_sim(best_im, eh_im)
            best_score = (best_score * 0.85) + (color_sim * 0.15)

        # Scope penalty if candidate is outside the detected document
        scope_mult = 1.0 if eh["in_scope"] else 0.20
        final_score = float(round(min(0.99, float(best_score * scope_mult)), 2))

        return {
            "score": float(final_score),
            "rep": str(best_rep),
            "diff_p": int(best_diff_p),
            "diff_d": int(best_diff_d),
            "color_sim": float(round(color_sim, 2)),
            "full_score": float(round(score_full, 2))
        }

    def match_figures(self, pdf_figures: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        Matches PDF figures against Excel records with 1-to-1 global assignment,
        composite multi-part figure matching, core-graphic normalization, and alt-text fusion.
        """
        # Non-visual fallback (preserves equation / non-image manifest behavior)
        if not self.excel_hashes:
            results = []
            for idx, fig in enumerate(pdf_figures):
                matched_rec = self.excel_records[idx] if idx < len(self.excel_records) else None
                results.append({
                    **fig,
                    "matched": bool(matched_rec),
                    "confidence": 0.50 if matched_rec else 0.0,
                    "status_label": "Sequential" if matched_rec else "No Match",
                    "excel_match": matched_rec
                })
            return results

        # 1. Preprocess PDF figures: normalize, extract core graphic, detect duplicate marks
        pdf_reps = {}
        seen_core_hashes = {}
        repeated_figures = {}

        for idx, fig in enumerate(pdf_figures):
            crop_path = fig.get("crop_path")
            if not crop_path or not os.path.exists(crop_path):
                continue
            try:
                with Image.open(crop_path) as orig_im:
                    norm_im = normalize_image(orig_im)
                    core_im = extract_core_graphic(norm_im)
                    pdf_reps[idx] = {
                        "norm_im": norm_im,
                        "core_im": core_im,
                        "w": norm_im.width,
                        "h": norm_im.height
                    }
                    core_ph = str(imagehash.phash(core_im))
                    core_dh = str(imagehash.dhash(core_im))
                    h_key = (core_ph, core_dh)
                    if h_key in seen_core_hashes:
                        repeated_figures[idx] = seen_core_hashes[h_key]
                    else:
                        seen_core_hashes[h_key] = idx
            except Exception:
                pass

        # 2. Score standard single candidate pairs
        pair_candidates = []
        fig_best_single = {}

        for fig_idx, reps in pdf_reps.items():
            norm_im = reps["norm_im"]
            core_im = reps["core_im"]

            best_for_fig = None
            best_score_for_fig = -1.0

            for r, eh in self.excel_hashes.items():
                s_res = self._score_image_pair(norm_im, core_im, eh)
                score = s_res["score"]

                if score > best_score_for_fig:
                    best_score_for_fig = score
                    best_for_fig = {
                        "row": r,
                        "score": score,
                        "rec": eh["rec"],
                        "details": s_res,
                        "in_scope": eh["in_scope"]
                    }

                if score >= self.MATCH_THRESHOLD:
                    pair_candidates.append({
                        "score": score,
                        "fig_idx": fig_idx,
                        "row": r,
                        "rec": eh["rec"],
                        "match_type": "single",
                        "in_scope": eh["in_scope"],
                        "details": s_res
                    })

            fig_best_single[fig_idx] = best_for_fig

        # 3. Multi-Part Figure Matching for tall/stacked composite figures
        # Evaluates vertical slices against Excel images and enables alt text fusion
        multi_part_candidates = []

        for fig_idx, reps in pdf_reps.items():
            best_single = fig_best_single.get(fig_idx)
            single_score = best_single["score"] if best_single else 0.0

            # Only evaluate multi-part if single match is not overwhelmingly confident
            if single_score < 0.75 and reps["h"] >= 150 and reps["w"] >= 150:
                norm_im = reps["norm_im"]
                h, w = reps["h"], reps["w"]

                scope_rows = sorted([r for r, eh in self.excel_hashes.items() if eh["in_scope"]])

                # Check adjacent rows in the manifest for composite arrangement
                for i in range(len(scope_rows) - 1):
                    r_top = scope_rows[i]
                    r_bot = scope_rows[i + 1]
                    eh_top = self.excel_hashes[r_top]
                    eh_bot = self.excel_hashes[r_bot]

                    # Expected split boundary based on top image aspect ratio
                    expected_top_h = int(w / eh_top["aspect"])
                    if 0.10 * h <= expected_top_h <= 0.80 * h:
                        min_y = max(int(0.10 * h), int(expected_top_h * 0.85))
                        max_y = min(int(0.85 * h), int(expected_top_h * 1.30))

                        best_pair_score = -1.0
                        best_pair_split = None
                        best_top_dict = None
                        best_bot_dict = None

                        step = max(2, (max_y - min_y) // 12)
                        for split_y in range(min_y, max_y + 1, step):
                            slice_top = normalize_image(norm_im.crop((0, 0, w, split_y)))
                            slice_bot = normalize_image(norm_im.crop((0, split_y, w, h)))

                            score_top_dict = self._score_image_pair(slice_top, slice_top, eh_top, is_sub_slice=True)
                            score_bot_dict = self._score_image_pair(slice_bot, slice_bot, eh_bot, is_sub_slice=True)

                            s_top = score_top_dict["score"]
                            s_bot = score_bot_dict["score"]

                            if s_top >= self.MATCH_THRESHOLD and s_bot >= self.MATCH_THRESHOLD:
                                avg_score = (s_top + s_bot) / 2.0
                                if avg_score > best_pair_score:
                                    best_pair_score = avg_score
                                    best_pair_split = split_y
                                    best_top_dict = score_top_dict
                                    best_bot_dict = score_bot_dict

                        if best_pair_score >= self.MATCH_THRESHOLD:
                            # Authoritative Alt Text Fusion preserving exact wording
                            alt_top = (eh_top["rec"].get("alt_text") or "").strip()
                            alt_bot = (eh_bot["rec"].get("alt_text") or "").strip()
                            if alt_top and alt_bot:
                                sep = " " if alt_top.endswith((".", "!", "?")) else ". "
                                fused_alt = f"{alt_top}{sep}{alt_bot}"
                            else:
                                fused_alt = alt_top or alt_bot

                            fused_rec = {
                                **eh_top["rec"],
                                "row": [int(r_top), int(r_bot)],
                                "filename": f"{eh_top['rec'].get('filename')} + {eh_bot['rec'].get('filename')}",
                                "alt_text": fused_alt,
                                "match_type": "composite_multi_part",
                                "source_rows": [int(r_top), int(r_bot)],
                                "components": [
                                    {"row": int(r_top), "filename": eh_top['rec'].get('filename'), "confidence": float(best_top_dict["score"]), "region": "upper_region"},
                                    {"row": int(r_bot), "filename": eh_bot['rec'].get('filename'), "confidence": float(best_bot_dict["score"]), "region": "lower_region"}
                                ]
                            }

                            multi_part_candidates.append({
                                "score": float(round(best_pair_score, 2)),
                                "fig_idx": int(fig_idx),
                                "rows": [int(r_top), int(r_bot)],
                                "rec": fused_rec,
                                "match_type": "composite_multi_part",
                                "in_scope": True,
                                "top_score": float(best_top_dict["score"]),
                                "bot_score": float(best_bot_dict["score"]),
                                "split_y": int(best_pair_split) if best_pair_split is not None else None
                            })

        # 4. Global Priority Assignment:
        # Combine single and multi-part proposals, sort by score descending
        all_proposals = []
        for cand in pair_candidates:
            all_proposals.append({
                "score": cand["score"],
                "fig_idx": cand["fig_idx"],
                "rows": [cand["row"]],
                "rec": cand["rec"],
                "match_type": "single",
                "details": cand.get("details", {})
            })

        for m_cand in multi_part_candidates:
            all_proposals.append({
                "score": m_cand["score"],
                "fig_idx": m_cand["fig_idx"],
                "rows": m_cand["rows"],
                "rec": m_cand["rec"],
                "match_type": "composite_multi_part",
                "details": {
                    "top_score": m_cand["top_score"],
                    "bot_score": m_cand["bot_score"],
                    "split_y": m_cand["split_y"]
                }
            })

        all_proposals.sort(key=lambda x: x["score"], reverse=True)

        assigned_figs = {}
        claimed_rows = set()

        for prop in all_proposals:
            f_idx = prop["fig_idx"]
            prop_rows = prop["rows"]
            # Assign if figure unassigned and ALL required rows are completely unclaimed
            if f_idx not in assigned_figs and not any(r in claimed_rows for r in prop_rows):
                assigned_figs[f_idx] = prop
                for r in prop_rows:
                    claimed_rows.add(r)

        # 5. Build final output list with strict confidence gating
        results = []
        for idx, fig in enumerate(pdf_figures):
            assignment = assigned_figs.get(idx)
            best_single = fig_best_single.get(idx, {})

            if assignment:
                score = assignment["score"]
                status_label = "Auto Match" if score >= 0.75 else "Review"
                results.append({
                    **fig,
                    "matched": True,
                    "confidence": float(score),
                    "status_label": status_label,
                    "excel_match": assignment["rec"],
                    "match_debug": {
                        "match_type": str(assignment["match_type"]),
                        "assigned_rows": [int(r) for r in assignment["rows"]],
                        "score": float(score),
                        "details": assignment.get("details")
                    }
                })
            elif idx in repeated_figures and repeated_figures[idx] in assigned_figs:
                # Repeated Figure Alt Inheritance: clones inherit authoritative Alt text from parent figure
                parent_idx = repeated_figures[idx]
                parent_assignment = assigned_figs[parent_idx]
                parent_rec = parent_assignment["rec"]
                score = float(parent_assignment["score"])
                status_label = "Auto Match" if score >= 0.75 else "Review"

                results.append({
                    **fig,
                    "matched": True,
                    "confidence": score,
                    "status_label": status_label,
                    "excel_match": parent_rec,
                    "match_debug": {
                        "match_type": "repeated_figure_instance",
                        "inherited_from_figure": pdf_figures[parent_idx].get("figure_id"),
                        "assigned_rows": [int(r) for r in parent_assignment["rows"]],
                        "score": score
                    }
                })
            else:
                conf = float(best_single.get("score", 0.0)) if best_single else 0.0
                rejection = "below_threshold" if conf < self.MATCH_THRESHOLD else "candidate_claimed"
                if idx in repeated_figures:
                    rejection = "repeated_figure_mark_unmatched_parent"

                best_row = best_single.get("row") if best_single else None
                results.append({
                    **fig,
                    "matched": False,
                    "confidence": float(conf),
                    "status_label": "No Match",
                    "excel_match": None,
                    "match_debug": {
                        "match_type": "none",
                        "best_single_row": int(best_row) if best_row is not None else None,
                        "score": float(conf),
                        "rejection_reason": str(rejection)
                    }
                })

        return results
