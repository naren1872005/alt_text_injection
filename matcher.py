import os
import io
import re
from pathlib import Path
from PIL import Image
import numpy as np
import imagehash
import pymupdf
import cv2
from typing import List, Dict, Any, Optional, Tuple, Callable

# -------------------------------------------------------------------------
# Mathematical Constants & Dictionaries
# -------------------------------------------------------------------------

GREEK_WORDS = {
    "alpha": "alpha", "beta": "beta", "gamma": "gamma", "delta": "delta",
    "epsilon": "epsilon", "zeta": "zeta", "eta": "eta", "theta": "theta",
    "iota": "iota", "kappa": "kappa", "lambda": "lambda", "mu": "mu",
    "nu": "nu", "xi": "xi", "omicron": "omicron", "pi": "pi",
    "rho": "rho", "sigma": "sigma", "tau": "tau", "upsilon": "upsilon",
    "phi": "phi", "chi": "chi", "psi": "psi", "omega": "omega"
}

DISTINCT_GREEK = set(GREEK_WORDS.keys())

GREEK_LATIN_MAP = {
    'a': 'alpha', 'alpha': 'a',
    'b': 'beta', 'beta': 'b',
    'g': 'gamma', 'gamma': 'g',
    'd': 'delta', 'delta': 'd',
    'e': 'epsilon', 'epsilon': 'e',
    't': 'theta', 'theta': 't',
    'l': 'lambda', 'lambda': 'l',
    'm': 'mu', 'mu': 'm',
    'n': 'nu', 'nu': 'n',
    'p': 'phi', 'phi': 'p',
    'o': 'omega', 'w': 'omega', 'omega': 'o',
    'r': 'rho', 'rho': 'r',
    's': 'sigma', 'sigma': 's'
}

MATH_FN_MAP = {
    "cosine": "cos", "cos": "cos",
    "sine": "sin", "sin": "sin",
    "tangent": "tan", "tan": "tan",
    "arcsin": "arcsin", "arccos": "arccos", "arctan": "arctan",
    "exp": "exp", "log": "log", "ln": "ln",
    "sqrt": "sqrt", "square root": "sqrt"
}

GLYPH_MAP = {
    'í': 'vector ',
    'ù': ' times ',
    'Ÿ': ' => ',
    'Çk': ' k_hat ',
    'Ç{': ' i_hat ',
    'Ç|': ' j_hat ',
    'Ç': ' hat ',
    'Ü✓': ' theta_dot ',
    'á✓': ' theta_ddot ',
    'Ü휙': ' phi_dot ',
    'á휙': ' phi_ddot ',
    '✓': ' theta ',
    '!': ' omega ',
    '↵': ' alpha ',
    '\"': ' alpha ',
    '휙': ' phi ',
    '*': ' - ',
    '‘': ' ( ',
    '’': ' ) ',
    '×': ' times ',
    '÷': ' / ',
    '²': ' squared ',
    '³': ' cubed ',
    '°': ' deg '
}

def clean_and_map_glyph_text(text: str) -> str:
    """Translates PDF MathType font glyph streams into standard textual math symbols."""
    if not text:
        return ""
    out = text
    for k in sorted(GLYPH_MAP.keys(), key=lambda x: len(x), reverse=True):
        if k in out:
            out = out.replace(k, GLYPH_MAP[k])
    return out

def parse_lhs_subject_and_sub(lhs_str: str) -> Tuple[Optional[str], Optional[str]]:
    """
    Extracts the leading mathematical subject (variable/Greek) and its subscript
    from the left-hand side of an equation string.
    Correctly ignores subscripts on LHS when determining subject (e.g. 'r vector subscript A' -> subject 'r').
    """
    if not lhs_str:
        return None, None
    s = lhs_str.lower().strip()
    s = re.sub(r'^line\s*\d+\s*[-:]\s*', '', s).strip()
    s = re.sub(r'\b(vector|hat|open\s+bracket|open\s+parenthesis|the)\b', '', s).strip()
    
    lead_var = None
    lhs_sub = None
    
    if 'subscript' in s:
        parts = s.split('subscript')
        subj = parts[0].strip()
        sub = parts[1].strip()
        for g in GREEK_WORDS:
            if g in subj:
                lead_var = g
                break
        if not lead_var:
            letters = re.findall(r'[a-zA-Z0-9]', subj)
            if letters:
                lead_var = letters[0]
        sub_cleaned = re.sub(r'[^a-zA-Z0-9_/]', '', sub)
        if sub_cleaned:
            lhs_sub = sub_cleaned
    elif '_' in s:
        parts = s.split('_')
        subj = parts[0].strip()
        sub = parts[1].strip()
        for g in GREEK_WORDS:
            if g in subj:
                lead_var = g
                break
        if not lead_var:
            letters = re.findall(r'[a-zA-Z0-9]', subj)
            if letters:
                lead_var = letters[0]
        sub_cleaned = re.sub(r'[^a-zA-Z0-9_/]', '', sub)
        if sub_cleaned:
            lhs_sub = sub_cleaned
    else:
        for g in GREEK_WORDS:
            if g in s:
                lead_var = g
                rest = s.replace(g, '').strip()
                if rest:
                    lhs_sub = re.sub(r'[^a-zA-Z0-9_/]', '', rest)
                break
        if not lead_var:
            m = re.findall(r'([a-zA-Z0-9])([a-zA-Z0-9_/]*)', s)
            if m:
                lead_var = m[0][0]
                if m[0][1]:
                    lhs_sub = m[0][1]

    return lead_var, lhs_sub

def parse_math_tokens(text: str, is_pdf_glyph: bool = False) -> Dict[str, Any]:
    """
    Extracts categorized math tokens from either PDF bbox text or Excel Alt text.
    """
    if not text:
        return {
            "variables": set(),
            "leading_var": None,
            "leading_sub": None,
            "subscripts": set(),
            "greek": set(),
            "functions": set(),
            "powers": set(),
            "derivatives": set(),
            "operators": set(),
            "numbers": set(),
            "is_equation": False,
            "is_fragment": True,
            "raw_text": ""
        }

    raw = clean_and_map_glyph_text(text) if is_pdf_glyph else text
    raw_lower = raw.lower()

    variables = set()
    subscripts = set()
    greek = set()
    functions = set()
    powers = set()
    derivatives = set()
    operators = set()
    numbers = set()

    # 1. Greek letters
    for g in GREEK_WORDS:
        if re.search(r'(?:\b|_)' + g + r'(?:\b|_)', raw_lower) or (g in raw_lower and len(g) > 4):
            greek.add(g)

    # 2. Mathematical functions
    for fn_key, fn_val in MATH_FN_MAP.items():
        if re.search(r'\b' + fn_key + r'\b', raw_lower):
            functions.add(fn_val)

    # 3. Derivatives
    if 'ddot' in raw_lower or 'double dot' in raw_lower or 'á' in text or 'two dots' in raw_lower:
        derivatives.add('ddot')
    if 'dot' in raw_lower or 'Ü' in text or 'with single dot' in raw_lower or 'one dot' in raw_lower:
        derivatives.add('dot')
    if 'prime' in raw_lower:
        derivatives.add('prime')

    # 4. Powers
    if 'squared' in raw_lower or 'superscript 2' in raw_lower or '²' in text or '^2' in raw or 's2' in raw_lower or 'rad/s2' in raw_lower:
        powers.add('2')
    if 'cubed' in raw_lower or 'superscript 3' in raw_lower or '³' in text or '^3' in raw or 's3' in raw_lower:
        powers.add('3')
    pow_matches = re.findall(r'(?:superscript|\^)\s*([0-9a-zA-Z]+)', raw_lower)
    for pm in pow_matches:
        powers.add(pm)

    # 5. Numbers / numerical constants
    num_matches = re.findall(r'\b\d+(?:\.\d+)?\b', raw)
    for nm in num_matches:
        if float(nm) > 0:
            numbers.add(nm)

    # 6. Operators & Equation status
    is_equation = ('=' in raw or 'equals' in raw_lower or '=>' in raw)
    for op_word, op_sym in [('equals', '='), ('times', '*'), ('multiplied by', '*'), ('plus', '+'), ('minus', '-'), ('cross', 'times'), ('vector', 'vector'), ('hat', 'hat'), ('divided by', '/')]:
        if op_word in raw_lower:
            operators.add(op_sym)
    for sym in ['=', '+', '-', '*', '/']:
        if sym in raw:
            operators.add(sym)

    # 7. Subscripts
    sub_matches = re.findall(r'subscript\s+([a-zA-Z0-9_/]+(?:\s+[a-zA-Z0-9_/]+)*?)(?=\s+(?:equals|multiplied|plus|minus|vector|with|and|in|on|bracket|close|open|hat|\)|\()|$)', raw_lower)
    for sm in sub_matches:
        cleaned_sub = sm.strip().replace(" ", "")
        if cleaned_sub:
            subscripts.add(cleaned_sub)
    
    if is_pdf_glyph:
        pdf_subs = re.findall(r'[a-zA-Z!✓↵휙]_?([A-Z0-9]+(?:/[A-Z0-9]+)?)', text)
        for ps in pdf_subs:
            subscripts.add(ps.lower().replace("/", "_"))

    # 8. Leading variable & leading subscript on LHS
    leading_var = None
    leading_sub = None
    if is_equation:
        eq_parts = raw_lower.split("equals") if "equals" in raw_lower else raw_lower.split("=")
        lhs = eq_parts[0] if len(eq_parts) > 1 else ""
        leading_var, leading_sub = parse_lhs_subject_and_sub(lhs)
        if leading_var:
            if leading_var in GREEK_WORDS:
                greek.add(leading_var)
            else:
                variables.add(leading_var)
        if leading_sub:
            subscripts.add(leading_sub)

    # Variables
    all_single_vars = re.findall(r'\b([a-zA-Z])\b', raw_lower)
    for sv in all_single_vars:
        if sv not in ['a', 'i', 'e', 'o', 'in', 'on', 'to', 'by', 'of', 'is', 'the', 'and', 'or']:
            variables.add(sv)
        elif sv in ['a', 'i', 'e'] and (f"vector {sv}" in raw_lower or f"{sv} subscript" in raw_lower or f"{sv}_" in raw_lower or f" {sv} " in raw):
            variables.add(sv)

    # Detect if this is an ungrounded fragment crop (< 3 chars, no equals, no expressions)
    is_fragment = (not is_equation and len(raw.strip()) <= 4 and len(functions) == 0 and len(derivatives) == 0 and len(greek) <= 1)

    return {
        "variables": variables,
        "leading_var": leading_var,
        "leading_sub": leading_sub,
        "subscripts": subscripts,
        "greek": greek,
        "functions": functions,
        "powers": powers,
        "derivatives": derivatives,
        "operators": operators,
        "numbers": numbers,
        "is_equation": is_equation,
        "is_fragment": is_fragment,
        "raw_text": raw
    }

def verify_math_compatibility(pdf_tokens: Dict[str, Any], excel_tokens: Dict[str, Any]) -> Tuple[bool, float, str]:
    """
    Checks for critical mathematical conflicts and computes token compatibility score.
    Returns (is_compatible, token_score, reason).
    """
    p_has_math = bool(pdf_tokens.get("greek") or pdf_tokens.get("subscripts") or pdf_tokens.get("functions") or pdf_tokens.get("derivatives") or pdf_tokens.get("powers") or pdf_tokens.get("is_equation") or pdf_tokens.get("leading_var"))
    e_has_math = bool(excel_tokens.get("greek") or excel_tokens.get("subscripts") or excel_tokens.get("functions") or excel_tokens.get("derivatives") or excel_tokens.get("powers") or excel_tokens.get("is_equation") or excel_tokens.get("leading_var"))

    if p_has_math != e_has_math:
        return False, 0.0, "Category mismatch: one item has math structure, the other does not"

    # 0. Fragment Gating: An ungrounded snippet cannot match a full equation
    if pdf_tokens.get("is_fragment") and excel_tokens.get("is_equation"):
        return False, 0.0, "Fragment mismatch: PDF is a single snippet, Excel is a full equation"

    # 1. Leading variable conflict (e.g. v = ... vs a = ..., or a = ... vs r = ...)
    p_lead = pdf_tokens.get("leading_var")
    e_lead = excel_tokens.get("leading_var")
    if p_lead and e_lead:
        is_direct_match = (p_lead == e_lead)
        is_alias_match = (GREEK_LATIN_MAP.get(p_lead) == e_lead or GREEK_LATIN_MAP.get(e_lead) == p_lead)
        if not (is_direct_match or is_alias_match):
            return False, 0.0, f"Leading variable conflict: PDF '{p_lead}' vs Excel '{e_lead}'"

    # 2. Leading subscript conflict on same leading variable (e.g. omega_c vs omega_d)
    p_lead_sub = pdf_tokens.get("leading_sub")
    e_lead_sub = excel_tokens.get("leading_sub")
    if p_lead_sub and e_lead_sub:
        if p_lead_sub != e_lead_sub and p_lead_sub not in e_lead_sub and e_lead_sub not in p_lead_sub:
            return False, 0.0, f"Leading subscript conflict: PDF '_{p_lead_sub}' vs Excel '_{e_lead_sub}'"

    # 3. Greek symbol conflict (e.g. gamma vs theta, phi vs beta)
    p_greek = pdf_tokens.get("greek", set())
    e_greek = excel_tokens.get("greek", set())
    for g in DISTINCT_GREEK:
        if (g in p_greek and g not in e_greek) or (g in e_greek and g not in p_greek):
            return False, 0.0, f"Greek symbol conflict on '{g}': PDF {p_greek} vs Excel {e_greek}"

    if p_greek and e_greek and not (p_greek & e_greek):
        return False, 0.0, f"Greek symbol mismatch: PDF {p_greek} vs Excel {e_greek}"

    # 4. Subscript conflict (e.g. {b, ob} vs {arm, e})
    p_subs = pdf_tokens.get("subscripts", set())
    e_subs = excel_tokens.get("subscripts", set())
    if p_subs and e_subs:
        sub_overlap = any(
            (ps == es) or (ps in es) or (es in ps)
            for ps in p_subs for es in e_subs
        )
        if not sub_overlap:
            return False, 0.0, f"Subscript conflict: PDF {p_subs} vs Excel {e_subs}"

    # 5. Exponent / Power conflict (e.g. squared vs none)
    p_pow = pdf_tokens.get("powers", set())
    e_pow = excel_tokens.get("powers", set())
    if p_pow != e_pow:
        if ('2' in p_pow and '2' not in e_pow) or ('2' in e_pow and '2' not in p_pow):
            return False, 0.0, f"Exponent mismatch (squared): PDF {p_pow} vs Excel {e_pow}"
        if ('3' in p_pow and '3' not in e_pow) or ('3' in e_pow and '3' not in p_pow):
            return False, 0.0, f"Exponent mismatch (cubed): PDF {p_pow} vs Excel {e_pow}"

    # 6. Mathematical function conflict (e.g. cos/sin vs none)
    p_fn = pdf_tokens.get("functions", set())
    e_fn = excel_tokens.get("functions", set())
    if p_fn and e_fn:
        if not (p_fn & e_fn):
            return False, 0.0, f"Function mismatch: PDF {p_fn} vs Excel {e_fn}"
    elif (p_fn and not e_fn and len(p_fn) >= 2) or (e_fn and not p_fn and len(e_fn) >= 2):
        return False, 0.0, f"Function missing: PDF {p_fn} vs Excel {e_fn}"

    # 7. Derivative conflict
    p_der = pdf_tokens.get("derivatives", set())
    e_der = excel_tokens.get("derivatives", set())
    if 'ddot' in p_der and 'ddot' not in e_der and 'dot' not in e_der and len(e_der) > 0:
        return False, 0.0, f"Derivative mismatch: PDF {p_der} vs Excel {e_der}"

    # 8. Distinct numeric constant mismatch
    p_num = pdf_tokens.get("numbers", set())
    e_num = excel_tokens.get("numbers", set())
    if p_num and e_num:
        overlap = p_num & e_num
        if not overlap and len(p_num) > 1 and len(e_num) > 1:
            return False, 0.0, f"Numerical constants mismatch: PDF {p_num} vs Excel {e_num}"

    # 9. Disjoint variables
    p_vars = pdf_tokens.get("variables", set())
    e_vars = excel_tokens.get("variables", set())
    if p_vars and e_vars:
        overlap = p_vars & e_vars
        if not overlap and len(p_vars) > 1 and len(e_vars) > 1:
            return False, 0.0, f"Complete variable disjoint: PDF {p_vars} vs Excel {e_vars}"

    # Token compatibility score (Jaccard similarity across categories)
    scores = []
    if p_greek or e_greek:
        g_sim = len(p_greek & e_greek) / max(1, len(p_greek | e_greek))
        scores.append((g_sim, 0.30))
    if p_subs or e_subs:
        s_sim = len(p_subs & e_subs) / max(1, len(p_subs | e_subs))
        scores.append((s_sim, 0.30))
    if p_vars or e_vars:
        v_sim = len(p_vars & e_vars) / max(1, len(p_vars | e_vars))
        scores.append((v_sim, 0.20))
    if p_num or e_num:
        n_sim = len(p_num & e_num) / max(1, len(p_num | e_num))
        scores.append((n_sim, 0.10))
    if p_fn or e_fn:
        f_sim = len(p_fn & e_fn) / max(1, len(p_fn | e_fn))
        scores.append((f_sim, 0.10))

    if scores:
        total_w = sum(w for _, w in scores)
        final_token_score = sum(s * w for s, w in scores) / total_w
    else:
        final_token_score = 0.50

    return True, float(round(final_token_score, 2)), "Compatible"

def compute_projection_profiles(img: Image.Image) -> Tuple[np.ndarray, np.ndarray]:
    """Computes normalized 64-bin horizontal and vertical projection profiles."""
    try:
        g = np.array(img.convert('L'))
        b = (g < 220).astype(np.float32)
        h_proj = np.sum(b, axis=1)
        v_proj = np.sum(b, axis=0)

        if len(h_proj) > 0:
            h_res = np.interp(np.linspace(0, 1, 64), np.linspace(0, 1, len(h_proj)), h_proj)
            norm_h = np.linalg.norm(h_res)
            h_res = (h_res / norm_h) if norm_h > 0 else np.zeros(64, dtype=np.float32)
        else:
            h_res = np.zeros(64, dtype=np.float32)

        if len(v_proj) > 0:
            v_res = np.interp(np.linspace(0, 1, 64), np.linspace(0, 1, len(v_proj)), v_proj)
            norm_v = np.linalg.norm(v_res)
            v_res = (v_res / norm_v) if norm_v > 0 else np.zeros(64, dtype=np.float32)
        else:
            v_res = np.zeros(64, dtype=np.float32)

        return h_res, v_res
    except Exception:
        return np.zeros(64, dtype=np.float32), np.zeros(64, dtype=np.float32)

def compute_profile_sim(p1: Tuple[np.ndarray, np.ndarray], p2: Tuple[np.ndarray, np.ndarray]) -> float:
    """Computes fast cosine similarity between precomputed projection profiles."""
    h1, v1 = p1
    h2, v2 = p2
    dot_h = float(np.dot(h1, h2))
    dot_v = float(np.dot(v1, v2))
    return float(round(max(0.0, min(1.0, 0.5 * dot_h + 0.5 * dot_v)), 2))

def is_math_element(item: Dict[str, Any], img: Optional[Image.Image] = None) -> bool:
    """
    Multi-signal math detector.
    Accurately differentiates visual diagrams from inline math equation crops.
    """
    tag = str(item.get("tag_name") or item.get("role") or "").lower()
    if tag in ["formula", "math", "equation"]:
        return True

    if item.get("is_formula"):
        return True

    fn = str(item.get("filename") or item.get("image_filename") or "").lower()
    if re.search(r'(_equation_|_eqn_|_formula_)', fn):
        return True
    if re.search(r'(_img_|_fig_|_figure_|_photo_)', fn):
        return False

    alt = str(item.get("alt_text") or "").lower().strip()
    if re.match(r'^(diagram|mechanical diagram|schematic|drawing|illustration|graph|plot|photo|figure|image|view|cross-section|a |an |the |tilted )\b', alt):
        return False

    if re.match(r'^(?:open parenthesis|line \d+|[a-zA-Z0-9_/\\]+\s*(?:equals|=|subscript|vector|hat))\b', alt) and len(alt) < 150:
        return True

    return False

# -------------------------------------------------------------------------
# Image Normalization & Helpers
# -------------------------------------------------------------------------

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
    Derives an optional 'core graphic' representation for small distinctive graphics.
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

def compute_orb_feature_sim(im1: Image.Image, im2: Image.Image) -> float:
    """
    Computes keypoint feature match ratio between two images using OpenCV ORB.
    Differentiates visually similar diagrams with different internal elements (e.g. 3D disc vs 2D rod).
    """
    try:
        g1 = cv2.cvtColor(np.array(im1.convert('RGB')), cv2.COLOR_RGB2GRAY)
        g2 = cv2.cvtColor(np.array(im2.convert('RGB')), cv2.COLOR_RGB2GRAY)

        orb = cv2.ORB_create(nfeatures=500)
        kp1, des1 = orb.detectAndCompute(g1, None)
        kp2, des2 = orb.detectAndCompute(g2, None)

        if des1 is None or des2 is None or len(des1) < 8 or len(des2) < 8:
            return 0.50

        bf = cv2.BFMatcher(cv2.NORM_HAMMING, crossCheck=True)
        matches = bf.match(des1, des2)

        if not matches:
            return 0.0

        good_matches = [m for m in matches if m.distance < 48]
        max_possible = min(len(kp1), len(kp2))
        if max_possible == 0:
            return 0.0

        match_ratio = len(good_matches) / max_possible
        return float(min(1.0, max(0.0, match_ratio * 1.6)))
    except Exception:
        return 0.50

def extract_img_num(text: Optional[str]) -> Optional[int]:
    """
    Extracts explicit figure / image number from filename or text (e.g. 'img_010.png' -> 10, 'Figure 14' -> 14).
    """
    if not text:
        return None
    s = str(text).lower()
    m = re.search(r'(?:_img_|_fig_|_figure_|_photo_|^img_|^fig_|figure\s*)\s*0*(\d+)', s)
    if m:
        try:
            return int(m.group(1))
        except Exception:
            pass
    return None

def extract_scope_key(fn: str) -> str:
    """
    Extracts the normalized canonical document prefix/scope identifier from a filename.
    """
    if not fn:
        return ""
    s = str(fn).strip()
    s = re.sub(r'\.(png|jpe?g|emf|wmf|gif)$', '', s, flags=re.IGNORECASE)
    s = re.sub(r'(_img_\d+|_equation_\d+|_\d+)$', '', s, flags=re.IGNORECASE)
    s = re.sub(r'^epub_img_', '', s, flags=re.IGNORECASE)
    parts = s.split('_')
    if len(parts) >= 4:
        half = len(parts) // 2
        if parts[:half] == parts[half:]:
            s = "_".join(parts[:half])
    return s.lower()

# -------------------------------------------------------------------------
# VisualMatcher Engine
# -------------------------------------------------------------------------

class VisualMatcher:
    """
    Matches PDF figure crops and math formula crops against candidate Excel manifest images.
    Features:
      1. Multi-signal Math Detection (StructTree, bbox text/glyphs, Alt keywords, visual density)
      2. Categorized Math Token Verification (LHS subject, subscripts, greek, functions, powers, derivatives)
      3. Strict Critical Conflict Gating (mismatches result in instant REJECT before assignment)
      4. Pre-assignment Math Verification (invalid candidates discarded BEFORE 1-to-1 global assignment)
      5. Fast precomputed Structural Projection Profiles
      6. Document Scope Filter
      7. White-margin image normalization
      8. Multi-part figure matching & Alt text fusion
      9. Strict confidence gating (formula threshold >= 0.55)
    """

    MATCH_THRESHOLD = 0.50
    FORMULA_THRESHOLD = 0.55

    def __init__(self, excel_records: List[Dict[str, Any]], excel_images_dir: str, pdf_filename: Optional[str] = None):
        self.excel_images_dir = excel_images_dir
        self.excel_records = excel_records
        with_images = [r for r in excel_records if r.get('has_image')]
        substantial = [r for r in with_images if (r.get('width') or 100) >= 8 and (r.get('height') or 100) >= 8]
        self.candidate_records = substantial if len(substantial) >= 5 else with_images

        self.pdf_filename = pdf_filename
        self.active_scope = self._detect_document_scope()

        self.excel_hashes = {}
        self._precompute_excel_hashes()

    def _detect_document_scope(self) -> Optional[str]:
        prefix_counts = {}
        for r in self.candidate_records:
            fn = r.get("filename") or ""
            pfx = extract_scope_key(fn)
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
            tokens = [t.lower() for t in pfx.split("_") if len(t) > 2]
            overlap = sum(1 for t in tokens if t in sample_text)
            if overlap > best_overlap:
                best_overlap = overlap
                best_pfx = pfx

        return best_pfx if best_overlap > 0 else None

    def _precompute_excel_hashes(self):
        """Precomputes normalized perceptual hashes, projection profiles, and categorized math tokens."""
        self.drawing_hashes = {}
        self.formula_hashes = {}

        for rec in self.candidate_records:
            r = rec["row"]
            img_fn = rec.get("image_filename")
            orig_fn = rec.get("filename") or ""
            if not img_fn:
                continue
            img_path = os.path.join(self.excel_images_dir, img_fn)
            if not os.path.exists(img_path):
                continue

            rec_key = extract_scope_key(orig_fn)
            in_scope = (self.active_scope is None) or (rec_key == self.active_scope) or (self.active_scope in rec_key) or (rec_key in self.active_scope)

            alt = rec.get("alt_text") or ""
            math_tokens = parse_math_tokens(alt, is_pdf_glyph=False)
            is_math = is_math_element(rec)

            try:
                with Image.open(img_path) as orig_im:
                    norm_im = normalize_image(orig_im)
                    w, h = norm_im.size
                    asp = round(w / max(1, h), 2)
                    ph = imagehash.phash(norm_im)
                    dh = imagehash.dhash(norm_im)
                    profiles = compute_projection_profiles(norm_im)

                    data = {
                        "row": r,
                        "phash": ph,
                        "dhash": dh,
                        "aspect": asp,
                        "width": w,
                        "height": h,
                        "in_scope": in_scope,
                        "prefix": rec_key,
                        "rec": rec,
                        "norm_im": norm_im,
                        "profiles": profiles,
                        "math_tokens": math_tokens,
                        "is_math": is_math
                    }
                    self.excel_hashes[r] = data
                    if is_math:
                        self.formula_hashes[r] = data
                    else:
                        self.drawing_hashes[r] = data
            except Exception:
                pass

    def _score_image_pair(self, fig_norm: Image.Image, fig_core: Image.Image, eh: Dict[str, Any], fig_profiles: Optional[Tuple[np.ndarray, np.ndarray]] = None, is_sub_slice: bool = False, fig_page: Optional[Any] = None) -> Dict[str, Any]:
        eh_ph = eh["phash"]
        eh_dh = eh["dhash"]
        eh_asp = eh["aspect"]
        eh_im = eh["norm_im"]
        eh_rec = eh.get("rec", {})
        eh_page = eh_rec.get("page_hint") or eh_rec.get("page")

        # Page match check
        page_matched = False
        if fig_page is not None and eh_page is not None:
            try:
                page_matched = str(fig_page).strip().lower() == str(eh_page).strip().lower() or int(fig_page) == int(eh_page)
            except Exception:
                page_matched = str(fig_page).strip().lower() == str(eh_page).strip().lower()

        # 1. Aspect ratio check
        asp_full = fig_norm.width / max(1, fig_norm.height)
        asp_diff = abs(asp_full - eh_asp) / max(0.5, eh_asp)
        max_asp_tol = 0.65 if page_matched else 0.45
        if asp_diff > max_asp_tol and not is_sub_slice and not page_matched:
            return {"score": 0.0, "rep": "full", "rejection_reason": "ASPECT_MISMATCH"}

        # 2. Perceptual hash
        ph_full = imagehash.phash(fig_norm)
        dh_full = imagehash.dhash(fig_norm)
        diff_p_full = int(ph_full - eh_ph)
        diff_d_full = int(dh_full - eh_dh)
        
        max_hash_diff = 24 if page_matched else 20
        if diff_p_full > max_hash_diff and diff_d_full > max_hash_diff and not is_sub_slice and not page_matched:
            return {"score": 0.0, "rep": "full", "rejection_reason": "LOW_CONFIDENCE"}

        hash_sim_full = max(0.0, 1.0 - (diff_p_full / 32.0)) * 0.6 + max(0.0, 1.0 - (diff_d_full / 32.0)) * 0.4

        # 3. Structural projection profile similarity
        struct_sim = 1.0
        if fig_profiles is not None and "profiles" in eh:
            struct_sim = compute_profile_sim(fig_profiles, eh["profiles"])
            if struct_sim < 0.25 and not is_sub_slice and not page_matched:
                return {"score": 0.0, "rep": "full", "rejection_reason": "LOW_STRUCTURE"}

        # 4. Color similarity
        color_sim = compute_color_sim(fig_norm, eh_im)

        # 5. ORB Keypoint Feature Descriptor similarity
        orb_sim = compute_orb_feature_sim(fig_norm, eh_im)

        asp_pen = min(0.20, asp_diff * 0.15)
        score_full = (0.35 * hash_sim_full) + (0.25 * struct_sim) + (0.25 * orb_sim) + (0.15 * color_sim) - asp_pen
        if page_matched:
            score_full += 0.20
        score_full = max(0.0, min(0.99, score_full))

        best_score = score_full
        best_rep = "full"
        best_diff_p = diff_p_full
        best_diff_d = diff_d_full

        # Core graphic comparison for small elements
        if fig_core is not fig_norm and not is_sub_slice:
            ph_core = imagehash.phash(fig_core)
            dh_core = imagehash.dhash(fig_core)
            diff_p_core = int(ph_core - eh_ph)
            diff_d_core = int(dh_core - eh_dh)
            asp_core = fig_core.width / max(1, fig_core.height)
            asp_diff_core = abs(asp_core - eh_asp) / max(0.5, eh_asp)
            asp_pen_core = min(0.20, asp_diff_core * 0.15)
            hash_sim_core = max(0.0, 1.0 - (diff_p_core / 32.0)) * 0.6 + max(0.0, 1.0 - (diff_d_core / 32.0)) * 0.4
            score_core = (0.50 * hash_sim_core) + (0.35 * struct_sim) + (0.15 * color_sim) - asp_pen_core
            if page_matched:
                score_core += 0.20
            score_core = max(0.0, min(0.99, score_core))

            if score_core > best_score:
                best_score = score_core
                best_rep = "core_graphic"
                best_diff_p = diff_p_core
                best_diff_d = diff_d_core

        scope_mult = 1.0 if eh.get("in_scope", True) else 0.40
        final_score = float(round(min(0.99, float(best_score * scope_mult)), 2))

        return {
            "score": float(final_score),
            "rep": str(best_rep),
            "diff_p": int(best_diff_p),
            "diff_d": int(best_diff_d),
            "struct_sim": float(round(struct_sim, 2)),
            "color_sim": float(round(color_sim, 2)),
            "full_score": float(round(score_full, 2)),
            "page_matched": page_matched
        }

    def match_figures(
        self,
        pdf_figures: List[Dict[str, Any]],
        is_cancelled: Optional[Callable[[], bool]] = None,
        progress_callback: Optional[Callable[[int, str], None]] = None
    ) -> List[Dict[str, Any]]:
        """Matches PDF figures against Excel records with strict category isolation, structural verification, and validated repeated figure propagation."""
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

        # Target candidate pool: Strictly Drawing Records (no equation crops)
        drawing_pool = self.drawing_hashes if self.drawing_hashes else {r: eh for r, eh in self.excel_hashes.items() if not eh.get("is_math")}
        if not drawing_pool:
            drawing_pool = self.excel_hashes

        pdf_reps = {}
        seen_core_hashes = {}
        repeated_figures = {}
        total_figs = max(1, len(pdf_figures))

        for idx, fig in enumerate(pdf_figures):
            if is_cancelled and is_cancelled():
                raise RuntimeError("Visual matching cancelled by user")
            if progress_callback and (idx % 2 == 0 or idx == total_figs - 1):
                pct = int((idx / total_figs) * 100)
                progress_callback(pct, f"Matching figure {idx + 1} of {total_figs}...")
            crop_path = fig.get("crop_path")
            if not crop_path or not os.path.exists(crop_path):
                continue
            try:
                with Image.open(crop_path) as orig_im:
                    norm_im = normalize_image(orig_im)
                    core_im = extract_core_graphic(norm_im)
                    bbox_txt = fig.get("bbox_text") or ""
                    fig_math_tokens = parse_math_tokens(bbox_txt, is_pdf_glyph=True)
                    fig_is_math = is_math_element(fig, norm_im)
                    profiles = compute_projection_profiles(norm_im)
                    w, h = norm_im.size
                    asp = round(w / max(1, h), 2)
                    ph = imagehash.phash(norm_im)
                    dh = imagehash.dhash(norm_im)

                    pdf_reps[idx] = {
                        "norm_im": norm_im,
                        "core_im": core_im,
                        "w": w,
                        "h": h,
                        "aspect": asp,
                        "phash": ph,
                        "dhash": dh,
                        "profiles": profiles,
                        "math_tokens": fig_math_tokens,
                        "is_math": fig_is_math
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

        pair_candidates = []
        fig_best_single = {}

        for fig_idx, reps in pdf_reps.items():
            norm_im = reps["norm_im"]
            core_im = reps["core_im"]
            fig_is_math = reps["is_math"]
            fig_tokens = reps["math_tokens"]
            fig_prof = reps["profiles"]

            fig_data = pdf_figures[fig_idx] if fig_idx < len(pdf_figures) else {}
            fig_num = extract_img_num(fig_data.get("title")) or extract_img_num(fig_data.get("bbox_text")) or fig_data.get("figure_id")
            fig_page = fig_data.get("page_number")

            best_for_fig = None
            best_score_for_fig = -1.0
            rejection_for_fig = "LOW_STRUCTURE"

            for r, eh in drawing_pool.items():
                s_res = self._score_image_pair(norm_im, core_im, eh, fig_profiles=fig_prof, fig_page=fig_page)
                score = s_res["score"]

                if score > 0:
                    rec_fn = eh["rec"].get("image_filename") or eh["rec"].get("filename") or ""
                    rec_num = extract_img_num(rec_fn)
                    if fig_num is not None and rec_num is not None:
                        if fig_num == rec_num:
                            score = float(round(min(0.99, score + 0.15), 2))
                            s_res["score"] = score
                        elif abs(fig_num - rec_num) >= 3:
                            score = float(round(max(0.0, score - 0.20), 2))
                            s_res["score"] = score

                if score > best_score_for_fig:
                    best_score_for_fig = score
                    best_for_fig = {
                        "row": r,
                        "score": score,
                        "rec": eh["rec"],
                        "details": s_res,
                        "in_scope": eh.get("in_scope", True)
                    }
                elif score == 0.0 and "rejection_reason" in s_res:
                    rejection_for_fig = s_res["rejection_reason"]

                if score >= self.MATCH_THRESHOLD:
                    pair_candidates.append({
                        "score": score,
                        "fig_idx": fig_idx,
                        "row": r,
                        "rec": eh["rec"],
                        "match_type": "single",
                        "in_scope": eh.get("in_scope", True),
                        "details": s_res
                    })

            if best_for_fig:
                best_for_fig["rejection_reason"] = "LOW_CONFIDENCE" if 0 < best_score_for_fig < self.MATCH_THRESHOLD else rejection_for_fig
            fig_best_single[fig_idx] = best_for_fig

        # Sort candidate proposals by confidence score descending
        pair_candidates.sort(key=lambda x: x["score"], reverse=True)

        assigned_figs = {}
        claimed_rows = set()

        for prop in pair_candidates:
            f_idx = prop["fig_idx"]
            r = prop["row"]
            if f_idx not in assigned_figs and r not in claimed_rows:
                assigned_figs[f_idx] = prop
                claimed_rows.add(r)

        results = []
        for idx, fig in enumerate(pdf_figures):
            assignment = assigned_figs.get(idx)
            best_single = fig_best_single.get(idx, {})

            if assignment and assignment["score"] >= self.MATCH_THRESHOLD:
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
                        "assigned_rows": [int(assignment["row"])],
                        "score": float(score),
                        "details": assignment.get("details")
                    }
                })
            elif idx in repeated_figures and repeated_figures[idx] in assigned_figs:
                parent_idx = repeated_figures[idx]
                parent_assignment = assigned_figs[parent_idx]
                
                # Check structural consistency between child and parent
                parent_score = float(parent_assignment["score"])
                child_rep = pdf_reps.get(idx)
                parent_rep = pdf_reps.get(parent_idx)
                child_parent_sim = compute_profile_sim(child_rep["profiles"], parent_rep["profiles"]) if (child_rep and parent_rep) else 0.0

                # Strict gating: Only propagate if parent passed validation and child matches parent structure
                if parent_score >= 0.75 and child_parent_sim >= 0.80:
                    parent_rec = parent_assignment["rec"]
                    results.append({
                        **fig,
                        "matched": True,
                        "confidence": parent_score,
                        "status_label": "Auto Match",
                        "excel_match": parent_rec,
                        "match_debug": {
                            "match_type": "repeated_figure_instance",
                            "inherited_from_figure": pdf_figures[parent_idx].get("figure_id"),
                            "assigned_rows": [int(parent_assignment["row"])],
                            "score": parent_score
                        }
                    })
                else:
                    conf = float(best_single.get("score", 0.0)) if best_single else 0.0
                    results.append({
                        **fig,
                        "matched": False,
                        "confidence": conf,
                        "status_label": "No Match",
                        "excel_match": None,
                        "match_debug": {
                            "match_type": "none",
                            "score": conf,
                            "rejection_reason": "PARENT_UNVALIDATED"
                        }
                    })
            else:
                conf = float(best_single.get("score", 0.0)) if best_single else 0.0
                rejection = best_single.get("rejection_reason", "LOW_STRUCTURE") if best_single else "NO_IMAGE"
                best_row = best_single.get("row") if best_single else None
                best_rec = best_single.get("rec") if (best_single and best_single.get("score", 0.0) >= 0.20) else None
                results.append({
                    **fig,
                    "matched": False,
                    "confidence": float(conf),
                    "status_label": "No Match",
                    "excel_match": None,
                    "candidate_match": best_rec,
                    "match_debug": {
                        "match_type": "none",
                        "best_single_row": int(best_row) if best_row is not None else None,
                        "score": float(conf),
                        "rejection_reason": str(rejection)
                    }
                })

        return results

    def match_formulas(
        self,
        pdf_formulas: List[Dict[str, Any]],
        is_cancelled: Optional[Callable[[], bool]] = None,
        progress_callback: Optional[Callable[[int, str], None]] = None
    ) -> List[Dict[str, Any]]:
        """
        Matches PDF formula tags against Excel manifest records with fast precomputed
        projection profiles, strict LHS & token conflict gating, and aspect-ratio validation.
        """
        if not pdf_formulas:
            return []

        if not self.excel_hashes:
            results = []
            for idx, form in enumerate(pdf_formulas):
                matched_rec = self.excel_records[idx] if idx < len(self.excel_records) else None
                results.append({
                    **form,
                    "matched": bool(matched_rec),
                    "confidence": 0.50 if matched_rec else 0.0,
                    "status_label": "Sequential" if matched_rec else "No Match",
                    "excel_match": matched_rec
                })
            return results

        eqn_candidates = []
        gen_candidates = []

        for r, eh in self.excel_hashes.items():
            if eh.get("is_math"):
                eqn_candidates.append(eh)
            else:
                gen_candidates.append(eh)

        target_excel_pool = eqn_candidates if len(eqn_candidates) > 0 else list(self.excel_hashes.values())

        form_reps = {}
        seen_formula_hashes = {}
        repeated_formulas = {}
        total_forms = max(1, len(pdf_formulas))

        for idx, form in enumerate(pdf_formulas):
            if is_cancelled and is_cancelled():
                raise RuntimeError("Formula matching cancelled by user")
            if progress_callback and (idx % 2 == 0 or idx == total_forms - 1):
                pct = int((idx / total_forms) * 100)
                progress_callback(pct, f"Matching formula {idx + 1} of {total_forms}...")
            crop_path = form.get("crop_path")
            if not crop_path or not os.path.exists(crop_path):
                continue
            try:
                with Image.open(crop_path) as orig_im:
                    norm_im = normalize_image(orig_im)
                    w, h = norm_im.size
                    asp = round(w / max(1, h), 2)
                    ph = imagehash.phash(norm_im)
                    dh = imagehash.dhash(norm_im)
                    bbox_txt = form.get("bbox_text") or ""
                    tokens = parse_math_tokens(bbox_txt, is_pdf_glyph=True)
                    profiles = compute_projection_profiles(norm_im)

                    form_reps[idx] = {
                        "norm_im": norm_im,
                        "phash": ph,
                        "dhash": dh,
                        "aspect": asp,
                        "w": w,
                        "h": h,
                        "profiles": profiles,
                        "math_tokens": tokens,
                        "page": form.get("page_number", 1)
                    }
                    h_key = (str(ph), str(dh))
                    if h_key in seen_formula_hashes:
                        repeated_formulas[idx] = seen_formula_hashes[h_key]
                    else:
                        seen_formula_hashes[h_key] = idx
            except Exception:
                pass

        pair_candidates = []
        form_best_single = {}
        
        target_excel_pool.sort(key=lambda x: int(x["row"]))
        num_excel = len(target_excel_pool)

        for f_idx, reps in form_reps.items():
            f_norm_im = reps["norm_im"]
            f_ph = reps["phash"]
            f_dh = reps["dhash"]
            f_asp = reps["aspect"]
            f_tokens = reps["math_tokens"]
            f_prof = reps["profiles"]
            f_h = reps["h"]
            f_w = reps["w"]

            best_for_form = None
            best_score_for_form = -1.0

            for e_pos, eh in enumerate(target_excel_pool):
                r = eh["row"]
                eh_asp = eh["aspect"]
                eh_h = eh["height"]
                eh_w = eh["width"]

                # 1. Aspect ratio gating
                asp_diff = abs(f_asp - eh_asp) / max(0.5, eh_asp)
                if asp_diff > 0.85:
                    continue

                # 2. Scale / Dimension sanity check: prevent tiny fragments matching large multi-line equations
                h_ratio = eh_h / max(1, f_h)
                w_ratio = eh_w / max(1, f_w)
                if (h_ratio > 3.0 or w_ratio > 3.5) and f_tokens.get("is_fragment"):
                    continue

                # 3. Strict Mathematical conflict verification
                e_tokens = eh.get("math_tokens", {})
                is_compat, token_score, reason = verify_math_compatibility(f_tokens, e_tokens)
                if not is_compat:
                    continue

                diff_p = int(f_ph - eh["phash"])
                diff_d = int(f_dh - eh["dhash"])
                asp_pen = min(0.35, asp_diff * 0.25)
                hash_sim = max(0.0, 1.0 - (diff_p / 32.0)) * 0.6 + max(0.0, 1.0 - (diff_d / 32.0)) * 0.4
                vis_score = max(0.0, hash_sim - asp_pen)

                # 4. Fast precomputed projection profile structural similarity
                struct_sim = compute_profile_sim(f_prof, eh["profiles"])

                scope_mult = 1.0 if eh["in_scope"] else 0.40
                
                # Estimated proportional progression position bonus
                if len(pdf_formulas) > 0 and num_excel > 0:
                    expected_e_pos = (f_idx / len(pdf_formulas)) * num_excel
                    pos_dist = abs(e_pos - expected_e_pos) / max(10, num_excel)
                    pos_bonus = max(0.0, 0.12 * (1.0 - min(1.0, pos_dist * 2.0)))
                else:
                    pos_bonus = 0.0

                combined_math_score = (0.40 * token_score + 0.35 * struct_sim + 0.25 * vis_score) * scope_mult + pos_bonus
                final_score = float(round(min(0.99, float(combined_math_score)), 2))

                # Guard: Candidate MUST have either decent visual similarity (>= 0.15) OR high token alignment (>= 0.70)
                if vis_score < 0.15 and token_score < 0.70:
                    final_score = min(final_score, 0.45)

                if final_score > best_score_for_form:
                    best_score_for_form = final_score
                    best_for_form = {
                        "row": r,
                        "score": final_score,
                        "rec": eh["rec"],
                        "in_scope": eh["in_scope"],
                        "e_pos": e_pos,
                        "details": {
                            "token_score": token_score,
                            "struct_sim": struct_sim,
                            "vis_score": vis_score
                        }
                    }

                # Formula candidate threshold: 0.55
                if final_score >= self.FORMULA_THRESHOLD:
                    pair_candidates.append({
                        "score": final_score,
                        "form_idx": f_idx,
                        "row": r,
                        "e_pos": e_pos,
                        "rec": eh["rec"],
                        "match_type": "formula_math_verified",
                        "in_scope": eh["in_scope"],
                        "details": {
                            "token_score": token_score,
                            "struct_sim": struct_sim,
                            "vis_score": vis_score
                        }
                    })

            form_best_single[f_idx] = best_for_form

        # Global Priority Assignment
        pair_candidates.sort(key=lambda x: x["score"], reverse=True)
        assigned_forms = {}
        claimed_rows = set()

        for cand in pair_candidates:
            f_idx = cand["form_idx"]
            r = cand["row"]
            if f_idx not in assigned_forms and r not in claimed_rows:
                assigned_forms[f_idx] = cand
                claimed_rows.add(r)

        results = []
        for idx, form in enumerate(pdf_formulas):
            assignment = assigned_forms.get(idx)
            best_single = form_best_single.get(idx, {})

            if assignment:
                score = assignment["score"]
                status_label = "Auto Match" if score >= 0.70 else "Review"
                results.append({
                    **form,
                    "matched": True,
                    "confidence": float(score),
                    "status_label": status_label,
                    "excel_match": assignment["rec"],
                    "match_debug": {
                        "match_type": str(assignment["match_type"]),
                        "assigned_rows": [int(assignment["row"])],
                        "score": float(score),
                        "details": assignment.get("details")
                    }
                })
            elif idx in repeated_formulas and repeated_formulas[idx] in assigned_forms:
                parent_idx = repeated_formulas[idx]
                parent_assignment = assigned_forms[parent_idx]
                parent_rec = parent_assignment["rec"]
                score = float(parent_assignment["score"])
                status_label = "Auto Match" if score >= 0.70 else "Review"
                results.append({
                    **form,
                    "matched": True,
                    "confidence": score,
                    "status_label": status_label,
                    "excel_match": parent_rec,
                    "match_debug": {
                        "match_type": "repeated_formula_instance",
                        "inherited_from_formula": pdf_formulas[parent_idx].get("formula_id"),
                        "assigned_rows": [int(parent_assignment["row"])],
                        "score": score
                    }
                })
            else:
                conf = float(best_single.get("score", 0.0)) if best_single else 0.0
                rejection = "below_threshold" if conf < self.FORMULA_THRESHOLD else "candidate_claimed"
                if idx in repeated_formulas:
                    rejection = "repeated_formula_unmatched_parent"

                best_row = best_single.get("row") if best_single else None
                best_rec = best_single.get("rec") if (best_single and best_single.get("score", 0.0) >= 0.20) else None
                results.append({
                    **form,
                    "matched": False,
                    "confidence": float(conf),
                    "status_label": "No Match",
                    "excel_match": None,
                    "candidate_match": best_rec,
                    "match_debug": {
                        "match_type": "none",
                        "best_single_row": int(best_row) if best_row is not None else None,
                        "score": float(conf),
                        "rejection_reason": str(rejection)
                    }
                })

        return results
