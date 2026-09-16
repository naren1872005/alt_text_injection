import io
import os
import shutil
import uuid
import zipfile
import csv
from pathlib import Path
from typing import Optional, Dict, Any, List
from pydantic import BaseModel

from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks, Form
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter

from extractor import FigureExtractor
from excel_parser import ExcelParser
from matcher import VisualMatcher

BASE_DIR = Path(__file__).resolve().parent
SESSIONS_DIR = BASE_DIR / "sessions"
STATIC_DIR = BASE_DIR / "static"

SESSIONS_DIR.mkdir(exist_ok=True)
STATIC_DIR.mkdir(exist_ok=True)

app = FastAPI(title="PDF Figure Tag & Image Extractor", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# In-memory session cache
session_cache = {}

def auto_attach_excel_if_available(session_id: str, figures: list, allow_sample: bool = False):
    """
    Checks if an Excel manifest was uploaded to the session (session_path / "manifest.xlsx").
    Only falls back to default sample candidates if allow_sample is explicitly True.
    """
    session_path = SESSIONS_DIR / session_id
    excel_candidates = [session_path / "manifest.xlsx"]
    if allow_sample:
        excel_candidates.extend([
            BASE_DIR / "Plesha_EngineeringMechanics_3e_Chap015_ISM 1 (1).xlsx",
            Path(r"D:\pdf_alt_textconverter\Plesha_EngineeringMechanics_3e_Chap015_ISM 1 (1).xlsx"),
            Path(r"D:\py_automation_alt\pdf_alt_text_automation\input\Plesha_EngineeringMechanics_3e_Chap015_ISM 1 (1).xlsx"),
            Path(r"C:\Users\NarenKG\Downloads\Plesha_EngineeringMechanics_3e_Chap015_ISM 1 (1).xlsx")
        ])
    excel_path = None
    for p in excel_candidates:
        if p.exists():
            excel_path = p
            break
            
    if not excel_path:
        return figures, [], None

    excel_images_dir = session_path / "excel_images"
    excel_images_dir.mkdir(parents=True, exist_ok=True)
    
    try:
        parser = ExcelParser(str(excel_path))
        excel_records = parser.parse_all_records(output_dir=str(excel_images_dir))
        parser.close()
        
        for rec in excel_records:
            if rec.get("image_filename"):
                rec["image_url"] = f"/api/excel-image/{session_id}/{rec['image_filename']}"
                
        if figures:
            matcher = VisualMatcher(excel_records, str(excel_images_dir))
            matched_figures = matcher.match_figures(figures)
            for fig in matched_figures:
                ex = fig.get("excel_match")
                if ex and ex.get("image_filename"):
                    ex["image_url"] = f"/api/excel-image/{session_id}/{ex['image_filename']}"
                fig["excel_match"] = ex
            figures = matched_figures
            
        return figures, excel_records, excel_path.name
    except Exception as e:
        print(f"Excel attach failed: {e}")
        return figures, [], None

@app.get("/api/health")
def health_check():
    return {"status": "ok", "service": "PDF Figure Tag Extractor"}

@app.post("/api/upload-pdf")
async def upload_pdf(file: UploadFile = File(...), session_id: Optional[str] = Form(None)):
    """
    Receives an uploaded PDF, parses its StructTreeRoot for all /Figure tags,
    computes bounding boxes from MCIDs, and renders high-res crops for each figure.
    Does NOT use default sample files. Works for any user PDF.
    """
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Uploaded file must be a PDF.")

    if not session_id or session_id not in session_cache:
        session_id = str(uuid.uuid4())
        session_cache[session_id] = {
            "session_id": session_id,
            "figures": [],
            "figures_count": 0
        }

    session_path = SESSIONS_DIR / session_id
    session_path.mkdir(parents=True, exist_ok=True)

    input_pdf_path = session_path / "input.pdf"
    figures_output_dir = session_path / "figures"
    figures_output_dir.mkdir(parents=True, exist_ok=True)

    # Save uploaded file
    contents = await file.read()
    with open(input_pdf_path, "wb") as f:
        f.write(contents)

    # Run extraction
    extractor = None
    try:
        extractor = FigureExtractor(str(input_pdf_path))
        raw_figures = extractor.extract_figures_with_crops(str(figures_output_dir), dpi=150)
        
        # Enrich figures with image URLs
        figures = []
        for fig in raw_figures:
            crop_fn = fig.get("crop_filename")
            image_url = f"/api/figure-image/{session_id}/{crop_fn}" if crop_fn else None
            figures.append({
                **fig,
                "image_url": image_url
            })

        # Attach Excel manifest ONLY if an Excel file was uploaded to this session (allow_sample=False)
        figures, excel_records, excel_fn = auto_attach_excel_if_available(session_id, figures, allow_sample=False)

        has_alt_count = sum(1 for f in figures if f["has_alt"] or f.get("excel_match"))
        missing_alt_count = len(figures) - has_alt_count

        excel_images_count = sum(1 for r in excel_records if r.get("has_image"))
        excel_has_alt = sum(1 for r in excel_records if r.get("has_alt"))

        result = {
            "session_id": session_id,
            "filename": file.filename,
            "total_pages": extractor.page_count,
            "has_struct_tree": extractor.has_struct_tree,
            "figures_count": len(figures),
            "has_alt_count": has_alt_count,
            "missing_alt_count": missing_alt_count,
            "figures": figures,
            "excel_filename": excel_fn,
            "excel_records": excel_records,
            "excel_total_records": len(excel_records),
            "excel_images_count": excel_images_count,
            "excel_has_alt_count": excel_has_alt,
            "matched_figures": figures
        }

        # Cache session data
        session_cache[session_id] = result
        return JSONResponse(content=result)

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error processing PDF: {str(e)}")
    finally:
        if extractor:
            extractor.close()

@app.post("/api/load-sample")
async def load_sample():
    """
    Quick test endpoint that loads the local sample PDF Chap_015 126-187_A11y.pdf
    and automatically matches with the attached Excel manifest.
    """
    candidate_paths = [
        BASE_DIR / "Chap_015 126-187_A11y.pdf",
        Path(r"D:\pdf_alt_textconverter\Chap_015 126-187_A11y.pdf"),
        Path(r"D:\py_automation_alt\pdf_alt_text_automation\input\Chap_015 126-187_A11y.pdf"),
        Path(r"D:\demo\Chap_015 126-187_A11y.pdf")
    ]
    sample_path = None
    for p in candidate_paths:
        if p.exists():
            sample_path = p
            break

    if not sample_path:
        raise HTTPException(status_code=404, detail="Sample PDF not found on disk.")

    session_id = str(uuid.uuid4())
    session_path = SESSIONS_DIR / session_id
    session_path.mkdir(parents=True, exist_ok=True)

    input_pdf_path = session_path / "input.pdf"
    figures_output_dir = session_path / "figures"
    figures_output_dir.mkdir(parents=True, exist_ok=True)

    shutil.copyfile(sample_path, input_pdf_path)

    extractor = None
    try:
        extractor = FigureExtractor(str(input_pdf_path))
        raw_figures = extractor.extract_figures_with_crops(str(figures_output_dir), dpi=150)
        
        figures = []
        for fig in raw_figures:
            crop_fn = fig.get("crop_filename")
            image_url = f"/api/figure-image/{session_id}/{crop_fn}" if crop_fn else None
            figures.append({
                **fig,
                "image_url": image_url
            })

        # Auto-attach Excel manifest if present in project
        figures, excel_records, excel_fn = auto_attach_excel_if_available(session_id, figures, allow_sample=True)

        has_alt_count = sum(1 for f in figures if f["has_alt"] or f.get("excel_match"))
        missing_alt_count = len(figures) - has_alt_count

        excel_images_count = sum(1 for r in excel_records if r.get("has_image"))
        excel_has_alt = sum(1 for r in excel_records if r.get("has_alt"))

        result = {
            "session_id": session_id,
            "filename": sample_path.name,
            "total_pages": extractor.page_count,
            "has_struct_tree": extractor.has_struct_tree,
            "figures_count": len(figures),
            "has_alt_count": has_alt_count,
            "missing_alt_count": missing_alt_count,
            "figures": figures,
            "excel_filename": excel_fn,
            "excel_records": excel_records,
            "excel_total_records": len(excel_records),
            "excel_images_count": excel_images_count,
            "excel_has_alt_count": excel_has_alt,
            "matched_figures": figures
        }

        session_cache[session_id] = result
        return JSONResponse(content=result)

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error processing sample PDF: {str(e)}")
    finally:
        if extractor:
            extractor.close()

@app.post("/api/upload-excel")
async def upload_excel(file: UploadFile = File(...), session_id: Optional[str] = Form(None)):
    """
    Upload an Excel ALT Manifest (.xlsx), extract all 4,848 image records and authoritative ALT texts,
    and match them visually with the current PDF figures.
    """
    if not file.filename.lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(status_code=400, detail="Uploaded file must be an Excel workbook (.xlsx).")

    if not session_id or session_id not in session_cache:
        # If no active session, create one
        session_id = str(uuid.uuid4())
        session_cache[session_id] = {
            "session_id": session_id,
            "figures": [],
            "figures_count": 0
        }

    session_path = SESSIONS_DIR / session_id
    session_path.mkdir(parents=True, exist_ok=True)
    excel_path = session_path / "manifest.xlsx"
    excel_images_dir = session_path / "excel_images"
    excel_images_dir.mkdir(parents=True, exist_ok=True)

    contents = await file.read()
    with open(excel_path, "wb") as f:
        f.write(contents)

    try:
        parser = ExcelParser(str(excel_path))
        # Extract images and records
        excel_records = parser.parse_all_records(output_dir=str(excel_images_dir))
        parser.close()

        # Add image URLs
        for rec in excel_records:
            if rec.get("image_filename"):
                rec["image_url"] = f"/api/excel-image/{session_id}/{rec['image_filename']}"

        # Run visual matcher if session already has PDF figures
        pdf_figures = session_cache[session_id].get("figures", [])
        if pdf_figures:
            matcher = VisualMatcher(excel_records, str(excel_images_dir))
            matched_figures = matcher.match_figures(pdf_figures)
            # Enrich with image URLs for matched records
            for fig in matched_figures:
                ex = fig.get("excel_match")
                if ex and ex.get("image_filename"):
                    ex["image_url"] = f"/api/excel-image/{session_id}/{ex['image_filename']}"
                fig["excel_match"] = ex
            session_cache[session_id]["figures"] = matched_figures
        else:
            matched_figures = []

        # Calculate metrics for Excel records
        excel_images_count = sum(1 for r in excel_records if r.get("has_image"))
        excel_has_alt = sum(1 for r in excel_records if r.get("has_alt"))
        excel_missing_alt = len(excel_records) - excel_has_alt

        session_cache[session_id]["excel_records"] = excel_records
        session_cache[session_id]["excel_filename"] = file.filename
        session_cache[session_id]["excel_total_records"] = len(excel_records)
        session_cache[session_id]["excel_images_count"] = excel_images_count
        session_cache[session_id]["excel_has_alt_count"] = excel_has_alt
        session_cache[session_id]["excel_missing_alt_count"] = excel_missing_alt

        return JSONResponse(content={
            "session_id": session_id,
            "excel_filename": file.filename,
            "excel_total_records": len(excel_records),
            "excel_images_count": excel_images_count,
            "excel_has_alt_count": excel_has_alt,
            "excel_missing_alt_count": excel_missing_alt,
            "matched_figures": session_cache[session_id].get("figures", []),
            "figures_count": len(session_cache[session_id].get("figures", [])),
            "excel_records": excel_records
        })

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error processing Excel manifest: {str(e)}")

@app.post("/api/load-sample-excel")
async def load_sample_excel(session_id: Optional[str] = None):
    """
    Quick test endpoint to load the client's Plesha Chapter 15 Excel manifest.
    """
    candidate_paths = [
        BASE_DIR / "Plesha_EngineeringMechanics_3e_Chap015_ISM 1 (1).xlsx",
        Path(r"D:\pdf_alt_textconverter\Plesha_EngineeringMechanics_3e_Chap015_ISM 1 (1).xlsx"),
        Path(r"D:\py_automation_alt\pdf_alt_text_automation\input\Plesha_EngineeringMechanics_3e_Chap015_ISM 1 (1).xlsx"),
        Path(r"C:\Users\NarenKG\Downloads\Plesha_EngineeringMechanics_3e_Chap015_ISM 1 (1).xlsx")
    ]
    sample_path = None
    for p in candidate_paths:
        if p.exists():
            sample_path = p
            break

    if not sample_path:
        raise HTTPException(status_code=404, detail="Sample Excel file not found on disk.")

    if not session_id or session_id not in session_cache:
        # Find the latest session if available
        if session_cache:
            session_id = list(session_cache.keys())[-1]
        else:
            session_id = str(uuid.uuid4())
            session_cache[session_id] = {
                "session_id": session_id,
                "figures": [],
                "figures_count": 0
            }

    session_path = SESSIONS_DIR / session_id
    session_path.mkdir(parents=True, exist_ok=True)
    excel_path = session_path / "manifest.xlsx"
    excel_images_dir = session_path / "excel_images"
    excel_images_dir.mkdir(parents=True, exist_ok=True)

    if not excel_path.exists():
        shutil.copyfile(sample_path, excel_path)

    try:
        parser = ExcelParser(str(excel_path))
        excel_records = parser.parse_all_records(output_dir=str(excel_images_dir))
        parser.close()

        for rec in excel_records:
            if rec.get("image_filename"):
                rec["image_url"] = f"/api/excel-image/{session_id}/{rec['image_filename']}"

        pdf_figures = session_cache[session_id].get("figures", [])
        if pdf_figures:
            matcher = VisualMatcher(excel_records, str(excel_images_dir))
            matched_figures = matcher.match_figures(pdf_figures)
            for fig in matched_figures:
                ex = fig.get("excel_match")
                if ex and ex.get("image_filename"):
                    ex["image_url"] = f"/api/excel-image/{session_id}/{ex['image_filename']}"
                fig["excel_match"] = ex
            session_cache[session_id]["figures"] = matched_figures
        else:
            matched_figures = []

        excel_images_count = sum(1 for r in excel_records if r.get("has_image"))
        excel_has_alt = sum(1 for r in excel_records if r.get("has_alt"))
        excel_missing_alt = len(excel_records) - excel_has_alt

        session_cache[session_id]["excel_records"] = excel_records
        session_cache[session_id]["excel_filename"] = sample_path.name
        session_cache[session_id]["excel_total_records"] = len(excel_records)
        session_cache[session_id]["excel_images_count"] = excel_images_count
        session_cache[session_id]["excel_has_alt_count"] = excel_has_alt
        session_cache[session_id]["excel_missing_alt_count"] = excel_missing_alt

        return JSONResponse(content={
            "session_id": session_id,
            "excel_filename": sample_path.name,
            "excel_total_records": len(excel_records),
            "excel_images_count": excel_images_count,
            "excel_has_alt_count": excel_has_alt,
            "excel_missing_alt_count": excel_missing_alt,
            "matched_figures": session_cache[session_id].get("figures", []),
            "figures_count": len(session_cache[session_id].get("figures", [])),
            "excel_records": excel_records
        })

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error processing sample Excel: {str(e)}")

@app.get("/api/excel-image/{session_id}/{filename}")
def get_excel_image(session_id: str, filename: str):
    image_path = SESSIONS_DIR / session_id / "excel_images" / filename
    if not image_path.exists():
        raise HTTPException(status_code=404, detail="Excel image not found")
    return FileResponse(image_path, media_type="image/png", headers={"Cache-Control": "public, max-age=86400"})

@app.get("/api/excel-records/{session_id}")
def get_excel_records(session_id: str, search: Optional[str] = None, page: int = 1, page_size: int = 50):
    if session_id not in session_cache or "excel_records" not in session_cache[session_id]:
        raise HTTPException(status_code=404, detail="No Excel records in session")
    records = session_cache[session_id]["excel_records"]
    if search:
        s = search.lower()
        records = [r for r in records if s in (r.get("filename") or "").lower() or s in (r.get("alt_text") or "").lower()]
    total = len(records)
    start = (page - 1) * page_size
    end = start + page_size
    return {
        "total": total,
        "page": page,
        "page_size": page_size,
        "records": records[start:end]
    }

@app.get("/api/figure-image/{session_id}/{filename}")
def get_figure_image(session_id: str, filename: str):
    image_path = SESSIONS_DIR / session_id / "figures" / filename
    if not image_path.exists():
        raise HTTPException(status_code=404, detail="Image not found")
    return FileResponse(image_path, media_type="image/png", headers={"Cache-Control": "public, max-age=86400"})

@app.get("/api/download-excel-zip/{session_id}")
def download_excel_zip(session_id: str):
    session_path = SESSIONS_DIR / session_id
    excel_dir = session_path / "excel_images"
    if not excel_dir.exists():
        raise HTTPException(status_code=404, detail="Excel images not found for session")

    zip_path = session_path / f"excel_images_{session_id[:8]}.zip"
    csv_buffer = io.StringIO()
    writer = csv.writer(csv_buffer)
    writer.writerow(["Row", "Sr No", "Filename", "Has Image", "Has Alt", "Authoritative Alt Text", "Original Alt", "Updated Alt"])

    records = session_cache.get(session_id, {}).get("excel_records", [])
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for rec in records:
            if rec.get("image_filename"):
                img_path = excel_dir / rec["image_filename"]
                if img_path.exists():
                    zip_file.write(img_path, arcname=f"excel_images/{rec['image_filename']}")
            writer.writerow([
                rec.get("row", ""),
                rec.get("sr_no", ""),
                rec.get("filename", ""),
                rec.get("has_image", False),
                rec.get("has_alt", False),
                rec.get("alt_text", ""),
                rec.get("original_alt", ""),
                rec.get("updated_alt", "")
            ])

        zip_file.writestr("excel_alt_manifest.csv", csv_buffer.getvalue())

    return FileResponse(
        zip_path,
        media_type="application/zip",
        filename=f"excel_manifest_{session_id[:8]}.zip"
    )

@app.get("/api/download-zip/{session_id}")
def download_zip(session_id: str):
    session_path = SESSIONS_DIR / session_id
    figures_dir = session_path / "figures"
    if not figures_dir.exists():
        raise HTTPException(status_code=404, detail="Session figures not found")

    zip_path = session_path / f"pdf_figures_{session_id[:8]}.zip"
    
    # Generate CSV manifest inside ZIP
    csv_buffer = io.StringIO()
    writer = csv.writer(csv_buffer)
    writer.writerow(["Figure ID", "Page", "MCIDs", "Status", "Alt Text", "BBox [x0, y0, x1, y1]", "Width (pt)", "Height (pt)", "Image File"])

    figures_meta = session_cache.get(session_id, {}).get("figures", [])
    meta_by_file = {f.get("crop_filename"): f for f in figures_meta}

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for img_file in figures_dir.glob("*.png"):
            zip_file.write(img_file, arcname=f"figures/{img_file.name}")
            meta = meta_by_file.get(img_file.name, {})
            writer.writerow([
                meta.get("figure_id", ""),
                meta.get("page_number", ""),
                str(meta.get("mcids", "")),
                meta.get("status", ""),
                meta.get("alt_text", "") or "",
                str(meta.get("bbox", "")),
                meta.get("bbox_width", ""),
                meta.get("bbox_height", ""),
                img_file.name
            ])

        zip_file.writestr("figures_manifest.csv", csv_buffer.getvalue())

    return FileResponse(
        zip_path,
        media_type="application/zip",
        filename=f"extracted_figures_{session_id[:8]}.zip"
    )

@app.get("/api/download-json/{session_id}")
def download_json(session_id: str):
    if session_id not in session_cache:
        raise HTTPException(status_code=404, detail="Session not found")
    return JSONResponse(content=session_cache[session_id])

class SingleAltInjection(BaseModel):
    figure_id: int
    alt_text: str

class BatchAltInjection(BaseModel):
    injections: Optional[Dict[int, str]] = None

class BatchAltRemoval(BaseModel):
    figure_ids: Optional[List[int]] = None

class UnselectedFiguresDownload(BaseModel):
    figure_ids: List[int]

@app.post("/api/inject-alt/{session_id}")
async def inject_all_alt(session_id: str, payload: Optional[BatchAltInjection] = None):
    """
    Injects alt texts into the PDF's StructTreeRoot elements where /S == '/Figure'.
    Uses either the provided injections map or defaults to all matched Excel alt texts.
    Saves the new accessible PDF to injected_accessible.pdf and returns updated metrics.
    """
    if session_id not in session_cache:
        raise HTTPException(status_code=404, detail="Session not found")

    session_path = SESSIONS_DIR / session_id
    base_pdf = session_path / "injected_accessible.pdf"
    if not base_pdf.exists():
        base_pdf = session_path / "input.pdf"

    if not base_pdf.exists():
        raise HTTPException(status_code=404, detail="Input PDF not found for session")

    temp_output = session_path / "injected_temp.pdf"
    target_output = session_path / "injected_accessible.pdf"
    figures = session_cache[session_id].get("figures", [])

    # Build injection dictionary (1-indexed figure_id -> alt_text)
    injections_map = {}
    if payload and payload.injections:
        injections_map = {int(k): str(v).strip() for k, v in payload.injections.items()}
    else:
        for f in figures:
            fig_id = f.get("figure_id")
            # Use matched Excel alt text or existing alt text
            alt = (f.get("excel_match") and f["excel_match"].get("alt_text")) or f.get("alt_text")
            if alt and fig_id:
                injections_map[fig_id] = alt.strip()

    if not injections_map:
        raise HTTPException(status_code=400, detail="No alt texts found to inject.")

    # Run injection via FigureExtractor
    extractor = None
    try:
        extractor = FigureExtractor(str(base_pdf))
        count = extractor.inject_alt_texts(injections_map, str(temp_output))
    finally:
        if extractor:
            extractor.close()

    if temp_output.exists():
        shutil.move(str(temp_output), str(target_output))

    # Update session cache
    for f in figures:
        fid = f.get("figure_id")
        if fid in injections_map:
            f["alt_text"] = injections_map[fid]
            f["has_alt"] = True
            f["status_label"] = "Injected"

    has_alt_count = sum(1 for f in figures if f["has_alt"])
    missing_alt_count = len(figures) - has_alt_count

    session_cache[session_id]["figures"] = figures
    session_cache[session_id]["has_alt_count"] = has_alt_count
    session_cache[session_id]["missing_alt_count"] = missing_alt_count
    session_cache[session_id]["has_injected_pdf"] = True

    return {
        "session_id": session_id,
        "status": "success",
        "injected_count": count,
        "total_figures": len(figures),
        "has_alt_count": has_alt_count,
        "missing_alt_count": missing_alt_count,
        "download_url": f"/api/download-injected-pdf/{session_id}",
        "figures": figures
    }

@app.post("/api/inject-single-alt/{session_id}")
async def inject_single_alt(session_id: str, payload: SingleAltInjection):
    """
    Injects alt text for a single figure into the PDF.
    """
    if session_id not in session_cache:
        raise HTTPException(status_code=404, detail="Session not found")

    session_path = SESSIONS_DIR / session_id
    base_pdf = session_path / "injected_accessible.pdf"
    if not base_pdf.exists():
        base_pdf = session_path / "input.pdf"

    if not base_pdf.exists():
        raise HTTPException(status_code=404, detail="Base PDF not found")

    temp_output = session_path / "injected_temp.pdf"
    target_output = session_path / "injected_accessible.pdf"

    extractor = None
    try:
        extractor = FigureExtractor(str(base_pdf))
        count = extractor.inject_alt_texts({payload.figure_id: payload.alt_text.strip()}, str(temp_output))
    finally:
        if extractor:
            extractor.close()

    if temp_output.exists():
        shutil.move(str(temp_output), str(target_output))

    # Update session cache
    figures = session_cache[session_id].get("figures", [])
    for f in figures:
        if f.get("figure_id") == payload.figure_id:
            f["alt_text"] = payload.alt_text.strip()
            f["has_alt"] = True
            f["status_label"] = "Injected"

    has_alt_count = sum(1 for f in figures if f["has_alt"])
    missing_alt_count = len(figures) - has_alt_count
    session_cache[session_id]["figures"] = figures
    session_cache[session_id]["has_alt_count"] = has_alt_count
    session_cache[session_id]["missing_alt_count"] = missing_alt_count
    session_cache[session_id]["has_injected_pdf"] = True

    return {
        "session_id": session_id,
        "status": "success",
        "figure_id": payload.figure_id,
        "injected_alt": payload.alt_text.strip(),
        "has_alt_count": has_alt_count,
        "missing_alt_count": missing_alt_count,
        "download_url": f"/api/download-injected-pdf/{session_id}"
    }

@app.post("/api/remove-alt/{session_id}")
async def remove_alt(session_id: str, payload: Optional[BatchAltRemoval] = None):
    """
    Removes /Alt accessibility texts from the PDF's StructTreeRoot elements where /S == '/Figure'.
    Can remove for selected figure IDs or for all figures if none specified.
    Saves the updated accessible PDF and returns updated metrics.
    """
    if session_id not in session_cache:
        raise HTTPException(status_code=404, detail="Session not found")

    session_path = SESSIONS_DIR / session_id
    base_pdf = session_path / "injected_accessible.pdf"
    if not base_pdf.exists():
        base_pdf = session_path / "input.pdf"

    if not base_pdf.exists():
        raise HTTPException(status_code=404, detail="Base PDF not found for session")

    temp_output = session_path / "injected_temp.pdf"
    target_output = session_path / "injected_accessible.pdf"
    figures = session_cache[session_id].get("figures", [])

    target_ids = payload.figure_ids if (payload and payload.figure_ids) else None

    extractor = None
    try:
        extractor = FigureExtractor(str(base_pdf))
        count = extractor.remove_alt_texts(target_ids, str(temp_output))
    finally:
        if extractor:
            extractor.close()

    if temp_output.exists():
        shutil.move(str(temp_output), str(target_output))

    target_set = set(target_ids) if target_ids else None
    for f in figures:
        fid = f.get("figure_id")
        if target_set is None or fid in target_set:
            f["alt_text"] = ""
            f["has_alt"] = False
            f["status_label"] = "Missing /Alt"

    has_alt_count = sum(1 for f in figures if f["has_alt"])
    missing_alt_count = len(figures) - has_alt_count

    session_cache[session_id]["figures"] = figures
    session_cache[session_id]["has_alt_count"] = has_alt_count
    session_cache[session_id]["missing_alt_count"] = missing_alt_count
    session_cache[session_id]["has_injected_pdf"] = True

    return {
        "session_id": session_id,
        "status": "success",
        "removed_count": count,
        "total_figures": len(figures),
        "has_alt_count": has_alt_count,
        "missing_alt_count": missing_alt_count,
        "download_url": f"/api/download-injected-pdf/{session_id}",
        "figures": figures
    }

@app.get("/api/download-injected-pdf/{session_id}")
def download_injected_pdf(session_id: str):
    session_path = SESSIONS_DIR / session_id
    injected_pdf = session_path / "injected_accessible.pdf"
    if not injected_pdf.exists():
        raise HTTPException(status_code=404, detail="Injected accessible PDF not found for this session")

    orig_name = session_cache.get(session_id, {}).get("filename", "document.pdf")
    base_name = Path(orig_name).stem
    download_filename = f"{base_name}_with_injected_alt.pdf"

    return FileResponse(
        injected_pdf,
        media_type="application/pdf",
        filename=download_filename
    )

@app.post("/api/download-unselected-excel/{session_id}")
async def download_unselected_excel(session_id: str, payload: UnselectedFiguresDownload):
    if session_id not in session_cache:
        raise HTTPException(status_code=404, detail="Session not found")

    session_path = SESSIONS_DIR / session_id
    figures_dir = session_path / "figures"
    if not figures_dir.exists():
        raise HTTPException(status_code=404, detail="Session figures not found")

    unselected_ids = set(payload.figure_ids)
    all_figures = session_cache[session_id].get("figures", [])

    unselected_figures = [f for f in all_figures if f.get("figure_id") in unselected_ids]
    if not unselected_figures:
        raise HTTPException(status_code=400, detail="No unselected figures found")

    excel_file_path = session_path / f"unselected_figures_{session_id[:8]}.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Unselected Figures"

    ws.append(["Figure ID", "Page", "MCIDs", "Status", "Alt Text", "BBox [x0, y0, x1, y1]", "Width (pt)", "Height (pt)", "Image File"])

    row_idx = 2
    for fig in unselected_figures:
        ws.append([
            fig.get("figure_id", ""),
            fig.get("page_number", ""),
            str(fig.get("mcids", "")),
            fig.get("status", ""),
            fig.get("alt_text", "") or "",
            str(fig.get("bbox", "")),
            fig.get("bbox_width", ""),
            fig.get("bbox_height", ""),
            fig.get("crop_filename", "")
        ])

        img_filename = fig.get("crop_filename")
        if img_filename:
            img_path = figures_dir / img_filename
            if img_path.exists():
                try:
                    excel_img = XLImage(str(img_path))
                    excel_img.width = 150
                    excel_img.height = 150
                    ws.add_image(excel_img, f"I{row_idx}")
                except Exception as e:
                    print(f"Warning: Could not embed image {img_filename}: {e}")

        row_idx += 1

    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 8
    ws.column_dimensions['C'].width = 12
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 25
    ws.column_dimensions['F'].width = 20
    ws.column_dimensions['G'].width = 12
    ws.column_dimensions['H'].width = 12
    ws.column_dimensions['I'].width = 18

    wb.save(excel_file_path)

    return FileResponse(
        excel_file_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=f"unselected_figures_{session_id[:8]}.xlsx"
    )

# Mount static frontend
app.mount("/", StaticFiles(directory=str(STATIC_DIR), html=True), name="static")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app:app", host="0.0.0.0", port=8000, reload=True)
