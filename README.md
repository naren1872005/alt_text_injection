# PDF Alt-Text Injector & Accessibility Tool

An automated tool for auditing, visually matching, and injecting authoritative alternative text (`/Alt`) into PDF documents using Computer Vision and Excel manifests.

---

## What is this project?

PDF documents (such as textbooks, research papers, and educational workbooks) require alternative text (`/Alt`) on figures, diagrams, and equations for screen readers and accessibility compliance (WCAG / PDF/UA).

Often, authors and editorial teams prepare authoritative descriptions and image previews in an **Excel spreadsheet**, while the **PDF** only contains tagged figures without alt-text (or with draft text).

This project completely automates the workflow:
1. **Extracts Figures & Formulas from PDF**: Reads the PDF's accessibility structure tree (`StructTreeRoot`), calculates exact bounding boxes, and generates high-resolution crops of all figures and equations.
2. **Extracts Images & Alt-Text from Excel**: Reads OpenXML workbooks to extract embedded images (floating drawings, in-cell pictures, and vector metafiles) alongside their authoritative descriptions.
3. **Visual & Structural AI Matching**: Automatically pairs PDF figures with corresponding Excel entries using Computer Vision (Perceptual Hashing, Projection Profiles, SIFT + RANSAC Containment Matching, and MathType equation parsing).
4. **Direct Alt-Text Injection**: Writes the approved alt-text directly into the PDF's structure tree tags without altering the visual appearance or layout of the document.
5. **Interactive Web UI**: Offers side-by-side inspection, split-screen zoom, manual review, and single-click accessible PDF download.

---

## How It Works

```
[ Upload PDF ]      ──▶  Extracts /Figure & /Formula tags + high-res crops
[ Upload Excel ]    ──▶  Extracts images & authoritative alt-text descriptions
         │
         ▼
[ Computer Vision ] ──▶  Multi-tier matching (pHash + SIFT + MathType token analysis)
         │
         ▼
[ Interactive UI ]  ──▶  Review side-by-side matches with confidence scores
         │
         ▼
[ Inject & Export ] ──▶  Injects /Alt into PDF StructTree and downloads remediated PDF
```

---

## Quick Start

### 1. Install Dependencies
```bash
pip install -r requirements.txt
```

### 2. Run the Application
```bash
python -m uvicorn app:app --host 127.0.0.1 --port 8000 --reload
```

### 3. Open in Browser
Visit **`http://localhost:8000`** to start using the tool.

---

## Key Modules

- **`app.py`**: FastAPI backend server, REST endpoints, and session management.
- **`extractor.py`**: PDF structure tree parser, MCID bounding box engine, crop rasterizer, and `/Alt` injector.
- **`excel_parser.py`**: OpenXML DrawingML and In-Cell picture extractor with multi-language column detection.
- **`matcher.py`**: Computer Vision matching engine (pHash, dHash, SIFT homography containment, and MathType parser).
- **`static/`**: Modern dark-mode web application (HTML5, Vanilla CSS, and JavaScript).
